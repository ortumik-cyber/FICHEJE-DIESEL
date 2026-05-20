import os
import json
import hashlib
import secrets
import httpx
import io
from datetime import datetime, timedelta, timezone
from typing import Optional
from fastapi import FastAPI, Depends, HTTPException, Request
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
from jose import JWTError, jwt
from pydantic import BaseModel
import firebase_admin
from firebase_admin import credentials, firestore
from apscheduler.schedulers.background import BackgroundScheduler
import openpyxl

# ── CONFIG ────────────────────────────────────────────────────────────────
EMPRESA_PREFIX = "DIESEL"
APP_URL        = os.getenv("APP_URL", "https://diesel-fichaje.onrender.com")
REDIRECT_URI   = f"{APP_URL}/api/auth/onedrive/callback"
JWT_SECRET     = os.getenv("JWT_SECRET", "CAMBIA_ESTO_EN_PRODUCCION")
JWT_ALGORITHM  = "HS256"
JWT_EXPIRE_H   = 12
AZURE_CLIENT_ID     = os.getenv("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "")
ONEDRIVE_FOLDER     = "AutoescuelaDiesel-Fichaje"
MS_SCOPES           = "Files.ReadWrite offline_access"

# ── FIREBASE ──────────────────────────────────────────────────────────────
_creds_json = os.getenv("FIREBASE_CREDENTIALS", "{}")
cred = credentials.Certificate(json.loads(_creds_json))
firebase_admin.initialize_app(cred)
db = firestore.client()

# ── HELPERS ───────────────────────────────────────────────────────────────
hash_pwd = lambda p: hashlib.sha256(f"{EMPRESA_PREFIX}:{p}".encode()).hexdigest()

def now_utc() -> str:
    return datetime.now(timezone.utc).isoformat()

def create_token(uid: str, role: str) -> str:
    exp = datetime.now(timezone.utc) + timedelta(hours=JWT_EXPIRE_H)
    return jwt.encode({"sub": uid, "role": role, "exp": exp}, JWT_SECRET, algorithm=JWT_ALGORITHM)

def normalize_key(nombre: str, apellido: str) -> str:
    key = f"{nombre.lower().strip()}.{apellido.lower().strip()}".replace(" ", "")
    for a, b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ü","u"),("ñ","n")]:
        key = key.replace(a, b)
    return key

def audit(type_: str, uid: str, details: str = ""):
    db.collection("audit").add({"type": type_, "userId": uid, "ts": now_utc(), "details": details})

def build_excel(fichajes_data: list, emp_map: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fichajes"
    ws.append(["Empleado", "Tipo", "Fecha", "Hora", "IP", "Latitud", "Longitud", "Precisión"])
    for f in fichajes_data:
        ts_str = f.get("ts", "").replace("Z", "")
        try:
            ts = datetime.fromisoformat(ts_str)
            fecha = ts.strftime("%Y-%m-%d")
            hora = ts.strftime("%H:%M:%S")
        except Exception:
            fecha = hora = ts_str
        nombre = emp_map.get(f.get("uid", ""), "Desconocido")
        ws.append([nombre, f.get("tipo",""), fecha, hora, f.get("ip",""), f.get("lat",""), f.get("lon",""), f.get("acc","")])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

async def upload_to_onedrive(token: str, folder: str, filename: str, content: bytes) -> bool:
    path = f"{ONEDRIVE_FOLDER}/{folder}/{filename}"
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{path}:/content"
    async with httpx.AsyncClient() as client:
        r = await client.put(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            },
            content=content,
            timeout=30
        )
    return r.status_code in [200, 201]

async def get_valid_token(uid: str) -> Optional[str]:
    doc = db.collection("users").document(uid).get()
    if not doc.exists:
        return None
    u = doc.to_dict()
    if not u.get("onedrive_connected"):
        return None
    expiry_str = u.get("onedrive_token_expiry")
    if expiry_str:
        expiry = datetime.fromisoformat(expiry_str.replace("Z",""))
        if expiry.tzinfo is None:
            expiry = expiry.replace(tzinfo=timezone.utc)
        if datetime.now(timezone.utc) > expiry - timedelta(minutes=5):
            async with httpx.AsyncClient() as client:
                r = await client.post(
                    "https://login.microsoftonline.com/common/oauth2/v2.0/token",
                    data={
                        "client_id": AZURE_CLIENT_ID,
                        "client_secret": AZURE_CLIENT_SECRET,
                        "refresh_token": u.get("onedrive_refresh_token",""),
                        "grant_type": "refresh_token",
                    }
                )
            if r.status_code == 200:
                tokens = r.json()
                new_expiry = (datetime.now(timezone.utc) + timedelta(seconds=tokens.get("expires_in", 3600))).isoformat()
                db.collection("users").document(uid).update({
                    "onedrive_access_token": tokens["access_token"],
                    "onedrive_refresh_token": tokens.get("refresh_token", u.get("onedrive_refresh_token","")),
                    "onedrive_token_expiry": new_expiry,
                })
                return tokens["access_token"]
            return None
    return u.get("onedrive_access_token")

async def _run_backup(tipo: str, uid: str):
    token = await get_valid_token(uid)
    if not token:
        return
    today = datetime.now(timezone.utc)
    emp_docs = db.collection("users").where("deleted","==",False).get()
    emp_map = {d.id: d.to_dict().get("fullname","") for d in emp_docs}
    if tipo == "daily":
        ayer = (today - timedelta(days=1)).strftime("%Y-%m-%d")
        docs = db.collection("fichajes").where("ts",">=",f"{ayer}T00:00:00").where("ts","<=",f"{ayer}T23:59:59").get()
        content = build_excel([d.to_dict() for d in docs], emp_map)
        await upload_to_onedrive(token, "Diario", f"{ayer}.xlsx", content)
    elif tipo == "monthly":
        mes = today.strftime("%Y-%m")
        docs = db.collection("fichajes").where("ts",">=",f"{mes}-01T00:00:00").get()
        content = build_excel([d.to_dict() for d in docs], emp_map)
        await upload_to_onedrive(token, "Mensual", f"{mes}.xlsx", content)
    elif tipo == "annual":
        year = today.strftime("%Y")
        docs = db.collection("fichajes").where("ts",">=",f"{year}-01-01T00:00:00").get()
        content = build_excel([d.to_dict() for d in docs], emp_map)
        await upload_to_onedrive(token, "Anual", f"{year}.xlsx", content)
    elif tipo == "full":
        docs = db.collection("fichajes").get()
        content = build_excel([d.to_dict() for d in docs], emp_map)
        await upload_to_onedrive(token, "Backup-completo", f"{today.strftime('%Y-%m-%d')}_completo.xlsx", content)
    audit("BACKUP", uid, f"Backup {tipo}")

def scheduled_backup(tipo: str):
    import asyncio
    admins = db.collection("users").where("role","==","admin").get()
    for adoc in admins:
        ad = adoc.to_dict()
        if ad.get("onedrive_connected") and not ad.get("deleted"):
            try:
                loop = asyncio.new_event_loop()
                loop.run_until_complete(_run_backup(tipo, adoc.id))
                loop.close()
            except Exception as e:
                print(f"[Backup] Error uid={adoc.id}: {e}")

# ── APP ───────────────────────────────────────────────────────────────────
app = FastAPI(title="Diesel Fichaje")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"]
)

# ── MODELOS ───────────────────────────────────────────────────────────────
class LoginBody(BaseModel):
    loginKey: str
    password: str

class SetupBody(BaseModel):
    nombre: str; apellido: str; email: str; password: str

class ChangePwdBody(BaseModel):
    oldPassword: str; newPassword: str

class ProfileBody(BaseModel):
    nombre: Optional[str] = None
    apellido: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    dept: Optional[str] = None

class EmpleadoBody(BaseModel):
    nombre: str; apellido: str; email: str
    phone: Optional[str] = ""
    dept: Optional[str] = ""
    role: Optional[str] = "employee"
    password: str

class EmpleadoUpdateBody(BaseModel):
    nombre: Optional[str] = None
    apellido: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    dept: Optional[str] = None
    role: Optional[str] = None
    active: Optional[bool] = None

class ResetPwdBody(BaseModel):
    userId: str; newPassword: str

# ── DEPS ──────────────────────────────────────────────────────────────────
async def get_current_user(request: Request):
    token = request.headers.get("Authorization", "").replace("Bearer ", "").strip()
    if not token:
        raise HTTPException(401, "Token requerido")
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
        uid = payload.get("sub")
    except JWTError:
        raise HTTPException(401, "Token inválido o expirado")
    doc = db.collection("users").document(uid).get()
    if not doc.exists:
        raise HTTPException(401, "Usuario no encontrado")
    user = doc.to_dict()
    if user.get("deleted") or not user.get("active"):
        raise HTTPException(403, "Acceso denegado")
    user["id"] = uid
    return user

async def get_admin(user=Depends(get_current_user)):
    if user.get("role") != "admin":
        raise HTTPException(403, "Solo administradores")
    return user

# ── ENDPOINTS PÚBLICOS ────────────────────────────────────────────────────
@app.get("/api/status")
def status_check():
    docs = db.collection("users").where("role","==","admin").limit(1).get()
    return {"configured": len(docs) > 0}

@app.post("/api/setup")
def setup(body: SetupBody):
    if len(db.collection("users").where("role","==","admin").limit(1).get()) > 0:
        raise HTTPException(400, "Sistema ya configurado")
    uid = secrets.token_urlsafe(16)
    login_key = normalize_key(body.nombre, body.apellido)
    db.collection("users").document(uid).set({
        "nombre": body.nombre, "apellido": body.apellido,
        "fullname": f"{body.nombre} {body.apellido}",
        "loginKey": login_key, "email": body.email,
        "phone": "", "dept": "Dirección", "role": "admin",
        "hash": hash_pwd(body.password), "active": True, "firstLogin": False,
        "failedAttempts": 0, "lockedUntil": None,
        "createdAt": now_utc(), "createdBy": None,
        "onedrive_connected": False, "onedrive_access_token": None,
        "onedrive_refresh_token": None, "onedrive_token_expiry": None,
        "backup_time": "23:00", "deleted": False, "deletedBy": None, "deletedAt": None,
    })
    audit("CREATE_USER", uid, f"Setup inicial: {login_key}")
    return {"ok": True, "loginKey": login_key}

@app.post("/api/auth/login")
def login(body: LoginBody):
    docs = [d for d in db.collection("users").where("loginKey","==",body.loginKey.lower().strip()).get()
            if not d.to_dict().get("deleted")]
    if not docs:
        raise HTTPException(401, "Credenciales incorrectas")
    doc = docs[0]; user = doc.to_dict(); uid = doc.id
    # Bloqueo temporal
    if user.get("lockedUntil"):
        locked = datetime.fromisoformat(user["lockedUntil"].replace("Z","")).replace(tzinfo=timezone.utc)
        if datetime.now(timezone.utc) < locked:
            secs = int((locked - datetime.now(timezone.utc)).total_seconds())
            raise HTTPException(403, f"Cuenta bloqueada. Intenta en {secs//60}m {secs%60}s")
        db.collection("users").document(uid).update({"lockedUntil": None, "failedAttempts": 0})
    if not user.get("active"):
        raise HTTPException(403, "Usuario desactivado")
    if user.get("hash") != hash_pwd(body.password):
        attempts = user.get("failedAttempts", 0) + 1
        upd = {"failedAttempts": attempts}
        if attempts >= 3:
            upd["lockedUntil"] = (datetime.now(timezone.utc) + timedelta(minutes=5)).isoformat()
        db.collection("users").document(uid).update(upd)
        raise HTTPException(401, "Credenciales incorrectas")
    db.collection("users").document(uid).update({"failedAttempts": 0, "lockedUntil": None})
    audit("LOGIN", uid, f"IP: {body.loginKey}")
    token = create_token(uid, user.get("role","employee"))
    return {
        "token": token,
        "user": {
            "id": uid, "nombre": user["nombre"], "apellido": user["apellido"],
            "fullname": user.get("fullname",""), "role": user.get("role"),
            "firstLogin": user.get("firstLogin", False),
        }
    }

# ── ENDPOINTS EMPLEADO ────────────────────────────────────────────────────
@app.get("/api/me")
def me(user=Depends(get_current_user)):
    return {k: v for k, v in user.items() if k not in ["hash","onedrive_access_token","onedrive_refresh_token"]}

@app.post("/api/auth/change-password")
def change_password(body: ChangePwdBody, user=Depends(get_current_user)):
    if user["hash"] != hash_pwd(body.oldPassword):
        raise HTTPException(400, "Contraseña actual incorrecta")
    if len(body.newPassword) < 6:
        raise HTTPException(400, "Mínimo 6 caracteres")
    db.collection("users").document(user["id"]).update({
        "hash": hash_pwd(body.newPassword), "firstLogin": False
    })
    audit("PWD_RESET", user["id"], "Auto-cambio contraseña")
    return {"ok": True}

@app.put("/api/profile")
def update_profile(body: ProfileBody, user=Depends(get_current_user)):
    upd = {k: v for k, v in body.dict().items() if v is not None}
    if "nombre" in upd or "apellido" in upd:
        upd["fullname"] = f"{upd.get('nombre', user['nombre'])} {upd.get('apellido', user['apellido'])}"
    db.collection("users").document(user["id"]).update(upd)
    return {"ok": True}

@app.post("/api/fichar")
async def fichar(request: Request, user=Depends(get_current_user)):
    body = {}
    try:
        body = await request.json()
    except Exception:
        pass
    docs = db.collection("fichajes").where("uid","==",user["id"]).order_by("ts", direction=firestore.Query.DESCENDING).limit(1).get()
    ultimo = docs[0].to_dict() if docs else None
    tipo = "entrada" if not ultimo or ultimo["tipo"] == "salida" else "salida"
    ip = request.headers.get("X-Forwarded-For", request.client.host or "").split(",")[0].strip()
    fichaje = {
        "uid": user["id"], "tipo": tipo, "ts": now_utc(), "ip": ip,
        "lat": body.get("lat"), "lon": body.get("lon"), "acc": body.get("acc"),
    }
    ref = db.collection("fichajes").add(fichaje)
    audit("FICHAJE", user["id"], f"{tipo} · {ip}")
    return {"ok": True, "tipo": tipo, "id": ref[1].id}

@app.get("/api/fichajes")
def mis_fichajes(user=Depends(get_current_user)):
    docs = db.collection("fichajes").where("uid","==",user["id"]).order_by("ts").get()
    return [{"id": d.id, **d.to_dict()} for d in docs]

# ── ONEDRIVE ──────────────────────────────────────────────────────────────
@app.get("/api/auth/onedrive/connect")
def onedrive_connect(token: str):
    try:
        payload = jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
        uid = payload["sub"]
    except Exception:
        raise HTTPException(401, "Token inválido")
    url = (
        f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
        f"?client_id={AZURE_CLIENT_ID}&response_type=code"
        f"&redirect_uri={REDIRECT_URI}&scope={MS_SCOPES}"
        f"&state={uid}&prompt=select_account"
    )
    return RedirectResponse(url)

@app.get("/api/auth/onedrive/callback")
async def onedrive_callback(code: str, state: str):
    async with httpx.AsyncClient() as client:
        r = await client.post(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            data={
                "client_id": AZURE_CLIENT_ID, "client_secret": AZURE_CLIENT_SECRET,
                "code": code, "redirect_uri": REDIRECT_URI, "grant_type": "authorization_code",
            }
        )
    if r.status_code != 200:
        return RedirectResponse(f"{APP_URL}?onedrive=error")
    tokens = r.json()
    expiry = (datetime.now(timezone.utc) + timedelta(seconds=tokens.get("expires_in", 3600))).isoformat()
    db.collection("users").document(state).update({
        "onedrive_connected": True,
        "onedrive_access_token": tokens["access_token"],
        "onedrive_refresh_token": tokens.get("refresh_token"),
        "onedrive_token_expiry": expiry,
    })
    audit("ONEDRIVE_CONNECT", state, "OAuth completado")
    return RedirectResponse(f"{APP_URL}?onedrive=success")

@app.post("/api/auth/onedrive/disconnect")
def onedrive_disconnect(user=Depends(get_current_user)):
    db.collection("users").document(user["id"]).update({
        "onedrive_connected": False, "onedrive_access_token": None,
        "onedrive_refresh_token": None, "onedrive_token_expiry": None,
    })
    audit("ONEDRIVE_CONNECT", user["id"], "Desconectado")
    return {"ok": True}

@app.get("/api/onedrive/status")
def onedrive_status(user=Depends(get_current_user)):
    return {"connected": bool(user.get("onedrive_connected"))}

@app.post("/api/backup/{tipo}")
async def backup_manual(tipo: str, user=Depends(get_current_user)):
    if tipo not in ["daily","monthly","annual","full"]:
        raise HTTPException(400, "Tipo inválido: daily|monthly|annual|full")
    token = await get_valid_token(user["id"])
    if not token:
        raise HTTPException(400, "OneDrive no conectado")
    await _run_backup(tipo, user["id"])
    return {"ok": True}

# ── ADMIN ─────────────────────────────────────────────────────────────────
@app.get("/api/admin/empleados")
def admin_empleados(admin=Depends(get_admin)):
    docs = db.collection("users").where("deleted","==",False).get()
    skip = {"hash","onedrive_access_token","onedrive_refresh_token"}
    return [{"id": d.id, **{k:v for k,v in d.to_dict().items() if k not in skip}} for d in docs]

@app.post("/api/admin/empleados")
def admin_create(body: EmpleadoBody, admin=Depends(get_admin)):
    uid = secrets.token_urlsafe(16)
    login_key = normalize_key(body.nombre, body.apellido)
    db.collection("users").document(uid).set({
        "nombre": body.nombre, "apellido": body.apellido,
        "fullname": f"{body.nombre} {body.apellido}",
        "loginKey": login_key, "email": body.email,
        "phone": body.phone, "dept": body.dept, "role": body.role,
        "hash": hash_pwd(body.password), "active": True, "firstLogin": True,
        "failedAttempts": 0, "lockedUntil": None,
        "createdAt": now_utc(), "createdBy": admin["id"],
        "onedrive_connected": False, "onedrive_access_token": None,
        "onedrive_refresh_token": None, "onedrive_token_expiry": None,
        "backup_time": "23:00", "deleted": False, "deletedBy": None, "deletedAt": None,
    })
    audit("CREATE_USER", admin["id"], f"Creado: {login_key}")
    return {"ok": True, "id": uid, "loginKey": login_key}

@app.put("/api/admin/empleados/{uid}")
def admin_update(uid: str, body: EmpleadoUpdateBody, admin=Depends(get_admin)):
    upd = {k: v for k, v in body.dict().items() if v is not None}
    if "nombre" in upd or "apellido" in upd:
        doc_data = db.collection("users").document(uid).get().to_dict()
        upd["fullname"] = f"{upd.get('nombre', doc_data['nombre'])} {upd.get('apellido', doc_data['apellido'])}"
    db.collection("users").document(uid).update(upd)
    audit("UPDATE_USER", admin["id"], f"Actualizado uid={uid}")
    return {"ok": True}

@app.delete("/api/admin/empleados/{uid}")
async def admin_delete(uid: str, admin=Depends(get_admin)):
    doc = db.collection("users").document(uid).get()
    if not doc.exists:
        raise HTTPException(404, "Usuario no encontrado")
    user = doc.to_dict()
    # Backup previo en OneDrive si hay fichajes
    token = await get_valid_token(admin["id"])
    if token:
        fichajes = db.collection("fichajes").where("uid","==",uid).get()
        if fichajes:
            emp_map = {uid: user.get("fullname","")}
            content = build_excel([d.to_dict() for d in fichajes], emp_map)
            nombre_safe = user.get("fullname","").replace(" ","_")
            await upload_to_onedrive(
                token, "Backup-completo",
                f"{datetime.now(timezone.utc).strftime('%Y-%m-%d')}_{nombre_safe}.xlsx",
                content
            )
    db.collection("users").document(uid).update({
        "deleted": True, "deletedBy": admin["id"],
        "deletedAt": now_utc(), "active": False
    })
    audit("DELETE_USER", admin["id"], f"Eliminado: {user.get('loginKey','')}")
    return {"ok": True}

@app.post("/api/admin/reset-pwd")
def admin_reset_pwd(body: ResetPwdBody, admin=Depends(get_admin)):
    if len(body.newPassword) < 6:
        raise HTTPException(400, "Mínimo 6 caracteres")
    db.collection("users").document(body.userId).update({
        "hash": hash_pwd(body.newPassword),
        "firstLogin": True, "failedAttempts": 0, "lockedUntil": None,
    })
    audit("PWD_RESET", admin["id"], f"Reset uid={body.userId}")
    return {"ok": True}

@app.get("/api/admin/fichajes")
def admin_fichajes(admin=Depends(get_admin)):
    docs = db.collection("fichajes").order_by("ts", direction=firestore.Query.DESCENDING).limit(500).get()
    emp_docs = db.collection("users").where("deleted","==",False).get()
    emp_map = {d.id: d.to_dict().get("fullname","") for d in emp_docs}
    result = []
    for d in docs:
        f = {"id": d.id, **d.to_dict()}
        f["empleado"] = emp_map.get(f.get("uid",""), "Desconocido")
        result.append(f)
    return result

@app.delete("/api/admin/fichajes/{fid}")
def admin_delete_fichaje(fid: str, admin=Depends(get_admin)):
    doc = db.collection("fichajes").document(fid).get()
    if not doc.exists:
        raise HTTPException(404, "Fichaje no encontrado")
    f = doc.to_dict()
    ts = datetime.fromisoformat(f["ts"].replace("Z",""))
    if datetime.utcnow() - ts < timedelta(days=4*365):
        raise HTTPException(400, "Retención mínima 4 años (RDL 8/2019)")
    db.collection("fichajes").document(fid).delete()
    audit("DELETE_FICHAJE", admin["id"], f"Eliminado {fid}")
    return {"ok": True}

@app.get("/api/admin/auditoria")
def admin_auditoria(admin=Depends(get_admin)):
    docs = db.collection("audit").order_by("ts", direction=firestore.Query.DESCENDING).limit(500).get()
    return [{"id": d.id, **d.to_dict()} for d in docs]

@app.get("/api/admin/export")
def admin_export(admin=Depends(get_admin)):
    skip = {"hash","onedrive_access_token","onedrive_refresh_token"}
    fichajes  = [{"id": d.id, **d.to_dict()} for d in db.collection("fichajes").get()]
    users     = [{"id": d.id, **{k:v for k,v in d.to_dict().items() if k not in skip}} for d in db.collection("users").get()]
    auditoria = [{"id": d.id, **d.to_dict()} for d in db.collection("audit").order_by("ts", direction=firestore.Query.DESCENDING).limit(500).get()]
    audit("EXPORT", admin["id"], f"{len(fichajes)} fichajes · {len(users)} usuarios")
    return {"fichajes": fichajes, "users": users, "auditoria": auditoria}

# ── SCHEDULER ─────────────────────────────────────────────────────────────
scheduler = BackgroundScheduler()
scheduler.add_job(lambda: scheduled_backup("daily"),   "cron", hour=23, minute=0)
scheduler.add_job(lambda: scheduled_backup("monthly"), "cron", day=1,  hour=0,  minute=30)
scheduler.add_job(lambda: scheduled_backup("annual"),  "cron", month=1, day=1, hour=1, minute=0)
scheduler.start()

# ── STATIC ────────────────────────────────────────────────────────────────
app.mount("/", StaticFiles(directory="static", html=True), name="static")
