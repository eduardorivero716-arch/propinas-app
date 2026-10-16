"""
Filtrado y Cálculo de Propinas - con login por sucursal y Panel de Control de Acceso
====================================================================================
Versión autocontenida: TODO está en este archivo (lectura del Excel y estilos incluidos).
Solo necesitas app.py y requirements.txt en el repositorio.

requirements.txt
----------------
    streamlit>=1.50
    pandas>=2.0
    openpyxl>=3.1
    xlrd>=2.0.1        # solo si subes archivos .xls antiguos

.streamlit/config.toml  (recomendado: evita que el modo oscuro del navegador rompa los colores)
----------------------------------------------------------------------------------------------
    [theme]
    base = "light"

.streamlit/secrets.toml  (local)  /  Settings > Secrets  (Streamlit Cloud)
--------------------------------------------------------------------------
    [admin]
    username = "admin"
    name     = "Administrador"
    password_hash = "pbkdf2_sha256$600000$..."  # recomendado; password en texto plano solo por compatibilidad

    # Opcional: sucursales iniciales. También se pueden crear desde el panel.
    [users.centro]
    name          = "Local Centro"
    password_hash = "pbkdf2_sha256$600000$..."   # o: password = "clave-inicial"
    active        = true
    pct_atencion  = 70                         # opcional: % de reparto de la sucursal (deben sumar 100)
    pct_barra     = 10
    pct_cocina    = 20
    alfredo_atencion = 500                     # opcional: aporte del bachero por asistencia (lunes a sábado)
    alfredo_barra    = 100

Envío de resumen por correo (opcional)
--------------------------------------
Sin configuración, el botón del resumen semanal SIMULA el envío.
Para enviar de verdad, agrega en Secrets:

    [smtp]
    host       = "smtp.gmail.com"
    port       = 587                       # 587 con STARTTLS; usa 465 con use_ssl = true
    username   = "propinas@tulocal.cl"
    password   = "clave-de-aplicacion"
    from_email = "propinas@tulocal.cl"     # opcional
    use_ssl    = false

Con SMTP configurado, solo se puede enviar al correo cargado en la tabla de trabajadores
(no a direcciones libres) y hay un límite de envíos por hora.

Cómo se guardan los accesos
---------------------------
* st.secrets es de SOLO LECTURA mientras la app corre. El administrador y las sucursales
  "semilla" viven en st.secrets. Los cambios hechos desde el panel se guardan en users_db.json
  (solo hashes, nunca contraseñas en texto plano).
* En Streamlit Cloud el disco es EFÍMERO: users_db.json se pierde al reiniciar o redesplegar.
  Usa en el panel el botón "Generar configuración para st.secrets" y pega el resultado en
  Settings > Secrets para hacerlo permanente.
* No subas users_db.json ni secrets.toml a GitHub (agrégalos a .gitignore).

Turnos que cruzan medianoche
----------------------------
Una marca de madrugada (antes de las 05:00) se asigna al día anterior si ese día quedó con una
marca abierta (cantidad impar). Así, entrar a las 20:00 y salir a las 01:00 cuenta como un solo
turno del día de entrada.
"""
from __future__ import annotations

import hashlib
import html
import hmac
import io
import json
import math
import os
import re
import secrets as pysecrets
import smtplib
import ssl
import threading
import time
import unicodedata
from collections import defaultdict
from datetime import date, datetime, time as dtime, timedelta
from email.message import EmailMessage
from pathlib import Path
from urllib.parse import quote

import pandas as pd
import streamlit as st

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
st.set_page_config(
    page_title="Propinas | Gestión semanal",
    page_icon="💠",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_CSS = """
:root {
  --ink: #17263c;
  --muted: #5a6a7e;
  --line: #e2e8f0;
  --canvas: #f4f7fb;
  --paper: #ffffff;
  --navy: #17365d;
  --navy-2: #244b78;
  --teal: #167d77;
  --gold: #d6a84f;
  --ok-bg: #f2faf8;
  --ok-line: #c8e4de;
  --ok-ink: #16665f;
  --step-ink: #66758a;
  --radius: 18px;
  color-scheme: light;
}
[data-testid="stAppViewContainer"] { background: var(--canvas); }
[data-testid="stHeader"] { background: rgba(244,247,251,.90); }
.block-container { max-width: 1500px; padding-top: 1.7rem; padding-bottom: 3.5rem; }
[data-testid="stSidebar"] { background: #fff; border-right: 1px solid var(--line); }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: var(--muted); }
h1, h2, h3, h4 { color: var(--ink) !important; letter-spacing: -.025em; }
h1 { font-weight: 750 !important; }
h2, h3 { font-weight: 700 !important; }
p, label, li { color: #334155; }
[data-testid="stCaptionContainer"] { color: var(--muted); }
div[class*="st-key-"][class*="-card"] { background: #fff; border-radius: var(--radius); border-color: var(--line); }
[data-testid="stMetric"] {
  background: var(--paper); border: 1px solid var(--line); border-radius: 16px;
  padding: 15px 18px; box-shadow: 0 5px 18px rgba(19,43,74,.035);
}
[data-testid="stMetricLabel"] p { color: var(--muted) !important; font-size: .83rem; }
[data-testid="stMetricValue"] { color: var(--navy) !important; font-weight: 750; }
[data-testid="stExpander"] { background: #fff; border: 1px solid var(--line); border-radius: 14px; overflow: hidden; }
[data-testid="stForm"] { border-color: var(--line) !important; border-radius: 14px !important; background: #fff; }
[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 12px; overflow: hidden; }
[data-testid="stFileUploaderDropzone"] { background: #fbfdff; border: 1.5px dashed #b8c7d9; border-radius: 15px; }
[data-testid="stAlert"] { border-radius: 12px; }
.stButton > button, [data-testid="stFormSubmitButton"] > button {
  border-radius: 11px; min-height: 2.7rem; font-weight: 650;
  transition: transform .12s ease, box-shadow .12s ease;
}
.stButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
  transform: translateY(-1px); box-shadow: 0 5px 14px rgba(23,54,93,.12);
}
.stButton > button:focus-visible, [data-testid="stFormSubmitButton"] > button:focus-visible {
  outline: 2px solid var(--teal); outline-offset: 2px;
}
button[kind="primary"], [data-testid="stFormSubmitButton"] button[kind="primary"] {
  background: var(--navy) !important; border-color: var(--navy) !important; color: white !important;
}
[data-baseweb="input"] > div, [data-baseweb="select"] > div, [data-testid="stDateInput"] input,
[data-testid="stTimeInput"] input { border-radius: 10px; }
.brand-lockup { display:flex; align-items:center; gap:14px; margin: 0 0 1.6rem 0; }
.brand-name { font-weight: 800; color: var(--ink); font-size: 1rem; }
.brand-sub { font-size: .75rem; color: var(--muted); }
.brand-mark { width:48px; height:48px; display:flex; align-items:center; justify-content:center;
  border-radius:15px; color:#fff; font-size:24px; font-weight:800; background:linear-gradient(140deg,#17365d,#267f83);
  box-shadow:0 8px 20px rgba(23,54,93,.18); flex-shrink:0; }
.eyebrow { text-transform:uppercase; letter-spacing:.15em; font-size:.70rem; font-weight:800; color:var(--teal); margin-bottom:7px; }
.hero-title { color:var(--ink); font-size:clamp(1.8rem, 3vw, 2.6rem); font-weight:800; line-height:1.1; letter-spacing:-.04em; }
.hero-copy { color:var(--muted); font-size:1rem; margin-top:9px; max-width:850px; line-height:1.55; }
.hero-shell { background:linear-gradient(115deg,#fff 10%,#f7fafc 65%,#edf5f5 100%); border:1px solid var(--line);
  border-radius:22px; padding:25px 28px; margin-bottom:22px; box-shadow:0 8px 26px rgba(19,43,74,.035); }
.hero-row { display:flex; align-items:flex-start; gap:17px; }
.hero-mark { width:56px; height:56px; border-radius:17px; flex-shrink:0; display:flex; align-items:center; justify-content:center;
  background:var(--navy); color:white; font-size:27px; box-shadow:0 8px 20px rgba(23,54,93,.15); }
.role-pill { display:inline-block; background:#e9f3f2; color:#126b65; border:1px solid #cee5e2; border-radius:999px;
  padding:5px 10px; font-size:.75rem; font-weight:750; letter-spacing:.02em; }
.section-heading { display:flex; align-items:flex-start; gap:13px; padding:3px 0 8px 0; margin-bottom:8px; }
.section-number { min-width:38px; height:38px; display:flex; align-items:center; justify-content:center; border-radius:12px;
  background:#eaf0f8; color:var(--navy); font-size:.84rem; font-weight:850; }
.section-title { color:var(--ink); font-size:1.24rem; font-weight:780; letter-spacing:-.025em; margin-top:1px; }
.section-copy { color:var(--muted); margin-top:4px; font-size:.91rem; line-height:1.5; }
.workflow { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:10px; margin: 0 0 24px 0; }
.workflow-step { display:flex; align-items:center; gap:10px; padding:12px 13px; border-radius:13px; border:1px solid var(--line); background:#fff; min-height:63px; }
.workflow-dot { width:29px; height:29px; display:flex; align-items:center; justify-content:center; border-radius:50%; background:#eef2f7; color:#5f6f85; font-size:.76rem; font-weight:800; flex-shrink:0; }
.workflow-label { color:var(--step-ink); font-size:.82rem; font-weight:680; line-height:1.25; }
.workflow-step.done { background:var(--ok-bg); border-color:var(--ok-line); }
.workflow-step.done .workflow-dot { background:var(--teal); color:#fff; }
.workflow-step.done .workflow-label { color:var(--ok-ink); }
.workflow-step.current { border-color:#9db4d1; box-shadow:inset 0 0 0 1px #d8e3ef; }
.workflow-step.current .workflow-dot { background:var(--navy); color:#fff; }
@media (max-width: 820px) { .workflow { grid-template-columns:repeat(2,minmax(0,1fr)); } .hero-shell { padding:20px; } .block-container { padding-top:1rem; } }
@media (max-width: 480px) { .hero-mark { width:45px; height:45px; } }
@media (prefers-reduced-motion: reduce) { .stButton > button, [data-testid="stFormSubmitButton"] > button { transition: none; transform: none !important; } }
"""
st.markdown(f"<style>\n{APP_CSS}\n</style>", unsafe_allow_html=True)

DB_PATH = Path(os.environ.get("USERS_DB_PATH", "users_db.json"))
_DB_LOCK = threading.Lock()

PBKDF2_ITERATIONS = 600_000
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_SECONDS = 120
FAIL_WINDOW_SECONDS = 900
MIN_PASSWORD_LEN = 8
USERNAME_RE = re.compile(r"^[a-z0-9_-]{3,32}$")
LOGIN_ERROR = "Acceso no autorizado o usuario desactivado"

DAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
AREAS = ["atencion", "barra", "cocina", "Omitir"]
DEFAULT_PCTS = {"atencion": 70.0, "barra": 10.0, "cocina": 20.0}
PCT_AREAS = [("atencion", "Atención / Garzones"), ("barra", "Barra"), ("cocina", "Cocina")]
AREA_ORDER = ["atencion", "cocina", "barra"]
TURNO_ORDER = {"Día": 0, "Noche": 1, "Único": 2, "Día completo": 3}

DEFAULT_ALFREDO = {"atencion": 500, "barra": 100}
ALFREDO_AREAS = [("atencion", "Garzón / Garzona"), ("barra", "Barra")]
ALFREDO_WEEKDAYS = range(0, 6)  # 0 = Lunes ... 5 = Sábado

REPLACEMENT_AREAS = ["atencion", "barra", "cocina"]
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
MAX_EMAILS_PER_USER_HOUR = 20
MAX_EMAILS_GLOBAL_HOUR = 60

DAY_SECONDS = 86400
EARLY_CUTOFF = 5 * 3600  # marcas antes de las 05:00 pueden ser la salida de un turno de la noche anterior
MAX_SHIFT_SECONDS = 16 * 3600


# ============================================================================
# LECTURA DEL EXCEL DE ASISTENCIA (antes en attendance_import.py)
# ============================================================================
def norm(s) -> str:
    """Minúsculas, sin tildes y sin espacios sobrantes (para buscar y comparar)."""
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.casefold().split())


def parse_time(v):
    """Convierte hora (time, datetime, fracción de día de Excel o texto 'HH:MM') a segundos desde medianoche."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, datetime):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, dtime):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        v = float(v)
        if 0 <= v < 1:
            sec = int(round(v * DAY_SECONDS))
        elif 1 <= v < 80000 and v - int(v) > 0:
            sec = int(round((v - int(v)) * DAY_SECONDS))
        else:
            return None
        return sec % DAY_SECONDS
    s = str(v).strip()
    if not s:
        return None
    m = re.match(r"^(\d{1,2})\s*[:.hH]\s*(\d{2})(?:\s*[:.]\s*(\d{2}))?\s*([aApP]\.?[mM]\.?)?$", s)
    if not m:
        return None
    hh, mm, ss = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    ampm = (m.group(4) or "").lower().replace(".", "")
    if ampm == "pm" and hh < 12:
        hh += 12
    if ampm == "am" and hh == 12:
        hh = 0
    if hh > 23 or mm > 59 or ss > 59:
        return None
    return hh * 3600 + mm * 60 + ss


def parse_date(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if 30000 < v < 80000:
            return date(1899, 12, 30) + timedelta(days=int(v))
        return None
    s = str(v).strip().split(" ")[0] if str(v).strip() else ""
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y", "%d/%m/%y", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def read_marks(file_bytes: bytes):
    """
    Lee el Excel (columnas Nombre, Fecha, Hora; una fila por marca).
    Devuelve (marks, empties, skipped, sheet):
      marks   = [(nombre, fecha, segundos)]
      empties = [(nombre, fecha)]  filas con hora vacía
      skipped = filas ignoradas por nombre o fecha inválidos
      sheet   = nombre de la hoja usada
    """
    sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, dtype=object)
    wanted = {"nombre", "fecha", "hora"}
    for sheet_name, raw in sheets.items():
        header_row, cols = None, {}
        for i in range(min(15, len(raw))):
            found = {norm(v): j for j, v in enumerate(raw.iloc[i].tolist()) if not is_missing(v)}
            if wanted <= set(found):
                header_row, cols = i, found
                break
        if header_row is None:
            continue
        marks, empties, skipped = [], [], 0
        canonical = {}
        for rec in raw.iloc[header_row + 1:].itertuples(index=False):
            name_raw, date_raw, time_raw = rec[cols["nombre"]], rec[cols["fecha"]], rec[cols["hora"]]
            if all(is_missing(x) or str(x).strip() == "" for x in (name_raw, date_raw, time_raw)):
                continue  # fila completamente vacía
            name = " ".join(str(name_raw).split()) if not is_missing(name_raw) else ""
            d = parse_date(date_raw)
            if not name or d is None:
                skipped += 1
                continue
            name = canonical.setdefault(name.casefold(), name)  # "ANA" y "Ana" son la misma persona
            sec = parse_time(time_raw)
            if sec is None:
                empties.append((name, d))
            else:
                marks.append((name, d, sec))
        return marks, empties, skipped, str(sheet_name)
    raise ValueError("No encontré las columnas Nombre, Fecha y Hora en ninguna hoja del archivo.")


# ============================================================================
# SEGURIDAD: HASH Y VERIFICACIÓN DE CONTRASEÑAS
# ============================================================================
def hash_password(password: str) -> str:
    salt = pysecrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return f"pbkdf2_sha256${PBKDF2_ITERATIONS}${salt.hex()}${dk.hex()}"


def _verify_hash(password: str, stored: str) -> bool:
    try:
        algo, iters_raw, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        iters = int(iters_raw)
        if not 100_000 <= iters <= 1_000_000:
            return False
        salt = bytes.fromhex(salt_hex)
        expected = bytes.fromhex(hash_hex)
        if len(salt) < 16 or len(expected) != 32:
            return False
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iters)
        return hmac.compare_digest(dk, expected)
    except (ValueError, TypeError):
        return False


def verify_password(record: dict, password: str) -> bool:
    if record.get("password_hash"):
        return _verify_hash(password, str(record["password_hash"]))
    if record.get("password") is not None:
        return hmac.compare_digest(str(record["password"]).encode("utf-8"), password.encode("utf-8"))
    return False


@st.cache_resource
def _dummy_hash() -> str:
    """Hash de relleno para igualar tiempos cuando el usuario no existe (se calcula una sola vez)."""
    return hash_password("dummy-password")


# ============================================================================
# LÍMITES EN EL SERVIDOR (intentos de login y correos): sobreviven a refrescar la página
# ============================================================================
@st.cache_resource
def _guard_store() -> dict:
    return {"lock": threading.Lock(), "fails": {}, "emails": []}


def _fail_key(username: str) -> str:
    return username.strip().lower()[:64]


def lock_remaining(username: str) -> int:
    store = _guard_store()
    with store["lock"]:
        rec = store["fails"].get(_fail_key(username))
        if not rec:
            return 0
        return max(0, math.ceil(rec["until"] - time.time()))


def register_failure(username: str) -> None:
    store, now = _guard_store(), time.time()
    with store["lock"]:
        fails = store["fails"]
        if len(fails) > 5000:  # evita que crezca sin límite
            for k in [k for k, v in fails.items() if now - v["last"] > FAIL_WINDOW_SECONDS and v["until"] < now]:
                fails.pop(k, None)
        key = _fail_key(username)
        rec = fails.get(key)
        if not rec or now - rec["last"] > FAIL_WINDOW_SECONDS:
            rec = {"n": 0, "until": 0.0, "last": now}
        rec["n"] += 1
        rec["last"] = now
        if rec["n"] >= MAX_FAILED_ATTEMPTS:
            rec["until"] = now + LOCKOUT_SECONDS
            rec["n"] = 0
        fails[key] = rec


def clear_failures(username: str) -> None:
    store = _guard_store()
    with store["lock"]:
        store["fails"].pop(_fail_key(username), None)


def email_quota_ok(user: str) -> bool:
    store, now = _guard_store(), time.time()
    with store["lock"]:
        store["emails"] = [(t, u) for t, u in store["emails"] if now - t < 3600]
        mine = sum(1 for _t, u in store["emails"] if u == user)
        return mine < MAX_EMAILS_PER_USER_HOUR and len(store["emails"]) < MAX_EMAILS_GLOBAL_HOUR


def register_email(user: str) -> None:
    store = _guard_store()
    with store["lock"]:
        store["emails"].append((time.time(), user))


# ============================================================================
# ALMACENAMIENTO DE USUARIOS (st.secrets + archivo JSON superpuesto)
# ============================================================================
def _secrets_section(name, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:  # no existe secrets.toml
        return default


def get_admin():
    cfg = _secrets_section("admin")
    if not cfg:
        return None
    cfg = dict(cfg)
    if not cfg.get("username") or not (cfg.get("password") or cfg.get("password_hash")):
        return None
    return {
        "username": str(cfg["username"]).strip().lower(),
        "name": str(cfg.get("name", "Administrador")),
        "password": cfg.get("password"),
        "password_hash": cfg.get("password_hash"),
    }


def _valid_pcts(p) -> bool:
    try:
        vals = [float(p[a]) for a, _ in PCT_AREAS]
    except (KeyError, TypeError, ValueError):
        return False
    return all(0 <= v <= 100 for v in vals) and round(sum(vals), 6) == 100.0


def _valid_alfredo(r) -> bool:
    try:
        vals = [float(r[a]) for a, _ in ALFREDO_AREAS]
    except (KeyError, TypeError, ValueError):
        return False
    return all(0 <= v <= 100000 for v in vals)


def _alfredo_from_cfg(cfg: dict):
    r = {a: cfg.get(f"alfredo_{a}") for a, _ in ALFREDO_AREAS}
    return {a: int(float(v)) for a, v in r.items()} if _valid_alfredo(r) else None


def branch_alfredo(rec) -> dict:
    r = (rec or {}).get("alfredo")
    return {a: int(float(r[a])) for a, _ in ALFREDO_AREAS} if _valid_alfredo(r) else dict(DEFAULT_ALFREDO)


def _pcts_from_cfg(cfg: dict):
    p = {a: cfg.get(f"pct_{a}") for a, _ in PCT_AREAS}
    return {a: float(v) for a, v in p.items()} if _valid_pcts(p) else None


def branch_pcts(rec) -> dict:
    p = (rec or {}).get("pcts")
    return {a: float(p[a]) for a, _ in PCT_AREAS} if _valid_pcts(p) else dict(DEFAULT_PCTS)


def _secret_users() -> dict:
    users = {}
    section = _secrets_section("users", {}) or {}
    for uname, cfg in section.items():
        cfg = dict(cfg)
        users[str(uname).lower()] = {
            "name": str(cfg.get("name", uname)),
            "password": cfg.get("password"),
            "password_hash": cfg.get("password_hash"),
            "active": bool(cfg.get("active", True)),
            "pcts": _pcts_from_cfg(cfg),
            "alfredo": _alfredo_from_cfg(cfg),
        }
    return users


def _read_overlay() -> dict:
    try:
        with open(DB_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def _write_overlay(data: dict) -> None:
    tmp = DB_PATH.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, DB_PATH)


def load_users() -> dict:
    """Usuarios SUCURSAL: base de st.secrets + cambios guardados desde el panel."""
    secret_users = _secret_users()
    users = dict(secret_users)
    overlay = _read_overlay()
    for uname, ov in overlay.items():
        rec = dict(users.get(uname, {"name": uname, "password": None, "password_hash": None, "active": True}))
        rec.update({k: v for k, v in ov.items() if k in ("name", "password_hash", "active", "pcts", "alfredo")})
        if "password_hash" in ov:
            rec["password"] = None
        users[uname] = rec
    secret_names = set(secret_users)
    overlay_names = set(overlay)
    for uname, rec in users.items():
        if uname in secret_names and uname in overlay_names:
            rec["source"] = "secrets + panel"
        elif uname in secret_names:
            rec["source"] = "secrets"
        else:
            rec["source"] = "panel"
    return users


def _update_overlay(uname: str, fields: dict) -> None:
    with _DB_LOCK:
        data = _read_overlay()
        data.setdefault(uname, {}).update(fields)
        _write_overlay(data)


def create_branch(username: str, name: str, password: str) -> str | None:
    """Devuelve un mensaje de error, o None si se creó bien."""
    uname = username.strip().lower()
    admin = get_admin()
    if not USERNAME_RE.match(uname):
        return "El usuario debe tener 3 a 32 caracteres: letras minúsculas, números, guion o guion bajo."
    if not name.strip():
        return "Escribe el nombre del local o sucursal."
    if len(password) < MIN_PASSWORD_LEN:
        return f"La contraseña debe tener al menos {MIN_PASSWORD_LEN} caracteres."
    if (admin and uname == admin["username"]) or uname in load_users():
        return "Ese nombre de usuario ya existe."
    try:
        _update_overlay(uname, {"name": name.strip(), "password_hash": hash_password(password), "active": True})
    except OSError as e:
        return f"No se pudo guardar: {e}"
    return None


def get_branch_pcts(username: str) -> dict:
    return branch_pcts(load_users().get(username))


def set_pcts(uname: str, pcts: dict) -> str | None:
    if not _valid_pcts(pcts):
        return "Los porcentajes deben estar entre 0 y 100 y sumar exactamente 100%."
    try:
        _update_overlay(uname, {"pcts": {a: float(pcts[a]) for a, _ in PCT_AREAS}})
    except OSError as e:
        return f"No se pudo guardar: {e}"
    return None


def get_branch_alfredo(username: str) -> dict:
    return branch_alfredo(load_users().get(username))


def set_alfredo(uname: str, rates: dict) -> str | None:
    if not _valid_alfredo(rates):
        return "Los montos deben ser números entre 0 y 100.000."
    try:
        _update_overlay(uname, {"alfredo": {a: int(float(rates[a])) for a, _ in ALFREDO_AREAS}})
    except OSError as e:
        return f"No se pudo guardar: {e}"
    return None


def set_active(uname: str, active: bool) -> str | None:
    try:
        _update_overlay(uname, {"active": bool(active)})
    except OSError as e:
        return f"No se pudo guardar: {e}"
    return None


def set_password(uname: str, password: str) -> str | None:
    if len(password) < MIN_PASSWORD_LEN:
        return f"La contraseña debe tener al menos {MIN_PASSWORD_LEN} caracteres."
    try:
        _update_overlay(uname, {"password_hash": hash_password(password)})
    except OSError as e:
        return f"No se pudo guardar: {e}"
    return None


def _toml_str(s: str) -> str:
    return str(s).replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ").replace("\r", " ")


def build_secrets_toml(users: dict) -> str:
    lines = ["# Pega este bloque en Settings > Secrets (junto a la sección [admin]).", ""]
    for uname in sorted(users):
        rec = users[uname]
        if rec.get("password_hash"):
            pw_hash = rec["password_hash"]
        elif rec.get("password"):
            pw_hash = hash_password(str(rec["password"]))
        else:  # sin contraseña: no se exporta (evitaría dejar una cuenta con clave vacía)
            lines += [f"# Se omitió '{uname}': no tiene contraseña definida.", ""]
            continue
        lines += [
            f"[users.{uname}]",
            f'name = "{_toml_str(rec["name"])}"',
            f'password_hash = "{pw_hash}"',
            f'active = {"true" if rec["active"] else "false"}',
        ]
        if rec.get("pcts"):
            lines += [f"pct_{a} = {float(rec['pcts'][a])}" for a, _ in PCT_AREAS]
        if rec.get("alfredo"):
            lines += [f"alfredo_{a} = {int(rec['alfredo'][a])}" for a, _ in ALFREDO_AREAS]
        lines.append("")
    return "\n".join(lines)


def render_hero(eyebrow: str, title: str, subtitle: str, role: str | None = None):
    role_html = f'<div style="margin-top:12px"><span class="role-pill">{html.escape(role)}</span></div>' if role else ""
    st.markdown(
        f"""
        <div class="hero-shell">
          <div class="hero-row">
            <div class="hero-mark">✦</div>
            <div>
              <div class="eyebrow">{html.escape(eyebrow)}</div>
              <div class="hero-title">{html.escape(title)}</div>
              <div class="hero-copy">{html.escape(subtitle)}</div>
              {role_html}
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_section_heading(number: str, title: str, subtitle: str = ""):
    desc = f'<div class="section-copy">{html.escape(subtitle)}</div>' if subtitle else ""
    st.markdown(
        f'<div class="section-heading"><div class="section-number">{html.escape(number)}</div>'
        f'<div><div class="section-title">{html.escape(title)}</div>{desc}</div></div>',
        unsafe_allow_html=True,
    )


def render_workflow_progress():
    has_file = st.session_state.get("attendance") is not None
    has_workers = st.session_state.get("workers_df") is not None
    tip_keys = [f"tip_{mode}_{i}" for mode in "udn" for i in range(7)]
    has_tips = any(int(st.session_state.get(k, 0) or 0) > 0 for k in tip_keys)
    has_result = st.session_state.get("result") is not None
    states = [has_file, has_workers, has_tips, has_result]
    labels = ["Cargar asistencia", "Configurar equipo", "Ingresar propinas", "Revisar resultados"]
    current = next((i for i, done in enumerate(states) if not done), 3)
    bits = []
    for i, (label, done) in enumerate(zip(labels, states)):
        cls = "done" if done else ("current" if i == current else "")
        mark = "✓" if done else f"{i + 1:02d}"
        bits.append(
            f'<div class="workflow-step {cls}"><div class="workflow-dot">{mark}</div>'
            f'<div class="workflow-label">{html.escape(label)}</div></div>'
        )
    st.markdown('<div class="workflow">' + "".join(bits) + '</div>', unsafe_allow_html=True)


# ============================================================================
# AUTENTICACIÓN Y SESIÓN
# ============================================================================
def authenticate(username: str, password: str):
    uname = username.strip().lower()
    admin = get_admin()
    if admin and hmac.compare_digest(uname.encode(), admin["username"].encode()):
        if verify_password(admin, password):
            return {"username": admin["username"], "name": admin["name"], "role": "admin"}
        return None
    rec = load_users().get(uname)
    if rec is None:
        _verify_hash(password, _dummy_hash())
        return None
    if verify_password(rec, password) and rec["active"]:
        return {"username": uname, "name": rec["name"], "role": "sucursal"}
    return None


def current_user():
    """Sesión activa, revalidada en cada interacción: si la sucursal fue desactivada, se cierra."""
    auth = st.session_state.get("auth")
    if not auth:
        return None
    if auth["role"] == "admin":
        admin = get_admin()
        if admin and admin["username"] == auth["username"]:
            return auth
    else:
        rec = load_users().get(auth["username"])
        if rec and rec["active"]:
            auth["name"] = rec["name"]
            return auth
    st.session_state.clear()
    st.session_state["login_error"] = True
    return None


def logout():
    st.session_state.clear()
    st.rerun()


def render_login():
    _, col, _ = st.columns([1, 1.15, 1])
    with col:
        st.markdown(
            '<div style="padding: 7vh 0 8px 0; text-align:center">'
            '<div class="brand-mark" style="margin:0 auto 18px auto;width:62px;height:62px;font-size:28px">✦</div>'
            '<div class="eyebrow">GESTIÓN OPERATIVA</div>'
            '<div class="hero-title" style="font-size:2rem">Control de propinas</div>'
            '<div class="hero-copy" style="margin:10px auto 25px auto;text-align:center">'
            'Una vista clara del equipo, los turnos y la distribución semanal.</div></div>',
            unsafe_allow_html=True,
        )

        if get_admin() is None:
            st.error(
                "Falta configurar el administrador. Agrega una sección [admin] con username y "
                "password (o password_hash) en .streamlit/secrets.toml o en Settings > Secrets."
            )
            st.stop()

        if st.session_state.pop("login_error", False):
            st.error(LOGIN_ERROR)

        with st.form("login_form"):
            username = st.text_input("Nombre de Usuario / Sucursal", autocomplete="username")
            password = st.text_input("Contraseña", type="password", autocomplete="current-password")
            submitted = st.form_submit_button("Iniciar Sesión")

        if submitted:
            wait = lock_remaining(username)
            if wait > 0:
                st.error(f"Demasiados intentos fallidos. Espera {wait} segundos.")
                return
            user = authenticate(username, password)
            if user:
                clear_failures(username)
                st.session_state.clear()
                st.session_state["auth"] = user
                st.rerun()
            register_failure(username)
            st.error(LOGIN_ERROR)


# ============================================================================
# HERRAMIENTA: ASISTENCIA CON 4 MARCAS FIJAS POR DÍA Y REVISIÓN DE EXCEPCIONES
# ============================================================================
SLOTS = ["Entrada 1", "Salida 1", "Entrada 2", "Salida 2"]
ATT_COLS = ["Nombre", "Fecha"] + SLOTS + ["Marcas originales", "Nota"]


def is_missing(v) -> bool:
    return v is None or (not isinstance(v, str) and bool(pd.isna(v)))


def sec_to_time(sec):
    if is_missing(sec):
        return None
    sec = int(sec) % DAY_SECONDS
    return dtime(sec // 3600, (sec % 3600) // 60, sec % 60)


def to_date(v):
    return v.date() if isinstance(v, datetime) else v


def hhmm(sec: int) -> str:
    base = f"{(sec % DAY_SECONDS) // 3600:02d}:{(sec % 3600) // 60:02d}"
    return base + " (+1 día)" if sec >= DAY_SECONDS else base


def normalize_slots(vals: list) -> list:
    """Si una marca es menor o igual a la anterior pero razonable al día siguiente, se entiende que cruzó medianoche."""
    out, prev = [], None
    for v in vals:
        if v is None:
            out.append(None)
            continue
        v = int(v)
        if prev is not None and v <= prev and v + DAY_SECONDS - prev <= MAX_SHIFT_SECONDS:
            v += DAY_SECONDS
        out.append(v)
        prev = v
    return out


def slot_problems(vals) -> list:
    """Valida un día con el esquema [Entrada 1, Salida 1, Entrada 2, Salida 2]."""
    e1, s1, e2, s2 = [None if is_missing(v) else int(v) for v in vals]
    if e1 is None and s1 is None and e2 is None and s2 is None:
        return ["Sin marcas"]
    out = []
    if e1 is None and s1 is None:
        out.append("Falta la jornada 1 (Entrada 1 y Salida 1)")
    else:
        if e1 is None:
            out.append("Falta Entrada 1")
        if s1 is None:
            out.append("Falta Salida 1")
    if e2 is None and s2 is not None:
        out.append("Falta Entrada 2")
    if s2 is None and e2 is not None:
        out.append("Falta Salida 2")
    filled = [v for v in (e1, s1, e2, s2) if v is not None]
    if any(b <= a for a, b in zip(filled, filled[1:])):
        out.append("Horas fuera de orden (cada marca debe ser posterior a la anterior)")
    return out


def needs_review(rec: dict) -> bool:
    return bool(slot_problems([rec[c] for c in SLOTS])) or bool(rec.get("Nota"))


def problem_text(rec: dict) -> str:
    parts = slot_problems([rec[c] for c in SLOTS])
    if rec.get("Nota"):
        parts.append(rec["Nota"])
    return "; ".join(parts)


def assign_marks_to_days(marks) -> dict:
    """
    Agrupa las marcas por (nombre, fecha). Una marca de madrugada (antes de las 05:00) se asigna al
    día anterior si ese día quedó con una marca abierta (cantidad impar): es la salida de un turno nocturno.
    Los segundos de esas marcas se guardan sumando 24 h (ej.: 01:00 -> 25:00).
    """
    times_by = defaultdict(set)
    for name, d, sec in marks:
        times_by[(name, d)].add(sec)
    final = {}
    for name, d in sorted(times_by, key=lambda k: (k[0].casefold(), k[1])):
        ts = sorted(times_by[(name, d)])
        prev = final.get((name, d - timedelta(days=1)))
        if prev is not None and len(prev) % 2 == 1 and ts and ts[0] < EARLY_CUTOFF:
            prev.append(ts.pop(0) + DAY_SECONDS)
            prev.sort()
        final[(name, d)] = ts
    return final


def build_attendance(marks, empties) -> pd.DataFrame:
    """
    DataFrame principal: UNA fila por trabajador y fecha, con 4 columnas fijas de tiempo
    (Entrada 1, Salida 1, Entrada 2, Salida 2) en segundos desde medianoche (pueden pasar de 86400
    si la salida fue de madrugada). Si hay más de 4 marcas se usan las 4 primeras y se deja una Nota.
    """
    valid = assign_marks_to_days(marks)
    blank = defaultdict(int)
    for name, d in empties:
        blank[(name, d)] += 1

    keys = {k for k, v in valid.items() if v} | set(blank)
    rows = []
    for (name, d) in sorted(keys, key=lambda k: (k[1], k[0].casefold())):
        times = sorted(valid.get((name, d), ()))
        n_blank = blank.get((name, d), 0)
        slots = (times[:4] + [None] * 4)[:4]

        notes = []
        if len(times) > 4:
            notes.append(f"Hay {len(times)} marcas en el archivo; solo se cargaron las 4 primeras")
        if n_blank:
            notes.append(f"{n_blank} fila(s) con la hora vacía en el archivo")

        original = ", ".join(hhmm(t) for t in times) if times else "(ninguna)"
        if n_blank:
            original += f" (+{n_blank} sin hora)"

        rows.append({
            "Nombre": name, "Fecha": d,
            **dict(zip(SLOTS, slots)),
            "Marcas originales": original,
            "Nota": "; ".join(notes),
        })
    return pd.DataFrame(rows, columns=ATT_COLS, dtype=object)


def review_mask(df: pd.DataFrame) -> pd.Series:
    return pd.Series([needs_review(r) for r in df.to_dict("records")], index=df.index, dtype=bool)


def apply_corrections(main: pd.DataFrame, fixes: dict) -> pd.DataFrame:
    out = []
    for rec in main.to_dict("records"):
        if needs_review(rec):
            fx = fixes.get((rec["Nombre"], rec["Fecha"]))
            if fx:
                if fx["drop"]:
                    continue
                rec = {**rec, **dict(zip(SLOTS, fx["s"])), "Nota": ""}
        out.append(rec)
    return pd.DataFrame(out, columns=ATT_COLS, dtype=object)


def resolve_pairs(final: pd.DataFrame):
    """Convierte cada día en intervalos [Entrada 1 -> Salida 1] y [Entrada 2 -> Salida 2]."""
    pairs, pending = [], []
    for rec in final.to_dict("records"):
        vals = [rec[c] for c in SLOTS]
        if slot_problems(vals) or rec.get("Nota"):
            pending.append(rec)
            continue
        e1, s1, e2, s2 = vals
        pairs.append((rec["Nombre"], rec["Fecha"], int(e1), int(s1)))
        if not is_missing(e2):
            pairs.append((rec["Nombre"], rec["Fecha"], int(e2), int(s2)))
    return pairs, pending


def render_missing_marks():
    att = st.session_state.get("attendance")
    if att is None or att.empty:
        return

    flagged = att[review_mask(att)]
    if flagged.empty:
        st.success("Todas las marcas están completas: no hay registros por corregir.")
        return

    st.markdown("**Registros con marcas faltantes**")
    st.caption(
        f"{len(flagged)} de {len(att)} días tienen marcas faltantes o por revisar "
        f"({flagged['Nombre'].nunique()} personas). Los {len(att) - len(flagged)} días completos no se muestran. "
        "Cada día admite hasta 4 marcas: Entrada 1, Salida 1, Entrada 2 y Salida 2. "
        "La columna Marcas originales muestra las horas que venían en el archivo, ya cargadas en las columnas de la derecha. "
        "Completa lo que falte; si no hubo segunda jornada, deja Entrada 2 y Salida 2 vacías. "
        "Si una salida fue después de medianoche, escribe la hora normal (ej. 01:00): se entiende como del día siguiente. "
        "Marca Descartar para ignorar todo ese día."
    )

    recs = flagged.to_dict("records")
    base = pd.DataFrame({
        "Nombre": [r["Nombre"] for r in recs],
        "Fecha": [r["Fecha"] for r in recs],
        "Marcas originales": [r["Marcas originales"] for r in recs],
        "Problema": [problem_text(r) for r in recs],
        **{c: pd.Series([sec_to_time(r[c]) for r in recs], dtype=object) for c in SLOTS},
        "Descartar": [False] * len(recs),
    })

    time_cfg = {c: st.column_config.TimeColumn(c, format="HH:mm", step=60) for c in SLOTS}
    with st.form("fix_form"):
        edited = st.data_editor(
            base,
            key="fix_editor",
            hide_index=True,
            num_rows="fixed",
            width="stretch",
            disabled=["Nombre", "Fecha", "Marcas originales", "Problema"],
            column_config={
                "Fecha": st.column_config.DateColumn("Fecha", format="DD-MM-YYYY"),
                "Marcas originales": st.column_config.TextColumn("Marcas originales", help="Horas que venían en el archivo."),
                "Problema": st.column_config.TextColumn("Problema detectado"),
                **time_cfg,
                "Descartar": st.column_config.CheckboxColumn("Descartar", help="Ignora todo ese día en el cálculo."),
            },
        )
        saved = st.form_submit_button("Guardar correcciones", type="primary")

    if saved:
        st.session_state["fixes"] = {
            (rec["Nombre"], to_date(rec["Fecha"])): {
                "s": normalize_slots([parse_time(rec[c]) for c in SLOTS]),
                "drop": bool(rec["Descartar"]),
            }
            for rec in edited.to_dict("records")
        }

    fixes = st.session_state.get("fixes")
    if fixes is None:
        st.info("Aún no guardas correcciones.")
    else:
        _, pending = resolve_pairs(apply_corrections(att, fixes))
        if pending:
            st.warning(
                f"Quedan {len(pending)} días sin resolver: faltan marcas o están fuera de orden. "
                "Complétalos o márcalos como Descartar, y vuelve a guardar."
            )
        else:
            st.success("Todos los registros fueron resueltos y se integraron a los datos. Ya puedes calcular.")


# ============================================================================
# HERRAMIENTA: CÁLCULO
# ============================================================================
def money(n) -> str:
    return "$" + f"{int(round(n)):,}".replace(",", ".")


def hours_fmt(h) -> str:
    return f"{h:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_date(d: date) -> str:
    return d.strftime("%d-%m-%Y")


def compute(pairs, areas: dict, modes: list, cut_sec: int, tips: dict, pcts: dict, alfredo_rates: dict,
            adjustments=None, extra_hours=None) -> dict:
    """
    modes[i] = True  -> el día i (0 = Lunes) está en Modo Pozo Único.
    modes[i] = False -> el día i está en Modo Turnos (Día / Noche según la hora de corte).

    areas = {nombre: [área_lunes, ..., área_domingo]}; "Omitir" excluye ese día del reparto.

    Reglas: atención y barra por turno (o por día entero en Pozo Único); cocina SIEMPRE por día entero.
    Las propinas se reparten en pesos enteros con el método del mayor residuo: la suma de cada grupo
    coincide exactamente con su pozo (sin diferencias de redondeo).

    Descuento Alfredo: por cada día de asistencia de lunes a sábado se descuenta el monto configurado según el área.
    """
    missing_pct = [a for a in AREA_ORDER if a not in pcts]
    if missing_pct or any(v < 0 for v in pcts.values()) or abs(sum(pcts.values()) - 1) > 1e-9:
        raise ValueError("Los porcentajes de reparto deben sumar exactamente 100%.")

    if any(a not in alfredo_rates or alfredo_rates[a] < 0 for a, _ in ALFREDO_AREAS):
        raise ValueError("Los montos de Alfredo deben ser números mayores o iguales a 0.")

    def area_of(name: str, d: date) -> str:
        week = areas.get(name)
        return week[d.weekday()] if week else "Omitir"

    if not pairs and not adjustments and not extra_hours:
        raise ValueError("No hay tramos completos (entrada y salida) para calcular.")
    if all(a == "Omitir" for week in areas.values() for a in week):
        raise ValueError('Todos los trabajadores están en "Omitir". Asigna un área al menos a uno.')

    # Una sola semana: dos fechas del mismo día de la semana duplicarían el pozo
    wd_dates = defaultdict(set)
    for d in ({p[1] for p in pairs} | {e["date"] for e in extra_hours or []} | {a["date"] for a in adjustments or []}):
        wd_dates[d.weekday()].add(d)
    repeated = [f"{DAYS[wd]} ({', '.join(fmt_date(x) for x in sorted(ds))})"
                for wd, ds in sorted(wd_dates.items()) if len(ds) > 1]
    if repeated:
        raise ValueError(
            "El archivo contiene más de una semana: " + "; ".join(repeated) + ". "
            "Carga una sola semana (7 días) para no duplicar los pozos de propinas."
        )

    def day_total(wd: int) -> float:
        return tips["unico"][wd] if modes[wd] else tips["dia"][wd] + tips["noche"][wd]

    def pool_of(d: date, turno: str) -> float:
        wd = d.weekday()
        if turno == "Único":
            return tips["unico"][wd]
        return tips["dia"][wd] if turno == "Día" else tips["noche"][wd]

    warnings = []

    # 1) y 2) Cada tramo completo se reparte según el modo de su día (las salidas pasadas de medianoche cuentan en el turno Noche)
    hours = defaultdict(float)  # (fecha, turno, nombre) -> horas
    for name, d, ent, sal in pairs:
        if area_of(name, d) == "Omitir":
            continue
        if modes[d.weekday()]:
            hours[(d, "Único", name)] += (sal - ent) / 3600
        else:
            if min(sal, cut_sec) - ent > 0:
                hours[(d, "Día", name)] += (min(sal, cut_sec) - ent) / 3600
            if sal - max(ent, cut_sec) > 0:
                hours[(d, "Noche", name)] += (sal - max(ent, cut_sec)) / 3600

    # 2a) Reemplazos / personal sin marcaje
    for e in extra_hours or []:
        d, name = e["date"], e["name"]
        if area_of(name, d) == "Omitir":
            continue
        turno = "Único" if modes[d.weekday()] else (e["turno"] if e["turno"] in ("Día", "Noche") else "Día")
        hours[(d, turno, name)] += e["hours"]

    # 2b) Ajustes manuales
    adj_applied = defaultdict(float)
    cocina_adj = defaultdict(float)
    n_adj = 0
    for adj in adjustments or []:
        name, d, delta = adj["name"], adj["date"], adj["minutes"] / 60
        area = area_of(name, d)
        label = f"{name} ({DAYS[d.weekday()]} {fmt_date(d)})"
        if area == "Omitir":
            warnings.append(f"Ajuste ignorado: {label} no tiene un área asignada ese día.")
            continue
        n_adj += 1
        adj_applied[name] += delta
        if area == "cocina":
            cocina_adj[(d, name)] += delta
            continue
        turno = "Único" if modes[d.weekday()] else (adj["turno"] if adj["turno"] in ("Día", "Noche") else "Día")
        new = hours.get((d, turno, name), 0.0) + delta
        if new > 1e-9:
            hours[(d, turno, name)] = new
        else:
            hours.pop((d, turno, name), None)
            if new < -1e-9:
                warnings.append(f"Ajuste de {label}: restaba más horas de las trabajadas en ese turno; quedó en 0.")
    if not hours and not cocina_adj:
        raise ValueError("No se pudieron calcular horas. Revisa que haya tramos completos de personas con un área asignada.")

    # 3) Horas por (fecha, turno, área) para atención y barra; horas por día para cocina
    area_hours = defaultdict(lambda: {"atencion": 0.0, "barra": 0.0})
    cocina_day = defaultdict(float)
    for (d, turno, name), h in hours.items():
        area = area_of(name, d)
        if area == "cocina":
            cocina_day[(d, name)] += h
        else:
            area_hours[(d, turno)][area] += h
    for key, delta in cocina_adj.items():
        new = cocina_day.get(key, 0.0) + delta
        if new > 1e-9:
            cocina_day[key] = new
        else:
            cocina_day.pop(key, None)
            if new < -1e-9:
                warnings.append(f"Ajuste de {key[1]} ({DAYS[key[0].weekday()]} {fmt_date(key[0])}): restaba más horas de las trabajadas; quedó en 0.")
    cocina_total = defaultdict(float)
    for (d, _name), h in cocina_day.items():
        cocina_total[d] += h

    # 4) Reparto proporcional
    rows = []
    for (d, turno, name), h in hours.items():
        area = area_of(name, d)
        if area == "cocina":
            continue
        total_area = area_hours[(d, turno)][area]
        tip = (h / total_area) * (pool_of(d, turno) * pcts[area]) if total_area > 0 else 0.0
        rows.append({"date": d, "turno": turno, "name": name, "area": area, "hours": h, "tip": tip})
    for (d, name), h in cocina_day.items():
        pool = day_total(d.weekday()) * pcts["cocina"]
        tip = (h / cocina_total[d]) * pool if cocina_total[d] > 0 else 0.0
        rows.append({"date": d, "turno": "Día completo", "name": name, "area": "cocina", "hours": h, "tip": tip})

    # 4c) Redondeo exacto (mayor residuo): cada grupo suma lo mismo que su pozo, en pesos enteros
    groups = defaultdict(list)
    for i, r in enumerate(rows):
        groups[(r["date"], r["turno"], r["area"])].append(i)
    for idxs in groups.values():
        target = int(math.floor(sum(rows[i]["tip"] for i in idxs) + 0.5 + 1e-9))
        floors = [int(math.floor(rows[i]["tip"] + 1e-9)) for i in idxs]
        remainder = target - sum(floors)
        order = sorted(range(len(idxs)),
                       key=lambda k: (-(rows[idxs[k]]["tip"] - floors[k]), rows[idxs[k]]["name"].casefold()))
        for k, i in enumerate(idxs):
            rows[i]["tip"] = floors[k]
        for k in order[:max(remainder, 0)]:
            rows[idxs[k]]["tip"] += 1

    rows.sort(key=lambda r: (r["date"], TURNO_ORDER[r["turno"]], AREA_ORDER.index(r["area"]), r["name"].casefold()))

    # 5) Avisos por dinero sin repartir
    unassigned = 0.0
    dates_present = sorted({d for (d, _t, _n) in hours} | {d for (d, _n) in cocina_day})
    for d in dates_present:
        wd = d.weekday()
        label = f"{DAYS[wd]} {fmt_date(d)}"
        for turno in (["Único"] if modes[wd] else ["Día", "Noche"]):
            pool = pool_of(d, turno)
            for area in ("atencion", "barra"):
                share = pool * pcts[area]
                if share > 0 and area_hours[(d, turno)][area] == 0:
                    unassigned += share
                    where = "ese día" if turno == "Único" else f"el turno {turno}"
                    warnings.append(f"{label}: nadie de {area} trabajó {where}, quedan {money(share)} sin repartir.")
        share = day_total(wd) * pcts["cocina"]
        if share > 0 and cocina_total[d] == 0:
            unassigned += share
            warnings.append(f"{label}: nadie de cocina trabajó ese día, quedan {money(share)} sin repartir.")

    present_wd = {d.weekday() for d in dates_present}
    for wd in range(7):
        total = day_total(wd)
        if total > 0 and wd not in present_wd:
            unassigned += total
            warnings.append(f"{DAYS[wd]}: ingresaste {money(total)} pero no hay horas registradas. Quedan sin repartir.")

    # 6) Descuento Alfredo y resumen por colaborador
    alfredo_dates = defaultdict(set)
    for r in rows:
        if r["area"] in alfredo_rates and r["date"].weekday() in ALFREDO_WEEKDAYS:
            alfredo_dates[(r["name"], r["area"])].add(r["date"])

    people = {}
    for r in rows:
        p = people.setdefault(r["name"], {"areas": set(), "hours": 0.0, "tip": 0})
        p["areas"].add(r["area"])
        p["hours"] += r["hours"]
        p["tip"] += r["tip"]

    repl_names = {e["name"] for e in extra_hours or []}
    summary = []
    for name, p in people.items():
        days = sum(len(v) for (n, _a), v in alfredo_dates.items() if n == name)
        alfredo = sum(alfredo_rates[a] * len(v) for (n, a), v in alfredo_dates.items() if n == name)
        tip = int(p["tip"])
        areas_sorted = sorted(p["areas"], key=AREA_ORDER.index)
        summary.append({
            "name": name, "area": " / ".join(areas_sorted), "_order": AREA_ORDER.index(areas_sorted[0]),
            "hours": p["hours"], "adj_hours": adj_applied.get(name, 0.0), "tip": tip,
            "alfredo_days": days, "alfredo": alfredo, "total": tip - alfredo,
            "replacement": name in repl_names,
        })
    summary.sort(key=lambda x: (x["_order"], x["name"].casefold()))

    if repl_names:
        warnings.insert(0, f"Se incluyeron {len(repl_names)} reemplazo{'s' if len(repl_names) != 1 else ''} sin marcaje: "
                           + ", ".join(sorted(repl_names, key=str.casefold)) + ".")
    if n_adj:
        net_txt = f"{sum(adj_applied.values()):+.2f}".replace(".", ",")
        warnings.insert(0, f"Se aplicaron {n_adj} ajuste{'s' if n_adj != 1 else ''} manual{'es' if n_adj != 1 else ''} de horas ({net_txt} h netas).")

    return {
        "alfredo_rates": dict(alfredo_rates),
        "summary": summary,
        "alfredo_total": sum(x["alfredo"] for x in summary),
        "rows": rows,
        "warnings": warnings,
        "unassigned": unassigned,
        "collected": sum(day_total(wd) for wd in range(7)),
        "distributed": sum(r["tip"] for r in rows),
        "total_hours": sum(r["hours"] for r in rows),
        "people": len({r["name"] for r in rows}),
    }


def results_dataframe(result: dict) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Área": r["area"],
                "Día": f"{DAYS[r['date'].weekday()]} {fmt_date(r['date'])}",
                "Turno": r["turno"],
                "Nombre": r["name"],
                "Horas Trabajadas": round(r["hours"], 2),
                "Propina ($)": int(round(r["tip"])),
            }
            for r in result["rows"]
        ]
    )


SUMMARY_COLS = ["Nombre", "Área", "Horas Trabajadas", "Ajuste manual (h)", "Propina ($)",
                "Asistencias (L-S)", "Alfredo", "Total menos alfredo"]


def summary_dataframe(result: dict) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Nombre": s["name"],
                "Área": s["area"],
                "Horas Trabajadas": round(s["hours"], 2),
                "Ajuste manual (h)": round(s["adj_hours"], 2),
                "Propina ($)": s["tip"],
                "Asistencias (L-S)": s["alfredo_days"],
                "Alfredo": s["alfredo"],
                "Total menos alfredo": s["total"],
            }
            for s in result["summary"]
        ],
        columns=SUMMARY_COLS,
    )


WEEK_COLS = ["Día", "Área", "Horas Trabajadas", "Propina ($)", "Alfredo", "Neto del día ($)"]


def person_week_dataframe(result: dict, name: str) -> pd.DataFrame:
    by_date = {}
    for r in result["rows"]:
        if r["name"] != name:
            continue
        b = by_date.setdefault(r["date"], {"area": r["area"], "hours": 0.0, "tip": 0})
        b["hours"] += r["hours"]
        b["tip"] += r["tip"]
    out = []
    for d in sorted(by_date):
        b = by_date[d]
        alfredo = result["alfredo_rates"].get(b["area"], 0) if d.weekday() in ALFREDO_WEEKDAYS else 0
        tip = int(round(b["tip"]))
        out.append({
            "Día": f"{DAYS[d.weekday()]} {fmt_date(d)}", "Área": b["area"],
            "Horas Trabajadas": round(b["hours"], 2), "Propina ($)": tip,
            "Alfredo": alfredo, "Neto del día ($)": tip - alfredo,
        })
    return pd.DataFrame(out, columns=WEEK_COLS)


def signed_hours(v) -> str:
    return "-" if not v else f"{v:+.2f}".replace(".", ",")


def _neutralize_formulas(ws) -> None:
    """Evita inyección de fórmulas: un texto que empieza con = + - @ se guarda como texto, no como fórmula."""
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value[:1] in ("=", "+", "-", "@"):
                c.data_type = "s"


def to_excel_bytes(df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    total = pd.DataFrame(
        [{"Área": "Total", "Día": "", "Turno": "", "Nombre": "",
          "Horas Trabajadas": round(df["Horas Trabajadas"].sum(), 2), "Propina ($)": int(df["Propina ($)"].sum())}]
    )
    s_total = pd.DataFrame(
        [{"Nombre": "Total", "Área": "",
          "Horas Trabajadas": round(summary_df["Horas Trabajadas"].sum(), 2),
          "Ajuste manual (h)": round(summary_df["Ajuste manual (h)"].sum(), 2),
          "Propina ($)": int(summary_df["Propina ($)"].sum()),
          "Asistencias (L-S)": int(summary_df["Asistencias (L-S)"].sum()),
          "Alfredo": int(summary_df["Alfredo"].sum()),
          "Total menos alfredo": int(summary_df["Total menos alfredo"].sum())}]
    )
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.concat([summary_df, s_total], ignore_index=True).to_excel(writer, index=False, sheet_name="Resumen")
        for col, width in zip("ABCDEFGH", [26, 16, 18, 18, 14, 18, 10, 20]):
            writer.sheets["Resumen"].column_dimensions[col].width = width
        pd.concat([df, total], ignore_index=True).to_excel(writer, index=False, sheet_name="Distribución")
        for col, width in zip("ABCDEF", [10, 20, 14, 26, 18, 14]):
            writer.sheets["Distribución"].column_dimensions[col].width = width
        for ws in writer.sheets.values():
            _neutralize_formulas(ws)
    return out.getvalue()


# ============================================================================
# ENVÍO DEL RESUMEN SEMANAL POR CORREO
# ============================================================================
def get_smtp_config():
    cfg = _secrets_section("smtp")
    if not cfg:
        return None
    cfg = dict(cfg)
    if not (cfg.get("host") and cfg.get("username") and cfg.get("password")):
        return None
    return {
        "host": str(cfg["host"]),
        "port": int(cfg.get("port", 587)),
        "username": str(cfg["username"]),
        "password": str(cfg["password"]),
        "from_email": str(cfg.get("from_email") or cfg["username"]),
        "use_ssl": bool(cfg.get("use_ssl", False)),
    }


def send_email(cfg: dict, to_addr: str, subject: str, body: str) -> None:
    msg = EmailMessage()
    msg["From"] = cfg["from_email"]
    msg["To"] = to_addr
    msg["Subject"] = " ".join(subject.split())
    msg.set_content(body)
    context = ssl.create_default_context()
    if cfg["use_ssl"]:
        with smtplib.SMTP_SSL(cfg["host"], cfg["port"], timeout=20, context=context) as server:
            server.login(cfg["username"], cfg["password"])
            server.send_message(msg)
    else:
        with smtplib.SMTP(cfg["host"], cfg["port"], timeout=20) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(cfg["username"], cfg["password"])
            server.send_message(msg)


def build_email_body(result: dict, srec: dict, wdf: pd.DataFrame, signer: str) -> str:
    dates = [r["date"] for r in result["rows"]]
    week = f"{fmt_date(min(dates))} al {fmt_date(max(dates))}"
    lines = [f"Hola {srec['name']},", "", f"Este es el resumen de tus propinas de la semana del {week}.", "",
             "Detalle por día:"]
    for r in wdf.to_dict("records"):
        lines.append(
            f"- {r['Día']} | {r['Área']} | {hours_fmt(r['Horas Trabajadas'])} h | "
            f"Propina {money(r['Propina ($)'])} | Alfredo {money(r['Alfredo'])} | Neto {money(r['Neto del día ($)'])}"
        )
    horas = f"- Horas trabajadas: {hours_fmt(srec['hours'])} h"
    if srec["adj_hours"]:
        horas += f" (incluye un ajuste manual de {signed_hours(srec['adj_hours'])} h)"
    lines += [
        "", "Resumen:", horas,
        f"- Propina: {money(srec['tip'])}",
        f"- Alfredo: {money(srec['alfredo'])} ({srec['alfredo_days']} asistencias de lunes a sábado)",
        f"- Total menos alfredo: {money(srec['total'])}",
        "", "Saludos,", signer,
    ]
    return "\n".join(lines)


def render_email_panel(result: dict, srec: dict, wdf: pd.DataFrame):
    person = srec["name"]
    auth = st.session_state.get("auth") or {}
    signer = auth.get("name") if auth.get("role") == "sucursal" else "Administración"
    registered = str(st.session_state.get("emails", {}).get(person, "") or "").strip()
    cfg = get_smtp_config()

    st.markdown("**Enviar resumen por correo**")
    # La clave incluye el correo registrado: si cambia en la tabla de trabajadores, el campo se actualiza.
    if cfg is not None:
        if registered:
            st.text_input("Correo del colaborador (cargado en la tabla de trabajadores)", value=registered,
                          disabled=True, key=f"mail_to_{person}_{registered}")
        else:
            st.warning("Este colaborador no tiene correo. Agrégalo en la tabla de trabajadores para poder enviarle el resumen.")
        to_addr = registered
    else:
        to_addr = st.text_input("Correo del colaborador", value=registered, key=f"mail_to_{person}_{registered}",
                                placeholder="nombre@correo.cl").strip()

    subject = st.text_input("Asunto", value=f"Resumen semanal de propinas - {person}", key=f"mail_subject_{person}")
    default_body = build_email_body(result, srec, wdf, signer)
    tag = hashlib.sha256(default_body.encode("utf-8")).hexdigest()[:8]
    body = st.text_area("Mensaje (editable)", value=default_body, height=320, key=f"mail_body_{person}_{tag}")

    if cfg is None:
        st.info('El envío real no está configurado (falta la sección [smtp] en Secrets): el botón simula el envío.')

    valid = bool(EMAIL_RE.match(to_addr))
    b1, b2 = st.columns(2)
    with b1:
        send = st.button("Enviar correo" if cfg else "Simular envío", key=f"mail_send_{person}", type="primary")
    with b2:
        if valid:
            url = f"mailto:{quote(to_addr)}?subject={quote(subject)}&body={quote(body)}"
            if len(url) < 1900:
                st.link_button("Abrir en mi programa de correo", url)
            else:
                st.caption("El mensaje es muy largo para abrirlo en el programa de correo.")

    if send:
        user_id = auth.get("username", "?")
        if not valid:
            st.error("Escribe un correo válido." if cfg is None else "Este colaborador no tiene un correo válido cargado.")
        elif not subject.strip() or not body.strip():
            st.error("El asunto y el mensaje no pueden estar vacíos.")
        elif cfg is None:
            st.success(f"Simulación: el correo para {to_addr} NO se envió realmente (SMTP no configurado).")
        elif to_addr.lower() != registered.lower():
            st.error("Solo se puede enviar al correo cargado en la tabla de trabajadores.")
        elif not email_quota_ok(user_id):
            st.error("Alcanzaste el límite de correos por hora. Intenta más tarde.")
        else:
            try:
                send_email(cfg, to_addr, subject, body)
                register_email(user_id)
                st.success(f"Correo enviado a {to_addr}.")
            except (smtplib.SMTPException, OSError, ValueError) as e:
                st.error(f"No se pudo enviar el correo: {e}")


# ============================================================================
# HERRAMIENTA: INTERFAZ
# ============================================================================
def describe_week(week: list) -> str:
    parts, i = [], 0
    while i < 7:
        j = i
        while j + 1 < 7 and week[j + 1] == week[i]:
            j += 1
        parts.append(f"{DAYS[i]}: {week[i]}" if i == j else f"{DAYS[i]} a {DAYS[j]}: {week[i]}")
        i = j + 1
    return "; ".join(parts)


def expand_replacements(reps: list) -> list:
    out = []
    for rp in reps:
        for d, h in sorted(rp["hours_by_date"].items()):
            if h > 0:
                out.append({"name": rp["name"], "date": d, "turno": rp["turno"], "hours": h})
    return out


def file_dates(att: pd.DataFrame) -> list:
    ds = sorted({to_date(x) for x in att["Fecha"]})
    if not ds:
        return []
    last = min(ds[-1], ds[0] + timedelta(days=30))
    return [ds[0] + timedelta(days=i) for i in range((last - ds[0]).days + 1)]


def add_replacement(taken: set):
    ss = st.session_state
    name = " ".join(str(ss.get("rep_name") or "").split())
    area = ss.get("rep_area")
    days = sorted(ss.get("rep_dates") or [])
    hours_by_date = {d: float(ss.get(f"rep_h_{d.isoformat()}") or 0.0) for d in days}
    mail = (ss.get("rep_email") or "").strip()
    reps = ss.setdefault("replacements", [])

    def fail(msg):
        ss["rep_msg"] = ("error", msg)

    if not name:
        return fail("Escribe el nombre del reemplazo.")
    if len(name) > 60:
        return fail("El nombre es demasiado largo (máximo 60 caracteres).")
    if name.casefold() in taken or any(r["name"].casefold() == name.casefold() for r in reps):
        return fail("Ya existe un trabajador con ese nombre. Usa otro nombre o el ajuste manual de horas.")
    if area not in REPLACEMENT_AREAS:
        return fail("Elige el área del reemplazo.")
    if not days:
        return fail("Elige al menos un día trabajado.")
    missing = [d for d, h in hours_by_date.items() if h <= 0]
    if missing:
        return fail("Indica las horas trabajadas de cada día. Faltan: "
                    + ", ".join(f"{DAYS[d.weekday()]} {d:%d-%m}" for d in missing) + ".")
    if any(h > 24 for h in hours_by_date.values()):
        return fail("Un día no puede tener más de 24 horas.")
    if mail and not EMAIL_RE.match(mail):
        return fail("El correo no es válido.")

    reps.append({"name": name, "area": area, "hours_by_date": hours_by_date,
                 "turno": ss.get("rep_turno", "Día"), "email": mail})
    for d in days:
        ss.pop(f"rep_h_{d.isoformat()}", None)
    ss["rep_name"], ss["rep_dates"], ss["rep_email"] = "", [], ""
    ss["rep_msg"] = ("ok", f"Reemplazo \"{name}\" agregado.")


def remove_replacement(i: int):
    lst = st.session_state.get("replacements", [])
    if 0 <= i < len(lst):
        lst.pop(i)


def render_replacements(att, taken: set, cutoff):
    st.caption(
        "Registra a quienes trabajaron sin marcar en el reloj: indica nombre, área, días trabajados y horas totales. "
        "Participan en el reparto de propinas y en el descuento Alfredo como cualquier colaborador. "
        "Al elegir los días aparece un campo para escribir las horas exactas de cada uno."
    )
    if att is None:
        st.info("Sube un archivo para registrar reemplazos: los días se eligen entre las fechas del archivo.")
        return

    msg = st.session_state.pop("rep_msg", None)
    if msg:
        (st.success if msg[0] == "ok" else st.error)(msg[1])

    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Nombre del reemplazo", key="rep_name", max_chars=60)
    with c2:
        st.selectbox("Área asignada", REPLACEMENT_AREAS, key="rep_area")
    st.multiselect("Días trabajados", file_dates(att), key="rep_dates",
                   format_func=lambda d: f"{DAYS[d.weekday()]} {fmt_date(d)}")
    selected = sorted(st.session_state.get("rep_dates") or [])
    if selected:
        st.markdown("**Horas trabajadas por día**")
        hour_cols = st.columns(4)
        for i, d in enumerate(selected):
            with hour_cols[i % 4]:
                st.number_input(f"{DAYS[d.weekday()]} {d:%d-%m}", min_value=0.0, max_value=24.0, step=0.25,
                                value=0.0, format="%.2f", key=f"rep_h_{d.isoformat()}")
        total = sum(float(st.session_state.get(f"rep_h_{d.isoformat()}") or 0.0) for d in selected)
        st.caption(f"Total: {hours_fmt(total)} h en {len(selected)} día(s).")
    else:
        st.caption("Elige los días trabajados para ingresar las horas de cada uno.")

    c3, c4 = st.columns(2)
    with c3:
        st.selectbox("Turno", ["Día", "Noche"], key="rep_turno",
                     help=f"Se usa en los días en Modo Turnos (Día: antes de las {cutoff:%H:%M}; Noche: desde esa hora). "
                          "En Pozo Único y en cocina se ignora.")
    with c4:
        st.text_input("Correo (opcional)", key="rep_email", help="Para enviarle su resumen semanal.")
    st.button("Agregar reemplazo", on_click=add_replacement, args=(taken,))

    reps = st.session_state.get("replacements", [])
    if reps:
        st.markdown("**Reemplazos cargados**")
        for i, rp in enumerate(reps):
            col_t, col_b = st.columns([5, 1])
            days_txt = ", ".join(
                f"{DAYS[d.weekday()][:3]} {d:%d-%m}: {hours_fmt(h)} h" for d, h in sorted(rp["hours_by_date"].items())
            )
            col_t.markdown(
                f"{rp['name']}: {rp['area']}, {hours_fmt(sum(rp['hours_by_date'].values()))} h en "
                f"{len(rp['hours_by_date'])} día(s) ({days_txt})"
            )
            col_b.button("Quitar", key=f"rep_del_{i}", on_click=remove_replacement, args=(i,))


def fmt_pct(v: float) -> str:
    return f"{v:.2f}".rstrip("0").rstrip(".").replace(".", ",")


def save_my_pcts():
    ss = st.session_state
    auth = ss.get("auth") or {}
    vals = {a: float(ss.get(f"pct_{a}") or 0.0) for a, _ in PCT_AREAS}
    if auth.get("role") != "sucursal":
        return
    if round(sum(vals.values()), 6) != 100.0:
        ss["pct_msg"] = ("error", "Los porcentajes deben sumar exactamente 100% para guardarlos.")
        return
    err = set_pcts(auth["username"], vals)
    ss["pct_msg"] = ("error", err) if err else ("ok", "Porcentajes guardados como los de tu sucursal.")


def render_pct_config(user: dict):
    ss = st.session_state
    if not ss.get("pct_ready"):
        saved = get_branch_pcts(user["username"]) if user["role"] == "sucursal" else dict(DEFAULT_PCTS)
        for a, _ in PCT_AREAS:
            ss[f"pct_{a}"] = float(saved[a])
        ss["pct_ready"] = True

    st.markdown("**Porcentajes de reparto por área**")
    st.caption(
        "Define qué parte del pozo recibe cada área. Deben sumar exactamente 100%. "
        "Si tu local no tiene un área, déjala en 0%."
    )
    cols = st.columns(3)
    for col, (a, label) in zip(cols, PCT_AREAS):
        with col:
            st.number_input(f"{label} (%)", min_value=0.0, max_value=100.0, step=0.5, format="%.2f", key=f"pct_{a}")

    vals = {a: float(ss.get(f"pct_{a}") or 0.0) for a, _ in PCT_AREAS}
    total = round(sum(vals.values()), 6)
    if total == 100.0:
        st.success("Los porcentajes suman 100%.")
    else:
        diff = round(100.0 - total, 6)
        st.warning(
            f"Los porcentajes suman {fmt_pct(total)}% ({'faltan' if diff > 0 else 'sobran'} {fmt_pct(abs(diff))} puntos). "
            "Deben sumar exactamente 100%: el cálculo queda bloqueado hasta corregirlo."
        )

    if user["role"] == "sucursal":
        st.button("Guardar como porcentajes de mi sucursal", on_click=save_my_pcts, disabled=(total != 100.0))
    else:
        st.caption("Estos porcentajes valen solo para esta sesión. Para fijarlos por sucursal usa el Panel de Control de Acceso.")
    msg = ss.pop("pct_msg", None)
    if msg:
        (st.success if msg[0] == "ok" else st.error)(msg[1])
    return vals, total


def save_my_alfredo():
    ss = st.session_state
    auth = ss.get("auth") or {}
    if auth.get("role") != "sucursal":
        return
    vals = {a: int(ss.get(f"alf_{a}") or 0) for a, _ in ALFREDO_AREAS}
    err = set_alfredo(auth["username"], vals)
    ss["alf_msg"] = ("error", err) if err else ("ok", "Montos guardados como los de tu sucursal.")


def render_alfredo_config(user: dict) -> dict:
    ss = st.session_state
    if not ss.get("alf_ready"):
        saved = get_branch_alfredo(user["username"]) if user["role"] == "sucursal" else dict(DEFAULT_ALFREDO)
        for a, _ in ALFREDO_AREAS:
            ss[f"alf_{a}"] = int(saved[a])
        ss["alf_ready"] = True

    with st.expander("Aporte del bachero (Alfredo)"):
        st.caption(
            "Monto que se descuenta a cada persona por cada día que asistió de lunes a sábado, según su área ese día. "
            "Se multiplica por los días reales asistidos y se resta en la columna Alfredo. "
            "Cocina no aporta. Puedes subir o bajar los montos."
        )
        cols = st.columns(2)
        for col, (a, label) in zip(cols, ALFREDO_AREAS):
            with col:
                st.number_input(f"{label}: aporte por asistencia ($)", min_value=0, max_value=100000,
                                step=100, key=f"alf_{a}")
        if user["role"] == "sucursal":
            st.button("Guardar como montos de mi sucursal", on_click=save_my_alfredo)
        else:
            st.caption("Estos montos valen solo para esta sesión. Para fijarlos por sucursal usa el Panel de Control de Acceso.")
        msg = ss.pop("alf_msg", None)
        if msg:
            (st.success if msg[0] == "ok" else st.error)(msg[1])
    return {a: int(ss.get(f"alf_{a}") or 0) for a, _ in ALFREDO_AREAS}


def fmt_minutes(mins: int) -> str:
    m = abs(int(mins))
    return f"{'+' if mins > 0 else '-'}{m // 60} h {m % 60:02d} min"


def add_adjustment(areas: dict, modes: list):
    ss = st.session_state
    name, d = ss.get("adj_name"), ss.get("adj_date")
    mins = int(ss.get("adj_h", 0)) * 60 + int(ss.get("adj_m", 0))
    if not name or d is None:
        ss["adj_msg"] = ("error", "Elige un colaborador y una fecha.")
        return
    if mins == 0:
        ss["adj_msg"] = ("error", "Indica las horas o los minutos a ajustar.")
        return
    wd = d.weekday()
    area = areas.get(name, ["Omitir"] * 7)[wd]
    if area == "Omitir":
        ss["adj_msg"] = ("error", f"{name} no tiene un área asignada el {DAYS[wd].lower()} (está en Omitir).")
        return
    if modes[wd]:
        turno = "Único"
    elif area == "cocina":
        turno = "Día completo"
    else:
        turno = ss.get("adj_turno", "Día")
    if ss.get("adj_sign") == "Restar":
        mins = -mins
    ss.setdefault("adjustments", []).append(
        {"name": name, "date": d, "turno": turno, "minutes": mins, "note": (ss.get("adj_note") or "").strip()}
    )
    ss["adj_h"], ss["adj_m"], ss["adj_note"] = 0, 0, ""
    ss["adj_msg"] = ("ok", "Ajuste agregado.")


def remove_adjustment(i: int):
    lst = st.session_state.get("adjustments", [])
    if 0 <= i < len(lst):
        lst.pop(i)


def render_adjustments(areas: dict, modes: list, cutoff):
    att = st.session_state["attendance"]
    names = sorted(areas, key=str.casefold)
    dates = sorted({to_date(x) for x in att["Fecha"]})
    st.caption(
        "Suma o resta tiempo a un colaborador en un día concreto: por ejemplo, alguien que se fue sin marcar "
        "o que se quedó después de su marca. El ajuste se suma a sus horas antes de repartir las propinas."
    )

    msg = st.session_state.pop("adj_msg", None)
    if msg:
        (st.success if msg[0] == "ok" else st.error)(msg[1])

    c1, c2 = st.columns(2)
    with c1:
        st.selectbox("Colaborador (escribe para buscar)", names, index=None,
                     placeholder="Buscar colaborador…", key="adj_name")
    with c2:
        st.selectbox("Fecha", dates, format_func=lambda d: f"{DAYS[d.weekday()]} {fmt_date(d)}", key="adj_date")

    name, d = st.session_state.get("adj_name"), st.session_state.get("adj_date")
    if name and d is not None:
        wd = d.weekday()
        area = areas.get(name, ["Omitir"] * 7)[wd]
        if area == "Omitir":
            st.warning(f"{name} está en Omitir el {DAYS[wd].lower()}: no se pueden ajustar sus horas ese día.")
        elif modes[wd]:
            st.caption("Ese día está en Pozo Único: el ajuste se suma al día completo.")
        elif area == "cocina":
            st.caption("Cocina se calcula por día entero: el ajuste se suma al total del día.")
        else:
            st.selectbox("Turno", ["Día", "Noche"], key="adj_turno",
                         help=f"Día: antes de las {cutoff:%H:%M}. Noche: desde esa hora.")

    c3, c4, c5 = st.columns(3)
    with c3:
        st.radio("Operación", ["Sumar", "Restar"], horizontal=True, key="adj_sign")
    with c4:
        st.number_input("Horas", min_value=0, max_value=24, step=1, value=0, key="adj_h")
    with c5:
        st.number_input("Minutos", min_value=0, max_value=59, step=5, value=0, key="adj_m")
    st.text_input("Motivo (opcional)", key="adj_note")
    st.button("Agregar ajuste", on_click=add_adjustment, args=(areas, modes))

    adjs = st.session_state.get("adjustments", [])
    if adjs:
        st.markdown("**Ajustes cargados**")
        for i, a in enumerate(adjs):
            col_t, col_b = st.columns([5, 1])
            turno_txt = f", turno {a['turno']}" if a["turno"] in ("Día", "Noche") else ""
            note_txt = f" ({a['note']})" if a["note"] else ""
            col_t.markdown(
                f"{a['name']}: {DAYS[a['date'].weekday()]} {fmt_date(a['date'])}{turno_txt}, "
                f"**{fmt_minutes(a['minutes'])}**{note_txt}"
            )
            col_b.button("Quitar", key=f"adj_del_{i}", on_click=remove_adjustment, args=(i,))


def reset_tool_state():
    for k in ("file_sig", "file_info", "attendance", "fixes", "adjustments", "replacements", "emails", "workers_df",
              "result", "result_sig", "workers_editor", "fix_editor", "person_view"):
        st.session_state.pop(k, None)
    for k in [k for k in st.session_state if str(k).startswith(("mixdays_", "adj_", "rep_", "mail_"))]:
        st.session_state.pop(k, None)


def render_tool(user: dict):
    role_label = "ADMINISTRACIÓN" if user["role"] == "admin" else f"SUCURSAL · {user['name']}"
    render_hero(
        "CONTROL FINANCIERO · RESTAURANTES",
        "Filtrado y cálculo de propinas",
        "Importa asistencia, configura turnos y áreas, registra los pozos y revisa una distribución trazable.",
        role_label,
    )

    with st.container(border=True, key="upload-card"):
        render_section_heading("01", "Cargar asistencia", "Importa el Excel semanal y revisa las marcas que necesitan corrección.")
        st.caption("El archivo debe tener las columnas Nombre, Fecha y Hora, con una fila por cada marca. Cada persona puede tener hasta 4 marcas por día (Entrada 1, Salida 1, Entrada 2, Salida 2). Carga una sola semana.")
        uploaded = st.file_uploader("Archivo Excel", type=["xlsx", "xls"], key="uploader")

        if uploaded is None:
            if "file_sig" in st.session_state:
                reset_tool_state()
        else:
            uploaded_bytes = uploaded.getvalue()
            sig = f"{uploaded.name}:{hashlib.sha256(uploaded_bytes).hexdigest()}"
            if st.session_state.get("file_sig") != sig:
                reset_tool_state()
                try:
                    marks, empties, skipped, sheet = read_marks(uploaded_bytes)
                    if not marks and not empties:
                        raise ValueError("El archivo no tiene marcas válidas (revisa los formatos de Fecha y Hora).")
                    attendance = build_attendance(marks, empties)
                    names = sorted(attendance["Nombre"].unique(), key=str.casefold)
                    extras = []
                    if empties:
                        extras.append(f"{len(empties)} con la hora vacía")
                    if skipped:
                        extras.append(f"{skipped} ignoradas por nombre o fecha inválidos")
                    st.session_state.update(
                        file_sig=sig,
                        attendance=attendance,
                        workers_df=pd.DataFrame({
                            "Nombre": names,
                            "Área Principal": ["atencion"] * len(names),
                            "Área mixta": [False] * len(names),
                            "Segunda Área": pd.Series([None] * len(names), dtype=object),
                            "Correo": pd.Series([None] * len(names), dtype=object),
                        }),
                        file_info=f'{len(marks) + len(empties)} filas leídas de la hoja "{sheet}"'
                        + (f" ({'; '.join(extras)})." if extras else "."),
                    )
                except Exception as e:  # noqa: BLE001
                    st.error(f"No se pudo leer el archivo: {e}")
            if "file_info" in st.session_state and "attendance" in st.session_state:
                st.caption(st.session_state["file_info"])
                render_missing_marks()

    render_workflow_progress()

    with st.container(border=True, key="setup-card"):
        render_section_heading("02", "Configurar turnos y equipo", "Define el modo de cada día, el reparto por áreas y las excepciones de asistencia.")
        all_single = all(st.session_state.get(f"single_{i}", False) for i in range(7))
        cutoff = st.time_input("Hora de corte de turno", value=dtime(17, 30), key="cutoff",
                               step=300, disabled=all_single,
                               help="En los días en Modo Turnos: antes de esta hora es turno Día; desde esta hora, turno Noche.")

        st.markdown("**Modo de cada día**")
        st.caption(
            "Desactivado = Modo Turnos (pozo de Día y pozo de Noche, según la hora de corte). "
            "Activado = Modo Pozo Único (un solo pozo para todo el día, ideal para festivos)."
        )
        mode_cols = st.columns(4)
        for i, day in enumerate(DAYS):
            with mode_cols[i % 4]:
                st.toggle(f"{day}: Pozo Único", value=False, key=f"single_{i}")
        modes = [bool(st.session_state.get(f"single_{i}", False)) for i in range(7)]

        pcts_vals, pcts_total = render_pct_config(user)
        pcts_ok = pcts_total == 100.0

        alfredo_vals = render_alfredo_config(user)

        st.markdown("**Trabajadores encontrados**")
        areas = {}
        emails = {}
        mixed_errors = []
        if "workers_df" not in st.session_state:
            st.info("Sube un archivo para ver aquí a los trabajadores y asignarles un área.")
        else:
            st.caption(
                "Elige el Área Principal de cada persona. Si trabaja en dos áreas según el día, marca Área mixta, "
                "elige su Segunda Área y luego indica abajo en qué días trabaja en ella."
            )
            edited = st.data_editor(
                st.session_state["workers_df"],
                key="workers_editor",
                hide_index=True,
                width="stretch",
                disabled=["Nombre"],
                column_config={
                    "Nombre": st.column_config.TextColumn("Nombre"),
                    "Área Principal": st.column_config.SelectboxColumn("Área Principal", options=AREAS, required=True),
                    "Área mixta": st.column_config.CheckboxColumn("Área mixta", help="Trabaja en dos áreas según el día de la semana."),
                    "Segunda Área": st.column_config.SelectboxColumn(
                        "Segunda Área", options=AREAS, required=False, help="Solo se usa si Área mixta está marcada."),
                    "Correo": st.column_config.TextColumn("Correo", help="Opcional: para enviarle su resumen semanal."),
                },
            )

            for k in [k for k in st.session_state if str(k).startswith("mixdays_")]:
                st.session_state[k] = st.session_state[k]

            records = edited.to_dict("records")
            if any(bool(r["Área mixta"]) for r in records):
                st.markdown("**Días de la Segunda Área**")
                st.caption("Selecciona los días en que cada persona trabaja en su Segunda Área. Los demás días cuentan en su Área Principal.")

            for r in records:
                name, main = r["Nombre"], r["Área Principal"]
                mail = r.get("Correo")
                if not is_missing(mail) and str(mail).strip():
                    emails[name] = str(mail).strip()
                if not bool(r["Área mixta"]):
                    areas[name] = [main] * 7
                    continue
                second = r["Segunda Área"]
                if is_missing(second):
                    mixed_errors.append(name)
                    second = main
                second_days = st.multiselect(
                    f"{name}: días en Segunda Área ({'sin elegir' if name in mixed_errors else second}); resto: {main}",
                    DAYS,
                    key=f"mixdays_{name}",
                )
                areas[name] = [second if DAYS[i] in second_days else main for i in range(7)]
                st.caption("Semana de " + name + " → " + describe_week(areas[name]))

            if mixed_errors:
                st.warning("Elige la Segunda Área de: " + ", ".join(mixed_errors) + ".")

        taken = {str(n).casefold() for n in areas}
        with st.expander("Reemplazos y personal sin marcaje (opcional)",
                         expanded=bool(st.session_state.get("replacements"))):
            render_replacements(st.session_state.get("attendance"), taken, cutoff)
        replacements = list(st.session_state.get("replacements", []))
        for rp in replacements:
            areas[rp["name"]] = [rp["area"]] * 7
            if rp.get("email"):
                emails[rp["name"]] = rp["email"]
        extra_hours = expand_replacements(replacements)
        st.session_state["emails"] = emails

        zero_areas = [a for a, _ in PCT_AREAS if pcts_vals[a] == 0]
        if zero_areas and areas:
            affected = sorted({n for n, week in areas.items() if any(x in zero_areas for x in week)}, key=str.casefold)
            if affected:
                st.warning(
                    f"Hay personas en un área con 0% ({', '.join(zero_areas)}): no recibirán propina por ese tiempo. "
                    "Revisa su área o los porcentajes. Personas: " + ", ".join(affected) + "."
                )

        if st.session_state.get("attendance") is not None and areas:
            with st.expander("Ajuste manual de horas (excepciones y turnos extra)",
                             expanded=bool(st.session_state.get("adjustments"))):
                render_adjustments(areas, modes, cutoff)
        adjustments = list(st.session_state.get("adjustments", []))

    with st.container(border=True, key="tips-card"):
        render_section_heading("03", "Registrar propinas recaudadas", "Ingresa el total de cada pozo según el modo operativo de cada día.")
        for k in [f"tip_{m}_{i}" for m in "udn" for i in range(7)]:
            if k in st.session_state:
                st.session_state[k] = st.session_state[k]

        def tip_input(label, key):
            return st.number_input(label, min_value=0, step=1000, value=0, format="%d", key=key)

        st.caption("Cada día muestra los campos que corresponden a su modo. Abre un día para ingresar sus pozos.")
        for i, day in enumerate(DAYS):
            with st.expander(day, expanded=(i == 0)):
                if modes[i]:
                    st.caption("Modo Pozo Único: un solo pozo para todo el día.")
                    tip_input("Pozo del día", f"tip_u_{i}")
                else:
                    st.caption(f"Modo Turnos: Día antes de las {cutoff:%H:%M}; Noche desde esa hora (incluye lo trabajado después de medianoche).")
                    a, b = st.columns(2)
                    with a:
                        tip_input("Turno Día", f"tip_d_{i}")
                    with b:
                        tip_input("Turno Noche", f"tip_n_{i}")
                if modes[i]:
                    day_sum = int(st.session_state.get(f"tip_u_{i}", 0))
                else:
                    day_sum = int(st.session_state.get(f"tip_d_{i}", 0)) + int(st.session_state.get(f"tip_n_{i}", 0))
                st.caption(f"Total del día: {money(day_sum)}")

        tips = {
            "unico": [int(st.session_state.get(f"tip_u_{i}", 0)) for i in range(7)],
            "dia": [int(st.session_state.get(f"tip_d_{i}", 0)) for i in range(7)],
            "noche": [int(st.session_state.get(f"tip_n_{i}", 0)) for i in range(7)],
        }
        week_total = sum(
            tips["unico"][i] if modes[i] else tips["dia"][i] + tips["noche"][i] for i in range(7)
        )
        st.markdown(f"Total semana: **{money(week_total)}**")

    with st.container(border=True, key="calculate-card"):
        render_section_heading("04", "Calcular y revisar distribución", "El cálculo respeta las áreas, los turnos, los ajustes manuales y los aportes configurados.")
        st.caption(
            f"Atención ({fmt_pct(pcts_vals['atencion'])}%) y barra ({fmt_pct(pcts_vals['barra'])}%) se reparten por turno "
            "en los días en Modo Turnos, o por el día entero en Pozo Único. "
            f"Cocina ({fmt_pct(pcts_vals['cocina'])}%) siempre se reparte por día entero, sobre el total recaudado ese día."
        )

        signature = json.dumps(
            {"f": st.session_state.get("file_sig"), "m": modes, "c": str(cutoff), "a": areas, "t": tips, "p": pcts_vals, "q": alfredo_vals,
             "j": [(a["name"], str(a["date"]), a["turno"], a["minutes"]) for a in adjustments],
             "r": [(rp["name"], rp["area"], sorted((str(d), h) for d, h in rp["hours_by_date"].items()), rp["turno"])
                   for rp in replacements],
             "x": sorted((str(k), v["s"], v["drop"]) for k, v in (st.session_state.get("fixes") or {}).items())},
            sort_keys=True, ensure_ascii=False,
        )

        if st.button("Procesar y Ver Resultados", type="primary"):
            try:
                cut_sec = cutoff.hour * 3600 + cutoff.minute * 60 + cutoff.second
                attendance = st.session_state.get("attendance")
                if attendance is None:
                    raise ValueError("Primero sube el archivo Excel con las horas.")
                fixes = st.session_state.get("fixes") or {}
                final = apply_corrections(attendance, fixes)
                pairs, pending = resolve_pairs(final)
                if pending:
                    raise ValueError(
                        f"Quedan {len(pending)} registros con marcas faltantes sin resolver. "
                        "Complétalos en la tabla de registros con marcas faltantes (o márcalos como Descartar) y guarda las correcciones."
                    )
                if not pcts_ok:
                    raise ValueError(
                        f"Los porcentajes de reparto suman {fmt_pct(pcts_total)}%: deben sumar exactamente 100% para calcular."
                    )
                if mixed_errors:
                    raise ValueError("Elige la Segunda Área de: " + ", ".join(mixed_errors) + ".")
                pcts = {a: v / 100 for a, v in pcts_vals.items()}
                result = compute(pairs, areas, modes, cut_sec, tips, pcts, alfredo_vals, adjustments, extra_hours)
                incomplete_keys = [
                    (r["Nombre"], r["Fecha"])
                    for r in attendance[review_mask(attendance)].to_dict("records")
                ]
                n_drop = sum(1 for k in incomplete_keys if fixes.get(k, {}).get("drop"))
                n_fix = sum(1 for k in incomplete_keys if k in fixes and not fixes[k]["drop"])
                if n_fix or n_drop:
                    result["warnings"].insert(0, f"Se integraron {n_fix} días corregidos a mano y se descartaron {n_drop}.")
                st.session_state["result"] = result
                st.session_state["result_sig"] = signature
            except ValueError as e:
                st.session_state.pop("result", None)
                st.error(str(e))

    result = st.session_state.get("result")
    if result:
        render_results(result, stale=st.session_state.get("result_sig") != signature)


def render_results(result: dict, stale: bool):
    with st.container(border=True, key="results-card"):
        st.divider()
        render_section_heading("✓", "Resultados y distribución", "Consulta los totales, revisa cada colaborador y exporta el detalle a Excel.")

        if stale:
            st.warning('Cambiaste datos después de calcular. Presiona "Procesar y Ver Resultados" para actualizar la tabla.')
        for w in result["warnings"]:
            st.warning(w)

        k1, k2, k3 = st.columns(3)
        k1.metric("Total Propinas Recaudadas ($)", money(result["collected"]))
        k2.metric("Total Horas Trabajadas", hours_fmt(result["total_hours"]))
        k3.metric("Personal Procesado", result["people"])
        if result["unassigned"] > 0:
            st.caption(f"{money(result['distributed'])} repartidos; {money(result['unassigned'])} quedaron sin repartir.")

        df = results_dataframe(result)
        sdf = summary_dataframe(result)
        query = norm(st.text_input("Buscar por nombre, área, día o turno", key="search"))
        view, sview = df, sdf
        if query:
            view = df[df.apply(lambda r: query in norm(" ".join(str(v) for v in r.values)), axis=1)]
            sview = sdf[sdf.apply(lambda r: query in norm(f"{r['Nombre']} {r['Área']}"), axis=1)]

        st.markdown(f"**Resumen por colaborador** ({len(sview)} personas)")
        st.caption(
            "Alfredo: se descuentan " + " y ".join(
                f"{money(result['alfredo_rates'][a])} en {a}" for a, _ in ALFREDO_AREAS) +
            " por cada día de asistencia de lunes a sábado (atencion = garzón/garzona). "
            "Total menos alfredo = Propina − Alfredo."
        )
        s_styled = sview.style.format({
            "Horas Trabajadas": hours_fmt, "Ajuste manual (h)": signed_hours,
            "Propina ($)": money, "Alfredo": money, "Total menos alfredo": money,
        })
        st.dataframe(s_styled, hide_index=True, width="stretch")
        st.markdown(
            f"**Total{' (filas filtradas)' if query else ''}:** propinas {money(sview['Propina ($)'].sum())}, "
            f"Alfredo {money(sview['Alfredo'].sum())}, total menos alfredo {money(sview['Total menos alfredo'].sum())}"
        )

        st.markdown("**Resumen semanal por colaborador**")
        names = [x["name"] for x in result["summary"]]
        if st.session_state.get("person_view") not in names:
            st.session_state.pop("person_view", None)
        person = st.selectbox("Colaborador (escribe para buscar)", names, index=None,
                              placeholder="Elige un colaborador para ver su semana…", key="person_view")
        if person:
            srec = next(x for x in result["summary"] if x["name"] == person)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Horas trabajadas", hours_fmt(srec["hours"]))
            m2.metric("Propina ($)", money(srec["tip"]))
            m3.metric("Alfredo", money(srec["alfredo"]))
            m4.metric("Total menos alfredo", money(srec["total"]))
            extra = f"Área: {srec['area']}. Asistencias de lunes a sábado: {srec['alfredo_days']}."
            if srec["adj_hours"]:
                extra += f" Ajuste manual incluido en las horas: {signed_hours(srec['adj_hours'])} h."
            if srec.get("replacement"):
                extra += " Reemplazo sin marcaje (horas cargadas a mano)."
            st.caption(extra)
            wdf = person_week_dataframe(result, person)
            st.dataframe(
                wdf.style.format({"Horas Trabajadas": hours_fmt, "Propina ($)": money, "Alfredo": money, "Neto del día ($)": money}),
                hide_index=True, width="stretch",
            )
            render_email_panel(result, srec, wdf)

        st.markdown(f"**Detalle por día y turno** ({len(view)} filas)")
        styled = view.style.format({"Horas Trabajadas": hours_fmt, "Propina ($)": money})
        st.dataframe(styled, hide_index=True, width="stretch")
        st.markdown(
            f"**Total{' (filas filtradas)' if query else ''}:** {hours_fmt(view['Horas Trabajadas'].sum())} horas, "
            f"{money(view['Propina ($)'].sum())}"
        )

        st.download_button(
            "Descargar copia en Excel (.xlsx)",
            data=to_excel_bytes(df, sdf),
            file_name="distribucion_propinas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ============================================================================
# PANEL DE CONTROL DE ACCESO (SOLO ADMIN)
# ============================================================================
def render_admin_panel():
    users = load_users()
    render_hero(
        "ADMINISTRACIÓN · ACCESOS",
        "Panel de control de acceso",
        "Gestiona sucursales, permisos, credenciales y parámetros de reparto desde un solo lugar.",
        "SOLO ADMINISTRACIÓN",
    )
    active_count = sum(1 for rec in users.values() if rec.get("active"))
    total_count = len(users)
    inactive_count = total_count - active_count
    metric_cols = st.columns(3)
    metric_cols[0].metric("Sucursales registradas", total_count)
    metric_cols[1].metric("Accesos activos", active_count)
    metric_cols[2].metric("Accesos desactivados", inactive_count)
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    flash = st.session_state.pop("flash", None)
    if flash:
        st.success(flash)

    st.info(
        "Los cambios hechos aquí se guardan en un archivo del servidor. En Streamlit Cloud ese archivo se borra "
        "cuando la app se reinicia o se vuelve a desplegar: para hacerlos permanentes, usa "
        '"Generar configuración para st.secrets" al final de esta página.'
    )

    render_section_heading("01", "Crear una sucursal", "Cada sucursal obtiene un usuario propio y parámetros iniciales de seguridad.")
    with st.form("create_branch", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        new_user = c1.text_input("Usuario", help="Minúsculas, números, guion o guion bajo (3 a 32 caracteres).")
        new_name = c2.text_input("Nombre del Local / Sucursal")
        new_pass = c3.text_input("Contraseña inicial", type="password", help=f"Mínimo {MIN_PASSWORD_LEN} caracteres.")
        if st.form_submit_button("Crear sucursal"):
            err = create_branch(new_user, new_name, new_pass)
            if err:
                st.error(err)
            else:
                st.session_state["flash"] = f'Sucursal "{new_name.strip()}" creada.'
                st.rerun()

    render_section_heading("02", "Sucursales registradas", "Estado de acceso y origen de la configuración de cada cuenta.")
    if not users:
        st.info("Todavía no hay sucursales. Crea la primera con el formulario de arriba.")
    else:
        table = pd.DataFrame(
            [
                {"Usuario": u, "Sucursal": r["name"], "Estado": "Activa" if r["active"] else "Desactivada", "Origen": r["source"]}
                for u, r in sorted(users.items())
            ]
        )
        st.dataframe(table, hide_index=True, width="stretch")

        render_section_heading("03", "Gestionar una sucursal", "Actualiza credenciales, estado de acceso y reglas de reparto.")
        selected = st.selectbox(
            "Sucursal", options=sorted(users), format_func=lambda u: f"{users[u]['name']} ({u})", key="manage_select"
        )
        rec = users[selected]
        a, b = st.columns(2)

        with a:
            st.markdown(f"**Estado actual:** {'Activa' if rec['active'] else 'Desactivada'}")
            label = "Desactivar sucursal" if rec["active"] else "Activar sucursal"
            if st.button(label, key=f"toggle_{selected}"):
                err = set_active(selected, not rec["active"])
                if err:
                    st.error(err)
                else:
                    st.session_state["flash"] = (
                        f'Sucursal "{rec["name"]}" {"desactivada: perderá el acceso de inmediato" if rec["active"] else "activada"}.'
                    )
                    st.rerun()

        with b:
            with st.form(f"pw_form_{selected}", clear_on_submit=True):
                st.markdown("**Cambiar contraseña**")
                p1 = st.text_input("Nueva contraseña", type="password")
                p2 = st.text_input("Repetir contraseña", type="password")
                if st.form_submit_button("Cambiar contraseña"):
                    if p1 != p2:
                        st.error("Las contraseñas no coinciden.")
                    else:
                        err = set_password(selected, p1)
                        if err:
                            st.error(err)
                        else:
                            st.session_state["flash"] = f'Contraseña de "{rec["name"]}" actualizada.'
                            st.rerun()

        st.markdown("**Porcentajes de reparto de la sucursal**")
        cur = branch_pcts(rec)
        with st.form(f"pct_form_{selected}"):
            pc = st.columns(3)
            new_p = {}
            for col, (area, label) in zip(pc, PCT_AREAS):
                with col:
                    new_p[area] = st.number_input(f"{label} (%)", min_value=0.0, max_value=100.0, step=0.5,
                                                  value=float(cur[area]), format="%.2f", key=f"adm_pct_{selected}_{area}")
            if st.form_submit_button("Guardar porcentajes"):
                total = round(sum(new_p.values()), 6)
                if total != 100.0:
                    st.error(f"Los porcentajes suman {fmt_pct(total)}%. Deben sumar exactamente 100%.")
                else:
                    err = set_pcts(selected, new_p)
                    if err:
                        st.error(err)
                    else:
                        st.session_state["flash"] = f'Porcentajes de "{rec["name"]}" guardados.'
                        st.rerun()

        st.markdown("**Aporte del bachero (Alfredo) de la sucursal**")
        cur_alf = branch_alfredo(rec)
        with st.form(f"alf_form_{selected}"):
            ac = st.columns(2)
            new_alf = {}
            for col, (area, label) in zip(ac, ALFREDO_AREAS):
                with col:
                    new_alf[area] = st.number_input(f"{label}: aporte por asistencia ($)", min_value=0, max_value=100000,
                                                    step=100, value=int(cur_alf[area]), key=f"adm_alf_{selected}_{area}")
            if st.form_submit_button("Guardar montos"):
                err = set_alfredo(selected, new_alf)
                if err:
                    st.error(err)
                else:
                    st.session_state["flash"] = f'Montos de Alfredo de "{rec["name"]}" guardados.'
                    st.rerun()

        render_section_heading("04", "Hacer los cambios permanentes", "Exporta la configuración cifrada para guardarla en los Secrets de Streamlit.")
        st.caption("Genera el bloque de configuración (con contraseñas ya cifradas) y pégalo en Settings > Secrets de tu app.")
        if st.button("Generar configuración para st.secrets"):
            toml = build_secrets_toml(users)
            st.code(toml, language="toml")
            st.download_button("Descargar como .toml", data=toml, file_name="sucursales_secrets.toml", mime="text/plain")


# ============================================================================
# PUNTO DE ENTRADA
# ============================================================================
def main():
    user = current_user()
    if not user:
        render_login()
        return

    with st.sidebar:
        st.markdown(
            '<div class="brand-lockup"><div class="brand-mark">✦</div>'
            '<div><div class="brand-name">Propinas</div>'
            '<div class="brand-sub">Gestión semanal</div></div></div>',
            unsafe_allow_html=True,
        )
        st.markdown("<div class='eyebrow' style='margin-bottom:8px'>SESIÓN ACTIVA</div>", unsafe_allow_html=True)
        st.markdown(f"**{html.escape(user['name'])}**")
        st.caption("Administrador" if user["role"] == "admin" else "Cuenta de sucursal")
        st.divider()
        page = "Cálculo de Propinas"
        if user["role"] == "admin":
            page = st.radio("Menú", ["Cálculo de Propinas", "Panel de Control de Acceso"], key="page")
        if st.button("Cerrar sesión"):
            logout()

    if page == "Panel de Control de Acceso" and user["role"] == "admin":
        render_admin_panel()
    else:
        render_tool(user)


if __name__ == "__main__":
    main()
