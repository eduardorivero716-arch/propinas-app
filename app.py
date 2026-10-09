"""
Filtrado y Cálculo de Propinas  -  con login por sucursal y Panel de Control de Acceso
======================================================================================

requirements.txt
----------------
    streamlit>=1.50.0
    pandas>=2.0
    openpyxl>=3.1
    xlrd>=2.0.1        # solo si subes archivos .xls antiguos

.streamlit/secrets.toml  (local)  /  Settings > Secrets  (Streamlit Cloud)
--------------------------------------------------------------------------
    [admin]
    username = "admin"
    name     = "Administrador"
    password = "una-clave-larga-y-unica"      # o bien: password_hash = "pbkdf2_sha256$..."

    # Opcional: sucursales iniciales. También se pueden crear desde el panel.
    [users.centro]
    name          = "Local Centro"
    password_hash = "pbkdf2_sha256$210000$..."   # o: password = "clave-inicial"
    active        = true

Cómo se guardan los accesos
---------------------------
* st.secrets es de SOLO LECTURA mientras la app corre. Por eso:
    - El administrador y las sucursales "semilla" viven en st.secrets.
    - Los cambios hechos desde el panel (crear, activar/desactivar, cambiar clave)
      se guardan en un archivo JSON (users_db.json) que se superpone a los secrets.
      Solo guarda hashes PBKDF2, nunca contraseñas en texto plano.
* En Streamlit Cloud el disco es EFÍMERO: users_db.json se pierde cuando la app se
  reinicia o se vuelve a desplegar. Para que los cambios sean permanentes, usa en el
  panel el botón "Generar configuración para st.secrets" y pega el resultado en
  Settings > Secrets de tu app.
* La clave del administrador solo se cambia editando st.secrets.
* No subas users_db.json ni secrets.toml a GitHub (agrégalos a .gitignore).
"""

import hashlib
import hmac
import io
import json
import math
import numbers
import os
import re
import secrets as pysecrets
import threading
import time
import unicodedata
from collections import defaultdict
from datetime import date, datetime, time as dtime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
st.set_page_config(page_title="Filtrado y Cálculo de Propinas", page_icon="💵", layout="wide")

DB_PATH = Path(os.environ.get("USERS_DB_PATH", "users_db.json"))
_DB_LOCK = threading.Lock()

PBKDF2_ITERATIONS = 210_000
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_SECONDS = 60
MIN_PASSWORD_LEN = 8
USERNAME_RE = re.compile(r"^[a-z0-9_-]{3,32}$")
LOGIN_ERROR = "Acceso no autorizado o usuario desactivado"

DAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
AREAS = ["atencion", "barra", "cocina", "Omitir"]
AREA_PCT = {"atencion": 0.70, "cocina": 0.20, "barra": 0.10}
AREA_ORDER = ["atencion", "cocina", "barra"]
TURNO_ORDER = {"Día": 0, "Noche": 1, "Único": 2}


# ============================================================================
# SEGURIDAD: HASH Y VERIFICACIÓN DE CONTRASEÑAS
# ============================================================================
def hash_password(password: str) -> str:
    salt = pysecrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PBKDF2_ITERATIONS)
    return f"pbkdf2_sha256${PBKDF2_ITERATIONS}${salt.hex()}${dk.hex()}"


def _verify_hash(password: str, stored: str) -> bool:
    try:
        algo, iters, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), int(iters))
        return hmac.compare_digest(dk, bytes.fromhex(hash_hex))
    except (ValueError, TypeError):
        return False


def verify_password(record: dict, password: str) -> bool:
    if record.get("password_hash"):
        return _verify_hash(password, str(record["password_hash"]))
    if record.get("password") is not None:  # contraseña en texto plano definida en st.secrets
        return hmac.compare_digest(str(record["password"]).encode("utf-8"), password.encode("utf-8"))
    return False


_DUMMY_HASH = hash_password("dummy-password")  # iguala tiempos cuando el usuario no existe


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
    users = _secret_users()
    for uname, ov in _read_overlay().items():
        rec = dict(users.get(uname, {"name": uname, "password": None, "password_hash": None, "active": True}))
        rec.update({k: v for k, v in ov.items() if k in ("name", "password_hash", "active")})
        if "password_hash" in ov:
            rec["password"] = None
        users[uname] = rec
    secret_names = set(_secret_users())
    overlay_names = set(_read_overlay())
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
        pw_hash = rec.get("password_hash") or hash_password(str(rec.get("password") or ""))
        lines += [
            f"[users.{uname}]",
            f'name = "{_toml_str(rec["name"])}"',
            f'password_hash = "{pw_hash}"',
            f'active = {"true" if rec["active"] else "false"}',
            "",
        ]
    return "\n".join(lines)


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
        _verify_hash(password, _DUMMY_HASH)
        return None
    if verify_password(rec, password) and rec["active"]:
        return {"username": uname, "name": rec["name"], "role": "sucursal"}
    return None


def current_user():
    """Devuelve la sesión activa. Revalida en cada interacción: si la sucursal fue
    desactivada o eliminada, la sesión se cierra de inmediato."""
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
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        st.title("Filtrado y Cálculo de Propinas")
        st.caption("Inicia sesión con el usuario de tu sucursal.")

        if get_admin() is None:
            st.error(
                "Falta configurar el administrador. Agrega una sección [admin] con username y "
                "password (o password_hash) en .streamlit/secrets.toml o en Settings > Secrets."
            )
            st.stop()

        if st.session_state.pop("login_error", False):
            st.error(LOGIN_ERROR)

        wait = st.session_state.get("lock_until", 0) - time.time()
        with st.form("login_form"):
            username = st.text_input("Nombre de Usuario / Sucursal", autocomplete="username")
            password = st.text_input("Contraseña", type="password", autocomplete="current-password")
            submitted = st.form_submit_button("Iniciar Sesión")

        if submitted:
            if wait > 0:
                st.error(f"Demasiados intentos fallidos. Espera {math.ceil(wait)} segundos.")
                return
            user = authenticate(username, password)
            if user:
                st.session_state.clear()
                st.session_state["auth"] = user
                st.rerun()
            fails = st.session_state.get("fails", 0) + 1
            st.session_state["fails"] = fails
            if fails >= MAX_FAILED_ATTEMPTS:
                st.session_state["lock_until"] = time.time() + LOCKOUT_SECONDS
                st.session_state["fails"] = 0
            st.error(LOGIN_ERROR)


# ============================================================================
# HERRAMIENTA: LECTURA DEL EXCEL
# ============================================================================
def norm(s) -> str:
    s = "" if s is None or (not isinstance(s, str) and pd.isna(s)) else str(s)
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").strip().lower()


def parse_date(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, numbers.Real) and not isinstance(v, bool):
        return date(1899, 12, 30) + timedelta(days=int(math.floor(v)))
    s = str(v).strip()
    try:
        m = re.match(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", s)
        if m:
            return date(int(m[1]), int(m[2]), int(m[3]))
        m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", s)
        if m:
            y = int(m[3]) + (2000 if int(m[3]) < 100 else 0)
            return date(y, int(m[2]), int(m[1]))
    except ValueError:
        return None
    return None


def parse_time(v):
    """Devuelve segundos desde medianoche, o None."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (datetime, dtime)):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, numbers.Real) and not isinstance(v, bool):
        return int(round((v - math.floor(v)) * 86400))
    m = re.search(r"(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([apAP])?\.?[mM]?\.?", str(v).strip())
    if not m:
        return None
    h, mi, sec = int(m[1]), int(m[2]), int(m[3] or 0)
    if m[4]:
        pm = m[4].lower() == "p"
        if pm and h < 12:
            h += 12
        if not pm and h == 12:
            h = 0
    if h > 23 or mi > 59:
        return None
    return h * 3600 + mi * 60 + sec


def read_marks(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    for sheet in xls.sheet_names:
        raw = xls.parse(sheet, header=None, dtype=object)
        for r in range(min(len(raw), 25)):
            cells = [norm(c) for c in raw.iloc[r].tolist()]
            if "nombre" not in cells or "fecha" not in cells:
                continue
            i_n, i_f = cells.index("nombre"), cells.index("fecha")
            i_h = cells.index("hora") if "hora" in cells else None
            marks, skipped = [], 0
            for i in range(r + 1, len(raw)):
                row = raw.iloc[i].tolist()
                name_v = row[i_n]
                name = " ".join(str(name_v).split()) if not pd.isna(name_v) else ""
                if not name:
                    if any(not pd.isna(c) for c in row):
                        skipped += 1
                    continue
                d = parse_date(row[i_f])
                sec = parse_time(row[i_h]) if i_h is not None else None
                if sec is None and isinstance(row[i_f], (datetime, numbers.Real)):
                    sec = parse_time(row[i_f])  # fecha y hora en la misma celda
                if d is None or sec is None:
                    skipped += 1
                    continue
                marks.append((name, d, sec))
            return marks, skipped, sheet
    raise ValueError('No encontré las columnas "Nombre", "Fecha" y "Hora" en el archivo.')


# ============================================================================
# HERRAMIENTA: CÁLCULO
# ============================================================================
def money(n) -> str:
    return "$" + f"{int(round(n)):,}".replace(",", ".")


def hours_fmt(h) -> str:
    return f"{h:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_date(d: date) -> str:
    return d.strftime("%d-%m-%Y")


def hhmm(sec: int) -> str:
    return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}"


def compute(marks, areas: dict, single: bool, cut_sec: int, tips: dict) -> dict:
    active = {n: a for n, a in areas.items() if a != "Omitir"}
    if not marks:
        raise ValueError("Primero sube el archivo Excel con las horas.")
    if not active:
        raise ValueError('Todos los trabajadores están en "Omitir". Asigna un área al menos a uno.')

    warnings = []

    # 1) Marcas por persona y fecha
    by_pd = defaultdict(set)
    for name, d, sec in marks:
        if name in active:
            by_pd[(name, d)].add(sec)

    # 2) Emparejar marcas (entrada, salida) y repartir el tiempo entre turnos
    hours = defaultdict(float)  # (fecha, turno, nombre) -> horas
    for (name, d), times in sorted(by_pd.items(), key=lambda kv: (kv[0][1], kv[0][0])):
        t = sorted(times)
        for i in range(0, len(t) - 1, 2):
            a, b = t[i], t[i + 1]
            if single:
                hours[(d, "Único", name)] += (b - a) / 3600
            else:
                if min(b, cut_sec) - a > 0:
                    hours[(d, "Día", name)] += (min(b, cut_sec) - a) / 3600
                if b - max(a, cut_sec) > 0:
                    hours[(d, "Noche", name)] += (b - max(a, cut_sec)) / 3600
        if len(t) % 2:
            warnings.append(f"{name}: marca sin pareja el {fmt_date(d)} a las {hhmm(t[-1])}. Se ignoró.")
    if not hours:
        raise ValueError("No se pudieron calcular horas. Revisa que cada persona tenga entrada y salida por día.")

    # 3) Horas por (fecha, turno, área)
    area_hours = defaultdict(lambda: {"atencion": 0.0, "cocina": 0.0, "barra": 0.0})
    for (d, turno, name), h in hours.items():
        area_hours[(d, turno)][active[name]] += h

    def pool_of(d, turno):
        key = "unico" if turno == "Único" else ("dia" if turno == "Día" else "noche")
        return tips[key][d.weekday()]

    # 4) Reparto: Propina = (Horas_Persona / Total_Horas_Area) * (Pozo_Turno * Pct_Area)
    rows = []
    for (d, turno, name), h in hours.items():
        area = active[name]
        total_area = area_hours[(d, turno)][area]
        tip = (h / total_area) * (pool_of(d, turno) * AREA_PCT[area]) if total_area > 0 else 0.0
        rows.append({"date": d, "turno": turno, "name": name, "area": area, "hours": h, "tip": tip})
    rows.sort(key=lambda r: (r["date"], TURNO_ORDER[r["turno"]], AREA_ORDER.index(r["area"]), r["name"].casefold()))

    # 5) Avisos por dinero sin repartir
    unassigned = 0.0
    for (d, turno), ah in sorted(area_hours.items(), key=lambda kv: (kv[0][0], TURNO_ORDER[kv[0][1]])):
        pool = pool_of(d, turno)
        for area in AREA_ORDER:
            share = pool * AREA_PCT[area]
            if ah[area] == 0 and share > 0:
                unassigned += share
                warnings.append(
                    f"{DAYS[d.weekday()]} {fmt_date(d)} ({turno}): nadie de {area} trabajó, quedan {money(share)} sin repartir."
                )
    present = {(d.weekday(), turno) for (d, turno) in area_hours}
    for wd in range(7):
        for turno in (["Único"] if single else ["Día", "Noche"]):
            key = "unico" if turno == "Único" else ("dia" if turno == "Día" else "noche")
            pool = tips[key][wd]
            if pool > 0 and (wd, turno) not in present:
                unassigned += pool
                warnings.append(f"{DAYS[wd]} ({turno}): ingresaste {money(pool)} pero no hay horas registradas. Quedan sin repartir.")
    dates_by_wd = defaultdict(set)
    for r in rows:
        dates_by_wd[r["date"].weekday()].add(r["date"])
    for wd, ds in sorted(dates_by_wd.items()):
        if len(ds) > 1:
            warnings.append(
                f"El archivo tiene {len(ds)} fechas distintas de {DAYS[wd]}. Se usa el mismo pozo para cada una; "
                "carga una sola semana para evitar duplicar propinas."
            )

    collected = sum(tips["unico"]) if single else sum(tips["dia"]) + sum(tips["noche"])
    return {
        "rows": rows,
        "warnings": warnings,
        "unassigned": unassigned,
        "collected": collected,
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


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    total = pd.DataFrame(
        [{"Área": "Total", "Día": "", "Turno": "", "Nombre": "",
          "Horas Trabajadas": round(df["Horas Trabajadas"].sum(), 2), "Propina ($)": int(df["Propina ($)"].sum())}]
    )
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.concat([df, total], ignore_index=True).to_excel(writer, index=False, sheet_name="Distribución")
        ws = writer.sheets["Distribución"]
        for col, width in zip("ABCDEF", [10, 20, 8, 26, 18, 14]):
            ws.column_dimensions[col].width = width
    return out.getvalue()


# ============================================================================
# HERRAMIENTA: INTERFAZ
# ============================================================================
def reset_tool_state():
    for k in ("file_sig", "marks", "workers_df", "result", "result_sig", "workers_editor"):
        st.session_state.pop(k, None)


def render_tool(user: dict):
    st.title("Filtrado y Cálculo de Propinas")
    if user["role"] == "sucursal":
        st.caption(f"Sucursal: {user['name']}")

    # ---- Sección 1
    st.subheader("1. Cargar Tus Horas")
    st.caption("El archivo debe tener las columnas Nombre, Fecha y Hora, con una fila por cada marca de entrada o salida.")
    uploaded = st.file_uploader("Archivo Excel", type=["xlsx", "xls"], key="uploader")

    if uploaded is None:
        if "file_sig" in st.session_state:
            reset_tool_state()
    else:
        sig = f"{uploaded.name}-{uploaded.size}"
        if st.session_state.get("file_sig") != sig:
            reset_tool_state()
            try:
                marks, skipped, sheet = read_marks(uploaded.getvalue())
                if not marks:
                    raise ValueError("El archivo no tiene marcas válidas (revisa los formatos de Fecha y Hora).")
                names = sorted({m[0] for m in marks}, key=str.casefold)
                st.session_state.update(
                    file_sig=sig,
                    marks=marks,
                    workers_df=pd.DataFrame({"Nombre": names, "Área": ["atencion"] * len(names)}),
                    file_info=f'{len(marks)} marcas leídas de la hoja "{sheet}"'
                    + (f"; {skipped} filas ignoradas por datos incompletos." if skipped else "."),
                )
            except Exception as e:  # noqa: BLE001
                st.error(f"No se pudo leer el archivo: {e}")
        if "file_info" in st.session_state and "marks" in st.session_state:
            st.caption(st.session_state["file_info"])

    # ---- Sección 2
    st.subheader("2. Configuración de Turnos y Mapeo")
    c1, c2 = st.columns(2)
    with c1:
        single = st.toggle("Modo Pozo Único (Festivos / Turno Único)", value=False, key="single",
                           help="Todas las horas del día se reparten en un solo pozo.")
    with c2:
        cutoff = st.time_input("Hora de corte de turno", value=dtime(17, 30), key="cutoff",
                               step=300, disabled=single,
                               help="Antes de esta hora es turno Día; desde esta hora, turno Noche.")

    st.markdown("**Trabajadores encontrados**")
    areas = {}
    if "workers_df" not in st.session_state:
        st.info("Sube un archivo para ver aquí a los trabajadores y asignarles un área.")
    else:
        edited = st.data_editor(
            st.session_state["workers_df"],
            key="workers_editor",
            hide_index=True,
            width="stretch",
            disabled=["Nombre"],
            column_config={
                "Nombre": st.column_config.TextColumn("Nombre"),
                "Área": st.column_config.SelectboxColumn("Área", options=AREAS, required=True),
            },
        )
        areas = dict(zip(edited["Nombre"], edited["Área"]))

    # ---- Sección 3
    st.subheader("3. Propinas Recaudadas ($)")
    # Streamlit descarta el valor de widgets que dejan de dibujarse (al cambiar de modo); lo conservamos.
    for k in [f"tip_{m}_{i}" for m in "udn" for i in range(7)]:
        if k in st.session_state:
            st.session_state[k] = st.session_state[k]

    def tip_input(label, key):
        return st.number_input(label, min_value=0, step=1000, value=0, format="%d", key=key)

    if single:
        st.caption("Ingresa el total recaudado en cada día. Se reparte entre todos los que trabajaron ese día.")
        cols = st.columns(4)
        for i, day in enumerate(DAYS):
            with cols[i % 4]:
                tip_input(day, f"tip_u_{i}")
    else:
        st.caption("Ingresa lo recaudado en cada turno.")
        for i, day in enumerate(DAYS):
            with st.expander(day, expanded=(i == 0)):
                a, b = st.columns(2)
                with a:
                    tip_input("Turno Día", f"tip_d_{i}")
                with b:
                    tip_input("Turno Noche", f"tip_n_{i}")

    tips = {
        "unico": [int(st.session_state.get(f"tip_u_{i}", 0)) for i in range(7)],
        "dia": [int(st.session_state.get(f"tip_d_{i}", 0)) for i in range(7)],
        "noche": [int(st.session_state.get(f"tip_n_{i}", 0)) for i in range(7)],
    }
    week_total = sum(tips["unico"]) if single else sum(tips["dia"]) + sum(tips["noche"])
    st.markdown(f"Total semana: **{money(week_total)}**")

    # ---- Sección 4
    st.subheader("4. Calcular Propinas")
    st.caption("Reparte 70% a atención, 20% a cocina y 10% a barra, en proporción a las horas de cada persona dentro de su área.")

    signature = json.dumps(
        {"f": st.session_state.get("file_sig"), "s": single, "c": str(cutoff), "a": areas, "t": tips},
        sort_keys=True, ensure_ascii=False,
    )

    if st.button("Procesar y Ver Resultados", type="primary"):
        try:
            cut_sec = cutoff.hour * 3600 + cutoff.minute * 60 + cutoff.second
            st.session_state["result"] = compute(st.session_state.get("marks"), areas, single, cut_sec, tips)
            st.session_state["result_sig"] = signature
        except ValueError as e:
            st.session_state.pop("result", None)
            st.error(str(e))

    # ---- Resultados
    result = st.session_state.get("result")
    if result:
        render_results(result, stale=st.session_state.get("result_sig") != signature)


def render_results(result: dict, stale: bool):
    st.divider()
    st.subheader("Resultados y Distribución")

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
    query = norm(st.text_input("Buscar por nombre, área, día o turno", key="search"))
    view = df
    if query:
        mask = df.apply(lambda r: query in norm(" ".join(str(v) for v in r.values)), axis=1)
        view = df[mask]

    st.markdown(f"**Detalle por día y turno** ({len(view)} filas)")
    styled = view.style.format({"Horas Trabajadas": hours_fmt, "Propina ($)": money})
    st.dataframe(styled, hide_index=True, width="stretch")
    st.markdown(
        f"**Total{' (filas filtradas)' if query else ''}:** {hours_fmt(view['Horas Trabajadas'].sum())} horas, "
        f"{money(view['Propina ($)'].sum())}"
    )

    st.download_button(
        "Descargar copia en Excel (.xlsx)",
        data=to_excel_bytes(df),
        file_name="distribucion_propinas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================================
# PANEL DE CONTROL DE ACCESO (SOLO ADMIN)
# ============================================================================
def render_admin_panel():
    st.title("Panel de Control de Acceso")

    flash = st.session_state.pop("flash", None)
    if flash:
        st.success(flash)

    st.info(
        "Los cambios hechos aquí se guardan en un archivo del servidor. En Streamlit Cloud ese archivo se borra "
        "cuando la app se reinicia o se vuelve a desplegar: para hacerlos permanentes, usa "
        '"Generar configuración para st.secrets" al final de esta página.'
    )

    # ---- Crear sucursal
    st.subheader("Crear nueva sucursal")
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

    # ---- Tabla de sucursales
    st.subheader("Sucursales registradas")
    users = load_users()
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

        # ---- Acciones
        st.subheader("Gestionar una sucursal")
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

        # ---- Exportar para secrets
        st.subheader("Hacer los cambios permanentes")
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
        st.markdown(f"**{user['name']}**")
        st.caption("Administrador" if user["role"] == "admin" else "Sucursal")
        page = "Cálculo de Propinas"
        if user["role"] == "admin":
            page = st.radio("Menú", ["Cálculo de Propinas", "Panel de Control de Acceso"], key="page")
        if st.button("Cerrar sesión"):
            logout()

    if page == "Panel de Control de Acceso" and user["role"] == "admin":  # doble verificación de rol
        render_admin_panel()
    else:
        render_tool(user)


main()
