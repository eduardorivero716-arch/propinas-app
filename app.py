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
TURNO_ORDER = {"Día": 0, "Noche": 1, "Único": 2, "Día completo": 3}


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
            marks, empties, skipped = [], [], 0
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
                    alt = parse_time(row[i_f])  # fecha y hora en la misma celda (0 = solo fecha)
                    sec = alt if alt else None
                if d is None:
                    skipped += 1
                    continue
                if sec is None:
                    empties.append((name, d))  # marca con la hora vacía: se corrige en la tabla de excepciones
                    continue
                marks.append((name, d, sec))
            return marks, empties, skipped, sheet
    raise ValueError('No encontré las columnas "Nombre", "Fecha" y "Hora" en el archivo.')


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
    sec = int(sec)
    return dtime(sec // 3600, (sec % 3600) // 60, sec % 60)


def to_date(v):
    return v.date() if isinstance(v, datetime) else v


def slot_problems(vals) -> list:
    """
    Valida un día con el esquema fijo [Entrada 1, Salida 1, Entrada 2, Salida 2].
    Es válido tener solo la jornada 1 (Entrada 2 y Salida 2 vacías) o las dos jornadas completas.
    """
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


def build_attendance(marks, empties) -> pd.DataFrame:
    """
    DataFrame principal: UNA fila por trabajador y fecha, con 4 columnas fijas de tiempo
    (Entrada 1, Salida 1, Entrada 2, Salida 2), en segundos desde medianoche (o None si falta).

    Las marcas del día se ordenan y se asignan en ese orden a las 4 columnas. Si hay menos de 4,
    las últimas columnas quedan vacías; si hay más de 4, se usan las 4 primeras y se deja una
    Nota para revisión. 'Marcas originales' conserva todas las horas que traía el archivo.
    """
    valid, blank = defaultdict(set), defaultdict(int)
    for name, d, sec in marks:
        valid[(name, d)].add(sec)
    for name, d in empties:
        blank[(name, d)] += 1

    rows = []
    for (name, d) in sorted(set(valid) | set(blank), key=lambda k: (k[1], k[0].casefold())):
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
    """Integra las correcciones guardadas al DataFrame principal (y quita los días descartados)."""
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
    """
    Convierte cada día en sus intervalos [Entrada 1 -> Salida 1] y [Entrada 2 -> Salida 2].
    La jornada 2 se omite si sus dos celdas están vacías. Los días con problemas quedan pendientes.
    """
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
    """Muestra SOLO los días con marcas faltantes o por revisar y permite completarlos."""
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
        "Si una marca quedó en una columna equivocada, muévela. Marca Descartar para ignorar todo ese día."
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
                "s": [parse_time(rec[c]) for c in SLOTS],
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


def hhmm(sec: int) -> str:
    return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}"


def compute(pairs, areas: dict, modes: list, cut_sec: int, tips: dict) -> dict:
    """
    modes[i] = True  -> el día i (0 = Lunes) está en Modo Pozo Único.
    modes[i] = False -> el día i está en Modo Turnos (Día / Noche según la hora de corte).

    Reglas de reparto:
      * atención (70%) y barra (10%): por turno en Modo Turnos; por día entero en Pozo Único.
      * cocina (20%): SIEMPRE por día entero, sobre el pozo total del día (Día + Noche o Pozo Único).
    """
    active = {n: a for n, a in areas.items() if a != "Omitir"}
    if not pairs:
        raise ValueError("No hay tramos completos (entrada y salida) para calcular.")
    if not active:
        raise ValueError('Todos los trabajadores están en "Omitir". Asigna un área al menos a uno.')

    def day_total(wd: int) -> float:
        return tips["unico"][wd] if modes[wd] else tips["dia"][wd] + tips["noche"][wd]

    def pool_of(d: date, turno: str) -> float:
        wd = d.weekday()
        if turno == "Único":
            return tips["unico"][wd]
        return tips["dia"][wd] if turno == "Día" else tips["noche"][wd]

    warnings = []

    # 1) y 2) Cada tramo completo (entrada, salida) se reparte según el modo de su día
    hours = defaultdict(float)  # (fecha, turno, nombre) -> horas
    for name, d, ent, sal in pairs:
        if name not in active:
            continue
        if modes[d.weekday()]:
            hours[(d, "Único", name)] += (sal - ent) / 3600
        else:
            if min(sal, cut_sec) - ent > 0:
                hours[(d, "Día", name)] += (min(sal, cut_sec) - ent) / 3600
            if sal - max(ent, cut_sec) > 0:
                hours[(d, "Noche", name)] += (sal - max(ent, cut_sec)) / 3600
    if not hours:
        raise ValueError("No se pudieron calcular horas. Revisa que haya tramos completos de personas con un área asignada.")

    # 3) Horas por (fecha, turno, área) para atención y barra; horas por día para cocina
    area_hours = defaultdict(lambda: {"atencion": 0.0, "barra": 0.0})
    cocina_day = defaultdict(float)  # (fecha, nombre) -> horas del día completo
    for (d, turno, name), h in hours.items():
        area = active[name]
        if area == "cocina":
            cocina_day[(d, name)] += h
        else:
            area_hours[(d, turno)][area] += h
    cocina_total = defaultdict(float)  # fecha -> horas totales de cocina
    for (d, _name), h in cocina_day.items():
        cocina_total[d] += h

    # 4) Reparto proporcional
    rows = []
    # 4a) Atención y barra: por turno (o por día entero si el día es Pozo Único)
    for (d, turno, name), h in hours.items():
        area = active[name]
        if area == "cocina":
            continue
        total_area = area_hours[(d, turno)][area]
        tip = (h / total_area) * (pool_of(d, turno) * AREA_PCT[area]) if total_area > 0 else 0.0
        rows.append({"date": d, "turno": turno, "name": name, "area": area, "hours": h, "tip": tip})
    # 4b) Cocina: siempre por día entero sobre el pozo total del día
    for (d, name), h in cocina_day.items():
        pool = day_total(d.weekday()) * AREA_PCT["cocina"]
        tip = (h / cocina_total[d]) * pool if cocina_total[d] > 0 else 0.0
        rows.append({"date": d, "turno": "Día completo", "name": name, "area": "cocina", "hours": h, "tip": tip})
    rows.sort(key=lambda r: (r["date"], TURNO_ORDER[r["turno"]], AREA_ORDER.index(r["area"]), r["name"].casefold()))

    # 5) Avisos por dinero sin repartir
    unassigned = 0.0
    dates_present = sorted({d for (d, _t, _n) in hours})
    for d in dates_present:
        wd = d.weekday()
        label = f"{DAYS[wd]} {fmt_date(d)}"
        for turno in (["Único"] if modes[wd] else ["Día", "Noche"]):
            pool = pool_of(d, turno)
            for area in ("atencion", "barra"):
                share = pool * AREA_PCT[area]
                if share > 0 and area_hours[(d, turno)][area] == 0:
                    unassigned += share
                    where = "ese día" if turno == "Único" else f"el turno {turno}"
                    warnings.append(f"{label}: nadie de {area} trabajó {where}, quedan {money(share)} sin repartir.")
        share = day_total(wd) * AREA_PCT["cocina"]
        if share > 0 and cocina_total[d] == 0:
            unassigned += share
            warnings.append(f"{label}: nadie de cocina trabajó ese día, quedan {money(share)} sin repartir.")

    present_wd = {d.weekday() for d in dates_present}
    for wd in range(7):
        total = day_total(wd)
        if total > 0 and wd not in present_wd:
            unassigned += total
            warnings.append(f"{DAYS[wd]}: ingresaste {money(total)} pero no hay horas registradas. Quedan sin repartir.")

    dates_by_wd = defaultdict(set)
    for d in dates_present:
        dates_by_wd[d.weekday()].add(d)
    for wd, ds in sorted(dates_by_wd.items()):
        if len(ds) > 1:
            warnings.append(
                f"El archivo tiene {len(ds)} fechas distintas de {DAYS[wd]}. Se usa el mismo pozo para cada una; "
                "carga una sola semana para evitar duplicar propinas."
            )

    return {
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
    for k in ("file_sig", "attendance", "fixes", "workers_df", "result", "result_sig", "workers_editor", "fix_editor"):
        st.session_state.pop(k, None)


def render_tool(user: dict):
    st.title("Filtrado y Cálculo de Propinas")
    if user["role"] == "sucursal":
        st.caption(f"Sucursal: {user['name']}")

    # ---- Sección 1
    st.subheader("1. Cargar Tus Horas")
    st.caption("El archivo debe tener las columnas Nombre, Fecha y Hora, con una fila por cada marca. Cada persona puede tener hasta 4 marcas por día (Entrada 1, Salida 1, Entrada 2, Salida 2).")
    uploaded = st.file_uploader("Archivo Excel", type=["xlsx", "xls"], key="uploader")

    if uploaded is None:
        if "file_sig" in st.session_state:
            reset_tool_state()
    else:
        sig = f"{uploaded.name}-{uploaded.size}"
        if st.session_state.get("file_sig") != sig:
            reset_tool_state()
            try:
                marks, empties, skipped, sheet = read_marks(uploaded.getvalue())
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
                    workers_df=pd.DataFrame({"Nombre": names, "Área": ["atencion"] * len(names)}),
                    file_info=f'{len(marks) + len(empties)} filas leídas de la hoja "{sheet}"'
                    + (f" ({'; '.join(extras)})." if extras else "."),
                )
            except Exception as e:  # noqa: BLE001
                st.error(f"No se pudo leer el archivo: {e}")
        if "file_info" in st.session_state and "attendance" in st.session_state:
            st.caption(st.session_state["file_info"])
            render_missing_marks()

    # ---- Sección 2
    st.subheader("2. Configuración de Turnos y Mapeo")
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
    # Streamlit descarta el valor de widgets que dejan de dibujarse (al cambiar el modo de un día); lo conservamos.
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
                st.caption(f"Modo Turnos: Día antes de las {cutoff:%H:%M}; Noche desde esa hora.")
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

    # ---- Sección 4
    st.subheader("4. Calcular Propinas")
    st.caption(
        "Atención (70%) y barra (10%) se reparten por turno en los días en Modo Turnos, o por el día entero en Pozo Único. "
        "Cocina (20%) siempre se reparte por día entero, sobre el total recaudado ese día."
    )

    signature = json.dumps(
        {"f": st.session_state.get("file_sig"), "m": modes, "c": str(cutoff), "a": areas, "t": tips,
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
            final = apply_corrections(attendance, fixes)  # DataFrame principal + correcciones
            pairs, pending = resolve_pairs(final)
            if pending:
                raise ValueError(
                    f"Quedan {len(pending)} registros con marcas faltantes sin resolver. "
                    "Complétalos en la tabla de registros con marcas faltantes (o márcalos como Descartar) y guarda las correcciones."
                )
            result = compute(pairs, areas, modes, cut_sec, tips)
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
