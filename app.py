import streamlit as st
import pandas as pd
from datetime import datetime, time as dtime, timedelta
import io
import math
import numbers
import unicodedata
import re
from collections import defaultdict

# ============================================================================
# CONFIGURACIÓN
# ============================================================================
st.set_page_config(page_title="Filtrado y Cálculo de Propinas", page_icon="💵", layout="wide")

DAYS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
AREAS = ["atencion", "barra", "cocina", "Omitir"]
AREA_PCT = {"atencion": 0.70, "cocina": 0.20, "barra": 0.10}
AREA_ORDER = ["atencion", "cocina", "barra"]
TURNO_ORDER = {"Día": 0, "Noche": 1, "Único": 2, "Día completo": 3}

# ============================================================================
# LECTURA DEL EXCEL
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
                    alt = parse_time(row[i_f])
                    sec = alt if alt else None
                if d is None:
                    skipped += 1
                    continue
                if sec is None:
                    empties.append((name, d))
                    continue
                marks.append((name, d, sec))
            return marks, empties, skipped, sheet
    raise ValueError('No encontré las columnas "Nombre", "Fecha" y "Hora" en el archivo.')

# ============================================================================
# ASISTENCIA CON 4 MARCAS FIJAS POR DÍA
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
    e1, s1, e2, s2 = [None if is_missing(v) else int(v) for v in vals]
    if e1 is None and s1 is None and e2 is None and s2 is None:
        return ["Sin marcas"]
    out = []
    if e1 is None and s1 is None:
        out.append("Falta la jornada 1 (Entrada 1 y Salida 1)")
    else:
        if e1 is None: out.append("Falta Entrada 1")
        if s1 is None: out.append("Falta Salida 1")
    if e2 is None and s2 is not None: out.append("Falta Entrada 2")
    if s2 is None and e2 is not None: out.append("Falta Salida 2")
    filled = [v for v in (e1, s1, e2, s2) if v is not None]
    if any(b <= a for a, b in zip(filled, filled[1:])):
        out.append("Horas fuera de orden")
    return out

def needs_review(rec: dict) -> bool:
    return bool(slot_problems([rec[c] for c in SLOTS])) or bool(rec.get("Nota"))

def problem_text(rec: dict) -> str:
    parts = slot_problems([rec[c] for c in SLOTS])
    if rec.get("Nota"): parts.append(rec["Nota"])
    return "; ".join(parts)

def build_attendance(marks, empties) -> pd.DataFrame:
    valid, blank = defaultdict(set), defaultdict(int)
    for name, d, sec in marks: valid[(name, d)].add(sec)
    for name, d in empties: blank[(name, d)] += 1

    rows = []
    for (name, d) in sorted(set(valid) | set(blank), key=lambda k: (k[1], k[0].casefold())):
        times = sorted(valid.get((name, d), ()))
        n_blank = blank.get((name, d), 0)
        slots = (times[:4] + [None] * 4)[:4]
        notes = []
        if len(times) > 4: notes.append(f"Hay {len(times)} marcas; solo se cargaron las 4 primeras")
        if n_blank: notes.append(f"{n_blank} fila(s) con hora vacía")
        original = ", ".join(hhmm(t) for t in times) if times else "(ninguna)"
        if n_blank: original += f" (+{n_blank} sin hora)"
        rows.append({"Nombre": name, "Fecha": d, **dict(zip(SLOTS, slots)), "Marcas originales": original, "Nota": "; ".join(notes)})
    return pd.DataFrame(rows, columns=ATT_COLS, dtype=object)

def review_mask(df: pd.DataFrame) -> pd.Series:
    return pd.Series([needs_review(r) for r in df.to_dict("records")], index=df.index, dtype=bool)

def apply_corrections(main: pd.DataFrame, fixes: dict) -> pd.DataFrame:
    out = []
    for rec in main.to_dict("records"):
        if needs_review(rec):
            fx = fixes.get((rec["Nombre"], rec["Fecha"]))
            if fx:
                if fx["drop"]: continue
                rec = {**rec, **dict(zip(SLOTS, fx["s"])), "Nota": ""}
        out.append(rec)
    return pd.DataFrame(out, columns=ATT_COLS, dtype=object)

def resolve_pairs(final: pd.DataFrame):
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
    if att is None or att.empty: return
    flagged = att[review_mask(att)]
    if flagged.empty:
        st.success("Todas las marcas están completas: no hay registros por corregir.")
        return

    st.markdown("**Registros con marcas faltantes**")
    recs = flagged.to_dict("records")
    base = pd.DataFrame({
        "Nombre": [r["Nombre"] for r in recs],
        "Fecha": [r["Fecha"] for r in recs],
        "Marcas originales": [r["Marcas originales"] for r in recs],
        "Problema": [problem_text(r) for r in recs],
        **{c: pd.Series([sec_to_time(r[c]) for r in recs], dtype=object) for c in SLOTS},
        "Descartar": [False] * len(recs),
    })

    with st.form("fix_form"):
        edited = st.data_editor(base, key="fix_editor", hide_index=True, num_rows="fixed", disabled=["Nombre", "Fecha", "Marcas originales", "Problema"])
        saved = st.form_submit_button("Guardar correcciones", type="primary")

    if saved:
        st.session_state["fixes"] = {
            (rec["Nombre"], to_date(rec["Fecha"])): {"s": [parse_time(rec[c]) for c in SLOTS], "drop": bool(rec["Descartar"])}
            for rec in edited.to_dict("records")
        }

# ============================================================================
# CÁLCULO Y REGLAS DE NEGOCIO (INCLUYENDO ALFREDO Y AJUSTES MANUALES)
# ============================================================================
def money(n) -> str: return "$" + f"{int(round(n)):,}".replace(",", ".")
def hours_fmt(h) -> str: return f"{h:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
def fmt_date(d: date) -> str: return d.strftime("%d-%m-%Y")
def hhmm(sec: int) -> str: return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}"

def compute(pairs, areas: dict, modes: list, cut_sec: int, tips: dict, manual_adj: dict, roles_dict: dict):
    def area_of(name: str, d: date) -> str:
        week = areas.get(name)
        return week[d.weekday()] if week else "Omitir"

    if not pairs: raise ValueError("No hay tramos completos para calcular.")

    def day_total(wd: int) -> float:
        return tips["unico"][wd] if modes[wd] else tips["dia"][wd] + tips["noche"][wd]

    def pool_of(d: date, turno: str) -> float:
        wd = d.weekday()
        if turno == "Único": return tips["unico"][wd]
        return tips["dia"][wd] if turno == "Día" else tips["noche"][wd]

    hours = defaultdict(float)
    days_worked_count = defaultdict(set) # Para contar asistencias reales por colaborador (lunes a sábado para Alfredo)

    for name, d, ent, sal in pairs:
        if area_of(name, d) == "Omitir": continue
        days_worked_count[name].add(d)
        if modes[d.weekday()]:
            hours[(d, "Único", name)] += (sal - ent) / 3600
        else:
            if min(sal, cut_sec) - ent > 0: hours[(d, "Día", name)] += (min(sal, cut_sec) - ent) / 3600
            if sal - max(ent, cut_sec) > 0: hours[(d, "Noche", name)] += (sal - max(ent, cut_sec)) / 3600

    # Aplicar ajuste manual de horas por colaborador
    for name, adj_val in manual_adj.items():
        if adj_val != 0.0:
            # Añadimos el ajuste al primer turno/fecha disponible o creamos registro base si no tiene horas
            found = False
            for k in hours:
                if k[2] == name:
                    hours[k] += adj_val
                    found = True
                    break
            if not found and adj_val > 0:
                # Si no tiene tramos pero se le agregaron horas manuales, las asignamos al primer lunes o fecha base
                dummy_date = list(set(d for _, d, _, _ in pairs))[0] if pairs else date.today()
                hours[(dummy_date, "Único", name)] += adj_val

    area_hours = defaultdict(lambda: {"atencion": 0.0, "barra": 0.0})
    cocina_day = defaultdict(float)
    for (d, turno, name), h in hours.items():
        area = area_of(name, d)
        if area == "cocina": cocina_day[(d, name)] += h
        else: area_hours[(d, turno)][area] += h

    cocina_total = defaultdict(float)
    for (d, _name), h in cocina_day.items(): cocina_total[d] += h

    rows = []
    for (d, turno, name), h in hours.items():
        area = area_of(name, d)
        if area == "cocina": continue
        total_area = area_hours[(d, turno)][area]
        tip = (h / total_area) * (pool_of(d, turno) * AREA_PCT[area]) if total_area > 0 else 0.0
        rows.append({"date": d, "turno": turno, "name": name, "area": area, "hours": h, "tip": tip})

    for (d, name), h in cocina_day.items():
        pool = day_total(d.weekday()) * AREA_PCT["cocina"]
        tip = (h / cocina_total[d]) * pool if cocina_total[d] > 0 else 0.0
        rows.append({"date": d, "turno": "Día completo", "name": name, "area": "cocina", "hours": h, "tip": tip})

    # Agrupar por persona para la tabla resumen estilo Google Sheets
    person_summary = defaultdict(lambda: {"hours": 0.0, "tip": 0.0, "role": "", "days_count": 0})
    for r in rows:
        name = r["name"]
        person_summary[name]["hours"] += r["hours"]
        person_summary[name]["tip"] += r["tip"]
        person_summary[name]["role"] = roles_dict.get(name, "Garzona")

    summary_rows = []
    for name, data in person_summary.items():
        role = data["role"]
        # Regla de Alfredo: 500 de lunes a sábado si es Garzón/Garzona, 100 si es Barra
        # Contamos cuántos días de lunes a sábado asistió
        lunes_sabado_asistencias = sum(1 for d in days_worked_count[name] if d.weekday() < 6)
        alfredo_desc = (500 * lunes_sabado_asistencias) if ("Garz" in role or "garz" in role.lower()) else (100 * lunes_sabado_asistencias)
        
        summary_rows.append({
            "Rol": role,
            "Equipo": name,
            "Suma Horas": round(data["hours"], 2),
            "Suma Total Propinas": int(round(data["tip"])),
            "Alfredo": alfredo_desc,
            "Total menos alfredo": int(round(data["tip"])) - alfredo_desc
        })

    return {
        "rows": rows,
        "summary_rows": summary_rows,
        "collected": sum(day_total(wd) for wd in range(7)),
        "distributed": sum(r["tip"] for r in rows),
        "total_hours": sum(r["hours"] for r in rows),
        "people": len(person_summary),
        "warnings": []
    }

def to_excel_bytes(summary_df: pd.DataFrame, detailed_df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Resumen Propinas")
        detailed_df.to_excel(writer, index=False, sheet_name="Detalle Turnos")
    return out.getvalue()

def describe_week(week: list) -> str:
    parts, i = [], 0
    while i < 7:
        j = i
        while j + 1 < 7 and week[j + 1] == week[i]: j += 1
        parts.append(f"{DAYS[i]}: {week[i]}" if i == j else f"{DAYS[i]} a {DAYS[j]}: {week[i]}")
        i = j + 1
    return "; ".join(parts)

def reset_tool_state():
    for k in ("file_sig", "attendance", "fixes", "workers_df", "result", "result_sig"):
        st.session_state.pop(k, None)

# ============================================================================
# INTERFAZ PRINCIPAL STREAMLIT
# ============================================================================
def main():
    st.title("💵 Sistema de Cálculo de Propinas - Mr. Sushi")
    st.markdown("---")

    st.subheader("1. Cargar Archivo de Asistencia")
    uploaded = st.file_uploader("Sube tu archivo de asistencias (Excel)", type=["xlsx", "xls"])

    if uploaded is None:
        if "file_sig" in st.session_state: reset_tool_state()
    else:
        sig = f"{uploaded.name}-{uploaded.size}"
        if st.session_state.get("file_sig") != sig:
            reset_tool_state()
            try:
                marks, empties, skipped, sheet = read_marks(uploaded.getvalue())
                attendance = build_attendance(marks, empties)
                names = sorted(attendance["Nombre"].unique(), key=str.casefold)
                st.session_state.update(
                    file_sig=sig,
                    attendance=attendance,
                    workers_df=pd.DataFrame({
                        "Nombre": names,
                        "Rol / Área Principal": ["Garzona"] * len(names),
                        "Área mixta": [False] * len(names),
                        "Segunda Área": pd.Series([None] * len(names), dtype=object),
                    }),
                    file_info=f"{len(marks) + len(empties)} filas leídas de la hoja '{sheet}'."
                )
            except Exception as e:
                st.error(f"Error al leer archivo: {e}")
        
        if "file_info" in st.session_state:
            st.caption(st.session_state["file_info"])
            render_missing_marks()

    st.subheader("2. Configuración de Turnos, Mapeo y Áreas Mixtas")
    cutoff = st.time_input("Hora de corte de turno", value=dtime(17, 30), step=300)
    
    mode_cols = st.columns(4)
    modes = []
    for i, day in enumerate(DAYS):
        with mode_cols[i % 4]:
            modes.append(st.toggle(f"{day}: Pozo Único", value=False, key=f"single_{i}"))

    areas = {}
    roles_dict = {}
    if "workers_df" in st.session_state:
        edited = st.data_editor(st.session_state["workers_df"], key="workers_editor", hide_index=True, disabled=["Nombre"])
        for r in edited.to_dict("records"):
            name, main = r["Nombre"], r["Rol / Área Principal"]
            roles_dict[name] = main
            if not bool(r["Área mixta"]):
                areas[name] = [main.lower() if main in AREAS else "atencion"] * 7
            else:
                second = str(r["Segunda Área"]).lower() if not is_missing(r["Segunda Área"]) else main.lower()
                second_days = st.multiselect(f"{name}: días en Segunda Área ({second}); resto: {main}", DAYS, key=f"mixdays_{name}")
                m_val = main.lower() if main in AREAS else "atencion"
                s_val = second if second in AREAS else m_val
                areas[name] = [s_val if DAYS[i] in second_days else m_val for i in range(7)]

    st.subheader("3. Buscador y Ajuste Manual de Horas (Opcional)")
    st.markdown("Suma o resta horas de forma manual para colaboradores que se fueron sin marcar o hicieron turnos extra.")
    manual_adjustments = {}
    if "workers_df" in st.session_state:
        with st.expander("Desplegar panel de ajustes de horas manuales"):
            for name in st.session_state["workers_df"]["Nombre"]:
                manual_adjustments[name] = st.number_input(f"Ajuste (hrs) para {name}", value=0.0, step=0.5, format="%.1f", key=f"adj_{name}")

    st.subheader("4. Propinas
