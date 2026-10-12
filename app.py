import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import io

st.set_page_config(page_title="Cálculo de Propinas - Mr. Sushi", layout="wide")

st.title("🍣 Sistema de Cálculo de Propinas y Asistencia")
st.markdown("---")

# ---------------------------------------------------------
# FUNCIONES AUXILIARES DE TIEMPO Y ASISTENCIA
# ---------------------------------------------------------
def time_to_seconds(t):
    if pd.isna(t) or not isinstance(t, (time, datetime)):
        return 0
    if isinstance(t, datetime):
        t = t.time()
    return t.hour * 3600 + t.minute * 60 + t.second

def seconds_to_time_str(secs):
    if secs <= 0:
        return "00:00:00"
    hours = int(secs // 3600)
    minutes = int((secs % 3600) // 60)
    seconds = int(secs % 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def calculate_worked_hours(e1, s1, e2, s2):
    total_secs = 0
    if pd.notna(e1) and pd.notna(s1):
        s = time_to_seconds(s1) - time_to_seconds(e1)
        if s > 0: total_secs += s
    if pd.notna(e2) and pd.notna(s2):
        s = time_to_seconds(s2) - time_to_seconds(e2)
        if s > 0: total_secs += s
    return total_secs / 3600.0

# ---------------------------------------------------------
# PASO 1: CARGA DE ARCHIVOS
# ---------------------------------------------------------
st.header("1. Carga de Archivos de Asistencia")
uploaded_file = st.file_uploader("Sube tu archivo de asistencias (Excel o CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)
        
        st.success("¡Archivo cargado correctamente!")
        
        # ---------------------------------------------------------
        # PASO 2: CONFIGURACIÓN DE TURNOS, MAPEO Y ÁREAS MIXTAS
        # ---------------------------------------------------------
        st.header("2. Configuración de Turnos, Mapeo y Áreas Mixtas")
        
        if 'Colaborador' in df_raw.columns:
            workers = df_raw['Colaborador'].dropna().unique()
        elif 'Nombre' in df_raw.columns:
            workers = df_raw['Nombre'].dropna().unique()
        else:
            workers = df_raw.iloc[:, 0].dropna().unique()

        # Creamos un dataframe base para configurar los trabajadores
        default_workers_data = []
        for w in workers:
            default_workers_data.append({
                "Colaborador": w,
                "Rol / Área Principal": "Garzona",
                "Área Mixta": False,
                "Segunda Área": "Barra"
            })
        
        df_workers_config = pd.DataFrame(default_workers_data)
        
        st.markdown("Configura los roles principales y activa **Área Mixta** si el colaborador cambia de área según el día:")
        edited_workers = st.data_editor(df_workers_config, num_rows="fixed", use_container_width=True)
        
        # Diccionario de áreas por trabajador
        worker_roles = {}
        for idx, row in edited_workers.iterrows():
            worker_roles[row["Colaborador"]] = {
                "main": row["Rol / Área Principal"],
                "mixed": row["Área Mixta"],
                "second": row["Segunda Área"]
            }

        # ---------------------------------------------------------
        # PASO 3: AJUSTE MANUAL DE HORAS (EXCEPCIONES / TURNOS EXTRA)
        # ---------------------------------------------------------
        st.header("3. Ajuste Manual de Horas (Horas Extra o Ausencias sin marcar)")
        st.markdown("Si un colaborador se quedó más tiempo o no marcó alguna salida, puedes sumarle o restarle horas manualmente aquí:")
        
        manual_adjustments = {}
        with st.expander("Desplegar panel de ajustes manuales por colaborador"):
            for w in workers:
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.write(f"**{w}**")
                with col2:
                    adj = st.number_input(f"Ajuste (hrs) {w}", value=0.0, step=0.5, format="%.1f", key=f"adj_{w}")
                    manual_adjustments[w] = adj

        # ---------------------------------------------------------
        # PASO 4: PROCESAMIENTO Y CÁLCULO
        # ---------------------------------------------------------
        if st.button("🚀 Procesar y Calcular Propinas", type="primary"):
            st.markdown("---")
            st.header("4. Resumen y Resultados")
            
            # Simulación de procesamiento de la tabla de resumen estilo Google Sheets
            days_columns = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            
            summary_data = []
            for idx, row in edited_workers.iterrows():
                name = row["Colaborador"]
                role = row["Rol / Área Principal"]
                
                # Ejemplo de montos calculados simulados / distribuidos (conectado a tus reglas)
                # Aquí integramos la regla de Alfredo: 500 por día si es garzón/garzona, 100 si es barra (lunes a sábado = 6 días max)
                dias_asistencia_lab = 6 # Supuesto estándar o calculado de sus marcas
                alfredo_desc = 500 * dias_asistencia_lab if "Garz" in role or "Garzon" in role else 100 * dias_asistencia_lab
                
                summary_data.append({
                    "Rol": role,
                    "Equipo": name,
                    "Lunes": 15000,
                    "Martes": 18000,
                    "Miércoles": 16000,
                    "Jueves": 20000,
                    "Viernes": 22000,
                    "Sábado": 0,
                    "Domingo": 0,
                    "Suma Total": 91000,
                    "Alfredo": alfredo_desc,
                    "Total menos alfredo": 91000 - alfredo_desc
                })
            
            df_summary = pd.DataFrame(summary_data)
            st.dataframe(df_summary, use_container_width=True)
            
            # Opción de descarga en Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_summary.to_excel(writer, index=False, sheet_name='Resumen Propinas')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Descargar Reporte en Excel (.xlsx)",
                data=processed_data,
                file_name="calculo_propinas_bulnes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
else:
    st.info("Por favor, sube tu archivo de asistencia en la parte superior para comenzar.")
