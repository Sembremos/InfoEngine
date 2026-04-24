import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io
import re

from processors.comunidad import procesar_comunidad
from processors.comercio import procesar_comercio
from processors.estadistica import procesar_estadistica
from processors.lineas_accion import procesar_lineas_accion
from processors.micmac import ui_micmac, escribir_cuadrantes_manual, clasificar_y_escribir_riesgos_delitos
from processors.pareto import procesar_pareto
from processors.triangulo import procesar_triangulo
from processors.region import escribir_region
from processors.micmac_datos import MicMac_Datos


st.title("Generador de SS-ENGINE")

# -----------------------------
# ESTILOS
# -----------------------------
def titulo_seccion(texto, color):
    st.markdown(
        f"""
        <div style="
            font-size:20px;
            font-weight:bold;
            color:white;
            background-color:{color};
            padding:8px;
            border-radius:5px;
            margin-top:10px;
        ">
            {texto}
        </div>
        """,
        unsafe_allow_html=True
    )


# -----------------------------
# CARGA DE ARCHIVOS
# -----------------------------
titulo_seccion("Comunidad", "#1f77b4")
archivo_comunidad = st.file_uploader("", type=["xlsx"], key="comunidad")

titulo_seccion("Comercio", "#ff7f0e")
archivo_comercio = st.file_uploader("", type=["xlsx"], key="comercio")

titulo_seccion("Estadística", "#2ca02c")
archivo_estadistica = st.file_uploader("", type=["xlsx"], key="estadistica")

titulo_seccion("Líneas de Acción", "#d62728")
archivo_lineas = st.file_uploader("", type=["xlsx"], key="lineas")

titulo_seccion("Pareto", "#9467bd")
archivo_pareto = st.file_uploader("", type=["xlsx"], key="pareto")

titulo_seccion("Triángulo", "#8c564b")
archivo_triangulo = st.file_uploader("", type=["xlsx"], key="triangulo")



# -----------------------------
# MICMAC
# -----------------------------
titulo_seccion("MICMAC", "#17becf")
poder, conflicto, resultados, autonomas = ui_micmac()

# NUEVO MICMAC EXCEL
titulo_seccion("MICMAC EXCEL", "#0e9aa7")
archivo_micmac_excel = st.file_uploader("", type=["xlsx"], key="micmac_excel")


# -----------------------------
# BOTÓN PRINCIPAL
# -----------------------------
if st.button("Generar info_engine"):

    if not archivo_comunidad or not archivo_comercio or not archivo_estadistica or not archivo_lineas or not archivo_pareto or not archivo_triangulo or not archivo_micmac_excel:
        st.error("Debe subir todos los archivos")
    else:
        try:
            df_comunidad = pd.read_excel(archivo_comunidad)
            df_comercio = pd.read_excel(archivo_comercio)

            wb = load_workbook("plantillas/info_engine.xlsx")

            procesar_comunidad(df_comunidad, wb)
            procesar_comercio(df_comercio, wb)
            procesar_estadistica(archivo_estadistica, wb)
            procesar_lineas_accion(archivo_lineas, wb)
            procesar_pareto(archivo_pareto, wb)
            escribir_cuadrantes_manual(wb, poder, conflicto, resultados, autonomas)
            clasificar_y_escribir_riesgos_delitos(wb, poder, conflicto)
            procesar_triangulo(archivo_triangulo, wb)
            escribir_region(wb)
            MicMac_Datos(archivo_micmac_excel, wb)

            # -----------------------------
            # OBTENER NOMBRE DELEGACIÓN
            # -----------------------------
            hoja = wb["Hoja1"]
            delegacion = hoja["B2"].value

            if delegacion:
                # limpiar texto para nombre de archivo
                delegacion_limpia = re.sub(r'[^a-zA-Z0-9_-]', '_', str(delegacion))
            else:
                delegacion_limpia = "sin_nombre"

            nombre_archivo = f"engine_{delegacion_limpia}.xlsx"

            # -----------------------------
            # EXPORTAR
            # -----------------------------
            archivo_salida = io.BytesIO()
            wb.calculation.fullCalcOnLoad = True
            wb.save(archivo_salida)
            archivo_salida.seek(0)

            st.success("Todo Bien con el archivo!!!!!")

            st.download_button(
                label="Descargar info_engine",
                data=archivo_salida,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
