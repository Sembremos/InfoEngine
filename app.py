import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

from processors.comunidad import procesar_comunidad
from processors.comercio import procesar_comercio
from processors.estadistica import procesar_estadistica
from processors.lineas_accion import procesar_lineas_accion
from processors.micmac import procesar_micmac


st.title("Generador de info_engine")

# -----------------------------
# CARGA DE ARCHIVOS
# -----------------------------
archivo_comunidad = st.file_uploader("Subir Comunidad", type=["xlsx"])
archivo_comercio = st.file_uploader("Subir Comercio", type=["xlsx"])
archivo_estadistica = st.file_uploader("Subir Estadística", type=["xlsx"])
archivo_lineas = st.file_uploader("Subir Líneas de Acción", type=["xlsx"])
archivo_micmac = st.file_uploader("Subir MICMAC", type=["xlsx"])

# -----------------------------
# BOTÓN PRINCIPAL
# -----------------------------
if st.button("Generar info_engine"):

    # Validación
    if not archivo_comunidad or not archivo_comercio or not archivo_estadistica or not archivo_lineas or not archivo_micmac:
        st.error("Debe subir todos los archivos")
    else:
        try:
            # -----------------------------
            # LECTURA DE ARCHIVOS
            # -----------------------------
            df_comunidad = pd.read_excel(archivo_comunidad)
            df_comercio = pd.read_excel(archivo_comercio)

            # -----------------------------
            # CARGAR PLANTILLA BASE
            # -----------------------------
            wb = load_workbook("plantillas/info_engine.xlsx")

            # -----------------------------
            # PROCESADORES
            # -----------------------------
            procesar_comunidad(df_comunidad, wb)
            procesar_comercio(df_comercio, wb)
            procesar_estadistica(archivo_estadistica, wb)
            procesar_lineas_accion(archivo_lineas, wb)
            procesar_micmac(archivo_micmac, wb)

            # -----------------------------
            # EXPORTAR RESULTADO
            # -----------------------------
            archivo_salida = io.BytesIO()
            wb.save(archivo_salida)
            archivo_salida.seek(0)

            st.success("Archivo generado correctamente")

            st.download_button(
                label="Descargar info_engine",
                data=archivo_salida,
                file_name="info_engine_resultado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
