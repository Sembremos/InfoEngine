import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

from processors.comunidad import procesar_comunidad
from processors.comercio import procesar_comercio
from processors.estadistica import procesar_estadistica

st.title("Generador de info_engine")

archivo_comunidad = st.file_uploader("Subir Comunidad", type=["xlsx"])
archivo_comercio = st.file_uploader("Subir Comercio", type=["xlsx"])
archivo_estadistica = st.file_uploader("Subir Estadística", type=["xlsx"])

if st.button("Generar info_engine"):

    if not archivo_comunidad or not archivo_comercio or not archivo_estadistica:
        st.error("Debe subir los tres archivos")
    else:
        try:
            # leer datos
            df_comunidad = pd.read_excel(archivo_comunidad)
            df_comercio = pd.read_excel(archivo_comercio)

            # abrir UNA sola plantilla
            wb = load_workbook("plantillas/info_engine.xlsx")

            # procesar
            procesar_comunidad(df_comunidad, wb)
            procesar_comercio(df_comercio, wb)
            procesar_estadistica(archivo_estadistica, wb)

            # guardar resultado final
            archivo = io.BytesIO()
            wb.save(archivo)
            archivo.seek(0)

            st.download_button(
                "Descargar archivo generado",
                archivo,
                file_name="info_engine_resultado.xlsx"
            )

        except Exception as e:
            st.error(f"Error: {e}")
