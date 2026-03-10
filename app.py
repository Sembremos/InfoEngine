import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

st.title("Generador de Diagnóstico")

encuesta_comunidad = st.file_uploader(
    "Subir encuesta comunidad", type=["xlsx"]
)

encuesta_comercio = st.file_uploader(
    "Subir encuesta comercio", type=["xlsx"]
)

if encuesta_comunidad and encuesta_comercio:

    if st.button("Generar diagnóstico"):

        df_comunidad = pd.read_excel(encuesta_comunidad)
        df_comercio = pd.read_excel(encuesta_comercio)

        percepcion = df_comunidad.iloc[:,1].mean()
        robos = df_comercio.iloc[:,1].sum()

        wb = load_workbook("plantillas/diagnostico_base.xlsx")
        ws = wb["Hoja1"]

        ws["C10"] = percepcion
        ws["C11"] = robos

        archivo = io.BytesIO()
        wb.save(archivo)
        archivo.seek(0)

        st.download_button(
            "Descargar diagnóstico",
            archivo,
            file_name="diagnostico.xlsx"
        )
