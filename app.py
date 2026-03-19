import streamlit as st
from processors.comunidad import procesar_comunidad

st.title("Generador de info_engine")

# Preparado para múltiples archivos en el futuro
tipo_archivo = st.selectbox(
    "Seleccione el tipo de archivo",
    ["Comunidad"]  # luego aquí agregas más
)

archivo = st.file_uploader("Subir archivo Excel", type=["xlsx"])


if archivo:

    if st.button("Generar info_engine"):

        if tipo_archivo == "Comunidad":
            output = procesar_comunidad(archivo).getvalue()

            st.download_button(
                "Descargar archivo generado",
                output,
                file_name="info_engine_resultado.xlsx"
            )
