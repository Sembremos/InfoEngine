import streamlit as st
from processors.comunidad import procesar_comunidad
from processors.comercio import procesar_comercio

st.title("Generador de info_engine")

tipo_archivo = st.selectbox(
    "Seleccione el tipo de archivo",
    ["Comunidad", "Comercio"]
)

archivo = st.file_uploader("Subir archivo Excel", type=["xlsx"])

if archivo:

    if st.button("Generar info_engine"):

        if tipo_archivo == "Comunidad":
            output = procesar_comunidad(archivo)

        elif tipo_archivo == "Comercio":
            output = procesar_comercio(archivo)

        if output is not None:
            st.download_button(
                "Descargar archivo generado",
                output,
                file_name="info_engine_resultado.xlsx"
            )
        else:
            st.error("Error generando el archivo")
