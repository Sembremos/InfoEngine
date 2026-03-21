import streamlit as st
from processors.comunidad import procesar_comunidad
from processors.comercio import procesar_comercio

st.title("Generador de info_engine")

archivo_comunidad = st.file_uploader("Subir archivo Comunidad", type=["xlsx"])
archivo_comercio = st.file_uploader("Subir archivo Comercio", type=["xlsx"])

if st.button("Generar info_engine"):

    if not archivo_comunidad or not archivo_comercio:
        st.error("Debe subir ambos archivos")
    else:
        try:
            output_comunidad = procesar_comunidad(archivo_comunidad)
            output_comercio = procesar_comercio(archivo_comercio)

            # usar el último (ambos escriben sobre la misma plantilla)
            st.download_button(
                "Descargar archivo generado",
                output_comercio,
                file_name="info_engine_resultado.xlsx"
            )

        except Exception as e:
            st.error(f"Error: {e}")
