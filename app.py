archivo_lineas = st.file_uploader("Subir Líneas de Acción", type=["xlsx"])

if st.button("Generar info_engine"):

    if not archivo_comunidad or not archivo_comercio or not archivo_estadistica or not archivo_lineas:
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
