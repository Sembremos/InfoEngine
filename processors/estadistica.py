from openpyxl import load_workbook


def procesar_estadistica(archivo_estadistica, wb_info):
    
    # cargar excel desde streamlit (archivo en memoria)
    wb_est = load_workbook(archivo_estadistica, data_only=True)
    ws_est = wb_est.active

    ws_info = wb_info.active

    # función para copiar rangos
    def copiar_rango(ws_origen, ws_destino, fila_ini_o, col_ini_o, fila_fin_o, col_fin_o,
                     fila_ini_d, col_ini_d):

        for i, fila in enumerate(range(fila_ini_o, fila_fin_o + 1)):
            for j, col in enumerate(range(col_ini_o, col_fin_o + 1)):
                ws_destino.cell(
                    row=fila_ini_d + i,
                    column=col_ini_d + j,
                    value=ws_origen.cell(row=fila, column=col).value
                )

    # 1. denuncias_distrito
    copiar_rango(ws_est, ws_info, 4, 1, 16, 2, 165, 1)

    # 2. horarios
    copiar_rango(ws_est, ws_info, 18, 6, 27, 17, 179, 6)

    # 3. modalidad
    copiar_rango(ws_est, ws_info, 34, 4, 43, 15, 195, 4)

    # 4. dias
    copiar_rango(ws_est, ws_info, 48, 4, 55, 15, 222, 4)
