import pandas as pd
from openpyxl import load_workbook


# -----------------------------
# PROCESAR PARETO
# -----------------------------
def procesar_pareto(archivo_pareto, wb_destino):

    # -----------------------------
    # LEER EXCEL (PANDAS)
    # -----------------------------
    df = pd.read_excel(archivo_pareto)

    # normalizar columnas
    df.columns = [str(c).strip().upper() for c in df.columns]

    # columnas esperadas
    col_categoria = "CATEGORIA"
    col_descriptor = "DESCRIPTOR"
    col_segmento = "SEGMENTO"

    if col_categoria not in df.columns:
        col_categoria = df.columns[0]
    if col_descriptor not in df.columns:
        col_descriptor = df.columns[1]
    if col_segmento not in df.columns:
        col_segmento = df.columns[6]

    # -----------------------------
    # FILTRAR PRIORIZADOS
    # -----------------------------
    df_filtrado = df[df[col_segmento].astype(str).str.upper() == "PRIORIZADO"]

    # -----------------------------
    # ESCRIBIR TABLA COMPLETA (A-G)
    # -----------------------------
    ws = wb_destino["Hoja1"]

    fila_excel = 2  # empieza en A2

    for _, row in df_filtrado.iterrows():

        if fila_excel > 31:
            break

        # escribir columnas A-G
        ws[f"A{fila_excel}"] = row.iloc[0]  # categoria
        ws[f"B{fila_excel}"] = row.iloc[1]  # descriptor
        ws[f"C{fila_excel}"] = row.iloc[2]  # frecuencia
        ws[f"D{fila_excel}"] = row.iloc[3]  # porcentaje
        ws[f"E{fila_excel}"] = row.iloc[4]  # pct_acum
        ws[f"F{fila_excel}"] = row.iloc[5]  # acumulado
        ws[f"G{fila_excel}"] = row.iloc[6]  # segmento

        fila_excel += 1

    # -----------------------------
    # BUSCAR TOTAL
    # -----------------------------
    df_raw = pd.read_excel(archivo_pareto, header=None)

    total_valor = None

    for i in range(len(df_raw)):
        celda = str(df_raw.iloc[i, 1]).strip().upper()  # columna B

        if celda == "TOTAL:":
            total_valor = df_raw.iloc[i, 2]  # columna C
            break

    if total_valor is not None:
        ws["B88"] = total_valor

    # -----------------------------
    # CONTAR DESCRIPTORES PRIORIZADOS
    # -----------------------------
    cantidad_descriptores = len(df_filtrado)

    ws["B93"] = cantidad_descriptores

    # -----------------------------
    # (TU CÓDIGO ORIGINAL SIGUE)
    # -----------------------------
    # SEPARAR POR CATEGORIA
    delitos = df_filtrado[df_filtrado[col_categoria].astype(str).str.upper() == "DELITO"]
    riesgos = df_filtrado[df_filtrado[col_categoria].astype(str).str.upper() == "RIESGO SOCIAL"]

    lista_delitos = delitos[col_descriptor].tolist()
    lista_riesgos = riesgos[col_descriptor].tolist()

    # limpiar rangos
    for fila in range(97, 118):
        ws[f"B{fila}"] = None
        ws[f"C{fila}"] = None

    # escribir delitos
    fila = 97
    for item in lista_delitos:
        if fila > 117:
            break
        ws[f"B{fila}"] = item
        fila += 1

    # escribir riesgos
    fila = 97
    for item in lista_riesgos:
        if fila > 117:
            break
        ws[f"C{fila}"] = item
        fila += 1
