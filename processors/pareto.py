import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import tempfile


# -----------------------------
# PROCESAR PARETO
# -----------------------------
def procesar_pareto(archivo_pareto, wb_destino):

    # -----------------------------
    # LEER EXCEL (PANDAS)
    # -----------------------------
    df = pd.read_excel(archivo_pareto)

    # normalizar nombres por si acaso
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
    # SEPARAR POR CATEGORIA
    # -----------------------------
    delitos = df_filtrado[df_filtrado[col_categoria].astype(str).str.upper() == "DELITO"]
    riesgos = df_filtrado[df_filtrado[col_categoria].astype(str).str.upper() == "RIESGO SOCIAL"]

    lista_delitos = delitos[col_descriptor].tolist()
    lista_riesgos = riesgos[col_descriptor].tolist()

    # -----------------------------
    # ESCRIBIR EN INFO_ENGINE
    # -----------------------------
    ws = wb_destino["Hoja1"]

    # limpiar rangos
    for fila in range(97, 118):
        ws[f"B{fila}"] = None
        ws[f"C{fila}"] = None

    # escribir delitos (columna B)
    fila = 97
    for item in lista_delitos:
        if fila > 117:
            break
        ws[f"B{fila}"] = item
        fila += 1

    # escribir riesgos (columna C)
    fila = 97
    for item in lista_riesgos:
        if fila > 117:
            break
        ws[f"C{fila}"] = item
        fila += 1

    # -----------------------------
    # COPIAR GRÁFICO
    # -----------------------------
    try:
        wb_pareto = load_workbook(archivo_pareto)
        ws_pareto = wb_pareto.active

        if ws_pareto._charts:
            chart = ws_pareto._charts[0]
            ws.add_chart(chart, "E95")

    except Exception as e:
        print("No se pudo copiar el gráfico:", e)
