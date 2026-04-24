import pandas as pd
import re
from openpyxl import load_workbook


def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).upper().strip()
    texto = re.sub(r'[ÁÀÄÂ]', 'A', texto)
    texto = re.sub(r'[ÉÈËÊ]', 'E', texto)
    texto = re.sub(r'[ÍÌÏÎ]', 'I', texto)
    texto = re.sub(r'[ÓÒÖÔ]', 'O', texto)
    texto = re.sub(r'[ÚÙÜÛ]', 'U', texto)
    return texto


def MicMac_Datos(archivo_micmac, wb):

    # =============================
    # CARGA HOJAS
    # =============================
    df_matriz = pd.read_excel(archivo_micmac, sheet_name="MATRIZ", header=None)
    df_desc = pd.read_excel(archivo_micmac, sheet_name="DESCRIPTORES")

    hoja_engine = wb["Hoja1"]

    # =============================
    # 1. INSTITUCIONES
    # =============================
    instituciones = []

    # header fijo en E8
    col_inicio = 4  # E
    col_fin = 7     # H

    for i in range(9, len(df_matriz)):  # debajo del header
        fila = df_matriz.iloc[i, col_inicio:col_fin+1]

        valor = " ".join([str(x) for x in fila if pd.notna(x)]).strip()

        if valor == "":
            break

        instituciones.append(valor)

    # eliminar duplicados manteniendo orden
    instituciones_unicas = list(dict.fromkeys(instituciones))

    # escribir en excel
    for idx, inst in enumerate(instituciones_unicas[:10]):
        hoja_engine[f"B{150 + idx}"] = inst

    # =============================
    # FECHA
    # =============================
    fecha = None

    for i in range(len(df_matriz)):
        for j in range(len(df_matriz.columns)):
            valor = df_matriz.iloc[i, j]

            if isinstance(valor, str) and "Fecha:" in valor:
                match = re.search(r'Fecha:\s*(.*)', valor)
                if match:
                    fecha = match.group(1).strip()
                    break
        if fecha:
            break

    if fecha:
        for i in range(10):
            hoja_engine[f"C{150 + i}"] = fecha

    # =============================
    # 2. MATRIZ MICMAC
    # =============================

    # detectar fila encabezado (códigos tipo CON.DROG)
    fila_header = None

    for i in range(len(df_matriz)):
        fila = df_matriz.iloc[i]

        valores = [str(x) for x in fila if pd.notna(x)]

        # si detecta varios códigos tipo "XXX.XXX"
        count_codigos = sum(1 for v in valores if re.match(r'^[A-Z]+\.[A-Z]+', v))

        if count_codigos >= 3:
            fila_header = i
            break

    if fila_header is None:
        return  # no rompe el flujo

    # columnas desde B en adelante
    col_inicio = 1

    encabezados = []
    j = col_inicio

    while j < len(df_matriz.columns):
        val = df_matriz.iloc[fila_header, j]
        if pd.isna(val):
            break
        encabezados.append(str(val).strip())
        j += 1

    size = len(encabezados)

    # matriz
    matriz = df_matriz.iloc[fila_header+1:fila_header+1+size, col_inicio:col_inicio+size]

    # columna A (problemas influyentes)
    problemas_fila = df_matriz.iloc[fila_header+1:fila_header+1+size, 0].tolist()

    # =============================
    # MAPEO DESCRIPTORES
    # =============================
    mapa_desc = {}

    for i in range(len(df_desc)):
        corto = df_desc.iloc[i, 0]
        largo = df_desc.iloc[i, 1]

        if pd.notna(corto) and pd.notna(largo):
            mapa_desc[str(corto).strip()] = str(largo).strip()

    # =============================
    # PROBLEMAS DE INFOENGINE
    # =============================
    problemas_engine = []

    for i in range(242, 254):
        val = hoja_engine[f"B{i}"].value
        if val:
            problemas_engine.append(val)

    problemas_engine_norm = [normalizar(x) for x in problemas_engine]

    # =============================
    # COLUMNAS DESTINO
    # =============================
    columnas_destino = [
        "G", "M", "S", "Y", "AE", "AK",
        "AQ", "AW", "BC", "BI", "BO", "BU"
    ]

    # =============================
    # PROCESO PRINCIPAL
    # =============================
    for idx_col, problema_header in enumerate(encabezados):

        if idx_col >= len(columnas_destino):
            break

        # convertir a nombre completo
        problema_largo = mapa_desc.get(problema_header, problema_header)

        if normalizar(problema_largo) not in problemas_engine_norm:
            continue

        col_destino = columnas_destino[idx_col]

        influyentes = []

        for i in range(size):
            valor = matriz.iloc[i, idx_col]

            if valor in [2, 3]:
                problema_corto = problemas_fila[i]
                problema_largo_inf = mapa_desc.get(problema_corto, problema_corto)
                influyentes.append(problema_largo_inf)

        # eliminar duplicados
        influyentes = list(dict.fromkeys(influyentes))

        # escribir max 30
        for i, val in enumerate(influyentes[:30]):
            hoja_engine[f"{col_destino}{247 + i}"] = val
