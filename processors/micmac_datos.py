import pandas as pd
import re


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
    xls = pd.ExcelFile(archivo_micmac)

    hoja_desc = None
    hoja_matriz = None

    for nombre in xls.sheet_names:
        nombre_limpio = nombre.strip().upper()

        if "DESCRIPTOR" in nombre_limpio:
            hoja_desc = nombre

        if "MATRIZ" in nombre_limpio:
            hoja_matriz = nombre

    if hoja_desc is None or hoja_matriz is None:
        raise Exception(f"Hojas no encontradas: {xls.sheet_names}")

    df_desc = pd.read_excel(xls, sheet_name=hoja_desc)
    df_matriz = pd.read_excel(xls, sheet_name=hoja_matriz, header=None)

    hoja_engine = wb["Hoja1"]

    # =============================
    # 1. INSTITUCIONES
    # =============================
    instituciones = []

    col_inicio = 4  # E
    col_fin = 7     # H

    for i in range(9, len(df_matriz)):
        fila = df_matriz.iloc[i, col_inicio:col_fin+1]

        valor = " ".join([str(x) for x in fila if pd.notna(x)]).strip()

        if valor == "":
            break

        instituciones.append(valor)

    instituciones_unicas = list(dict.fromkeys(instituciones))

    for idx, inst in enumerate(instituciones_unicas[:10]):
        hoja_engine[f"B{150 + idx}"] = inst

    # =============================
    # 2. FECHA
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
    # 3. DETECTAR MATRIZ MICMAC
    # =============================
    fila_header = None

    for i in range(len(df_matriz)):
        fila = df_matriz.iloc[i]
        valores = [str(x) for x in fila if pd.notna(x)]

        count_codigos = sum(
            1 for v in valores
            if isinstance(v, str) and "." in v
        )

        if count_codigos >= 3:
            fila_header = i
            break

    if fila_header is None:
        return

    # =============================
    # EXTRAER MATRIZ
    # =============================
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

    matriz = df_matriz.iloc[
        fila_header + 1 : fila_header + 1 + size,
        col_inicio : col_inicio + size
    ]

    problemas_fila = df_matriz.iloc[
        fila_header + 1 : fila_header + 1 + size,
        0
    ].tolist()

    # =============================
    # 4. MAPEO DESCRIPTORES
    # =============================
    mapa_desc = {}

    for i in range(len(df_desc)):
        corto = df_desc.iloc[i, 0]
        largo = df_desc.iloc[i, 1]

        if pd.notna(corto) and pd.notna(largo):
            mapa_desc[str(corto).strip()] = str(largo).strip()

    # =============================
    # 5. PROBLEMAS POR LÍNEA (B,C,D)
    # =============================
    problemas_engine_por_linea = []

    for fila in range(242, 254):
        problemas_linea = []

        for col in ["B", "C", "D"]:
            val = hoja_engine[f"{col}{fila}"].value
            if val:
                problemas_linea.append(val)

        problemas_engine_por_linea.append(
            [normalizar(x) for x in problemas_linea]
        )

    # =============================
    # 6. COLUMNAS DESTINO
    # =============================
    columnas_destino = [
        "G", "M", "S", "Y", "AE", "AK",
        "AQ", "AW", "BC", "BI", "BO", "BU"
    ]

    # =============================
    # 7. PROCESAMIENTO FINAL
    # =============================
    for idx_linea, problemas_linea_norm in enumerate(problemas_engine_por_linea):

        if idx_linea >= len(columnas_destino):
            break

        col_destino = columnas_destino[idx_linea]
        influyentes = []

        # recorrer encabezados de la matriz
        for idx_col, problema_header in enumerate(encabezados):

            problema_largo = mapa_desc.get(problema_header, problema_header)

            if normalizar(problema_largo) not in problemas_linea_norm:
                continue

            # buscar influencias
            for i in range(size):
                valor = matriz.iloc[i, idx_col]

                if valor in [2, 3]:
                    problema_corto = problemas_fila[i]
                    problema_largo_inf = mapa_desc.get(problema_corto, problema_corto)
                    influyentes.append(problema_largo_inf)

        # eliminar duplicados
        influyentes = list(dict.fromkeys(influyentes))

        # escribir máximo 30
        for i, val in enumerate(influyentes[:30]):
            hoja_engine[f"{col_destino}{247 + i}"] = val
