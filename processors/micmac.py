import pandas as pd

# -----------------------------
# DETECTAR MATRIZ MICMAC (ROBUSTO REAL)
# -----------------------------
def detectar_matriz_micmac(archivo):
    df_raw = pd.read_excel(archivo, sheet_name="MATRIZ", header=None)

    header_row = None

    for i in range(len(df_raw) - 1):

        fila = df_raw.iloc[i]
        fila_siguiente = df_raw.iloc[i + 1]

        # Detectar encabezado:
        # fila actual = texto
        # fila siguiente = números
        if (
            isinstance(fila[1], str) and
            isinstance(fila[2], str) and
            isinstance(fila_siguiente[1], (int, float))
        ):
            header_row = i
            break

    if header_row is None:
        return None

    # -----------------------------
    # VARIABLES (ENCABEZADO)
    # -----------------------------
    variables = df_raw.iloc[header_row, 1:].dropna().tolist()
    size = len(variables)

    # -----------------------------
    # MATRIZ NUMÉRICA
    # -----------------------------
    data = []

    for j in range(size):
        fila_data = df_raw.iloc[header_row + 1 + j, 1:1 + size]
        fila_data = pd.to_numeric(fila_data, errors='coerce').fillna(0)
        data.append(fila_data.tolist())

    df = pd.DataFrame(data, columns=variables, index=variables)

    return df


# -----------------------------
# DESCRIPTORES (CON FALLBACK)
# -----------------------------
def obtener_descriptores(archivo):
    try:
        df_desc = pd.read_excel(archivo, sheet_name="DESCRIPTORES")
    except:
        df_desc = pd.read_excel(archivo, sheet_name="DESCRIPTORES ")

    mapping = dict(zip(df_desc.iloc[:, 0], df_desc.iloc[:, 1]))
    return mapping


# -----------------------------
# CLASIFICACIÓN MICMAC
# -----------------------------
def clasificar_variables(df):
    influencia = df.sum(axis=1)
    dependencia = df.sum(axis=0)

    resultado = pd.DataFrame({
        "Variable": df.index,
        "Influencia": influencia,
        "Dependencia": dependencia
    })

    prom_inf = resultado["Influencia"].median()
    prom_dep = resultado["Dependencia"].median()

    # pequeño ajuste de tolerancia
    epsilon = 0.01
    

    def clasificar(row):
        if row["Influencia"] >= prom_inf and row["Dependencia"] < prom_dep:
            return "Poder"
        elif row["Influencia"] >= prom_inf and row["Dependencia"] >= prom_dep:
            return "Conflicto"
        elif row["Influencia"] < prom_inf and row["Dependencia"] >= prom_dep:
            return "Resultados"
        else:
            return "Autonomas"

    resultado["Clasificacion"] = resultado.apply(clasificar, axis=1)

    return resultado


# -----------------------------
# FUNCIÓN PRINCIPAL
# -----------------------------
def procesar_micmac(archivo_micmac, wb):

    df = detectar_matriz_micmac(archivo_micmac)

    if df is None:
        raise ValueError("No se pudo detectar la matriz MICMAC")

    mapping = obtener_descriptores(archivo_micmac)

    resultado = clasificar_variables(df)

    # Reemplazar nombres cortos por descriptores
    resultado["Variable"] = resultado["Variable"].map(mapping).fillna(resultado["Variable"])

    # Separar grupos
    poder = resultado[resultado["Clasificacion"] == "Poder"]["Variable"].tolist()
    conflicto = resultado[resultado["Clasificacion"] == "Conflicto"]["Variable"].tolist()
    resultados = resultado[resultado["Clasificacion"] == "Resultados"]["Variable"].tolist()
    autonomas = resultado[resultado["Clasificacion"] == "Autonomas"]["Variable"].tolist()

    # -----------------------------
    # ESCRIBIR EN EXCEL (FILAS)
    # -----------------------------
    ws = wb.active  # o la hoja que estés usando
    
    # Limpiar antes (opcional pero recomendado)
    for col in ["B", "C", "D", "E"]:
        for fila in range(124, 141):
            ws[f"{col}{fila}"] = None
    
    # Función para escribir
    def escribir_columna(ws, lista, columna):
        fila = 124
        for item in lista:
            if fila > 140:
                break
            ws[f"{columna}{fila}"] = item
            fila += 1
    
    # Escribir datos
    escribir_columna(ws, poder, "B")
    escribir_columna(ws, conflicto, "C")
    escribir_columna(ws, resultados, "D")
    escribir_columna(ws, autonomas, "E")
