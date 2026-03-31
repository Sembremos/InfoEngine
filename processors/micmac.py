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

    prom_inf = resultado["Influencia"].mean()
    prom_dep = resultado["Dependencia"].mean()

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

    # Convertir a texto
    poder_txt = "\n".join(poder)
    conflicto_txt = "\n".join(conflicto)
    resultados_txt = "\n".join(resultados)
    autonomas_txt = "\n".join(autonomas)

    # Escribir en Excel
    ws = wb.active  # cambiar si usás una hoja específica

    ws["B124"] = poder_txt
    ws["C124"] = conflicto_txt
    ws["D124"] = resultados_txt
    ws["E124"] = autonomas_txt
