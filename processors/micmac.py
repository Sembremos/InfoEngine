import pandas as pd
import numpy as np

# -----------------------------
# DETECTAR MATRIZ MICMAC
# -----------------------------
def detectar_matriz_micmac(archivo):
    df_raw = pd.read_excel(archivo, sheet_name="MATRIZ", header=None)

    header_row = None

    for i in range(len(df_raw) - 1):

        fila = df_raw.iloc[i]
        fila_siguiente = df_raw.iloc[i + 1]

        if (
            isinstance(fila[1], str) and
            isinstance(fila[2], str) and
            isinstance(fila_siguiente[1], (int, float))
        ):
            header_row = i
            break

    if header_row is None:
        return None

    variables = df_raw.iloc[header_row, 1:].dropna().tolist()
    size = len(variables)

    data = []

    for j in range(size):
        fila_data = df_raw.iloc[header_row + 1 + j, 1:1 + size]
        fila_data = pd.to_numeric(fila_data, errors='coerce').fillna(0)
        data.append(fila_data.tolist())

    df = pd.DataFrame(data, columns=variables, index=variables)

    return df


# -----------------------------
# DESCRIPTORES
# -----------------------------
def obtener_descriptores(archivo):
    try:
        df_desc = pd.read_excel(archivo, sheet_name="DESCRIPTORES")
    except:
        df_desc = pd.read_excel(archivo, sheet_name="DESCRIPTORES ")

    mapping = dict(zip(df_desc.iloc[:, 0], df_desc.iloc[:, 1]))
    return mapping


# -----------------------------
# CLASIFICACIÓN SIMPLE (DIRECTA)
# -----------------------------
def clasificar_variables(df):

    variables = list(df.index)
    n = len(variables)

    # -----------------------------
    # MATRIZ BASE
    # -----------------------------
    M = df.values.astype(float)

    # -----------------------------
    # MATRIZ BINARIA (grafo)
    # -----------------------------
    # MICMAC realmente trabaja sobre existencia + peso
    G = (M > 0).astype(int)

    # -----------------------------
    # ACUMULACIÓN DE CAMINOS
    # -----------------------------
    # caminos de longitud 1 a 5 (clave real)
    max_iter = 5

    influencia_total = np.zeros(n)
    dependencia_total = np.zeros(n)

    G_power = G.copy()

    for _ in range(max_iter):

        # influencia = salidas
        influencia_total += G_power.sum(axis=1)

        # dependencia = entradas
        dependencia_total += G_power.sum(axis=0)

        # siguiente nivel de caminos
        G_power = np.dot(G_power, G)

        # evitar crecimiento infinito
        G_power = (G_power > 0).astype(int)

    resultado = pd.DataFrame({
        "Variable": variables,
        "Influencia": influencia_total,
        "Dependencia": dependencia_total
    })

    # -----------------------------
    # CENTRO DEL PLANO
    # -----------------------------
    centro_inf = influencia_total.mean()
    centro_dep = dependencia_total.mean()

    # -----------------------------
    # CLASIFICACIÓN
    # -----------------------------
    def clasificar(row):
        if row["Influencia"] > centro_inf and row["Dependencia"] < centro_dep:
            return "Poder"
        elif row["Influencia"] > centro_inf and row["Dependencia"] > centro_dep:
            return "Conflicto"
        elif row["Influencia"] < centro_inf and row["Dependencia"] > centro_dep:
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
    # ESCRIBIR EN EXCEL
    # -----------------------------
    ws = wb.active

    # Limpiar rango
    for col in ["B", "C", "D", "E"]:
        for fila in range(124, 141):
            ws[f"{col}{fila}"] = None

    def escribir(lista, columna):
        fila = 124
        for item in lista:
            if fila > 140:
                break
            ws[f"{columna}{fila}"] = item
            fila += 1

    escribir(poder, "B")
    escribir(conflicto, "C")
    escribir(resultados, "D")
    escribir(autonomas, "E")
