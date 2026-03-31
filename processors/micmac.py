import pandas as pd
import numpy as np
from sklearn.cluster import KMeans

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
# CLASIFICACIÓN MICMAC REAL
# -----------------------------
def clasificar_variables(df):

    # MATRIZ BASE
    M = df.values.astype(float)

    if M.max() != 0:
        M = M / M.max()

    # MATRIZ INDIRECTA
    M_total = M.copy()
    M_power = M.copy()

    for _ in range(1, 12):
        M_power = np.dot(M_power, M)
        M_total += M_power

    # INFLUENCIA / DEPENDENCIA
    influencia = M_total.sum(axis=1)
    dependencia = M_total.sum(axis=0)

    resultado = pd.DataFrame({
        "Variable": df.index,
        "Influencia": influencia,
        "Dependencia": dependencia
    })

    # -----------------------------
    # CLUSTERING (CLAVE)
    # -----------------------------
    X = resultado[["Influencia", "Dependencia"]].values

    kmeans = KMeans(n_clusters=4, random_state=0, n_init=10)
    resultado["cluster"] = kmeans.fit_predict(X)

    centros = kmeans.cluster_centers_

    clasificacion = {}

    for i, (inf, dep) in enumerate(centros):

        if inf > np.mean(influencia) and dep < np.mean(dependencia):
            clasificacion[i] = "Poder"
        elif inf > np.mean(influencia) and dep > np.mean(dependencia):
            clasificacion[i] = "Conflicto"
        elif inf < np.mean(influencia) and dep > np.mean(dependencia):
            clasificacion[i] = "Resultados"
        else:
            clasificacion[i] = "Autonomas"

    resultado["Clasificacion"] = resultado["cluster"].map(clasificacion)

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
