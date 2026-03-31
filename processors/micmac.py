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
import numpy as np

def clasificar_variables(df):

    # -----------------------------
    # MATRIZ BASE
    # -----------------------------
    M = df.values.astype(float)

    # Normalización (clave)
    if M.max() != 0:
        M = M / M.max()

    # -----------------------------
    # CONVERGENCIA REAL (NO ITERACIONES FIJAS)
    # -----------------------------
    M_total = M.copy()
    M_power = M.copy()

    for _ in range(100):  # límite alto de seguridad

        M_next = np.dot(M_power, M)

        # criterio de convergencia real
        if np.allclose(M_next, M_power, atol=1e-6):
            break

        M_total += M_next
        M_power = M_next

    # -----------------------------
    # INFLUENCIA Y DEPENDENCIA REALES
    # -----------------------------
    influencia = M_total.sum(axis=1)
    dependencia = M_total.sum(axis=0)

    resultado = pd.DataFrame({
        "Variable": df.index,
        "Influencia": influencia,
        "Dependencia": dependencia
    })

    # -----------------------------
    # CENTRO REAL DEL PLANO
    # -----------------------------
    centro_inf = np.mean(influencia)
    centro_dep = np.mean(dependencia)

    # -----------------------------
    # CLASIFICACIÓN EXACTA
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
