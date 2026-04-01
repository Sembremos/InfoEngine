import pandas as pd
import numpy as np

# -----------------------------
# DETECTAR MATRIZ
# -----------------------------
def detectar_matriz_micmac(archivo):
    df_raw = pd.read_excel(archivo, sheet_name="MATRIZ", header=None)

    header_row = None

    for i in range(len(df_raw) - 1):
        if isinstance(df_raw.iloc[i, 1], str) and isinstance(df_raw.iloc[i+1, 1], (int, float)):
            header_row = i
            break

    if header_row is None:
        return None

    variables = df_raw.iloc[header_row, 1:].dropna().tolist()
    size = len(variables)

    data = []

    for j in range(size):
        fila = df_raw.iloc[header_row + 1 + j, 1:1 + size]
        fila = pd.to_numeric(fila, errors='coerce')
        fila.iloc[j] = 0
        fila = fila.fillna(0)
        data.append(fila.tolist())

    return pd.DataFrame(data, columns=variables, index=variables)


# -----------------------------
# DESCRIPTORES
# -----------------------------
def obtener_descriptores(archivo):
    try:
        df = pd.read_excel(archivo, sheet_name="DESCRIPTORES")
    except:
        df = pd.read_excel(archivo, sheet_name="DESCRIPTORES ")

    return dict(zip(df.iloc[:, 0], df.iloc[:, 1]))


# -----------------------------
# MOTOR MICMAC REAL
# -----------------------------
def calcular_matriz_acumulada(M, iteraciones=4):

    M_power = M.copy()
    M_acum = M.copy()

    for _ in range(2, iteraciones + 1):
        M_power = np.dot(M_power, M)
        M_acum += M_power

    return M_acum


# -----------------------------
# CLASIFICACIÓN CORRECTA
# -----------------------------
def clasificar_variables(df):

    M = df.values.astype(float)

    # 🔥 MATRIZ ACUMULADA (clave MICMAC)
    M_acum = calcular_matriz_acumulada(M, iteraciones=4)

    # -----------------------------
    # MOTRICIDAD Y DEPENDENCIA
    # -----------------------------
    influencia = M_acum.sum(axis=1)
    dependencia = M_acum.sum(axis=0)

    # normalizar (solo para estabilidad visual)
    if influencia.max() != 0:
        influencia = influencia / influencia.max()
    if dependencia.max() != 0:
        dependencia = dependencia / dependencia.max()

    resultado = pd.DataFrame({
        "Variable": df.index,
        "Influencia": influencia,
        "Dependencia": dependencia
    })

    # -----------------------------
    # CENTRO DEL SISTEMA
    # -----------------------------
    centro_I = np.mean(influencia)
    centro_D = np.mean(dependencia)

    # -----------------------------
    # CLASIFICACIÓN POR CUADRANTES
    # -----------------------------
    def clasificar(row):
        I = row["Influencia"]
        D = row["Dependencia"]

        if I >= centro_I and D < centro_D:
            return "Poder"
        elif I >= centro_I and D >= centro_D:
            return "Conflicto"
        elif I < centro_I and D >= centro_D:
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

    resultado["Variable"] = resultado["Variable"].map(mapping).fillna(resultado["Variable"])

    poder = resultado[resultado["Clasificacion"] == "Poder"]["Variable"].tolist()
    conflicto = resultado[resultado["Clasificacion"] == "Conflicto"]["Variable"].tolist()
    resultados = resultado[resultado["Clasificacion"] == "Resultados"]["Variable"].tolist()
    autonomas = resultado[resultado["Clasificacion"] == "Autonomas"]["Variable"].tolist()

    ws = wb.active

    # limpiar
    for col in ["B", "C", "D", "E"]:
        for fila in range(124, 141):
            ws[f"{col}{fila}"] = None

    def escribir(lista, col):
        fila = 124
        for item in lista:
            if fila > 140:
                break
            ws[f"{col}{fila}"] = item
            fila += 1

    escribir(poder, "B")
    escribir(conflicto, "C")
    escribir(resultados, "D")
    escribir(autonomas, "E")
