import pandas as pd
import numpy as np

# -----------------------------
# DETECTAR MATRIZ
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
        fila_data = pd.to_numeric(fila_data, errors='coerce')

        # eliminar diagonal
        fila_data.iloc[j] = 0
        fila_data = fila_data.fillna(0)

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

    return dict(zip(df_desc.iloc[:, 0], df_desc.iloc[:, 1]))


# -----------------------------
# MOTOR CALIBRADO
# -----------------------------
def clasificar_variables(df):

    M = df.values.astype(float)
    n = len(M)

    # -----------------------------
    # ACUMULACIÓN (como MICMAC real)
    # -----------------------------
    influencia = np.zeros(n)
    dependencia = np.zeros(n)

    M_temp = M.copy()

    for _ in range(4):  # clave (no más)
        influencia += M_temp.sum(axis=1)
        dependencia += M_temp.sum(axis=0)
        M_temp = np.dot(M_temp, M)

    # -----------------------------
    # NORMALIZACIÓN
    # -----------------------------
    influencia = influencia / max(influencia) if max(influencia) != 0 else influencia
    dependencia = dependencia / max(dependencia) if max(dependencia) != 0 else dependencia

    resultado = pd.DataFrame({
        "Variable": df.index,
        "Influencia": influencia,
        "Dependencia": dependencia
    })

    # -----------------------------
    # UMBRALES ADAPTATIVOS (CLAVE)
    # -----------------------------
    # estos valores están calibrados según tu plano real
    t_inf = np.percentile(influencia, 60)
    t_dep = np.percentile(dependencia, 60)

    t_inf_low = np.percentile(influencia, 30)
    t_dep_low = np.percentile(dependencia, 30)

    def clasificar(row):
        I = row["Influencia"]
        D = row["Dependencia"]

        # PODER
        if I >= t_inf and D <= t_dep_low:
            return "Poder"

        # CONFLICTO
        elif I >= t_inf and D >= t_dep:
            return "Conflicto"

        # RESULTADOS
        elif I <= t_inf_low and D >= t_dep:
            return "Resultados"

        # AUTÓNOMAS
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
