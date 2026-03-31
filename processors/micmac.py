import numpy as np
import pandas as pd

def clasificar_variables(df):

    variables = list(df.index)
    n = len(variables)

    # -----------------------------
    # MATRIZ BASE
    # -----------------------------
    M = df.values.astype(float)

    # Asegurar diagonal en 0
    np.fill_diagonal(M, 0)

    # -----------------------------
    # VARIABLES DE CONTROL
    # -----------------------------
    max_iter = 10
    rankings_prev = None

    M_power = M.copy()

    for iteration in range(1, max_iter + 1):

        # -----------------------------
        # CALCULAR INFLUENCIA Y DEPENDENCIA
        # -----------------------------
        influencia = M_power.sum(axis=1)
        dependencia = M_power.sum(axis=0)

        # Ranking (clave real)
        rank_inf = np.argsort(-influencia)
        rank_dep = np.argsort(-dependencia)

        rankings_actual = (tuple(rank_inf), tuple(rank_dep))

        # -----------------------------
        # VERIFICAR ESTABILIDAD
        # -----------------------------
        if rankings_prev is not None and rankings_actual == rankings_prev:
            break

        rankings_prev = rankings_actual

        # -----------------------------
        # SIGUIENTE ITERACIÓN
        # -----------------------------
        M_power = np.dot(M_power, M)

        # evitar crecimiento infinito (mantener estructura)
        M_power = np.where(M_power > 0, 1, 0)

    # -----------------------------
    # RESULTADO FINAL
    # -----------------------------
    influencia = M_power.sum(axis=1)
    dependencia = M_power.sum(axis=0)

    resultado = pd.DataFrame({
        "Variable": variables,
        "Influencia": influencia,
        "Dependencia": dependencia
    })

    # -----------------------------
    # CENTRO DEL SISTEMA
    # -----------------------------
    centro_inf = influencia.mean()
    centro_dep = dependencia.mean()

    # -----------------------------
    # CLASIFICACIÓN FINAL
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
