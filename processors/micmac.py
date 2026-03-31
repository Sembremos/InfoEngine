import pandas as pd

def detectar_matriz_micmac(archivo):
    df_raw = pd.read_excel(archivo, sheet_name="MATRIZ", header=None)

    for i in range(len(df_raw)):
        fila = df_raw.iloc[i]

        # Contar textos en la fila (posibles variables)
        textos = [x for x in fila[1:] if isinstance(x, str)]

        # Condición: al menos 3 variables detectadas
        if len(textos) >= 3:

            # Intentar leer matriz desde ahí
            try:
                df = pd.read_excel(
                    archivo,
                    sheet_name="MATRIZ",
                    skiprows=i,
                    index_col=0
                )

                # Limpiar completamente
                df = df.dropna(how="all")
                df = df.dropna(axis=1, how="all")

                # Forzar numérico
                df = df.apply(pd.to_numeric, errors='coerce')

                # Validar que tenga suficientes datos numéricos
                if df.notna().sum().sum() > (df.shape[0] * df.shape[1] * 0.6):

                    # Validar cuadrada
                    if df.shape[0] == df.shape[1]:
                        return df

            except:
                continue

    return None


def obtener_descriptores(archivo):
    df_desc = pd.read_excel(archivo, sheet_name="DESCRIPTORES")

    mapping = dict(zip(df_desc.iloc[:, 0], df_desc.iloc[:, 1]))
    return mapping


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


def procesar_micmac(archivo_micmac, wb):

    df = detectar_matriz_micmac(archivo_micmac)

    if df is None:
        raise ValueError("No se pudo detectar la matriz MICMAC")

    mapping = obtener_descriptores(archivo_micmac)

    resultado = clasificar_variables(df)

    # Reemplazar nombres cortos por descriptores
    resultado["Variable"] = resultado["Variable"].map(mapping).fillna(resultado["Variable"])

    # Separar listas
    poder = resultado[resultado["Clasificacion"] == "Poder"]["Variable"].tolist()
    conflicto = resultado[resultado["Clasificacion"] == "Conflicto"]["Variable"].tolist()
    resultados = resultado[resultado["Clasificacion"] == "Resultados"]["Variable"].tolist()
    autonomas = resultado[resultado["Clasificacion"] == "Autonomas"]["Variable"].tolist()

    # Convertir a texto (una por línea)
    poder_txt = "\n".join(poder)
    conflicto_txt = "\n".join(conflicto)
    resultados_txt = "\n".join(resultados)
    autonomas_txt = "\n".join(autonomas)

    ws = wb.active  # o cambiar si usás una hoja específica

    ws["B124"] = poder_txt
    ws["C124"] = conflicto_txt
    ws["D124"] = resultados_txt
    ws["E124"] = autonomas_txt
