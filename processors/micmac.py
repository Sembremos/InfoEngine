import pandas as pd

# -----------------------------
# DETECTAR MATRIZ MICMAC (ROBUSTO)
# -----------------------------
def detectar_matriz_micmac(archivo):
    df_raw = pd.read_excel(archivo, sheet_name="MATRIZ", header=None)

    # Convertir todo a numérico donde se pueda
    df_num = df_raw.applymap(lambda x: pd.to_numeric(x, errors='coerce'))

    # Crear máscara de valores válidos (0–3)
    mask = df_num.applymap(lambda x: x in [0,1,2,3] if pd.notna(x) else False)

    max_area = 0
    best_block = None

    rows, cols = mask.shape

    # Buscar el bloque cuadrado más grande de números
    for i in range(rows):
        for j in range(cols):

            if not mask.iloc[i, j]:
                continue

            size = 1
            while True:
                if i + size > rows or j + size > cols:
                    break

                bloque = mask.iloc[i:i+size, j:j+size]

                if bloque.all().all():
                    area = size * size

                    if area > max_area:
                        max_area = area
                        best_block = (i, j, size)

                    size += 1
                else:
                    break

    if best_block is None:
        return None

    i, j, size = best_block

    # Extraer matriz con encabezados
    df = pd.read_excel(
        archivo,
        sheet_name="MATRIZ",
        skiprows=i-1,
        usecols=range(j, j+size+1),
        index_col=0
    )

    df = df.iloc[:size, :size]

    # Forzar numérico
    df = df.apply(pd.to_numeric, errors='coerce').fillna(0)

    return df


# -----------------------------
# DESCRIPTORES
# -----------------------------
def obtener_descriptores(archivo):
    df_desc = pd.read_excel(archivo, sheet_name="DESCRIPTORES")

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
    ws = wb.active  # si usás otra hoja, decímelo

    ws["B124"] = poder_txt
    ws["C124"] = conflicto_txt
    ws["D124"] = resultados_txt
    ws["E124"] = autonomas_txt
