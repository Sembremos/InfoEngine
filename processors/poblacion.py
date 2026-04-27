import pandas as pd


def procesar_poblacion(wb):
    hoja = wb["Hoja1"]

    # -----------------------------
    # 1. OBTENER CODIGO (B3)
    # -----------------------------
    codigo = str(hoja["B3"].value).strip()

    if not codigo:
        raise ValueError("No hay código en B3")

    # -----------------------------
    # 2. LEER ARCHIVO LOCAL
    # -----------------------------
    ruta = "plantillas/Poblaciones.xlsx"
    df = pd.read_excel(ruta)

    # limpiar nombres de columnas
    df.columns = [str(c).strip().upper() for c in df.columns]

    # limpiar columna código
    df["CODIGO"] = df["CODIGO"].astype(str).str.strip()

    # -----------------------------
    # 3. BUSCAR COINCIDENCIA
    # -----------------------------
    fila = df[df["CODIGO"] == codigo]

    if fila.empty:
        raise ValueError(f"No se encontró el código {codigo} en Poblaciones.xlsx")

    # -----------------------------
    # 4. OBTENER POBLACIÓN
    # -----------------------------
    poblacion = fila.iloc[0]["POBLACIO TERRITORIAL"]

    # asegurar número
    poblacion = int(poblacion)

    # -----------------------------
    # 5. ESCRIBIR EN INFOENGINE
    # -----------------------------
    hoja["B60"] = poblacion
