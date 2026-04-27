import pandas as pd


def procesar_metas(archivo_metas, wb):
    hoja = wb["Hoja1"]
    hoja2 = wb["Hoja2"]

    # -----------------------------
    # 1. OBTENER CANTON
    # -----------------------------
    canton = str(hoja["B2"].value).strip()

    if not canton:
        raise ValueError("No hay cantón en B2")

    # -----------------------------
    # 2. BUSCAR CODIGO EN HOJA2
    # -----------------------------
    codigo = None

    for fila in range(159, 257):
        nombre = hoja2[f"A{fila}"].value
        cod = hoja2[f"B{fila}"].value

        if nombre and str(nombre).strip().upper() == canton.upper():
            codigo = str(cod).strip()
            break

    if not codigo:
        raise ValueError(f"No se encontró código para: {canton}")

    hoja["B3"] = codigo

    # -----------------------------
    # 3. BUSCAR HOJA CORRECTA
    # -----------------------------
    xls = pd.ExcelFile(archivo_metas)

    hoja_objetivo = None
    for nombre in xls.sheet_names:
        if codigo in nombre:
            hoja_objetivo = nombre
            break

    if not hoja_objetivo:
        raise ValueError(f"No existe hoja para {codigo}")

    # -----------------------------
    # 4. LEER DATA
    # -----------------------------
    df = pd.read_excel(xls, sheet_name=hoja_objetivo)

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    
    # 🔥 importante para celdas combinadas
    df["Tipo"] = df["Tipo"].ffill()
    
    # 🔥 eliminar filas basura sin distrito
    df = df[df["Distrito"].notna()]
    
    # normalización
    df["Tipo"] = df["Tipo"].astype(str).str.strip().str.upper()
    df["Distrito"] = df["Distrito"].astype(str).str.strip()

    # -----------------------------
    # 5. COMUNIDAD
    # -----------------------------
    comunidad = df[df["Tipo"] == "COMUNIDAD"]

    fila_inicio = 64
    total_meta = 0
    total_conta = 0

    for i, (_, row) in enumerate(comunidad.iterrows()):
        fila = fila_inicio + i

        distrito = str(row["Distrito"]).upper()
        meta = int(row["Meta"])
        conta = int(row["Contabilidad"])

        hoja[f"B{fila}"] = distrito
        hoja[f"C{fila}"] = meta
        hoja[f"D{fila}"] = conta

        total_meta += meta
        total_conta += conta

    hoja["C60"] = total_meta
    hoja["D60"] = total_conta

    # -----------------------------
    # 6. COMERCIO
    # -----------------------------
    comercio = df[df["Tipo"] == "COMERCIO"]

    if not comercio.empty:
        row = comercio.iloc[0]
        hoja["I64"] = int(row["Meta"])
        hoja["J64"] = int(row["Contabilidad"])

    # -----------------------------
    # 7. POLICIAL
    # -----------------------------
    policial = df[df["Tipo"] == "POLICIAL"]

    if not policial.empty:
        row = policial.iloc[0]
        hoja["I65"] = int(row["Meta"])
        hoja["J65"] = int(row["Contabilidad"])
