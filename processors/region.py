def normalizar(valor):
    if not valor:
        return None
    return str(valor).upper().strip()


def mapear_region_por_codigo(codigo):
    """
    Mapea directamente códigos tipo:
    D1, D2, ..., D82 E, D82 O
    """

    if not codigo:
        return None

    codigo = normalizar(codigo)

    # --- REGION 1 ---
    if codigo in [f"D{i}" for i in range(0, 26)]:
        return 1

    # --- REGION 2 ---
    elif codigo in [f"D{i}" for i in range(26, 37)]:
        return 2

    # --- REGION 3 ---
    elif codigo in [f"D{i}" for i in range(37, 49)]:
        return 3

    # --- REGION 4 ---
    elif codigo in ["D49","D50","D51","D52","D53","D54","D56","D57"]:
        return 4

    # --- REGION 5 ---
    elif codigo in ["D60","D61","D62","D63","D64","D65","D66","D67","D68","D69"]:
        return 5

    # --- REGION 6 ---
    elif codigo in ["D71","D72","D73","D74","D75","D76","D77","D78"]:
        return 6

    # --- REGION 7 ---
    elif codigo in ["D79","D80","D81"]:
        return 7

    # --- REGION 8 ---
    elif codigo in ["D82 E", "D82 O", "D83","D86","D87"]:
        return 8

    # --- REGION 9 ---
    elif codigo in ["D88","D90","D91","D92"]:
        return 9

    # --- REGION 10 ---
    elif codigo in ["D94","D95","D96","D97"]:
        return 10

    # --- REGION 11 ---
    elif codigo in ["D70","D84","D85"]:
        return 11

    # --- REGION 12 ---
    elif codigo in ["D89","D93","D58","D59","D95"]:
        return 12

    return None


def escribir_region(wb):
    hoja = wb["Hoja1"]

    valor_b3 = hoja["B3"].value
    codigo = normalizar(valor_b3)

    region = mapear_region_por_codigo(codigo)

    if region is not None:
        hoja["D2"] = region
