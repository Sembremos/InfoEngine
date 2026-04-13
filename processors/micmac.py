import streamlit as st

# -----------------------------
# LISTA DE PROBLEMÁTICAS
# -----------------------------
def obtener_problematicas():
    return [
        "ABANDONO DE PERSONAS (MENOR DE EDAD, ADULTO MAYOR O CON CAPACIDADES DIFERENTES)",
        "ABIGEATO (ROBO Y DESTACE DE GANADO)",
        "ABORTO",
        "ABUSO DE AUTORIDAD",
        "ACCIDENTES DE TRANSITO",
        "ACCIONAMIENTO DE ARMA DE FUEGO (BALACERAS)",
        "ACOSO ESCOLAR (BULLYING)",
        "ACOSO LABORAL (MOBBING)",
        "ACOSO SEXUAL CALLEJERO",
        "ACTOS OBSCENOS EN VIA PUBLICA",
        "ADMINISTRACION FRAUDULENTA, APROPIACIONES INDEBIDAS O ENRIQUECIMIENTO ILICITO",
        "AGRESION CON ARMAS",
        "AGRUPACIONES DELINCUENCIALES NO ORGANIZADAS",
        "ALTERACIÓN DE DATOS Y SABOTAJE INFORMÁTICO",
        "AMBIENTE LABORAL INADECUADO",
        "AMENAZAS",
        "ANALFABETISMO",
        "BAJOS SALARIOS",
        "BARRAS DE FUTBOL",
        "BUNKER (EJE DE EXPENDIO DE DROGAS)",
        "CALUMNIA",
        "CAZA ILEGAL",
        "CONDUCCION TEMERARIA",
        "CONSUMO DE ALCOHOL EN VÍA PÚBLICA",
        "CONSUMO DE DROGAS",
        "CONTAMINACION SONICA",
        "CONTRABANDO",
        "CORRUPCION",
        "CORRUPCION POLICIAL",
        "CULTIVO DE DROGA (MARIHUANA)",
        "DAÑO AMBIENTAL",
        "DAÑOS/VANDALISMO",
        "DEFICENCIA EN LA INFRAESTRUCTURA VIAL",
        "DEFICIENCIA EN LA LINEA 9-1-1",
        "DEFICIENCIAS EN EL ALUMBRADO PUBLICO",
        "DELICUENCIA ORGANIZADA",
        "DELITOS SEXUALES",
        "DESAPARICION DE PERSONAS",
        "DESEMPLEO",
        "DISTURBIOS (RIÑAS)",
        "ESTAFA O DEFRAUDACION",
        "EXTORSION",
        "FALTA DE CAMARAS DE SEGURIDAD",
        "FALTA DE CAPACITACION POLICIAL",
        "FALTA DE CONTROL A PATENTES",
        "FALTA DE CONTROL FRONTERIZO",
        "FALTA DE CULTURA VIAL",
        "FALTA DE INVERSION SOCIAL",
        "FALTA DE PERSONAL POLICIAL",
        "FALTA DE PRESENCIA POLICIAL",
        "FAMILIAS DISFUNCIONALES",
        "FRAUDE INFORMATICO",
        "GROOMING",
        "HOMICIDIO (PROFESIONAL)",
        "HURTO",
        "INEFECTIVIDAD EN EL SERVICIO DE POLICIA",
        "INEFICIENCIA EN LA ADMINISTRACION DE JUSTICIA",
        "INFRAESTRUCTURA INADECUADA",
        "INTOLERANCIA SOCIAL",
        "LAVADO DE ACTIVOS",
        "LESIONES",
        "MALTRATO ANIMAL",
        "NARCOTRAFICO",
        "PERCEPCION DE INSEGURIDAD",
        "PERSONAS CON EXCESO DE TIEMPO DE OCIO",
        "PERSONAS EN SITUACION DE CALLE",
        "PORTACION ILEGAL DE ARMAS",
        "PROBLEMAS VECINALES",
        "RECEPTACION",
        "RESISTENCIA (IRRESPETO A LA AUTORIDAD)",
        "ROBO A COMERCIO (INTIMIDACION)",
        "ROBO A COMERCIO (TACHA)",
        "ROBO A PERSONAS",
        "ROBO A VEHICULOS (TACHA)",
        "ROBO A VIVIENDA (INTIMIDACION)",
        "ROBO A VIVIENDA (TACHA)",
        "ROBO DE BICICLETA",
        "ROBO DE MOTOCICLETAS/VEHICULOS(BAJONAZO)",
        "ROBO DE VEHICULOS",
        "SECUESTRO",
        "SUICIDIO",
        "TALA ILEGAL",
        "TENTATIVA DE HOMICIDIO",
        "TRAFICO DE ARMAS",
        "TRAFICO DE INFLUENCIAS",
        "TRATA DE PERSONAS",
        "USURPACION DE TERRENOS",
        "VENTA DE DROGAS",
        "VENTAS INFORMALES",
        "VIOLACIÓN DE DOMICILIO",
        "VIOLENCIA DE GENERO",
        "VIOLENCIA INTRAFAMILIAR",
        "XENOFOBIA",
        "ZONAS DE PROSTITUCION",
        "ZONAS VULNERABLES",
        "ROBO A TRANSPORTE PÚBLICO CON INTIMIDACIÓN",
        "EXPLOTACIÓN SEXUAL INFANTIL",
        "EXPLOTACIÓN LABORAL INFANTIL",
        "TRÁFICO ILEGAL DE PERSONAS",
        "FEMICIDIO"
    ]


# -----------------------------
# UI MICMAC
# -----------------------------
def ui_micmac():

    st.subheader("Clasificación MICMAC")

    lista = obtener_problematicas()

    poder = st.multiselect("Poder", lista)
    conflicto = st.multiselect("Conflicto", lista)
    resultados = st.multiselect("Resultados", lista)
    autonomas = st.multiselect("Autónomas", lista)

    # validación
    todas = poder + conflicto + resultados + autonomas

    if len(todas) != len(set(todas)):
        st.error("Hay problemáticas repetidas en múltiples cuadrantes")
        st.stop()

    return poder, conflicto, resultados, autonomas


# -----------------------------
# ESCRIBIR EN EXCEL
# -----------------------------
def escribir_cuadrantes_manual(wb, poder, conflicto, resultados, autonomas):

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

#clasificador
def clasificar_y_escribir_riesgos_delitos(wb, poder, conflicto):

    ws1 = wb["Hoja1"]
    ws2 = wb["Hoja2"]

    # -----------------------------
    # LEER LISTAS DESDE EXCEL
    # -----------------------------
    delitos = []
    for fila in range(3, 90):  # B3 a B89
        valor = ws2[f"B{fila}"].value
        if valor:
            delitos.append(valor.strip())

    riesgos = []
    for fila in range(92, 155):  # B92 a B154
        valor = ws2[f"B{fila}"].value
        if valor:
            riesgos.append(valor.strip())

    # -----------------------------
    # LIMPIAR COLUMNAS DESTINO
    # -----------------------------
    for fila in range(123, 200):
        ws1[f"N{fila}"] = None  # Riesgos
        ws1[f"O{fila}"] = None  # Delitos

    # -----------------------------
    # CLASIFICAR
    # -----------------------------
    seleccionadas = poder + conflicto

    lista_riesgos = []
    lista_delitos = []

    for item in seleccionadas:
        item_clean = item.strip()

        if item_clean in riesgos:
            lista_riesgos.append(item_clean)

        elif item_clean in delitos:
            lista_delitos.append(item_clean)

    # -----------------------------
    # ESCRIBIR RESULTADOS
    # -----------------------------
    fila_r = 123
    for r in lista_riesgos:
        ws1[f"N{fila_r}"] = r
        fila_r += 1

    fila_d = 123
    for d in lista_delitos:
        ws1[f"O{fila_d}"] = d
        fila_d += 1
    

    escribir(poder, "B")
    escribir(conflicto, "C")
    escribir(resultados, "D")
    escribir(autonomas, "E")
