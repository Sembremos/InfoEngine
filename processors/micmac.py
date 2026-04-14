import streamlit as st

# -----------------------------
# LISTA DE PROBLEMÁTICAS
# -----------------------------
def obtener_problematicas():
    return [
        "ABANDONO DE PERSONAS (MENOR DE EDAD, ADULTO MAYOR O CON CAPACIDADES DIFERENTES)"
        "ABIGEATO (ROBO Y DESTACE DE GANADO)"
        "ABORTO"
        "ABUSO DE AUTORIDAD"
        "ACCIONAMIENTO DE ARMA DE FUEGO (BALACERAS)"
        "ADMINISTRACION FRAUDULENTA, APROPIACIONES INDEBIDAS O ENRIQUECIMIENTO ILICITO"
        "AGRESION CON ARMAS"
        "ALTERACION DE DATOS Y SABOTAJE INFORMATICO"
        "AMENAZAS"
        "CALUMNIA"
        "CAZA ILEGAL"
        "CONDUCCION TEMERARIA"
        "CONTRABANDO"
        "CORRUPCION"
        "CORRUPCION POLICIAL"
        "CULTIVO DE DROGA (MARIHUANA)"
        "DANO AMBIENTAL"
        "DANOS/VANDALISMO"
        "DELICUENCIA ORGANIZADA"
        "DELITOS CONTRA EL AMBITO DE INTIMIDAD (VIOLACION DE SECRETOS (CORRESPONDENCIA Y COMUNICACIONES ELECTRONICAS))"
        "DELITOS SEXUALES"
        "DESOBEDIENCIA"
        "DESORDENES EN VIA PUBLICA"
        "DISTURBIOS (RINAS)"
        "ESTAFA O DEFRAUDACION"
        "ESTUPRO (DELITOS SEXUALES CONTRA MENOR DE EDAD)"
        "EVASION Y QUEBRANTAMIENTO DE PENA"
        "EXPLOSIVOS"
        "EXTORSION"
        "FABRICACION, PRODUCCION O REPRODUCCION DE PORNOGRAFIA"
        "FALSIFICACION DE MONEDA Y OTROS VALORES"
        "FRAUDE INFORMATICO"
        "GROOMING"
        "HOMICIDIO (PROFESIONAL)"
        "HURTO"
        "INCUMPLIMIENTO DEL DEBER ALIMENTARIO"
        "LAVADO DE ACTIVOS"
        "LESIONES"
        "LEY DE ARMAS Y EXPLOSIVOS N° 7530"
        "MALTRATO ANIMAL"
        "NARCOTRAFICO"
        "MENORES EN VULNERABILIDAD"
        "PESCA ILEGAL"
        "PORTACION ILEGAL DE ARMAS"
        "PRIVACION DE LIBERTAD SIN ANIMO DE LUCRO"
        "RECEPTACION"
        "RELACIONES IMPROPIAS"
        "RESISTENCIA (IRRESPETO A LA AUTORIDAD)"
        "ROBO A COMERCIO (INTIMIDACION)"
        "ROBO A COMERCIO (TACHA)"
        "ROBO A EDIFICACION (TACHA)"
        "ROBO A PERSONAS"
        "ROBO A TRANSPORTE COMERCIAL"
        "ROBO A VEHICULOS (TACHA)"
        "ROBO A VIVIENDA (INTIMIDACION)"
        "ROBO A VIVIENDA(TACHA)"
        "ROBO DE BICICLETA"
        "ROBO DE CULTIVOS"
        "ROBO DE MOTOCICLETAS/VEHICULOS(BAJONAZO)"
        "ROBO DE VEHICULOS"
        "SECUESTRO"
        "SIMULACION DE DELITO"
        "SUSTRACCION DE UNA PERSONA MENOR DE EDAD O INCAPAZ"
        "TENTATIVA DE HOMICIDIO"
        "TERRORISMO"
        "TRAFICO DE ARMAS"
        "TRAFICO DE INFLUENCIAS"
        "TRATA DE PERSONAS"
        "VIOLENCIA DOMESTICA"
        "USO ILEGAL DE UNIFORMES, INSIGNIAS O DISPOSITIVOS POLICIALES"
        "USURPACION DE TERRENOS (PRECARIOS)"
        "VENTA DE DROGAS"
        "VIOLACION DE DOMICILIO"
        "VIOLACION DE LA CUSTODIA DE LAS COSAS"
        "VIOLACION DE SELLOS"
        "VIOLENCIA DE GENERO"
        "VIOLENCIA INTRAFAMILIAR"
        "ROBO A VIVIENDA (TACHA)"
        "ROBO A TRANSPORTE PUBLICO CON INTIMIDACION"
        "ROBO DE CABLE"
        "EXPLOTACION SEXUAL INFANTIL"
        "EXPLOTACION LABORAL INFANTIL"
        "TRAFICO ILEGAL DE PERSONAS"
        "FEMICIDIO"
        "ROBO DE EMBARCACIONES"
        "ROBO DE EQUIPO AGRICOLA"
        "ROBO A EMBARCACIONES (TACHA)"
        "RIESGO"
        "ACCIDENTES DE TRANSITO"
        "ACOSO ESCOLAR (BULLYING)"
        "ACOSO LABORAL (MOBBING)"
        "ACOSO SEXUAL CALLEJERO"
        "ACTOS OBSCENOS EN VIA PUBLICA"
        "AGRUPACIONES DELINCUENCIALES NO ORGANIZADAS"
        "ANALFABETISMO"
        "BAJOS SALARIOS"
        "BARRAS DE FUTBOL"
        "BUNKER (EJE DE EXPENDIO DE DROGAS)"
        "CONSUMO DE ALCOHOL EN VIA PUBLICA"
        "CONSUMO DE DROGAS"
        "CONTAMINACION SONICA"
        "DEFICIENCIAS EN LA INFRAESTRUCTURA VIAL"
        "DEFICIENCIAS EN EL ALUMBRADO PUBLICO"
        "DESAPARICION DE PERSONAS"
        "DESARTICULACION INTERINSTITUCIONAL"
        "DESEMPLEO"
        "FALTA DE OPORTUNIDADES LABORALES"
        "DESVINCULACION ESTUDIANTIL"
        "ENFRENTAMIENTOS ESTUDIANTILES"
        "FACILISMO ECONOMICO"
        "FALTA DE CAMARAS DE SEGURIDAD"
        "FALTA DE CONTROL A PATENTES"
        "FALTA DE CONTROL FRONTERIZO"
        "FALTA DE CORRESPONSABILIDAD EN SEGURIDAD"
        "FALTA DE CULTURA VIAL"
        "FALTA DE CULTURA Y COMPROMISO CIUDADANO"
        "FALTA DE EDUCACION FAMILIAR"
        "FALTA DE INVERSION SOCIAL"
        "FALTA DE LEGISLACION DE EXTINCION DE DOMINIO"
        "FALTA DE POLITICAS PUBLICAS EN SEGURIDAD"
        "FALTA DE PRESENCIA POLICIAL"
        2FALTA DE SALUBRIDAD PUBLICA"
        "FAMILIAS DISFUNCIONALES"
        "HACINAMIENTO CARCELARIO"
        "HACINAMIENTO POLICIAL"
        "HOSPEDAJES ILEGALES (CUARTERIAS)"
        "INCUMPLIMIENTO AL PLAN REGULADOR DE LA MUNICIPALIDAD"
        "INDIFERENCIA SOCIAL"
        "INEFICIENCIA EN LA ADMINISTRACION DE JUSTICIA"
        "INTOLERANCIA SOCIAL"
        "LEY DE CONTROL DE TABACO (LEY 9028)"
        "LOTES BALDIOS"
        "NECESIDADES BASICAS INSATISFECHAS"
        "PERCEPCION DE INSEGURIDAD"
        "PERDIDA DE ESPACIOS PUBLICOS"
        "PERSONAS CON EXCESO DE TIEMPO DE OCIO"
        "PERSONAS EN ESTADO MIGRATORIO IRREGULAR"
        "PERSONAS EN SITUACION DE CALLE"
        "PRESENCIA MULTICULTURAL"
        "PROBLEMAS VECINALES"
        "SISTEMA JURIDICO DESACTUALIZADO"
        "SUICIDIO"
        "TENDENCIA SOCIAL HACIA EL DELITO (PAUTAS DE CRIANZA VIOLENTA)"
        "TENENCIA DE DROGA"
        "TRABAJO INFORMAL"
        "TRANSPORTE INFORMAL (UBER, PORTEADORES, PIRATAS)"
        "VENTAS INFORMALES (AMBULANTES)"
        "VIGILANCIA INFORMAL"
        "XENOFOBIA"
        "ZONAS DE PROSTITUCION"
        "ZONAS VULNERABLES"
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

    todas = poder + conflicto + resultados + autonomas

    if len(todas) != len(set(todas)):
        st.error("Hay problemáticas repetidas en múltiples cuadrantes")
        st.stop()

    return poder, conflicto, resultados, autonomas


# -----------------------------
# ESCRIBIR CUADRANTES
# -----------------------------
def escribir_cuadrantes_manual(wb, poder, conflicto, resultados, autonomas):

    ws = wb.active

    for col in ["B", "C", "D", "E"]:
        for fila in range(124, 141):
            try:
                ws[f"{col}{fila}"].value = None
            except:
                pass

    def escribir(lista, columna):
        fila = 124
        for item in lista:
            if fila > 140:
                break
            try:
                ws[f"{columna}{fila}"].value = item
            except:
                pass
            fila += 1

    escribir(poder, "B")
    escribir(conflicto, "C")
    escribir(resultados, "D")
    escribir(autonomas, "E")


# -----------------------------
# CLASIFICADOR
# -----------------------------
def clasificar_y_escribir_riesgos_delitos(wb, poder, conflicto):

    ws1 = wb["Hoja1"]
    ws2 = wb["Hoja2"]

    def normalizar(texto):
        return (
            texto.strip()
            .upper()
            .replace("Á", "A")
            .replace("É", "E")
            .replace("Í", "I")
            .replace("Ó", "O")
            .replace("Ú", "U")
        )

    delitos = []
    for fila in range(3, 90):
        val = ws2[f"B{fila}"].value
        if val:
            delitos.append(normalizar(val))

    riesgos = []
    for fila in range(92, 155):
        val = ws2[f"B{fila}"].value
        if val:
            riesgos.append(normalizar(val))

    for fila in range(124, 141):
        try:
            ws1[f"K{fila}"].value = None  # Riesgos
            ws1[f"L{fila}"].value = None  # Delitos
        except:
            pass

    lista_riesgos = []
    lista_delitos = []

    for item in poder + conflicto:
        item_norm = normalizar(item)

        if item_norm in riesgos:
            lista_riesgos.append(item)
        elif item_norm in delitos:
            lista_delitos.append(item)

    fila_r = 124
    for r in lista_riesgos:
        if fila_r > 140:
            break
        try:
            ws1[f"K{fila_r}"].value = r
            fila_r += 1
        except:
            pass

    fila_d = 124
    for d in lista_delitos:
        if fila_d > 140:
            break
        try:
            ws1[f"L{fila_d}"].value = d
            fila_d += 1
        except:
            pass
