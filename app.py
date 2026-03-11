import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

st.title("Generador de Diagnóstico")

archivo_comunidad = st.file_uploader(
    "Subir archivo comunidad", type=["xlsx"]
)


# -----------------------------
# FUNCIONES GENERALES
# -----------------------------

def contar_frecuencias(df, columna, orden):
    serie = df[columna].dropna()
    conteo = serie.value_counts()
    return [conteo.get(opcion, 0) for opcion in orden]


def escribir_lista(ws, columna, fila_inicio, lista):
    for i, valor in enumerate(lista):
        ws[f"{columna}{fila_inicio+i}"] = int(valor)


def escribir_texto(ws, celda, texto):
    ws[celda] = texto


def limpiar_lista(ws, columna, fila_inicio, cantidad):
    for i in range(cantidad):
        ws[f"{columna}{fila_inicio+i}"] = None


# -----------------------------
# BOTON PRINCIPAL
# -----------------------------

if archivo_comunidad:

    if st.button("Generar diagnóstico"):

        df = pd.read_excel(archivo_comunidad)

        wb = load_workbook("plantillas/info_engine.xlsx")
        ws = wb["Hoja1"]


        # -----------------------------------
        # 1 CANTON
        # -----------------------------------

        canton_raw = df["T"].dropna().iloc[0]

        canton = canton_raw.replace("_", " ").title()

        canton = canton.replace(" Ramon", " Ramón")

        escribir_texto(ws, "B2", canton)


        # -----------------------------------
        # 2 DISTRITOS LISTA
        # -----------------------------------

        distritos = (
            df["U"]
            .dropna()
            .unique()
        )

        distritos = sorted(distritos)

        if len(distritos) == 16:

            for i, d in enumerate(distritos):
                ws[f"A{8+i}"] = d

        else:

            limpiar_lista(ws,"A",8,16)


        # -----------------------------------
        # 3 FRECUENCIA DISTRITOS
        # -----------------------------------

        conteo = df["U"].value_counts()

        if len(distritos) == 16:

            frecuencias = [conteo.get(d,0) for d in distritos]

            escribir_lista(ws,"E",8,frecuencias)

        else:

            limpiar_lista(ws,"E",8,16)


        # -----------------------------------
        # 4 RELACION CON ZONA
        # -----------------------------------

        orden = [
            "estudio_en_la_zona",
            "trabajo_en_la_zona",
            "visito_la_zona",
            "vivo_en_la_zona"
        ]

        frec = contar_frecuencias(df,"Y",orden)

        escribir_lista(ws,"I",8,frec)


        # -----------------------------------
        # 5 EDADES
        # -----------------------------------

        orden = [
            "De 18 a 29",
            "De 30 a 44",
            "De 45 a 64",
            "65 años o más",
            "Vacio"
        ]

        frec = contar_frecuencias(df,"V",orden)

        escribir_lista(ws,"D",29,frec)


        # -----------------------------------
        # 6 ESCOLARIDAD
        # -----------------------------------

        orden = [
            "ninguna",
            "primaria_completa",
            "primaria_incompleta",
            "secundaria_completa",
            "secundaria_incompleta",
            "tecnico",
            "Universidad_completa",
            "universidad_incompleta"
        ]

        frec = contar_frecuencias(df,"X",orden)

        escribir_lista(ws,"D",39,frec)


        # -----------------------------------
        # 7 GENERO
        # -----------------------------------

        orden = [
            "masculino",
            "femenino",
            "persona_no_binaria"
        ]

        frec = contar_frecuencias(df,"W",orden)

        escribir_lista(ws,"D",52,frec)


        # -----------------------------------
        # 8 SEGURIDAD ZONA
        # -----------------------------------

        orden = [
            "muy_inseguro",
            "inseguro",
            "ni_seguro_ni_inseguro",
            "seguro",
            "muy_seguro"
        ]

        frec = contar_frecuencias(df,"AA",orden)

        escribir_lista(ws,"C",283,frec)


        # -----------------------------------
        # 9 COMPARACION AÑO
        # -----------------------------------

        orden = [
            "mucho_menos_seguro",
            "menos_seguro",
            "se_mantiene_igual",
            "mas_seguro",
            "mucho_mas_seguro"
        ]

        frec = contar_frecuencias(df,"AD",orden)

        escribir_lista(ws,"C",291,frec)


        # -----------------------------------
        # 10 LUGARES INSEGUROS
        # -----------------------------------

        columnas = list(df.loc[:, "AF":"AQ"].columns)

        orden = [
            "muy_inseguro",
            "inseguro",
            "ni_seguro_ni_inseguro",
            "seguro",
            "muy_seguro",
            "no_aplica"
        ]

        destinos = ["B","D","F","H","J","L"]

        for op, col in zip(orden, destinos):

            resultados = []

            for c in columnas:

                resultados.append(
                    (df[c]==op).sum()
                )

            escribir_lista(ws,col,300,resultados)


        # -----------------------------------
        # 11 VICTIMA
        # -----------------------------------

        orden = [
            "no",
            "si_he_sido_víctima_pero_no_denuncie",
            "si_he_sido_víctima_y_si_denuncie"
        ]

        frec = contar_frecuencias(df,"BZ",orden)

        escribir_lista(ws,"D",314,frec)


        # -----------------------------------
        # 12 NO DENUNCIA
        # -----------------------------------

        orden = [
            "Distancia",
            "Miedo a represalias.",
            "Falta de respuesta oportuna.",
            "Complejidad al colocar la denuncia.",
            "Desconocimiento de dónde colocar la denuncia.",
            "El Policía me dijo que era mejor no denunciar.",
            "Falta de tiempo para colocar la denuncia",
            "Desconfianza en las autoridades o en el proceso de denuncia"
        ]

        frec = contar_frecuencias(df,"CF",orden)

        escribir_lista(ws,"D",322,frec)


        # -----------------------------------
        # 13 HORARIOS
        # -----------------------------------

        orden = [
            "00:00-02:59 a.m",
            "03:00-05:59 a.m",
            "06:00-08:59 a.m",
            "09:00-11:59 a.m",
            "12:00-14:59 p.m",
            "15:00-17:59 p.m",
            "18:00-20:59 p.m",
            "21:00-23:59 p.m",
            "Desconocido"
        ]

        frec = contar_frecuencias(df,"CH",orden)

        escribir_lista(ws,"D",336,frec)


        # -----------------------------------
        # 14 METODOLOGIA
        # -----------------------------------

        orden = [
            "Arma blanca (cuchillo, machete, tijeras).",
            "Arma de fuego.",
            "Amenazas",
            "Arrebato",
            "Boquete",
            "Ganzúa (pata de chancho)",
            "Engaño",
            "Escalamiento",
            "No sé",
            "Otro"
        ]

        frec = contar_frecuencias(df,"CI",orden)

        escribir_lista(ws,"D",350,frec)


        # -----------------------------------
        # GUARDAR ARCHIVO
        # -----------------------------------

        archivo = io.BytesIO()

        wb.save(archivo)

        archivo.seek(0)

        st.download_button(
            "Descargar info_engine generado",
            archivo,
            file_name="info_engine_resultado.xlsx"
        )
