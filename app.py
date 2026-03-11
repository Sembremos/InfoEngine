import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

st.title("Generador de info_engine desde Encuesta Comunidad")

archivo_comunidad = st.file_uploader("Subir archivo comunidad", type=["xlsx"])


# --------------------------------------
# FUNCIONES GENERALES
# --------------------------------------

def contar_frecuencias(df, columna, orden):
    serie = df[columna].dropna()
    conteo = serie.value_counts()
    return [conteo.get(opcion, 0) for opcion in orden]


def escribir_lista(ws, columna, fila_inicio, lista):
    for i, valor in enumerate(lista):
        ws[f"{columna}{fila_inicio+i}"] = int(valor)


def limpiar_lista(ws, columna, fila_inicio, cantidad):
    for i in range(cantidad):
        ws[f"{columna}{fila_inicio+i}"] = None


def formatear_canton(texto):
    texto = str(texto).replace("_", " ").title()
    texto = texto.replace(" Ramon", " Ramón")
    return texto


# --------------------------------------
# BOTON PRINCIPAL
# --------------------------------------

if archivo_comunidad:

    if st.button("Generar info_engine"):

        df = pd.read_excel(archivo_comunidad)

        wb = load_workbook("plantillas/info_engine.xlsx")
        ws = wb["Hoja1"]

        # -----------------------------------
        # 1 CANTON
        # -----------------------------------

        canton_raw = df["1. Cantón"].dropna().iloc[0]
        ws["B2"] = formatear_canton(canton_raw)

        # -----------------------------------
        # 2 LISTA DISTRITOS
        # -----------------------------------

        distritos = sorted(df["2. Distrito:"].dropna().unique())

        if len(distritos) == 16:

            for i, d in enumerate(distritos):
                ws[f"A{8+i}"] = d

            conteo = df["2. Distrito:"].value_counts()
            frecuencias = [conteo.get(d, 0) for d in distritos]

            escribir_lista(ws, "E", 8, frecuencias)

        else:

            limpiar_lista(ws, "A", 8, 16)
            limpiar_lista(ws, "E", 8, 16)

        # -----------------------------------
        # 4 RELACION ZONA
        # -----------------------------------

        orden = [
            "estudio_en_la_zona",
            "trabajo_en_la_zona",
            "visito_la_zona",
            "vivo_en_la_zona"
        ]

        frec = contar_frecuencias(df, "6. ¿Cuál es su relación con la zona?", orden)

        escribir_lista(ws, "I", 8, frec)

        # -----------------------------------
        # 5 EDAD
        # -----------------------------------

        orden = [
            "De 18 a 29",
            "De 30 a 44",
            "De 45 a 64",
            "65 años o más",
            "Vacio"
        ]

        frec = contar_frecuencias(
            df,
            "3. Edad (en años cumplidos): marque una categoría que incluya su edad.",
            orden
        )

        escribir_lista(ws, "D", 29, frec)

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

        frec = contar_frecuencias(df, "5. Escolaridad:", orden)

        escribir_lista(ws, "D", 39, frec)

        # -----------------------------------
        # 7 GENERO
        # -----------------------------------

        orden = [
            "masculino",
            "femenino",
            "persona_no_binaria"
        ]

        frec = contar_frecuencias(
            df,
            "4. ¿Con cuál de estas opciones se identifica?",
            orden
        )

        escribir_lista(ws, "D", 52, frec)

        # -----------------------------------
        # 8 SEGURIDAD DISTRITO
        # -----------------------------------

        orden = [
            "muy_inseguro",
            "inseguro",
            "ni_seguro_ni_inseguro",
            "seguro",
            "muy_seguro"
        ]

        frec = contar_frecuencias(
            df,
            "7. ¿Qué tan seguro percibe usted el distrito donde reside o transita?",
            orden
        )

        escribir_lista(ws, "C", 283, frec)

        # -----------------------------------
        # 9 CAMBIO SEGURIDAD
        # -----------------------------------

        orden = [
            "mucho_menos_seguro",
            "menos_seguro",
            "se_mantiene_igual",
            "mas_seguro",
            "mucho_mas_seguro"
        ]

        frec = contar_frecuencias(
            df,
            "8. En comparación con los 12 meses anteriores, ¿cómo percibe que ha cambiado la seguridad en este distrito?",
            orden
        )

        escribir_lista(ws, "C", 291, frec)

        # -----------------------------------
        # 10 SEGURIDAD LUGARES
        # -----------------------------------

        columnas = [
            "seg_discotecas_bares",
            "seg_espacios_recreativos",
            "seg_lugar_residencia",
            "seg_paradas_estaciones",
            "seg_puentes_peatonales",
            "seg_transporte_publico",
            "seg_zona_bancaria",
            "seg_zona_comercio",
            "seg_zonas_residenciales",
            "seg_zonas_francas",
            "seg_lugares_turisticos",
            "seg_centros_educativos"
        ]

        orden = [
            "muy_inseguro",
            "inseguro",
            "ni_seguro_ni_inseguro",
            "seguro",
            "muy_seguro",
            "no_aplica"
        ]

        columnas_destino = ["B", "D", "F", "H", "J", "L"]

        for opcion, col_excel in zip(orden, columnas_destino):

            resultados = []

            for c in columnas:
                resultados.append((df[c] == opcion).sum())

            escribir_lista(ws, col_excel, 300, resultados)

        # -----------------------------------
        # 11 VICTIMIZACION
        # -----------------------------------

        orden = [
            "no",
            "si_he_sido_víctima_pero_no_denuncie",
            "si_he_sido_víctima_y_si_denuncie"
        ]

        frec = contar_frecuencias(
            df,
            "30. Durante los últimos 12 meses, ¿usted o algún miembro de su hogar fue afectado por algún delito?",
            orden
        )

        escribir_lista(ws, "D", 314, frec)

        # -----------------------------------
        # 12 MOTIVOS NO DENUNCIA
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

        frec = contar_frecuencias(
            df,
            "30.2 En caso de NO haber realizado la denuncia, indique ¿cuál o cuáles fueron el motivo?",
            orden
        )

        escribir_lista(ws, "D", 322, frec)

        # -----------------------------------
        # 13 HORARIO DELITO
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

        frec = contar_frecuencias(
            df,
            "30.3 ¿Tiene conocimiento sobre el horario en el cual se presentó el hecho o situación que le afectó a usted o un familiar?",
            orden
        )

        escribir_lista(ws, "D", 336, frec)

        # -----------------------------------
        # 14 METODOLOGIA DELITO
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

        frec = contar_frecuencias(
            df,
            "30.4 ¿Cuál fue la forma o modo en que ocurrió la situación que afectó a usted o a algún miembro de su hogar?",
            orden
        )

        escribir_lista(ws, "D", 350, frec)

        archivo = io.BytesIO()
        wb.save(archivo)
        archivo.seek(0)

        st.download_button(
            "Descargar info_engine generado",
            archivo,
            file_name="info_engine_resultado.xlsx"
        )
