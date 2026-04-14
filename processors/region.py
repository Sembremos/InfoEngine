import re

def obtener_numero_delegacion(valor):
    if not valor:
        return None

    texto = str(valor)

    # Busca cualquier número dentro del texto (más confiable)
    numeros = re.findall(r'\d+', texto)

    if numeros:
        return int(numeros[0])  # toma el primer número completo

    return None

def mapear_region(numero):
    """
    Retorna el número de región según el número de delegación
    """

    if numero in range(0, 26):
        return 1

    elif numero in range(26, 37):
        return 2

    elif numero in range(37, 49):
        return 3

    elif numero in [49,50,51,52,53,54,56,57]:
        return 4

    elif numero in [60,61,62,63,64,65,66,67,68,69]:
        return 5

    elif numero in [71,72,73,74,75,76,77,78]:
        return 6

    elif numero in [79,80,81]:
        return 7

    elif numero in [82,83,86,87]:
        return 8

    elif numero in [88,90,91,92]:
        return 9

    elif numero in [94,95,96,97]:
        return 10

    elif numero in [70,84,85]:
        return 11

    elif numero in [89,93,58,59,98]:
        return 12

    else:
        return None


def escribir_region(wb):
    """
    Lee B3 (delegación), calcula la región y la escribe en D2
    """
    hoja = wb["Hoja1"]

    valor_b3 = hoja["B3"].value
    numero = obtener_numero_delegacion(valor_b3)

    if numero is None:
        return

    region = mapear_region(numero)

    if region is not None:
        hoja["D2"] = region
