import numpy as np
import pandas as pd
from openpyxl import load_workbook


def _find_matrix_bounds(ws):
    """Encuentra fila/col de inicio de la matriz MICMAC en la hoja MATRIZ."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            # La celda superior izquierda de la matriz tiene None,
            # pero la fila de encabezados tiene None en col A y nombres cortos desde col B
            if cell.column == 1 and cell.value is None:
                continue
            # Buscamos la fila donde col A es None y col B tiene un nombre de variable
            pass

    # Estrategia: buscar fila donde col A = None y col B = string (encabezado de columnas)
    # y la siguiente fila col A = mismo string que col B de encabezado
    header_row = None
    for r in range(1, ws.max_row + 1):
        a_val = ws.cell(row=r, column=1).value
        b_val = ws.cell(row=r, column=2).value
        if a_val is None and isinstance(b_val, str) and len(b_val) > 0:
            # Verificar que la siguiente fila col A == b_val (confirma que es encabezado)
            next_a = ws.cell(row=r + 1, column=1).value
            if next_a == b_val:
                header_row = r
                break

    if header_row is None:
        raise ValueError("No se encontró la tabla MICMAC en la hoja MATRIZ.")

    # Leer nombres de variables desde encabezado (col B en adelante)
    variables = []
    col = 2
    while True:
        val = ws.cell(row=header_row, column=col).value
        if val is None:
            break
        variables.append(str(val).strip())
        col += 1

    data_start_row = header_row + 1
    n = len(variables)

    return data_start_row, n, variables


def _extract_matrix(ws, data_start_row, n):
    """Extrae la matriz numérica n×n desde la hoja."""
    matrix = np.zeros((n, n), dtype=float)
    for i in range(n):
        for j in range(n):
            val = ws.cell(row=data_start_row + i, column=2 + j).value
            if isinstance(val, (int, float)) and val is not None:
                matrix[i][j] = float(val)
    return matrix


def _load_descriptors(wb):
    """Carga el mapeo NOMBRE_CORTO → DESCRIPTOR desde la hoja DESCRIPTORES."""
    # El nombre de la hoja puede tener espacios al final
    sheet_name = next((s for s in wb.sheetnames if s.strip().upper() == "DESCRIPTORES"), None)
    if sheet_name is None:
        raise ValueError("No se encontró la hoja DESCRIPTORES en el archivo.")

    ws = wb[sheet_name]
    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        corto = row[0]
        descriptor = row[1]
        if corto and descriptor:
            mapping[str(corto).strip()] = str(descriptor).strip()
    return mapping


def _classify_micmac(variables, matrix, k=3):
    """
    Calcula M* = M + M² + M³ y clasifica variables por cuadrante.
    Corte: mediana de motricidad y mediana de dependencia (fiel a MICMAC).
    Retorna dict con listas: poder, conflicto, resultados, autonomas.
    """
    # Normalizar para evitar explosión numérica en potencias
    m_max = matrix.max()
    if m_max == 0:
        raise ValueError("La matriz MICMAC está vacía o todos los valores son cero.")

    M = matrix / m_max

    # Matriz acumulada M* = M + M² + M³
    M_acc = np.zeros_like(M)
    M_pot = M.copy()
    for _ in range(k):
        M_acc += M_pot
        M_pot = M_pot @ M

    # Motricidad = suma de filas, Dependencia = suma de columnas
    motricidad = M_acc.sum(axis=1)
    dependencia = M_acc.sum(axis=0)

    # Umbral: mediana (fiel a distribución relativa de MICMAC)
    umbral_I = np.median(motricidad)
    umbral_D = np.median(dependencia)

    cuadrantes = {"poder": [], "conflicto": [], "resultados": [], "autonomas": []}

    for i, var in enumerate(variables):
        alta_I = motricidad[i] >= umbral_I
        alta_D = dependencia[i] >= umbral_D

        if alta_I and not alta_D:
            cuadrantes["poder"].append(var)
        elif alta_I and alta_D:
            cuadrantes["conflicto"].append(var)
        elif not alta_I and alta_D:
            cuadrantes["resultados"].append(var)
        else:
            cuadrantes["autonomas"].append(var)

    return cuadrantes


def _write_to_engine(wb_engine, cuadrantes, descriptor_map):
    """
    Escribe los resultados en info_engine:
      B124 = Poder, C124 = Conflicto, D124 = Resultados, E124 = Autónomas
    Una variable por fila, bajando desde fila 124.
    Usa el DESCRIPTOR completo en lugar del nombre corto.
    """
    sheet = wb_engine.active

    col_map = {
        "poder":      2,   # B
        "conflicto":  3,   # C
        "resultados": 4,   # D
        "autonomas":  5,   # E
    }
    start_row = 124

    for cuadrante, col in col_map.items():
        for i, var_corto in enumerate(cuadrantes[cuadrante]):
            descriptor = descriptor_map.get(var_corto, var_corto)
            sheet.cell(row=start_row + i, column=col, value=descriptor)


def procesar_micmac(archivo_micmac, wb_engine):
    """
    Función principal llamada desde app.py.
    archivo_micmac: objeto file-like (BytesIO desde Streamlit)
    wb_engine: workbook openpyxl ya cargado de info_engine
    """
    wb = load_workbook(archivo_micmac)

    # Cargar mapeo de descriptores
    descriptor_map = _load_descriptors(wb)

    # Leer matriz MICMAC
    ws_matriz = wb["MATRIZ"]
    data_start_row, n, variables = _find_matrix_bounds(ws_matriz)
    matrix = _extract_matrix(ws_matriz, data_start_row, n)

    # Clasificar por cuadrantes
    cuadrantes = _classify_micmac(variables, matrix, k=3)

    # Escribir en info_engine
    _write_to_engine(wb_engine, cuadrantes, descriptor_map)
