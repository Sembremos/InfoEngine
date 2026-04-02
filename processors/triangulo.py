from openpyxl import load_workbook

def procesar_triangulo(archivo_triangulo, wb_destino):
    wb_origen = load_workbook(archivo_triangulo, data_only=True)

    total_socio = 0
    total_estructural = 0

    for hoja in wb_origen.worksheets:
        # Buscar encabezados en fila 4
        fila_encabezados = 4

        col_socio = None
        col_estructural = None

        for col in range(1, hoja.max_column + 1):
            valor = hoja.cell(row=fila_encabezados, column=col).value

            if valor:
                valor = str(valor).strip()

                if "Socio" in valor:
                    col_socio = col
                elif "Estructurales" in valor:
                    col_estructural = col

        # Si no encuentra ambas columnas, saltar hoja
        if not col_socio or not col_estructural:
            continue

        # Leer desde fila 5 hacia abajo
        fila = 5

        while True:
            val_socio = hoja.cell(row=fila, column=col_socio).value
            val_estructural = hoja.cell(row=fila, column=col_estructural).value

            # Condición de corte: ambas vacías (fin de tabla)
            if not val_socio and not val_estructural:
                break

            # Contar solo si hay contenido real
            if val_socio:
                total_socio += 1

            if val_estructural:
                total_estructural += 1

            fila += 1

    # Escribir en el archivo destino
    hoja_destino = wb_destino["Hoja1"]

    hoja_destino["B147"] = total_socio
    hoja_destino["C147"] = total_estructural
