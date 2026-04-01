def escribir_cuadrantes_manual(wb, poder, conflicto, resultados, autonomas):

    ws = wb.active  # o especificá hoja si aplica

    # limpiar rango
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

    escribir(poder, "B")
    escribir(conflicto, "C")
    escribir(resultados, "D")
    escribir(autonomas, "E")
