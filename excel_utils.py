from openpyxl import load_workbook

def horario_disponible(archivo, hoja, dia, bloque):
    """Verifica si el bloque horario está libre en la matriz."""
    libro = load_workbook(archivo, read_only=True, data_only=True)
    try:
        planilla = libro[hoja]
        filas = list(planilla.iter_rows(values_only=True))
        
        if not filas:
            return False

        encabezado = filas[0]
        columna_dia_idx = None
        for i, col in enumerate(encabezado):
            if col and str(col).strip().lower() == str(dia).strip().lower():
                columna_dia_idx = i
                break

        if columna_dia_idx is None:
            return False

        for fila in filas[1:]:
            bloque_excel = str(fila[0]).strip() if fila[0] else ""
            if bloque_excel == str(bloque).strip():
                val_celda = fila[columna_dia_idx]
                return val_celda is None or str(val_celda).strip() == ""

        return False
    finally:
        libro.close()


def reservar_horario(archivo, hoja, dia, bloque, nombre_docente):
    """Escribe el nombre del docente en el bloque si está libre."""
    libro = load_workbook(archivo)
    try:
        planilla = libro[hoja]
        filas = list(planilla.iter_rows())
        if not filas:
            return False

        encabezado = [celda.value for celda in filas[0]]
        columna_dia_idx = None
        for i, col in enumerate(encabezado):
            if col and str(col).strip().lower() == str(dia).strip().lower():
                columna_dia_idx = i
                break

        if columna_dia_idx is None:
            return False

        for fila in filas[1:]:
            bloque_excel = str(fila[0].value).strip() if fila[0].value else ""
            if bloque_excel == str(bloque).strip():
                celda_objetivo = fila[columna_dia_idx]
                if celda_objetivo.value is None or str(celda_objetivo.value).strip() == "":
                    celda_objetivo.value = nombre_docente
                    libro.save(archivo)
                    return True
                return False

        return False
    finally:
        libro.close()