from excel_utils import obtener_horarios_libres

archivo = "calendario_sala.xlsx"
hoja = "Hoja1"

horarios = obtener_horarios_libres(archivo, hoja)

if horarios:
    print("\nHORARIOS DISPONIBLES DE LA SALA\n")

    for horario in horarios:
        print(
            f"{horario['dia']} - "
            f"{horario['bloque']}"
        )
else:
    print("No se encontraron horarios disponibles.")
