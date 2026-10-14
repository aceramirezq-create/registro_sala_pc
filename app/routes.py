from flask import Blueprint, render_template, request, flash, current_app
from app.services import procesar_reserva, horario_disponible

main_bp = Blueprint('main', __name__)

@main_bp.route('/', methods=['GET', 'POST'])
def index():
    estado_disponibilidad = None
    datos_previos = {}

    if request.method == 'POST':
        accion = request.form.get('accion')
        nombre = request.form.get('nombre')
        correo = request.form.get('correo')
        dia = request.form.get('dia')
        bloque = request.form.get('bloque')

        datos_previos = {
            'nombre': nombre,
            'correo': correo,
            'dia': dia,
            'bloque': bloque
        }

        # 1. Comprobar disponibilidad
        if accion == 'comprobar':
            libre = horario_disponible(
                current_app.config['EXCEL_PATH'],
                current_app.config['EXCEL_HOJA'],
                dia,
                bloque
            )
            if libre:
                estado_disponibilidad = 'DISPONIBLE'
                flash(f"El bloque {bloque} del día {dia} está LIBRE. Puedes confirmar la reserva.", "success")
            else:
                estado_disponibilidad = 'OCUPADO'
                flash(f"El bloque {bloque} del día {dia} ya se encuentra OCUPADO.", "danger")

        # 2. Confirmar y guardar reserva
        elif accion == 'reservar':
            exito, mensaje = procesar_reserva(nombre, correo, dia, bloque)
            if exito:
                flash(mensaje, 'success')
                datos_previos = {}  # Limpiar formulario
            else:
                flash(mensaje, 'danger')

    return render_template('index.html', estado=estado_disponibilidad, datos=datos_previos)