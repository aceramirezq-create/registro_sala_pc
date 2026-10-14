import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from flask import current_app
from excel_utils import horario_disponible, reservar_horario

def enviar_correo_confirmacion(destinatario, nombre_docente, dia, bloque, estado):
    """Envía un correo electrónico automático al docente solicitante."""
    servidor_email = current_app.config['MAIL_USERNAME']
    password_email = current_app.config['MAIL_PASSWORD']

    msg = MIMEMultipart()
    msg['From'] = servidor_email
    msg['To'] = destinatario

    if estado == "APROBADO":
        msg['Subject'] = "Aprobada: Reserva de Sala de Computación"
        cuerpo = f"""
        Hola {nombre_docente},

        Tu solicitud para utilizar la Sala de Computación ha sido CONFIRMADA y registrada con éxito.

        Detalles de la reserva:
        - Día: {dia}
        - Bloque Horario: {bloque}

        Por favor, asegúrate de dejar la sala ordenada al finalizar el bloque.

        Saludos cordiales,
        Sistema de Gestión de Sala de Computación
        """
    else:
        msg['Subject'] = "Rechazada: Reserva de Sala de Computación"
        cuerpo = f"""
        Hola {nombre_docente},

        Lamentamos informarte que la reserva NO se pudo realizar.
        El bloque horario {bloque} del día {dia} ya se encuentra ocupado.

        Por favor ingresa a la plataforma para verificar otros bloques disponibles.

        Saludos cordiales,
        Sistema de Gestión de Sala de Computación
        """

    msg.attach(MIMEText(cuerpo, 'plain'))

    try:
        server = smtplib.SMTP(current_app.config['MAIL_SERVER'], current_app.config['MAIL_PORT'])
        server.starttls()
        server.login(servidor_email, password_email)
        server.send_message(msg)
        server.quit()
        print(f"Correo enviado exitosamente a {destinatario}")
        return True
    except Exception as e:
        print(f"Error al enviar el correo: {e}")
        return False


def procesar_reserva(nombre_docente, correo_docente, dia, bloque):
    """Maneja la verificación, reserva en Excel y el envío del correo."""
    archivo = current_app.config['EXCEL_PATH']
    hoja = current_app.config['EXCEL_HOJA']

    # 1. Verificar si está disponible
    disponible = horario_disponible(archivo, hoja, dia, bloque)

    if disponible:
        # 2. Registrar en el archivo Excel
        reservado = reservar_horario(archivo, hoja, dia, bloque, nombre_docente)
        if reservado:
            # 3. Enviar correo de éxito
            enviado = enviar_correo_confirmacion(correo_docente, nombre_docente, dia, bloque, "APROBADO")
            if enviado:
                return True, f"¡Reserva confirmada! Se ha enviado el correo a {correo_docente}."
            else:
                return True, "Reserva registrada en Excel, pero hubo un error al enviar el correo (revisa la clave SMTP)."

    # Si el horario ya está ocupado
    enviar_correo_confirmacion(correo_docente, nombre_docente, dia, bloque, "RECHAZADO")
    return False, f"El bloque {bloque} del día {dia} ya está ocupado."