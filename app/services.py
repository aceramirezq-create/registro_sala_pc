"""Lógica de reserva automática y notificaciones por correo."""
import smtplib
from email.message import EmailMessage

from flask import current_app
from excel_utils import reservar_horario


def enviar_correo(destinatario, nombre_docente, dia, bloque, aprobado):
    """Envía resultado de la solicitud. Devuelve False si SMTP no está configurado/falla."""
    usuario = current_app.config.get("MAIL_USERNAME", "")
    clave = current_app.config.get("MAIL_PASSWORD", "")
    if not usuario or not clave:
        current_app.logger.error("Correo no enviado: faltan MAIL_USERNAME o MAIL_PASSWORD.")
        return False

    mensaje = EmailMessage()
    mensaje["From"] = usuario
    mensaje["To"] = destinatario
    if aprobado:
        mensaje["Subject"] = "Confirmada: reserva de Sala de Computación"
        cuerpo = (
            f"Hola {nombre_docente},\n\n"
            "Tu reserva de la Sala de Computación fue CONFIRMADA y registrada en el calendario.\n\n"
            f"Día: {dia}\nBloque: {bloque}\n\n"
            "Saludos,\nSistema de Gestión de Sala de Computación"
        )
    else:
        mensaje["Subject"] = "Rechazada: reserva de Sala de Computación"
        cuerpo = (
            f"Hola {nombre_docente},\n\n"
            "No fue posible realizar tu reserva porque el bloque solicitado ya está ocupado.\n\n"
            f"Día: {dia}\nBloque: {bloque}\n\n"
            "Por favor, elige otro bloque disponible.\n\n"
            "Saludos,\nSistema de Gestión de Sala de Computación"
        )
    mensaje.set_content(cuerpo)

    try:
        with smtplib.SMTP(
            current_app.config["MAIL_SERVER"],
            current_app.config["MAIL_PORT"],
            timeout=20,
        ) as servidor:
            if current_app.config.get("MAIL_USE_TLS", True):
                servidor.starttls()
            servidor.login(usuario, clave)
            servidor.send_message(mensaje)
        return True
    except (OSError, smtplib.SMTPException) as error:
        current_app.logger.exception("No se pudo enviar correo a %s: %s", destinatario, error)
        return False


def procesar_reserva(nombre_docente, correo_docente, dia, bloque):
    """Reserva atómicamente y notifica el resultado real al solicitante."""
    archivo = current_app.config["EXCEL_PATH"]
    hoja = current_app.config["EXCEL_HOJA"]

    # No hacer una consulta previa separada: la comprobación y escritura ocurren
    # dentro de reservar_horario, bajo el mismo bloqueo entre procesos.
    reservado = reservar_horario(archivo, hoja, dia, bloque, nombre_docente)
    correo_enviado = enviar_correo(correo_docente, nombre_docente, dia, bloque, reservado)

    if reservado and correo_enviado:
        return True, f"Reserva confirmada. Se envió el correo a {correo_docente}."
    if reservado:
        return True, "La reserva quedó registrada en Excel, pero no se pudo enviar el correo. Avise al administrador para revisar la configuración SMTP."
    if correo_enviado:
        return False, f"El bloque {bloque} del día {dia} ya está ocupado. Se envió un correo de rechazo a {correo_docente}."
    return False, f"El bloque {bloque} del día {dia} ya está ocupado. No se pudo enviar el correo de rechazo; avise al administrador."
