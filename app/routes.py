
"""Rutas web para reserva automática y consulta de ocupación."""

import re
import hmac
import unicodedata

from filelock import Timeout
from flask import (
    Blueprint,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    url_for,
)

from app.services import procesar_reserva
from excel_utils import obtener_ocupacion, liberar_horario

main_bp = Blueprint("main", __name__)

DIAS_PERMITIDOS = {
    "Lunes",
    "Martes",
    "Miércoles",
    "Jueves",
    "Viernes",
}

BLOQUES_PERMITIDOS = {
    "08:30 - 10:00",
    "10:15 - 11:45",
    "12:00 - 14:15",
    "15:00 - 17:15",
}


def normalizar_texto(valor):
    """Normaliza texto ignorando mayúsculas, tildes y espacios repetidos."""
    texto = str(valor or "").strip().casefold()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )
    return " ".join(texto.split())


def buscar_valor_permitido(valor, valores_permitidos):
    """Devuelve el valor oficial que coincide, ignorando tildes y mayúsculas."""
    valor_normalizado = normalizar_texto(valor)

    for permitido in valores_permitidos:
        if normalizar_texto(permitido) == valor_normalizado:
            return permitido

    return None


def _leer_ocupacion():
    return obtener_ocupacion(
        current_app.config["EXCEL_PATH"],
        current_app.config["EXCEL_HOJA"],
    )


@main_bp.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nombre = (request.form.get("nombre") or "").strip()
        correo = (request.form.get("correo") or "").strip()
        dia_recibido = (request.form.get("dia") or "").strip()
        bloque_recibido = (request.form.get("bloque") or "").strip()

        dia = buscar_valor_permitido(dia_recibido, DIAS_PERMITIDOS)
        bloque = buscar_valor_permitido(
            bloque_recibido,
            BLOQUES_PERMITIDOS,
        )

        if not nombre or len(nombre) > 120:
            flash(
                "Ingresa un nombre válido (máximo 120 caracteres).",
                "danger",
            )
        elif (
            len(correo) > 254
            or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", correo)
        ):
            flash("Ingresa un correo electrónico válido.", "danger")
        elif dia is None or bloque is None:
            flash(
                "Selecciona un día y bloque de la lista disponible.",
                "danger",
            )
        else:
            try:
                exito, mensaje = procesar_reserva(
                    nombre,
                    correo,
                    dia,
                    bloque,
                )
                flash(mensaje, "success" if exito else "danger")
            except (OSError, ValueError, Timeout):
                current_app.logger.exception(
                    "No se pudo procesar la reserva."
                )
                flash(
                    "No se pudo procesar la reserva por un problema "
                    "con el calendario. Inténtalo más tarde o avisa "
                    "al administrador.",
                    "danger",
                )

        return redirect(url_for("main.index"))

    try:
        ocupacion = _leer_ocupacion()
        error_ocupacion = None
    except (OSError, ValueError, Timeout):
        current_app.logger.exception(
            "No se pudo leer el calendario."
        )
        ocupacion = {"dias": [], "horarios": []}
        error_ocupacion = (
            "No se pudo leer el calendario Excel. Revisa la ruta, "
            "el nombre de la hoja y que el archivo no esté dañado."
        )

    return render_template(
        "index.html",
        ocupacion=ocupacion,
        error_ocupacion=error_ocupacion,
    )


@main_bp.route("/liberar", methods=["POST"])
def liberar():
    dia_recibido = (request.form.get("dia") or "").strip()
    bloque_recibido = (request.form.get("bloque") or "").strip()
    clave_ingresada = request.form.get("admin_password", "")
    clave_admin = current_app.config.get("ADMIN_PASSWORD", "")

    if (
        not clave_admin
        or not hmac.compare_digest(clave_ingresada, clave_admin)
    ):
        flash(
            "Contraseña de administrador incorrecta o no configurada.",
            "danger",
        )
        return redirect(url_for("main.index"))

    dia = buscar_valor_permitido(dia_recibido, DIAS_PERMITIDOS)
    bloque = buscar_valor_permitido(
        bloque_recibido,
        BLOQUES_PERMITIDOS,
    )

    if dia is None or bloque is None:
        flash("Selecciona un día y bloque válidos.", "danger")
        return redirect(url_for("main.index"))

    try:
        liberado = liberar_horario(
            current_app.config["EXCEL_PATH"],
            current_app.config["EXCEL_HOJA"],
            dia,
            bloque,
        )

        if liberado:
            flash(
                f"Se liberó el bloque {bloque} del día {dia}.",
                "success",
            )
        else:
            flash("Ese bloque ya estaba libre.", "danger")

    except (OSError, ValueError, Timeout):
        current_app.logger.exception(
            "No se pudo liberar el bloque."
        )
        flash(
            "No se pudo liberar el bloque. Revisa el calendario.",
            "danger",
        )

    return redirect(url_for("main.index"))
