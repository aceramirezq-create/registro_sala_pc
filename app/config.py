import os
import secrets


class Config:
    # Define SECRET_KEY como variable de entorno en producción.
    SECRET_KEY = os.environ.get("SECRET_KEY") or secrets.token_hex(32)

    EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "calendario_sala.xlsx")
    EXCEL_HOJA = os.environ.get("EXCEL_HOJA", "Hoja1")

    MAIL_SERVER = os.environ.get("MAIL_SERVER", "smtp.gmail.com")
    MAIL_PORT = int(os.environ.get("MAIL_PORT", "587"))
    MAIL_USE_TLS = os.environ.get("MAIL_USE_TLS", "true").lower() in {"1", "true", "yes", "on"}
    # Nunca guardar usuario/clave de correo en el código ni en GitHub.
    MAIL_USERNAME = os.environ.get("MAIL_USERNAME", "")
    MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD", "")
    ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")
