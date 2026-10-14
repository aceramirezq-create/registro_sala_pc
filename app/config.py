import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'clave_secreta_proyecto_sala_pc'
    
    # Ruta absoluta hacia el archivo subido por tu compañero
    EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'calendario_sala.xlsx')
    
    # Nombre exacto de la hoja detectado en tu archivo
    EXCEL_HOJA = 'Hoja1'

    # Configuración de correo SMTP (Ajustar según corresponda)
    MAIL_SERVER = 'smtp.gmail.com'
    MAIL_PORT = 587
    MAIL_USE_TLS = True
    MAIL_USERNAME = os.environ.get('MAIL_USERNAME') or 'a_ceramirezq@colegiosanlorenzo.org'
    MAIL_PASSWORD = os.environ.get('MAIL_PASSWORD') or 'bixqrlhdkhdywbsn'