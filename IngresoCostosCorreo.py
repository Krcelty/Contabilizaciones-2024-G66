import imaplib
import email
from email.header import decode_header
import openpyxl

# Parámetros de conexión
IMAP_SERVER = "imap.gmail.com"  # Cambiar al servidor IMAP correcto
EMAIL_ACCOUNT = "constanza.perez@global66.com"
PASSWORD = "oxen tifh gizb sgtm"
EMAIL_SUBJECT = "Ingresos y costos operativos - Octubre 2024"  

def conectar_email():
    # Conectar a la bandeja de entrada de IMAP
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, PASSWORD)
    return mail

def buscar_correo(mail, asunto):
    # Seleccionar la bandeja de entrada
    mail.select("inbox")
    # Buscar correo con el asunto específico
    status, mensajes = mail.search(None, f'SUBJECT "{asunto}"')
    if status != "OK":
        print("No se encontraron correos con el asunto especificado.")
        return None

    # Obtener el ID del primer correo que coincide con el asunto
    mensaje_ids = mensajes[0].split()
    if mensaje_ids:
        return mensaje_ids[0]
    else:
        return None

def leer_correo(mail, mensaje_id):
    # Obtener el correo completo por su ID
    status, mensaje_datos = mail.fetch(mensaje_id, "(RFC822)")
    if status != "OK":
        print("Error al obtener el correo.")
        return None

    # Decodificar y analizar el mensaje
    mensaje = email.message_from_bytes(mensaje_datos[0][1])

    # Asegurarse de que el mensaje no es None
    if not mensaje:
        print("El mensaje no se ha podido obtener correctamente.")
        return None

    contenido = ""

    # Si el correo tiene varias partes (texto y HTML, por ejemplo)
    if mensaje.is_multipart():
        for parte in mensaje.walk():
            # Usar solo la parte de texto (descartar HTML)
            if parte.get_content_type() == "text/plain":
                # Obtener el payload de la parte (que es el texto del correo)
                contenido += parte.get_payload(decode=True).decode('utf-8', errors='ignore')
    else:
        # Si no es multipart, simplemente obtener el payload
        contenido = mensaje.get_payload(decode=True).decode('utf-8', errors='ignore')

    return contenido

def guardar_en_excel(texto):
    # Crear un archivo Excel y una hoja
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contenido del Correo"

    # Guardar cada línea del contenido en una fila de Excel
    for i, linea in enumerate(texto.splitlines(), start=1):
        ws.cell(row=i, column=1, value=linea)

    # Guardar el archivo
    wb.save("contenido_correo.xlsx")
    print("Archivo Excel creado con éxito.")

def main():
    # Conectar al servidor de correo y buscar el correo deseado
    mail = conectar_email()
    mensaje_id = buscar_correo(mail, EMAIL_SUBJECT)

    if mensaje_id:
        # Leer el correo y guardar en Excel
        contenido = leer_correo(mail, mensaje_id)
        if contenido:
            guardar_en_excel(contenido)
    else:
        print("No se encontró el correo con el asunto especificado.")

    # Cerrar la conexión
    mail.logout()

if __name__ == "__main__":
    main()
