import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Iniciamos los parámetros del script

def Envio_pdf(dir_destino, e_vendedor, e_ejecutivo, e_asunto, PdfFilename):
    folderOps = dir_destino
    mail_vendedor = str(e_vendedor)
    mail_ejecutivo = str(e_ejecutivo)

    remitente = 'notificaciones@jiroseguros.com.mx'
    destinatarios = [mail_vendedor, 'mesadecontrol.leon@jiro.mx']
    asunto = e_asunto 
    cuerpo = 'FAVOR DE NO CONTESTAR ESTE MENSAJE, esta cuenta es solo para enviar notificaciones.\n\nEstimado agente Jiro, se anexa detalle de comisiones a pagar ésta semana.\nAntes de facturar favor de confirmar con Brenda Ramírez,\n correo: mesadecontrol.leon@jiro.mx\n\nCualquier aclaración al correo sobre las comisiones reflejadas en este correo escribir a la dirección:\nmesadecontrol.leon@jiro.mx\n\nMesa de Control'
    nombre_adjunto = PdfFilename
    ruta_adjunto = folderOps + "/final/" + nombre_adjunto

    # Creamos el objeto mensaje
    mensaje = MIMEMultipart()
    
    # Establecemos los atributos del mensaje
    mensaje['From'] = remitente
    mensaje['To'] = ", ".join(destinatarios)
    mensaje['Subject'] = asunto
    
    # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
    mensaje.attach(MIMEText(cuerpo, 'plain'))
    
    # Abrimos el archivo que vamos a adjuntar
    archivo_adjunto = open(ruta_adjunto, 'rb')
    
    # Creamos un objeto MIME base
    adjunto_MIME = MIMEBase('application', 'octet-stream')
    # Y le cargamos el archivo adjunto
    adjunto_MIME.set_payload((archivo_adjunto).read())
    # Codificamos el objeto en BASE64
    encoders.encode_base64(adjunto_MIME)
    # Agregamos una cabecera al objeto
    adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
    # Y finalmente lo agregamos al mensaje
    mensaje.attach(adjunto_MIME)
    
    # Creamos la conexión con el servidor
    sesion_smtp = smtplib.SMTP('smtp.1and1.mx', 587)
    
    # Ciframos la conexión
    sesion_smtp.starttls()

    # Iniciamos sesión en el servidor
    sesion_smtp.login('jironotificaciones@gmail.com','jiromar*22')

    # Convertimos el objeto mensaje a texto
    texto = mensaje.as_string()

    # Enviamos el mensaje
    sesion_smtp.sendmail(remitente, destinatarios, texto)

    # Cerramos la conexión
    sesion_smtp.quit()
