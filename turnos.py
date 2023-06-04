import os
import pandas as pd
import openpyxl
import smtplib
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import shutil
from dotenv import load_dotenv
import os
 
load_dotenv()

# Datos del correo electrónico
destinatario =  os.environ.get("DESTINATARIO")
remitente = os.environ.get("REMITENTE")
contraseña = os.environ.get("PASSWORD")
servidor_smtp = os.environ.get("SERVER")
puerto_smtp = os.environ.get("PUERTO")
asunto = 'Archivo CSV adjunto'
cuerpo = 'Adjunto encontrarás el archivo CSV que solicitaste.'


destino_archivo = '/home/fer/projects/python/turnos_Vic/files/archivosRecibidos/'

def comprobar_ficheros_descargados():
    source_folder = '/home/fer/Descargas'
    final_path_file = ''
    for root, dirs, files in os.walk(source_folder):
        if root == source_folder:
            for file_name in files:
                source_path = os.path.join(root, file_name)
                if source_path.endswith(".xlsx"):
                    final_path_file = source_path
                    # print(f"Es un excel {final_path_file}")
                    data = leer_excel_load(final_path_file)
                    archivo_csv = verificar_data(data, source_path)
                    mover_archivo_descargado()
                    shutil.move(source_path, destino_archivo)
                    # enviar_email(destinatario, asunto, cuerpo, archivo_csv, remitente, contraseña, servidor_smtp, puerto_smtp)

def mover_archivo_descargado():
    destino_archivo = '/home/fer/projects/python/turnos_Vic/files/archivosRecibidos/'
    for root, dirs, files in os.walk(destino_archivo):
        if root == destino_archivo:
            for file_name in files:
                source_path = os.path.join(root, file_name)
                if source_path.endswith(".xlsx"):
                    os.remove(source_path)
                    print('Archivo : ', file_name , ' borrado correctamente' )

            
def leer_excel_load(nombre_archivo):
    try:
        libro_excel = openpyxl.load_workbook(nombre_archivo)
        hoja_activa = libro_excel.active
        datos = []
        contador = 0
        for fila in hoja_activa.iter_rows(values_only=True):
            if contador == 0:
                datos.append(fila)
                contador += 1
                continue
            if contador == 1:
                datos.append(fila)
                contador += 1
                continue
            if fila[0] == 'VICTORIA':
                datos.append(fila)
                break
        # print(datos)
        return datos
    except FileNotFoundError:
        print("No se encontró el archivo:", nombre_archivo)
    except Exception as e:
        print("Ocurrió un error al leer el archivo:", str(e))


def verificar_data(data, path):
    dias = data[0]
    dias_numero = data[1]
    turno = data[2]
    nombre_archivo = '/home/fer/projects/python/turnos_Vic/files/turnosProcesados/Turnos.csv' 
    datos = []
    for i in range(len(dias)):
        if i == 0:
            datos.append(['Día     Día Semana   Turno de ' +turno[i]])
            continue
        datos.append([f"{dias_numero[i]}-> {dias[i]}       {turno[i]}"])
    try:
        with open(nombre_archivo, 'w', newline='') as archivo_csv:
            escritor = csv.writer(archivo_csv)
            for fila in datos:
                escritor.writerow(fila)
        print("Archivo CSV creado exitosamente:", nombre_archivo)
        return nombre_archivo
    except Exception as e:
        print("Ocurrió un error al crear el archivo CSV:", str(e))


def enviar_email(destinatario, asunto, cuerpo, archivo_adjunto, remitente, contraseña, servidor_smtp, puerto_smtp):
    mensaje = MIMEMultipart()
    mensaje['From'] = remitente
    mensaje['To'] = destinatario
    mensaje['Subject'] = asunto

    mensaje.attach(MIMEBase('application', 'octet-stream'))
    with open(archivo_adjunto, 'rb') as adjunto:
        parte_adjunto = MIMEBase('application', 'octet-stream')
        parte_adjunto.set_payload(adjunto.read())
        encoders.encode_base64(parte_adjunto)
        parte_adjunto.add_header('Content-Disposition', f'attachment; filename= {archivo_adjunto}')
        mensaje.attach(parte_adjunto)

    try:
        servidor = smtplib.SMTP(servidor_smtp, puerto_smtp)
        servidor.starttls()
        servidor.login(remitente, contraseña)
        servidor.send_message(mensaje)
        servidor.quit()
        print("Correo electrónico enviado exitosamente a", destinatario)
    except Exception as e:
        print("Ocurrió un error al enviar el correo electrónico:", str(e))




if (__name__ == '__main__'):
    print("******************************")
    comprobar_ficheros_descargados()