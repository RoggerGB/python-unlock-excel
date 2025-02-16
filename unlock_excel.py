import zipfile
import shutil
import os
from io import BytesIO
import time
import class_correo
import datetime


# RUTA DE LA CARPETA DONDE ESTA CARGADO EL ARCHIVO DEL CORREO
path = r"C:\Users\70262920\Desktop\Doc_unlock"
ruta_origen = path


# FUNCIONES
def listar_archivos(ruta):
    archivos = os.listdir(ruta)
    for archivo in archivos:
        return archivo


def validar_xlsx(archivo):
    if ".zip" in archivo:
        return False
    elif ".xlsx" in archivo:
        return True


def temporizador():
    intervalo = 60
    repeticiones = 5
    for _ in range(repeticiones):
        tiempo_inicio = time.time()
        while time.time() - tiempo_inicio < intervalo:
            if listar_archivos(ruta_origen) is not None:
                print("¡Archivo encontrado! Terminando el temporizador.")
                return True
                break
            time.sleep(1)
        print("¡Tiempo agotado!\n")


try:
    if temporizador():
        archivo_name = listar_archivos(ruta_origen)
        ruta_origen = ruta_origen + "\\" + archivo_name
        print(ruta_origen)
        if not validar_xlsx(ruta_origen):
            ruta_destino = ruta_origen.replace("\\" + archivo_name, "\\extraido")
            with zipfile.ZipFile(ruta_origen, "r") as zip_ref:
                zip_ref.extractall(ruta_destino)
            archivo_name = listar_archivos(ruta_destino)
            ruta_origen = ruta_destino + "\\" + archivo_name
        # REASIGNAR EXTENSION .ZIP
        ruta_del_archivo_zip = ruta_origen[:-5] + ".zip"
        # MODIFICAR A .ZIP
        shutil.move(ruta_origen, ruta_del_archivo_zip)
        archivo_a_modificar = "xl/workbook.xml"
        # ABRIR EL WORKBOOK.XML
        with zipfile.ZipFile(ruta_del_archivo_zip, "r") as zip_original:
            contenido_archivo = zip_original.read(archivo_a_modificar)
        # CAMBIAR LAS PROPIEDADES DEL WORKBOOK.XML
        contenido_str = contenido_archivo.decode("utf-8")
        nuevo_contenido_str = contenido_str.replace(
            'lockStructure="1"', 'lockStructure="0"'
        )
        for x in ['state="hidden"', 'state="veryHidden"']:
            nuevo_contenido_str = nuevo_contenido_str.replace(x, "")
        nuevo_contenido = nuevo_contenido_str.encode("utf-8")
        try:
            with BytesIO() as nuevo_zip_buffer:
                # COPIAR EL NUEVO WORKBOOK AL ZIP
                with zipfile.ZipFile(
                    nuevo_zip_buffer, "a", zipfile.ZIP_DEFLATED, False
                ) as nuevo_zip:
                    with zipfile.ZipFile(ruta_del_archivo_zip, "r") as zip_original:
                        for nombre_archivo in zip_original.namelist():
                            if nombre_archivo == archivo_a_modificar:
                                nuevo_zip.writestr(nombre_archivo, nuevo_contenido)
                            else:
                                nuevo_zip.writestr(
                                    nombre_archivo, zip_original.read(nombre_archivo)
                                )
                with open(ruta_origen, "wb") as nuevo_archivo_zip:
                    nuevo_archivo_zip.write(nuevo_zip_buffer.getvalue())
        except Exception as e:
            print("ERROR EN REEMPLAZO DE WORKBOOK")
        # ENVIAR CORREO
        try:
            # CAMBIAR EL DESTINATARIO
            correo = class_correo.Class_Correo(["rbarrientos@mfsac.pe","rhamm@mfsac.pe"])
            cuerpo_html = "<html><body><h1>Hola, este es el reenvio del excel desbloqueado</h1></body></html>"
            correo.Cuerpo_HTML(cuerpo_html)
            correo.Adjuntar_Archivo(ruta_origen, archivo_name, "text/plain")
            correo.Enviar_Correo_Estado("ARCHIVO: " + archivo_name)

            try:
                # Elimina todos los archivos y carpetas dentro de la carpeta
                for root, dirs, files in os.walk(path, topdown=False):
                    for archivo in files:
                        ruta_completa = os.path.join(root, archivo)
                        try:
                            os.remove(ruta_completa)
                            print(f"Archivo eliminado: {ruta_completa}")
                        except Exception as e:
                            print(
                                f"No se pudo eliminar el archivo {ruta_completa}. Error: {str(e)}"
                            )

                    for carpeta in dirs:
                        ruta_carpeta = os.path.join(root, carpeta)
                        try:
                            os.rmdir(ruta_carpeta)
                            print(f"Carpeta eliminada: {ruta_carpeta}")
                        except Exception as e:
                            print(
                                f"No se pudo eliminar la carpeta {ruta_carpeta}. Error: {str(e)}"
                            )

                print(f"Contenido de la carpeta eliminado: {path}")
            except Exception as e:
                print(
                    f"No se pudo eliminar el contenido de la carpeta {path}. Error: {str(e)}"
                )
        except Exception as e:
            print("ERROR EN ENVIAR EL CORREO")

    else:
        try:
            # Instancia un objeto de la clase Class_Correo con la lista de destinatarios
            correo = class_correo.Class_Correo(["rbarrientos@mfsac.pe"])
            fecha_actual = datetime.datetime.now()
            fecha_actual = fecha_actual.strftime("%d/%m/%Y %H:%M:%S")
            cuerpo_html = f"<html><body><h1>NO SE RECIBIO EL ARCHIVO EL DÍA {fecha_actual}</h1></body></html>"
            correo.Cuerpo_HTML(cuerpo_html)
            # Adjunta un archivo al correo
            # Envia el correo con un asunto especifico
            correo.Enviar_Correo_Estado("EXCEL NO RECIBIDO")
        except Exception as e:
            print("ERROR EN EL ENVIO DEL CORREO")
except Exception as e:
    print(e)
