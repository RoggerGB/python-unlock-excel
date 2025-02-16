import traceback
from datetime import datetime, timedelta
import shutup

shutup.please()
import os, glob
import shutil
import traceback
import time
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email import encoders


class Class_Correo:
    def __init__(self, reci):
        self.__Correo = MIMEMultipart("alternative")
        self.__Correo["Date"] = time.strftime(
            "%a, %d %b %Y %H:%M:%S %z", time.localtime()
        )
        self.__Correo["From"] = "servidorbender@mfsac.pe"
        self.recipients = reci
        self.__Correo["To"] = ", ".join(reci)

    def __Enviar_Correo(self, email_subject):
        email_username = "biespecialista@mfsac.pe"
        email_password = "Tag09219"
        server = smtplib.SMTP("smtp-mail.outlook.com", "587")
        server.ehlo()
        server.starttls()
        server.login(email_username, email_password)
        self.__Correo["Subject"] = email_subject
        server.sendmail(
            self.__Correo["From"],
            self.__Correo["To"].split(","),
            self.__Correo.as_string(),
        )
        server.quit()

    def Enviar_Correo_Estado(self, Asunto):
        self.__Enviar_Correo(Asunto)

    def Adjuntar_Imagen(self, ruta_img, nombre_imagen):
        with open(ruta_img, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header(
                "Content-Disposition", "attachment", filename=ruta_img.split("\\")[-1]
            )
            img.add_header("Content-ID", f"<{nombre_imagen}>")
            self.__Correo.attach(img)

    def Cuerpo_HTML(self, HTML):
        partHTML = MIMEText(HTML, "html")
        self.__Correo.attach(partHTML)

    def Adjuntar_Archivo(self, ruta_archivo, nombre_archivo, tipo):
        part = MIMEBase("application", "octet-stream")
        part.set_payload(open(ruta_archivo, "rb").read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={nombre_archivo}")
        self.__Correo.attach(part)
