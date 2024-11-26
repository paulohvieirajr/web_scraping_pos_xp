import os
from os import listdir
from datetime import datetime
from os.path import isfile, join

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def dispara_email():
    diretorio_arquivos = './movimento_falimentar' 
    data_formatada = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

    files = [f for f in listdir(diretorio_arquivos) if isfile(join(diretorio_arquivos, f))]

    enviar_email(data_formatada, files)

    for file_path in files or []:
        os.remove(file_path)

def enviar_email(data_execucao:str, files:[str]):
    assunto = "Dados capturados de movimento falimentar"
    corpo_email = f"Segue em anexo dados capturados da execução na data {data_execucao} do movimento falimentar."
    email_origem = ""
    email_destino = ""
    email_senha = ""
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    path_to_file = 'example.csv'
    
    message = MIMEMultipart()
    message['Subject'] = assunto
    message['From'] = email_origem
    message['To'] = email_destino
    body_part = MIMEText(corpo_email)
    message.attach(body_part)
    

    for file_path in files or []:
        with open(file_path, 'rb') as file:
            message.attach(MIMEApplication(file.read(), Name="example.csv"))
    
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
       server.login(email_origem, email_senha)
       server.sendmail(email_origem, email_destino, message.as_string())