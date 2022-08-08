from optparse import Option
from turtle import color, left, right, width
from hamcrest import none
from openpyxl import load_workbook
import requests
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.message import EmailMessage
from email.mime.base import MIMEBase 
from email import encoders 
import openpyxl
from sqlalchemy import func
from xarray import align
import time
from datetime import datetime
import datetime as dt

   
def Enviar_Email(email):
    print('Enviando E-mail...')
    time.sleep(3)
    Agora = dt.datetime.now()
    GeradoEm = Agora.strftime("%d/%m/%Y %H:%M")
    fromaddr = "sistemas.plural@plural.com.br"
    msg = MIMEMultipart() 
    msg['From'] = fromaddr 
    msg['To'] = email
    msg['Subject'] = "Teste e-mail"
    html = """\
    <html>
        <head></head>
        <body>
            <p>Relatório gerado em: {0}</p>
            <p style="margin-top:20px"></p>
            <p>Abra a planilha dos Relatórios Diarios <a href="\\\srvsao028\Automação Python\Relatorio_Diario.xlsx"
                                                target="_blank">clicando aqui</a>
            </p>
            <p>Abra a planilha de Controle de Estoque Simpress <a href="\\\srvsao040\Departamentos\TI\Suporte\Estoque Simpress\Estoque (Simpress).xlsx""
                                                target="_blank">clicando aqui</a>
            </p>
            <p>Suporte TI - (11) 4152-9518 / 9821</p>
        </body>
    </html>
    """.format(GeradoEm)
    body = MIMEText(html, 'html')
    msg.attach(body)
    s = smtplib.SMTP('email.plural.com.br') 
    s.ehlo()
    s.login("Sistemas.plural","asdf321!@#") 
    text = msg.as_string() 
    s.sendmail(fromaddr, email, text) 
    s.quit() 
    print('E-mail Enviado com Sucesso!')
    time.sleep(3)

listaemail = ["kelvin.rocha@plural.com.br"]
for email in listaemail:
    print("Enviando email para:",email)
    Enviar_Email(email)