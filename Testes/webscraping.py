from optparse import Option
import requests
import time
import pandas as pd
import numpy as np
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



navegador = webdriver.Chrome()

                                            ########## recepçao ##########
try:
    navegador.get("https://10.10.4.182/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgRecepcao = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerRecepcao = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')                                        
        
    UnidImgRecepcao = UnidImgRecepcao.text
    TonerRecepcao = TonerRecepcao.text
except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgRecepcao = tabela.loc[tabela["Local"] == "Recepção","Toner"]
    TonerRecepcao = tabela.loc[tabela["Local"] == "Recepção","Unid de Imagem"]
    
    
    
    
                                            ########## Impressão Digital ##########
try:
    navegador.get("https://10.10.4.169/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgImpDigital = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerImpDigital = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgImpDigital = UnidImgImpDigital.text
    TonerImpDigital = TonerImpDigital.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerImpDigital = tabela.loc[tabela["Local"] == "Impressão Digital","Toner"]
    UnidImgImpDigital = tabela.loc[tabela["Local"] == "Impressão Digital","Unid de Imagem"]
    

    
    
                                            ########## Produção ##########
try:

    navegador.get("https://10.10.4.156/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgProducao = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerProducao = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgProducao = UnidImgProducao.text
    TonerProducao = TonerProducao.text
    
except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgProducao = tabela.loc[tabela["Local"] == "Produção","Unid de Imagem"]
    TonerProducao = tabela.loc[tabela["Local"] == "Produção","Toner"]
    
    
    
                                            ########## Portaria ##########
try:        
    
    navegador.get("https://10.10.4.160/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgPortaria = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerPortaria = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')    
    
    UnidImgPortaria = UnidImgPortaria.text
    TonerPortaria = TonerPortaria.text
    
except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerPortaria = tabela.loc[tabela["Local"] == "Portaria","Toner"]
    UnidImgPortaria = tabela.loc[tabela["Local"] == "Portaria","Unid de Imagem"]
    
    
    
                                            ########## Manutenção ##########
try:        
    
    navegador.get("https://10.10.4.158/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgManutencao = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerManutencao = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgManutencao = UnidImgManutencao.text
    TonerManutencao = TonerManutencao.text
    
except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerManutencao = tabela.loc[tabela["Local"] == "Manutenção","Toner"]
    UnidImgManutencao = tabela.loc[tabela["Local"] == "Manutenção","Unid de Imagem"]
    
    
    

                                            ########## Expedição ##########
try:
    
    navegador.get("https://10.10.4.163/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgExpedição = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerExpedição = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgExpedição = UnidImgExpedição.text
    TonerExpedição = TonerExpedição.text
    
except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerExpedição = tabela.loc[tabela["Local"] == "Expedição","Toner"]
    UnidImgExpedição = tabela.loc[tabela["Local"] == "Expedição","Unid de Imagem"]

    
    
                                            ########## Papel e Tinta ##########
try:        

    navegador.get("https://10.10.4.164/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgPapelTinta = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerPapelTinta = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgPapelTinta = UnidImgPapelTinta.text
    TonerPapelTinta = TonerPapelTinta.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerPapelTinta = tabela.loc[tabela["Local"] == "Papel e Tinta","Toner"]
    UnidImgPapelTinta = tabela.loc[tabela["Local"] == "Papel e Tinta","Unid de Imagem"]
    
    
    
    
                                            ########## Comercial ##########
try:        

    navegador.get("https://10.10.4.162/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgComercial = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerComercial = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgComercial = UnidImgComercial.text
    TonerComercial = TonerComercial.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerComercial = tabela.loc[tabela["Local"] == "Comercial","Toner"]
    UnidImgComercial = tabela.loc[tabela["Local"] == "Comercial","Unid de Imagem"]
    
    
    

                                            ########## Segurança Trabalho ##########
try:        

    navegador.get("https://10.10.4.173/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgSegurancaTrabalho = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerSegurancaTrabalho = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgSegurancaTrabalho = UnidImgSegurancaTrabalho.text
    TonerSegurancaTrabalho = TonerSegurancaTrabalho.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerSegurancaTrabalho = tabela.loc[tabela["Local"] == "Segurança Trabalho","Toner"]
    UnidImgSegurancaTrabalho = tabela.loc[tabela["Local"] == "Segurança Trabalho","Unid de Imagem"]
    
    
    

                                            ########## Sala TI ##########
try:        

    navegador.get("https://10.10.4.174/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgTI = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerTI = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    UnidImgTI = UnidImgTI.text
    TonerTI = TonerTI.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerTI = tabela.loc[tabela["Local"] == "Sala TI","Toner"]
    UnidImgTI = tabela.loc[tabela["Local"] == "Sala TI","Unid de Imagem"]
    
    
    

                                            ########## Pré Impressão ##########
try:        

    navegador.get("https://10.10.4.151/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgPre = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerPre = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    UnidImgPre = UnidImgPre.text
    TonerPre = TonerPre.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerPre = tabela.loc[tabela["Local"] == "Pré Impressão","Toner"]
    UnidImgPre = tabela.loc[tabela["Local"] == "Pré Impressão","Unid de Imagem"]
    
    
    

                                            ########## Ambulatório Dra Renata ##########
try:

    navegador.get("https://10.10.4.177/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgDraRenata = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerDraRenata = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgDraRenata = UnidImgDraRenata.text
    TonerDraRenata = TonerDraRenata.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgDraRenata = tabela.loc[tabela["Local"] == "Ambulatório Dra Renata","Unid de Imagem"]
    TonerDraRenata = tabela.loc[tabela["Local"] == "Ambulatório Dra Renata","Toner"]

    
    
                                            ########## Ambulatório ##########
try:        

    navegador.get("https://10.10.4.155/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgAmbulatorio = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerAmbulatorio = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   

    UnidImgAmbulatorio = UnidImgAmbulatorio.text
    TonerAmbulatorio = TonerAmbulatorio.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerAmbulatorio = tabela.loc[tabela["Local"] == "Ambulatório","Toner"]
    UnidImgAmbulatorio = tabela.loc[tabela["Local"] == "Ambulatório","Unid de Imagem"]
    

    
    
                                            ########## RH ##########
try:        

    navegador.get("https://10.10.4.157/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgRH = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerRH = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgRH = UnidImgRH.text
    TonerRH = TonerRH.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerRH = tabela.loc[tabela["Local"] == "RH","Toner"]
    UnidImgRH = tabela.loc[tabela["Local"] == "RH","Unid de Imagem"]

    
    
                                            ########## Juridico ##########
try:        

    navegador.get("https://10.10.4.150/sws/index.html")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)
    
    UnidImgJuridico = navegador.find_element(By.XPATH,'//*[@id="ext-gen344"]/div/table/tbody/tr/td[2]/div/div/div[2]')
    TonerJuridico = navegador.find_element(By.XPATH,'//*[@id="ext-gen300"]/div/table/tbody/tr/td[2]/div/div/div[2]')   
    
    UnidImgJuridico = UnidImgJuridico.text
    TonerJuridico = TonerJuridico.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    TonerJuridico = tabela.loc[tabela["Local"] == "Juridico","Toner"]
    UnidImgJuridico = tabela.loc[tabela["Local"] == "Juridico","Unid de Imagem"]

##################################### E52645 #######################################

  ########## Almoxarifado ##########
try:

    navegador.get("https://10.10.4.152/")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgAlmoxarifado = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[2]/span[2]')
    TonerAlmoxarifado = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[1]/span[2]')                                        
        
    UnidImgAlmoxarifado = UnidImgAlmoxarifado.text
    TonerAlmoxarifado = TonerAlmoxarifado.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgAlmoxarifado = tabela.loc[tabela["Local"] == "Almoxarifado","Toner"]
    TonerAlmoxarifado = tabela.loc[tabela["Local"] == "Almoxarifado","Unid de Imagem"]

  ########## Qualidade ##########
try:

    navegador.get("https://10.10.4.153/")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgQualidade = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[2]/span[2]')
    TonerQualidade = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[1]/span[2]')                                        
        
    UnidImgQualidade = UnidImgQualidade.text
    TonerQualidade = TonerQualidade.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgQualidade = tabela.loc[tabela["Local"] == "Qualidade","Toner"]
    TonerQualidade = tabela.loc[tabela["Local"] == "Qualidade","Unid de Imagem"]

  ########## Atendimento PB ##########
try:

    navegador.get("https://10.10.4.154/")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgAtendimentoPB = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[2]/span[2]')
    TonerAtendimentoPB = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[1]/span[2]')                                        
        
    UnidImgAtendimentoPB = UnidImgAtendimentoPB.text
    TonerAtendimentoPB = TonerAtendimentoPB.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgAtendimentoPB = tabela.loc[tabela["Local"] == "Atendimento PB","Toner"]
    TonerAtendimentoPB = tabela.loc[tabela["Local"] == "Atendimento PB","Unid de Imagem"]

  ########## Financeiro ##########
try:

    navegador.get("https://10.10.4.171/")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgFinanceiro = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[2]/span[2]')
    TonerFinanceiro = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[1]/span[2]')                                        
        
    UnidImgFinanceiro = UnidImgFinanceiro.text
    TonerFinanceiro = TonerFinanceiro.text

except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgFinanceiro = tabela.loc[tabela["Local"] == "Financeiro","Toner"]
    TonerFinanceiro = tabela.loc[tabela["Local"] == "Financeiro","Unid de Imagem"]

  ########## Compras ##########
try:

    navegador.get("https://10.10.4.159/")
    navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
    navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
    time.sleep(35)

    UnidImgCompras = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[2]/span[2]')
    TonerCompras = navegador.find_element(By.XPATH,'/html/body/div[2]/div/div/div[2]/div/div[2]/form/div/div[2]/div[2]/div/div[1]/span[2]')                                        
        
    UnidImgCompras = UnidImgCompras.text
    TonerCompras = TonerCompras.text
    
except:
    tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco")
    UnidImgCompras = tabela.loc[tabela["Local"] == "Compras","Toner"]
    TonerCompras = tabela.loc[tabela["Local"] == "Compras","Unid de Imagem"]



########################## Atualizar a porcentagem do toner ##########################
tabela = pd.read_excel('Relatorio_Impressoras_Plural.xlsx')

########## Almoxarifado ##########
tabela.loc[tabela["Local"] == "Almoxarifado","Toner"] = TonerAlmoxarifado

########## Qualidade ##########
tabela.loc[tabela["Local"] == "Qualidade","Toner"] = TonerQualidade

########## Atendimento PB ##########
tabela.loc[tabela["Local"] == "Atendimento PB","Toner"] = TonerAtendimentoPB

########## Financeiro ##########
tabela.loc[tabela["Local"] == "Financeiro","Toner"] = TonerFinanceiro

########## Compras ##########
tabela.loc[tabela["Local"] == "Compras","Toner"] = TonerCompras

########## recepçao ##########
tabela.loc[tabela["Local"] == "Recepção","Toner"] = TonerRecepcao

########## Impressão Digital ##########
tabela.loc[tabela["Local"] == "Impressão Digital","Toner"] = TonerImpDigital


########## Produção ##########
tabela.loc[tabela["Local"] == "Produção","Toner"] = TonerProducao


########## Portaria ##########
tabela.loc[tabela["Local"] == "Portaria","Toner"] = TonerPortaria


########## Manutenção ##########
tabela.loc[tabela["Local"] == "Manutenção","Toner"] = TonerManutencao


########## Expedição ##########
tabela.loc[tabela["Local"] == "Expedição","Toner"] = TonerExpedição


########## Papel e Tinta ##########
tabela.loc[tabela["Local"] == "Papel e Tinta","Toner"] = TonerPapelTinta


########## Comercial ##########
tabela.loc[tabela["Local"] == "Comercial","Toner"] = TonerComercial


########## Segurança Trabalho ##########
tabela.loc[tabela["Local"] == "Segurança Trabalho","Toner"] = TonerSegurancaTrabalho


########## Sala TI ##########
tabela.loc[tabela["Local"] == "Sala TI","Toner"] = TonerTI


########## Pré Impressão ##########
tabela.loc[tabela["Local"] == "Pré Impressão","Toner"] = TonerPre


########## Ambulatório Dra Renata ##########
tabela.loc[tabela["Local"] == "Ambulatório Dra Renata","Toner"] = TonerDraRenata


########## Ambulatório ##########
tabela.loc[tabela["Local"] == "Ambulatório","Toner"] = TonerAmbulatorio


########## RH ##########
tabela.loc[tabela["Local"] == "RH","Toner"] = TonerRH


########## Juridico ##########
tabela.loc[tabela["Local"] == "Juridico","Toner"] = TonerJuridico





########################## Atualizar a Unidade de Imagem ##########################

########## Almoxarifado ##########
tabela.loc[tabela["Local"] == "Almoxarifado","Unid de Imagem"] = UnidImgAlmoxarifado

########## Qualidade ##########
tabela.loc[tabela["Local"] == "Qualidade","Unid de Imagem"] = UnidImgQualidade

########## Atendimento PB ##########
tabela.loc[tabela["Local"] == "Atendimento PB","Unid de Imagem"] = UnidImgAtendimentoPB

########## Financeiro ##########
tabela.loc[tabela["Local"] == "Financeiro","Unid de Imagem"] = UnidImgFinanceiro

########## Compras ##########
tabela.loc[tabela["Local"] == "Compras","Unid de Imagem"] = UnidImgCompras

########## recepçao ##########
tabela.loc[tabela["Local"] == "Recepção","Unid de Imagem"] = UnidImgRecepcao


########## Impressão Digital ##########
tabela.loc[tabela["Local"] == "Impressão Digital","Unid de Imagem"] = UnidImgImpDigital


########## Produção ##########
tabela.loc[tabela["Local"] == "Produção","Unid de Imagem"] = UnidImgProducao


########## Portaria ##########
tabela.loc[tabela["Local"] == "Portaria","Unid de Imagem"] = UnidImgPortaria


########## Manutenção ##########
tabela.loc[tabela["Local"] == "Manutenção","Unid de Imagem"] = UnidImgManutencao


########## Expedição ##########
tabela.loc[tabela["Local"] == "Expedição","Unid de Imagem"] = UnidImgExpedição


########## Papel e Tinta ##########
tabela.loc[tabela["Local"] == "Papel e Tinta","Unid de Imagem"] = UnidImgPapelTinta


########## Comercial ##########
tabela.loc[tabela["Local"] == "Comercial","Unid de Imagem"] = UnidImgComercial


########## Segurança Trabalho ##########
tabela.loc[tabela["Local"] == "Segurança Trabalho","Unid de Imagem"] = UnidImgSegurancaTrabalho


########## Sala TI ##########
tabela.loc[tabela["Local"] == "Sala TI","Unid de Imagem"] = UnidImgTI


########## Pré Impressão ##########
tabela.loc[tabela["Local"] == "Pré Impressão","Unid de Imagem"] = UnidImgPre


########## Ambulatório Dra Renata ##########
tabela.loc[tabela["Local"] == "Ambulatório Dra Renata","Unid de Imagem"] = UnidImgDraRenata


########## Ambulatório ##########
tabela.loc[tabela["Local"] == "Ambulatório","Unid de Imagem"] = UnidImgAmbulatorio


########## RH ##########
tabela.loc[tabela["Local"] == "RH","Unid de Imagem"] = UnidImgRH


########## Juridico ##########
tabela.loc[tabela["Local"] == "Juridico","Unid de Imagem"] = UnidImgJuridico

tabela.to_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name="Preto e Branco",index=False)


################################ COLORIDAS ################################


##### Adriana Gasparini #####


navegador.get("https://10.10.4.178/#hId-pgConsumables")
navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
time.sleep(10)

tabela2 = pd.read_excel('Relatorio_Impressoras_Plural.xlsx',sheet_name='Coloridas')

AdrianaB = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[2]').text
AdrianaB = AdrianaB.split(' ')
try:
    AdrianaB = AdrianaB[1]
except:
    AdrianaB = AdrianaB[0]

tabela2.loc[tabela2["Local"] == "Adriana Gasparini","Black"] = AdrianaB

AdrianaC = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[3]').text
AdrianaC = AdrianaC.split(' ')
try:
    AdrianaC = AdrianaC[1]
except:
    AdrianaC = AdrianaC[0]

tabela2.loc[tabela2["Local"] == "Adriana Gasparini","Cyan"] = AdrianaC

AdrianaY = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[5]').text
AdrianaY = AdrianaY.split(' ')
try:
    AdrianaY = AdrianaY[1]
except:
    AdrianaY = AdrianaY[0]

tabela2.loc[tabela2["Local"] == "Adriana Gasparini","Yellow"] = AdrianaY

AdrianaM = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[4]').text
AdrianaM = AdrianaM.split(' ')
try:
    AdrianaM = AdrianaM[1]
except:
    AdrianaM = AdrianaM[0]

tabela2.loc[tabela2["Local"] == "Adriana Gasparini","Magenta"] = AdrianaM


##### Anatilia #####


navegador.get("https://10.10.4.167/#hId-pgConsumables")
navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
time.sleep(10)

AnaB = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[2]').text
AnaB = AnaB.split(' ')
try:
    AnaB = AnaB[1]
except:
    AnaB = AnaB[0]

tabela2.loc[tabela2["Local"] == "Anatilia","Black"] = AnaB

AnaC = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[3]').text
AnaC = AnaC.split(' ')
try:
    AnaC = AnaC[1]
except:
    AnaC = AnaC[0]

tabela2.loc[tabela2["Local"] == "Anatilia","Cyan"] = AnaC

AnaY = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[5]').text
AnaY = AnaY.split(' ')
try:
    AnaY = AnaY[1]
except:
    AnaY = AnaY[0]

tabela2.loc[tabela2["Local"] == "Anatilia","Yellow"] = AnaY

AnaM = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[4]').text
AnaM = AnaM.split(' ')
try:
    AnaM = AnaM[1]
except:
    AnaM = AnaM[0]

tabela2.loc[tabela2["Local"] == "Anatilia","Magenta"] = AnaM


##### Atendimento Color #####


navegador.get("https://10.10.4.168/#hId-pgConsumables")
navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
time.sleep(10)

AtendimentoB = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[2]').text
AtendimentoB = AtendimentoB.split(' ')
try:
    AtendimentoB = AtendimentoB[1]
except:
    AtendimentoB = AtendimentoB[0]

tabela2.loc[tabela2["Local"] == "Atendimento Color","Black"] = AtendimentoB

AtendimentoC = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[3]').text
AtendimentoC = AtendimentoC.split(' ')
try:
    AtendimentoC = AtendimentoC[1]
except:
    AtendimentoC = AtendimentoC[0]

tabela2.loc[tabela2["Local"] == "Atendimento Color","Cyan"] = AtendimentoC

AtendimentoY = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[5]').text
AtendimentoY = AtendimentoY.split(' ')
try:
    AtendimentoY = AtendimentoY[1]
except:
    AtendimentoY = AtendimentoY[0]

tabela2.loc[tabela2["Local"] == "Atendimento Color","Yellow"] = AtendimentoY

AtendimentoM = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[4]').text
AtendimentoM = AtendimentoM.split(' ')
try:
    AtendimentoM = AtendimentoM[1]
except:
    AtendimentoM = AtendimentoM[0]

tabela2.loc[tabela2["Local"] == "Atendimento Color","Magenta"] = AtendimentoM


##### Carlos Jacomine #####


navegador.get("https://10.10.4.178/#hId-pgConsumables")
navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
time.sleep(10)

CarlosB = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[2]').text
CarlosB = CarlosB.split(' ')
try:
    CarlosB = CarlosB[1]
except:
    CarlosB = CarlosB[0]

tabela2.loc[tabela2["Local"] == "Carlos Jacomine","Black"] = CarlosB

CarlosC = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[3]').text
CarlosC = CarlosC.split(' ')
try:
    CarlosC = CarlosC[1]
except:
    CarlosC = CarlosC[0]

tabela2.loc[tabela2["Local"] == "Carlos Jacomine","Cyan"] = CarlosC

CarlosY = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[5]').text
CarlosY = CarlosY.split(' ')
try:
    CarlosY = CarlosY[1]
except:
    CarlosY = CarlosY[0]

tabela2.loc[tabela2["Local"] == "Carlos Jacomine","Yellow"] = CarlosY

CarlosM = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[4]').text
CarlosM = CarlosM.split(' ')
try:
    CarlosM = CarlosM[1]
except:
    CarlosM = CarlosM[0]

tabela2.loc[tabela2["Local"] == "Carlos Jacomine","Magenta"] = CarlosM



##### Rogerio Cordeiro #####


navegador.get("https://10.10.4.175/#hId-pgConsumables")
navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
time.sleep(10)

RogerioB = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[2]').text
RogerioB = RogerioB.split(' ')
try:
    RogerioB = RogerioB[1]
except:
    RogerioB = RogerioB[0]

tabela2.loc[tabela2["Local"] == "Rogerio Cordeiro","Black"] = RogerioB

RogerioC = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[3]').text
RogerioC = RogerioC.split(' ')
try:
    RogerioC = RogerioC[1]
except:
    RogerioC = RogerioC[0]

tabela2.loc[tabela2["Local"] == "Rogerio Cordeiro","Cyan"] = RogerioC

RogerioY = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[5]').text
RogerioY = RogerioY.split(' ')
try:
    RogerioY = RogerioY[1]
except:
    RogerioY = RogerioY[0]

tabela2.loc[tabela2["Local"] == "Rogerio Cordeiro","Yellow"] = RogerioY

RogerioM = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[4]').text
RogerioM = RogerioM.split(' ')
try:
    RogerioM = RogerioM[1]
except:
    RogerioM = RogerioM[0]

tabela2.loc[tabela2["Local"] == "Rogerio Cordeiro","Magenta"] = RogerioM



##### Welinton Martins #####


navegador.get("https://10.10.4.172/#hId-pgConsumables")
navegador.find_element(By.XPATH, '/html/body/div/div[2]/button[3]').click()
navegador.find_element(By.XPATH, '/html/body/div/div[3]/p[2]/a').click()
time.sleep(10)

WelintonB = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[2]').text
WelintonB = WelintonB.split(' ')
try:
    WelintonB = WelintonB[1]
except:
    WelintonB = WelintonB[0]

tabela2.loc[tabela2["Local"] == "Welinton Martins","Black"] = WelintonB

WelintonC = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[3]').text
WelintonC = WelintonC.split(' ')
try:
    WelintonC = WelintonC[1]
except:
    WelintonC = WelintonC[0]

tabela2.loc[tabela2["Local"] == "Welinton Martins","Cyan"] = WelintonC

WelintonY = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[5]').text
WelintonY = WelintonY.split(' ')
try:
    WelintonY = WelintonY[1]
except:
    WelintonY = WelintonY[0]

tabela2.loc[tabela2["Local"] == "Welinton Martins","Yellow"] = WelintonY

WelintonM = navegador.find_element(By.XPATH,'//*[@id="appConsumable-inkCart-tbl-Tbl"]/tbody/tr[8]/td[4]').text
WelintonM = WelintonM.split(' ')
try:
    WelintonM = WelintonM[1]
except:
    WelintonM = WelintonM[0]

tabela2.loc[tabela2["Local"] == "Welinton Martins","Magenta"] = WelintonM
navegador.quit()

tabela2.to_excel('ImpressorasColoridas.xlsx',sheet_name='Coloridas',index=False)


##################################### Verificar toners do estoque #######################################

tipos_colunas = {'Modelo' : str,
                 'Estoque': str
                }

lista_colunas = ['Modelo',
                 'Cor',
                 'Estoque'
                ]

tabela_simpress = pd.read_excel(r'\\srvsao040\Departamentos\TI\Suporte\Estoque Simpress\Estoque (Simpress).xlsx', sheet_name='Python')

pd.set_option('display.precision',0)

Toner_MFP432 = tabela_simpress.loc[tabela_simpress["Modelo"] == "432FDN","Estoque"]

Toner_MFPE52645 = tabela_simpress.loc[tabela_simpress["Modelo"] == "E52645","Estoque"]

Toner_MFP7E77830 = tabela_simpress.loc[tabela_simpress["Modelo"] == "HP 7E77830","Estoque"]

Toner_MFP479 = tabela_simpress.loc[tabela_simpress["Modelo"] == "M479","Estoque"]



############################################ Envio de E-mail ############################################

relatorio1 = tabela.loc[:,"Local":"Unid de Imagem"]
relatorio2 = tabela2.loc[:,"Local":"--"]
fromaddr = "sistemas.plural@plural.com.br"
toaddr = "kelvin.rocha@plural.com.br"

msg = MIMEMultipart() 
msg['From'] = fromaddr 
msg['To'] = toaddr 
msg['Subject'] = "Relatório de Impressoras"
html = """\
<html>
    <head></head>
    <body>
        <h3>Segue relatorio de consumo das impressoras P&B MF432 e E52645:</h3>
        <p></p>
        {0}
        <p></p>
        <p>Toners MFP 432FDN em estoque : <strong>{1}</strong></p>
        <p>Toners E52645 em estoque : <strong>{2}</strong></p>
        <p></p>
        <h3>Segue relatorio de consumo das impressoras Coloridas:</h3>
        <p></p>
        {3}
        <p></p>
        <p>Automação Python desenvolvida por <strong>Kelvin Santos da Rocha</strong>.</p>
        <p>Suporte TI - (11) 4152-9518 / 9821</p>
    </body>
</html>
""".format(relatorio1.to_html(),Toner_MFP432.to_string(index=False),Toner_MFPE52645.to_string(index=False),relatorio2.to_html())
body = MIMEText(html, 'html')
msg.attach(body)
filename = "Relatorio_Impressoras_Plural.xlsx"
attachment = open(r"C:\Users\kelvin.rocha\Desktop\Area de Trabalho\Kelvin\Python\Automação Python\Breatifulsoap\Relatorio_Impressoras_Plural.xlsx","rb") 
p = MIMEBase('application', 'octet-stream') 
p.set_payload((attachment).read()) 
encoders.encode_base64(p) 
   
p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
msg.attach(p) 
s = smtplib.SMTP('email.plural.com.br') 
s.ehlo()
s.login("Sistemas.plural","asdf321!@#") 
text = msg.as_string() 
s.sendmail(fromaddr, toaddr, text) 
s.quit() 