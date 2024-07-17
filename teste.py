# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
from datetime import date
import pytesseract
import datetime
#import cv2
from ahk import AHK
import pyautogui as bot
import pandas as pd
from colorama import Fore, Back, Style
from onedrivedownloader import download as one_download
from funcoes import procura_imagem, extrai_txt_img, marca_lancado
from valida_pedido import valida_pedido
from Materia_Prima import processo_transferencia

''' #*Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from pynfe.processamento.comunicacao import ComunicacaoSefaz
'''
#import winsound
#import pygetwindow as gw

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
acabou_pedido = ''
numero_nf = "965999"
transportador = "111594"
chave_xml, silo2, silo1 = '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
time.sleep(1)
start = time.time()
hoje = datetime.date.today()

#! Variavel de teste
silo1 = 'SILO 1'
filial_estoq = 'JAGUARE'
centro_custo = filial_estoq
cracha_mot = '112480'

#! Funções



''' #* Cria banco de dados
#Cria a conexão com o banco de dados
con = sqlite3.connect("informacoes.db")

#Cursor para realizar comandos dentro do banco de dados
cur = con.cursor()

#Utilizando o cursor, executa a ação da criação da tabela informacoes, com as seguintes colunas: XML, CRACHA, TEMPO
#cur.execute("CREATE TABLE informacoes(xml, cracha, tempo)")
#! Continuar tutorial de banco de dados https://docs.python.org/3/library/sqlite3.html
exit()
'''

 #* Consulta as telas abertas

''' #* Parte da tentativa de comunicação com o SEFAZ

 





certificado = "/certificado_nfe/certificado.pfx"
senha = '123456'
uf = 'sp'
homologacao = True

con = ComunicacaoSefaz(uf, certificado, senha, homologacao)
xml = con.status_servico('nfe')
print(xml.text)
'''
