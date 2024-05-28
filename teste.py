# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
from datetime import date
import pytesseract
#import cv2
from ahk import AHK
import pyautogui as bot
from funcoes import procura_imagem, extrai_txt_img, marca_lancado
from acoes_planilha import valida_lancamento
from valida_pedido import valida_pedido
'''
#*Selenium
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
bot.FAILSAFE = True
acabou_pedido = ''
numero_nf = "965999"
transportador = "111594"
chave_xml, silo2, silo1 = '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
time.sleep(0.8)
start = time.time()


#! Variavel de teste
silo1 = 'SILO 1'
filial_estoq = 'JAGUARE'
centro_custo = filial_estoq
cracha_mot = '112480'


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

''' #* Consulta as telas abertas
for telas in ahk.list_windows():
    print(telas.text)
'''


#ahk.win_activate('TopCompras', title_match_mode= 2)
#ahk.win_activate('db_alltrips', title_match_mode= 2)
time.sleep(0.5)
#! Utilizado apenas para estar trechos de codigo.
#bot.click(procura_imagem(imagem='img_topcon/icone_topcon.png', continuar_exec=True))

'''
while True:
    bot.press('alt') #Ativa os atalhos
    #Clica no botão para alterar Edição / Exibição
    bot.press('z')
    bot.press('m')
    #Habilita a opção: Exibição
    bot.press('e')
'''

'''
certificado = "/certificado_nfe/certificado.pfx"
senha = '123456'
uf = 'sp'
homologacao = True

con = ComunicacaoSefaz(uf, certificado, senha, homologacao)
xml = con.status_servico('nfe')
print(xml.text)
'''

print(ahk.win_get_title())
