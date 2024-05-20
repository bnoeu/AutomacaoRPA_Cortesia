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
#*Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

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



#! Variavel de teste
silo1 = 'SILO 6'
filial_estoq = 'JAGUARE'
centro_custo = filial_estoq
cracha_mot = '112480'


'''
#Cria a conexão com o banco de dados
con = sqlite3.connect("informacoes.db")

#Cursor para realizar comandos dentro do banco de dados
cur = con.cursor()

#Utilizando o cursor, executa a ação da criação da tabela informacoes, com as seguintes colunas: XML, CRACHA, TEMPO
#cur.execute("CREATE TABLE informacoes(xml, cracha, tempo)")
#! Continuar tutorial de banco de dados https://docs.python.org/3/library/sqlite3.html
exit()
'''

ahk.win_activate('TopCompras', title_match_mode= 2)
#ahk.win_activate('db_alltrips', title_match_mode= 2)
time.sleep(0.5)
#! Utilizado apenas para estar trechos de codigo.


bot.click(1006, 345)  # Campo data da operação150524
time.sleep(1)
hoje = date.today()
hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
print(F'--- Alterando a data para {hoje}')
data_NfeFaturada = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(895, 299, 20, 20))
if data_NfeFaturada < '11':
    print('Data menor que 11')
    bot.write('11052024')
    bot.press('enter')
    time.sleep(0.5)
    if procura_imagem(imagem='img_topcon/txt_NaoPermitidoData.png', continuar_exec=True, limite_tentativa= 12):
        print('Precisa mudar a data')
        bot.press('enter')
        bot.write(hoje)
        bot.press('enter')
        time.sleep(0.5)
else:
    bot.write(hoje)
    bot.press('enter')
    time.sleep(0.5)

