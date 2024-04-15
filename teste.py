# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import datetime
import pytesseract
import pygetwindow as gw
import cv2
from ahk import AHK
import pyautogui as bot
<<<<<<< HEAD
from datetime import date
from selenium import webdriver
from openpyxl import load_workbook
=======
#from selenium import webdriver
>>>>>>> bd0b452094036a7972f64d01887334bb4277e777
from funcoes import procura_imagem, extrai_txt_img, marca_lancado
from acoes_planilha import valida_lancamento
from valida_pedido import valida_pedido

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
acabou_pedido = ''
numero_nf = "965999"
transportador = "111594"
chave_xml, silo2, silo1 = '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"
time.sleep(1)


#! Variavel de teste
silo1 = 'SILO 6'
filial_estoq = 'JAGUARE'
centro_custo = filial_estoq
cracha_mot = '112480'


ahk.win_activate('TopCompras', title_match_mode= 2)

#ahk.win_activate('db_alltrips', title_match_mode= 2)
#! Utilizado apenas para estar trechos de codigo.

<<<<<<< HEAD

texto = extrai_txt_img(imagem='item_nota.png',area_tela=(170, 400, 280, 30))
print(F'Texto extraido do campo Itens XML: {texto}')  

cimento = ['IE-40 RS', 'Bruno', 'Ana']

for item in cimento:
    if item in texto:
        print(F'Está aqui, o item: {item}')
    else:
        print(item + ' não está aqui')
=======
valida_lancamento()
>>>>>>> bd0b452094036a7972f64d01887334bb4277e777
