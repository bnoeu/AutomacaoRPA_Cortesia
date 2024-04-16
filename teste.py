# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import datetime
import pytesseract
import pygetwindow as gw
import cv2
from ahk import AHK
import pyautogui as bot
#from selenium import webdriver
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


#ahk.win_activate('TopCompras', title_match_mode= 2)

ahk.win_activate('db_alltrips', title_match_mode= 2)
#! Utilizado apenas para estar trechos de codigo.

bot.click(procura_imagem(imagem='img_planilha/titulo_re.png'))
bot.press('DOWN')