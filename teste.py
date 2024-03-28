# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import datetime
import pytesseract
import pygetwindow as gw
import cv2
from ahk import AHK
import pyautogui as bot
from funcoes import marca_lancado, procura_imagem, verifica_tela, extrai_txt_img
from coleta_planilha import coleta_planilha

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
acabou_pedido = ''
numero_nf = "965999"
transportador = "111594"
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"

time.sleep(1)
#ahk.win_activate('TopCompras')
ahk.win_activate('db_alltrips', title_match_mode= 2)
#! Utilizado apenas para estar trechos de codigo.
'''
if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8) is not False:
    time.sleep(2)
    print('--- Já está filtrado, continuando!')
else:
    print('--- Não está filtrado, executando o filtro!')
    bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1463, 419, 100, 100)))
    time.sleep(1)
    bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png'))
    time.sleep(1)
    bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
    time.sleep(1)
    bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
    time.sleep(1)
'''