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
bot.PAUSE = 1.2  # Pausa padrão do bot
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
ahk.win_activate('TopCompras')
#! Utilizado apenas para estar trechos de codigo.

while ahk.win_exists('Não está respondendo', title_match_mode= 2):
        time.sleep(2)
while procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec=True, limite_tentativa= 8) is False:
        time.sleep(2)
        print('Aguardando')