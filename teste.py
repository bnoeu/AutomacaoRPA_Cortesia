# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import datetime
import pytesseract
#import pygetwindow as gw
#import cv2
from ahk import AHK
import pyautogui as bot
from funcoes import procura_imagem, extrai_txt_img, marca_lancado

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
acabou_pedido = ''
numero_nf = "965999"
transportador = "111594"
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"
time.sleep(1)


#! Variavel de teste
cracha_mot = '112251'

ahk.win_activate('TopCompras', title_match_mode= 2)
#ahk.win_activate('db_alltrips', title_match_mode= 2)
#! Utilizado apenas para estar trechos de codigo.

tentativa = 0
lista_erros = ['botao_sim.jpg', 'chave_invalida.png', 'naoencontrado_xml.png', 'chave_44digitos.png', 'nfe_cancelada.png']
while tentativa < 10:
    for item in lista_erros:
        img = 'img_topcon/' + item
        if procura_imagem(imagem = img, limite_tentativa= 1, continuar_exec=True) is not False:
            print('Achou!')
    tentativa += 1
    print(F'Tentativa: {tentativa}')