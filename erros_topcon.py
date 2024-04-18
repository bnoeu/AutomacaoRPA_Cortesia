# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import datetime
import cv2
#import pygetwindow as gw
import numpy as np
import pytesseract
from ahk import AHK
from funcoes import procura_imagem
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"
bot.useImageNotFoundException(False)

#Abre o Topcon novamente
ahk.win_activate('TopCompras', title_match_mode=2)
ahk.win_wait_active('TopCompras', title_match_mode=2)

#Verifica se encontrou o icone de erro
if procura_imagem(imagem='img_topcon/icone_erro.png', continuar_exec=True) is not False:
    img = bot.screenshot('img_geradas/tela_erro.png')
    print('Encontrado icone de erro, tirando print da tela!')
    
    #Clica no botão ok para sair da tela do erro. 
    bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg'))
    
    #! Aqui precisa analisar a tela que será apresentada após isso. 
    pass