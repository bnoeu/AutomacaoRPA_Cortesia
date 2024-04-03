# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import cv2
#import pygetwindow as gw
import pytesseract
from ahk import AHK
from datetime import date
from funcoes import procura_imagem
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 2  # Pausa padrão do bot
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"  # Valor para teste
transportador = "111594"  # Valor para teste
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"

# * ------------------------------------------
'''
telas = ahk.list_windows()
for tela in telas:
    print(tela.title)
'''



ahk.win_activate('GeneXus', title_match_mode=2)
ahk.win_wait('GeneXus', title_match_mode= 2 ,timeout= 10)
bot.click(procura_imagem(imagem='img_gerenciador/bt_fechar.png'))
bot.click(procura_imagem(imagem='img_gerenciador/bt_distribuicao.png', limite_tentativa= 50))
bot.click(procura_imagem(imagem='img_gerenciador/autorizada.png', limite_tentativa= 500))
bot.click(954, 157) #Campo da data
hoje = date.today()
hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
bot.write('311224', interval=0.05)
bot.press('enter', presses= 3)
bot.doubleClick(procura_imagem(imagem='img_gerenciador/icone_nfe.png', limite_tentativa= 50))
ahk.win_wait('Não está respondendo', title_match_mode= 2, timeout= 50)
while ahk.win_exists('Não está respondendo', title_match_mode= 2):
    print('Aguardando')
    time.sleep(2)

#TODO --- TELA REJEIÇÃO
while True:
    procura_imagem(imagem='img_gerenciador/rejeicao.png', continuar_exec= True)
#TODO --- TELA QUE INFORMA QUANTOS BAIXOU