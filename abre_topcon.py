# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# 35240433039223000979550010003667421381161110
import time
import pytesseract
from ahk import AHK
from funcoes import marca_lancado, procura_imagem, extrai_txt_img
from acoes_planilha import valida_lancamento
from valida_pedido import valida_pedido
import pyautogui as bot
import os
#from datetime import date
#import sqlite3

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()

chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.PAUSE = 1.3

#Executa o RemoteApp
os.startfile("RemoteApp-Cortesia.rdp")

#Clica no campo da senha para inserção
bot.click(procura_imagem('img_topcon/campo_senha.png', continuar_exec=True))
bot.write('C0rtesi@01')
bot.press('ENTER')