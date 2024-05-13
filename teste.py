# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import datetime
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
bot.FAILSAFE = False
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
time.sleep(0.2)
#! Utilizado apenas para estar trechos de codigo.

while procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec=True) is False:
    while ahk.win_exists('Não está respondendo', title_match_mode= 2):
        time.sleep(0.4)
        ahk.win_activate('TopCompras', title_match_mode= 2)
        
    if procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa= 1, continuar_exec=True) is not False:
        print('--- Nota já lançada, marcando planilha!')
        bot.press('ENTER')
        marca_lancado(texto_marcacao='Lancado_Manual')
        #programa_principal()

bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec=True))

#Verifica se apareceu a tela de transferencia 
if procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True, limite_tentativa= 10) is not False:
    bot.click(procura_imagem('img_topcon/bt_sim.png', continuar_exec=True))
    while True:  # Aguardar o .PDF
        try:
            ahk.win_wait('.pdf', title_match_mode=2, timeout=2)
            time.sleep(0.6)
        except TimeoutError:
            print('--- Aguardando .PDF da transferencia')
        else:
            ahk.win_activate('.pdf', title_match_mode=2)
            ahk.win_close('pdf - Google Chrome', title_match_mode=2)
            print('--- Fechou o PDF da transferencia')
            break
    time.sleep(0.8)
    ahk.win_activate('Transmissão', title_match_mode=2)
    bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))

# * -------------------------------------- Marca planilha --------------------------------------
marca_lancado(texto_marcacao='Lancado_RPA')