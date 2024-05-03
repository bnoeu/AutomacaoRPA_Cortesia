# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import datetime
import pytesseract
#import cv2
from ahk import AHK
import pyautogui as bot
#from selenium import webdriver
from funcoes import procura_imagem, extrai_txt_img, marca_lancado
from acoes_planilha import valida_lancamento
from valida_pedido import valida_pedido
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


#ahk.win_activate('TopCompras', title_match_mode= 2)
time.sleep(1)
#ahk.win_activate('db_alltrips', title_match_mode= 2)
#! Utilizado apenas para estar trechos de codigo.

bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec=True))
#Verifica se apareceu a tela de transferencia 
if procura_imagem('img_topcon/txt_transfMateriaPrima.png', continuar_exec=True) is not False:
    if procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True) is not False:
        bot.click(procura_imagem('img_topcon/bt_sim.png',
                continuar_exec=True, limite_tentativa=4))
        while True:  # Aguardar o .PDF
            try:
                ahk.win_wait('.pdf', title_match_mode=2, timeout=2)
                time.sleep(0.4)
            except TimeoutError:
                print('Aguardando .PDF')
            else:
                ahk.win_activate('.pdf', title_match_mode=2)
                ahk.win_close('pdf - Google Chrome', title_match_mode=2)
                print('Fechou o PDF')
                break
        time.sleep(0.8)
        ahk.win_activate('Transmissão', title_match_mode=2)
        bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))
        time.sleep(1)
