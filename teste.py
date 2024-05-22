# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
from datetime import date
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
bot.FAILSAFE = True
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

for telas in ahk.list_windows():
    print(telas.text)


#ahk.win_activate('TopCompras', title_match_mode= 2)
#ahk.win_activate('db_alltrips', title_match_mode= 2)
time.sleep(0.5)
#! Utilizado apenas para estar trechos de codigo.

'''
if procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True, limite_tentativa= 12, confianca= 0.7) is not False:
    bot.click(procura_imagem('img_topcon/bt_sim.png', continuar_exec=True))
    while True:  # Aguardar o .PDF
        try:
            ahk.win_wait('.pdf', title_match_mode=2, timeout= 15)
        except TimeoutError:
            print('--- Aguardando .PDF da transferencia')
        else:
            ahk.win_activate('.pdf', title_match_mode=2)
            ahk.win_close('pdf - Google Chrome', title_match_mode=2)
            print('--- Fechou o PDF da transferencia')
            break
    time.sleep(0.4)
    ahk.win_activate('Transmissão', title_match_mode=2)
    bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))
    ahk.win_wait_active('TopCom', timeout=10, title_match_mode=2)
    ahk.win_activate('TopCom', title_match_mode=2)
'''

'''
while True:
    bot.press('alt') #Ativa os atalhos
    #Clica no botão para alterar Edição / Exibição
    bot.press('z')
    bot.press('m')
    #Habilita a opção: Exibição
    bot.press('e')

'''