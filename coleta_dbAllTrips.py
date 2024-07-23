# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link do site do Senior
# https://logincloud.senior.com.br/logon/LogonPoint/tmindex.html

import time
import pytesseract
from ahk import AHK
import pyautogui as bot
from datetime import date
from selenium import webdriver
from funcoes import procura_imagem
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
#from baixa_controle import download_registros, inicia_navegador

# --- Definição de parametros
ahk = AHK()
posicao_img = 0 
bot.PAUSE = 0.5
continuar = True
bot.FAILSAFE = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
planilha_original_leitura = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg?e=RYfgcA"
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"

#! Funções
def inicia_navegador():
    #Definições Chrome Driver
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(options=options,service=servico)
    navegador.implicitly_wait(30)
    return navegador


def abre_planilha_chrome():
    navegador = inicia_navegador()
    navegador.get(planilha_debug)
    navegador.maximize_window()
    time.sleep(5)


if __name__ == "__main__":
    abre_planilha_chrome()
    
    ahk.win_exists('debug_db_alltrips', title_match_mode= 2)
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    
    # Clica no meio da planilha e navega até a coluna A1
    bot.click(800, 800)
    bot.hotkey('CTRL', 'HOME')
    bot.click(800, 800)
    