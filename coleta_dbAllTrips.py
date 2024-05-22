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
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
#from baixa_controle import download_registros, inicia_navegador

# --- Definição de parametros
ahk = AHK()
posicao_img = 0 
bot.PAUSE = 1
continuar = True
bot.FAILSAFE = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


#! Link do db_alltrips_old
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/EeKEw02Y8wxNsUl20Ye6AXEBI6hSgj_U9zmkYI5O9pN6Lw?e=ouNo7U

# * ---------------------------------------------------------------------------------------------------
# *                                        Inicio do Programa
# * ---------------------------------------------------------------------------------------------------
#Execução apenas caso rode o arquivo principal



if __name__ == "__main__":
    '''
    hoje = date.today()
    day = hoje.strftime("%d")
    month = hoje.strftime("%m")
    year =  hoje.strftime("%y")
    #print(day, month, year)
    '''
    def inicia_navegador():
        #Definições Chrome Driver
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(options=options,service=servico)
        navegador.implicitly_wait(15)
        return navegador

    
    navegador = inicia_navegador()
    navegador.get("https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/EeKEw02Y8wxNsUl20Ye6AXEBI6hSgj_U9zmkYI5O9pN6Lw?e=ouNo7U")
    #navegador.get("https://g1.globo.com/")
    navegador.maximize_window()
    navegador.find_element('xpath', '//*[@id="ModeSwitcher"]/span[1]').click()
    exit()
    #Preencher a senha

        
    #Clica no botão de login
    navegador.find_element('xpath', '//*[@id="loginBtn"]').click()
    time.sleep(1.5)
    #Clicar no botão "Detecte o aplicativo"
    navegador.find_element('xpath', '//*[@id="protocolhandler-welcome-installButton"]').click()
    time.sleep(1)
    bot.click(procura_imagem(img='img_chrome/abrir_citrix.png', continuar_exec=True))
    time.sleep(1)
    navegador.find_element('xpath', '//*[@id="protocolhandler-detect-alreadyInstalledLink"]').click()
    navegador.find_element('xpath', '//*[@id="home-screen"]/div[2]/section[5]/div[5]/div/ul/li[5]').click()
    #Clica no botão "Downloads"
    #bot.click(procura_imagem(img='img_chrome/icone_downloads.png'))
    #Clica no arquivo .ICA
    while procura_imagem(img='img_chrome/icone_ica.png') is False:
        time.sleep(0.1)
    else:
        bot.click(procura_imagem(img='img_chrome/icone_ica.png'))

#* -------------------- Processos Sênior Desktop. ---------------------
    ahk.win_wait('Controle de Ponto e Refeitório', title_match_mode= 2, timeout= 1000)
    ahk.win_activate('Controle de Ponto e Refeitório', title_match_mode= 2)
    time.sleep(0.5)
    
    #Aguarda até encontrar o campo para inserir os dados de login.
    while procura_imagem(img='img_senior/txt_usuario.png', continuar_exec= True) is False:
        time.sleep(0.3)
    else:
        print('--- Realizando o processo de login no Sênior Desktop App')
        bot.click(procura_imagem(img='img_senior/txt_usuario.png'))

    #! Dados de login do Sênior Desktop
    bot.write('eliane')
    bot.click(procura_imagem(img='img_senior/txt_senha.png'))
    bot.write('sol2020')
    
    bot.click(procura_imagem(img='img_senior/bt_autenticar.png', continuar_exec=True))

    #Aguarda o modulo carregar por completo
    ahk.win_wait('Controle de Ponto e Refeitório', title_match_mode= 2, timeout= 1000)
    ahk.win_activate('Controle de Ponto e Refeitório', title_match_mode= 2)
    
    
    while procura_imagem(img='img_senior/logo_senior.png', continuar_exec=True) is False:
        time.sleep(0.5)
    else:
        bot.click(procura_imagem(img='img_senior/logo_senior.png', continuar_exec=True))
        print('--- Carregou o Sênior')
        time.sleep(0.5)
        
    #Abre o menu de pesquisa, e abre a tela "Calculos / Apuração > Calcular"
    bot.click(procura_imagem(img='img_senior/icon_pesquisa.png', continuar_exec=True))
    bot.write('FRCALAPU')
    bot.click(procura_imagem(img='img_senior/bt_calcular.png', continuar_exec=True))
    time.sleep(0.5)
    
    #Caso apareça a tela de "Um aplicativo "
    if ahk.win_exists('Citrix Workspace - Aviso de segurança', title_match_mode= 2):
        bot.click(procura_imagem(img='img_senior/bt_acessoTotal.png', continuar_exec=True))
    
    
    while procura_imagem(img='img_senior/txt_periodo.png', continuar_exec=True) is False:
        time.sleep(0.5)
    else:
        print('--- Carregou a tela de calculo da apuração')
        ahk.win_activate('FRCALAPU', title_match_mode= 2)
        time.sleep(1)
        print('--- Carregou a tela de calculo da apuração')
        time.sleep(0.5)
        bot.write('02052024')
        bot.click(893, 191)
        time.sleep(0.5)
        bot.write('03052024')
        bot.click(procura_imagem(img='img_senior/bt_marcaTodos.png', continuar_exec=True))   
     
    #Clica no botão processar, e aguarda a tela para confirmar o calculo de apuração 
    bot.click(procura_imagem(img='img_senior/txt_processar.png', continuar_exec=True))
    
    ahk.win_wait('Confirmação', title_match_mode= 2, timeout= 1000)
    ahk.win_activate('Confirmação', title_match_mode= 2)
    bot.press('alt') 
    bot.press('s') #Seleciona a opção SIM
    
    ahk.win_activate('Senior', title_match_mode= 2)
    ahk.win_wait('FRPROAPU', title_match_mode= 2, timeout= 1000)
    ahk.win_activate('FRPROAPU', title_match_mode= 2)
    while (procura_imagem(img='img_senior/txt_StatusErroLog.png', limite_tentativa= 1,  continuar_exec = True)) is False:
        print('Aguardando o processamento.')
    ahk.send('LEFT')
    ahk.send('ENTER')
    

    #TODO --- Caso apareça o erro "Número do PIS/CPF não cadastrado"
