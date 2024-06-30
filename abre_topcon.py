# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# 35240433039223000979550010003667421381161110
import time
import pytesseract
from ahk import AHK
from funcoes import procura_imagem
import pyautogui as bot
import os
#from datetime import date
#import sqlite3

# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1.5
posicao_img = 0
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
#! Dados para acesso ao remoteApp
login_rdp = 'bruno.s'
senha_rdp = 'C0ncret0'

#* ---------------- PROGRAMA PRINCIPAL ------------
def fecha_execucoes():
    print(F'Pausa do bot {bot.PAUSE}')
    #Primeiro força o fechamento das telas, para evitar erros de validações, e depois abre o RDP
    print('--- Fechando as execuções atuais.')
    os.system('taskkill /im mstsc.exe /f /t') #Força o fechamento do processo do RDP por completo
    
    ''' #! Subistituido por apenas essa linha acima.
    ahk.win_kill('Segurança do Windows', title_match_mode= 2)
    ahk.win_kill('mstsc', title_match_mode= 2)
    ahk.win_kill('RemoteApp', title_match_mode= 2)
    ahk.win_kill('RemoteApp', title_match_mode= 2)
    ahk.win_kill('TopCon', title_match_mode= 2)
    ahk.win_close('VM-CortesiaApli.CORTESIA.com', title_match_mode= 2)
    '''
    print('--- Iniciando o Topcon')

def abre_topcon():
    os.startfile('RemoteApp-Cortesia.rdp')
    time.sleep(2)
    ahk.win_wait_active('Windows Security', title_match_mode= 2, timeout = 15)
    if ahk.win_exists('Windows Security', title_match_mode= 2):
        print('--- Abriu a tela "Windows Security, realizando o login" ')
        #Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
        ahk.win_activate('Windows Security', title_match_mode= 2)
        ahk.win_wait_active('Windows Security', title_match_mode= 2, timeout = 30)
        #bot.click(procura_imagem(imagem='img_windows/txt_seguranca.png'))
        #bot.write('C0rtesi@01') #Senha BARBARA.K
        time.sleep(1)
        bot.write(senha_rdp) #Senha BRUNO.S 
        bot.press('TAB', presses= 3, interval= 0.02)
        bot.press('ENTER')
        print('--- Login realizado no RemoteApp-Cortesia.rdp')
    elif ahk.win_exists('TopCon (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2): 
        print('--- Tela de login do Topcon já está aberta, prosseguindo para o login')

    #Verifica se nessa sessão o TopCompras já está aberto.
    if ahk.win_exists('TopCompras', title_match_mode= 2) is False:
        #Realiza login no TopCon
        while procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png', limite_tentativa= 24, continuar_exec= True) is False: #Aguarda até aparecer o campo do servidor preenchido
            time.sleep(0.2)
        else:
            bot.click(procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png'))
            print('--- Tela de login do topcon aberta')
            bot.press('tab', presses= 2, interval= 0.005)
            bot.press('backspace')
            
            #Insere os dados de login do usuario BRUNO.S
            bot.write('BRUNO.S')
            bot.press('tab')
            bot.write('rockie')
            bot.press('tab')
            bot.press('enter')
            time.sleep(3)
    
    #Abre o modulo de compras e navega até a tela de lançamento
    print('--- Tentando abrir o modulo de compras')
    
    ''' #! Substituido pela logica informada a baixo.
    ahk.win_activate('TopCompras - Versão', title_match_mode= 2)
    try:
        ahk.win_wait('TopCompras', title_match_mode=2, timeout= 10)
    except TimeoutError:
        print('--- Não encontrou a janela "TopCompras", tentando pelo icone do topcon')
        icone_carrinho = procura_imagem(imagem='img_topcon/icone_topcon.png', continuar_exec=True)
        if icone_carrinho is not False: #Caso encontre o icone
            bot.click(icone_carrinho)
            ahk.win_set_title(new_title= 'TopCompras', title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)
        else:
            exit(bot.alert('---Tela de Compras não abriu.'))
    bot.press('ENTER')
    '''
    ahk.win_activate('TopCon -', title_match_mode= 2)
    bot.press('TAB', presses= 20)
    bot.press('ENTER')
    time.sleep(3)

    if ahk.win_exists(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1):
        print('Encontrou a tela sem o nome, alterando')
        ahk.win_set_title(new_title= 'TopCompras', title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)
    #Verifica se aparece a tela "interveniente"
    if procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True, confianca= 0.7):
        print('--- Encontrou a tela do interveniente, clicando no botão "OK"')
        bot.press('ENTER')
        
    #Navegando entre os menus para abrir a opção "Compras - Mercantil"
    ahk.win_activate('TopCompras', title_match_mode= 2)
    bot.press('ALT')
    bot.press('RIGHT', presses= 2, interval= 0.05)
    bot.press('DOWN', presses= 7, interval= 0.05)
    bot.press('ENTER')
    time.sleep(3)

if __name__ == '__main__':
    fecha_execucoes()
    abre_topcon()