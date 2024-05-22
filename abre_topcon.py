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
bot.PAUSE = 0.5
posicao_img = 0
continuar = True
bot.FAILSAFE = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
#! Dados para acesso ao remoteApp
login_rdp = 'bruno.s'
senha_rdp = 'C0rtesi@'

#* ---------------- PROGRAMA PRINCIPAL ------------
def abre_topcon():
    #Primeiro força o fechamento das telas, para evitar erros de validações, e depois abre o RDP
    ahk.win_kill('Segurança do Windows', title_match_mode= 2)
    ahk.win_kill('RemoteApp', title_match_mode= 2)
    ahk.win_kill('RemoteApp', title_match_mode= 2)
    ahk.win_kill('TopCon', title_match_mode= 2)
    ahk.win_close('VM-CortesiaApli.CORTESIA.com', title_match_mode= 2)
    #os.close()
    os.startfile('RemoteApp-Cortesia.rdp')
    if ahk.win_exists('Segurança do Windows', title_match_mode= 2):
        print('--- Abriu a tela "Segurança do Windows, realizando o login" ')
        #Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
        ahk.win_activate('Segurança do Windows')
        ahk.win_wait_active('Segurança do Windows')
        bot.click(procura_imagem(imagem='img_windows/txt_seguranca.png'))
        bot.write('C0rtesi@01') #Senha BARBARA.K
        bot.press('TAB', presses= 3, interval= 0.02)
        bot.press('ENTER')
        print('--- Login realizado no RemoteApp-Cortesia.rdp')
    elif ahk.win_exists('TopCon (VM-CortesiaApli.CORTESIA.com)'): 
        print('--- Tela de login do Topcon já está aberta, prosseguindo para o login')

    #Realiza login no TopCon
    while procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png') is False: #Aguarda até aparecer o campo do servidor preenchido
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

    #Abre o modulo de compras e navega até a tela de lançamento
    print('--- Abrindo modulo de compras')
    bot.click(procura_imagem(imagem='img_topcon/icone_compras.png'))
    ahk.win_wait('TopCompras - Versão', title_match_mode= 2, timeout= 100)
    ahk.win_activate('TopCompras - Versão', title_match_mode= 2)
    bot.press('ENTER')
    
    #Verifica se aparece a tela "interveniente"
    procura_imagem(imagem='img_topcon/txt_interveniente.png', continuar_exec= True)
    ahk.win_activate('TopCompras - Versão', title_match_mode= 2)

if __name__ == '__main__':
    abre_topcon()