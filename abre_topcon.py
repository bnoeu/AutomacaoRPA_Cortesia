# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pytesseract
from ahk import AHK
from funcoes import procura_imagem, corrige_topcompras
import pyautogui as bot
import os
from colorama import Fore, Style, Back

# Definição de parametros
ahk = AHK()
bot.PAUSE = 0.6
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
#! Dados para acesso ao remoteApp
login_rdp = 'bruno.s'
senha_rdp = '1Pyth0n@'


def abre_mercantil():
    
    print(Fore.RED + '--- Executando a função: ABRE MERCANTIL ' + Style.RESET_ALL)
    # Inicia fechando o modulo de compras.
    verifica_topcompras = 0
    time.sleep(0.2)
    corrige_topcompras()
    
    while ahk.win_exists('TopCompras', title_match_mode=2):
        ahk.win_close('TopCompras', title_match_mode=2)
        time.sleep(0.5)
        if verifica_topcompras > 10:
            print('--- Tentou fechar o TopCompras porém não conseguiu! Fechando o remote por completo')
            abre_topcon()
            
        verifica_topcompras += 1
            
    print('--- Fechou o modulo de compras, reabrindo uma nova execução!')
    # Ativa o Topcon, e clica no topcompras, e executa a função para correção do nome.
    ahk.win_activate('TopCon', title_match_mode= 2)
    bot.click(procura_imagem(imagem='img_topcon/logo_topcompras.png'))
    time.sleep(1)
    corrige_topcompras()

    #Abre o TopCompras, e verifica se aparece a tela "interveniente"
    ahk.win_activate('TopCompras', title_match_mode= 2)
    time.sleep(1)
    if procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True):
        print('--- Encontrou a tela do interveniente, clicando no botão "OK"')
        bot.press('ENTER')
    else:
        print('--- Não exibiu a tela de interveniente.')
    
    navega_topcompras()
        
def navega_topcompras():
    # Navegando entre os menus para abrir a opção "Compras - Mercantil"
    ahk.win_activate('TopCompras', title_match_mode= 2)
    bot.click(900, 900)
    bot.press('ALT')
    bot.press('RIGHT', presses= 2, interval= 0.05)
    bot.press('DOWN', presses= 7, interval= 0.05)
    bot.press('ENTER')
    time.sleep(3)
    print(Fore.GREEN +  '--- TopCompras aberto!' + Style.RESET_ALL)

#* ---------------- PROGRAMA PRINCIPAL ------------
def fecha_execucoes():
    # Encerra todos os processos do AHK
    os.system('taskkill /im AutoHotkey.exe /f /t')    
    
    print(Back.RED + '--- iniciando fecha_execucoes --- Realizando fechamento do TopCompras' + Style.RESET_ALL)
    
    # Primeiro força o fechamento das telas, para evitar erros de validações
    while ahk.win_exists('TopCompras', title_match_mode= 2):
        ahk.win_close('TopCompras', title_match_mode= 2)   
    else:
        print(Back.RED + '--- Fechou o TopCompras, para garantir que seja uma execução limpa' + Style.RESET_ALL)
    
    # Força o fechamento do processo do RDP por completo
    os.system('taskkill /im mstsc.exe /f /t')
    print(Back.RED + '--- Fechou o processo do RemoteDesktop' + Style.RESET_ALL)


def abre_topcon():
    bot.PAUSE = 1
    # Começa garantindo que fechou todas as execuções antigas.
    fecha_execucoes()
    
    print('--- Iniciando o RemoteApp')
    os.startfile('RemoteApp-Cortesia.rdp')
    #os.startfile('RemoteApp-CortesiaVPN.rdp')
    time.sleep(0.5)
    
    # Tenta encontrar em ingles e portugues.
    telas_seguranca = ['Windows Security', 'Segurança do Windows']
    tela_login_rdp = False
    
    for tela in telas_seguranca:
        try:
            time.sleep(0.5)
            ahk.win_wait_active(tela, title_match_mode= 2, timeout = 5)
        except TimeoutError:
            time.sleep(0.5)
        else:
            #print(F'--- Encontrou com o nome {tela}')
            tela_login_rdp = tela
    
    if tela_login_rdp is False: # Caso não encontro nenhuma das telas da lista "telas_seguranca"
        exit(bot.alert('--- Nehuma das telas de segurança abriu.'))
    
    # Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
    if ahk.win_exists(tela_login_rdp, title_match_mode= 2):
        print(F'--- Abriu a tela "{tela_login_rdp}", realizando o login" ')
        ahk.win_activate(tela_login_rdp, title_match_mode= 2)
        ahk.win_wait_active(tela_login_rdp, title_match_mode= 2, timeout = 10)
        
        # Insere os dados de login.
        bot.write(senha_rdp) #Senha BRUNO.S 
        bot.press('TAB', presses= 3, interval= 0.05) # Navega até o botão "Ok"
        bot.press('ENTER')
        print('--- Login realizado no RemoteApp-Cortesia.rdp')
        
    elif ahk.win_exists('TopCon (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2): 
        print('--- Tela de login do Topcon já está aberta, prosseguindo para o login')

    # Se o modulo de compras estiver fechado, realiza o login no TopCon
    if ahk.win_exists('TopCompras', title_match_mode= 2) is False:
        # Realiza login no TopCon
        while True:
            ahk.win_activate('TopCon', title_match_mode= 2)
            ahk.win_wait_active('TopCon', title_match_mode= 2, timeout= 30)
            if procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png', continuar_exec= True): # Aguarda até aparecer o campo do servidor preenchido
                print('--- Tela de login do topcon aberta')
                bot.click(procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png'))
                bot.press('tab', presses= 2, interval= 0.005)
                bot.press('backspace')
                
                # Insere os dados de login do usuario BRUNO.S
                bot.write('BRUNO.S')
                bot.press('tab')
                bot.write('rockie')
                bot.press('tab')
                bot.press('enter')
                time.sleep(3)
                break
            
            if procura_imagem(imagem='img_topcon/logo_principal.png', continuar_exec= True):
                print('--- Tela do Topcon já está aberta.')
                break
    
    # Com o programa TopCon aberto, realiza a abertura do modulo de compras (mercantil). 
    abre_mercantil()

if __name__ == '__main__':
    #abre_mercantil()
    #fecha_execucoes()
    abre_topcon()