# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pytesseract
from ahk import AHK
from funcoes import procura_imagem, corrige_nometela
import pyautogui as bot
import os
from colorama import Fore, Style

# Definição de parametros
ahk = AHK()
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
#! Dados para acesso ao remoteApp
login_rdp = 'bruno.s'
senha_rdp = '1Pyth0n@'
bot.FAILSAFE = False


def abre_mercantil():
    print(Fore.GREEN + '\n--- Executando a função: ABRE MERCANTIL ' + Style.RESET_ALL)
    bot.pause = 0.5
    verifica_topcompras = 0
    
    corrige_nometela() # Verifica se o TopCompras não está com outro nome.
    
    while ahk.win_exists('TopCompras', title_match_mode= 2): # Realiza o fechamento do TopCompras, caso esteja aberto
        ahk.win_close('TopCompras', title_match_mode=2)
        time.sleep(0.5)
        if verifica_topcompras > 10:
            print('--- Tentou fechar o TopCompras porém não conseguiu! Fechando o remote por completo')
            return False
            
        verifica_topcompras += 1
    else:
        print('--- Modulo de compras não está aberto, abrindo o modulo')
            
    # Ativa o Topcon, e clica no topcompras, e executa a função para correção do nome.
    ahk.win_activate('TopCon', title_match_mode= 2)
    bot.click(procura_imagem(imagem='img_topcon/logo_topcompras.png'))
    time.sleep(2)
    if ahk.win_exists(title= 'TopCompras - Versão', title_match_mode = 1) is False:
        corrige_nometela()

    try: #Abre o TopCompras, e verifica se aparece a tela "interveniente"
        ahk.win_wait('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 10)
        while ahk.win_exists('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2):
            print('--- Encontrou a tela do interveniente, clicando no botão "OK"')
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2)
            ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 5)
            time.sleep(0.2)
            
            while procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True, confianca= 0.74, limite_tentativa= 1) is not False:
                if procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True):
                    print('--- Encontrou a tela do interveniente, clicando no botão "OK"')
                    bot.press('ENTER')
                else:
                    print('--- Não exibiu a tela de interveniente.')
        else:
            ahk.win_wait_close('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 10)
            print('--- Fechou a tela dos intervenientes')
    except TimeoutError:
        print('--- Apresentou um erro')
        time.sleep(1)
        abre_topcon()
    
    return navega_topcompras()
        
def navega_topcompras():
    print(Fore.GREEN + '--- Executando a função: navega topcompras ' + Style.RESET_ALL)
    # Navegando entre os menus para abrir a opção "Compras - Mercantil"
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 30)
    time.sleep(0.5)
    bot.click(900, 900)
    bot.press('ALT')
    bot.press('RIGHT', presses= 2, interval= 0.05)
    bot.press('DOWN', presses= 7, interval= 0.05)
    bot.press('ENTER')
    time.sleep(1)
    print(Fore.GREEN +  '--- TopCompras aberto!' + Style.RESET_ALL)
    return True

def fecha_execucoes():
    print('--- Iniciando fecha_execucoes, para fechar o TopCompras e o RDP ---')
    
    # Primeiro força o fechamento do TopCompras, para evitar erros de validações
    while ahk.win_exists(title= "TopCompras", title_match_mode= 2):
        time.sleep(0.2)
        ahk.win_close(title= 'TopCompras', title_match_mode = 2)   
    else:
        time.sleep(0.2)

    #os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
    os.system('taskkill /im mstsc.exe /f /t') # Força o fechamento do processo do RDP por completo

def abre_topcon():
    while True:
        print(Fore.GREEN + '--- Executando a função: ABRE TOPCON ' + Style.RESET_ALL)
        bot.pause = 0.5
        fecha_execucoes() # Começa garantindo que fechou todas as execuções antigas.
        
        print('--- Iniciando o RemoteApp')
        os.startfile('RemoteApp-Cortesia.rdp')
        #os.startfile('RemoteApp-CortesiaVPN.rdp')
        
        telas_seguranca = ['Windows Security', 'Segurança do Windows'] # Tenta encontrar em ingles e portugues.
        tela_login_rdp = False
        
        for tela in telas_seguranca:
            print(F'--- Tentando a tela: {tela}')
            try:
                ahk.win_wait_active(tela, title_match_mode= 2, timeout= 3)
            except TimeoutError:
                time.sleep(0.5)
            else:
                #print(F'--- Encontrou com o nome {tela}')
                tela_login_rdp = tela
                ahk.win_activate(tela, title_match_mode= 2)
        
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

        if ahk.win_exists('TopCompras', title_match_mode= 2) is False: # Se o modulo de compras estiver fechado, realiza o login no TopCon
            while True: # Realiza login no TopCon
                ahk.win_activate('TopCon', title_match_mode= 2)
                ahk.win_wait_active('TopCon', title_match_mode= 2, timeout= 30)
                if procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74): # Aguarda até aparecer o campo do servidor preenchido
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
                    time.sleep(1)
                    break
                
                if procura_imagem(imagem='img_topcon/logo_principal.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74):
                    time.sleep(0.5)
                    print('--- Tela do Topcon já está aberta.')
                    break
        
        # Com o programa TopCon aberto, realiza a abertura do modulo de compras (mercantil). 
        if abre_mercantil() is True:
            print('--- processo concluido')
            break
        else:
            print('--- deu algum xabu')

if __name__ == '__main__':
    bot.pause = 0.25
    #abre_mercantil()
    #fecha_execucoes()
    abre_topcon()