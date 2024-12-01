# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import os
import time
import pytesseract
import pyautogui as bot
from utils.funcoes import ahk as ahk
from utils.configura_logger import get_logger
from utils.funcoes import procura_imagem, corrige_nometela
#from colorama import Style

# Definição de parametros
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
logger = get_logger("script1")

#! Dados para acesso ao remoteApp
login_rdp = 'bruno.s'
senha_rdp = '3Pyth0n@'
bot.FAILSAFE = False


def abre_mercantil():
    logger.info('--- Realizando a abertura do modulo de compras')
    bot.PAUSE = 0.2
    verifica_topcompras = 0
    
    while ahk.win_exists('TopCompras', title_match_mode= 2): # Realiza o fechamento do TopCompras, caso esteja aberto
        logger.info('--- Janela do TopCompras já aberta, forçando o fechamento')
        ahk.win_close('TopCompras', title_match_mode=2, seconds_to_wait= 3)
        
        if ahk.win_exists('TopCompras', title_match_mode= 2) is False:
            logger.info('--- Modulo de compras realmente está fechado, abrindo nova execução do modulo')
            break
        
        verifica_topcompras += 1
        if verifica_topcompras > 10:
            logger.warning('--- Não foi possivel fechar apenas o TopCompras, reiniciando Topcon & TopCompras')
            fecha_execucoes()

    # Clica para abrir o modulo de compras
    time.sleep(8)
    ahk.win_activate('TopCon', title_match_mode= 2)
    ahk.win_wait_active('TopCon', title_match_mode= 2)
    bot.click(procura_imagem(imagem='imagens/img_topcon/icone_modulo_compras.png', confianca= 0.67, limite_tentativa= 10))
    if ahk.win_exists(title= 'TopCompras - Versão', title_match_mode = 1) is False: # Caso não encontre o TopCompras, tenta corrigir o nome
        corrige_nometela()
    
    #* Verifica se o pop-up "interveniente" está aberto
    bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec= True))
    '''
    if ahk.win_exists(title= 'TopCompras', title_match_mode = 3):
        logger.info('--- Pop-up de inverveniente aberto, fechando para continuar!')
        ahk.win_activate(title= 'TopCompras', title_match_mode = 3)
        bot.press('enter')
        ahk.win_wait_close(title= 'TopCompras', title_match_mode = 3)
    '''

    return navega_topcompras()
        
def navega_topcompras():
    bot.PAUSE = 1
    logger.info('--- Executando a função: navega topcompras ' )
    # Navegando entre os menus para abrir a opção "Compras - Mercantil"
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 30)
    time.sleep(2)
    bot.click(900, 900)
    bot.press('ALT')
    bot.press('RIGHT', presses= 2, interval= 0.05)
    bot.press('DOWN', presses= 7, interval= 0.05)
    bot.press('ENTER')
    time.sleep(0.5)
    logger.info( '--- TopCompras aberto!' )
    return True

def fecha_execucoes():
    bot.PAUSE = 1
    logger.info('--- Iniciando fecha execucoes, para fechar o TopCompras e o RDP ---')
    
    if ahk.win_exists('Vinculação Itens da Nota', title_match_mode = 2):
        ahk.win_close('Vinculação Itens da Nota', title_match_mode = 2, seconds_to_wait= 5)
        ahk.win_wait_close('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
        logger.info('--- Fechou a tela "Vinculação itens da nota" ')
        
    # Primeiro força o fechamento do TopCompras, para evitar erros de validações
    limite_tentativas = 0
    while ahk.win_exists(title= "TopCompras", title_match_mode= 2):
        ahk.win_close(title= 'TopCompras', title_match_mode = 2, seconds_to_wait= 1)   
        ahk.win_kill(title='TopCompras', title_match_mode= 2, seconds_to_wait= 1)
        
        limite_tentativas += 1
        if limite_tentativas > 10:
            os.system('taskkill /im mstsc.exe /f /t') # Força o fechamento do processo do RDP por completo
    else:
        time.sleep(0.25)
        logger.info('--- Fechou a tela "TopCompras" ')

    #os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
    os.system('taskkill /im mstsc.exe /f /t') # Força o fechamento do processo do RDP por completo
    os.system('taskkill /im chrome.exe /f /t') # Força o fechamento do processo do RDP por completo
    logger.info('--- Fechou o processo MSTSC.EXE" ')

def abre_topcon():
    bot.PAUSE = 1.5
    logger.info('--- Executando a função: ABRE TOPCON' )
    if ahk.win_exists('RemoteApp Disconnected', title_match_mode= 2): # Caso fique aparecendo a tela "RemoteApp Disconnected"
        logger.debug('Fechando a tela "RemoteApp Disconnected" ')
        ahk.win_activate('RemoteApp Disconnected', title_match_mode= 2)
        time.sleep(0.5)
        bot.press('ENTER')
    
    while True:
        fecha_execucoes() # Começa garantindo que fechou todas as execuções antigas.
        
        logger.info('--- Iniciando o RemoteApp')
        os.startfile('RemoteApp\RemoteApp-Cortesia.rdp')
        time.sleep(10)
        
        telas_seguranca = ['Windows Security', 'Segurança do Windows'] # Tenta encontrar em ingles e portugues.
        tela_login_rdp = False
        
        for tela in telas_seguranca:
            logger.info(F'--- Tentando a tela: {tela}')
            try:
                ahk.win_wait_active(tela, title_match_mode= 2, timeout= 7)
            except TimeoutError:
                time.sleep(0.5)
            else:
                logger.info(F'--- Encontrou com o nome {tela}')
                tela_login_rdp = tela
                ahk.win_activate(tela, title_match_mode= 2)
        
        if tela_login_rdp is False: # Caso não encontro nenhuma das telas da lista "telas_seguranca"
            raise TimeoutError
        
        # Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
        if ahk.win_exists(tela_login_rdp, title_match_mode= 2):
            logger.info(F'--- Abriu a tela "{tela_login_rdp}", realizando o login" ')
            ahk.win_activate(tela_login_rdp, title_match_mode= 2)
            ahk.win_wait_active(tela_login_rdp, title_match_mode= 2, timeout = 10)
            
            # Insere os dados de login.
            bot.write(senha_rdp) #Senha BRUNO.S 
            bot.press('TAB', presses= 3, interval= 0.05) # Navega até o botão "Ok"
            bot.press('ENTER')
            logger.info('--- Login realizado no RemoteApp-Cortesia.rdp')
            
        elif ahk.win_exists('TopCon (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2): 
            logger.info('--- Tela de login do Topcon já está aberta, prosseguindo para o login')

        '''
        if ahk.win_exists('TopCompras', title_match_mode= 2) is False: # Se o modulo de compras estiver fechado, realiza o login no TopCon
            while True: # Realiza login no TopCon
                ahk.win_activate('TopCon', title_match_mode= 2)
                ahk.win_wait_active('TopCon', title_match_mode= 2, timeout= 30)
                if procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74): # Aguarda até aparecer o campo do servidor preenchido
                    logger.info('--- Tela de login do topcon aberta')
                    bot.click(procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png'))
                    bot.press('tab', presses= 2, interval= 0.005)
                    bot.press('backspace')
                    
                    # Insere os dados de login do usuario BRUNO.S
                    bot.write('BRUNO.S')
                    bot.press('tab')
                    bot.write('rockie')
                    bot.press('tab')
                    bot.press('enter')
                    time.sleep(0.5)
                    break
                
                if procura_imagem(imagem='imagens/img_topcon/logo_principal.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74):
                    time.sleep(0.5)
                    logger.info('--- Tela do Topcon já está aberta.')
                    break
        '''        
        abre_mercantil()
        return True

if __name__ == '__main__':
    bot.PAUSE = 1
    #print(abre_mercantil())
    #fecha_execucoes()
    abre_topcon()