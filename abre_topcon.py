# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
import pytesseract
from ahk import AHK
from utils.funcoes import procura_imagem, corrige_nometela, configurar_logging
import pyautogui as bot
import os
from colorama import Style

# Definição de parametros
ahk = AHK()
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
#! Dados para acesso ao remoteApp
login_rdp = 'bruno.s'
senha_rdp = '2Pyth0n@'
bot.FAILSAFE = False


def abre_mercantil():
    bot.PAUSE = 1
    verifica_topcompras = 0
    
    # Verificar se o TopCompras está aberto, caso não esteja, abre ele
    if ahk.win_exists('TopCon', title_match_mode= 2):
        logging.info('--- TopCon está aberto, pode prosseguir com a abertura do TopCompras')
    else:
        logging.warning('--- TopCon está fechado! executando ABRE TOPCON')
        abre_topcon()       
    
    if ahk.win_exists('TopCompras') is False:
        corrige_nometela() # Verifica se o TopCompras não está com outro nome.
        
    while ahk.win_exists('TopCompras', title_match_mode= 2): # Realiza o fechamento do TopCompras, caso esteja aberto
        logging.info('--- Tentando fechar o TopCompras')
        ahk.win_close('TopCompras', title_match_mode=2, seconds_to_wait= 3)
        
        if ahk.win_exists('TopCompras', title_match_mode= 2) is False:
            logging.info('--- Modulo de compras realmente está fechado, abrindo uma nova execução')
            return True
        
        verifica_topcompras += 1
        if verifica_topcompras > 10:
            logging.warning('--- Não foi possivel fechar apenas o TopCompras, reiniciando Topcon & TopCompras')
            fecha_execucoes()
            
    else:
        logging.info('--- Modulo de compras realmente está fechado, abrindo uma nova execução')
            
    # Ativa o Topcon, e clica no topcompras, e executa a função para correção do nome.
    ahk.win_activate('TopCon', title_match_mode= 2)
    bot.click(procura_imagem(imagem='imagens/img_topcon/logo_topcompras.png'))
    logging.info('--- Na tela do Topcon, clicou no logo do topcompras ')
    time.sleep(5)
    if ahk.win_exists(title= 'TopCompras - Versão', title_match_mode = 1) is False: # Caso não encontre o TopCompras
        corrige_nometela()

    try: #Abre o TopCompras, e verifica se aparece a tela "interveniente"
        ahk.win_wait('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 10)
        while ahk.win_exists('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2):
            logging.info('--- Encontrou a tela do interveniente, clicando no botão "OK"')
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2)
            ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 5)
            time.sleep(0.4)
            
            while procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec= True, confianca= 0.74, limite_tentativa= 1) is not False:
                if procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec= True):
                    logging.info('--- Encontrou a tela do interveniente, clicando no botão "OK"')
                    bot.press('ENTER')
                else:
                    logging.warning('--- Não exibiu a tela de interveniente.')
        else:
            ahk.win_wait_close('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 10)
            logging.info('--- Fechou a tela dos intervenientes')
    except TimeoutError:
        logging.error('--- Apresentou um erro')
        time.sleep(0.5)
        abre_topcon()
    corrige_nometela() # Verifica se o TopCompras não está com outro nome.
    return navega_topcompras()
        
def navega_topcompras():
    bot.PAUSE = 1
    logging.info('--- Executando a função: navega topcompras ' )
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
    logging.info( '--- TopCompras aberto!' )
    return True

def fecha_execucoes():
    bot.PAUSE = 1
    logging.info('--- Iniciando fecha_execucoes, para fechar o TopCompras e o RDP ---')
    
    if ahk.win_exists('Vinculação Itens da Nota', title_match_mode = 2):
        ahk.win_close('Vinculação Itens da Nota', title_match_mode = 2, seconds_to_wait= 5)
        
    
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

    #os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
    os.system('taskkill /im mstsc.exe /f /t') # Força o fechamento do processo do RDP por completo

def abre_topcon():
    bot.PAUSE = 1
    if ahk.win_exists('RemoteApp Disconnected', title_match_mode= 2): # Caso fique aparecendo a tela "RemoteApp Disconnected"
        logging.debug('Fechando a tela "RemoteApp Disconnected" ')
        ahk.win_activate('RemoteApp Disconnected', title_match_mode= 2)
        time.sleep(0.5)
        bot.press('ENTER')
    
    while True:
        logging.info('--- Executando a função: ABRE TOPCON ' )
        fecha_execucoes() # Começa garantindo que fechou todas as execuções antigas.
        
        logging.info('--- Iniciando o RemoteApp')
        os.startfile('RemoteApp\RemoteApp-Cortesia.rdp')
        #os.startfile('RemoteApp/RemoteApp-CortesiaVPN.rdp')
        
        telas_seguranca = ['Windows Security', 'Segurança do Windows'] # Tenta encontrar em ingles e portugues.
        tela_login_rdp = False
        
        for tela in telas_seguranca:
            logging.info(F'--- Tentando a tela: {tela}')
            try:
                ahk.win_wait_active(tela, title_match_mode= 2, timeout= 7)
            except TimeoutError:
                time.sleep(0.5)
            else:
                logging.info(F'--- Encontrou com o nome {tela}')
                tela_login_rdp = tela
                ahk.win_activate(tela, title_match_mode= 2)
        
        if tela_login_rdp is False: # Caso não encontro nenhuma das telas da lista "telas_seguranca"
            raise TimeoutError
        
        # Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
        if ahk.win_exists(tela_login_rdp, title_match_mode= 2):
            logging.info(F'--- Abriu a tela "{tela_login_rdp}", realizando o login" ')
            ahk.win_activate(tela_login_rdp, title_match_mode= 2)
            ahk.win_wait_active(tela_login_rdp, title_match_mode= 2, timeout = 10)
            
            # Insere os dados de login.
            bot.write(senha_rdp) #Senha BRUNO.S 
            bot.press('TAB', presses= 3, interval= 0.05) # Navega até o botão "Ok"
            bot.press('ENTER')
            logging.info('--- Login realizado no RemoteApp-Cortesia.rdp')
            
        elif ahk.win_exists('TopCon (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2): 
            logging.info('--- Tela de login do Topcon já está aberta, prosseguindo para o login')

        if ahk.win_exists('TopCompras', title_match_mode= 2) is False: # Se o modulo de compras estiver fechado, realiza o login no TopCon
            while True: # Realiza login no TopCon
                ahk.win_activate('TopCon', title_match_mode= 2)
                ahk.win_wait_active('TopCon', title_match_mode= 2, timeout= 30)
                if procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74): # Aguarda até aparecer o campo do servidor preenchido
                    logging.info('--- Tela de login do topcon aberta')
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
                    logging.info('--- Tela do Topcon já está aberta.')
                    break
        
        return True
        ''' #! Incluir todos os processo num chamado "reabertura_completa"
        # Com o programa TopCon aberto, realiza a abertura do modulo de compras (mercantil). 
        if abre_mercantil() is True:
            logging.info('--- processo concluido')
            break
        else:
            logging.critical('--- deu algum xabu')
        '''


if __name__ == '__main__':
    bot.PAUSE = 1
    configurar_logging(nome_arquivo= "logs/debug", nivel_log = logging.DEBUG)
    #fecha_execucoes()
    abre_mercantil()