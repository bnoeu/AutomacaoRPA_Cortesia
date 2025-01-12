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
logger = get_logger("script1", print_terminal = True)

#! Dados para acesso ao remoteApp
'''Usuario antigo
login_rdp = 'bruno.s'
senha_rdp = '4Pyth0n@'
'''
#* Novo usuario
bot.FAILSAFE = False
login_rdp = 'b.santos'
senha_rdp = 'Cortesia@123'

def abre_mercantil():
    #bot.alert(F"Valor da pausa: {bot.PAUSE}")
    logger.info('--- Realizando a abertura do modulo de compras')
    verifica_topcompras = 0
    
    #* Realiza o fechamento do TopCompras, caso esteja aberto
    while ahk.win_exists('TopCompras', title_match_mode= 2): 
        logger.info('--- Janela do TopCompras já aberta, forçando o fechamento')
        ahk.win_close('TopCompras', title_match_mode=2, seconds_to_wait= 3)
        
        if ahk.win_exists('TopCompras', title_match_mode= 2) is False:
            logger.info('--- Modulo de compras realmente está fechado, abrindo nova execução do modulo')
            break
        
        verifica_topcompras += 1
        if verifica_topcompras > 10:
            logger.warning('--- Não foi possivel fechar apenas o TopCompras, reiniciando Topcon & TopCompras')
            fecha_execucoes()

    #* Clica para abrir o modulo de compras
    time.sleep(8)
    ahk.win_activate('TopCon', title_match_mode= 2)
    ahk.win_wait_active('TopCon', title_match_mode= 2, timeout= 30)
    bot.click(procura_imagem(imagem='imagens/img_topcon/icone_modulo_compras.png', confianca= 0.67, limite_tentativa= 10))
    
    #* Caso não encontre o TopCompras, tenta corrigir o nome
    #if ahk.win_exists(title= 'TopCompras - Versão', title_match_mode = 1) is False: 
    corrige_nometela()
    
    #* Verifica se o pop-up "interveniente" está aberto
    for i in range (0, 5):
        bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec= True))

    logger.success("Concluiu a task ABRE MERCANTIL")
        
def navega_topcompras():
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
    logger.success("Concluiu a task NAVEGA TOPCOMPRAS")
    return True

def fecha_execucoes():
    """#* Realiza o fechamento completo do TopCon e TopCompras
    """    
    
    logger.info('--- Iniciando fecha execucoes, para fechar o TopCompras e o RDP ---')
    limite_tentativas = 0

    #* Verifica se a tela "Vinculação itens da NFE" está aberta, e fecha ela.
    if ahk.win_exists('Vinculação Itens da Nota', title_match_mode = 2):
        ahk.win_close('Vinculação Itens da Nota', title_match_mode = 2, seconds_to_wait= 5)
        ahk.win_wait_close('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
        logger.info('--- Fechou a tela "Vinculação itens da nota" ')
        
    #* Primeiro força o fechamento do TopCompras, para evitar erros de validações
    while ahk.win_exists(title= "TopCompras", title_match_mode= 2):
        ahk.win_close(title= 'TopCompras', title_match_mode = 2, seconds_to_wait= 1)   
        ahk.win_kill(title='TopCompras', title_match_mode= 2, seconds_to_wait= 1)
        
        limite_tentativas += 1
        if limite_tentativas > 10:
            logger.error('--- Não conseguiu fechar a tela "TopCompras" ')
            os.system('taskkill /im mstsc.exe /f /t') # Força o fechamento do processo do RDP por completo
    else:
        time.sleep(0.4)
        logger.info('--- Fechou a tela "TopCompras" ')

    os.system('taskkill /PID 872 /f /t') # Força o fechamento do processo do RDP por completo
    os.system('taskkill /im wksprt.exe /f /t') # Força o fechamento do processo do RDP por completo
    os.system('taskkill /im mstsc.exe /f /t') # Força o fechamento do processo do RDP por completo
    logger.info('--- Os processos wksprt e mstsc.exe do RDP')

    logger.success('--- Concluiu a task FECHA EXECUÇÕES')

def login_topcon():
    #* Se o modulo de compras estiver fechado, realiza o login no TopCon
    if ahk.win_exists('TopCompras', title_match_mode= 2) is False: 
        time.sleep(8)
        ahk.win_activate('TopCon', title_match_mode= 2)
        ahk.win_wait_active('TopCon', title_match_mode= 2, timeout= 30)
        
        #* Valida se realmente realizou o Login no TopCon ou se já iniciou logado
        for i in range(0, 5):
            time.sleep(0.5)
            if procura_imagem(imagem='imagens/img_topcon/logo_topcon_grande.png', continuar_exec= True):
                time.sleep(0.5)
                logger.info('--- Tela do Topcon está aberta!')
                ahk.win_maximize('TopCon', title_match_mode= 2)
                logger.success("--- Concluiu a task LOGIN TOPCON")
                break

            #* Aguarda até aparecer o campo do servidor de aplicação
            if procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png', continuar_exec= True): 
                logger.info('--- Tela de login do topcon aberta')
                bot.click(procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png'))
                bot.press('tab', presses= 2, interval= 0.005)
                bot.press('backspace')
                
                #* Insere os dados de login do usuario BRUNO.S
                logger.info('--- Inserindo dados para login')
                bot.write('BRUNO.S')
                bot.press('tab')
                bot.write('rockie')
                bot.press('tab')
                bot.press('enter')
                time.sleep(1)
                logger.success("Login realizado com sucesso!")
                return True
            else: #* Caso não encontre a tela do para realizar o Login no TopCon
                raise Exception("Login no Topcon não foi concluido!")

def abre_topcon():
    logger.info('--- Executando a função: ABRE TOPCON' )

    #* Caso fique aparecendo a tela "RemoteApp Disconnected"
    if ahk.win_exists('RemoteApp Disconnected', title_match_mode= 2): 
        logger.debug('Fechando a tela "RemoteApp Disconnected" ')
        ahk.win_activate('RemoteApp Disconnected', title_match_mode= 2)
        time.sleep(0.5)
        bot.press('ENTER')
    
    while True:
        logger.info('--- Iniciando o RemoteApp')
        os.startfile('RemoteApp\RemoteApp-Cortesia.rdp')
        time.sleep(10)
        
        telas_seguranca = ['Windows Security', 'Segurança do Windows'] # Tenta encontrar em ingles e portugues.
        tela_login_rdp = False
        
        for i in range (0, 10):
            time.sleep(0.5)
            if ahk.win_exists('TopCon', title_match_mode= 2):
                return True

        for i in range (0, 5):
            time.sleep(0.4)
            if procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png', continuar_exec= True):
                tela_login_rdp = True
                break

            if tela_login_rdp is not False:
                break
            
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
            raise Exception("Não foi possivel encontrar nenhuma das telas de login do RDP")
        
        #* Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
        if ahk.win_exists(tela_login_rdp, title_match_mode= 2):
            logger.info(F'--- Abriu a tela "{tela_login_rdp}", realizando o login" ')
            ahk.win_activate(tela_login_rdp, title_match_mode= 2)
            ahk.win_wait_active(tela_login_rdp, title_match_mode= 2, timeout = 10)
            
            #* Insere os dados de login.
            bot.write(senha_rdp, interval= 0.25) #Senha BRUNO.S 
            bot.press('TAB', presses= 3, interval= 0.05) # Navega até o botão "Ok"
            bot.press('ENTER')
            logger.info('--- Login realizado no RemoteApp-Cortesia.rdp')
        elif ahk.win_exists('TopCon (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2): 
            logger.info('--- Tela de login do Topcon já está aberta, prosseguindo para o login')

        '''
        
        #* Se o modulo de compras estiver fechado, realiza o login no TopCon
        if ahk.win_exists('TopCompras', title_match_mode= 2) is False: 
            time.sleep(8)
            ahk.win_activate('TopCon', title_match_mode= 2)
            ahk.win_wait_active('TopCon', title_match_mode= 2, timeout= 30)
            
            for i in range(0, 5):
                #* Valida se realmente realizou o Login no TopCon ou se já iniciou logado
                time.sleep(2)
                if procura_imagem(imagem='imagens/img_topcon/logo_topcon_grande.png', continuar_exec= True):
                    time.sleep(0.5)
                    logger.info('--- Tela do Topcon está aberta!')
                    ahk.win_maximize('TopCompras', title_match_mode= 2)
                    break

                #* Aguarda até aparecer o campo do servidor de aplicação
                if procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png', continuar_exec= True): 
                    logger.info('--- Tela de login do topcon aberta')
                    bot.click(procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png'))
                    bot.press('tab', presses= 2, interval= 0.005)
                    bot.press('backspace')
                    
                    #* Insere os dados de login do usuario BRUNO.S
                    bot.write('BRUNO.S')
                    bot.press('tab')
                    bot.write('rockie')
                    bot.press('tab')
                    bot.press('enter')
                    time.sleep(0.5)
                else: #* Caso não encontre a tela do para realizar o Login no TopCon
                    raise Exception("Login no Topcon não foi concluido!")
'''

        logger.success("Concluiu a task ABRE TOPCON")
        return True


def main():
    ultimo_erro = ""
    for tentativa in range(0, 7):
        logger.info(F"Tentativa de abrir o topcon: {tentativa}")
        try:
            time.sleep(0.5)
            bot.PAUSE = 1
            fecha_execucoes()
            abre_topcon()
            login_topcon()
            abre_mercantil()
            navega_topcompras()
            if tentativa >= 6:
                break
        except Exception as e:
            ultimo_erro = e
            if tentativa >= 5:
                logger.critical(F"Função ABRE TOPCON apresentou erro critico! {ultimo_erro}")
                return ultimo_erro
            logger.error(F"Apresentou um erro! {ultimo_erro}")
        else:
            logger.success("Executou a TASK ABRE TOCPON com sucesso!")
            break
    else:
        raise ultimo_erro
    
if __name__ == '__main__':
    #fecha_execucoes()
    #abre_topcon()
    #login_topcon()
    main()