# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import os
import time
import subprocess
from math import e
import pytesseract
import pyautogui as bot
from utils.configura_logger import get_logger
from utils.funcoes import ahk as ahk, msg_box
from utils.funcoes import procura_imagem, corrige_nometela, ativar_janela
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
senha_rdp = 'C0rtesi@2025'

def fechar_tela_nota_compra():
    logger.info('--- Executando função FECHAR TELA NOTA COMPRA')
    ativar_janela('TopCompras')
    
    #* Verifica se a tela "6201 NOTA FISCAL DE COMPRA" está aberta
    if procura_imagem(imagem='imagens/img_topcon/produtos_servicos.png', continuar_exec= True):
        logger.info('--- Encontrou a tela "6201 NOTA FISCAL DE COMPRA" realizando fechamento')
        bot.click(procura_imagem(imagem='imagens/img_topcon/bt_fechar.PNG', continuar_exec= True))
        time.sleep(0.5)

    #* Verifica se realmente fechou a tela "6201 NOTA FISCAL DE COMPRA"
    if procura_imagem(imagem='imagens/img_topcon/produtos_servicos.png', continuar_exec= True) is False:
        logger.info('--- Realmente fechou a tela 6201 NOTA FISCAL DE COMPRA')

def fechar_topcompras():
    """ #* Garante o fechamento do TopCompras, caso ele esteja aberto
    """    
    
    for i in range (0, 11):
        logger.info('--- Janela do TopCompras está aberta, forçando o fechamento')
        ahk.win_close('TopCompras', title_match_mode=2, seconds_to_wait= 3)
        time.sleep(0.5)

        if ahk.win_exists('TopCompras', title_match_mode= 2) is False:
            logger.info('--- Modulo de compras realmente está fechado! Pode continuar a abertura')
            return True        
        
        if i > 10:
            logger.warning('--- Não foi possivel fechar apenas o TopCompras, reiniciando Topcon & TopCompras')
            raise Exception('--- Não foi possivel fechar apenas o TopCompras, reiniciando Topcon & TopCompras')


def abre_mercantil():
    logger.info('--- Realizando a abertura do modulo de compras')
    time.sleep(1)

    ativar_janela('TopCon', 30)
    logger.info('--- Clicando para abrir o modulo de compras')
    bot.click(procura_imagem(imagem='imagens/img_topcon/icone_mercantil.png', limite_tentativa= 15), clicks= 2)
    time.sleep(3)

    #* Caso não encontre o TopCompras, tenta corrigir o nome
    corrige_nometela("TopCompras (")
    
    #* Verifica se o pop-up "interveniente" está aberto
    logger.info('--- Verificando se o pop-up do interveniente está aberto')
    for i in range (0, 5):
        time.sleep(0.2)
        if ahk.win_exists("TopCompras (", title_match_mode= 1):
            logger.info('--- Fechando a tela "interveniente" ')
            ahk.win_close("TopCompras (", title_match_mode= 2, seconds_to_wait= 10)
        else:
            corrige_nometela("TopCompras")
            time.sleep(0.2)
            ativar_janela('TopCompras', 30)
            time.sleep(0.2)
            if procura_imagem('imagens/img_topcon/produtos_servicos.png', limite_tentativa= 12):
                logger.success("Concluiu a task ABRE MERCANTIL")
                break
        
        if i >= 4:
            logger.error(F"Não foi possivel fechar a tela 'interveniente' (TopCompras (! Tentativas executadas: {i}")
            raise Exception('Não foi possivel fechar a tela "TopCompras (", necessario reiniciar o TopCon')
    else:
        logger.error(F"Não foi possivel fechar a tela 'interveniente' (TopCompras (! Tentativas executadas: {i}")
        raise Exception('Não foi possivel fechar a tela "TopCompras (", necessario reiniciar o TopCon')

''' #! Não é mais necessario, agora abre direto na tela 6201 - Compras mercantil
def navega_topcompras():
    bot.PAUSE = 1.2
    logger.info('--- Executando a função: navega topcompras ' )
    # Navegando entre os menus para abrir a opção "Compras - Mercantil"
    ativar_janela('TopCompras', 30)
    time.sleep(2)
    bot.click(900, 900)
    bot.press('ALT')
    bot.press('RIGHT', presses= 2, interval= 0.05)
    bot.press('DOWN', presses= 7, interval= 0.05)
    bot.press('ENTER')
    time.sleep(3)
    logger.success("Concluiu a função NAVEGA TOPCOMPRAS")
    return True
'''


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
            subprocess.run(["taskkill", "/im", "mstsc.exe", "/f", "/t"], stderr=subprocess.DEVNULL)
    else:
        time.sleep(0.4)
        logger.info('--- Fechou a tela "TopCompras" ')

    logger.info('--- Fechando os processos do RemoteDesktop ---')
    subprocess.run(["taskkill", "/im", "wksprt.exe", "/f", "/t"], stderr=subprocess.DEVNULL)
    subprocess.run(["taskkill", "/im", "mstsc.exe", "/f", "/t"], stderr=subprocess.DEVNULL)
    logger.info('--- Os processos wksprt e mstsc.exe do RDP')

    logger.success('--- Concluiu a task FECHA EXECUÇÕES')


def login_topcon():
    bot.PAUSE = 1.4

    logger.info('--- Realizando login no TOPCON')
    #* Se o modulo de compras estiver fechado, realiza o login no TopCon
    if ahk.win_exists('TopCompras', title_match_mode= 2) is False: 
        time.sleep(1)
        ativar_janela('TopCon', 30)
        
        #* Valida se realmente realizou o Login no TopCon ou se já iniciou logado
        for i in range(0, 5):
            if procura_imagem(imagem='imagens/img_topcon/logo_topcon_grande.png', continuar_exec= True, limite_tentativa= 4):
                logger.info('--- Já está logado no Topcon!')
                ahk.win_maximize('TopCon', title_match_mode= 2)
                return True

            #* Aguarda até aparecer o campo do servidor de aplicação
            elif procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png', continuar_exec= True, limite_tentativa= 4): 
                logger.info('--- Tela de login do topcon aberta')
                bot.click(procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png'))
                bot.press('tab', presses= 2, interval= 0.05)
                bot.press('backspace')
                
                #* Insere os dados de login do usuario BRUNO.S
                logger.info('--- Inserindo dados para login')
                bot.write('BRUNO.S')
                bot.press('tab')
                bot.write('rockie')
                bot.press('tab')
                bot.press('enter')
                time.sleep(5)
                logger.success("Login realizado com sucesso!")
                return True
            else: #* Caso não encontre a tela do para realizar o Login no TopCon
                logger.warning(f"Tentativa {i+1}/5: não encontrou tela de login nem logo do TopCon.")
                time.sleep(1)
        else:
            raise Exception(f"Login no TopCon não foi concluído após {i} tentativas.")

    if procura_imagem('imagens/img_topcon/txt_OLA_BRUNO.png'):
        logger.success("Concluiu o Login no TopCon")


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
        ahk.win_wait("RemoteApp", title_match_mode= 3, timeout= 10)
        
        
        # Verifica se ao abrir, já começou logado no RemoteDesktop
        for i in range (0, 5):
            time.sleep(0.5)
            corrige_nometela("TopCon - Versão")
            if ahk.win_exists('TopCon', title_match_mode= 2):
                logger.info('--- RemoteDesktop já está logado! Não é necessario fazer o processo.')
                return True

        # Verifica se abriu alguma das telas de segurança de execução do RDP
        telas_seguranca = ['Windows Security', 'Segurança do Windows']
        tela_login_rdp = False
        for i in range (0, 10):
            time.sleep(0.4)
            if procura_imagem(imagem='imagens/img_topcon/txt_ServidorAplicacao.png', continuar_exec= True):
                tela_login_rdp = True
                break

            if tela_login_rdp is not False:
                break
            
            for tela in telas_seguranca:
                logger.info(F'--- Tentando a tela: {tela}')
                try:
                    ahk.win_wait_active(tela, title_match_mode= 2, timeout= 3)
                except TimeoutError:
                    time.sleep(0.2)
                else:
                    logger.info(F'--- Encontrou com o nome {tela}')
                    tela_login_rdp = tela
                    ahk.win_activate(tela, title_match_mode= 2)
            

        if tela_login_rdp is False: # Caso não encontro nenhuma das telas da lista "telas_seguranca"
            raise Exception("Não foi possivel encontrar nenhuma das telas de login do RDP")
        
        #* Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
        if ahk.win_exists(tela_login_rdp, title_match_mode= 2):
            logger.info(F'--- Abriu a tela "{tela_login_rdp}", realizando o login" ')

            ativar_janela(tela_login_rdp, 10)
            
            #* Insere os dados de login.
            bot.write(senha_rdp, interval= 0.1) #Senha BRUNO.S 
            bot.press('TAB', presses= 3, interval= 0.08) # Navega até o botão "Ok"
            bot.press('ENTER')
            time.sleep(3)
            logger.info('--- Login realizado no RemoteApp-Cortesia.rdp')
        elif ahk.win_exists('TopCon (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2): 
            logger.info('--- Tela de login do Topcon já está aberta, prosseguindo para o login')
            break

        #* Verifica se apareceu a tela para login já dentro do Topcon
        logger.info('--- Verificando se já está logado no Topcon')
        for i in range (0, 5):
            ativar_janela('TopCon', 30)
            time.sleep(2)

            # Caso apareça "Olá Bruno" já está logado!
            if procura_imagem(imagem='imagens/img_topcon/txt_OLA_BRUNO.png', continuar_exec= True) is False: 
                time.sleep(1)
                ativar_janela('TopCon', 30)

                # Caso apareça o logo de login, precisa entrar no topcon!
                if procura_imagem(imagem='imagens/img_topcon/logo_topcon_login.png', continuar_exec= True): 
                    logger.success("Concluiu a task ABRE TOPCON")
                    return True
            else:
                return True
        else:
            logger.success("Concluiu a task ABRE TOPCON")
            return True



def main():
    ultimo_erro = ""
    for tentativa in range(0, 10):
        bot.PAUSE = 1.2
        logger.info(F"Tentativa de abrir o topcon: {tentativa}")
        try:
            time.sleep(0.5)
            fecha_execucoes()
            abre_topcon()
            login_topcon()
            fechar_topcompras()
            abre_mercantil()
            if tentativa >= 6:
                break
        except Exception as e:
            logger.error(f"Erro: {type(e).__name__} - {str(e)}")
            #logger.error(F"Apresentou um erro! {ultimo_erro}")
            logger.debug("Traceback completo:", exc_info=True)

            if tentativa >= 9:
                logger.critical(F"Função ABRE TOPCON apresentou erro critico! {ultimo_erro}")
                return ultimo_erro
        else:
            logger.success("Executou o ABRE_TOPCON.PY com sucesso!")
            return True
    else:   
        raise ultimo_erro
    
if __name__ == '__main__':
    tempo_inicial = time.time()
    main()

    # Linha específica onde você quer medir o tempo
    end_time = time.time()
    elapsed_time = end_time - tempo_inicial
    medicao_minutos = elapsed_time / 60
    print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")
    time.sleep(1)
    bot.alert("acabou")

    '''
    fechar_tela_nota_compra()
    navega_topcompras()
    msg_box("Conclui a TASK Abre TopCon", 10)
    '''
