# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pyautogui as bot
from abre_topcon import navega_topcompras, fechar_tela_nota_compra
#from automacao.conferencia_xml import conferencia_xml
#from coleta_planilha import main as coleta_planilha
from utils.funcoes import ahk as ahk
from utils.funcoes import procura_imagem, ativar_janela
from utils.configura_logger import get_logger

# --- Definição de parametros
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
logger = get_logger("script1")

def alterar_localizar():
    logger.info('--- Alterando o TopCompras para o modo LOCALIZAR')
    
    for i in range(0, 5):
        ativar_janela("TopCompras", timeout= 15)
        time.sleep(0.2)
        
        if procura_imagem(imagem='imagens/img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
            logger.info('--- Está no modo "incluir", enviando comando F2 para entrar no modo "Localizar"')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('F2', presses= 2)

        '''
        if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
            logger.info('--- Está no modo "Localizar" Alterando para "Incluir"')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('F3', presses= 2)
        '''

        if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
            logger.success('--- Está no modo "LOCALIZAR", lançamento pode continuar!')
            return True

        #* Após 5 execuções, tenta reabrir o TopCompras antes de prosseguir
        if i == 2:
            fechar_tela_nota_compra()
            navega_topcompras()
            

        if i >= 4:
            logger.error('--- Atingiu o maximo de tentativas de alterar os botões ---')
            #bot.alert("Limite de tentativas de alterar os botões")        
            raise Exception("Atingiu o maximo de tentativas de alterar os botões")
        

if __name__ == '__main__':
    alterar_localizar()