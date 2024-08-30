# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import logging
import pyautogui as bot
from ahk import AHK

# --- Definição de parametros
ahk = AHK()


def abre_planilha(): # Realiza a abertura da planilha de debug
    logging.debug('--- Abrindo a planilha de debug.')
    if ahk.win_exists('debug_db_alltrips', title_match_mode= 2):
        ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
        ahk.win_wait('debug_db_alltrips', title_match_mode= 2, timeout= 15)
        
        logging.debug('--- Clicando no meio da planilha de debug.')
        bot.click(990, 700) # Clica no meio da planilha, para "firmar" a tela.
    else: # Caso a tela da planilha não seja encontrada
        #TODO É necessario criar um script para reabrir a planilha caso não encontre
        pass
    
if __name__ == '__main__':
    bot.PAUSE = 1.5
    abre_planilha()