# Para utilização na Cortesia Concreto.
# -*- Criado por Bruno da Silva Santos. -*-

from utils.funcoes import ahk as ahk
from utils.configura_logger import get_logger

'''
from datetime import datetime
import time
import pyautogui as bot
from funcoes import procura_imagem, abre_planilha_navegador, ativar_janela, reaplica_filtro_status
'''


# --- Definição de parametros
logger = get_logger("script1")
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"



def ativa_planilha_original():
    """Ativa a tela da planilha original "db_alltrips"

    Returns:
        boolean: True = achou | False = Não achou
    """

    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
    try:
        ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 1, timeout= 10)
        return True
    except TimeoutError:
        logger.warning('--- Planilha não encontrada!')
        return False
    
if __name__ == '__main__':
    ativa_planilha_original()