import pyautogui as bot
from utils.configura_logger import get_logger
from utils.funcoes import ahk as ahk # --- Definição de parametros
logger = get_logger("automacao") # Obter logger configurado

def abre_planilha(planilha = "debug_db_alltrips"): # Realiza a abertura da planilha de debug
    logger.info(F'--- Abrindo a planilha: {planilha}')
    if ahk.win_exists(planilha, title_match_mode= 2):
        ahk.win_activate(planilha, title_match_mode= 2)
        ahk.win_wait(planilha, title_match_mode= 2, timeout= 15)
        
        logger.debug('--- Clicando no meio da planilha de debug.')
        bot.click(990, 700) # Clica no meio da planilha, para "firmar" a tela.
    else: # Caso a tela da planilha não seja encontrada
        logger.warning('--- Não encontrou a planilha')
        #TODO É necessario criar um script para reabrir a planilha caso não encontre
        exit(bot.alert('Não deu certo! '))
    
if __name__ == '__main__':
    bot.PAUSE = 0.4
    abre_planilha()