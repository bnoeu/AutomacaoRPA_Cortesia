import time
import pyperclip
import pyautogui as bot
from utils.configura_logger import get_logger
from utils.funcoes import ahk as ahk # --- Definição de parametros


def copia_linha_atual():
    logger = get_logger("script1")
    bot.PAUSE = 0.2
    dados_planilha = []
    
    # Seleciona a linha inteira e copia, avaliando se coletou ou ficou "Recuperando"
    ahk.win_activate("debug_db", title_match_mode = 2)
    time.sleep(0.2)
    bot.hotkey('Shift', 'Space')
    for i in range (0, 30):
        logger.debug('--- Tentando copiar os dados.')
        bot.hotkey('ctrl', 'c')
        time.sleep(0.4)
        valor_copiado = pyperclip.paste()
        if "Recuperando" in valor_copiado:
            logger.debug(F'Encontrou RECUPERANDO, tentando copiar novamente a linha atual: {dados_planilha}, tentativa: {i}')
            time.sleep(0.5)
            continue
        else:
            dados_planilha = valor_copiado.split('\t')
            return dados_planilha
    else:
        logger.warning(F'Limite de copiar linha atual atingindo, ultima coleta: {dados_planilha}, tentativa: {i}')

    return dados_planilha


if __name__ == '__main__':
    copia_linha_atual()
