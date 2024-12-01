import time
import pyperclip
import pyautogui as bot
from utils.configura_logger import get_logger
from utils.funcoes import ahk as ahk # --- Definição de parametros


def  copia_linha_atual():
    logger = get_logger("script1")
    bot.PAUSE = 0.2
    dados_planilha = []
    coluna_atual = 0
    
    while coluna_atual < 7: # Navega entre os 6 campos, realizando a copia um por um, e inserindo na lista Dados Planilha.
        while True:
            bot.hotkey('ctrl', 'c')
            time.sleep(0.2)
            valor_copiado = pyperclip.paste()
            if "Recuperando" == valor_copiado:
                time.sleep(0.2)
            else:
                break
            
        if 'Recuperando' not in valor_copiado:
            dados_planilha.append(valor_copiado)
            bot.press('right')
            coluna_atual += 1
        else:
            logger.debug(F'Copiou o texto "Recuperando Dados", tentando copiar novamente {dados_planilha}')

    return dados_planilha