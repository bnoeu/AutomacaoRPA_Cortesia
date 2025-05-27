import time
import pyperclip
import pyautogui as bot
from utils.configura_logger import get_logger
from utils.funcoes import ahk as ahk # --- Definição de parametros


def copia_linha_atual():
    logger = get_logger("script1")
    bot.PAUSE = 0.1
    dados_planilha = []
    coluna_atual = 0
    
    # Seleciona a linha inteira e copia, avaliando se coletou ou ficou "Recuperando"
    bot.hotkey('Shift', 'Space')
    for i in range (0, 30):
        logger.debug('--- Tentando copiar os dados.')
        bot.hotkey('ctrl', 'c')
        time.sleep(0.5)
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

    '''
    while coluna_atual < 9:
        bot.hotkey('Shift', 'Space')
        for i in range (0, 30):
            bot.hotkey('ctrl', 'c')
            time.sleep(0.5)
            valor_copiado = pyperclip.paste()
            if "Recuperando" in valor_copiado:
                time.sleep(0.5)
                continue
            else:
                break
            
        if 'Recuperando' not in valor_copiado:
            dados_planilha.append(valor_copiado)
            bot.press('right')
            coluna_atual += 1
        else:
            if i >= 15:
                print("limite tentativas")
            logger.debug(F'Copiou o texto "Recuperando Dados", tentando copiar novamente {dados_planilha}, tentativa: {i}')
            time.sleep(0.3)
    '''

    return dados_planilha


if __name__ == '__main__':
    copia_linha_atual()
