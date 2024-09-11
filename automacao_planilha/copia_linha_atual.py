import time
from ahk import AHK
import pyautogui as bot

ahk = AHK() # --- Definição de parametros


def copia_linha_atual():
    bot.PAUSE = 0.1
    dados_planilha = []
    coluna_atual = 0
    
    while coluna_atual < 7: # Navega entre os 6 campos, realizando a copia um por um, e inserindo na lista Dados Planilha.
        while True:
            bot.hotkey('ctrl', 'c')
            time.sleep(0.5)
            if 'Recuperando' in ahk.get_clipboard():
                time.sleep(0.2)
            else:
                break
            
        if 'Recuperando' not in ahk.get_clipboard():
            dados_planilha.append(ahk.get_clipboard())
            print(dados_planilha)
            bot.press('right')
            coluna_atual += 1
        else:
            print(F'Copiou o texto "Recuperando Dados", tentando copiar novamente {dados_planilha}')

    return dados_planilha