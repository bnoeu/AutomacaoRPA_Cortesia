# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import cv2
import pygetwindow as gw
import pytesseract
from ahk import AHK
from datetime import date
from funcoes import procura_imagem, alteracao_filtro
#from valida_pedido import valida_pedido
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"  # Valor para teste
transportador = "111594"  # Valor para teste
#tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"

def coleta_planilha():
    bot.PAUSE = 1  # Pausa padrão do bot
    print('--- Abrindo planilha - COLETA_PLANILHA')
    ahk.win_activate('db_alltrips', title_match_mode= 2)
    time.sleep(1)
    if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True) is False:
        bot.click(procura_imagem(imagem='img_planilha/botao_edicao.png'))
        bot.click(procura_imagem(imagem='img_planilha/botao_exibicao.png'))
        time.sleep(1)
        procura_imagem(imagem='img_planilha/botao_exibicaoverde.png')
    else:
        print('--- Já está no modo de edição, continuando processo')
        # Validação se houve novo valor inserido
    alteracao_filtro()
    # * Coleta os dados da linha atual
    dados_planilha = []
    print('--- Copiando dados e formatando')
    bot.click(100, 510)  # Clica na primeira linha
    for n in range(0, 7, 1):  # Copia dados dos 6 campos
        while True:
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                print('Tentando copiar novamente')
                time.sleep(0.15)
            else:
                break
        dados_planilha.append(ahk.get_clipboard())
        bot.press('right')
    bot.PAUSE = 1.8  # Pausa padrão do bot
    return dados_planilha