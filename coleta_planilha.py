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
bot.PAUSE = 1.5  # Pausa padrão do bot
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"  # Valor para teste
transportador = "111594"  # Valor para teste
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"

def coleta_planilha():
    print('--- Abrindo planilha')
    ahk.win_activate('db_alltrips')
    time.sleep(1)
    if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True, limite_tentativa=2) is False:
        bot.click(procura_imagem(imagem='img_planilha/botao_edicao.png'))
        bot.click(procura_imagem(imagem='img_planilha/botao_exibicao.png'))
        procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', limite_tentativa= 8)
    else:
        print('--- Já está no modo de edição, continuando processo')
        # Validação se houve novo valor inserido
        time.sleep(1)
    alteracao_filtro()
    # * Coleta os dados da linha atual
    dados_planilha = []
    print('--- Copiando dados e formatando')
    bot.click(100, 510)  # Clica na primeira linha
    bot.PAUSE = 0.3
    for n in range(0, 7, 1):  # Copia dados dos 6 campos
        while True:
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                print('Tentando copiar novamente')
                time.sleep(0.2)
            else:
                break
        dados_planilha.append(ahk.get_clipboard())
        bot.press('right')
    return dados_planilha
