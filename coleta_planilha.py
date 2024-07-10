# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
# import cv2
# import pygetwindow as gw
import pytesseract
from ahk import AHK
from colorama import Fore, Style
from funcoes import procura_imagem
import pyautogui as bot

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

def coleta_planilha():
    bot.PAUSE = 0.5
    print(Fore.GREEN + '--- Abrindo planilha - COLETA_PLANILHA' + Style.RESET_ALL)
    if ahk.win_exists('debug_db_alltrips', title_match_mode= 2):
        ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
        ahk.win_wait('debug_db_alltrips', title_match_mode= 2)
    else:
        ahk.win_activate('db_alltrips', title_match_mode= 2)
        ahk.win_wait('db_alltrips', title_match_mode= 2)
        
    if ahk.win_exists('debug_db_alltrips', title_match_mode= 2) is False:
        bot.hotkey('CTRL', 'HOME')
        #Verifica se já está no modo de edição, caso esteja, muda para o modo "exibição"
        if procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', continuar_exec=True) is False:
            print('--- Não está no modo exibição! Realizando alteração.')
            while procura_imagem(imagem='img_planilha/bt_edicao.png', continuar_exec= True) is False: #Espera até encontar o botão "Exibição" (Lapis bloqueado)
                time.sleep(0.1)
                
            if procura_imagem(imagem='img_planilha/bt_TresPontos.png', continuar_exec= True) is not False:
                bot.click(procura_imagem(imagem='img_planilha/bt_TresPontos.png'))
                
            bot.click(procura_imagem(imagem='img_planilha/bt_edicao.png'))  
            time.sleep(0.5)
            bot.click(procura_imagem(imagem='img_planilha/txt_exibicao.png'))

            #Aguarda até aparecer o botão do modo "exibição"
            while procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', continuar_exec=True) is False:
                time.sleep(0.1)
            print('--- Alterado para o modo exibição, continuando.')
            
        else: # Caso não esteja no modo "Edição"
            print('--- A planilha já está no modo "Exibição", continuando processo')

    # * Coleta os dados da linha atual
    dados_planilha = []
    # Clica na primeira linha (Campo RE), e pressiona seta para baixo
    print('--- Copiando dados e formatando')
    bot.hotkey('CTRL', 'HOME')
    bot.press('DOWN')
    for n in range(0, 7, 1):  # Copia dados dos 6 campos
        while True:
            pausa_copia = 0.1
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                time.sleep(pausa_copia)
                pausa_copia += 0.1
            else:
                break
        dados_planilha.append(ahk.get_clipboard())
        bot.press('right')
        
    tempo_coleta = time.time() - tempo_inicio
    tempo_coleta = tempo_coleta
    print(F'--- Tempo que levou: {tempo_coleta:0f} segundos')
    print(Fore.GREEN + F'--- Dados copiados com sucesso: {dados_planilha}\n' + Style.RESET_ALL)
    return dados_planilha

if __name__ == '__main__':
    coleta_planilha()