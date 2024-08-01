# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import os
# import cv2
# import pygetwindow as gw
import pytesseract
from ahk import AHK
from colorama import Fore, Style
from funcoes import procura_imagem, marca_lancado
import pyautogui as bot

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

def abre_planilha():
        # Verifica quais das planilhas está aberta, debug ou o banco puro.
        print(Fore.GREEN + '--- Abrindo planilha - COLETA_PLANILHA' + Style.RESET_ALL)
        if ahk.win_exists('debug_db_alltrips', title_match_mode= 2):
            print('--- Abrindo a planilha de debug')
            ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
            ahk.win_wait('debug_db_alltrips', title_match_mode= 2)
        else:
            print('--- Abrindo a planilha com o banco puro')
            ahk.win_activate('db_alltrips', title_match_mode= 2)
            ahk.win_wait('db_alltrips', title_match_mode= 2)

def coleta_planilha():
    while True:
        bot.PAUSE = 0.2
        # Abre a tela da planilha, que já deve ter sido acessada no Edge
        abre_planilha()
        
        # Processo necessario apenas caso rode diretamente no db_alltrips.
        if ahk.win_exists('debug_db_alltrips', title_match_mode= 2) is False:
            bot.hotkey('CTRL', 'HOME')
            # Verifica se já está no modo de edição, caso esteja, muda para o modo "exibição"
            if procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', continuar_exec=True, limite_tentativa= 3, confianca= 0.74) is False:
                print('--- Não está no modo exibição! Realizando alteração.')
                while procura_imagem(imagem='img_planilha/bt_edicao.png', continuar_exec= True, limite_tentativa= 3, confianca= 0.74) is False: #Espera até encontar o botão "Exibição" (Lapis bloqueado)
                    time.sleep(0.1)
                    
                if procura_imagem(imagem='img_planilha/bt_TresPontos.png', continuar_exec= True) is not False:
                    bot.click(procura_imagem(imagem='img_planilha/bt_TresPontos.png'))
                    
                bot.click(procura_imagem(imagem='img_planilha/bt_edicao.png'))  
                time.sleep(0.3)
                bot.click(procura_imagem(imagem='img_planilha/txt_exibicao.png'))

                #Aguarda até aparecer o botão do modo "exibição"
                while procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', continuar_exec=True, limite_tentativa= 3, confianca= 0.74) is False:
                    time.sleep(0.1)
                print('--- Alterado para o modo exibição, continuando.')
                
            else: # Caso não esteja no modo "Edição"
                print('--- A planilha já está no modo "Exibição", continuando processo')

        # Coleta os dados da linha atual
        dados_planilha = []
        print('--- Copiando dados e formatando')
        # Navega para a celula A1 ( RE ), em seguida vai para a primeira linha com dados a serem copiado
        bot.hotkey('CTRL', 'HOME')
        bot.press('DOWN')
        
        # Navega entre os 6 campos, realizando a copia um por um, e inserindo na lista Dados Planilha.
        for n in range(0, 7, 1):  # Copia dados dos 6 campos
            pausa_copia = 0.1
            while True:
                bot.hotkey('ctrl', 'c')
                if 'Recuperando' in ahk.get_clipboard():
                    time.sleep(pausa_copia)
                    pausa_copia += 0.1
                    print(F'--- Pausa copia aumentada para, {pausa_copia}')
                else:
                    break
            dados_planilha.append(ahk.get_clipboard())
            bot.press('right')
            
        # Realiza a validação dos dados copiados.
        chave_xml = dados_planilha[4].strip()
        if len(dados_planilha[6]) > 0:
            print(F'--- Aguardando 30s para aparecer novas notas, campo preenchido com: {dados_planilha[6]}, tamanho do status: {len(dados_planilha[6])}')
            time.sleep(30)
        if len(chave_xml) < 42:
            marca_lancado('chave_invalida')
        elif (len(dados_planilha[0]) < 4) or (len(dados_planilha[0]) == 5):
            marca_lancado('RE_Invalido')
        else:
            break
    
    #os.system('cls')
    print(Fore.GREEN + F'--- Dados copiados com sucesso: {dados_planilha}' + Style.RESET_ALL)
    return dados_planilha

if __name__ == '__main__':
    coleta_planilha()