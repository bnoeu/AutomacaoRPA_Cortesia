# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link do site do Senior
# https://logincloud.senior.com.br/logon/LogonPoint/tmindex.html

import os
import time
import pytesseract
from ahk import AHK
import pyautogui as bot
from datetime import date
from funcoes import procura_imagem

# --- Definição de parametros
ahk = AHK()
posicao_img = 0 
bot.PAUSE = 0.8
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '35240733039223000979550010003744871224990676', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"
alltrips = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg?e=RYfgcA.com"

# Receber nos parametros a chave_xml da ultima NFE coletada
def abre_planilha_navegador():
    print('--- Abrindo a planilha no EDGE.')
    comando_iniciar = F'start msedge {alltrips} -inprivate'
    os.system(comando_iniciar)
    time.sleep(15)
    ahk.win_wait_active('db_alltrips.xlsx', title_match_mode = 2, timeout= 15)
    ahk.win_maximize('db_alltrips.xlsx')
    print('--- Planilha aberta e maximizada.')

def encontra_ultimo_xml(ultimo_xml = ''):
    print(F'--- Iniciando a navegação até a ultima chave XML: {ultimo_xml}')
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
    # Clica no meio da planilha para "ativar" a navegação dentro dela.
    bot.click(1000, 600)
    
    # Move a navegação até a celula A1
    bot.hotkey('CTRL', 'HOME')
    
    # Navega até o campo "D. Insercao"]
    bot.press('RIGHT', presses= 8, interval= 0.05)
    
    #Abre o menu do filtro
    bot.hotkey('ALT', 'DOWN')
    
    # Pressiona enter, para aplicar o filtro 'Do menor ao maior"
    bot.press('ENTER')
    
    # Clica no meio da planilha para "ativar" a navegação dentro dela.
    bot.click(1000, 600)
    
    #Abre o menu de pesquisa 
    bot.press('ALT')
    bot.press('C')
    bot.press('F')
    bot.press('D')
    bot.press('F')
    
    # Insere a ultima chave copiada da planilha de debug
    bot.write(ultimo_xml)
    bot.press('ENTER')
    
    # Fecha o menu de pesquisa
    bot.press('ESC')
    bot.press('ALT', presses= 2)
    
    # Finaliza
    print(F'--- Concluido a navegação até a ultima chave XML: {ultimo_xml}')
    
def copia_dados():
    print('--- Iniciando o processo de cópia.')
    #ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
    
    # Navega até a proxima linha após a ultima chave.
    bot.press('DOWN')
    
    # Realiza uma avaliação, 
    # 1. Caso o campo esteja vazio, significa que ainda não foram inseridas novas notas, e para o processo. 
    # 2. Caso o campo esteja com uma chave XML nova, prossegue
    while True:
        bot.hotkey('ctrl', 'c')
        if 'Recuperando' in ahk.get_clipboard():
            print('--- Tentando copiar novamente.')
            time.sleep(0.4)
        else:
            print('--- Dado copiado com sucesso, realizando avaliação.')
            valor_copiado = ahk.get_clipboard()
            break
    
    if valor_copiado == "":
        print('--- Campo vazio, aguardando 10 minutos.')
        time.sleep(100)
    else:
        print(F'--- Uma nova chave foi inserida: {valor_copiado}')
    
    # Inicia o processo de seleção dos dados
    print('--- Iniciando o processo de seleção dos dados')    
    bot.press('LEFT', presses= 4)
    
    #bot.hotkey('ctrlleft', 'shiftleft', 'DOWN', interval= 0.05)
    #bot.keyDown('shiftleft')
    #bot.press('down', presses= 10)
    #bot.keyUp('shift')

    ahk.key_down('Shift')
    ahk.key_down('Control')
    
    ahk.key_press('down')
    
    ahk.key_up('Control')
    ahk.key_up('Shift')  # Release the key
    
    ahk.key_down('Shift')
    ahk.key_press('Right')
    ahk.key_press('Right')
    ahk.key_press('Right')
    ahk.key_up('Shift')  # Release the key

if __name__ == '__main__':
    #abre_planilha_navegador()
    encontra_ultimo_xml(chave_xml)
    copia_dados()

