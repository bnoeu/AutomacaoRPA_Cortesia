# Para utilização na Cortesia Concreto.
# -*- Criado por Bruno da Silva Santos. -*-

#! Link do site do Senior
# https://logincloud.senior.com.br/logon/LogonPoint/tmindex.html

import os
import time
import pytesseract
from ahk import AHK
import pyautogui as bot
from datetime import date
from funcoes import procura_imagem
from coleta_planilha import abre_planilha as abre_planilha_debug

# --- Definição de parametros
ahk = AHK()
posicao_img = 0 
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '35240829067113027809550070003077291319740149', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"
alltrips = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg?e=RYfgcA.com"

# Receber nos parametros a chave_xml da ultima NFE coletada
def abre_planilha_navegador(link_planilha = alltrips):
    bot.PAUSE = 1
    # Garante que a planilha não esteja aberta
    while ahk.win_exists('db_alltrips.xlsx', title_match_mode= 1):
        ahk.win_close('db_alltrips.xlsx', title_match_mode= 1)
        time.sleep(0.5)
    
    print('--- Abrindo a planilha no EDGE.')
    comando_iniciar = F'start msedge {link_planilha} -new-window -inprivate'
    os.system(comando_iniciar)
    time.sleep(15)
    ahk.win_wait_active('db_alltrips.xlsx', title_match_mode = 2, timeout= 15)
    ahk.win_maximize('db_alltrips.xlsx')
    print('--- Planilha aberta e maximizada.')

def encontra_ultimo_xml(ultimo_xml = ''):
    print(F'--- Iniciando a navegação até a ultima chave XML: {ultimo_xml}')
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
    try:
        ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 2, timeout= 10)
    except TimeoutError:
        print('--- Planilha não encontrada!')
        return True
    # Clica no meio da planilha para "ativar" a navegação dentro dela.
    bot.click(1000, 600)
    
    # Move a navegação até a celula A1
    bot.hotkey('CTRL', 'HOME')
    
    # Navega até o campo "D. Insercao"]
    bot.press('RIGHT', presses= 8, interval= 0.05)
    time.sleep(0.5)
    
    #Abre o menu do filtro
    bot.hotkey('ALT', 'DOWN')
    time.sleep(2)
    
    # Pressiona enter, para aplicar o filtro 'Do menor ao maior"
    bot.press('ENTER')
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
    
    #bot.click(procura_imagem(imagem='img_planilha/classific_menor_maior.PNG'))
    
    # Clica no meio da planilha para "ativar" a navegação dentro dela.
    bot.click(1000, 600)
    time.sleep(1)
    
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
    bot.PAUSE = 1
    print('--- Iniciando o processo de cópia.')
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
    
    # Navega até a proxima linha após a ultima chave.
    bot.press('DOWN')
    
    # Realiza uma avaliação, 
    # 1. Caso o campo esteja vazio, significa que ainda não foram inseridas novas notas, e para o processo. 
    # 2. Caso o campo esteja com uma chave XML nova, prossegue
    while True:
        bot.hotkey('ctrl', 'c')
        if 'Recuperando' in ahk.get_clipboard():
            print('--- Tentando copiar novamente.')
            time.sleep(0.1)
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
    time.sleep(1)
    bot.press('LEFT', presses= 4) # Navega até a coluna "RE"
    time.sleep(1)
    ahk.key_down('Shift')
    time.sleep(1)
    ahk.key_down('Control')
    time.sleep(1)
    ahk.key_press('down') # Com shift + ctrl pressionado, navega até a ultima linha da planilha
    time.sleep(1)
    ahk.key_press('right') # Avança para a ultima coluna
    time.sleep(1)
    ahk.key_press('right') # Avança para a ultima coluna
    time.sleep(1)
    ahk.key_press('right') # Avança para a ultima coluna
    time.sleep(1)
    ahk.key_up('Control') # Solta a tecla ctrl
    ahk.key_up('Shift')  # Solta a tecla Shift
    time.sleep(1)
    bot.hotkey('ctrl', 'c')
    time.sleep(1)
    dados_copiados = ahk.get_clipboard()
    cola_dados(dados_copiados)
    print('--- Dados copiados da planilha db_alltrips')


def cola_dados(dados_copiados = "TESTE"):
    print('--- Acessando a planilha de debug')
    abre_planilha_debug()
    time.sleep(0.5)
    bot.hotkey('CTRL', 'HOME') # Navega até a celula A1.
    bot.press('DOWN', presses= 2) # Proxima linha que deveria estar sem informação.
    bot.hotkey('ctrl', 'c') # Copia dados da coluna RE, para avaliar se realmente não tem nenhuma outra linha.
    if "" not in ahk.get_clipboard():
        exit(bot.alert('Ainda existem linhas para processar'))
    else:
        ahk.set_clipboard(dados_copiados) # Devolve os dados copiados para o clipboard.
        
    bot.press('ALT') # Abre o menu para navegação via teclas
    bot.press('C') # Vai até a opção "Inicio"
    bot.press('V') # Abre o menu de "Colar"
    bot.press('V') # Seleciona a opção "Colar somente valores"
    time.sleep(1)
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)

def main(ultimo_xml = chave_xml):
    bot.pause = 0.8
    abre_planilha_navegador()
    encontra_ultimo_xml(ultimo_xml = ultimo_xml)
    copia_dados()

if __name__ == '__main__':
    bot.PAUSE = 0.5
    #cola_dados()
    #abre_planilha_navegador()
    main(ultimo_xml= "35240814038100000111550010003443181975318640")
    #ahk.win_close('db_alltrips.xlsx', title_match_mode = 2)
