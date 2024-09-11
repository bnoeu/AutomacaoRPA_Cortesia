# Para utilização na Cortesia Concreto.
# -*- Criado por Bruno da Silva Santos. -*-

import os
import time
import logging
from ahk import AHK
import pyautogui as bot
from utils.funcoes import procura_imagem, abre_planilha_navegador
from automacao_planilha.abre_planilha_debug import abre_planilha

# --- Definição de parametros
ahk = AHK()
chave_xml = ""
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"


'''
# Receber nos parametros a chave_xml da ultima NFE coletada
def abre_planilha_navegador(link_planilha = alltrips):
    while ahk.win_exists('db_alltrips.xlsx', title_match_mode= 1): # Garante que a planilha não esteja aberta
        ahk.win_close('db_alltrips.xlsx', title_match_mode= 1)
        time.sleep(0.5)
    
    logging.info('--- Abrindo a planilha no EDGE.')
    comando_iniciar = F'start msedge {link_planilha} -new-window -inprivate'
    os.system(comando_iniciar)
    time.sleep(10)
    ahk.win_wait_active('db_alltrips.xlsx', title_match_mode = 2, timeout= 15)
    ahk.win_maximize('db_alltrips.xlsx')
    logging.info('--- Planilha aberta e maximizada.')
'''

def encontra_ultimo_xml(ultimo_xml = ''):
    bot.PAUSE = 1
    while True:
        logging.info(F'--- Iniciando a navegação até a ultima chave XML: {ultimo_xml}')
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
        ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 2, timeout= 5)
        try:
            ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 2, timeout= 10)
        except TimeoutError:
            logging.critical('--- Planilha não encontrada!')
            return False
        
        # Navega até o campo da data, e organiza do menor para o maior.
        bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.
        bot.hotkey('CTRL', 'HOME') # Move a navegação até a celula A1
        bot.press('RIGHT', presses= 8, interval= 0.05) # Navega até o campo "D. Insercao"]
        bot.hotkey('ALT', 'DOWN') # Abre o menu do filtro
        time.sleep(5)
        bot.click(procura_imagem(imagem='imagens/img_planilha/txt_menor_maior.png')) # Clica no botão "organizar do menor ao maior"
        time.sleep(0.5)
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
        bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.
        
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
        
        # Verifica se realmente chegou no ultimo XML
        bot.hotkey('ctrl', 'c')
        if ahk.get_clipboard() == ultimo_xml:
            logging.info(F'--- Concluido a navegação até a ultima chave XML: {ultimo_xml}')
            return True
        else:
            logging.error(F'--- Ops... não está na ultima chave {ultimo_xml}, navegando novamente.')
    
def copia_dados(ultimo_xml):
    bot.PAUSE = 0.25
    dados_copiados = ""
    logging.info('--- Iniciando o processo de cópia.')
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 2)
    
    bot.press('DOWN') # Navega até a proxima linha após a ultima chave.
    
    # Realiza uma avaliação
    # 1. Caso o campo esteja vazio, significa que ainda não foram inseridas novas notas, e para o processo. 
    # 2. Caso o campo esteja com uma chave XML nova, prossegue
    while True:
        bot.hotkey('ctrl', 'c')
        if 'Recuperando' in ahk.get_clipboard():
            logging.info('--- Tentando copiar novamente.')
            time.sleep(0.2)
        else:
            logging.info('--- Dado copiado com sucesso, realizando avaliação.')
            valor_copiado = ahk.get_clipboard()
            break
    
    if valor_copiado == "":
        logging.info('--- Campo vazio, aguardando 10 minutos.')
        time.sleep(100)
    else:
        if valor_copiado == ultimo_xml:
            print('--- Erro, está na mesma chave!')
        else:
            logging.info(F'--- Uma nova chave foi inserida: {valor_copiado}')
    
    # Inicia o processo de seleção dos dados
    logging.info('--- Iniciando o processo de seleção dos dados')    
    time.sleep(0.5)
    bot.press('LEFT', presses= 4) # Navega até a coluna "RE"
    time.sleep(0.5)
    ahk.key_down('Shift') # Segura a tecla SHIFT
    time.sleep(0.5)
    ahk.key_down('Control') # Segura a tecla CTRL
    time.sleep(0.5)
    ahk.key_press('down') # Com shift + ctrl pressionado, navega até a ultima linha da planilha
    time.sleep(0.5)
    ahk.key_press('right') # Avança para a ultima coluna
    tentativa_copia = 0
    while True:
        ahk.key_up('Shift')
        ahk.key_up('Control')
        
        time.sleep(0.5)
        bot.hotkey('ctrl', 'c')
        time.sleep(0.5)
        dados_copiados = ahk.get_clipboard()
        if "," in dados_copiados:
            logging.info('--- Novos dados copiados com sucesso da planilha db_alltrips')
            return dados_copiados
        else:
            ahk.key_down('Shift') # Segura a tecla SHIFT
            time.sleep(0.5)
            ahk.key_down('Control')
            time.sleep(0.5)
            ahk.key_press('right') # Avança para a ultima coluna
            time.sleep(0.5)
        
        tentativa_copia += 1
        if tentativa_copia > 10:
            abre_planilha_navegador()


def cola_dados(dados_copiados = "TESTE"):
    logging.info('--- Acessando a planilha de debug')
    abre_planilha_navegador(planilha_debug)
    time.sleep(0.5)
    bot.hotkey('CTRL', 'HOME') # Navega até a celula A1.
    bot.press('DOWN', presses= 2) # Proxima linha que deveria estar sem informação.
        
    bot.press('ALT') # Abre o menu para navegação via teclas
    bot.press('C') # Vai até a opção "Inicio"
    bot.press('V') # Abre o menu de "Colar"
    bot.press('V') # Seleciona a opção "Colar somente valores"
    time.sleep(0.5)
    logging.info('--- Copiado e colado com sucesso! Fechando a planilha original.')
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
    while ahk.win_exists('db_alltrips.xlsx', title_match_mode= 1):
        time.sleep(0.25)
        ahk.win_kill('db_alltrips.xlsx', title_match_mode= 1) # Força o fechamento da planilha com o banco puro.
    else:
        return True

def main(ultimo_xml = chave_xml):
    bot.PAUSE = 1
    abre_planilha_navegador()
    encontra_ultimo_xml(ultimo_xml = ultimo_xml)
    dados_copiados = copia_dados(ultimo_xml)
    colou_dados = cola_dados(dados_copiados)
    if colou_dados is True:
        return True

if __name__ == '__main__':
    main(ultimo_xml= "35240956899602000268550010007998511676004194")