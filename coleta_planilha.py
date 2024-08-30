# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
import pytesseract
import pyautogui as bot

from ahk import AHK
from copia_alltrips import main as copia_banco
from utils.funcoes import marca_lancado, reaplica_filtro_status
from automacao_planilha.abre_planilha_debug import abre_planilha
from automacao_planilha.verifica_finalizou import verifica_finalizou_planilha

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.FAILSAFE = False

'''
def abre_planilha(): # Realiza a abertura da planilha de debug
    if ahk.win_exists('debug_db_alltrips', title_match_mode= 2):
        ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
        ahk.win_wait('debug_db_alltrips', title_match_mode= 2, timeout= 15)
        bot.click(990, 540)
'''


def copia_linha_atual():
    dados_planilha = []
    coluna_atual = 0
    while coluna_atual < 7: # Navega entre os 6 campos, realizando a copia um por um, e inserindo na lista Dados Planilha.
        while True:
            bot.hotkey('ctrl', 'c')
            time.sleep(0.2)
            if 'Recuperando' in ahk.get_clipboard():
                time.sleep(0.2)
            else:
                break
            
        if 'Recuperando' not in ahk.get_clipboard():
            dados_planilha.append(ahk.get_clipboard())
            bot.press('right')
            coluna_atual += 1
        else:
            logging.error('Copiou o texto "Recuperando Dados", tentando copiar novamente.')

    return dados_planilha


def coleta_planilha():
    valida_final_planilha = 0
    logging.info('--- Iniciando a função: coleta planilha ---' )
    
    while True:
        bot.PAUSE = 0.25    
        abre_planilha() # Abre a tela da planilha, que já deve ter sido acessada no Edge
        
        logging.info('--- Copiando dados e formatando na planilha de debug')
        ahk.set_clipboard("")
        bot.hotkey('CTRL', 'HOME') # Navega para a celula A1 ( RE ), em seguida vai para a primeira linha com dados a serem copiado
        bot.press('DOWN')
        logging.info('--- Navegou até a celula A1, e foi para a linha logo a baixo.')
        
        dados_planilha = copia_linha_atual()
        chave_xml = dados_planilha[4].strip() # Realiza a validação dos dados copiados
        
        #! Toda essa parte das validações precisa ser revisada.
        if len(dados_planilha[6]) > 0 and (valida_final_planilha < 2):
            print('--- coluna status preenchida, validando se chegou ao final da planilha.')
            abre_planilha() # Acessa novamente a planilha
            bot.hotkey('CTRL', 'HOME') # Navega para a celula A1 ( RE ), em seguida vai para a primeira linha com dados a serem copiado
            bot.press('down')
            
            dados_planilha = copia_linha_atual() # Executa uma nova copia para avaliar os dados
            verifica_finalizou_planilha(dados_planilha, chave_xml)
            
            if len(dados_planilha[6]) > 0: # Caso realmente esteja preenchido
                logging.warning(F'--- Realmente está na ultima chave: {chave_xml}, executando COPIA BANCO')
                #print('--- final da planilha?:')
                #exit(bot.alert('Final?'))
                time.sleep(1)
                copia_banco(ultimo_xml= chave_xml)
                time.sleep(1)
                abre_planilha()
                bot.press('F5')
                time.sleep(8)
                reaplica_filtro_status()
            
        elif len(chave_xml) < 42 and len(chave_xml) > 1: # Caso a chave tenha menos de 42 digitos, ela é invalida!
            marca_lancado('chave_invalida')
            
        elif (len(dados_planilha[0]) < 4) or (len(dados_planilha[0]) == 5):
            marca_lancado('RE_Invalido')
        
        # Ultima validação antes de realmente sair do loop
        if 'Recuperando' in dados_planilha:
            logging.error('Não copiou corretamente os dados, executando o loop novamente.')
        else:    
            logging.info(F'--- Dados copiados com sucesso: {dados_planilha}' )
            return dados_planilha

if __name__ == '__main__':
    bot.PAUSE = 1.5
    coleta_planilha()
