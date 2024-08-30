# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
import pytesseract
import pyautogui as bot

from ahk import AHK
from coleta_planilha import copia_linha_atual
from automacao_planilha.verifica_finalizou import verifica_finalizou_planilha
from utils.funcoes import marca_lancado

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.FAILSAFE = False


def coleta_planilha(dados_planilha = []):
        valida_final_planilha = 0
        if len(dados_planilha[6]) > 0 and (valida_final_planilha < 2):
            print('--- coluna status preenchida, validando se chegou ao final da planilha.')
            bot.hotkey('CTRL', 'HOME') # Navega para a celula A1 ( RE ), em seguida vai para a primeira linha com dados a serem copiado
            bot.press('down')
            
            dados_planilha = copia_linha_atual() # Executa uma nova copia para avaliar os dados
            verifica_finalizou_planilha(dados_planilha)
            
            '''
            def verifica_finalizou_planilha()
                if len(dados_planilha[6]) > 0: # Caso realmente esteja preenchido
                    logging.warning(F'--- Realmente está na ultima chave: {chave_xml}, executando COPIA BANCO')
                    print('--- final da planilha?:')
                    exit(bot.alert('Final?'))
                    time.sleep(1)
                    copia_banco(ultimo_xml= chave_xml)
                    time.sleep(1)
                    abre_planilha()
                    bot.press('F5')
                    time.sleep(8)
                    reaplica_filtro_status()
            '''
            
        elif len(chave_xml) < 42:
            marca_lancado('chave_invalida')
            
        elif (len(dados_planilha[0]) < 4) or (len(dados_planilha[0]) == 5):
            marca_lancado('RE_Invalido')
            
        elif 'Recuperando' in dados_planilha:
            logging.error('Não copiou corretamente os dados!')
        else:    
            logging.info(F'--- Dados copiados com sucesso: {dados_planilha}' )
            return dados_planilha

if __name__ == '__main__':
    coleta_planilha()
