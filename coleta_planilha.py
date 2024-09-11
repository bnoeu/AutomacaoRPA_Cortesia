# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
import pytesseract
from ahk import AHK
import pyautogui as bot
from copia_alltrips import main as copia_banco
from automacao_planilha.copia_linha_atual import copia_linha_atual
from automacao_planilha.abre_planilha_debug import abre_planilha
from automacao_planilha.verifica_finalizou import verifica_finalizou_planilha
from automacao_planilha.valida_dados_coletados import valida_dados_coletados
from utils.funcoes import marca_lancado, reaplica_filtro_status, abre_planilha_navegador

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.FAILSAFE = False
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"


''' #! Modularizado devido ao uso em outras partes.
def copia_linha_atual():
    bot.PAUSE = 0.1    
    dados_planilha = []
    coluna_atual = 0
    while coluna_atual < 7: # Navega entre os 6 campos, realizando a copia um por um, e inserindo na lista Dados Planilha.
        while True:
            bot.hotkey('ctrl', 'c')
            time.sleep(0.1)
            if 'Recuperando' in ahk.get_clipboard():
                time.sleep(0.1)
            else:
                break

        if 'Recuperando' not in ahk.get_clipboard():
            dados_planilha.append(ahk.get_clipboard())
            bot.press('right')
            coluna_atual += 1
        else:
            print(F'--- Copiou o texto "Recuperando Dados", tentando copiar novamente {dados_planilha}')
    return dados_planilha
'''

def coleta_planilha():
    logging.info('--- Iniciando a função: coleta planilha ---' )
    
    while True:
        bot.PAUSE = 0.2
        ahk.set_clipboard("")
        
        logging.info('--- Copiando dados e formatando na planilha de debug')
        abre_planilha_navegador(planilha_debug) # Reabre DO ZERO a planilha
        reaplica_filtro_status()
        bot.hotkey('CTRL', 'HOME') # Navega para a celula A1 ( RE ), em seguida vai para a primeira linha com dados a serem copiado
        bot.press('DOWN')
        logging.info('--- Navegou até a celula A1, e foi para a linha logo a baixo.')
        
        dados_planilha = copia_linha_atual()
        chave_xml = dados_planilha[4].strip()
        
        validou_dados = valida_dados_coletados(dados_planilha) # Valida se os campos foram copiados corretamente.   
            
        if validou_dados is True:    
            if len(dados_planilha[6]) > 1:
                logging.warning(F'--- Chegou na ultima NFE {chave_xml}')
                copia_banco(chave_xml)
            else:
                logging.info(F'--- Dados copiados com sucesso: {dados_planilha[4]}' )        
                return dados_planilha
        else:
            print(F'--- Copiando novamente pois os dados são invalidos: {dados_planilha}')
            

if __name__ == '__main__':
    dados_validados = False
    while dados_validados is False:
        dados_copiados = coleta_planilha()
        dados_validados = valida_dados_coletados(dados_copiados)
        print(dados_copiados)
