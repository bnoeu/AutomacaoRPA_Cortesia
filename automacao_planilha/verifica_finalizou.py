import time
import logging
import pyautogui as bot

from copia_alltrips import main as copia_banco
from automacao_planilha.abre_planilha_debug import abre_planilha
from utils.funcoes import reaplica_filtro_status

def verifica_finalizou_planilha(dados_planilha = [], chave_xml= ""):
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