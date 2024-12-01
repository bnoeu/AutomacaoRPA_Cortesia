# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
from utils.configura_logger import get_logger
import pyautogui as bot
from copia_alltrips import main as copia_banco
from automacao_planilha.copia_linha_atual import copia_linha_atual
from automacao_planilha.valida_dados_coletados import valida_dados_coletados
from utils.funcoes import reaplica_filtro_status, abre_planilha_navegador, msg_box
from utils.funcoes import ahk as ahk


# --- Definição de parametros
logger = get_logger("script1")
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=jFhSykjw3Eg"

def main():
    while True:
        abre_planilha_navegador(planilha_debug)
        dados_copiados = coleta_dados()
        if dados_copiados:
            logger.success('--- Processo de coleta da planilha foi executado corretamente.')
            return dados_copiados

def coleta_dados():
    dados_copiados = False
    while not dados_copiados:
        dados_copiados = coleta_planilha()
    return dados_copiados

def coleta_planilha():
    logger.info('--- Copiando dados e formatando na planilha de debug')
    bot.PAUSE = 0.5
    while True:
        reaplica_filtro_status()
        bot.hotkey('CTRL', 'HOME')
        bot.press('DOWN')
        
        dados_planilha = copia_linha_atual()
        if valida_dados(dados_planilha):
            return processa_dados(dados_planilha)
        else:
            logger.warning(F'--- Dados inválidos: {str(dados_planilha)}. Tentando novamente.')

def valida_dados(dados_planilha):
    return valida_dados_coletados(dados_planilha)

def processa_dados(dados_planilha):
    chave_xml = dados_planilha[4].strip()
    if len(dados_planilha[6]) > 1:
        logger.info(F'--- Dados copiados: {dados_planilha}')
        logger.info(F'--- Chegou na última NFE {chave_xml}')
        copia_banco(chave_xml)
    else:
        logger.info(F'--- Dados copiados com sucesso: {dados_planilha[4]}')
    return dados_planilha

def handle_timeout(texto_erro):
    logger.exception("Os processos de coleta na PLANILHA deram erro, fechando planilha e executando novamente.")
    msg_box(str(f"{texto_erro}"), tempo=10)
    ahk.win_kill('Edge', title_match_mode=2, seconds_to_wait=3)
    logger.warning('Fechando o EDGE')
    time.sleep(5)

if __name__ == '__main__':
    main()
