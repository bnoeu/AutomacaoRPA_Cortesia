# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
from ahk import AHK
import pyautogui as bot
from copia_alltrips import main as copia_banco
from automacao_planilha.copia_linha_atual import copia_linha_atual
from automacao_planilha.valida_dados_coletados import valida_dados_coletados
from utils.funcoes import reaplica_filtro_status, abre_planilha_navegador

# --- Definição de parametros
ahk = AHK()
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"



def main():
    dados_copiados = False
    dados_validados = False
    while dados_validados is False:
        try:
            abre_planilha_navegador(planilha_debug) # Reabre DO ZERO a planilha
            while dados_copiados is False:
                dados_copiados = coleta_planilha()
            
            dados_validados = valida_dados_coletados(dados_copiados)
        except (TimeoutError, ValueError, OSError):
            logging.error("Os processos de coleta na PLANILHA deram erro, fechando planilha e executando novamente.")
            ahk.win_kill('Edge', title_match_mode= 2, seconds_to_wait= 3)
            logging.warning('Fechando o EDGE')
            time.sleep(5)          
    else:
        logging.info('--- Processo de coleta da planilha foi executado corretamente.')
        return dados_copiados

def coleta_planilha():
    while True:
        bot.PAUSE = 0.5
        
        logging.info('--- Copiando dados e formatando na planilha de debug')
        reaplica_filtro_status()
        bot.hotkey('CTRL', 'HOME') # Navega para a celula A1 ( RE ), em seguida vai para a primeira linha com dados a serem copiado
        bot.press('DOWN')
        logging.info('--- Navegou até a celula A1, e foi para a linha logo a baixo.')
        
        dados_planilha = copia_linha_atual()
        chave_xml = dados_planilha[4].strip()
        
        validou_dados = valida_dados_coletados(dados_planilha) # Valida se os campos foram copiados corretamente.   
        
        if validou_dados is True:    
            if len(dados_planilha[6]) > 1:
                logging.warning(F'--- Dados copiados: {dados_planilha}')
                logging.warning(F'--- Chegou na ultima NFE {chave_xml}')
                #exit(bot.alert(' Verificar o script do copia banco'))
                copia_banco(chave_xml)
            else:
                logging.info(F'--- Dados copiados com sucesso: {dados_planilha[4]}' )        
                return dados_planilha
        else:
            logging.info(F'--- Copiando novamente pois os dados são invalidos: {dados_planilha}')
            if "RE_Invalido" in dados_planilha[6]:
                logging.warning(F'--- Dados copiados: {dados_planilha}')
                logging.warning(F'--- Chegou na ultima NFE {chave_xml}')
                #exit(bot.alert(' Verificar o script do copia banco'))
                copia_banco(chave_xml)
            return False
            

if __name__ == '__main__':
    main()

