# -*- Criado por Bruno da Silva Santos. -*-
 # Para utilização na Cortesia Concreto.

import time

import asyncio
from utils.configura_logger import get_logger
import pyautogui as bot
from coleta_novos_dados import main as copia_banco
from datetime import datetime
from automacao_planilha.copia_linha_atual import copia_linha_atual
from automacao_planilha.valida_dados_coletados import valida_dados_coletados
from utils.funcoes import abre_planilha_navegador, reaplica_filtro_status, matar_autohotkey
from utils.funcoes import ahk as ahk


# --- Definição de parametros
logger = get_logger("script1")
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=jFhSykjw3Eg"

erro_log = ""

def coleta_dados():
    dados_copiados = False
    while not dados_copiados:
        dados_copiados = coleta_planilha()
    return dados_copiados

def coleta_planilha():
    bot.PAUSE = 0.6
    logger.info('--- Copiando dados e formatando na planilha de debug')
    tentativa = 0
    while tentativa < 20:
        ahk.win_activate("debug_db", title_match_mode = 2)
        time.sleep(0.2)
        bot.hotkey('CTRL', 'HOME')
        bot.press('DOWN')

        dados_planilha = copia_linha_atual()
        
        if tentativa > 19:
            raise Exception(F"Dados inválidos: {str(dados_planilha)}, executou todas as tentativas")
        if valida_dados(dados_planilha):
            return processa_dados(dados_planilha)
        else:
            logger.warning(F'--- Dados inválidos: {str(dados_planilha)}. Tentando novamente.')

def valida_dados(dados_planilha):
    return valida_dados_coletados(dados_planilha)

def processa_dados(dados_planilha):
    chave_xml = dados_planilha[4].strip()
    powerapps_id = dados_planilha[5]
    
    #* Confere se essa chave é a ultima, utilizando a situação da coluna "STATUS" como base
    if len(dados_planilha[6]) > 1:
        logger.info(F'--- Dados copiados: {dados_planilha}')
        logger.info(F'--- Chegou na última NFE {chave_xml}')
        #exit(bot.alert("Verificar copia"))
        copia_banco(chave_xml, powerapps_id)
        raise ValueError
    else:
            logger.info(F'--- Dados copiados com sucesso: {dados_planilha}')
    return dados_planilha

def handle_timeout(texto_erro):
    logger.exception("Os processos de coleta na PLANILHA apresentaram erro")
    #ahk.win_kill('Edge', title_match_mode=2, seconds_to_wait=3)
    time.sleep(1)


def formata_data_coletada(dados_copiados):
    data_copiada = dados_copiados.split(' ')
    data_copiada = data_copiada[0]
    
    # Converter para objeto datetime
    data_obj = datetime.strptime(data_copiada, "%d/%m/%y")

    # Converter para o formato desejado
    data_formatada = data_obj.strftime("%d/%m/%Y")
    return data_formatada

def calculo_tempo_final(tempo_inicial: float):
    # Linha específica onde você quer medir o tempo
    end_time = time.time()
    elapsed_time = end_time - tempo_inicial
    medicao_minutos = elapsed_time / 60
    print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")


def main():
    bot.PAUSE = 0.6
    ultimo_erro = ""
    erro_log = "" 

    for i in range(0, 3):
    
        logger.debug(F"--- Executando a tentativa {i} de executar o COLETA PLANILHA.py ")
        try:
            abre_planilha_navegador(planilha_debug)
            reaplica_filtro_status()

            dados_copiados = coleta_dados()
            if dados_copiados:
                logger.info('--- Processo de coleta da planilha foi executado corretamente.')
                return dados_copiados

        except ValueError:
            return False
        
        except Exception as e:
            ultimo_erro = e
            erro_log = e
            handle_timeout(texto_erro = ultimo_erro)
    else:
        logger.critical("--- Número maximo de tentativas de executar o COLETA PLANILHA.PY ")
        #subprocess.run([ ", "/im", "msedge.exe", "/f", "/t"], stderr=subprocess.DEVNULL)
        asyncio.run(matar_autohotkey(nome_exec= "msedge.exe"))

        raise Exception(F"Número maximo de tentativas de executar o COLETA PLANILHA.py, erro coletado: {erro_log}")


if __name__ == '__main__':
    tempo_inicial = time.time()
    main()

    #dados_copiados = main()
    #data_formatada = formata_data_coletada(dados_copiados[8])
    #coleta_planilha()


    #Calcula o tempo que levou para realizar a ação
    #calculo_tempo_final(tempo_inicial)