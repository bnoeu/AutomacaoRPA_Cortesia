# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg
# Debug db alltrips
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg
# db_alltrips no paulo, apenas leitura
#  

#*'''
#* A resolução da maquina precisa ser: 1920 x 1080, com a aproximação em "150%"
#*'''

import os
import time
import platform
import traceback
import pytesseract
import pyautogui as bot
from datetime import date, datetime, timedelta, timedelta
from utils.funcoes import ahk as ahk
from abre_topcon import main as abre_topcon
from utils.enviar_email import enviar_email
from utils.configura_logger import get_logger
from valida_pedido import main as valida_pedido
from valida_lancamento import valida_lancamento
from preenche_local import main as preenche_local
from finaliza_lancamento import finaliza_lancamento
from utils.funcoes import marca_lancado, procura_imagem, verifica_horario, ativar_janela

#* Definição de parametros
posicao_img = 0
continuar = True
qtd_notas_lancadas = 0
bot.LOG_SCREENSHOTS = True  
bot.LOG_SCREENSHOTS_LIMIT = 5
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

logger = get_logger("automacao", print_terminal= True) # Obter logger configurado




def formata_data_coletada(dados_copiados):
    data_copiada = dados_copiados.split(' ')[0]  # Pega apenas a parte da data
    print(F"Data copiada: {data_copiada}")
    # Converter a string para objeto datetime.date
    data_obj = datetime.strptime(data_copiada, "%d/%m/%Y").date()

    # Obtém a data de amanhã como objeto date
    amanha_data = coleta_proximo_dia()

    # Comparação correta entre objetos date
    if data_obj >= amanha_data:
        print("--- A data coletada é do próximo dia! Alterando para a data atual.")
        return date.today().strftime("%d/%m/%Y")  # Retorna a data atual formatada
    
    print("--- A data coletada é válida!")
    return data_obj.strftime("%d/%m/%Y")  # Retorna a data coletada formatada

def coleta_proximo_dia():
    # Retorna a data de amanhã como objeto date
    return date.today() + timedelta(days=1)

def programa_principal():
    dados_planilha = ("","", "", "", "", "", "", "", "01/02/2025 15:05")
    '''
    # String com a data e hora
    data_hora_str = dados_planilha[8]

    # Convertendo para um objeto datetime
    data_hora = datetime.strptime(data_hora_str, "%Y-%m-%d %H:%M:%S")

    # Formatando para o formato "dd/mm/YYYY"
    data_formatada = data_hora.strftime("%d/%m/%Y")

    print(data_formatada)
    '''

    
    data_formatada = formata_data_coletada(dados_planilha[8])

    exit(print(data_formatada))

    #* Alteração da data
    logger.info('--- Realizando validação/alteração da data')
    hoje = date.today()
    hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
    logger.info(F'--- Inserindo a data coletada: {data_formatada} e apertando ENTER')
    bot.write(data_formatada)
    bot.press('ENTER')
    ativar_janela('TopCompras', 70)

    # Caso o sistema informe que a data deve ser maior/igual a data inserida acima.
    logger.info('--- Verificando se apareceu data')
    if procura_imagem('imagens/img_topcon/data_invalida.png', continuar_exec= True):
        logger.warning('--- Precisa mudar a data, inserindo a data de hoje!')
        #bot.alert("Apresentou tela erro")
        ahk.win_close("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode= 2)
        time.sleep(0.5)        
        bot.write(hoje)
        bot.press('enter')
        # Aguarda até o topcompras voltar a funcionar
        ativar_janela('TopCompras', 70)

    try: # Aguarda a tela de erro do TopCon 
        ahk.win_wait('Topsys', title_match_mode= 2, timeout= 3)
    except TimeoutError:
        pass
    else:
        if ahk.win_exists('Topsys', title_match_mode= 2):
            ahk.win_activate('Topsys', title_match_mode= 2)
            logger.warning('--- Precisa mudar a data')
            bot.press('enter')          
            bot.write(hoje)
            bot.press('enter')
            time.sleep(0.4)


def main():
    chave_xml = "35250200787528000233570010000012781002513941"

    tipo_chave = (chave_xml[20:22])

    if "57" in tipo_chave: # Verifica se o tipo da chave é de um CT-E (Verificando o digitos 21º e 22º)
        print("Chave é de um CTE-E")


if __name__ == '__main__':
    data_copiada = "24/02/2025 21:07"

    data_obj = datetime.strptime(data_copiada, "%d/%m/%y").date()
    
    print(data_obj)

    #main()
