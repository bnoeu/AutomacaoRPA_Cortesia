# Para utilização na Cortesia Concreto.
# -*- Criado por Bruno da Silva Santos. -*-

import os
import time
from datetime import datetime
import pyautogui as bot
from utils.configura_logger import get_logger
from utils.funcoes import procura_imagem, abre_planilha_navegador, msg_box

# --- Definição de parametros
from utils.funcoes import ahk as ahk
chave_xml = ""
logger = get_logger("script1")
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"


def encontra_ultimo_xml(ultimo_xml = ''):
    bot.PAUSE = 1.5
    while True:
        logger.info(F'--- Iniciando a navegação até a ultima chave XML: {ultimo_xml}')
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
        ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 1, timeout= 5)
        try:
            ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 1, timeout= 10)
        except TimeoutError:
            logger.warning('--- Planilha não encontrada!')
            return False

        # Navega até o campo da data, e organiza do menor para o maior.
        bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.
        bot.hotkey('CTRL', 'HOME') # Move a navegação até a celula A1
        logger.info('--- Navegou até a coluna/celula A1')
        bot.press('RIGHT', presses= 8, interval= 0.05) # Navega até o campo "D. Insercao"]
        logger.info('--- Navegou até a D. Inserção')
        bot.hotkey('ALT', 'DOWN') # Abre o menu do filtro
        time.sleep(5)
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
        time.sleep(0.5)
        bot.click(procura_imagem(imagem='imagens/img_planilha/icone_organiza_A_Z.png', continuar_exec= True)) # Clica no botão "organizar do mais antigo ao mais novo"
        logger.info('--- Organizou a planilha da forma "da menor para a maior" ')
        time.sleep(0.5)
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
        bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.

        #Abre o menu de pesquisa
        logger.info('--- Abrindo o menu de pesquisa na planilha')
        bot.press('ALT')
        bot.press('C')
        bot.press('F')
        bot.press('D')
        bot.press('F')

        # Insere a ultima chave copiada da planilha de debug
        logger.info(F'--- Digitando a ultima chave XML: {ultimo_xml}')
        bot.write(ultimo_xml)
        bot.press('ENTER', presses= 2)

        # Fecha o menu de pesquisa
        bot.press('ESC')
        bot.press('ALT', presses= 2)
        logger.info('--- Fechou o menu de pesquisa')

        # Verifica se realmente chegou no ultimo XML
        bot.hotkey('ctrl', 'c')
        if ahk.get_clipboard() == ultimo_xml:
            logger.info(F'--- Concluido a navegação até a ultima chave XML: {ultimo_xml}')
            return True
        else:
            logger.warning(F'--- Ops... não está na ultima chave {ultimo_xml}, navegando novamente.')
            raise TimeoutError

def copia_dados(ultimo_xml):
    bot.PAUSE = 1
    pausa_padrao = 600 # 10 Minutos
    tentativas = 0
    dados_copiados = ""

    logger.info('--- Iniciando o processo de cópia.')
    while tentativas < 4:
        hora_atual = datetime.now().time() # Obter o horário atual

        hora_inicio_pausa = datetime.strptime("02:00", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
        hora_final_pausa = datetime.strptime("04:00", "%H:%M").time() # Definir o horario final de referencia (04:00)
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
        bot.press('DOWN') # Navega até a proxima linha após a ultima chave.

        # 1. Caso o campo esteja vazio, significa que ainda não foram inseridas novas notas, e para o processo.
        # 2. Caso o campo esteja com uma chave XML nova, prossegue.
        while True: # Executa o processo de copia dos dados
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                logger.info('--- Tentando copiar novamente.')
                time.sleep(0.4)
            else:
                logger.info('--- Dado copiado com sucesso, realizando avaliação.')
                valor_copiado = ahk.get_clipboard()
                break

        if valor_copiado == "":
            logger.info(F'Verificando o horario atual: {hora_atual}')
            os.system('taskkill /im msedge.exe /f /t')
            if hora_atual > hora_inicio_pausa and hora_atual < hora_final_pausa: # Verificar se o horário atual é maior que 02:00
                logger.info(F"{hora_atual.strftime('%H:%M')} é maior que {hora_inicio_pausa}. pausando o script por 3 horas")
                msg_box(F"{hora_atual.strftime('%H:%M')} é maior que {hora_inicio_pausa}. pausando o script por 3 horas", tempo= 10800)
            else:
                msg_box(F"Campo vazio, aguardando {pausa_padrao / 60} minutos, tentativa: {tentativas}", tempo = pausa_padrao)
                logger.info(F"Campo vazio, aguardando {pausa_padrao / 60} minutos, tentativa {tentativas}")                
                if tentativas > 1:
                    pausa_padrao = pausa_padrao * 3
                    logger.warning(F'Tentou encontrar uma nova nota mais de {tentativas} vezes, aumentando tempo da pausa para: {pausa_padrao / 60}')
                tentativas += 1
        else:
            logger.info(F'--- Uma nova chave foi inserida: {valor_copiado}, saindo do loop')
            break
    else:
        logger.error(F'Tentou encontrar uma nova nota mais de {tentativas}')
        raise TimeoutError

    # Inicia o processo de seleção dos dados
    logger.info('--- Iniciando o processo de seleção dos dados')
    time.sleep(1)
    bot.press('LEFT', presses= 4) # Navega até a coluna "RE"
    time.sleep(1)
    ahk.key_down('Shift') # Segura a tecla SHIFT
    time.sleep(1)
    ahk.key_down('Control') # Segura a tecla CTRL
    time.sleep(1)
    ahk.key_press('down') # Com shift + ctrl pressionado, navega até a ultima linha da planilha
    time.sleep(1)
    ahk.key_press('right') # Avança para a ultima coluna
    tentativa_copia = 0
    logger.info('--- Pressionou SHIFT e CONTROL, indo até a ultima coluna preenchida')
    for i in range (0, 5):
        time.sleep(0.5)
        ahk.key_up('Shift')
        time.sleep(0.5)
        ahk.key_up('Control')
        time.sleep(0.5)
        bot.hotkey('ctrl', 'c')
        time.sleep(0.5)
        dados_copiados = ahk.get_clipboard()

        if ("/" in dados_copiados) or ("/2025" in dados_copiados) or ("," in dados_copiados): # Verifica se os dados foram copiados com sucesso
            logger.info('--- Novos dados copiados com sucesso da planilha db_alltrips')
            return dados_copiados
        else:
            ahk.key_down('Shift') # Segura a tecla SHIFT
            time.sleep(1)
            ahk.key_down('Control')
            time.sleep(1)
            ahk.key_press('right') # Avança para a ultima coluna
            time.sleep(1)

        tentativa_copia += 1
        if tentativa_copia > 10: # Verifica se excedeu o limite de tentativas de copiar os dados.
            logger.error('--- Excedeu o limite de tentativas de copiar os dados, soltando SHIFT e CONTROL')
            ahk.key_up('Shift')
            time.sleep(1)
            ahk.key_up('Control')
            raise Exception("Excedeu o limite de tentativas de copiar os dados, soltando SHIFT e CONTROL")

def cola_dados(dados_copiados = "TESTE"):
    abre_planilha_navegador(planilha_debug)
    bot.PAUSE = 2
    logger.info('--- Acessando a planilha de debug para COLAR os dados!')
    time.sleep(0.5)
    bot.hotkey('CTRL', 'HOME') # Navega até a celula A1.
    bot.press('DOWN', presses= 2) # Proxima linha que deveria estar sem informação.
    logger.info('--- Navegou até a proxima linha sem informações')

    bot.press('ALT') # Abre o menu para navegação via teclas
    bot.press('C') # Vai até a opção "Inicio"
    bot.press('V') # Abre o menu de "Colar"
    bot.press('V') # Seleciona a opção "Colar somente valores"
    time.sleep(0.5)
    logger.info('--- Copiado e colado com sucesso! Fechando a planilha original.')
    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)

    for i in range (0, 15):
        logger.info("--- Fechando a planilha do banco ORIGINAL antes de prosseguir.")
        time.sleep(0.4)
        ahk.win_kill('db_alltrips.xlsx', title_match_mode= 1) # Força o fechamento da planilha com o banco puro.
        if ahk.win_exists('db_alltrips.xlsx', title_match_mode= 1) is False:
            break
    else:
        logger.error('--- Não conseguiu fechar a planilha original')
        raise Exception("Não conseguiu fechar a planilha original")
    

def main(ultimo_xml = chave_xml):
    bot.PAUSE = 1.5
    dia_mes_atual = datetime.now() # Coleta a data atual, para validar se os dados são novos.
    quatro_dias_antes = dia_mes_atual.day - 4
    quatro_dias_antes = F"{quatro_dias_antes}/"

    #* Abre a planilha do db_alltrips (banco original)
    abre_planilha_navegador()
    encontra_ultimo_xml(ultimo_xml = ultimo_xml)

    dados_copiados = copia_dados(ultimo_xml)

    if quatro_dias_antes in dados_copiados:
        raise Exception(F'Dia "{quatro_dias_antes}" está nos dados copiados.')
    else:
        logger.info(F'Não encontrou: {quatro_dias_antes} nos dados copiados, os dados são novos!')

    ahk.key_up('Shift')
    ahk.key_up('Control')

    colou_dados = cola_dados(dados_copiados)
    if colou_dados is True:
        return True

if __name__ == '__main__':
    main(ultimo_xml= "35241233039223000979550010003836571296060514")
