import time
import pyautogui as bot
import logging

from ahk import AHK
from colorama import Style, Fore
from datetime import date, timedelta
#from Materia_Prima import programa_principal
from automacao.processo_transferencia import processo_transferencia
from abre_topcon import abre_mercantil
from utils.funcoes import marca_lancado, procura_imagem

# --- Definição de parametros
ahk = AHK()
bot.LOG_SCREENSHOTS = True
bot.LOG_SCREENSHOTS_LIMIT = 5
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''


def finaliza_lancamento(planilha_marcada = False, lancamento_concluido = False, realizou_transferencia = False, tentativas_telas = 0):
    logging.info('--- Iniciando a função de finalização de lançamento, enviando PAGEDOWN ---' )
    ahk.win_activate('TopCompras', title_match_mode=2)
    bot.press('pagedown')  # Conclui o lançamento
    
    while True:
        ahk.win_activate('TopCompras', title_match_mode=2) # Para manter o TopCompras aberto.
        ahk.win_wait('TopCompras', title_match_mode = 2, timeout= 50)
        
        #! Mover para a função do tratamento de erros
        if ahk.win_exists('CsjTb', title_match_mode= 2): # Caso apareça a tela de campo obrigatorio (Aparece quando não preencher nenhum campo.)
            ahk.win_close('CsjTb', title_match_mode= 2)
            abre_mercantil()
            logging.error('--- Reabriu o mercantil, recomeçando o processo.')
            #programa_principal()
        
        # 0. Verifica se ocorreu algo de transferencia
        realizou_transferencia = processo_transferencia()
        time.sleep(1)
        # 1. Caso chave invalida.  
        if procura_imagem(imagem='imagens/img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            logging.info('--- Nota já lançada, marcando planilha!')
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            marca_lancado(texto_marcacao='Lancado_Manual')
            break

        # 2. Caso operação realizada.
        if procura_imagem(imagem='imagens/img_topcon/operacao_realizada.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is not False:
            if planilha_marcada is False:
                logging.info('--- Operação realizada, marcando a planilha com "Lancado RPA" ')
                marca_lancado(texto_marcacao='Lancado_RPA')
                planilha_marcada = True
            
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.click(procura_imagem(imagem='imagens/img_topcon/operacao_realizada.png'))
            logging.info('--- Clicando na tela "Operação Realizada" ')
            bot.press('ENTER')
                    
        elif planilha_marcada is True: # Essa parte só pode rodar, se encontrar a opção "operação realizada"
            logging.info('--- Não encontrou a tela "operação realizada", porém a planilha está marcada!')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            
            #Validando se já fecharam todas as telas.
            if procura_imagem(imagem='imagens/img_topcon/bt_obslancamento.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is not False:
                logging.info(F'--- Encontrou o botão "OBS. Lancamento." encerrando loop das telas, valor do realizou transf: {realizou_transferencia}')
                if realizou_transferencia is True:
                    logging.info('--- Realizou transferencia, reabrindo o modulo do topcompras para evitar erros.')
                    time.sleep(0.25)
                    abre_mercantil()
                else: # Retorna a tela para o modo localizar
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.press('F3', presses = 1)
                    bot.press('F2', presses = 1)
                    time.sleep(0.25)
                    
                    if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
                        logging.info('--- Entrou no modo localizar, lançamento realmente concluido!\n')
                        lancamento_concluido = True
                        return True
                    
        # 3. Caso apareça "deseja imprimir o espelho da nota?"
        if procura_imagem(imagem='imagens/img_topcon/txt_espelhonota.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            logging.info('--- Apareceu a tela: deseja imprimir o espelho da nota?')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('ENTER')
            ahk.win_activate('Espelho de Nota Fiscal', title_match_mode= 2)
            ahk.win_wait('Espelho de Nota Fiscal', title_match_mode= 2, timeout= 30)
        
        # 4. Caso apareça tela "Espelho da nota fiscal"
        while ahk.win_exists('Espelho de Nota Fiscal', title_match_mode= 2):
            ahk.win_close('Espelho de Nota Fiscal', title_match_mode= 2)

        if lancamento_concluido is True:
            time.sleep(1)
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('F2') # Aperta F2 para retornar a tela para o modo "Localizar"
            marca_lancado(texto_marcacao='Lancado_RPA')
        
        # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        if tentativas_telas >= 20:
            logging.error(F'--- Excedeu o limite de tentativas de encontrar as telas, reabrindo o TopCompras, tentativa: {tentativas_telas}' )
            time.sleep(1)
            abre_mercantil()
            return False # Retorna False pois o lançamento não foi concluido
        else:
            tempo_pausa_telas = tentativas_telas * 0.2
            time.sleep(tempo_pausa_telas)
            logging.debug(F'--- Não encontrou nenhuma das telas do processo finaliza lançamento, executando novamente, {tentativas_telas}, tempo pausa: {tempo_pausa_telas}')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            tentativas_telas += 1
        
        # 6. Caso apareça o erro de vencimento
        if procura_imagem(imagem='imagens/img_topcon/txt_vencimento.PNG', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
            logging.warning('--- Apareceu a tela de vencimento, alterando para +3 dias')
            bot.press('ENTER')
            ahk.win_wait_close('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
            time.sleep(1)
            # Altera a data de vencimento para +3 dias
            bot.click(procura_imagem(imagem='imagens/img_topcon/bt_contasapagar.PNG'))
            bot.click(procura_imagem(imagem='imagens/img_topcon/bt_datavencimento.PNG', area= (419, 536, 811, 715)))
            data_vencimento = date.today() + timedelta(3)
            data_vencimento = data_vencimento.strftime("%d%m%y")
            bot.write(data_vencimento)
            bot.press('ENTER')
            time.sleep(1)
            bot.press('pagedown')  # Conclui o lançamento
