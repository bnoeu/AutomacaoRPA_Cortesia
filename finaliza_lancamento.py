from ast import Raise
import time
import pyautogui as bot

from datetime import date, timedelta
from abre_topcon import main as abre_topcon
from utils.configura_logger import get_logger
from alterar_localizar import alterar_localizar
from automacao.processo_transferencia import processo_transferencia
from utils.funcoes import marca_lancado, procura_imagem, ativar_janela, corrige_nometela


# --- Definição de parametros
from utils.funcoes import ahk as ahk
bot.LOG_SCREENSHOTS = True
bot.LOG_SCREENSHOTS_LIMIT = 5
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
logger = get_logger("finaliza_lancamento") # Obter logger configurado

def janelas_erro():
    telas_erro = ('Topsys', 'CsjTb')

    for tela in telas_erro:
        logger.debug(F'--- Tentando encontrar a tela: {tela}')
        if ahk.win_exists(tela, title_match_mode= 2):
            logger.error(F'--- Encontrou o pop-up de erro: "{tela}" necessario validar manualmente')
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec=True))
            marca_lancado(texto_marcacao = F'Erro_{tela}')
            return True


def janelas_sucesso():
    tela = "TopCompras (VM-CortesiaApli.CORTESIA.com)"

    logger.debug(F'--- Tentando encontrar a tela: {tela}')
    if ahk.win_exists(tela, title_match_mode= 1):
        logger.info(F'--- Encontrou uma tela de sucesso! Tela: {tela}')
        ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
        ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2, timeout= 30)
        return True


def verifica_popup_erro():
    # Identifica se algum pop_up de erro apareceu

    # Caso inconsistencia no local de estoque
    if procura_imagem(imagem='imagens/img_topcon/txt_existe_inconsistencia_local.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
        logger.info('--- Apresentou "Existe inconsistencia no local de estoque" ')
        marca_lancado(texto_marcacao='inconsistencia_local_estoque')
        return True

def finaliza_lancamento(planilha_marcada = False, lancamento_concluido = False, realizou_transferencia = False, tentativas_telas = 0, temp_inicial = ""):
    logger.info('--- Iniciando a função de finalização de lançamento, enviando PAGEDOWN ---' )
    ativar_janela('TopCompras')
    bot.press('pagedown')  # Conclui o lançamento

    logger.info('--- Tentando validar a tela que apresentou no sistema ---' )
    for i in range (0, 1):
        time.sleep(0.5)
        #* Verifica se apresentou alguma das telas de erro!
        if janelas_erro() is True:
            logger.info('--- Finalizou a task FINALIZA LANCAMENTO, pois apareceu uma tela de erro.' )
            return False 
        if janelas_sucesso() is True:
            break
        if i >= 29:
            raise Exception("Erro na função  FINALIZA LANCAMENTO: Não encontrou telas SUCESSO ou ERRO")

    while True:        
        # 0. Verifica se ocorreu algo de transferencia
        realizou_transferencia = processo_transferencia()
        time.sleep(0.25)
        # 1. Caso chave invalida.  
        if procura_imagem(imagem='imagens/img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            marca_lancado(texto_marcacao='Lancado_Manual')
            ahk.win_wait('TopCompras', title_match_mode = 2, timeout= 50)
            logger.info('--- Nota já lançada, marcando planilha!')
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            break

        if verifica_popup_erro() is True:
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
            time.sleep(0.6)
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            time.sleep(0.2)
            break

        if procura_imagem(imagem='imagens/img_topcon/txt_existe_diferenca_pedido.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            marca_lancado(texto_marcacao='Diferenca_ValorPedido')
            ahk.win_wait('TopCompras', title_match_mode = 2, timeout= 50)
            ahk.win_activate('TopCompras', title_match_mode = 2)
            logger.info('--- Existe diferença de valor entre o pedido e a nota fiscal!')
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            break

        # 2. Caso operação realizada.
        if procura_imagem(imagem='imagens/img_topcon/operacao_realizada.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is not False:
            if planilha_marcada is False:
                logger.info('--- Operação realizada, marcando a planilha com "Lancado RPA" ')
                marca_lancado(texto_marcacao='Lancado_RPA', temp_inicial = temp_inicial)
                planilha_marcada = True
            
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.click(procura_imagem(imagem='imagens/img_topcon/operacao_realizada.png'))
            logger.info('--- Clicando na tela "Operação Realizada" ')
            bot.press('ENTER')
            for i in range (0, 8):
                time.sleep(0.2)
                if procura_imagem(imagem='imagens/img_topcon/operacao_realizada.png', limite_tentativa= 1, continuar_exec= True) is False:
                    break

            #* Caso apareça a tela sobre o lançamento de CTE
            corrige_nometela("TopCompras (VM-CortesiaApli.CORTESIA.com)")
            ahk.win_activate("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode= 2)
            #* Caso apareça a tela sobre o lançamento de CTE
            if procura_imagem(imagem='imagens/img_topcon/txt_alerta_conhecimento.png', confianca= 0.75, limite_tentativa= 3, continuar_exec=True):
                bot.click(procura_imagem(imagem='imagens/img_topcon/bt_nao.png'))
                pass
                    
        elif planilha_marcada is True: # Essa parte só pode rodar, se encontrar a opção "operação realizada"
            logger.info('--- Não encontrou a tela "operação realizada", porém a planilha está marcada!')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            
            #Validando se já fecharam todas as telas.
            if procura_imagem(imagem='imagens/img_topcon/bt_obslancamento.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is not False:
                logger.info(F'--- Encontrou o botão "OBS. Lancamento." encerrando loop das telas, (valor do realizou transf: {realizou_transferencia})')
                if realizou_transferencia is True:
                    logger.info('--- Realizou transferencia, reabrindo o modulo do topcompras para evitar erros.')
                    time.sleep(0.4)
                    abre_topcon()
                else: # Retorna a tela para o modo localizar
                    logger.info(F'--- Não realizou transferencia! Pode continuar na mesma execução do TopCon! (valor do realizou transf: {realizou_transferencia})')
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    for i in range (0, 5):
                        if alterar_localizar():
                            lancamento_concluido = True
                            return True
                    
        # 3. Caso apareça "deseja imprimir o espelho da nota?"
        if procura_imagem(imagem='imagens/img_topcon/txt_espelhonota.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            logger.info('--- Apareceu a tela: "deseja imprimir o espelho da nota?" ')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('ENTER')
            ahk.win_activate('Espelho de Nota Fiscal', title_match_mode= 2)
            ahk.win_wait('Espelho de Nota Fiscal', title_match_mode= 2, timeout= 30)
        
        # 4. Caso apareça tela "Espelho da nota fiscal"
        while ahk.win_exists('Espelho de Nota Fiscal', title_match_mode= 2):
            ahk.win_close('Espelho de Nota Fiscal', title_match_mode= 2)

        # 6. Caso apareça a tela "Valor de frete maior"
        if procura_imagem(imagem='imagens/img_topcon/txt_valor_frete.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            logger.info('--- Nota já lançada, marcando planilha!')
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            marca_lancado(texto_marcacao='Erro_Frete')
            break
        
        # 7. Caso apareça a tela "Não é permitido informar contas a pagar"
        if procura_imagem(imagem='imagens/img_topcon/txt_contasapagar.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            logger.info('--- Erro de contas a pagar para essa operação!')
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            marca_lancado(texto_marcacao='Erro_OperacaoFiscal')
            break

        if lancamento_concluido is True:
            marca_lancado(texto_marcacao='Lancado_RPA')
            time.sleep(0.5)
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('F2') # Aperta F2 para retornar a tela para o modo "Localizar"
        
        # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        if tentativas_telas >= 20:
            logger.error(F'--- Excedeu o limite de tentativas de encontrar as telas, reabrindo o TopCompras, tentativa: {tentativas_telas}' )
            time.sleep(0.5)
            abre_topcon()
            return False # Retorna False pois o lançamento não foi concluido
        else:
            tempo_pausa_telas = tentativas_telas * 0.2
            time.sleep(tempo_pausa_telas)
            logger.debug(F'--- Não encontrou nenhuma das telas do processo finaliza lançamento, executando novamente, {tentativas_telas}, tempo pausa: {tempo_pausa_telas}')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            tentativas_telas += 1
        
        # 6. Caso apareça o erro de vencimento
        if procura_imagem(imagem='imagens/img_topcon/txt_vencimento.PNG', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
            logger.warning('--- Apareceu a tela de vencimento, alterando para +3 dias')
            bot.press('ENTER')
            ahk.win_wait_close('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
            time.sleep(0.5)
            # Altera a data de vencimento para +3 dias
            bot.click(procura_imagem(imagem='imagens/img_topcon/bt_contasapagar.PNG'))
            bot.click(procura_imagem(imagem='imagens/img_topcon/bt_datavencimento.PNG', area= (419, 536, 811, 715)))
            data_vencimento = date.today() + timedelta(3)
            data_vencimento = data_vencimento.strftime("%d%m%y")
            bot.write(data_vencimento)
            bot.press('ENTER')
            time.sleep(0.5)
            bot.press('pagedown')  # Conclui o lançamento


def main():
    finaliza_lancamento()

if __name__ == '__main__':        
    main()