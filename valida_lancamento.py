# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
from ahk import AHK
import pyautogui as bot
from abre_topcon import abre_topcon
from automacao.conferencia_xml import conferencia_xml
from coleta_planilha import main as coleta_planilha
from utils.funcoes import procura_imagem, corrige_nometela, configurar_logging
from abre_topcon import abre_mercantil


# --- Definição de parametros
ahk = AHK()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''


# Realiza o processo de validação do lançamento.
def valida_lancamento():
    validou_xml = False
    bot.PAUSE = 0.6
    
    while validou_xml is False:
        dados_planilha = False
        tentativa_alterar_botoes = 0
        while dados_planilha is False:
            dados_planilha = coleta_planilha() # Recebe os dados coletados da planilha, já validados e formatados.

        logging.info(F'--- Iniciando VALIDA LANÇAMENTO - dados copiados: {dados_planilha}' )
        chave_xml = dados_planilha[4].strip()
        
        logging.info('--- Abrindo TopCompras para iniciar o lançamento')
        ahk.win_close('Vinculação Itens da Nota', title_match_mode = 2, seconds_to_wait= 3)
        if ahk.win_exists('TopCompras', title_match_mode= 2) is False: # Caso não encontre o TopCompras
            corrige_nometela() # Realiza a correção do nome do modulo de compras
        # Verificar se o Topcon & TopCompras estão abertos
        if ahk.win_exists('TopCon', title_match_mode=2) and ahk.win_exists('TopCompras', title_match_mode=2):
            logging.info('--- TopCompras e Topcon estão abertos, processo pode iniciar')
        else:
            logging.error('--- TopCompras e TopCon NÃO estão abertos, realizando abertura.')
            abre_topcon()
            abre_mercantil()
        
        logging.info('--- Alterando o TopCompras para o modo incluir')
        ahk.win_activate('TopCompras', title_match_mode= 2)
        ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 3)
        
        while True: # Enquanto a tela não for alterada para o modo incluir
            logging.info('--- Verificando se está no modo Localizar.')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            
            if procura_imagem(imagem='imagens/img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                logging.info(F'--- Não está no modo Localizar, enviando comando F2 para tentar entrar no modo, tentativa: {tentativa_alterar_botoes}')
                ahk.win_activate('TopCompras', title_match_mode= 2)
                bot.press('F2', presses= 2)
                
            if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
                logging.info(F'--- Entrou no modo localizar, mudando para o modo incluir, tentativa: {tentativa_alterar_botoes}')
                ahk.win_activate('TopCompras', title_match_mode= 2)
                bot.press('F3', presses= 2)

                tentativa_alterar_botoes += 1
                if tentativa_alterar_botoes > 15:
                    logging.warning('--- Atingiu o maximo de tentativas de alterar os botões ---')
                    ahk.win_activate('TopCompras', title_match_mode= 2 )
                    bot.press('TAB')
                    # 3. Verificar se ficou o "1001 - Vila Prudente em azul"
                    if procura_imagem(imagem='imagens/img_topcon/txt_1001vila_prudente.png', continuar_exec= True):
                        # 4. Caso encontre, quebrar o loop e continuar a inserção.
                        logging.info('--- Está funcionando a inserção de NFE! não é necessario reabrir o TopCompras')
                        bot.press('F3')
                        break
                    else:
                        # Caso passe o limite de tentativas, provavelmente ocorreu algum problema.
                        time.sleep(0.5)
                        logging.warning('--- Excedeu o limite de tentativas de alteração para o modo localizar, reabrindo o TopCompras.')
                        abre_mercantil()
                        
            else:
                logging.info('--- Entrou no modo incluir, continuando inserção da NFE')
                ahk.win_activate('TopCompras', title_match_mode=2)
                break

        # Inicia inserção da chave XML
        bot.press('TAB', presses= 2, interval = 1)
        bot.write(chave_xml)
        bot.press('TAB')
        
        validou_xml = conferencia_xml() # Confere qual tela será apresentada. 
        if validou_xml is not False:
            logging.info(F'--- Validou o XML! Prosseguindo para a seleção do pedido: {validou_xml}')
            return dados_planilha # Após todas as validações, retorna os dados para a execução principal


if __name__ == '__main__':
    configurar_logging(nome_arquivo="debug", nivel_log= logging.DEBUG)
    valida_lancamento()
