# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pyautogui as bot
from utils.funcoes import ahk as ahk
from abre_topcon import abre_topcon
from utils.configura_logger import get_logger
from utils.funcoes import marca_lancado, procura_imagem

# Definição de parametros
logger = get_logger("script1")


def conferencia_xml():    
    logger.info('--- Iniciando a função: CONFERENCIA XML ---' )
    tentativa = 0
    maximo_tentativas = 50
    texto_erro = False
    
    ahk.win_activate('TopCompras', title_match_mode=2)  
    while tentativa < maximo_tentativas: # Aguarda até aparecer uma das telas que podem ser exibidas nesse processo.
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
        time.sleep(0.25)
        
        if procura_imagem(imagem='imagens/img_topcon/txt_fornecedor_cadastrado.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "cadastrar_fornecedor"
            bot.click(procura_imagem(imagem='imagens/img_topcon/bt_nao.png', continuar_exec= True))
        
        elif procura_imagem(imagem='imagens/img_topcon/botao_sim.jpg', continuar_exec= True, limite_tentativa= 1, confianca= 0.73) is not False:
            logger.info('--- XML Validado, indo para validação do pedido')
            ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_sim.jpg', continuar_exec=True))
            return True # Retorna os dados, confirmando que essa chave XML é valida.
    
        elif procura_imagem(imagem='imagens/img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Lancado_Manual"         
        
        elif procura_imagem(imagem='imagens/img_topcon/naoencontrado_xml.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Aguardando_SEFAZ"
        
        elif procura_imagem(imagem='imagens/img_topcon/chave_44digitos.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Chave_invalida"
        
        elif procura_imagem(imagem='imagens/img_topcon/nfe_cancelada.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "NFE_Cancelada"

        if texto_erro is not False: # Caso tenha apresentado alguma tela de erro.
            tentativa_localiza = 0

            logger.info(F'--- Apresentou um erro: {texto_erro}' )
            marca_lancado(texto_marcacao = texto_erro)
            while tentativa_localiza < 6:
                if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                    logger.info(F'--- Tela de erro fechada, alterando a tela para o modo "localiza" para ficar correto o proximo lançamento, tentativa: {tentativa_localiza}')
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec= True, limite_tentativa= 3, confianca= 0.75))
                    time.sleep(0.25)
                    bot.press('F2')
                else:
                    logger.info('--- Tela já está no modo localizar, saindo do loop!')
                    return False
                tentativa_localiza += 1
                if tentativa_localiza >= 5:
                    logger.error('--- Não foi possivel alterar a tela para o modo LOCALIZAR!')
                    raise TimeoutError

        else:       
            time.sleep(0.25)
        tentativa += 1
    else: # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        logger.warning('--- Excedeu o limite de tentativas de encontrar alguma tela, reabrindo mercantil')
        abre_topcon()
