# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pyautogui as bot
from utils.funcoes import ahk as ahk
from abre_topcon import main as abre_topcon
from utils.configura_logger import get_logger
from utils.funcoes import marca_lancado, procura_imagem, corrige_nometela

# Definição de parametros
logger = get_logger("script1")


def conferencia_xml():    
    logger.info('--- Iniciando a função: CONFERENCIA XML ---' )
    texto_erro = False
    
    #* Aguarda a tela "TopCompras (VM-CortesiaApli.CORTESIA.com)" que é exibida quando ocorre algum retorno de informação do TopCon
    for i in range(0, 140):
        time.sleep(1)
        if ahk.win_is_active("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode= 2):
            logger.info('--- Encontrou o pop-up "TopCompras (VM-CortesiaApli.CORTESIA.co" ---' )
            break
        if ahk.win_is_active(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1):
            corrige_nometela("TopCompras (VM-CortesiaApli.CORTESIA.com)")
    else:
        logger.critical("Não encontrou o POP-UP, algum problema ocorreu")
        raise Exception("Não encontrou o POP-UP na coferencia do lançamento, algum problema ocorreu")

    for i in range (0, 30):
        if procura_imagem(imagem='imagens/img_topcon/botao_sim.jpg', continuar_exec= True, limite_tentativa= 1, confianca= 0.73) is not False:
            logger.info('--- XML Validado, indo para validação do pedido (Encontro o botão para vincular pedido "SIM" )')
            ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_sim.jpg', continuar_exec=True))
            return True # Retorna os dados, confirmando que essa chave XML é valida.

        elif procura_imagem(imagem='imagens/img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Lancado_Manual"

        elif procura_imagem(imagem='imagens/img_topcon/txt_fornecedor_cadastrado.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "cadastrar_fornecedor"
            bot.click(procura_imagem(imagem='imagens/img_topcon/bt_nao.png', continuar_exec= True))       
        
        elif procura_imagem(imagem='imagens/img_topcon/naoencontrado_xml.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Aguardando_SEFAZ"
        
        elif procura_imagem(imagem='imagens/img_topcon/chave_44digitos.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Chave_invalida"
        
        elif procura_imagem(imagem='imagens/img_topcon/nfe_cancelada.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "NFE_Cancelada"

        #* Caso encontre algum erro no lançamento da NFE
        if texto_erro is not False: # Caso tenha apresentado alguma tela de erro.
            logger.warning(F'--- Apresentou um erro: {texto_erro}' )
            marca_lancado(texto_marcacao = texto_erro)

            for i in range (0, 10):
                if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                    logger.info(F'--- Tela de erro fechada, alterando a tela para o modo "localiza" para ficar correto o proximo lançamento, tentativa: {i}')
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec= True, limite_tentativa= 3, confianca= 0.75))
                    time.sleep(0.4)
                    bot.press('F2')
                else:
                    logger.info('--- Tela já está no modo localizar, saindo do loop!')
                    return False
    else:
        logger.critical("Task conferencia XML quebrou por completo!")
        raise Exception("Task conferencia XML quebrou por completo!")