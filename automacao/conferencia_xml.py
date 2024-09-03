# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
# import cv2
from ahk import AHK
import pyautogui as bot
from utils.funcoes import marca_lancado, procura_imagem
from abre_topcon import abre_mercantil

# --- Definição de parametros
ahk = AHK()


def conferencia_xml():    
    tentativa = 0
    maximo_tentativas = 50
    texto_erro = False
    
    
    logging.info('--- Iniciando a função: CONFERENCIA XML ---' )
    ahk.win_activate('TopCompras', title_match_mode=2)  
    while tentativa < maximo_tentativas: # Aguarda até aparecer uma das telas que podem ser exibidas nesse processo.
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
        time.sleep(0.5)
        
        if procura_imagem(imagem='imagens/img_topcon/botao_sim.jpg', continuar_exec= True, limite_tentativa= 1, confianca= 0.73) is not False:
            logging.info('--- XML Validado, indo para validação do pedido')
            ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_sim.jpg', continuar_exec=True))
            return True # Retorna os dados, confirmando que essa chave XML é valida.
    
        if procura_imagem(imagem='imagens/img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Lancado_Manual"         
        
        if procura_imagem(imagem='imagens/img_topcon/naoencontrado_xml.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Aguardando_SEFAZ"
        
        if procura_imagem(imagem='imagens/img_topcon/chave_44digitos.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Chave_invalida"
        
        if procura_imagem(imagem='imagens/img_topcon/nfe_cancelada.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "NFE_Cancelada"
        
        if texto_erro is not False: # Caso tenha apresentado algum erro.
            logging.info(F'--- Apresentou um erro: {texto_erro} ' )
            while True:
                if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec= True, limite_tentativa= 3, confianca= 0.75))
                    logging.info('--- Alterando a tela para o modo "localiza" para ficar correto o proximo lançamento.')
                    time.sleep(0.25)
                    bot.press('F2')
                else:
                    logging.info('--- Tela já está no modo localizar, saindo do loop!')
                    marca_lancado(texto_marcacao = texto_erro)
                    return False
        else:       
            time.sleep(0.25)
            tentativa += 1
    else: # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        logging.warning('--- Excedeu o limite de tentativas de encontrar alguma tela, reabrindo mercantil')
        abre_mercantil()
