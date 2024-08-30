# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import logging
# import cv2
import pytesseract
from ahk import AHK
import pyautogui as bot
from coleta_planilha import coleta_planilha
from utils.funcoes import marca_lancado, procura_imagem, corrige_nometela
from abre_topcon import abre_mercantil


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


# Realiza o processo de validação do lançamento.
def valida_lancamento():
    bot.PAUSE = 0.3
    while True:
        dados_planilha = coleta_planilha() # Recebe os dados coletados da planilha, já validados e formatados.
        logging.info('--- Iniciando a função: valida lancamento ---' )
        tentativa_alterar_botoes = 0
        chave_xml = dados_planilha[4].strip()

        # -------------------------------------- Lançamento Topcon --------------------------------------
        logging.info('--- Abrindo TopCompras para iniciar o lançamento')
        if ahk.win_exists('TopCompras', title_match_mode= 2) is False:
            corrige_nometela() # Realiza a correção do nome do modulo de compras
        
        logging.info('--- Alterando o TopCompras para o modo incluir')
        ahk.win_activate('TopCompras', title_match_mode= 2)
        ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 10)
        
        while True: # Enquanto a tela não for alterada para o modo incluir
            logging.info('--- Verificando se está no modo Localizar.')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            
            if procura_imagem(imagem='img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045), limite_tentativa= 2, confianca= 0.74) is False:
                logging.info(F'--- Não está no modo Localizar, enviando comando F2 para tentar entrar no modo, tentativa: {tentativa_alterar_botoes}')
                ahk.win_activate('TopCompras', title_match_mode= 2)
                bot.press('F2', presses= 2)
                
            if procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045), limite_tentativa= 2, confianca= 0.74):
                logging.info(F'--- Entrou no modo localizar, mudando para o modo incluir, tentativa: {tentativa_alterar_botoes}')
                ahk.win_activate('TopCompras', title_match_mode= 2)
                bot.press('F3', presses= 2)

                tentativa_alterar_botoes += 1
                if tentativa_alterar_botoes > 15:
                    logging.warning('--- Atingiu o maximo de tentativas de alterar os botões ---')
                    ahk.win_activate('TopCompras', title_match_mode= 2 )
                    bot.press('TAB')
                    # 3. Verificar se ficou o "1001 - Vila Prudente em azul"
                    if procura_imagem(imagem='img_topcon/txt_1001vila_prudente.png', continuar_exec= True):
                        # 4. Caso encontre, quebrar o loop e continuar a inserção.
                        logging.info('--- Está funcionando a inserção de NFE! não é necessario reabrir o TopCompras')
                        bot.press('F3')
                        break
                    else:
                        # Caso passe o limite de tentativas, provavelmente ocorreu algum problema.
                        time.sleep(1)
                        logging.warning('--- Excedeu o limite de tentativas de alteração para o modo localizar, reabrindo o TopCompras.')
                        abre_mercantil()
                        
            else:
                ahk.win_activate('TopCompras', title_match_mode=2)
                bot.press('F3', presses= 2)
                logging.info('--- Entrou no modo incluir, continuando inserção da NFE')
                break

        # Inicia inserção da chave XML
        bot.press('TAB', presses= 2, interval = 1)
        bot.write(chave_xml)
        bot.press('TAB')
        
        validou_xml = conferencia_xml(dados_planilha = dados_planilha) # Confere qual tela será apresentada. 
        if validou_xml is not None:
            logging.info(F'--- Validou o XML! Prosseguindo para a seleção do pedido {validou_xml}')
            return validou_xml

def conferencia_xml(tentativa = 0, maximo_tentativas = 50, texto_erro = False, dados_planilha = False):    
    logging.info('--- Iniciando a função: CONFERENCIA XML ---' )
    ahk.win_activate('TopCompras', title_match_mode=2)  
    
    while tentativa < maximo_tentativas: # Aguarda até aparecer uma das telas que podem ser exibidas nesse processo.
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
        
        if procura_imagem(imagem='img_topcon/botao_sim.jpg', continuar_exec= True, limite_tentativa= 1, confianca= 0.73) is not False:
            logging.info('--- XML Validado, indo para validação do pedido')
            ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
            bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg', continuar_exec=True))
            return dados_planilha
    
        if procura_imagem(imagem='img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Lancado_Manual"         
        
        if procura_imagem(imagem='img_topcon/naoencontrado_xml.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Aguardando_SEFAZ"
        
        if procura_imagem(imagem='img_topcon/chave_44digitos.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "Chave_invalida"
        
        if procura_imagem(imagem='img_topcon/nfe_cancelada.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.73) is not False:
            texto_erro = "NFE_Cancelada"
        
        if texto_erro is not False: # Caso tenha apresentado algum erro.
            logging.info(F'--- Apresentou um erro: {texto_erro} ' )
            while True:
                if procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True, limite_tentativa= 3, confianca= 0.75))
                    logging.info('--- Alterando a tela para o modo "localiza" para ficar correto o proximo lançamento.')
                    time.sleep(0.25)
                    bot.press('F2')
                else:
                    logging.info('--- Tela já está no modo localizar, saindo do loop!')
                    marca_lancado(texto_marcacao = texto_erro)
                    return
        else:       
            time.sleep(0.25)
            tentativa += 1
    else: # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        logging.info('--- Excedeu o limite de tentativas de encontrar alguma tela.')
        abre_mercantil()

if __name__ == '__main__':    
    #conferencia_xml()
    valida_lancamento()
