# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
# import cv2
import pytesseract
import os
from ahk import AHK
import pyautogui as bot
# import pygetwindow as gw
from colorama import Fore, Style
from coleta_planilha import coleta_planilha
from funcoes import marca_lancado, procura_imagem, corrige_topcompras
from abre_topcon import abre_topcon, abre_mercantil


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


# Realiza o processo de validação do lançamento.
def valida_lancamento():
    bot.PAUSE = 0.2
    while True:
        tentativa_alterar_botoes = 0
        # Recebe os dados coletados da planilha, já validados e formatados.
        dados_planilha = coleta_planilha()
        chave_xml = dados_planilha[4].strip()

        # -------------------------------------- Lançamento Topcon --------------------------------------
        print(Fore.GREEN +  '--- Abrindo TopCompras para iniciar o lançamento' + Style.RESET_ALL)
        corrige_topcompras() # Realiza a correção do nome do modulo de compras
        
        print('--- Alterando o TopCompras para o modo incluir')
        ahk.win_activate('TopCompras', title_match_mode= 2 )
        ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 15)
        
        while True: # Enquanto a tela não for alterada para o modo incluir
            ahk.win_activate('TopCompras', title_match_mode= 2)
            
            print('--- Verificando se está no modo Localizar.')
            if procura_imagem(imagem='img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045), limite_tentativa= 2, confianca= 0.74) is False:
                print(F'--- Não está no modo Localizar, enviando comando F2 para tentar entrar no modo, tentativa: {tentativa_alterar_botoes}')
                ahk.win_activate('TopCompras', title_match_mode= 2)
                bot.press('F2', presses= 2)
                
            if procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045), limite_tentativa= 2, confianca= 0.74):
                print(F'--- Entrou no modo localizar, mudando para o modo incluir, tentativa: {tentativa_alterar_botoes}')
                ahk.win_activate('TopCompras', title_match_mode= 2)
                bot.press('F3', presses= 2)

                tentativa_alterar_botoes += 1
                if tentativa_alterar_botoes > 10:
                    print('--- Atingiu o maximo de tentativas de alterar os botões ---')
                    # 1. Abrir topCompras
                    ahk.win_activate('TopCompras', title_match_mode= 2 )
                    # 2. Apertar TAB
                    bot.press('TAB')
                    # 3. Verificar se ficou o "1001 - Vila Prudente em azul"
                    if procura_imagem(imagem='img_topcon/txt_1001vila_prudente.png', continuar_exec= True):
                        # 4. Caso encontre, quebrar o loop e continuar a inserção.
                        print('--- Está funcionando a inserção de NFE! não é necessario reabrir o TopCompras')
                        bot.press('F3')
                        break
                    else:
                        # Caso passe o limite de tentativas, provavelmente ocorreu algum problema.
                        time.sleep(0.5)
                        print('--- Excedeu o limite de tentativas de alteração para o modo localizar, reabrindo o TopCompras.')
                        abre_mercantil()
                        #abre_topcon()
                        #exit(bot.alert('Verificar script'))
                        #abre_topcon()
                        

            else:
                ahk.win_activate('TopCompras', title_match_mode=2)
                bot.press('F3', presses= 2)
                print('--- Entrou no modo incluir, continuando inserção da NFE')
                break

        # Inicia inserção da chave XML
        bot.press('TAB', presses= 2, interval = 1)
        bot.write(chave_xml)
        bot.press('TAB')
        time.sleep(0.5)
        validou_xml = conferencia_xml(dados_planilha = dados_planilha)
        if validou_xml is None:
            print('--- Retornou os valores como None, mantendo o loop')
            pass
        else:
            #os.system('cls')
            print(Fore.GREEN + F'--- Validou os dados do XML, dados_planilha: {dados_planilha}' + Style.RESET_ALL)
            return validou_xml

def conferencia_xml(tentativa = 0, maximo_tentativas = 25, texto_erro = False, dados_planilha = False):
    # Aguarda até aparecer uma das telas que podem ser exibidas nesse processo.
    tentativas_telas = 0
    ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
    
    while tentativa < maximo_tentativas:
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True) 
        if procura_imagem(imagem='img_topcon/botao_sim.jpg', continuar_exec= True, limite_tentativa= 2, confianca= 0.74) is not False:
            bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg', continuar_exec=True))
            print(Fore.GREEN + '--- XML Validado, indo para validação do pedido' + Style.RESET_ALL)
            return dados_planilha
        
        else: # Caso não encontre o botão "Sim", verifica se apareceu alguma das outras telas.
            while True:
                if procura_imagem(imagem='img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 2, confianca= 0.74) is not False:
                    texto_erro = "Lancado_Manual"
                    break          
                elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', continuar_exec=True, limite_tentativa= 2, confianca= 0.74) is not False:
                    texto_erro = "Aguardando_SEFAZ"
                    break
                elif procura_imagem(imagem='img_topcon/chave_44digitos.png', continuar_exec=True, limite_tentativa= 2, confianca= 0.74) is not False:
                    texto_erro = "Chave_invalida"
                    break
                elif procura_imagem(imagem='img_topcon/nfe_cancelada.png', continuar_exec=True, limite_tentativa= 2, confianca= 0.74) is not False:
                    texto_erro = "NFE_Cancelada"
                    break
                else:
                    print('--- Não encontrou nenhuma das telas, executando novamente.')
                    break
                
        if texto_erro is not False:
            print(Fore.RED + F'--- Apresentou um erro: {texto_erro} ' + Style.RESET_ALL)
            ahk.win_activate('TopCompras', title_match_mode= 2)
            while True:
                if procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True, limite_tentativa= 3, confianca= 0.75))
                    print('--- Alterando a tela para o modo "localiza" para ficar correto o proximo lançamento.')
                    time.sleep(0.5)
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.press('F2')
                else:
                    print('--- Tela já está no modo localizar, saindo do loop!')
                    break
                
            marca_lancado(texto_marcacao = texto_erro)
            break
        
        tentativa += 1
    else: # Caso execute o maximo de tentativas sem sucesso.
        # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        print('--- Excedeu o limite de tentativas de encontrar alguma tela.')
        abre_mercantil()
        #exit(bot.alert(Fore.RED + F'--- Rodou {maximo_tentativas} verificações e não achou nenhuma tela, verificar!' + Style.RESET_ALL))
        #TODO --- Validar se o topcon ainda está aberto, caso não esteja, reiniciar o processo do zero.

if __name__ == '__main__':    
    conferencia_xml()
    #valida_lancamento()
