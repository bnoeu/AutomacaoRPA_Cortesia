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
from abre_topcon import abre_topcon


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


# Realiza o processo de validação do lançamento.
def valida_lancamento():
    tentativa_alterar_botoes = 0
    while True:
        # Recebe os dados coletados da planilha, já validados e formatados.
        dados_planilha = coleta_planilha()
        chave_xml = dados_planilha[4].strip()

        # -------------------------------------- Lançamento Topcon --------------------------------------
        print(Fore.GREEN +  '--- Abrindo TopCompras para iniciar o lançamento' + Style.RESET_ALL)
        corrige_topcompras() # Realiza a correção do nome do modulo de compras
        
        print('--- Alterando o TopCompras para o modo incluir')
        ahk.win_activate('TopCompras', title_match_mode= 2 )
        ahk.win_wait_active('TopCompras', title_match_mode= 2)
        
        while True: # Enquanto a tela não for alterada para o modo incluir
            if procura_imagem(imagem='img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                ahk.win_activate('TopCompras', title_match_mode= 2)
                bot.press('F2', presses= 2)
                if procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
                    print(F'--- Entrou no modo localizar, mudando para o modo incluir, tentativa {tentativa_alterar_botoes}')
                    bot.press('F3', presses= 2)
                    time.sleep(0.2)

            else:
                ahk.win_activate('TopCompras', title_match_mode=2)
                bot.press('F3', presses= 2)
                print('--- Entrou no modo incluir, continuando inserção da NFE')
                time.sleep(0.5)
                break

            tentativa_alterar_botoes += 1
            if tentativa_alterar_botoes > 5:
                # Caso passe o limite de tentativas, provavelmente ocorreu algum problema.
                print('--- Excedeu o limite de tentativas de alteração, reabrindo o TopCompras.')
                exit()
                abre_topcon()
                valida_lancamento()

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
            os.system('cls')
            print(Fore.GREEN + F'--- Validou os dados do XML, dados_planilha: {dados_planilha}\n' + Style.RESET_ALL)
            return validou_xml

def conferencia_xml(tentativa = 0, maximo_tentativas = 20, texto_erro = False, dados_planilha = False):
    # Aguarda até aparecer uma das telas que podem ser exibidas nesse processo.
    while tentativa < maximo_tentativas:
        
        # Caso a tela não esteja respondendo.
        while ahk.win_exists('Não está respondendo', title_match_mode= 2):
            time.sleep(0.5)
        
        # Verifica quais das telas apareceu.
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)   
        if procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa= 2, continuar_exec=True) is not False:
            bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa=2, continuar_exec=True))
            print(Fore.GREEN + '--- XML Validado, indo para validação do pedido' + Style.RESET_ALL)
            return dados_planilha
        
        else: # Caso não encontre o botão "Sim", verifica se apareceu alguma das outras telas.

            if procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa= 2, continuar_exec=True) is not False:
                texto_erro = "Lancado_Manual"             
            elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', limite_tentativa= 2, continuar_exec=True) is not False:
                texto_erro = "Aguardando_SEFAZ"
            elif procura_imagem(imagem='img_topcon/chave_44digitos.png', limite_tentativa= 2, continuar_exec=True) is not False:
                texto_erro = "Chave_invalida"
            elif procura_imagem(imagem='img_topcon/nfe_cancelada.png', limite_tentativa= 2, continuar_exec=True) is not False:
                texto_erro = "NFE_Cancelada"
            else:
                print('--- Não encontrou nenhuma das telas, executando novamente.')
                
        if texto_erro is not False:
            print(Fore.RED + F'--- Apresentou um erro: {texto_erro} ' + Style.RESET_ALL)
            while procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
                print('--- Alterando a tela para o modo "localiza" para ficar correto o proximo lançamento.')
                time.sleep(0.5)
                bot.press('ENTER')
                bot.press('F2')
                
            marca_lancado(texto_marcacao = texto_erro)
            break
        
        tentativa += 1
    else: # Caso execute o maximo de tentativas sem sucesso.
        exit(bot.alert(Fore.RED + F'--- Rodou {maximo_tentativas} verificações e não achou nenhuma tela, verificar!' + Style.RESET_ALL))
        #TODO --- Validar se o topcon ainda está aberto, caso não esteja, reiniciar o processo do zero.

if __name__ == '__main__':    
    valida_lancamento()
