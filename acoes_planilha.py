# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
# import cv2
import pytesseract
from ahk import AHK
import pyautogui as bot
# import pygetwindow as gw
from colorama import Fore, Style
from coleta_planilha import coleta_planilha
from funcoes import marca_lancado, procura_imagem


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


#Realiza o processo de validação do lançamento.
def valida_lancamento():
    tentativa = 0
    maximo_tentativas = 15
    texto_erro = False

    while True:
        # Trata os dados coletados em "dados_planilha"
        while True:
            dados_planilha = coleta_planilha()
            chave_xml = dados_planilha[4].strip()
            if len(dados_planilha[6]) > 0:
                print(F'--- Aguardando 30s para aparecer novas notas, campo preenchido com: {dados_planilha[6]}, tamanho do status: {len(dados_planilha[6])}')
                time.sleep(30)
            if len(chave_xml) < 10:
                marca_lancado('Chave Invalida')
                #exit(bot.alert('chave_xml invalida'))
            elif (len(dados_planilha[0]) < 4) or (len(dados_planilha[0]) == 5):
                marca_lancado('RE_Invalido')
            else:
                break

        # -------------------------------------- Lançamento Topcon --------------------------------------
        print(Fore.GREEN +  '--- Abrindo TopCompras para iniciar o lançamento' + Style.RESET_ALL)
        tempo_inicio = time.time() #* Tempo inicial do lançamento
        
        try: 
            ahk.win_wait(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 3)
        except TimeoutError:
            print('--- Não achou a tela sem nome, verificando se o TopCompras abriu')
            if ahk.win_wait('TopCompras', title_match_mode= 1):
                print('--- TopCompras abriu com o nome normal, prosseguindo.')
            else:
                bot.alert(exit('TopCompras não encontrado.'))
        else:
            print('--- Encontrou  tela sem o nome, corrigindo')
            ahk.win_set_title(new_title= 'TopCompras', title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)

        print('--- Alterando para o modo alteração')
        ahk.win_activate('TopCompras', title_match_mode=2)
        while procura_imagem(imagem='img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('F2', presses= 2)
            bot.press('F3', presses= 2)
            time.sleep(0.1)
        else:
            ahk.win_activate('TopCompras', title_match_mode = 2)
            bot.press('F3', presses= 2)
            print('--- Entrou no modo incluir, continuando inserção da NFE')
            time.sleep(3)

        bot.press('TAB', presses= 2, interval = 0.5) # Inicia inserção da chave XML
        bot.write(chave_xml)
        bot.press('TAB')
        time.sleep(1)

        while tentativa < maximo_tentativas: #* Aguarda até aparecer uma das telas que podem ser exibidas nesse processo.             
            #Verifica quais das telas apareceu.
            if procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa=1, continuar_exec=True))
                tempo_coleta = time.time() - tempo_inicio
                tempo_coleta = tempo_coleta
                print(F'--- Tempo que levou: {tempo_coleta:0f} segundos')
                print(Fore.GREEN + '--- XML Validado, indo para validação do pedido' + Style.RESET_ALL)
                return dados_planilha
            else:
                if procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa= 1, continuar_exec=True) is not False:
                    texto_erro = "Lancado_Manual"              
                elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', limite_tentativa= 1, continuar_exec=True) is not False:
                    texto_erro = "Aguardando_SEFAZ"
                elif procura_imagem(imagem='img_topcon/chave_44digitos.png', limite_tentativa= 1, continuar_exec=True) is not False:
                    texto_erro = "Chave_invalida"
                elif procura_imagem(imagem='img_topcon/nfe_cancelada.png', limite_tentativa= 1, continuar_exec=True) is not False:
                    texto_erro = "NFE_Cancelada"
                else:
                    texto_erro = False
            
            if texto_erro is not False:
                print(Fore.RED + F'Apresentou um erro: {texto_erro} ' + Style.RESET_ALL)
                bot.press('ENTER')
                marca_lancado(texto_marcacao = texto_erro)
                break
            
            # Verifica caso tenha travado e espera até que o topcom volte a responder
            while ahk.win_exists('Não está respondendo', title_match_mode= 2):
                time.sleep(0.4)
            tentativa += 1
        else:
            exit(Fore.RED + '--- Rodou 10 verificações e não achou nenhuma tela, verificar!' + Style.RESET_ALL)
            #TODO --- Validar se o topcon ainda está aberto, caso não esteja, reiniciar o processo do zero.

if __name__ == '__main__':
    valida_lancamento()
