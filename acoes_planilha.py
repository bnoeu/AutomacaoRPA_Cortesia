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
bot.PAUSE = 0.5
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


#Realiza o processo de validação do lançamento.
def valida_lancamento():
    while True:
        # Trata os dados coletados em "dados_planilha"
        while True:
            dados_planilha = coleta_planilha()
            chave_xml = dados_planilha[4].strip()
            if len(chave_xml) < 10:
                marca_lancado('Chave Invalida')
                #exit(bot.alert('chave_xml invalida'))
            elif (len(dados_planilha[0]) < 4) or (len(dados_planilha[0]) == 5):
                marca_lancado('RE Inexistente')
            else:
                break

        # -------------------------------------- Lançamento Topcon --------------------------------------
        print(Fore.GREEN +  '--- Abrindo TopCompras para iniciar o lançamento' + Style.RESET_ALL)
        ahk.win_activate('TopCompras', title_match_mode=2)
        try:
            ahk.win_wait('TopCompras', title_match_mode=2, timeout= 5)
        except TimeoutError:
            exit(bot.alert('Tela de Compras não abriu.'))
            #TODO --- Caso isso aconteça, tentar abrir a tela do Topcon.
        else:
            print('--- Tela compras está maximizada!')  
        
        print('--- Alterando para o modo alteração')
        bot.press('F2')
        bot.press('F3')
        while procura_imagem(imagem='img_topcon/txt_inclui.png') is False:
            time.sleep(0.1)

        bot.doubleClick(558, 235)  # Clica dentro do campo para inserir a chave XML
        bot.write(chave_xml)
        bot.press('ENTER')
        ahk.win_wait_active('TopCompras')
        
        tentativa = 0
        while tentativa < 15: #* Aguarda até aparecer uma das telas que podem ser exibidas nesse processo.              
            #Verifica quais das telas apareceu. 
            if procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa=1, continuar_exec=True))
                print('--- XML Validado, indo para validação do pedido')
                return dados_planilha
            elif procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa= 1, continuar_exec=True) is not False:
                print('--- Nota já lançada, marcando planilha!')
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Lancado_Manual')
                break
            elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Aguardando_SEFAZ')
                break
            elif procura_imagem(imagem='img_topcon/chave_44digitos.png', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Chave_invalida')
                break
            elif procura_imagem(imagem='img_topcon/nfe_cancelada.png', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='NFE_CANCELADA')
                break
            
            # Verifica caso tenha travado e espera até que o topcom volte a responder
            print('--- Aguardando TopCompras Retornar')
            while ahk.win_exists('Não está respondendo', title_match_mode= 2):
                time.sleep(0.2)
            tentativa += 1
        else:
            exit('Rodou 10 verificações e não achou nenhuma tela, verificar!')
            #TODO --- Validar se o topcon ainda está aberto, caso não esteja, reiniciar o processo do zero.

if __name__ == '__main__':
    valida_lancamento()