# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pytesseract
from ahk import AHK
from Material_Prima import programa_principal
from funcoes import marca_lancado, procura_imagem
from coleta_planilha import coleta_planilha
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"  # Valor para teste
transportador = "111594"  # Valor para teste
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"



#* -------------------------- Programa principal
def acoes_planilha():
    time.sleep(0.5)
    validou_xml = False
    while validou_xml is False:
        # * Trata os dados coletados em "dados_planilha"
        dados_planilha = coleta_planilha()
        exit()
        bot.PAUSE = 1.5
        print(F'--- Alterado PAUSE {bot.PAUSE} ')
        chave_xml = dados_planilha[4].strip()
        # * -------------------------------------- Lançamento Topcon --------------------------------------
        print('--- Abrindo TopCompras')
        ahk.win_activate('TopCompras', title_match_mode=2)
        if ahk.win_is_active('TopCompras', title_match_mode=2):
            print('Tela compras está maximizada! Iniciando o programa')
        else:
            exit(bot.alert('Tela de Compras não abriu.'))
        # Processo de lançamento
        bot.press('F2')
        bot.press('F3', presses=2, interval=0.3)
        bot.click(558, 235)  # Clica dentro do campo para inserir a chave XML
        bot.write(chave_xml)
        bot.press('ENTER')
        ahk.win_wait_active('TopCompras')
        while procura_imagem(imagem='img_topcon/naorespondendo.png', limite_tentativa=3, continuar_exec=True) is not False:
            time.sleep(0.5)
            print('Aguardando topvoltar')
        tentativa = 0
        while tentativa < 10:
            if procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa=3, continuar_exec=True) is not False:
                bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg',
                          limite_tentativa=2, continuar_exec=True))
                validou_xml is True
                return dados_planilha
            elif procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa=3, continuar_exec=True) is not False:
                print('--- Nota já lançada, marcando planilha!')
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Lancado_Manual')
                break
            elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', limite_tentativa=3, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Aguardando_SEFAZ')
                programa_principal()
            elif procura_imagem(imagem='img_topcon/chave_44digitos.png', limite_tentativa=3, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Chave_invalida')
                programa_principal()
            elif procura_imagem(imagem='img_topcon/transportador_incorreto.png', limite_tentativa=3, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Transportador_incorreto')
                programa_principal()
            tentativa += 1
        if tentativa >= 15:
            exit('Rodou 10 verificações e não achou nenhuma tela, aumentar o tempo')

