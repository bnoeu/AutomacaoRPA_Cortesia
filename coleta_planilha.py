# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pytesseract
from ahk import AHK
from funcoes import procura_imagem
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"

def coleta_planilha():
    bot.PAUSE = 1
    print('--- Abrindo planilha - COLETA_PLANILHA')
    ahk.win_activate('db_alltrips', title_match_mode= 2)

    #Verifica se já está no modo de edição.
    if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True) is False:
        bot.click(procura_imagem(imagem='img_planilha/botao_edicao.png'))
        bot.click(procura_imagem(imagem='img_planilha/botao_exibicao.png'))
        #procura_imagem(imagem='img_planilha/botao_exibicaoverde.png')
    else:
        print('--- Já está no modo de edição, continuando processo')

    #Altera o filtro para "vazio", para iniciar a coleta de dados.
    if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8, area= (1468, 400, 200, 200)) is not False:
        print('--- Já está filtrado, continuando!')
    else:
        print('--- Não está filtrado, executando o filtro!')
        bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1529, 459, 75, 75)))
        bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png', limite_tentativa= 30, confianca= 0.5))
        bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png', confianca= 0.5))
        bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png', confianca= 0.4))    
    
    # * Coleta os dados da linha atual
    bot.PAUSE = 0.5  # Pausa padrão do bot
    dados_planilha = []
    print('--- Copiando dados e formatando')
    bot.click(100, 510)  # Clica na primeira linha e coluna da planilha
    for n in range(0, 7, 1):  # Copia dados dos 6 campos
        while True:
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                time.sleep(0.15)
            else:
                break
        dados_planilha.append(ahk.get_clipboard())
        bot.press('right')
    print(F'--- Dados copiados com sucesso: {dados_planilha}')
    return dados_planilha