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
tempo_inicio = time.time() #Define o tempo inicial do lançamento
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"

def coleta_planilha():
    bot.PAUSE = 0.3  # Pausa padrão do bot
    print('--- Abrindo planilha - COLETA_PLANILHA')
    ahk.win_activate('db_alltrips', title_match_mode= 2)

    #Verifica se já está no modo de edição, caso esteja, muda para o modo "exibição"
    if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True) is False:
        bot.click(procura_imagem(imagem='img_planilha/botao_edicao.png'))
        bot.click(procura_imagem(imagem='img_planilha/botao_exibicao.png'))
    else: #Caso não esteja no modo "Edição"
        print('--- A planilha já está no modo "Exibição", continuando processo')

    #Altera o filtro para "vazio", para iniciar a coleta de dados.
    if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, area= (1468, 400, 200, 200)) is not False:
        print('--- Já está filtrado, continuando!')
    else:
        print('--- Não está filtrado, executando o filtro!')
        bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1529, 459, 75, 75)))
        ahk.win_activate('Sem título', title_match_mode= 2)
        #Caso não apareça o botão "Selecionar tudo" clica em "limpar filtro"
        if procura_imagem(imagem='img_planilha/botao_selecionartudo.png', confianca= 0.5, continuar_exec= True) is False:
            bot.click(procura_imagem(imagem='img_planilha/bt_limparFiltro.png', confianca= 0.5))
            coleta_planilha()
        else:
            bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png', confianca= 0.5))
            bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png', confianca= 0.5))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png', confianca= 0.4))
    
    # * Coleta os dados da linha atual
    dados_planilha = []
    print('--- Copiando dados e formatando')
    time.sleep(1)
    bot.click(100, 510)  # Clica na primeira linha e coluna da planilha
    for n in range(0, 7, 1):  # Copia dados dos 6 campos
        while True:
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                time.sleep(0.1)
            else:
                break
        dados_planilha.append(ahk.get_clipboard())
        bot.press('right')
    print(F'--- Dados copiados com sucesso: {dados_planilha}')
    tempo_coleta = time.time() - tempo_inicio
    print(F'\n Tempo que levou: {tempo_coleta}')
    return dados_planilha