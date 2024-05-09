# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
# import cv2
# import pygetwindow as gw
import pytesseract
from ahk import AHK
from colorama import Fore, Style
from funcoes import procura_imagem
import pyautogui as bot

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.PAUSE = 1

def coleta_planilha():
    bot.PAUSE = 0.8
    print(Fore.CYAN + '--- Abrindo planilha - COLETA_PLANILHA' + Style.RESET_ALL)
    ahk.win_activate('db_alltrips', title_match_mode= 2)
    ahk.win_wait('db_alltrips', title_match_mode= 2)

    #Verifica se já está no modo de edição, caso esteja, muda para o modo "exibição"
    if procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', continuar_exec=True) is False:
        print('--- Não está no modo exibição! Realizando alteração.')
        
        #Espera até encontar o botão "Exibição" (Lapis bloqueado) e realiza um click nele
        while procura_imagem(imagem='img_planilha/bt_edicao.png', continuar_exec= True) is False:
            time.sleep(0.2)
            
        #Clica no botão da edição (lapis branco), para mostrar o dropdown com as opções, e clica no botão "lapis bloqueado"
        bot.click(procura_imagem(imagem='img_planilha/bt_edicao.png', continuar_exec= True))
        bot.click(procura_imagem(imagem='img_planilha/bt_exibicao.png', confianca= 0.90, continuar_exec= True))
        #Aguarda até aparecer o botão do modo "exibição"
        while procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', continuar_exec=True) is False:
            time.sleep(0.2)
        print('--- Alterado para o modo exibição, continuando.')
        
    else: #Caso não esteja no modo "Edição"
        print('--- A planilha já está no modo "Exibição", continuando processo')

    #Altera o filtro para "vazio", para iniciar a coleta de dados.
    if procura_imagem(imagem='img_planilha/bt_filtro.png', confianca= 0.5, continuar_exec=True, area= (1468, 400, 200, 200)) is not False:
        print('--- Já está filtrado, continuando!')
    else:
        print('--- Não está filtrado, executando o filtro!')
        bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1529, 459, 75, 75)))
        #Caso não apareça o botão "Selecionar tudo" clica em "limpar filtro" e executa tudo novamente.
        if procura_imagem(imagem='img_planilha/botao_selecionartudo.png', continuar_exec= True) is False:
            bot.click(procura_imagem(imagem='img_planilha/bt_limparFiltro.png'))
            coleta_planilha()
        else: #Se tudo estiver ok, prossegue aplicando o filtro nas notas vazias. 
            bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png'))
            bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
            print('--- Filtrado pelas notas vazias!')

            #Aguarda aparecer o botão do filtro, para confirmar que está filtrado! 
            while procura_imagem(imagem='img_planilha/bt_filtro.png', limite_tentativa= 10, area= (1468, 400, 200, 200)) is False:
                print('--- Aguardando o botão do filtro na coluna "Status" ')
                time.sleep(0.6)
            else:
                print('--- Filtro das notas vazias aplicado!')
    
    # * Coleta os dados da linha atual
    dados_planilha = []
    print('--- Copiando dados e formatando')
    #Clica na primeira linha (Campo RE), e pressiona seta para baixo
    bot.PAUSE = 0.3
    bot.click(procura_imagem(imagem='img_planilha/titulo_re.png'))
    bot.press('DOWN')
    time.sleep(0.2)
    for n in range(0, 7, 1):  # Copia dados dos 6 campos
        time.sleep(0.2)
        while True:
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                time.sleep(0.2)
            else:
                break
        dados_planilha.append(ahk.get_clipboard())
        bot.press('right')
    bot.PAUSE = 1
    tempo_coleta = time.time() - tempo_inicio
    tempo_coleta = tempo_coleta
    print(F'--- Tempo que levou: {tempo_coleta:0f} segundos')
    print(Fore.GREEN +  F'--- Dados copiados com sucesso: {dados_planilha}' + Style.RESET_ALL + '\n')
    return dados_planilha

if __name__ == '__main__':
    coleta_planilha()