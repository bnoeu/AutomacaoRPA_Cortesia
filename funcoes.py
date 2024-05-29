# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import cv2
import datetime
import pytesseract
import numpy as np
from ahk import AHK
import pyautogui as bot
#import pygetwindow as gw
from colorama import Fore, Style


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
bot.PAUSE = 0.8
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.useImageNotFoundException(False)

def procura_imagem(imagem, limite_tentativa=12, area=(0, 0, 1920, 1080), continuar_exec=False, confianca = 0.75):
    tentativa = 0   
    #print(F'--- Tentando encontrar: {imagem}', end= ' ')
    while tentativa < limite_tentativa:
        time.sleep(0.2)
        posicao_img = bot.locateCenterOnScreen(imagem, grayscale= True, confidence= confianca, region= area)
        if posicao_img is not None:
            print(F'--- Encontrou {imagem} na posição: {posicao_img}')
            break
        tentativa += 1

    #Caso seja para continuar
    if (continuar_exec is True) and (posicao_img is None):
        #print('' + F'--- {imagem} não foi encontrada, continuando execução pois o parametro "continuar_exec" está habilitado')
        time.sleep(0.2)
        return False
    if tentativa >= limite_tentativa:
        #print('--- FECHANDO PLANILHA PARA EVITAR ERROS')
        bot.screenshot('img_geradas/ERRO_' + imagem)
        ahk.win_kill('db_alltrips')
        exit(bot.alert(text=F'Não foi possivel encontrar: {imagem}', title='Erro!', button='Fechar'))
    return posicao_img


def verifica_tela(nome_tela, manual=False):
    if ahk.win_exists(nome_tela):
        print(F'--- A tela: {nome_tela} está aberta')
        ahk.win_activate(nome_tela, title_match_mode=2)
        return True
    elif manual is True:
        print(F'--- A tela: {nome_tela} está fechada, Modo Manual: {True}, executando...')
        return False
    else:
        exit(print(F'--- Tela: {nome_tela} está fechada, saindo do programa.'))


def marca_lancado(texto_marcacao='Lancado'):
    time.sleep(0.5)
    print(Fore.GREEN + F'\n--- Abrindo planilha - MARCA_LANCADO, com parametro: {texto_marcacao}' + Style.RESET_ALL)
    ahk.win_activate('db_alltrips', title_match_mode= 2)
    ahk.win_wait_active('db_alltrips', title_match_mode= 2, timeout= 15)
    time.sleep(0.5)

    #Verifica se está no modo "Apenas exibição", caso esteja, altera para permitir edição.
    if procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', continuar_exec=True) is not False:
        bot.click(procura_imagem(imagem='img_planilha/bt_exibicaoverde.png', limite_tentativa= 50))
        bot.click(procura_imagem(imagem='img_planilha/botao_iniciaredicao.png', limite_tentativa= 50))
        print('--- Aguardando entrar no modo edição')
        while procura_imagem(imagem='img_planilha/bt_edicao.png', continuar_exec= True) is False: #Aguarda até entrar no modo Edição
            #Caso apareça a tela informando que houve alteração durante esse periodo, confirma que quer atualizar e prossegue.
            if procura_imagem(imagem='img_planilha/txt_modificada.png', continuar_exec=True, limite_tentativa= 10) is not False: 
                bot.click(procura_imagem(imagem='img_planilha/bt_sim.png', limite_tentativa= 10, area= (751, 521, 429, 218)))
    else:
        print(F'--- Planilha já no modo edição, continuando a inserção do texto: {texto_marcacao}')

    #Clica no campo status, e move para baixo, para permitir inserir o texto passado na função.
    bot.click(procura_imagem(imagem='img_planilha/txt_status.png', confianca= 0.75))
    bot.press('DOWN')
    #Informa o texto recebido pela função e passa para a celula ao lado, para inserir a data
    bot.write(texto_marcacao)
    bot.press('RIGHT')
    hoje = datetime.date.today()
    bot.write(str(hoje))

    #Retorna a planilha para o modo "Somente Exibição (Botão Verde)"
    if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, area = (1468, 400, 200, 200)) is not False:
        bot.click(procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, area= (1468, 400, 200, 200)))
        bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
    else:
        print('--- Não está filtrado, executando o filtro!')        
        bot.click(procura_imagem(imagem='img_planilha/txt_status.png'))
        bot.move(500, 500)
        bot.hotkey('alt', 'down')        
        
        while procura_imagem(imagem='img_planilha/botao_selecionartudo.png', limite_tentativa= 1) is None:
            time.sleep(0.1)
        bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png'))
        bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
        bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))

    print(Fore.GREEN + F'--------------------- Processou NFE, situação: {texto_marcacao} ---------------------\n' + Style.RESET_ALL)

def extrai_txt_img(imagem, area_tela, porce_escala = 400):
    # Captura uma screenshot da área especificada da tela
    img = bot.screenshot('img_geradas/' + imagem, region=area_tela)
    print(F'--- Tirou print da imagem: {imagem} ----')

    # Lê a imagem capturada usando o OpenCV
    img = cv2.imread('img_geradas/' + imagem)

    # Define uma porcentagem de escala para redimensionar a imagem
    porce_escala = porce_escala
    largura = int(img.shape[1] * porce_escala / 90)
    altura = int(img.shape[0] * porce_escala / 90)
    nova_dim = (largura, altura)
    img = cv2.resize(img, nova_dim, interpolation=cv2.INTER_AREA) # Redimensiona a imagem
    img_cinza = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) # Converte a imagem para tons de cinza
    kernel = np.ones((5,5),np.float32)/30
    smooth = cv2.filter2D(img_cinza,-1,kernel)
    img_thresh = cv2.threshold(smooth, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1] #OTSU threshold    
    
    # Utiliza o pytesseract para extrair texto da imagem binarizada
    texto = pytesseract.image_to_string(img_thresh, lang='eng', config='--psm 7').strip()
    cv2.imwrite('img_geradas\img_thresh.png', img_thresh)
    
    '''
    #Exibe as imagens em caso de debug
    cv2.imshow('img', img)
    cv2.imshow('img_cinza', img_cinza)
    #cv2.imshow('blur', blur)
    cv2.imshow('thresh', img_thresh)
    cv2.imshow('smooth', smooth)
    #cv2.imshow('erosion', erosion)
    cv2.waitKey()
    '''


    return texto

def verifica_ped_vazio(texto, pos):
    #Extrai o texto da imagem 
    #texto_xml = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(168, 400, 250, 30)).strip().replace('_','')
    texto_xml = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(168, 407, 250, 20))
    print(F'--- Item da nota: {texto}, texto que ainda ficou: {texto_xml}, tamanho do texto {len(texto_xml)}')

    #Verifica pelo tamanho do texto, se ainda ficou algum valor no campo "Itens do pedido"
    if len(texto_xml) > 5: 
        print('--- Itens XML ainda tem informação!')
        return False
    else:  # Caso fique vazio
        bot.click(procura_imagem(imagem='img_topcon/confirma.png', limite_tentativa= 100))
        bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', limite_tentativa= 100))
        return True



if __name__ == '__main__':
    verifica_ped_vazio()