# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import datetime
import cv2
import pygetwindow as gw
import pytesseract
from ahk import AHK
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1.2  # Pausa padrão do bot
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
acabou_pedido = ''
numero_nf = "965999"
transportador = "111594"
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"


def procura_imagem(imagem, limite_tentativa=4, area=(0, 0, 1920, 1080), continuar_exec=False):
    tentativa = 0
    print( F'--- Tentando encontrar: {imagem}')
    while tentativa < limite_tentativa:
        time.sleep(0.5)
        posicao_img = bot.locateCenterOnScreen(
            imagem, grayscale=True, confidence=0.88, region=area)
        if posicao_img is not None:
            print(F'--- Imagem na posição: {posicao_img}')
            break
        tentativa += 1
    if (continuar_exec is True) and (posicao_img is None):
        print('--- Não encontrada, continuando execução pois o parametro "continuar_exec" está habilitado')
        return False
    if tentativa >= limite_tentativa:
        exit(bot.alert(text=F'Não foi possivel encontrar: {imagem}', title='Erro!', button='Fechar'))
    return posicao_img


def alteracao_filtro():
    time.sleep(0.5)
    if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8) is not False:
        print('--- Já está filtrado, continuando!')
    else:
        print('--- Não está filtrado, executando o filtro!')
        bot.click(procura_imagem(
            imagem='img_planilha/bt_setabaixo.png', area=(1463, 419, 100, 100)))
        bot.click(procura_imagem(
            imagem='img_planilha/botao_selecionartudo.png'))
        bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
        bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
        time.sleep(2)


def verifica_tela(nome_tela, manual=False):
    if ahk.win_exists(nome_tela):
        print(F'--- A tela: {nome_tela} está aberta')
        ahk.win_activate(nome_tela, title_match_mode= 2)
        return True
    elif manual is True:
        print(
            F'--- A tela: {nome_tela} está fechada, Modo Manual: {True}, executando...')
        return False
    else:
        exit(print(F'--- Tela: {nome_tela} está fechada, saindo do programa.'))


def marca_lancado(texto_marcacao='Lancado'):
    print('--- Abrindo planilha - MARCA_LANCADO')
    ahk.win_activate('db_alltrips')
    time.sleep(1)
    if bot.click(procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True, limite_tentativa=3)) is not False:
        time.sleep(2)
        bot.click(procura_imagem(imagem='img_planilha/botao_iniciaredicao.png'))
        time.sleep(2)
        if procura_imagem(imagem='img_planilha/txt_modificada.png', continuar_exec=True, limite_tentativa=2) is not False:
            bot.click(procura_imagem(imagem='img_planilha/bt_sim.png'))
        procura_imagem(imagem='img_planilha/botao_edicao.png')
        bot.doubleClick(1494, 508)
        bot.write(texto_marcacao)
        bot.press('RIGHT')
        hoje = datetime.date.today()
        bot.write(str(hoje))
        bot.press("ENTER")
        time.sleep(0.5)
        if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8) is not False:
            bot.click(procura_imagem(imagem='img_planilha/bt_filtro.png',
                      continuar_exec=True, limite_tentativa=8))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
            time.sleep(2)
        else:
            print('--- Não está filtrado, executando o filtro!')
            bot.click(procura_imagem(
                imagem='img_planilha/bt_setabaixo.png', area=(1463, 419, 100, 100)))
            bot.click(procura_imagem(
                imagem='img_planilha/botao_selecionartudo.png'))
            bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
            time.sleep(2)
    else:
        print('Não achou o botao de edição')


def extrai_txt_img(imagem, area_tela):
    bot.screenshot('img_geradas/' + imagem, region=area_tela)
    print(F'--- Tirou print da imagem: {imagem} ----')
    img = cv2.imread('img_geradas/' + imagem)
    scale_percent = 180
    width = int(img.shape[1] * scale_percent / 110)
    height = int(img.shape[0] * scale_percent / 110)
    dim = (width, height)
    img = cv2.resize(img, dim, interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    # cv2.imshow('T', thresh)
    # cv2.waitKey()
    time.sleep(0.5)
    texto = pytesseract.image_to_string(thresh, lang='eng', config='--psm 6').strip()
    return texto
