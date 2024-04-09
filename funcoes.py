# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import datetime
import cv2
#import pygetwindow as gw
import pytesseract
from ahk import AHK
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1  # Pausa padrão do bot
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"
bot.useImageNotFoundException(False)


def procura_imagem(imagem, limite_tentativa=6, area=(0, 0, 1920, 1080), continuar_exec=False, confianca = 0.75):
    tentativa = 0   
    print(F'--- Tentando encontrar: {imagem}', end= ' ')
    while tentativa < limite_tentativa:
        time.sleep(1)
        posicao_img = bot.locateCenterOnScreen(imagem, grayscale= True, confidence= confianca, region= area)
        if posicao_img is not None:
            print(F'Encontrou na posição: {posicao_img}')
            break
        tentativa += 1

    #Caso seja para continuar
    if (continuar_exec is True) and (posicao_img is None):
        print(F'' + 'Não encontrada, continuando execução pois o parametro "continuar_exec" está habilitado')
        return False
    if tentativa >= limite_tentativa:
        print('--- FECHANDO PLANILHA PARA EVITAR ERROS')
        #ahk.win_kill('db_alltrips')
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
    print('--- Abrindo planilha - MARCA_LANCADO')
    ahk.win_activate('db_alltrips', title_match_mode= 2)

    #Verifica se está no modo "Apenas exibição", caso esteja, altera para permitir edição.
    if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True) is not False:
        bot.click(procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', confianca = 0.8))
        bot.click(procura_imagem(imagem='img_planilha/botao_iniciaredicao.png'))
        
        #Caso apareça a tela informando que houve alteração durante esse periodo, confirma que quer atualizar e prossegue.
        if procura_imagem(imagem='img_planilha/txt_modificada.png', continuar_exec=True, limite_tentativa= 4) is not False: 
            print('--- Planilha atualizada, confirmando alterações..')
            bot.click(procura_imagem(imagem='img_planilha/bt_sim.png', limite_tentativa= 8, confianca= 0.45, area= (751, 521, 429, 218)))

        #Verifica se realmente entrou no modo edição.
        #procura_imagem(imagem='img_planilha/botao_edicao.png')
        
        #Clica na coluna Status
        bot.doubleClick(1494, 508)

        #Informa o texto recebido pela função e passa para a celula ao lado, para inserir a data
        bot.write(texto_marcacao)
        bot.press('RIGHT')
        hoje = datetime.date.today()
        bot.write(str(hoje))
        bot.press("ENTER")
        time.sleep(1)

        #Retorna a planilha para o modo "Somente Exibição (Botão Verde)"
        if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8, area = (1468, 400, 200, 200)) is not False:
            bot.click(procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8, area= (1468, 400, 200, 200)))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
        else:
            print('--- Não está filtrado, executando o filtro!')
            bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1529, 459, 75, 75)))
            while procura_imagem(imagem='img_planilha/botao_selecionartudo.png', confianca= 0.5) is None:
                time.sleep(1)
            bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png', confianca= 0.5))
            bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png', confianca= 0.5))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png',  confianca= 0.4))
    else: #Caso já esteja no modo "Edição"
        exit('--- Planilha no modo edição! Necessario scriptar essa parte')


def extrai_txt_img(imagem, area_tela):
    # Captura uma screenshot da área especificada da tela
    bot.screenshot('img_geradas/' + imagem, region=area_tela)
    print(F'--- Tirou print da imagem: {imagem} ----')

    # Lê a imagem capturada usando o OpenCV
    img = cv2.imread('img_geradas/' + imagem)

    # Define uma porcentagem de escala para redimensionar a imagem
    porce_escala = 185
    largura = int(img.shape[1] * porce_escala / 50)
    altura = int(img.shape[0] * porce_escala / 50)
    nova_dim = (largura, altura)

    # Redimensiona a imagem
    img = cv2.resize(img, nova_dim, interpolation=cv2.INTER_AREA)

    # Converte a imagem para tons de cinza
    img_cinza = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Aplica uma operação de limiarização para binarizar a imagem
    img_thresh = cv2.threshold(img_cinza, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
    
    #Exibe as imagens em caso de debug
    #cv2.imshow('thresh', img_thresh)
    #cv2.waitKey()

    # Utiliza o pytesseract para extrair texto da imagem binarizada
    texto = pytesseract.image_to_string(img_thresh, lang='eng', config='--psm 6').strip()
    return texto

def verifica_ped_vazio(texto, pos):
    #Extrai o texto da imagem 
    texto_xml = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(168, 400, 250, 30)).strip()
    print(F'Item da nota: {texto}, texto que ainda ficou: {texto_xml}')

    #Verifica pelo tamanho do texto, se ainda ficou algum valor no campo "Itens do pedido"
    if len(texto_xml) > 4: 
        print('Itens XML ainda tem informação!')
        return False
    else:  # Caso fique vazio
        print('Itens XML ficou vazio! prosseguindo')
        bot.click(procura_imagem(imagem='img_topcon/confirma.png', limite_tentativa= 10))
        bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', limite_tentativa= 10))
        return True