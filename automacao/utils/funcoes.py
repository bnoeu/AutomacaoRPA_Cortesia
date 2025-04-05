# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import cv2
import datetime
import pytesseract
import numpy as np
import pyautogui as bot
from .configura_logger import get_logger


# --- Definição de parametros
from utils.funcoes import ahk as ahk
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.useImageNotFoundException(False)
logger = get_logger("automacao") # Obter logger configurado


def procura_imagem(imagem, limite_tentativa=5, area=(0, 0, 1920, 1080), continuar_exec=False, confianca = 0.78, msg_continuar_exec = False, msg_confianca = False):
    """Função que realiza o processo de OCR na tela, retornando as coordenadas onde localizou a imagem especificada.

    Args:
        imagem (Arquivo): imagem que deseja encontrar.
        limite_tentativa (int, optional): Quantas vezes deseja procurar. Defaults to 5.
        area (tuple, optional): Area onde deseja procurar. Defaults to (0, 0, 1920, 1080).
        continuar_exec (bool, optional): Continua a execução caso não encontre. Defaults to False.
        confianca (float, optional): _description_. Defaults to 0.78.
        msg_continuar_exec (bool, optional): _description_. Defaults to False.
        msg_confianca (bool, optional): _description_. Defaults to False.

    Returns:
        _type_: Retorna as posições onde encontrou a imagem.
    """    
    
    #from abre_topcon import main as abre_topcon
    #from Materia_Prima import programa_principal
    pausa_img = 0.3
    #hoje = datetime.date.today()
    maquina_viva = False
    tentativa = 0   
    logger.debug(F'--- Tentando encontrar: {imagem}')
    while tentativa < limite_tentativa:
        time.sleep(pausa_img)
        while maquina_viva is False:
            try:
                posicao_img = bot.locateCenterOnScreen(imagem, grayscale= True, confidence= confianca, region= area)
            except OSError:
                logger.critical('--- Erro devido a resolução da maquina virtual, aguardando')
                logger.exception()
                time.sleep(15)
                raise OSError
            else:
                maquina_viva = True
            
        if posicao_img is not None:
            logger.debug(F'--- Encontrou {imagem} na posição: {posicao_img}')
            break
        
        if msg_confianca is True:
            if confianca < 0.73:
                logger.debug(F'--- Valor atual da confiança da imagem: {confianca}', end= "")
            else:
                logger.debug(F', {confianca}', end= "")
        
        # Ajuste dos parametros
        confianca -= 0.01              
        tentativa += 1
        pausa_img += 0.25 
        

    #Caso seja para continuar
    if (continuar_exec is True) and (posicao_img is None): # Exibe a mensagem que o parametro está ativo
        if msg_continuar_exec is True:
            logger.info('' + F'--- {imagem} não foi encontrada, continuando execução pois o parametro "continuar_exec" está habilitado')
        return False
    
    if tentativa >= limite_tentativa: # Caso exceda o limite de tentativas
        time_atual = str(datetime.datetime.now()).replace(":","_").replace(".","_")
        caminho_erro = 'imagens/img_geradas/' + 'erro' + time_atual + '.png'
        img_erro = bot.screenshot()
        img_erro.save(fp= caminho_erro)
        raise TimeoutError
        
    return posicao_img

def verifica_tela(nome_tela, manual=False):
    if ahk.win_exists(nome_tela):
        logger.info(F'--- A tela: {nome_tela} está aberta')
        ahk.win_activate(nome_tela, title_match_mode=2)
        return True
    elif manual is True:
        logger.info(F'--- A tela: {nome_tela} está fechada, Modo Manual: {True}, executando...')
        return False
    else:
        exit(logger.error(F'--- Tela: {nome_tela} está fechada, saindo do programa.'))


def marca_lancado(texto_marcacao='Lancado'):
    bot.PAUSE = 0.4
    tentativa = 0
    
    logger.info('--- Abrindo planilha')
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    logger.info(F'--- Marcando planilha: {texto_marcacao}')
    
    while tentativa < 3:
        try:
            ahk.win_wait_active(title= 'debug_db', title_match_mode= 2, timeout= 5)
        except TimeoutError:
            ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
            tentativa += 1
        else:
            break
        
    time.sleep(0.5)
    bot.hotkey('CTRL', 'HOME')

    # Navega até o campo "Status"
    bot.press('RIGHT', presses= 6)
    bot.press('DOWN')
    
    # Informa o texto recebido pela função e passa para a celula ao lado, para inserir a data
    bot.write(texto_marcacao)
    bot.press('RIGHT')
    hoje = datetime.date.today()
    bot.write(str(hoje))
    time.sleep(0.5)
    bot.click(960, 640) # Clica no meio da planilha
    time.sleep(0.5)
    
    # Retorna a planilha para o modo "Somente Exibição (Botão Verde)"
    bot.hotkey('CTRL', 'HOME')
    reaplica_filtro_status() # Reaplica o filtro da coluna "Status"
    bot.hotkey('CTRL', 'HOME')
    logger.info(F'--------------------- Processou NFE, situação: {texto_marcacao} ---------------------')

def reaplica_filtro_status(): 
    bot.PAUSE = 1
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    logger.debug('--- Reaplicando o filtro na coluna "Status" ')
    time.sleep(0.5)
    bot.click(960, 640)
    
    bot.hotkey('CTRL', 'HOME') # Navega até o campo A1
    bot.press('RIGHT', presses= 6) # Navega até o campo "Status"
    bot.hotkey('ALT', 'DOWN') # Comando do excel para abrir o menu do filtro
    logger.debug('--- Navegou até celula A1 e abriu o filtro do status ')
    
    for i in range (0, 10):
        time.sleep(0.5)
        ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
        if procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png', continuar_exec= True, limite_tentativa= 3):
            break
    else:
        procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png', limite_tentativa= 1)


    bot.click(procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png', continuar_exec= True, limite_tentativa= 2, confianca= 0.73))
    logger.debug('--- na tela do menu de filtro, clicou no botão "Aplicar" para reaplicar o filtro ')
    
    if procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png', limite_tentativa= 3, confianca= 0.73, continuar_exec= True):
        bot.click(procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png'))
        time.sleep(2)
        logger.debug('--- Clicou para visualizar o filtro de todos.')
        


def extrai_txt_img(imagem, area_tela, porce_escala = 400):
    time.sleep(0.5)
    img = bot.screenshot('imagens/img_geradas/' + imagem, region=area_tela) # Captura uma screenshot da área especificada da tela
    logger.debug(F'--- Tirou uma screenshot da imagem: {imagem} ----')

    img = cv2.imread('imagens/img_geradas/' + imagem) # Lê a imagem capturada usando o OpenCV

    porce_escala = porce_escala # Define uma porcentagem de escala para redimensionar a imagem
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
    cv2.imwrite('imagens/img_geradas/img_thresh.png', img_thresh)
    
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
    texto_xml = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(168, 407, 250, 20))
    logger.info(F'--- Item da nota: {texto}, texto que ainda ficou: {texto_xml}, tamanho do texto {len(texto_xml)}')

    # Verifica pelo tamanho do texto, se ainda ficou algum valor no campo "Itens do pedido"
    if len(texto_xml) > 5: 
        logger.info('--- Itens XML ainda tem informação!')
        return False
    else:  # Caso fique vazio
        logger.info('--- Itens XML ficou vazio! saindo da tela de vinculação')
        ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
        time.sleep(0.4)
        bot.click(procura_imagem(imagem='imagens/img_topcon/confirma.png'))
        
        ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2, timeout= 30)
        while ahk.win_exists('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2):
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2)
            time.sleep(0.5)
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', confianca= 0.73, limite_tentativa= 10))

        ahk.win_wait_close('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
        logger.info('--- Encerrado a função verifica pedido vazio!')
        return True

def corrige_nometela(novo_nome = "TopCompras"):    
    try: # Verifica se o topcon abriu SEM NOME
        ahk.win_wait(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 5)
    
    except (TimeoutError, OSError): # Apresenta Timeout caso esteja aberto com o nome normal.
        try: # Verifica se REALMENTE abriu com o nome normal
            if ahk.win_wait(novo_nome, title_match_mode= 1, timeout= 8):
                logger.info('--- TopCompras abriu com o nome normal, prosseguindo.')
                return True
            else:
                exit(bot.alert('TopCompras não encontrado.'))
        except (TimeoutError, OSError):
            logger.warning("Não encontrou o TopCompras nem a tela sem nome")
            return
    else:
        ahk.win_set_title(new_title= novo_nome, title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)
        logger.error('--- Encontrou tela sem o nome, e realizou a correção!' )
            
if __name__ == '__main__':
    bot.PAUSE = 1
    bot.FAILSAFE = False
    reaplica_filtro_status()
    #verifica_ped_vazio()
    #corrige_nometela()
    #marca_lancado('Teste')