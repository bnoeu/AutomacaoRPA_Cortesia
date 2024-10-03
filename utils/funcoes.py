# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import os
import time
import cv2
from datetime import datetime
import logging
import pytesseract
import numpy as np
from ahk import AHK
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.useImageNotFoundException(False)
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"
alltrips = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg?e=RYfgcA.com"

def configurar_logging(nome_arquivo, nivel_log = logging.INFO):
    horario_inicio = datetime.now()
    horario_inicio = F"D{horario_inicio.day}-{horario_inicio.month}__H{horario_inicio.hour}-{horario_inicio.minute}_"

    logging.basicConfig(
        filename= F"logs/{nome_arquivo}_{horario_inicio}.log",
        filemode= "a",
        encoding= "utf-8",
        level= nivel_log,
        format= "{asctime} - {levelname} - {message}",
        style="{",
        )


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
    

    pausa_img = 0.2
    maquina_viva = False
    tentativa = 0   
    logging.debug(F'--- Tentando encontrar: {imagem}')
    while tentativa < limite_tentativa:
        time.sleep(pausa_img)
        while maquina_viva is False:
            try:
                posicao_img = bot.locateCenterOnScreen(imagem, grayscale= True, confidence= confianca, region= area)
            except OSError:
                logging.critical('--- Erro devido a resolução da maquina virtual, aguardando')
                time.sleep(15)
                raise OSError
            else:
                maquina_viva = True
            
        if posicao_img is not None:
            logging.debug(F'--- Encontrou {imagem} na posição: {posicao_img}')
            break
        
        if msg_confianca is True:
            if confianca < 0.73:
                logging.debug(F'--- Valor atual da confiança da imagem: {confianca}', end= "")
            else:
                logging.debug(F', {confianca}', end= "")
        
        # Ajuste dos parametros
        confianca -= 0.02           
        tentativa += 1
        pausa_img += 0.05
        

    #Caso seja para continuar
    if (continuar_exec is True) and (posicao_img is None): # Exibe a mensagem que o parametro está ativo
        if msg_continuar_exec is True:
            logging.info('' + F'--- {imagem} não foi encontrada, continuando execução pois o parametro "continuar_exec" está habilitado')
        return False
    
    if tentativa >= limite_tentativa: # Caso exceda o limite de tentativas
        logging.warning(F'--- Não encontrou a imagem: {imagem}')
        time_atual = str(datetime.now()).replace(":","_").replace(".","_")
        caminho_erro = 'imagens/img_geradas/' + 'erro' + time_atual + '.png'
        img_erro = bot.screenshot()
        img_erro.save(fp= caminho_erro)
        raise TimeoutError
        
    return posicao_img

def verifica_tela(nome_tela, manual=False):
    if ahk.win_exists(nome_tela):
        logging.info(F'--- A tela: {nome_tela} está aberta')
        ahk.win_activate(nome_tela, title_match_mode=2)
        return True
    elif manual is True:
        logging.info(F'--- A tela: {nome_tela} está fechada, Modo Manual: {True}, executando...')
        return False
    else:
        exit(logging.error(F'--- Tela: {nome_tela} está fechada, saindo do programa.'))

def marca_lancado(texto_marcacao='Lancado'):
    bot.PAUSE = 0.6
    tentativa = 0
    logging.info(F'--- Abrindo planilha - MARCA_LANCADO, com parametro: {texto_marcacao}' )
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    while tentativa < 3:
        try:
            ahk.win_wait_active(title= 'debug_db', title_match_mode= 2, timeout= 5)
        except TimeoutError:
            ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
            tentativa += 1
        else:
            break
        
    time.sleep(0.25)
    bot.hotkey('CTRL', 'HOME')

    # Navega até o campo "Status"
    bot.press('RIGHT', presses= 6)
    bot.press('DOWN')
    
    # Informa o texto recebido pela função e passa para a celula ao lado, para inserir a data
    bot.write(texto_marcacao)
    bot.press('RIGHT')
    hoje = datetime.now()
    hoje_formatado = hoje.strftime('%d/%m/%Y')
    bot.write(hoje_formatado)
    time.sleep(0.25)
    bot.click(960, 640) # Clica no meio da planilha
    time.sleep(0.25)
    
    # Retorna a planilha para o modo "Somente Exibição (Botão Verde)"
    bot.hotkey('CTRL', 'HOME')
    reaplica_filtro_status() # Reaplica o filtro da coluna "Status"
    bot.hotkey('CTRL', 'HOME')
    logging.info(F'--------------------- Processou NFE, situação: {texto_marcacao} ---------------------')

def reaplica_filtro_status(): 
    bot.PAUSE = 0.6
    logging.info('--- Executando a função REAPLICA FILTRO STATUS')
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    time.sleep(0.25)
    bot.click(960, 640)
    
    bot.hotkey('CTRL', 'HOME') # Navega até o campo A1
    bot.press('RIGHT', presses= 6) # Navega até o campo "Status"
    bot.hotkey('ALT', 'DOWN') # Comando do excel para abrir o menu do filtro
    logging.debug('--- Navegou até celula A1 e abriu o filtro do status ')
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    bot.click(procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png'))
    logging.debug('--- Na tela do menu de filtro, clicou no botão "Aplicar" para reaplicar o filtro ')
    
    if procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png', continuar_exec= True):
        bot.click(procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png', continuar_exec= True))
        logging.debug('--- Clicou para visualizar o filtro de todos.')
    
    time.sleep(1)
    bot.hotkey('CTRL', 'HOME') # Navega até o campo A1
        

def extrai_txt_img(imagem, area_tela, porce_escala = 400):
    time.sleep(0.25)
    img = bot.screenshot('imagens/img_geradas/' + imagem, region=area_tela) # Captura uma screenshot da área especificada da tela
    logging.debug(F'--- Tirou print da imagem: {imagem} ----')

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
    logging.info(F'--- Item da nota: {texto}, texto que ainda ficou: {texto_xml}, tamanho do texto {len(texto_xml)}')

    # Verifica pelo tamanho do texto, se ainda ficou algum valor no campo "Itens do pedido"
    if len(texto_xml) > 5: 
        logging.info('--- Itens XML ainda tem informação!')
        return False
    else:  # Caso fique vazio
        logging.info('--- Itens XML ficou vazio! saindo da tela de vinculação')
        ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
        time.sleep(0.25)
        bot.click(procura_imagem(imagem='imagens/img_topcon/confirma.png'))
        
        ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2, timeout= 30)
        while ahk.win_exists('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2):
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2)
            time.sleep(0.25)
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', confianca= 0.73, limite_tentativa= 10))

        ahk.win_wait_close('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
        logging.info('--- Encerrado a função verifica pedido vazio!')
        return True

def corrige_nometela(novo_nome = "TopCompras"):    
    
    
    try: # Verifica se o topcon abriu SEM NOME
        ahk.win_wait(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 5)
    
    except (TimeoutError, OSError): # Apresenta Timeout caso esteja aberto com o nome normal.
        try: # Verifica se REALMENTE abriu com o nome normal
            if ahk.win_wait(novo_nome, title_match_mode= 1, timeout= 8):
                logging.info('--- TopCompras abriu com o nome normal, prosseguindo.')
                return
            else:
                exit(bot.alert('TopCompras não encontrado.'))
        except (TimeoutError, OSError):
            logging.warning("Não encontrou o TopCompras nem a tela sem nome")
            return
    else:
        ahk.win_set_title(new_title= novo_nome, title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)
        logging.warning('--- Encontrou tela sem o nome, e realizou a correção!' )


def abre_planilha_navegador(link_planilha = alltrips):
    if link_planilha == alltrips: # Planilha original
        planilha = "db_alltrips.xlsx" 
    else: # Planilha de debug
        planilha = "debug_db_alltrips.xlsx"   
    
    logging.info(F'--- Tentando abrir a planilha: {planilha}')
    if ahk.win_exists(planilha):
        logging.info('--- Planilha já estava aberta, executou apenas um recarregamento')
        ahk.win_activate(planilha, title_match_mode= 2)
        bot.hotkey('CTRL', 'F5') # Recarrega a planilha limpando o cache
        time.sleep(8)
        logging.info('--- Recarregou a planilha com sucesso.')
        return True
    
    while ahk.win_exists("alltrips.xlsx", title_match_mode= 2): # Garante que a planilha não esteja aberta
        ahk.win_close('alltrips.xlsx', title_match_mode= 2)
        time.sleep(0.5)
    
    logging.info('--- Abrindo a planilha no EDGE.')
    comando_iniciar = F'start msedge {link_planilha} -new-window -inprivate'
    os.system(comando_iniciar)
    ahk.win_wait_active(planilha, title_match_mode = 2, timeout= 8)
    ahk.win_maximize(planilha)
    time.sleep(10)
    logging.info('--- Planilha aberta e maximizada.')

def msg_box(texto, tempo):
    """Exibe uma dialog box temporaria utilizando AHK

    Args:
        texto (str): O texto que irá aparecer na mensagem
        tempo (int): Tempo até o fechamento
    """
    ahk.msg_box(text= texto, blocking= False)
    time.sleep(tempo)
    bot.press('ENTER')
            
if __name__ == '__main__':
    bot.PAUSE = 0.6
    bot.FAILSAFE = False
    abre_planilha_navegador()
    #reaplica_filtro_status()
    #verifica_ped_vazio()
    #corrige_nometela()
    #marca_lancado('Teste')