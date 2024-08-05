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
# tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.useImageNotFoundException(False)

def procura_imagem(imagem, limite_tentativa=5, area=(0, 0, 1920, 1080), continuar_exec=False, confianca = 0.78, msg_continuar_exec = False, msg_confianca = False):
    pausa_img = 0.2
    hoje = datetime.date.today()
    maquina_viva = False
    tentativa = 0   
    #print(F'--- Tentando encontrar: {imagem}')
    while tentativa < limite_tentativa:
        time.sleep(pausa_img)
        while maquina_viva is False:
            try:
                posicao_img = bot.locateCenterOnScreen(imagem, grayscale= True, confidence= confianca, region= area)
            except OSError:
                print('--- Erro devido a resolução da maquina virtual, aguardando')
                time.sleep(30)
            else:
                maquina_viva = True
            
        if posicao_img is not None:
            #print(F'--- Encontrou {imagem} na posição: {posicao_img}')
            break
        tentativa += 1
        pausa_img += 0.25 
    
        #TODO Aqui deveria ter um IF para validar se a MSG Confiança está como True
        
        if msg_confianca is True:
            if confianca < 0.73:
                print(F'--- Valor atual da confiança da imagem: {confianca}', end= "")
            else:
                print(F', {confianca}', end= "")
        
    
        confianca -= 0.01              
        

    #Caso seja para continuar
    if (continuar_exec is True) and (posicao_img is None):
        if msg_continuar_exec is True:
            print('' + F'--- {imagem} não foi encontrada, continuando execução pois o parametro "continuar_exec" está habilitado')
        return False
    
    if tentativa >= limite_tentativa:
        time_atual = str(datetime.datetime.now()).replace(":","_").replace(".","_")
        caminho_erro = 'img_geradas/' + 'erro' + time_atual + '.png'
        img_erro = bot.screenshot()
        img_erro.save(fp= caminho_erro)
        #ahk.win_kill('debug_db_alltrips')
        #! Não funcionando.
        # raise 
        #reinicia_mp()
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
    bot.PAUSE = 0.2
    tentativa = 0
    print(Fore.GREEN + F'--- Abrindo planilha - MARCA_LANCADO, com parametro: {texto_marcacao}' + Style.RESET_ALL)
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    while tentativa < 3:
        try:
            ahk.win_wait_active(title= 'debug_db', title_match_mode= 2, timeout= 5)
        except TimeoutError:
            ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
            tentativa += 1
        else:
            break
        
    bot.hotkey('CTRL', 'HOME')

    # Navega até o campo "Status"
    bot.press('RIGHT', presses= 6)
    bot.press('DOWN')
    
    # Informa o texto recebido pela função e passa para a celula ao lado, para inserir a data
    bot.write(texto_marcacao)
    bot.press('RIGHT')
    hoje = datetime.date.today()
    bot.write(str(hoje))
    time.sleep(0.2)
    bot.click(500, 500)
    
    # Retorna a planilha para o modo "Somente Exibição (Botão Verde)"
    bot.hotkey('CTRL', 'HOME')
    if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, area = (1468, 400, 200, 200)) is not False:
        print('--- Encontrou o botão do filtro, navegando no menu do filtro')
        bot.press('RIGHT', presses= 6) # Navega até o campo "Status"
        bot.hotkey('ALT', 'DOWN') # Comando do excel para abrir o menu do filtro
        bot.press('TAB', presses= 10)
        bot.press('ENTER')
        print('--- Saindo do menu do filtro')
        #exit(bot.alert('Verificar se filtrou!'))
    else:
        print('--- Não está filtrado, executando o filtro!')
        bot.hotkey('CTRL', 'HOME')
        bot.press('RIGHT', presses= 6)
        bot.move(500, 500)
        bot.hotkey('alt', 'down') 
        
        print('--- Aguardando aparecer o botão selecionar tudo')
        while procura_imagem(imagem='img_planilha/botao_selecionartudo.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.73) is None:
            time.sleep(0.1)
            
        print('--- Reaplicando o filtro para as notas vazias')
        # Navega até a opção "Selecionar tudo", para reaplicar o filtro de notas vazias.
        bot.press('TAB', presses= 9)
        bot.press('ENTER')

    print(Fore.GREEN + F'--------------------- Processou NFE, situação: {texto_marcacao} ---------------------' + Style.RESET_ALL)

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
    texto_xml = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(168, 407, 250, 20))
    print(F'--- Item da nota: {texto}, texto que ainda ficou: {texto_xml}, tamanho do texto {len(texto_xml)}')

    # Verifica pelo tamanho do texto, se ainda ficou algum valor no campo "Itens do pedido"
    if len(texto_xml) > 5: 
        print('--- Itens XML ainda tem informação!')
        return False
    else:  # Caso fique vazio
        print('--- Itens XML ficou vazio! saindo da tela de vinculação')
        time.sleep(0.2)
        ahk.win_activate('TopCompras', title_match_mode= 2)
        bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
        bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg'))
        print('--- Encerrado a função verifica pedido vazio!')
        return True

def corrige_topcompras():
    try: 
        ahk.win_wait(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 5)
    except TimeoutError:
        try:
            if ahk.win_wait('TopCompras', title_match_mode= 1, timeout= 5):
                pass
                print('--- TopCompras abriu com o nome normal, prosseguindo.')
            else:
                bot.alert(exit('TopCompras não encontrado.'))
        except TimeoutError:
            return
    else:
        ahk.win_set_title(new_title= 'TopCompras', title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)
        print(Fore.GREEN + '--- Encontrou tela sem o nome, e realizou a correção!' + Style.RESET_ALL)
            
if __name__ == '__main__':
    corrige_topcompras()
    #marca_lancado('Teste')