# Para utilização na Cortesia Concreto.
# -*- Criado por Bruno da Silva Santos. -*-

import os
import time
import cv2
import pytesseract
import numpy as np
from ahk import AHK
import pyautogui as bot
from datetime import datetime

if __name__ == '__main__':
    from configura_logger import get_logger
else:
    from .configura_logger import get_logger


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
bot.useImageNotFoundException(False)
logger = get_logger("")
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=jFhSykjw3Eg"
alltrips = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg?e=RYfgcA.com"

def print_erro(nome_img = "erro"):
    time_atual = str(datetime.now()).replace(":","_").replace(".","_")
    caminho_erro = 'imagens/img_geradas/erros/' + F"{nome_img}" + time_atual + '.png'
    img_erro = bot.screenshot()
    img_erro.save(fp= caminho_erro)
    return caminho_erro

def procura_imagem(imagem, limite_tentativa=5, area=(0, 0, 1920, 1080), continuar_exec=False, confianca = 0.75):
    """Função que realiza o processo de OCR na tela, retornando as coordenadas onde localizou a imagem especificada.

    Args:
        imagem (Arquivo): imagem que deseja encontrar.
        limite_tentativa (int, optional): Quantas vezes deseja procurar. Defaults to 5.
        area (tuple, optional): Area onde deseja procurar. Defaults to (0, 0, 1920, 1080).
        continuar_exec (bool, optional): Continua a execução caso não encontre. Defaults to False.
        confianca (float, optional): _description_. Defaults to 0.78.

    Returns:
        _type_: Retorna as posições onde encontrou a imagem.
    """    
    
    pausa_img = 0.12
    maquina_viva = False
    tentativa = 0  
    logger.debug(F'--- Tentando encontrar: {imagem}')
    while tentativa < limite_tentativa:
        time.sleep(pausa_img)
        while maquina_viva is False:
            try:
                posicao_img = bot.locateCenterOnScreen(imagem, grayscale= True, confidence= confianca, region= area)
            except OSError as e:
                logger.critical(F'--- Erro devido a resolução da maquina virtual, aguardando, erro coletado: \n{e}')
                time.sleep(15)
                raise OSError
            else:
                maquina_viva = True
            
        if posicao_img is not None:
            logger.debug(F'--- Encontrou {imagem} na posição: {posicao_img} ( Tentativa: {tentativa}, Confiança: {confianca} )')
            break
                
        # Ajuste dos parametros
        confianca -= 0.02           
        tentativa += 1
        pausa_img += 0.05

    #* Caso seja para continuar
    if (continuar_exec is True) and (posicao_img is None): # Exibe a mensagem que o parametro está ativo
        logger.debug('' + F'--- {imagem} não foi encontrada, "continuar_exec" está habilitado (Tentativa: {tentativa}, Confiança {confianca}, Pausa: {pausa_img} Area: {area})')
        return False
    
    if tentativa >= limite_tentativa: # Caso exceda o limite de tentativas
        logger.warning(F'--- Não encontrou a imagem: {imagem}, ( Tentativa: {tentativa}, Confiança: {confianca} ) ')
        time_atual = str(datetime.now()).replace(":","_").replace(".","_")
        caminho_erro = 'imagens/img_geradas/erros/' + 'erro' + time_atual + '.png'
        img_erro = bot.screenshot()
        img_erro.save(fp= caminho_erro)
        ahk.get_active_window()
        raise Exception(F"--- Falhou ao procurar imagem: {imagem}, confiança: {confianca}")
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

def marca_lancado(texto_marcacao='teste_08_12'):
    bot.PAUSE = 0.2

    logger.info(F'--- Abrindo planilha - MARCA_LANCADO, com parametro: {texto_marcacao}' )
    ativar_janela('debug_db_alltrips', 30)
    ''' #! Substituido pela logica a cima
    time.sleep(1)
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    ahk.win_wait_active('debug_db_alltrips', title_match_mode= 2, timeout= 30, detect_hidden_windows = True)
    '''

    time.sleep(1)
    bot.hotkey('CTRL', 'HOME')

    # Navega até o campo "Status"
    bot.press('RIGHT', presses= 6)
    bot.press('DOWN')
    time.sleep(0.8)
    
    # Informa o texto recebido pela função e passa para a celula ao lado, para inserir a data
    bot.write(texto_marcacao)
    bot.press('RIGHT')
    time.sleep(0.2)
    hoje = datetime.now()
    hoje_formatado = hoje.strftime('%d/%m/%Y')
    bot.write(hoje_formatado)
    time.sleep(0.2)
    bot.click(960, 640) # Clica no meio da planilha
    time.sleep(0.2)
    
    # Retorna a planilha para o modo "Somente Exibição (Botão Verde)"
    bot.hotkey('CTRL', 'HOME')
    reaplica_filtro_status() # Reaplica o filtro da coluna "Status"
    bot.hotkey('CTRL', 'HOME')
    logger.info(F'--------------------- Processou NFE, situação: {texto_marcacao} ---------------------')

def reaplica_filtro_status(): 
    bot.PAUSE = 0.25
    logger.debug('--- Executando a função REAPLICA FILTRO STATUS')
    time.sleep(0.2)
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
    ahk.win_wait_active('debug_db_alltrips', title_match_mode= 2, timeout= 15)
    time.sleep(0.5)

    #* Clica no meio da tela, para garantir que está sem nenhuma outra tela aberta
    bot.click(960, 640)
    
    #* Inicia navamento até o campo "A1"
    bot.hotkey('CTRL', 'HOME') # Navega até o campo A1

    #* Navega até a coluna "STATUS" e abre o menu com as opções
    bot.press('RIGHT', presses= 6) # Navega até o campo "Status"
    bot.hotkey('ALT', 'DOWN') # Comando do excel para abrir o menu do filtro
    time.sleep(0.2)
    logger.info('--- Navegou até celula A1 e abriu o filtro do status ')
    ahk.win_activate('debug_db_alltrips', title_match_mode= 2)

    #* Procura pelo botão aplicar, e clica nele! Caso não encontre lança uma exceção.
    for i in range(0, 10):
        ahk.win_activate('debug_db_alltrips', title_match_mode= 2)
        time.sleep(1.5)

        #* Procura pelo botão "APLICAR"
        if procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png', continuar_exec= True):
            bot.click(procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png', continuar_exec= True))
            logger.info('--- Na tela do menu de filtro, clicou no botão "Aplicar" para reaplicar o filtro ')
            time.sleep(0.2)
            break
        
        #* Caso exceda o maximo de tentativas de encontrar o botão
        if i >= 9:
            logger.error('--- Não encontrou o botão APLICAR!')
            procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png', limite_tentativa= 1)
    
    #* Verifica se a tela "APLICAR FILTRO PARA TODOS" apareceu
    if procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png', continuar_exec= True, limite_tentativa= 3):
        bot.click(procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png', continuar_exec= True))
        logger.info('--- Clicou para visualizar o filtro de todos.')
    
    #* Concluiu a validação que o filtro está aplicado
    logger.success("--- Filtro da coluna status aplicado!")
    bot.hotkey('CTRL', 'HOME') # Navega até o campo A1

def extrai_txt_img(imagem, area_tela, porce_escala = 400):
    time.sleep(0.4)
    img = bot.screenshot('imagens/img_geradas/' + imagem, region=area_tela) # Captura uma screenshot da área especificada da tela
    logger.debug(F'--- Tirou print da imagem: {imagem} ----')

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
    ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
    texto_xml = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(168, 407, 250, 20))
    logger.debug(F'--- Item da nota: {texto}, texto que ainda ficou: {texto_xml}, tamanho do texto {len(texto_xml)}')

    # Verifica pelo tamanho do texto, se ainda ficou algum valor no campo "Itens do pedido"
    if len(texto_xml) > 5: 
        logger.warning('--- Itens XML ainda tem informação!')
        return False
    else:  # Caso fique vazio
        logger.info('--- Itens XML ficou vazio! saindo da tela de vinculação')
        bot.click(procura_imagem(imagem='imagens/img_topcon/confirma.png'))
        
        ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2, timeout= 30)
        while ahk.win_exists('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2):
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2)
            time.sleep(0.4)
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
                logger.debug('--- TopCompras abriu com o nome normal, prosseguindo.')
                return
            else:
                exit(bot.alert('TopCompras não encontrado.'))
        except (TimeoutError, OSError):
            logger.warning("Não encontrou o TopCompras nem a tela sem nome")
            return
    else:
        ahk.win_set_title(new_title= novo_nome, title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)
        logger.warning('--- Encontrou tela sem o nome, e realizou a correção!' )

def abre_planilha_navegador(link_planilha = alltrips):
    if link_planilha == alltrips: # Planilha original
        planilha = "db_alltrips.xlsx" 
    else: # Planilha de debug
        planilha = "debug_db_alltrips.xlsx"   

    logger.info(F'--- Iniciando função ABRE PLANILHA NAVEGADOR, abrindo: {planilha}')

    #* Verifica se a planilha já esta aberta
    if ahk.win_exists(planilha):
        logger.debug('--- Planilha já está aberta!')

        ativar_janela(planilha, 5)
        ahk.win_maximize(planilha, title_match_mode= 2)
        time.sleep(0.2)
        bot.hotkey('CTRL', 'F5') # Recarrega a planilha limpando o cache
        time.sleep(5)

    #* Verifica se a planilha realmente já recarrego
        for i in range (0, 30):
            ativar_janela(planilha, 5)
            time.sleep(0.5)
            
            if procura_imagem(imagem='imagens/img_planilha/txt_status.png', continuar_exec= True, limite_tentativa= 10):
                if procura_imagem(imagem='imagens/img_planilha/icone_nuvem.png', continuar_exec= True, limite_tentativa= 10):
                    break
            if i == 30:
                logger.erro('--- Planilha não carregou corretamente!')
                raise Exception('--- Planilha não carregou corretamente!')
            
        logger.info('--- Executou apenas um recarregamento')
        return True
    else:
        logger.info('--- Planilha fechada, abrindo uma nova execução!')
        os.system('taskkill /im msedge.exe /f /t')

    #* Garante que a planilha ORIGINAL não esteja aberta
    while ahk.win_exists("alltrips.xlsx", title_match_mode= 1): # Garante que a planilha não esteja aberta
        logger.info('--- Forçou o fechamento da planilha ORIGINAL do AllTrips')
        ahk.win_close('alltrips.xlsx', title_match_mode= 2)
        os.system('taskkill /im msedge.exe /f /t')
        time.sleep(1)
    
    logger.info(F'--- Realizando uma nova execução da planilha: {planilha}')    
    comando_iniciar = F'start msedge {link_planilha} -new-window -inprivate'
    os.system(comando_iniciar)
    time.sleep(5)
    
    #* Aguarda a planilha abrir no EDGE e maximiza
    for i in range (0, 10):
        time.sleep(1)
        ahk.win_activate(planilha, title_match_mode = 2)
        time.sleep(1)
        if ahk.win_is_active(planilha, title_match_mode = 2):
            ahk.win_maximize(planilha)
            time.sleep(2)
            logger.info('--- Planilha aberta e maximizada! procurando icone do EXCEL')
            if procura_imagem(imagem='imagens/img_planilha/icone_excel.png', continuar_exec= True):
                logger.info('--- Encontrou o icone do excel! procurando texto "STATUS" ')
                time.sleep(2)
                logger.success('--- Todas validações realizadas, planilha realmente aberta!')
                return True
        
        if i >= 9:
            logger.error('--- Planilha não abriu corretamente!')
            raise Exception('--- Planilha não abriu corretamente!')
        
        bot.press('F5')
    logger.success('--- Planilha aberta e maximizada.')
    

def msg_box(texto: str, tempo: int = 60):
    """Exibe uma dialog box temporaria utilizando AHK

    Args:
        texto (str): O texto que irá aparecer na mensagem
        tempo (int): Tempo até o fechamento em segundos
    """
    ahk.msg_box(text= texto, blocking= False)
    time.sleep(tempo)
    ahk.win_close('Message', title_match_mode= 2)

def verifica_horario():
    while True:
        hora_atual = datetime.now().time() # Obter o horário atual
        for i in range (0, 1):
            if i < 1:
                print('--- Verificando se passou das 23h')
                hora_inicio_pausa = datetime.strptime("23:30", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
                hora_final_pausa = datetime.strptime("23:59", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
            else:
                print('--- Verificando se é madrugada')
                hora_inicio_pausa = datetime.strptime("00:00", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
                hora_final_pausa = datetime.strptime("02:30", "%H:%M").time() # Definir o horário de inicio de referência (02:00)

            if hora_atual > hora_inicio_pausa and hora_atual < hora_final_pausa:
                logger.warning(F'--- São: {hora_atual}, aguardando 2 hora para tentar novamente.')
                msg_box(F"São: {hora_atual}, aguardando 2 hora para tentar novamente", 7200)
        else:
            return

def ativar_janela(nome_janela, timeout= 5):
    """ Tenta realizar a ativação de uma janela, e aguarda até ela estar aberta

    Args:
        nome_janela (_type_): Nome da janela que será aberta
        timeout (int, optional): Tempo em segundos que aguardará até a janela estar aberta. Valor padrão: 10.
    """
    logger.debug(F'--- Tentando ativar/abrir a janela: {nome_janela} ---' )
    ahk.win_activate(nome_janela, title_match_mode=2)
    ahk.win_wait_active(nome_janela, title_match_mode=2, timeout=timeout)
    time.sleep(0.25)

def move_telas_direita(tela:str):
    """ Move a {tela} para a esquerda

    Args:
        tela (str): nome da tela que será movida
    """    
    
    posicao = ahk.win_get_position(tela, title_match_mode = 2)

    if type(posicao) is not tuple:
        return

    if posicao[0] < -110:
        ahk.win_activate(tela, title_match_mode = 2)
    
        ahk.key_down('LShift')
        ahk.key_down('LWin')
        ahk.key_press('left')
        ahk.key_release('LShift')
        ahk.key_release('LWin')

        #ahk.win_maximize(tela, title_match_mode = 2)
    else:
        return


if __name__ == '__main__':
    bot.PAUSE = 0.6
    bot.FAILSAFE = False
    #print_erro()
    #msg_box("Teste", tempo = 1000)
    #abre_planilha_navegador(planilha_debug)
    #bot.alert("Executou")
    #reaplica_filtro_status()
    #verifica_ped_vazio()
    #corrige_nometela()
    marca_lancado()