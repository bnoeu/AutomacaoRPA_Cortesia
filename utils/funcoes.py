# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import cv2
import psutil
import pyperclip
import subprocess
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

def procura_imagem(imagem, limite_tentativa=8, area=(0, 0, 1920, 1080), continuar_exec=False, confianca=0.85, cinza=True):
    """Procura uma imagem na tela usando OCR e retorna suas coordenadas.

    Args:
        imagem (str): Caminho da imagem a ser encontrada
        limite_tentativa (int): Número máximo de tentativas. Default: 8
        area (tuple): Região da tela para buscar (x1,y1,x2,y2). Default: tela inteira
        continuar_exec (bool): Se True, continua execução mesmo sem encontrar. Default: False 
        confianca (float): Nível de confiança da busca (0-1). Default: 0.85
        cinza (bool): Se True, converte para escala de cinza. Default: True

    Returns:
        tuple: Coordenadas (x,y) onde encontrou a imagem, ou False se não encontrou
        
    Raises:
        OSError: Se houver erro de resolução da VM
        Exception: Se exceder limite de tentativas
    """
    pausa_img = 0.125
    confianca_min = 0.5
    pausa_max = 0.25
    
    for tentativa in range(limite_tentativa):
        try:
            posicao_img = bot.locateCenterOnScreen(
                imagem, 
                grayscale=cinza,
                confidence=confianca,
                region=area
            )
            
            if posicao_img:
                logger.debug(
                    f'Encontrou {imagem} em {posicao_img} '
                    f'(Tentativa: {tentativa}, Confiança: {confianca:.2f}, Pausa: {pausa_img})'
                )
                return posicao_img
                
        except OSError as e:
            logger.critical(f'Erro de resolução da VM: {e}')
            time.sleep(15)
            raise
            
        # Ajuste gradual dos parâmetros
        if confianca > confianca_min:
            confianca -= 0.02
        pausa_img = min(pausa_img + 0.025, pausa_max)
        time.sleep(pausa_img)

    if continuar_exec:
        logger.debug(
            f'{imagem} não encontrada com continuar_exec=True '
            f'(Tentativas: {limite_tentativa}, Confiança: {confianca})'
        )
        return False

    # Salva screenshot do erro
    time_atual = datetime.now().strftime('%Y%m%d_%H%M%S')
    caminho_erro = f'imagens/img_geradas/erros/erro_{time_atual}.png'
    bot.screenshot().save(caminho_erro)
    
    raise Exception(f"Falha ao procurar {imagem} após {limite_tentativa} tentativas")

def verifica_tela(nome_tela, manual=False):
    if ahk.win_exists(nome_tela):
        logger.info(F'--- A tela: {nome_tela} está aberta')
        ahk.win_activate(nome_tela, title_match_mode=2)
        return True
    elif manual is True:
        logger.info(F'--- A tela: {nome_tela} está fechada, Modo Manual: {True}, executando...')
        return False
    else:
        pass
        #exit(logger.error(F'--- Tela: {nome_tela} está fechada, saindo do programa.'))

def marca_lancado(texto_marcacao='texto_teste_marcacao', temp_inicial = ""):
    bot.PAUSE = 0.2

    logger.info(F'--- Abrindo planilha - MARCA_LANCADO, com parametro: {texto_marcacao}' )
    ativar_janela('debug_db', 30)
    time.sleep(0.4)
    bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.
    time.sleep(0.4)
    
    # Navega até o campo "Status"
    bot.hotkey('CTRL', 'HOME')
    bot.press('RIGHT', presses= 6)
    bot.press('DOWN')
    time.sleep(0.4)
    
    # Informa o texto recebido pela função e passa para a celula ao lado, para inserir a data
    bot.write(texto_marcacao)
    bot.press('RIGHT')
    hoje = datetime.now()
    hoje_formatado = hoje.strftime('%d/%m/%Y %H:%M:%S')
    bot.write(hoje_formatado)
    time.sleep(0.4)

    # Navega até a coluna "D. Insercao" para corrigir data
    bot.press('RIGHT', interval= 0.2)
    bot.press('F2', interval= 0.2)
    bot.press('ENTER', interval= 0.2)
    bot.press('UP', interval= 0.2)
    time.sleep(0.4)
    
    # Navega até a coluna da tentativa
    bot.press('RIGHT', presses= 1, interval= 0.1)

    # Copia a quantidade de tentativas atual
    for i in range (0, 5):
        # Verifica se já existe valor no campo
        time.sleep(0.2)
        bot.hotkey('ctrl', 'c', interval= 0.1)
        time.sleep(0.2)

        valor_copiado = pyperclip.paste()
        time.sleep(0.1)
        logger.debug(F'--- Coletando a tentativa e verificando o texto: {valor_copiado}' )
        if 'Recuperando' not in valor_copiado:
            logger.debug(F'--- Copia da tentativa deu certo! Texto coletado: {valor_copiado}' )
            break

        if i > 4:
            raise Exception("Não foi possivel copiar a quantidade de tentativas.")
        
    # Preenchimento da tentativa
    if valor_copiado != "":
        valor_copiado = int(valor_copiado)
        if valor_copiado > 0:
            valor_copiado += 1
            ativar_janela('debug_db', 30)
            bot.write(str(valor_copiado))
            time.sleep(0.1)
    else:
        # Caso o campo esteja vazio, significa que ainda não havia sido feito uma tentativa! Por isso, marca como 1º tentativa
        ativar_janela('debug_db', 30)
        bot.write("1")
        time.sleep(0.2)

    #* Caso precise informar o tempo que levou.
    if temp_inicial != "":
        bot.press('RIGHT', presses= 1, interval= 0.1)
        time.sleep(0.2)

        # Valida a medição de tempo que levou
        end_time = time.time()
        elapsed_time = end_time - temp_inicial
        #medicao_minutos = elapsed_time / 60
        bot.write(F"{round(elapsed_time)}")
        #print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")
        #logger.info(f"Tempo decorrido: {medicao_minutos:.2f} segundos")

    time.sleep(0.2)
    ativar_janela('debug_db', 30)

    bot.click(960, 640) # Clica no meio da planilha
    
    # Retorna a planilha para o modo "Somente Exibição (Botão Verde)"
    bot.hotkey('CTRL', 'HOME')
    reaplica_filtro_status() # Reaplica o filtro da coluna "Status"
    bot.hotkey('CTRL', 'HOME')
    logger.info(F'--------------------- Processou NFE, situação: {texto_marcacao} ---------------------')

def verifica_status_vazio():
    """No momento da aplicação do filtro, consulta para verificar se existe a opção "Vazio"

    Returns:
        boolean: True = Existe "Vazias" | False = Não existe "Vazias"
    """

    ativar_janela('debug_db_alltrips')

    if procura_imagem(imagem='imagens/img_planilha/txt_vazias.png', continuar_exec= True):
        return True
    else:
        return False

def verifica_existe_pendente():
    bot.PAUSE = 1
    # Verifica se existe mais alguma linha para ser lançada.

    ativar_janela('debug_db_alltrips')
    time.sleep(0.5)
    bot.hotkey('CTRL', 'HOME') # Navega até o campo A1

    for i in range (0, 2): # Avança para a terceira linha
        bot.press('DOWN') # Navega até o campo "Status"

    bot.hotkey('ctrl', 'c')
    time.sleep(0.4)
    valor_copiado = pyperclip.paste()
    tamanho_texto = len(valor_copiado)

    if tamanho_texto > 1:
        return True
    else:
        logger.info('--- Terceira linha da planilha está vazia! Provavelmente chegou ao final')
        return False

def reaplica_filtro_status(): 
    bot.PAUSE = 0.25

    for tentativa_filtro in range (0, 6):
        logger.debug(f'--- Executando a função REAPLICA FILTRO STATUS, tentativa: {tentativa_filtro}')

        ativar_janela('debug_db_alltrips')

        #* Clica no meio da tela, para garantir que está sem nenhuma outra tela aberta
        bot.click(960, 640)
        time.sleep(0.2)
        
        if verifica_existe_pendente() is False:
            return True

        #* Navega até a coluna "STATUS" e abre o menu com as opções
        #* Inicia navamento até o campo "A1"
        bot.hotkey('CTRL', 'HOME') # Navega até o campo A1
        bot.press('RIGHT', presses= 7, interval= 0.08) # Navega até o campo "Status"
        bot.press('LEFT') # Navega até o campo "Status"
        bot.hotkey('ALT', 'DOWN') # Comando do excel para abrir o menu do filtro
        logger.info('--- Navegou até celula A1 e abriu o filtro do status ')

        if verifica_status_vazio() is False:
            logger.info('--- Não existem notas c/ status "Vazias" ')
            bot.click(960, 640)
            time.sleep(0.2)
            return True

        if procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png', continuar_exec= True):
            bot.click(procura_imagem(imagem='imagens/img_planilha/bt_aplicar.png'))
            logger.info(f'--- Na tela do menu de filtro, clicou no botão "Aplicar" para reaplicar o filtro, tentativa: {tentativa_filtro}')

            #* Verifica se a tela "Outras pessoas também estão fazendo alterações" apareceu
            if procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png', limite_tentativa= 3, confianca= 0.73, continuar_exec= True):
                bot.click(procura_imagem(imagem='imagens/img_planilha/bt_visualizar_todos.png', continuar_exec= True))
                logger.info('--- Clicou para visualizar o filtro de todos.')
        
    
            #* Concluiu a validação que o filtro está aplicado
            logger.success("--- Filtro da coluna status aplicado!")
            bot.hotkey('CTRL', 'HOME') # Navega até o campo A1

            return True
    else:
        raise Exception("Falhou ao aplicar o filtro na coluna de status!")

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

        # Fecha a tela de confirmação de pedido vinculado.
        ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode = 2, timeout= 30)
        bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', limite_tentativa= 10))
        ahk.win_wait_close('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
        logger.info('--- Encerrado a função verifica pedido vazio!')
        return True

def corrige_nometela(novo_nome = "TopCompras"):  
    """_summary_ Verifica se tem alguma tela com o nome vazio, e realiza a correção.

    Args:
        novo_nome (str, optional): _description_. Nome da tela que quero tentar corrigir, Defaults to "TopCompras".
    """

    try: # Verifica se abriu alguma tela sem o nome.
        ahk.win_wait(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, timeout= 8)
    except (TimeoutError, OSError): # Apresenta Timeout caso esteja aberto com o nome normal.
        try: # Verifica se REALMENTE abriu com o nome normal
            if ahk.win_wait(novo_nome, title_match_mode= 1, timeout= 6):
                logger.debug('--- TopCompras abriu com o nome normal, prosseguindo.')
                return
            else:
                exit(bot.alert('TopCompras não encontrado.'))
        except (TimeoutError, OSError):
            logger.warning("Não encontrou o TopCompras nem a tela sem nome")
            return
    else:
        if ahk.win_exists(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1):
            time.sleep(1)
            ahk.win_set_title(new_title= novo_nome, title= ' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1, detect_hidden_windows= True)
            logger.warning('--- Encontrou tela sem o nome, e realizou a correção!' )

def valida_carregamento_planilha(nome_planilha = ""):
    # Verifica se a planilha realmente já recarregou
    for i in range (0, 3):
        ativar_janela(nome_planilha, 5)
        time.sleep(0.4)
        
        if procura_imagem(imagem='imagens/img_planilha/icones_inferior.png', limite_tentativa= 10, continuar_exec= True):
            logger.success('--- Todas validações realizadas, planilha realmente aberta!')
            return True 
        else:
            ahk.win_maximize(nome_planilha)
            time.sleep(0.4)

    else:
        logger.erro('--- Planilha não carregou corretamente!')
        raise Exception('--- Planilha não carregou corretamente!')


def abre_planilha_navegador(link_planilha = alltrips):
    for i in range (0, 3):
        try:
            # Identifica qual nome_planilha será utilizada.
            if link_planilha == alltrips: # Planilha original
                nome_planilha = "db_alltrips.xlsx" 
            else: # Planilha de debug
                nome_planilha = "debug_db_alltrips.xlsx"   

            logger.info(F'--- Iniciando função ABRE PLANILHA NAVEGADOR, abrindo: {nome_planilha}')

            #* Verifica se a planilha já esta aberta
            if ahk.win_exists(nome_planilha):
                logger.debug('--- Planilha já está aberta! Executando apenas um recarregamento')

                ativar_janela(nome_planilha)
                ahk.win_maximize(nome_planilha, title_match_mode= 2)
                time.sleep(0.5)
                bot.hotkey('CTRL', 'F5') # Recarrega a planilha limpando o cache

                valida_carregamento_planilha(nome_planilha)

                logger.info('--- Executou apenas um recarregamento')
                return True
            else:
                #subprocess.run(["cmd", "/c", "taskkill /im msedge.exe /f /t 2>nul"], shell=True)
                matar_autohotkey(nome_exec= "msedge.exe")
                logger.info('--- Planilha (EDGE) fechada, abrindo uma nova execução da planilha: {nome_planilha}')
                
                comando_iniciar = f'start msedge {link_planilha} -new-window -inprivate'
                subprocess.run(["cmd", "/c", comando_iniciar], shell=True)
                time.sleep(2)
                

                #* Aguarda a planilha abrir no EDGE e maximiza
                valida_carregamento_planilha(nome_planilha)
                return True
            
        except Exception as e:
            logger.error(f'--- Erro na função de abertura da planilha: {e}!')
            
            if i > 2:
                return False

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
        hora_atual = datetime.now().time()
        intervalos_pausa = [
            (datetime.strptime("23:20", "%H:%M").time(), datetime.strptime("23:59", "%H:%M").time()),
            (datetime.strptime("00:00", "%H:%M").time(), datetime.strptime("04:30", "%H:%M").time())
        ]

        for inicio, fim in intervalos_pausa:
            logger.debug(f'--- São: {hora_atual}, verificando intervalo: {inicio} até {fim}')
            if inicio < hora_atual < fim:
                logger.warning(f'São: {hora_atual}, aguardando 2 horas para tentar novamente.')
                msg_box(f"São: {hora_atual}, aguardando 2 horas para tentar novamente", 7200)
                break
        else:
            logger.debug('--- Horário validado! Pode prosseguir com o lançamento.')
            break

''' #! Substituido pela logica a cima.
def verifica_horario():
    validou_horarios = 0
    while True:
        hora_atual = datetime.now().time() # Obter o horário atual
        for i in range (2):
            if i < 1:
                print('--- Verificando se passou das 23h')
                hora_inicio_pausa = datetime.strptime("23:20", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
                hora_final_pausa = datetime.strptime("23:59", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
                logger.debug(F'--- São: {hora_atual}, Verificando se passou das 23h: {hora_inicio_pausa} vs {hora_final_pausa}')
            else:
                hora_inicio_pausa = datetime.strptime("00:00", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
                hora_final_pausa = datetime.strptime("04:30", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
                logger.debug(F'--- São: {hora_atual}, Verificando se é madrugada: {hora_inicio_pausa} vs {hora_final_pausa}')

            if hora_atual > hora_inicio_pausa and hora_atual < hora_final_pausa:
                logger.warning(F'--- São: {hora_atual}, aguardando 2 hora para tentar novamente.')
                msg_box(F"São: {hora_atual}, aguardando 2 hora para tentar novamente", 7200)
                validou_horarios = 0
            else:
                validou_horarios += 1

        if validou_horarios >= 2:
            logger.debug('--- Horario validado! Pode prosseguir com o lançamento.')
            break
'''

def ativar_janela(nome_janela= "TopCompras", timeout= 8):
    """ Tenta realizar a ativação de uma janela, e aguarda até ela estar aberta

    Args:
        nome_janela (_type_): Nome da janela que será aberta
        timeout (int, optional): Tempo em segundos que aguardará até a janela estar aberta. Valor padrão: 10.
    """

    logger.debug(F'--- Tentando ativar/abrir a janela: {nome_janela} ---' )
    ahk.win_activate(nome_janela, title_match_mode=2)
    time.sleep(0.2)
    ahk.win_wait_active(nome_janela, title_match_mode=2, timeout=timeout)
    time.sleep(0.2)

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


def matar_autohotkey(nome_exec = "", timeout=5):
    """
    Encerra todos os processos de forma segura.
    - Tenta matar via psutil (mais rápido e confiável).
    - Se não encontrar, tenta usar taskkill com timeout.
    """
    
    # 1️⃣ Tentativa via psutil (mais confiável e sem travamento)
    encontrado = False
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower() == nome_exec:
            try:
                proc.kill()
                encontrado = True
            except Exception as e:
                print(f"⚠️ Erro ao encerrar PID {proc.pid}: {e}")
    
    if encontrado:
        print(f"✅ Processo: {nome_exec} encerrado via psutil.")
        return
    
    # 2️⃣ Se não encontrou, tenta via taskkill (com timeout)
    try:
        subprocess.run(
            ["taskkill", "/im", nome_exec, "/f", "/t"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            timeout=timeout
        )
        print(f"✅ Processo: {nome_exec} encerrado via taskkill.")
    except subprocess.TimeoutExpired:
        print("⏳ taskkill demorou demais, processo ignorado.")
    except Exception as e:
        print(f"⚠️ Erro ao executar taskkill: {e}")


if __name__ == '__main__':
    bot.PAUSE = 0.6
    bot.FAILSAFE = False
    tempo_inicial = time.time() # Calculo do tempo de execução das funções
    
    corrige_nometela("TopCompras")
    #corrige_nometela('TopCon')
    #verifica_horario()
    #matar_autohotkey(nome_exec= "msedge.exe")
    #exit()
    #abre_planilha_navegador(alltrips)
    #verifica_existe_pendente()
    #valida_carregamento_planilha('db_alltrips.xlsx')
    #reaplica_filtro_status()
    #ativar_janela()
    #verifica_status_vazio()
    #verifica_horario()
    #print_erro()
    #msg_box("Teste", tempo = 1000)
    #abre_planilha_navegador(planilha_debug)
    #bot.alert("Executou")
    #verifica_ped_vazio()
    #corrige_nometela()
    #marca_lancado(temp_inicial= tempo_inicial, texto_marcacao= "teste_2025_07_25")

    # Calculo do tempo de execução das funções
    end_time = time.time()
    elapsed_time = end_time - tempo_inicial
    medicao_minutos = elapsed_time / 60
    print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")