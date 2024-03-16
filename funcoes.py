# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import datetime
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
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"


def procura_imagem(imagem, limite_tentativa=4, area=(0, 0, 1920, 1080), continuar_exec=False):
    tentativa = 0
    #print(F'--- Tentando encontrar "{imagem}", tentativa: {tentativa}', end='... ')
    while tentativa < limite_tentativa:
        tentativa += 1
        time.sleep(0.5)
        # Identifica posição do logo TOPCON
        posicao_img = bot.locateOnScreen(imagem, grayscale=True, confidence=0.88, region=area)
        if posicao_img is not None:
            #print(F'encontrou a imagem na posição: {posicao_img}')
            break
    if (continuar_exec is True) and (posicao_img is None):
        print('não achou imagem, continuando execução pois o parametro "continuar_exec" está habilitado')
        return False
    if tentativa >= limite_tentativa:
        exit(bot.alert(
            text=F'não foi possivel encontrar: {imagem}', title='Erro!', button='Fechar'))
    return posicao_img

def alteracao_filtro():
    time.sleep(1)
    if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8) is not False:
        print('--- Já está filtrado, continuando!')
    else:
        print('--- Não está filtrado, executando o filtro!')
        bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1463, 419, 100, 100)))
        bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png'))
        bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
        bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))

def verifica_tela(nome_tela, manual=False):
    if ahk.win_exists(nome_tela):
        print(F'--- A tela: {nome_tela} está aberta')
        ahk.win_activate(nome_tela)
        return True
    elif manual is True:
        print(
            F'--- A tela: {nome_tela} está fechada, Modo Manual: {True}, executando...')
        return False
    else:
        exit(print(F'--- Tela: {nome_tela} está fechada, saindo do programa.'))

def marca_lancado(texto_marcacao='Lancado'):
    ahk.win_activate('db_alltrips')
    time.sleep(1)
    if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True, limite_tentativa=3) is not False:
        bot.click(procura_imagem(imagem='img_planilha/botao_exibicaoverde.png'))
        bot.click(procura_imagem(imagem='img_planilha/botao_iniciaredicao.png'))
        if procura_imagem(imagem='img_planilha/txt_modificada.png', continuar_exec=True, limite_tentativa=2) is not False:
            bot.click(procura_imagem(imagem='img_planilha/bt_sim.png'))
        procura_imagem(imagem='img_planilha/botao_edicao.png')
        bot.doubleClick(1494, 508)
        bot.write(texto_marcacao)
        bot.press('RIGHT')
        hoje = datetime.date.today()
        bot.write(str(hoje))
        bot.press("ENTER")
        time.sleep(1)
        if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8) is not False:
            bot.click(procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, limite_tentativa=8))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
        else:
            print('--- Não está filtrado, executando o filtro!')
            bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1463, 419, 100, 100)))
            bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png'))
            bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
            bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
    else:
        print('Não achou o botao de edição')

def coleta_planilha():
    ahk.win_activate('db_alltrips')
    time.sleep(1)
    if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True, limite_tentativa=2) is False:
        bot.click(procura_imagem(imagem='img_planilha/botao_edicao.png'))
        bot.click(procura_imagem(imagem='img_planilha/botao_exibicao.png'))
        procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', limite_tentativa= 8)
    else:
        print('--- Já está no modo de edição, continuando processo')
        # Validação se houve novo valor inserido
        time.sleep(1)
    
    alteracao_filtro()

    # * Coleta os dados da linha atual
    dados_planilha = []
    print('--- Copiando dados e formatando')
    bot.click(100, 510)  # Clica na primeira linha
    bot.PAUSE = 0.3
    for n in range(0, 7, 1):  # Copia dados dos 6 campos
        while True:
            bot.hotkey('ctrl', 'c')
            if 'Recuperando' in ahk.get_clipboard():
                print('Tentando copiar novamente')
                time.sleep(0.2)
            else:
                break
        dados_planilha.append(ahk.get_clipboard())
        bot.press('right')
    return dados_planilha

def acoes_planilha():
    validou_xml = False
    while validou_xml is False:
        # * Trata os dados coletados em "dados_planilha"
        dados_planilha = coleta_planilha()
        chave_xml = dados_planilha[4]
        # * -------------------------------------- Lançamento Topcon --------------------------------------
        bot.PAUSE = 1  # Pausa padrão do bot
        time.sleep(1)
        #! Tela de lançamento, necessaria apenas se estiver em modo de teste
        menu_inicio = bot.confirm(
            text='Deseja iniciar o lançamento?', title='Atenção', buttons=['Iniciar', 'Fechar'])
        if menu_inicio == 'Fechar':
            exit(print('Fechando programa...'))
        print('--- Iniciando lançamento ----')
        time.sleep(1)
        tela_compras = gw.getWindowsWithTitle('TopCompras')[0]
        tela_compras.maximize()
        tela_compras.activate()  # Ativa a tela de compras
        ahk.win_activate('TopCompras')
        if tela_compras.isMaximized:
            print('Tela compras está maximizada! Iniciando o programa')
        else:
            exit(print('Tela compras não abriu... Fechando script'))
        # Processo de lançamento
        time.sleep(1)
        bot.press('F2')
        bot.press('F3')
        bot.press('F3')
        bot.click(558, 235)  # Clica dentro do campo para inserir a chave XML
        time.sleep(1)
        bot.write(chave_xml)
        bot.press('ENTER')
        time.sleep(3)
        ahk.win_wait_active('TopCompras')
        while procura_imagem(imagem='img_topcon/naorespondendo.png', limite_tentativa=2, continuar_exec=True) is not True:
            time.sleep(2)
            print('Aguardando topvoltar')
        tentativa = 0
        while tentativa < 10:
            print(F'Tentativa: {tentativa}')
            if procura_imagem(imagem='botao_sim.jpg', limite_tentativa=2, continuar_exec=True) is not False:
                bot.click(procura_imagem(imagem='botao_sim.jpg',
                          limite_tentativa=2, continuar_exec=True))
                return dados_planilha, (validou_xml is True)
            # Verifica se encontrou o erro de NFE lançada
            elif procura_imagem(imagem='chave_invalida.png', limite_tentativa=2, continuar_exec=True) is not False:
                print('--- Nota já lançada, marcando planilha!')
                bot.press('ENTER')
                marca_lancado()
                break
            # Verifica se encontrou o erro de NFE lançada
            elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', limite_tentativa=2, continuar_exec=True) is not False:
                bot.press('ENTER')
                exit(print('--- XML ainda não baixado!!'))
                # TODO ---- O que fazer nestes casos?
                break
            tentativa += 1
        if tentativa >= 10:
            exit('Rodou 10 verificações e não achou nenhuma tela, aumentar o tempo')
