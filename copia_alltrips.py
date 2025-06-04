# Para utilização na Cortesia Concreto.
# -*- Criado por Bruno da Silva Santos. -*-

import time
import pyautogui as bot
from datetime import datetime
from utils.funcoes import ahk as ahk
from utils.configura_logger import get_logger
from utils.funcoes import procura_imagem, abre_planilha_navegador, ativar_janela, reaplica_filtro_status


# --- Definição de parametros
chave_xml = ""
powerapps_id = ""
logger = get_logger("script1")
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg"


def abre_menu_pesquisa(valor_pesquisa = ""):
    # Abre o menu de pesquisa, e pesquisa pelo valor informado
    bot.PAUSE = 0.6

    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
    time.sleep(2)
    bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.
    time.sleep(2)

    # Abre o menu de pesquisa
    logger.info(f'--- Abrindo o menu de pesquisa na planilha para procurar o valor: {valor_pesquisa}')
    bot.press('ALT')
    bot.press('C')
    bot.press('F')
    bot.press('D')
    bot.press('F')

    # Insere a ultima chave copiada da planilha de debug
    logger.info(F'--- Digitando o valor: {valor_pesquisa}')
    bot.write(valor_pesquisa, interval= 0.01)
    bot.press('ENTER', presses= 2)

    # Fecha o menu de pesquisa
    bot.press('ESC')
    bot.press('ALT', presses= 2)
    logger.info('--- Fechou o menu de pesquisa da planilha')


def encontra_ultimo_xml(ultimo_xml = '', powerapps_id = ''):
    bot.PAUSE = 2.2
    tentativa = 0
    while True:
        tentativa += 1
        if tentativa > 5:
            raise Exception("Excedeu o limite de tentativas de encontrar o ultimo XML!")

        logger.info(F'--- Iniciando a navegação até a ultima chave XML: {ultimo_xml}')
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
        ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 1, timeout= 5)
        try:
            ahk.win_wait_active('db_alltrips.xlsx', title_match_mode= 1, timeout= 10)
        except TimeoutError:
            logger.warning('--- Planilha não encontrada!')
            return False

        # Navega até o campo da data, e organiza do menor para o maior.
        bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.
        bot.hotkey('CTRL', 'HOME') # Move a navegação até a celula A1
        logger.info('--- Navegou até a coluna/celula A1')
        bot.press('RIGHT', presses= 8, interval= 0.05) # Navega até o campo "D. Insercao"]
        logger.info('--- Navegou até a D. Inserção')
        bot.hotkey('ALT', 'DOWN') # Abre o menu do filtro
        time.sleep(3)
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
        time.sleep(0.5)
        bot.click(procura_imagem(imagem='imagens/img_planilha/icone_organiza_A_Z.png', continuar_exec= True, limite_tentativa= 30)) # Clica no botão "organizar do mais antigo ao mais novo"
        logger.info('--- Organizou a planilha da forma "da menor para a maior" ')
        time.sleep(0.5)
        ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
        bot.click(960, 630) # Clica no meio da planilha para "ativar" a navegação dentro dela.

        abre_menu_pesquisa(ultimo_xml)

        # Verifica se realmente chegou no ultimo XML
        bot.hotkey('ctrl', 'c')
        chave_encontrada = ahk.get_clipboard()
        if chave_encontrada == ultimo_xml:
            logger.info(F'--- Concluido a navegação até a ultima chave XML: {ultimo_xml}, validando PowerApps ID')

            #* Realiza uma validação também pelo PowerApps ID
            #* Caso o antigo seja = #NOME?, não realiza a validação.
            if powerapps_id != "#NOME?":
                bot.press('RIGHT')
                time.sleep(0.2)
                bot.hotkey('ctrl', 'c')
                novo_powerapps_id = ahk.get_clipboard()

                if novo_powerapps_id == powerapps_id:
                    logger.info(f'--- PowerApps ID também validado! valor: {novo_powerapps_id}, realmente está na ultima linha.')
                    bot.press('LEFT')
                    return True
                else:
                    logger.info(f'--- PowerApps ID coletado: {novo_powerapps_id} é diferente do da ultima nota: {powerapps_id}')
                    abre_menu_pesquisa(powerapps_id)
                    bot.press('LEFT')                    
            else:
                return True
        else:
            logger.warning(F'--- Ops... não está na ultima chave {ultimo_xml}, navegando novamente.')
            raise TimeoutError


def valida_nova_chave_inserida(tentativa):
    bot.PAUSE = 2.2
    tempo_pausa = tentativa * 1800  # Multiplica a tentativa por 30 minutos, como são 4, o maximo é 2 horas
    logger.info('--- Verificando se existe uma nova chave NFE.')

    ahk.win_activate('db_alltrips.xlsx', title_match_mode= 1)
    time.sleep(0.8)
    bot.press('DOWN') # Navega até a proxima linha após a ultima chave.
    time.sleep(0.2)

    while True: # Executa o processo de copia dos dados
        bot.hotkey('ctrl', 'c')
        if 'Recuperando' in ahk.get_clipboard():
            logger.info('--- Tentando copiar novamente.')
            time.sleep(0.4)
        else:
            logger.info('--- Dado copiado com sucesso, realizando avaliação.')
            valor_copiado = ahk.get_clipboard()
            break

    #* Executa a validação dos dados copiados
    if valor_copiado == "": # 1. Caso o campo esteja vazio, significa que ainda não foram inseridas novas notas, e para o processo.
        logger.info(F'--- Valor copiado está vazio! Aguardando {tempo_pausa / 60} minutos antes de tentar novamente')
        time.sleep(tempo_pausa)
        return False
    elif len(valor_copiado) < 20 or len(valor_copiado) > 44:
        logger.warning(F'--- Valor copiado é invalido: {valor_copiado}')
        return False
    else: # 2. Caso o campo esteja com uma chave XML nova, prossegue.
        logger.info(F'--- Uma nova chave foi inserida: {valor_copiado}, saindo do loop')
        return True


def copia_dados():
    bot.PAUSE = 2.2
    dados_copiados = ""

    # Inicia o processo de seleção dos dados
    ativar_janela("db_alltrips.xlsx")
    logger.info('--- Iniciando o processo de seleção dos dados novos')
    time.sleep(0.5)
    bot.press('LEFT', presses= 4) # Navega até a coluna "RE"
    time.sleep(0.5)
    ahk.key_down('Shift') # Segura a tecla SHIFT
    time.sleep(0.5)
    ahk.key_down('Control') # Segura a tecla CTRL
    time.sleep(0.5)
    ahk.key_press('down') # Com shift + ctrl pressionado, navega até a ultima linha da planilha
    time.sleep(0.5)
    ahk.key_press('right') # Avança para a ultima coluna
    logger.info('--- Pressionou SHIFT e CONTROL, indo até a ultima coluna preenchida')
    for i in range (0, 9):
        time.sleep(0.5)
        ahk.key_up('Shift')
        time.sleep(0.5)
        ahk.key_up('Control')
        time.sleep(0.5)
        bot.hotkey('ctrl', 'c')
        time.sleep(0.5)
        dados_copiados = ahk.get_clipboard()
        time.sleep(0.5)

         #* Script de validação original
        if "/2025" in dados_copiados:
            logger.info('--- Encontrou "/2025" que indica os dados da coluna "D. Inserção" nos dados copiados!')
            return dados_copiados

        if i >= 6:
            if ("/" in dados_copiados) or ("/2025" in dados_copiados) or ("," in dados_copiados): # Verifica se os dados foram copiados com sucesso
                logger.success('--- Novos dados copiados com sucesso da planilha db_alltrips')
                print('--- Novos dados copiados com sucesso da planilha db_alltrips')
                return dados_copiados
        else:
            ahk.key_down('Shift') # Segura a tecla SHIFT
            time.sleep(0.5)
            ahk.key_down('Control') # Segura a tecla CTRL
            time.sleep(0.5)
            ahk.key_press('right') # Avança para a ultima coluna
            time.sleep(0.5)
            logger.debug('--- Ainda não encontrou a coluna das datas! Tentando novamente')
            ativar_janela("db_alltrips.xlsx")

        if i >= 9: # Verifica se excedeu o limite de tentativas de copiar os dados.
            logger.error('--- Excedeu o limite de tentativas de copiar os dados, soltando SHIFT e CONTROL')
            ahk.key_up('Shift')
            time.sleep(1)
            ahk.key_up('Control')
            raise Exception("Excedeu o limite de tentativas de copiar os dados, soltando SHIFT e CONTROL")


def cola_dados(dados_copiados = "TESTE"):
    bot.PAUSE = 2.5
    
    logger.info('--- Acessando a planilha de debug para COLAR os dados!')
    abre_planilha_navegador(planilha_debug)
    time.sleep(1)
    ativar_janela('debug_db_alltrips.xlsx')
    
    bot.hotkey('CTRL', 'HOME') # Navega até a celula A1.
    bot.press('DOWN', presses= 2) # Proxima linha que deveria estar sem informação.
    logger.info('--- Navegou até a proxima linha sem informações')

    bot.press('ALT') # Abre o menu para navegação via teclas
    time.sleep(0.8)
    bot.press('C') # Vai até a opção "Inicio"
    time.sleep(0.8)
    bot.press('V') # Abre o menu de "Colar"
    time.sleep(0.8)
    bot.press('V') # Seleciona a opção "Colar somente valores"
    time.sleep(3)

    # Realiza o fechamento da planilha com os dados originais. 
    logger.info('--- Dados colados com sucesso! Fechando a planilha original.')

    for i in range (0, 15):
        logger.info("--- Fechando a planilha do banco ORIGINAL antes de prosseguir.")
        time.sleep(0.4)
        ahk.win_kill('db_alltrips.xlsx', title_match_mode= 1) # Força o fechamento da planilha com o banco puro.
        if ahk.win_exists('db_alltrips.xlsx', title_match_mode= 1) is False:
            break
    else:
        logger.error('--- Não conseguiu fechar a planilha original')
        raise Exception("Não conseguiu fechar a planilha original")

    reaplica_filtro_status()


def verifica_quatro_dias(dados_copiados):
    """ Compara os dados copiados e verifica se consta alguma nota que a data é de quatro dias atrás.

    Args:
        dados_copiados (_type_): Recebe os dados coletados

    Raises:
        Exception: "Dia "{quatro_dias_antes}" está nos dados copiados"
    """

    dia_mes_atual = datetime.now() # Coleta a data atual, para validar se os dados são novos.
    quatro_dias_antes = dia_mes_atual.day - 4
    quatro_dias_antes = F"{quatro_dias_antes}/25"

    if quatro_dias_antes in dados_copiados:
        logger.debug(F'Verificando se o dia: {quatro_dias_antes} está nos dados copiados...')
        raise Exception(F'Dia "{quatro_dias_antes}" está nos dados copiados.')
    else:
        logger.info(F'Não encontrou: {quatro_dias_antes} nos dados copiados, os dados são novos!')


def main(ultimo_xml = chave_xml, powerapps_id = powerapps_id):
    bot.PAUSE = 2.2
    logger.info('Iniciando função COPIA BANCO ( COPIA ALL TRIPS)')

    #* Abre a planilha do db_alltrips (banco original)
    for tentativa in range(0, 6):
        abre_planilha_navegador()
        encontra_ultimo_xml(ultimo_xml = ultimo_xml, powerapps_id = powerapps_id)

        if valida_nova_chave_inserida(tentativa) is True:
            dados_copiados = copia_dados()
            print(dados_copiados)
            if dados_copiados != "":
                break
    else:
        raise Exception(F"--- Falhou as: {tentativa} tentativas da task COPIA ALLTRIPS")

    #* Verifica se consta alguma nota que a data é de quatro dias atrás.
    verifica_quatro_dias(dados_copiados)

    #* Libera as teclas para evitar problemas.
    ahk.key_up('Shift')
    ahk.key_up('Control')
    ahk.key_release('Shift') # Segura a tecla SHIFT
    ahk.key_release('Control')

    # Força o fechamento da planilha ORIGINAL
    ahk.win_kill('db_alltrips.xlsx', title_match_mode= 1)
    time.sleep(2)

    colou_dados = cola_dados(dados_copiados)
    if colou_dados is True:
        return True

if __name__ == '__main__':
    ultimo_xml = "35250633039223000979550010003927081074547429"
    powerapps_ultima_nfe = "ncHouHl1Y4I"
    
    cola_dados()
    exit()
    main(ultimo_xml= ultimo_xml, powerapps_id= powerapps_ultima_nfe)
    exit(bot.alert("Terminou"))
    

    #abre_menu_pesquisa("35250533039223000979550010003925631934762896")
    #abre_planilha_navegador()
    #encontra_ultimo_xml(ultimo_xml = ultimo_xml, powerapps_id= powerapps_ultima_nfe)
    #dados = copia_dados()
    #print(dados)