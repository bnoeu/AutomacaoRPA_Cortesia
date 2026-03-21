# -*- Funções para interação com TopCompras. -*-

import time
import pyautogui as bot
from utils.funcoes import ahk, procura_imagem, ativar_janela, marca_lancado
from utils.configura_logger import get_logger
from datetime import datetime, timedelta
from abre_topcon import main as abre_topcon

logger = get_logger("topcompras_handler", print_terminal=True)


def verifica_incrementa_data(data_antiga_str: str = "03/06/2025", qtd_incremento: int = 1) -> bool:
    """
    Verifica e incrementa a data se necessário.
    
    Args:
        data_antiga_str (str): Data anterior em formato DD/MM/YY.
        qtd_incremento (int): Quantidade de dias a incrementar.
        
    Returns:
        bool: True se a data foi aceita, False caso contrário.
    """
    for i in range(0, 3):
        time.sleep(3)
        if ahk.win_exists("Topsys", title_match_mode=2):
            logger.warning(f'--- Precisa mudar a data, inserindo a data de hoje: {data_antiga_str}')
            ativar_janela('Topsys', 70)
            time.sleep(1)
            bot.press('ENTER')
            time.sleep(1)
            break
        else:
            return True
    
    # Converter para datetime, assumindo que '25' é o ano 2025
    data_dt = datetime.strptime(data_antiga_str, "%d/%m/%y")

    # Adiciona dias
    nova_data = data_dt + timedelta(days=qtd_incremento)

    # Converte de volta para string
    nova_data_str = nova_data.strftime("%d/%m/%Y")

    # Insere a data ajustada
    ativar_janela('TopCompras', 70)
    time.sleep(0.4)
    bot.write(nova_data_str)
    time.sleep(0.4)
    bot.press('ENTER')
    time.sleep(0.4)
    
    return False


def preenche_data(data_formatada: str = "") -> None:
    """
    Realiza validação e preenchimento da data no TopCompras.
    
    Args:
        data_formatada (str): Data formatada em DD/MM/YY.
        
    Raises:
        Exception: Se não conseguir preencher a data após múltiplas tentativas.
    """

    logger.info('--- Realizando validação/alteração da data')
    ativar_janela('TopCompras', 70)
    time.sleep(0.2)

    hoje = datetime.today().strftime("%d%m%y")  # dd/mm/YY
    logger.info(f'--- Inserindo a data coletada: {data_formatada} e apertando ENTER')
    bot.write(data_formatada, interval=0.025)
    bot.press('ENTER')
    time.sleep(2)

    qtd_incremento = 0
    while verifica_incrementa_data(data_formatada, qtd_incremento) is not True:
        time.sleep(0.25)

        if qtd_incremento > 4:
            if ahk.win_exists("Topsys", title_match_mode=2):
                ahk.win_close('Topsys', title_match_mode=2)
            marca_lancado(texto_marcacao='data_limite_ultrapassada')
            raise Exception(f'Tentou preencher a data por {qtd_incremento} vezes sem sucesso!')
        qtd_incremento += 1

    ativar_janela('TopCompras', 70)

    # Caso o sistema informe que a data deve ser maior/igual a data inserida acima
    logger.info('--- Verificando se apareceu data inválida')
    if procura_imagem('imagens/img_topcon/data_invalida.png', continuar_exec=True, limite_tentativa=2):
        logger.warning(f'--- Precisa mudar a data, inserindo a data de hoje: {hoje}')
        ahk.win_close("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode=2)
        time.sleep(0.5)        
        bot.write(f"{hoje}")
        bot.press('enter')
        time.sleep(1)
    else:
        logger.info('--- Não foi necessário alterar a data!')

    try:  # Aguarda a tela de erro do TopCon
        ahk.win_wait('Topsys', title_match_mode=2, timeout=2)
    except TimeoutError:
        return
    else:
        if ahk.win_exists('Topsys', title_match_mode=2):
            ahk.win_activate('Topsys', title_match_mode=2)
            logger.warning('--- Precisa mudar a data')
            bot.press('enter')          
            bot.write(f"{hoje}")
            bot.press('enter')
            time.sleep(0.4)


def preenche_filial_estoque(filial_estoq: str) -> None:
    """
    Preenche o campo de filial de estoque no TopCompras.
    
    Args:
        filial_estoq (str): Código da filial de estoque.
    """
    logger.info('--- Preenchendo filial de estoque')
    ativar_janela('TopCompras')
    time.sleep(0.4)

    # Coleta a posição do TXT e faz um clique relativo
    posicao_texto = procura_imagem(imagem='imagens/img_topcon/txt_filial_estoque.png')
    bot.click(posicao_texto.x + 296, posicao_texto.y)
    time.sleep(0.4)

    bot.write(filial_estoq, interval=0.025)
    time.sleep(0.2)
    bot.press('TAB', presses=2, interval=0.2)  # Confirma a informação da nova filial de estoque


def valida_transportador(cracha_mot: str = "112842") -> bool:
    """
    Valida e preenche o transportador no TopCompras.
    
    Args:
        cracha_mot (str): Crachá do motorista.
        
    Returns:
        bool: True se validado com sucesso, False caso contrário.
    """
    logger.info(f'--- Preenchendo transportador: {cracha_mot}')
    ahk.win_activate('TopCompras', title_match_mode=2)
    time.sleep(0.2)

    # Move a tela para encontrar o campo do transportador
    bot.click(1907, 970, 2)
    # Clique relativo a posição do campo "Transportador: RE"
    onde_achou = procura_imagem(imagem='imagens/img_topcon/txt_transportador.png')
    bot.click(onde_achou[0] + 190, onde_achou[1])

    # Verifica se o TAB realmente navegou até o campo "RE: 0"
    tentativa_achar_camp_re = 0
    while procura_imagem(imagem='imagens/img_topcon/campo_re_0.png', continuar_exec=True, 
                        limite_tentativa=2, confianca=0.74) is False:
        logger.info(f'Tentativa: {tentativa_achar_camp_re}')
        time.sleep(0.4)
        tentativa_achar_camp_re += 1
        if tentativa_achar_camp_re >= 10:
            logger.info('--- Limite de tentativas de achar o campo "RE", reabrindo topcompras e reiniciando o processo.')
            time.sleep(0.5)
            abre_topcon()
            return True
    else:
        logger.info('--- Campo RE habilitado, preenchendo.')
        # Preenche o campo do transportador e verifica se aconteceu algum erro
        bot.press("Backspace", presses=6)
        bot.write(cracha_mot, interval=0.08)  # ID transportador
        bot.press('enter')

    logger.info('--- Aguardando validar o campo do transportador')
    ahk.win_activate('TopCompras', title_match_mode=2)
    if procura_imagem(imagem='imagens/img_topcon/transportador_incorreto.png', 
                     continuar_exec=True, limite_tentativa=2) is not False:
        logger.info('--- Transportador incorreto!')
        bot.press('ENTER')
        bot.press('F2')
        marca_lancado(texto_marcacao='RE_Invalido')
        return False
    else:
        logger.info('--- Transportador validado! Prosseguindo para validação da placa')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.press('enter')

    # Verifica se o campo da placa ficou preenchido
    time.sleep(0.2)
    if procura_imagem('imagens/img_topcon/campo_placa.png', confianca=0.74, 
                     continuar_exec=True, limite_tentativa=4) is not False:
        logger.info('--- Encontrou o campo vazio, inserindo XXX0000')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.click(procura_imagem('imagens/img_topcon/campo_placa.png', continuar_exec=True))
        bot.write('XXX0000')
        bot.press('ENTER')

        # Volta a tela para a posição correta
        bot.click(1907, 78, 2)
    else:
        logger.info('--- Não achou o campo ou já está preenchido')
    
    return True
