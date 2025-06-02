import time
import pytesseract
import pyautogui as bot
from utils.funcoes import ahk as ahk
from utils.funcoes import procura_imagem  
from abre_topcon import fechar_topcompras, abre_mercantil
from utils.configura_logger import get_logger
#from utils.configura_logger import get_logger

'''
if __name__ == '__main__':
    from configura_logger import get_logger
else:
    from .configura_logger import get_logger
'''


# --- Definição de parametros
bot.LOG_SCREENSHOTS = True
bot.LOG_SCREENSHOTS_LIMIT = 5
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
logger = get_logger("automacao") # Obter logger configurado

def fecha_espelho_nota():
    logger.info('--- Fechando a tela espelho da nota fiscal ')
    ahk.win_wait('Espelho de Nota Fiscal', title_match_mode= 2, timeout= 30)
    ahk.win_activate('Espelho de Nota Fiscal', title_match_mode= 2)
    time.sleep(0.8)

    for i in range (0, 15):
        if ahk.win_exists('Espelho de Nota Fiscal', title_match_mode= 2):
            ahk.win_close('Espelho de Nota Fiscal', title_match_mode= 2)
        else:
            return True

def tela_deseja_imprimir_espelho_nota():
    """_summary_ Verifica se apareceu a tela "Deseja imprimir o Espelho nota" e fecha ela.
    """
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait("TopCompras", title_match_mode= 2)
    time.sleep(1)

    if procura_imagem(imagem='imagens/img_topcon/txt_espelhonota.png', limite_tentativa= 4, continuar_exec=True, confianca= 0.95):
        logger.info('--- Encontrou a tela "Deseja imprimir um espelho da nota fiscal" ')
        bot.click(procura_imagem('imagens/img_topcon/bt_sim.png'))
        time.sleep(1)
        ahk.win_wait('Espelho de Nota Fiscal', title_match_mode= 2, timeout= 30)
        ahk.win_activate('Espelho de Nota Fiscal', title_match_mode= 2)

        return True
    else:
        return False

def fecha_operacao_realizada():
    """_summary_ Fecha a tela de "Operação realizada com sucesso", caso ela apareça
    """
    logger.info('--- Executando a função "Fecha operação realizada"')
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait("TopCompras", title_match_mode= 2)

    if procura_imagem(imagem='imagens/img_topcon/operacao_realizada.png', continuar_exec= True, confianca= 0.95, limite_tentativa= 4):
        logger.info('--- Encontrou a tela "Operação Realizada", realizando o fechamento.')
        ahk.win_activate('TopCompras', title_match_mode= 2)
        time.sleep(0.4)
        bot.click(procura_imagem(imagem='imagens/img_topcon/operacao_realizada.png'))
        time.sleep(0.4)
        bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg'))
        time.sleep(0.4)
        logger.info('--- Fechou a tela "Operação Realizada"')
        return True

def fecha_pdf_transferencia():
    """_summary_ Aguardar 10s para verificar se apareceu o PDF, se aparecer realiza o fechamento

    Returns:
        Boleano: True caso tenha encontrado e fechado. False caso não encontre
    """
    tempo_pausa = 0.5
    logger.info('--- Executando a função "FECHA PDF TRANSFERENCIA" ')

    for i in range (0, 50):
        time.sleep(tempo_pausa)

        if ahk.win_exists('.pdf', title_match_mode= 2):
            ahk.win_close('pdf - Google Chrome', title_match_mode=2)
            logger.info('--- PDF da transferencia fechado! ')
            return True

        if (i == 10) or (i == 20) or (i == 30):
            tempo_pausa += 1.5
            logger.info(f'--- Aguardando .PDF da transferencia, tentativa: {i}, aumentando pausa para: {tempo_pausa} segundos')
    else:
        logger.info(f'--- Não encontrou nenhum PDF para fechar, tentativa: {i}, tempo pausa: {tempo_pausa}')
        return False

    '''
    ahk.win_wait_active('.pdf', title_match_mode= 2, timeout= 8)

    for i in range (0, 15):  # Aguardar o .PDF
        time.sleep(0.4)
        try:
            ahk.win_wait('.pdf', title_match_mode=2, timeout= 8)
        except TimeoutError:
            if contador_pdf >= 15:
                # Fechando a tela de transmissão
                while ahk.win_exists('Transmissão', title_match_mode= 2):
                    ahk.win_activate('Transmissão', title_match_mode=2)
                    bot.click(procura_imagem(imagem='imagens/img_topcon/sair_tela.png'))
                    ahk.win_wait_close('Transmissão', title_match_mode=2, timeout= 15)
                    
                    ahk.win_wait_active('TopCompras', timeout=10, title_match_mode=2)
                    ahk.win_activate('TopCompras', title_match_mode=2)
                    return True

            contador_pdf += 1
            logger.info(f'--- Aguardando .PDF da transferencia, tentativa: {contador_pdf}')
            
        else:
            ahk.win_activate('.pdf', title_match_mode=2)
            ahk.win_close('pdf - Google Chrome', title_match_mode=2)
            logger.info('--- Fechou o PDF da transferencia')
            break
    else:
        raise Exception("Falhou ao executar o processo de transferencia")
    '''

def fecha_tela_transmissao():

    # Fechando a tela de transmissão
    for i in range (0, 15):
        if ahk.win_exists('Transmissão', title_match_mode= 2):
            logger.debug('--- Fechando a tela de Transmissão da NFE')
            ahk.win_activate('Transmissão', title_match_mode=2)
            bot.click(procura_imagem(imagem='imagens/img_topcon/sair_tela.png'))
            
            ahk.win_wait_active('TopCompras', timeout=10, title_match_mode=2)
            ahk.win_activate('TopCompras', title_match_mode=2)
            return True

def tela_deseja_processar_nota():
    """_summary_ Aguarda até aparecer a tela "Deseja processar a nota fiscal eletronica"

    Raises:
        Exception: _description_

    Returns:
        _type_: _description_
    """
    contador_pdf = 0

    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait("TopCompras", title_match_mode= 2)
    time.sleep(1)

    if procura_imagem('imagens/img_topcon/deseja_processar.png', continuar_exec=True, confianca= 0.95):
        logger.warning('--- Encontrou a tela "Deseja processar NFE"')

        while procura_imagem('imagens/img_topcon/bt_sim.png', continuar_exec=True, limite_tentativa= 3, confianca= 0.74):
            bot.click(procura_imagem('imagens/img_topcon/bt_sim.png', continuar_exec=True))
    
            if procura_imagem("imagens/img_topcon/txt_numero_nf_alterado.png", continuar_exec= True, limite_tentativa= 4):
                ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2)
                bot.press('enter')
                return True
            else:
                return True
    else:
        logger.info('--- Não encontrou a tela "Deseja processar NFE" ')


def processo_transferencia():
    ahk.win_activate('TopCompras', title_match_mode=2)

    # Procura o texto "TRANSFERENCIA" que aparece na tela 6203 - Transferencia, no campo "Operação Fiscal"
    logger.info('--- Procura o texto "TRANSFERENCIA" que aparece na tela 6203 - Transferencia, no campo "Operação Fiscal" ')
    for i in range (0, 5):
        if procura_imagem(imagem='imagens/img_topcon/txt_transferencia.png', continuar_exec= True, confianca= 0.74):
            logger.info('--- Iniciando a função: processo transferencia ---' )
            ahk.win_activate('TopCompras', title_match_mode=2)
            ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
            time.sleep(0.5)

            if tela_deseja_imprimir_espelho_nota():
                fecha_espelho_nota()
                fecha_operacao_realizada()

            if fecha_operacao_realizada():
                time.sleep(1)

            if tela_deseja_processar_nota():
                fecha_pdf_transferencia()
                fecha_tela_transmissao()

            # Caso já tenha fechado a tela de transferencia.
            if procura_imagem(imagem='imagens/img_topcon/txt_transferencia.png', continuar_exec= True, confianca= 0.74) is False:
                #! Sempre que é executado o processo de transferencia, é necessario reabrir o TopCompras, pois a tela trava.
                logger.warning('--- Forçando reabertura do mercantil, pois foi executada uma transferencia.')
                fechar_topcompras() 
                abre_mercantil()

    else:
        logger.debug("Não houve transferencia para essa nota fiscal, prosseguindo!")


if __name__ == '__main__':
    bot.PAUSE = 0.6
    #tela_deseja_processar_nota()
    fecha_pdf_transferencia()
    fecha_tela_transmissao()
    #tela_deseja_processar_nota()
    #processo_transferencia()
    #tela_deseja_imprimir_espelho_nota()
    #fecha_operacao_realizada()
    #deleta_espelho_nota()