import time
import pytesseract
import pyautogui as bot
from utils.funcoes import procura_imagem  
from abre_topcon import main as abre_topcon
from utils.configura_logger import get_logger

# --- Definição de parametros
from utils.funcoes import ahk as ahk
bot.LOG_SCREENSHOTS = True
bot.LOG_SCREENSHOTS_LIMIT = 5
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
logger = get_logger("automacao") # Obter logger configurado

def processo_transferencia():
    ahk.win_activate('TopCompras', title_match_mode=2)
    if procura_imagem(imagem='imagens/img_topcon/txt_transferencia.png', continuar_exec= True, limite_tentativa= 3, confianca= 0.74):
        logger.info('--- Iniciando a função: processo transferencia ---' )
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        
        if procura_imagem('imagens/img_topcon/deseja_processar.png', continuar_exec=True, confianca= 0.75):
            logger.warning('--- Encontrou a tela "Deseja processar NFE" ')
            while procura_imagem('imagens/img_topcon/bt_sim.png', continuar_exec=True, limite_tentativa= 3, confianca= 0.74):
                bot.click(procura_imagem('imagens/img_topcon/bt_sim.png', continuar_exec=True))
                
            contador_pdf = 0
            while True:  # Aguardar o .PDF
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
                    logger.info('--- Aguardando .PDF da transferencia')
                    
                else:
                    ahk.win_activate('.pdf', title_match_mode=2)
                    ahk.win_close('pdf - Google Chrome', title_match_mode=2)
                    logger.info('--- Fechou o PDF da transferencia')
                    break
            
            # Fechando a tela de transmissão
            while ahk.win_exists('Transmissão', title_match_mode= 2):
                logger.debug('--- Fechando a tela de Transmissão da NFE')
                ahk.win_activate('Transmissão', title_match_mode=2)
                bot.click(procura_imagem(imagem='imagens/img_topcon/sair_tela.png'))
                
                ahk.win_wait_active('TopCompras', timeout=10, title_match_mode=2)
                ahk.win_activate('TopCompras', title_match_mode=2)
                logger.warning('--- Forçando reabertura do mercantil, pois foi executada uma transferencia.')
                abre_topcon() #! Sempre que é executado o processo de transferencia, é necessario reabrir o TopCompras, pois a tela trava.
                return True
        elif procura_imagem("imagens/img_topcon/txt_numero_nf_alterado.png", continuar_exec= True):
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2)
            bot.press('enter')
        else:
            logger.info('--- Não encontrou a tela "Deseja processar NFE ainda" ')

if __name__ == '__main__':
    processo_transferencia()