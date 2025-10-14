import time
import pytesseract
import pyautogui as bot
from utils.funcoes import ahk as ahk
from utils.configura_logger import get_logger
from utils.funcoes import marca_lancado, procura_imagem, extrai_txt_img

logger = get_logger("automacao") # Obter logger configurado
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

def extrai_qtd(valor_escala = 0):
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 10)

    #* Realiza a extração da quantidade de toneladas
    for i in range(0, 20):
        try:
            qtd_ton = extrai_txt_img(imagem='img_toneladas.png', area_tela=(892, 577, 70, 20), porce_escala= valor_escala).strip()
            qtd_ton = qtd_ton.replace(",", ".")
            qtd_ton = float(qtd_ton)
        except ValueError:
            valor_escala += 10
        else:
            logger.debug(F'--- Texto coletado da quantidade: {qtd_ton}, Valor escala: {valor_escala}')
            return qtd_ton
    else:
        logger.error('--- Não foi possivel extrair as toneladas da NFE!" ')
        raise Exception('--- Não foi possivel extrair as toneladas da NFE!')

def valida_preenchimento():
    logger.info('--- Verificando se apresentou erro no preenchimento do SILO ')

    for tentativa in range (0, 3):
        ahk.win_activate('TopCompras', title_match_mode=2)
        logger.info('--- Aguardando fechamento da tela: Itens da nota fiscal de compra ')
        time.sleep(0.2)

        if procura_imagem(imagem='imagens/img_topcon/local_estoque_obrigatorio.png', continuar_exec=True, limite_tentativa= 4):
            marca_lancado("erro_local_estoque")
            ahk.win_activate('TopCompras', title_match_mode=2)
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec=True))
            time.sleep(0.4)
            bot.click(procura_imagem(imagem='imagens/img_topcon/bt_cancela.png', continuar_exec=True))
            raise Exception('--- Falhou no preenchimento do SILO para essa nota fiscal')
            #return False

def preenche_local(silo1 = "", silo2 = ""):
    logger.info('--- Iniciando a função PREENCHE LOCAL ')
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 10)

    valor_escala = 200
    for i in range (0, 50):
        qtd_ton = extrai_qtd(valor_escala)

        logger.info('--- Abrindo a tela "Itens nota fiscal de compra" ')
        bot.click(procura_imagem(imagem='imagens/img_topcon/botao_alterar.png', area=(100, 839, 300, 400)))

        logger.info('--- Aguardando aparecer a tela "Itens nota fiscal de compra" ')
        for i in range (0, 100):
            time.sleep(0.4)
            if procura_imagem(imagem='imagens/img_topcon/valor_cofins.png', continuar_exec= True, limite_tentativa= 3, confianca= 0.73) is False:
                pass
            else:
                break
        else:
            raise Exception('Falhou ao encontrar a opção: VALOR COFINS')

        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        time.sleep(0.4)
        logger.info('--- Preenchendo SILO e quantidade')

        #* Verifica se a informação está duplicada nos dois campos
        if silo1 == silo2:
            logger.warning('--- A informação nos dois campos "SILO" estava igual, corrigindo apenas para um preenchimento')
            silo2 = ""

        #* Preenche os locais de estoque informados, e a quantidade.
        if ('SILO' in silo1) or ('SILO' in silo2):
            bot.click(851, 443)  # Clica na linha para informar o primeiro silo
            if ('SILO' in silo1) and ('SILO' in silo2):  
                qtd_ton = str((qtd_ton / 2)) # Realiza a divisão da quantidade de cimento, pois será distribuido em dois silos!
                qtd_ton = qtd_ton.replace(".", ",")
                logger.info(F'--- Foi informado dois silos, preenchendo... {silo1} e {silo2}, quantidade: {qtd_ton}')
                bot.write(silo1)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
                bot.write(silo2)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
            else:
                logger.info(F'--- Foi informado UM silo, preenchendo... {silo1}, quantidade: {qtd_ton}')
                qtd_ton = str(qtd_ton)
                qtd_ton = qtd_ton.replace(".", ",")
                bot.write(silo1)
                bot.write(silo2)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
        else: #* Caso não tenha coletado nenhum silo.            
            if procura_imagem(imagem='imagens/img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True):
                marca_lancado(texto_marcacao= 'Faltou_InfoSilo')
                logger.info('--- Não foi informado nenhum SILO, porém a nota é de cimento!' )
                ahk.win_activate('TopCompras', title_match_mode=2)
                bot.click(procura_imagem(imagem='imagens/img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True))
                bot.press('ESC')
                time.sleep(0.4)
                return True
            else: # Caso realmente seja de agregado.
                logger.info('--- Nota de agregado, continuando o processo!')
            
            bot.click(procura_imagem(imagem='imagens/img_topcon/confirma.png'))
            break

        #* Confirma ou cancela os processo executados na tela "itens nota fiscal de compras "
        bot.click(procura_imagem(imagem='imagens/img_topcon/confirma.png'))    
        time.sleep(1)   

        if procura_imagem(imagem='imagens/img_topcon/txt_ErroAtribuida.png', continuar_exec = True) is False:
            logger.info('--- Preenchimento completo, saindo do loop.' )
            break
        else:
            logger.warning(F'--- Falha, executando novamente a coleta das toneladas. Escala atual: {valor_escala}' )
            valor_escala += 10
            while procura_imagem(imagem='imagens/img_topcon/confirma.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
                bot.press('ENTER')
                bot.press('ESC')
                time.sleep(0.4)

        '''
        while procura_imagem(imagem='imagens/img_topcon/confirma.png', continuar_exec=True) is not False:
            ahk.win_activate('TopCompras', title_match_mode=2)
            logger.info('--- Aguardando fechamento da tela do botão "Alterar" ')
            time.sleep(0.4)

            if procura_imagem(imagem='imagens/img_topcon/local_estoque_obrigatorio.png', continuar_exec=True):
                marca_lancado("Erro local estoque")
                ahk.win_activate('TopCompras', title_match_mode=2)
                bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec=True))
                bot.click(procura_imagem(imagem='imagens/img_topcon/bt_cancela.png', continuar_exec=True))
                return False
                
            tentativa += 1
            if tentativa > 10: # Caso a tela não feche.
                raise Exception('Não foi possivel fechar a tela "itens nota fiscal de compra')
        else:
            logger.warning('--- Não fechou a tela "itens nota fiscal de compras" ')
        '''
    else:
        raise Exception('--- Falhou ao tentar preencher o local')


def main(silo1, silo2: str = ''):
    preenche_local(silo1, silo2)
    valida_preenchimento()


if __name__ == '__main__':
    main(silo1= 'SILO 8')
    #valida_preenchimento()