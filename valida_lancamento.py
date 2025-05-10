# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pyautogui as bot
from abre_topcon import abre_mercantil, main as fechar_tela_nota_compra
from automacao.conferencia_xml import conferencia_xml
from coleta_planilha import main as coleta_planilha
from utils.funcoes import ativar_janela, procura_imagem, corrige_nometela
from utils.configura_logger import get_logger


# --- Definição de parametros
from utils.funcoes import ahk as ahk
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
logger = get_logger("script1")

def altera_topcon_incluir():
    logger.info('--- Alterando o TopCompras para o modo incluir')
    
    for i in range(0, 6):
        logger.info('--- Verificando se está no modo Localizar.')
        ativar_janela('TopCompras')
        time.sleep(1)
        
        if procura_imagem(imagem='imagens/img_topcon/txt_inclui.png', limite_tentativa= 5, continuar_exec= True, area= (852, 956, 1368, 1045)):
            logger.info('--- Está no modo "incluir", enviando comando F2 para entrar no modo "Localizar"')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('F2', presses= 2)
            time.sleep(0.5)

        if procura_imagem(imagem='imagens/img_topcon/txt_localizar.png', limite_tentativa= 5, continuar_exec= True, area= (852, 956, 1368, 1045)):
            logger.info('--- Está no modo "Localizar" Alterando para "Incluir"')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            time.sleep(0.2)
            bot.press('F3', presses= 2)
            time.sleep(1)

        if procura_imagem(imagem='imagens/img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
            logger.success('--- Está no modo "Incluir", lançamento pode continuar!')
            break

        if procura_imagem(imagem='imagens/img_topcon/txt_existe_nota_transferencia.png', continuar_exec= True):
            logger.warning('--- Encontrou a tela "existe nota fiscal de transferencia" ')
            corrige_nometela("TopCompras (VM-CortesiaApli.CORTESIA.com)")

            ahk.win_activate("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode= 2)
            bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg'))
            logger.info('--- Clicou para fechar a tela "existe nota fiscal de transferencia" ')

            for tentativa in range (0, 4):
                if ahk.win_exists("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode= 2):
                    ahk.win_close("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode= 2)
                    if tentativa >= 3:
                        raise Exception("Não foi possivel fechar a tela de transferencia de NFE") 
                else:
                    break

        if i == 3:
            fechar_tela_nota_compra()
            abre_mercantil()

        if i >= 5:
            logger.error('--- Atingiu o maximo de tentativas de alterar os botões ---')
            #bot.alert("Limite de tentativas de alterar o topcon para o modo incluir")        
            raise Exception("Atingiu o maximo de tentativas de alterar os botões")
    

# Realiza o processo de validação do lançamento.
def valida_lancamento():
    validou_xml = False
    bot.PAUSE = 0.4
    
    logger.info('--- Iniciando função VALIDA LANÇAMENTO')
    while validou_xml is False:        
        dados_planilha = False
        while dados_planilha is False:
            time.sleep(0.2)
            dados_planilha = coleta_planilha() # Recebe os dados coletados da planilha, já validados e formatados.
            if dados_planilha == False:
                raise Exception("Copiou dados novos! Necessario reiniciar o processo!")
        #* Trata a chave XML, removendo os espaços caso exista.
        chave_xml = dados_planilha[4].strip()

        #* Enquanto a tela não for alterada para o modo incluir
        altera_topcon_incluir()
        
        # Inicia inserção da chave XML
        bot.press('TAB', presses= 2, interval = 1)
        bot.write(chave_xml)
        bot.press('TAB')
        
        validou_xml = conferencia_xml() # Confere qual tela será apresentada.

        if validou_xml is not False:
            logger.success(F'--- Validou o XML! Prosseguindo para a seleção do pedido: {validou_xml}')
            return dados_planilha # Após todas as validações, retorna os dados para a execução principal

if __name__ == '__main__':
    ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
    tempo_inicial = time.time()
    
    valida_lancamento()
    
    # Linha específica onde você quer medir o tempo
    end_time = time.time()
    elapsed_time = end_time - tempo_inicial
    medicao_minutos = elapsed_time / 60
    print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")
    bot.alert("acabou")

