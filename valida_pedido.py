# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pytesseract
import pyautogui as bot
from utils.funcoes import ahk as ahk
from utils.configura_logger import get_logger
from utils.funcoes import marca_lancado, procura_imagem, extrai_txt_img, verifica_ped_vazio, corrige_nometela, print_erro

# --- Definição de parametros
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
numero_nf = "965999"
transportador = "111594"
logger = get_logger("automacao") # Obter logger configurado
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
#Nome dos itens que constarem no "Itens XML"
PEDRA_1 = ('BRITA]','BRITA1', 'PEDRA 01', 'PEDRA DI', 'BRITADA 01', 'PEDRA 1', 'PEDRA BRITADA 01', 'PEDRAT', 'PEDRA BRITADA 1', 'BRITADA 1', 
           'BRITA 01', 'BRITA 1', 'BRITA NR "01"', 'BRITA O01', 'PEDRA No 1', 'PEDRA1')
PO_PEDRA = ('PO DE PEDRA', 'AREA INDUSTRIAL', 'AREIA INDUSTRIAL')
BRITA_0 = ('BRITA 0', 'PEDRISCO LIMPO', 'BRITAD™', 'BRITAO', 'PEDRISCO LAVADDO', 'BRITA ZERO')
CIMENTO_CP3 = ('CP 111', 'teste')
CIMENTO_CP2 = ('CP II-E-40', '£-40', 'CIMENTO PORTLAND CP IIE-40 RS |', "CIMENTO PORTLAND CP I'E-40 RS.", "CIMENTO PORTLAND CP IE-40 RS")
AREIA_RIO = ('AREIA LAVADA MEDIA', 'ARE A LAVADA MEDIA', 'AREA LAVADA MEDIA', 'AREIA LAVADA')
CIMENTO_CP5 = ('CPV', 'V-ARI')
AREIA_QUARTZO = ('AREIA DE QUARTZO VERMELHA', 'AREA QUARTZD', 'AREIA DE QUARTZ0 VERMELHA', 'P2 AREIA', 'AREI|A DE QUARTZ0', 'AREIA VERMELHA',
                 'AREIA MEDIA UMIDA BRANCA', 'AREI|A MEDIA UMIDA BRANCA')
AREIA_PRIME = ('AREA PRIME', 'AREIA PRIME')
AREIA_BRITADA = ('AR EIA ARTIF ClaL', 'AR EIA AR TIFICIAL', 'AREIA ARTIFICIAL', 'AREIA INDUSTRIAL DE BRITA')
PEDRISCO_MISTO = ('PEDRA MISTO', 'PEDRISCO MISTO')
nome_pedido = [PEDRA_1, PO_PEDRA, BRITA_0, CIMENTO_CP2, CIMENTO_CP3, CIMENTO_CP5, AREIA_RIO, AREIA_QUARTZO, AREIA_PRIME, AREIA_BRITADA, PEDRISCO_MISTO]
# Mapeamento de nomes para imagens
mapeamento_imagens = {
    PEDRA_1: 'PED_BRITA1.jpg',
    PO_PEDRA: 'PED_POPEDRA.png',
    BRITA_0: 'PED_BRITA0.jpg',
    CIMENTO_CP2: 'PED_CPIIE40.png',
    AREIA_RIO: 'PED_AREIARIO.png',
    CIMENTO_CP5: 'PED_CIMENTOCPV.png',
    AREIA_QUARTZO: 'PED_AREIAFINA.png',
    AREIA_PRIME: 'PED_AREIAPRIME.png',
    AREIA_BRITADA: 'PED_AREIABRITA.png',
    CIMENTO_CP3: 'PED_CPIII40.png',
    PEDRISCO_MISTO: 'PED_PEDRISCOMISTO.png'
}


def valida_pedido():
    logger.info('--- Executando função: valida pedido' )
    bot.PAUSE = 0.2
    tentativa = 0
    img_pedido = 0
    item_pedido = ''
    validou_itensXml = False

    #Confirma a abertura da tela de vinculação do pedido
    ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
    time.sleep(0.5)

    #* Coleta o texto do campo "item XML", que é o item a constar na nota fiscal, e com base nisso, trata o dado
    logger.debug('--- Extraindo a imagem para descobrir qual item consta no campo "Itens XML" ')
    txt_itensXML = extrai_txt_img(imagem='item_nota.png',area_tela=(170, 407, 280, 20))
    logger.info(F'--- Texto extraido do "Itens XML": {txt_itensXML} ')

    #* Indentifica qual o item que consta na extração.
    for nome in nome_pedido: #* Para cada item na lista de pedidos
        for item_pedido in nome: #Para cada item dentro das linhas.          
            #Pesquisa se o txt_itensXML bate com o item da lista atual
            if item_pedido in txt_itensXML:
                #Verificação do nome no mapeamento
                if (nome in mapeamento_imagens) and (validou_itensXml is False):
                    img_pedido = mapeamento_imagens[nome]
                    logger.debug(F'--- O item {item_pedido} é igual ao item extraido da NFE: {txt_itensXML}, procurando a imagem: {img_pedido}')
                    ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
                    validou_itensXml = True
                    break
    
    #* Caso não tenha encontrado o texto em nenhuma lista. 
    if validou_itensXml is False:
        logger.error(F'--- Não foi possivel encontrar: "{txt_itensXML}" em nenhuma lista, marcando planilha com "padronizar item" ')
        marca_lancado(texto_marcacao='Padronizar_Item')
        for i in range(0, 30): # Aguarda o fechamento da tela "Vinculação Itens da Nota"
            time.sleep(0.4)
            if ahk.win_exists('Vinculação Itens da Nota', title_match_mode = 2):
                ahk.win_close('Vinculação Itens da Nota', title_match_mode = 2)
            else:
                logger.error('--- Finalizou a task VALIDA PEDIDO ( Fechou a tela "Vinculação Itens da Nota" )')
                break
        return False

#* --------------------------------- Pedidos Encontrados ---------------------------------
    if ahk.win_exists('Vinculação Itens da Nota') is False:
        if ahk.win_is_active(' (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 1):
            corrige_nometela("Vinculação Itens da Nota")

    vazio = False
    while (tentativa < 4) and (vazio is False):
        vazio = False 
        time.sleep(0.2)
        ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
        ahk.win_wait_active('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
        
        #* Validação para saber se encontrou em algum local, caso não encontre, exibe um erro.
        #* Tenta encontrar a imagem do pedido e salva as posições onde encontrar
        
        #* Caso seja a segunda tentativa, já baixa a lista de pedidos.
        if tentativa >= 1 and tentativa <= 4: 
            logger.info(F'--- Está na tentativa: {tentativa} de vincular pedido, baixando a lista dos pedidos')
            time.sleep(0.2)
            ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
            ahk.win_wait_active('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
            bot.click(748, 310) # Clica para descer o menu e exibir o resto das opções
            time.sleep(0.3)

        #* Procura o pedido na tela "6201 - Vinculação Itens da Nota" e conta quantos itens encontrou
        #posicoes = bot.locateAllOnScreen('imagens/img_pedidos/' + img_pedido, confidence= 0.92, grayscale=True, region=(0, 0, 850, 400))
        posicoes = bot.locateAllOnScreen('imagens/img_pedidos/' + img_pedido, confidence= 0.92, grayscale=True, region=(0, 0, 650, 185))
        contagem = 0
        for pos in posicoes:
            contagem += 1

        #* Caso não encontre a imagem em na tela de vinculação de pedido ( significa que falta pedido. )
        if contagem == 0 and tentativa > 3 : 
            logger.warning(F'--- Não encontrou: {img_pedido}, NÃO EXISTE PEDIDO! Saindo do processo.')
            print_erro()
            marca_lancado(texto_marcacao= 'Pedido_Inexistente')
            ahk.win_close('Vinculação Itens da Nota', title_match_mode = 2)
            ahk.win_wait_close('Vinculação Itens da Nota', title_match_mode = 2, timeout= 30)
            time.sleep(0.5)
            bot.press('F2')
            vazio = False
            tentativa = 5
            logger.debug("--- Saindo da função VALIDA PEDIDO, acabou pedido = False")
            return False # Retorna false pois não concluiu o processo.
        else:
            logger.info(F"Existem: {contagem} pedidos para o item {txt_itensXML}")
            ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
            time.sleep(0.2)
            posicoes = bot.locateAllOnScreen('imagens/img_pedidos/' + img_pedido, confidence= 0.92, grayscale=True, region=(0, 0, 850, 400))

        #Verifica nas posições que encontrou
        for pos in posicoes:  # Tenta em todos pedidos encontrados
            logger.debug(F'--- Tentativa: {tentativa}, achou o {txt_itensXML} na posição {pos}')
            ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
            
            bot.doubleClick(pos)  # Marca o pedido encontrado
            bot.click(procura_imagem(imagem='imagens/img_topcon/localizar.png'))
            vazio = verifica_ped_vazio(texto=txt_itensXML, pos=pos)
            logger.debug(F'--- Valor campo "item XML": {vazio}')

            if vazio is not True:
                bot.click(procura_imagem('imagens/img_topcon/vinc_xml_pedido.png',continuar_exec=True, limite_tentativa= 4, confianca= 0.75))
                vazio = verifica_ped_vazio(texto=txt_itensXML, pos=pos)
                logger.info(F'--- Valor campo "vazio": {vazio}')
                if vazio is True:
                    tentativa = 5
                    break
                if procura_imagem('imagens/img_topcon/dife_valor.png', continuar_exec=True, limite_tentativa= 4):
                    bot.press('ENTER')
                if procura_imagem('imagens/img_topcon/operacao_fiscal_configurada.png', continuar_exec=True, limite_tentativa= 4):
                    bot.press('ENTER')
                
                #Confere se após clicar nos botões, ainda assim o campo ficou vazio.
                if verifica_ped_vazio(texto=txt_itensXML, pos=pos) is not True:
                    logger.debug(F'--- Não ficou vazio, desmarcando pedido, indo para proxima tentativa {tentativa}')
                    bot.doubleClick(pos) # Clica novamente no mesmo pedido, para desmarcar
            else:
                logger.success(F'--- Pedido validado! VALIDA PEDIDO concluida.  ( valor do campo: {vazio} )')
                return True
            
        # 1. Verificar se encontrou o campo cinza abaixo dos pedidos
        # 2. Caso encontrou, não precisa de outra tentativa, já testou todos os pedidos
        # Se achar, significa que não precisa realizar uma nova tentativa, pois já achou o "final" dos pedidos
        ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
        time.sleep(0.2)
        if procura_imagem('imagens/img_topcon/bt_final_pedido.png', continuar_exec=True, limite_tentativa= 4, area= (726, 285, 35, 60)):
            bot.click(procura_imagem('imagens/img_topcon/bt_final_pedido.png', continuar_exec=True, limite_tentativa= 4, area= (726, 285, 35, 60)))
            tentativa = 5

        tentativa += 1
        
    else:
        if vazio is False:
            logger.warning('--- Acabou o pedido, fechando "vincula pedido" e marcando informação na planilha' )
            marca_lancado('Erro_Pedido_' + img_pedido)
            while ahk.win_exists('Vinculação Itens da Nota', title_match_mode = 2):
                ahk.win_close('Vinculação Itens da Nota', title_match_mode = 2)
                ahk.win_wait_close('Vinculação Itens da Nota', title_match_mode = 2)
                
            time.sleep(0.2)
            bot.press('F2')
            return False


def verifica_tela_vinculacao():
    logger.info('--- Executando a função VERIFICA TELA VINCULAÇÃO --- ')
    #* Aguarda a abertura da tela "Vinculação Itens da Nota"
    for i in range(0, 20):
        ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
        time.sleep(0.2)
        if ahk.win_is_active('Vinculação Itens da Nota', title_match_mode = 2):
            logger.info('Tela "Vinculação Itens da Nota" aberta!')
            break
        if i == 10:
            logger.error('Não encontrou a tela "Vinculação Itens da Nota" ')
            ahk.win_wait_active('Vinculação Itens da Nota', title_match_mode= 2, timeout= 1)


#* Verifica se a tela "Vinculação itens da Nota" carregou e está exibindo o botão "localizar"
def valida_bt_localizar():
    logger.info('--- Executando a função VALIDA BT LOCALIZAR --- ')
    for i in range (0, 20):
        if procura_imagem(imagem='imagens/img_topcon/localizar.png', limite_tentativa= 5, continuar_exec= True):
            logger.info('--- Tela "Vinculação itens da NOTA" carregou e encontrou o botão "LOCALIZAR" ')
            break
        if i >= 9:
            logger.error('--- Tela "Vinculação itens da NOTA" não carregou corretamente')
            raise Exception('Tela "Vinculação itens da NOTA" não carregou corretamente')


def main():
    logger.info('--- Executando o arquivo VALIDA PEDIDO --- ')
    verifica_tela_vinculacao()
    valida_bt_localizar()
    return valida_pedido()
    

if __name__ == '__main__':
    tempo_inicial = time.time()
    main()

    # Linha específica onde você quer medir o tempo
    end_time = time.time()
    elapsed_time = end_time - tempo_inicial
    medicao_minutos = elapsed_time / 60
    print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")
    bot.alert("acabou")