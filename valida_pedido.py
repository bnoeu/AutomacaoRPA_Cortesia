# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import cv2
#import pytesseract
from ahk import AHK
from funcoes import marca_lancado, procura_imagem, extrai_txt_img, verifica_ped_vazio
import pyautogui as bot

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"
transportador = "111594"
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''

def valida_pedido(acabou_pedido=False):
    bot.PAUSE = 0.5
    tentativa = 0
    item_pedido = []
    
    #Nome dos itens que constarem no "Itens XML"
    PEDRA_1 = ['PEDRA 01', 'PEDRA DI', 'BRITADA 01', 'PEDRA 1', 'PEDRA BRITADA 01', 'PEDRAT', 'PEDRA BRITADA 1', 'BRITADA 1', 'BRITA 01', 'BRITA 1', 'BRITA NR "01"']
    PO_PEDRA = ['PO DE PEDRA', 'AREA INDUSTRIAL', 'INDUSTRIAL']
    BRITA_0 = ['BRITA 0', 'PEDRISCO LIMPO', 'LAVAD']
    CIMENTO_CP2 = ['-40', 'E-40', '£-40', 'II-E-40', 'CIMENTO PORTLAND CP II-E-40 RS |', 'CIMENTO PORTLAND CP IIE-40 RS', 'CIMENTO PORTLAND CP IIE-40 RS |',
                   "CIMENTO PORTLAND CP I'E-40 RS."]
    AREIA_RIO = ['AREIA LAVADA MEDIA']

    
    #Força a abertura da tela de vinculação de item versus nota
    time.sleep(1)
    ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
    ahk.win_wait_active('Vinculação Itens da Nota', title_match_mode = 2, timeout= 20)
    time.sleep(2)
    
    #Coleta o texto do campo "item XML", que é o item a constar na nota fiscal, e com base nisso, trata o dado
    texto = extrai_txt_img(imagem='item_nota.png',area_tela=(170, 400, 280, 30))
    print(F'Texto extraido do campo Itens XML: {texto}')   
    
    #Com base no texto extraido, identifica qual o pedido
    if texto in PEDRA_1:
        print('Contém PEDRA 1')
        item_pedido.append('PED_BRITA1.jpg')
    elif ('CPV' in texto):
        print('Contém CIMENTO CP V/5')
        item_pedido.append('PED_CIMENTOCPV.png')
    elif ('CP 1ll' in texto) or ('CP lll' in texto) or ('CP 111' in texto) or ('1-40' in texto):
        print(F'Contém Cimento CP III, texto coletado: {texto}')
    elif texto in CIMENTO_CP2:
        print(F'Contém Cimento CP II E 40, texto coletado: {texto}')
        item_pedido.append('PED_CPIIE40.png')
    elif texto in BRITA_0:
        print('Contém PEDRISCO LIMPO')
        item_pedido.append('PED_BRITA0.jpg')
    elif ('AREIA PRIME' in texto) or ('AREA PRIME' in texto):
        print('Contém AREIA PRIME')
        item_pedido.append('PED_AREIAPRIME.png')
    elif texto in PO_PEDRA:
        print('Contém PO DE PEDRA')
        item_pedido.append('PED_POPEDRA.png')
    elif 'QUARTZ' in texto:
        print('Contém AREIA QUARTZO')
        item_pedido.append('PED_AREIAFINA.png')
    elif 'AREIA ARTIFICIAL' in texto:
        print('Contém AREIA ARTIFICIAL')
        item_pedido.append('PED_AREIABRITA.png')
    elif texto in AREIA_RIO:
        print('Contém AREIA DE RIO')
        item_pedido.append('PED_AREIARIO.png')
    else:
        exit(bot.alert(F'Texto não padronizado, verificar script, texto: {texto.strip()}'))
        
#* --------------------------------- Pedidos Encontrados 
    while tentativa <= 2:
        vazio = ''
        print(F'--- Tentativa: {tentativa}')
        if tentativa > 1:
            bot.click(744, 230) #Clica para descer o menu e exibir o resto das opções 
            exit(F'Tentativa: {tentativa}')
        posicoes = bot.locateAllOnScreen('img_pedidos/' + item_pedido[0], confidence=0.9, grayscale=True, region=(0, 0, 850, 400))
        time.sleep(1)
        for pos in posicoes:  # Tenta em todos pedidos encontrados
            print(F'Achou o {texto} na posição {pos}')
            bot.doubleClick(pos)  # Marca o pedido encontrado
            bot.click(procura_imagem(imagem='img_topcon/localizar.png'))
            time.sleep(1)

            vazio = verifica_ped_vazio(texto=texto, pos=pos)
            print(F'--- Valor campo "vazio": {vazio}')
            if vazio is not True:
                bot.click(procura_imagem('img_topcon/vinc_xml_pedido.png',continuar_exec=True, limite_tentativa=2))
                time.sleep(1)
                vazio = verifica_ped_vazio(texto=texto, pos=pos)
                print(F'--- Valor campo "vazio": {vazio}')
                if procura_imagem('img_topcon/dife_valor.png', continuar_exec=True, limite_tentativa=2):
                    bot.press('ENTER')
                if procura_imagem('img_topcon/operacao_fiscal_configurada.png', continuar_exec=True, limite_tentativa=2):
                    bot.press('ENTER')
                if verifica_ped_vazio(texto=texto, pos=pos) is not True:
                    print('--- Não ficou vazio, desmarcando pedido')
                    bot.doubleClick(pos) # Clica novamente no mesmo pedido, para desmarcar
            else:
                print(F'--- Pedido validado, saindo do loop dos pedidos encontrados, valor do campo: {vazio}')
                break
            tentativa += 1
        else:
            bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png', confianca= 0.6))
            marca_lancado('Erro_Pedido')
            acabou_pedido = True
            return acabou_pedido
        '''
        if vazio is False:  # Caso já tenha realiza duas execuções
            bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png', confianca= 0.6))
            marca_lancado('Erro_Pedido')
            acabou_pedido = True
            return acabou_pedido
        else:
            break
        '''