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
bot.FAILSAFE = False
numero_nf = "965999"
transportador = "111594"
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''

def valida_pedido(acabou_pedido=False):
    bot.PAUSE = 0.8
    tentativa = 0
    item_pedido = ''
        
    #Nome dos itens que constarem no "Itens XML"
    PEDRA_1 = ['PEDRA 01', 'PEDRA DI', 'BRITADA 01', 'PEDRA 1', 'PEDRA BRITADA 01', 'PEDRAT', 'PEDRA BRITADA 1', 'BRITADA 1', 'BRITA 01', 'BRITA 1', 'BRITA NR "01"']
    PO_PEDRA = ['PO DE PEDRA', 'AREA INDUSTRIAL', 'INDUSTRIAL']
    BRITA_0 = ['BRITA 0', 'PEDRISCO LIMPO', 'LAVAD', 'BRITAD™']
    CIMENTO_CP2 = ['-40', 'E-40', '£-40', 'II-E-40', 'CIMENTO PORTLAND CP II-E-40 RS |', 'CIMENTO PORTLAND CP IIE-40 RS', 'CIMENTO PORTLAND CP IIE-40 RS |',
                    "CIMENTO PORTLAND CP I'E-40 RS.", "CIMENTO PORTLAND CP IE-40 RS"]
    AREIA_RIO = ['AREIA LAVADA MEDIA', 'AREIA MEDIA']
    CIMENTO_CP5 = ['CPV']
    AREIA_QUARTZO = ['AREIA DE QUARTZO VERMELHA', 'AREA QUARTZD']
    AREIA_PRIME = ['AREA PRIME']

    nome_pedido = [PEDRA_1, PO_PEDRA, BRITA_0, CIMENTO_CP2, AREIA_RIO, AREIA_QUARTZO, AREIA_PRIME]

    #Força a abertura da tela de vinculação de item versus nota
    ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
    ahk.win_wait_active('Vinculação Itens da Nota', title_match_mode = 2, timeout= 20)
    time.sleep(0.5)

    #Coleta o texto do campo "item XML", que é o item a constar na nota fiscal, e com base nisso, trata o dado
    txt_itensXML = extrai_txt_img(imagem='item_nota.png',area_tela=(170, 400, 280, 30))
    print(F'Texto extraido do campo Itens XML: {txt_itensXML}') 

    #Indentifica qual o item que consta na extração.
    definiu_pedido = False
    img_pedido = 0
    for nome in nome_pedido: #Para cada item na lista.
        if definiu_pedido is True:
            break
        print(F'\nVerificando lista: {nome}')
        for item_pedido in nome: #Para cada item dentro das linhas.
            #print(f'--- Comparando o texto extraido: {txt_itensXML }, com o texto: {item_pedido}')            
            #Pesquisa se o txt_itensXML bate com o item da lista atual
            if item_pedido in txt_itensXML:
                print(F'\n--- O item da lista: {item_pedido} é igual ao item extraido: {txt_itensXML}')
                print(F'Lista onde encontrou: {nome}\n')
                
                if nome == PEDRA_1:
                    img_pedido = 'PED_BRITA1.jpg'
                elif nome == PO_PEDRA:
                    img_pedido = 'PED_POPEDRA.png'
                elif nome == BRITA_0:
                    img_pedido = 'PED_BRITA0.jpg'
                elif nome == CIMENTO_CP2:
                    img_pedido = 'PED_CPIIE40.png'
                elif nome == AREIA_RIO:
                    img_pedido = 'PED_AREIARIO.png'
                elif nome == CIMENTO_CP5:
                    img_pedido = 'PED_CIMENTOCPV.png'
                elif nome == AREIA_QUARTZO:
                    img_pedido = 'PED_AREIAFINA.png'
                elif nome == AREIA_PRIME:
                    img_pedido = 'PED_AREIAPRIME.png'
                else:
                    exit(bot.alert(F'Texto não padronizado, verificar script, texto: {txt_itensXML.strip()}'))
                definiu_pedido = True
                break
            else:
                pass
                #print(F'--- O item da lista: {item_pedido} NÃO é igual ao item extraido: {txt_itensXML}')
            
    print(F'--- Valor da variavel img_pedido: {img_pedido}')
    
#* --------------------------------- Pedidos Encontrados ---------------------------------
    while tentativa <= 2:
        vazio = ''
        print(F'--- Tentativa: {tentativa}')
        if tentativa > 0:
            print('--- Baixando a lista dos pedidos')
            bot.click(744, 230) #Clica para descer o menu e exibir o resto das opções 
        #Tenta encontrar a imagem do pedido e salva as posições onde encontrar
        time.sleep(0.5)
        posicoes = bot.locateAllOnScreen('img_pedidos/' + img_pedido, confidence=0.9, grayscale=True, region=(0, 0, 850, 400))
        print('--- Procurando nas posições os pedidos')
        for pos in posicoes:  # Tenta em todos pedidos encontrados
            print(F'Achou o {txt_itensXML} na posição {pos}')
            bot.doubleClick(pos)  # Marca o pedido encontrado
            bot.click(procura_imagem(imagem='img_topcon/localizar.png'))

            vazio = verifica_ped_vazio(texto=txt_itensXML, pos=pos)
            print(F'--- Valor campo "vazio": {vazio}')
            if vazio is not True:
                bot.click(procura_imagem('img_topcon/vinc_xml_pedido.png',continuar_exec=True, limite_tentativa=2))
                vazio = verifica_ped_vazio(texto=txt_itensXML, pos=pos)
                print(F'--- Valor campo "vazio": {vazio}')
                if procura_imagem('img_topcon/dife_valor.png', continuar_exec=True, limite_tentativa=2):
                    bot.press('ENTER')
                if procura_imagem('img_topcon/operacao_fiscal_configurada.png', continuar_exec=True, limite_tentativa=2):
                    bot.press('ENTER')
                if verifica_ped_vazio(texto=txt_itensXML, pos=pos) is not True:
                    print('--- Não ficou vazio, desmarcando pedido')
                    bot.doubleClick(pos) # Clica novamente no mesmo pedido, para desmarcar
            else:
                print(F'--- Pedido validado, saindo do loop dos pedidos encontrados, valor do campo: {vazio}')
                break
            tentativa += 1
        
        #Marcou o pedido, saindo dos loop
        if vazio is True: 
            break
    else:
        if vazio is False:
            bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
            marca_lancado('Erro_Pedido')
            acabou_pedido = True
            return acabou_pedido