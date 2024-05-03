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
bot.PAUSE = 0.2
item_pedido = ''

#Nome dos itens que constarem no "Itens XML"
PEDRA_1 = ('BRITA]','BRITA1', 'PEDRA 01', 'PEDRA DI', 'BRITADA 01', 'PEDRA 1', 'PEDRA BRITADA 01', 'PEDRAT', 'PEDRA BRITADA 1', 'BRITADA 1', 'BRITA 01', 'BRITA 1', 'BRITA NR "01"', 'BRITA O01')
PO_PEDRA = ('PO DE PEDRA', 'AREA INDUSTRIAL', 'INDUSTRIAL')
BRITA_0 = ('BRITA 0', 'PEDRISCO LIMPO', 'BRITAD™')
CIMENTO_CP3 = ('CP 111', 'teste')
CIMENTO_CP2 = ('E-40', '£-40', 'II-E-40', 'CIMENTO PORTLAND CP II-E-40 RS |', 'CIMENTO PORTLAND CP IIE-40 RS', 'CIMENTO PORTLAND CP IIE-40 RS |', "CIMENTO PORTLAND CP I'E-40 RS.", "CIMENTO PORTLAND CP IE-40 RS")
AREIA_RIO = ('AREIA LAVADA MEDIA', 'AREIA MEDIA', 'ARE A LAVADA MEDIA')
CIMENTO_CP5 = ('CPV', 'TESTE')
AREIA_QUARTZO = ('AREIA DE QUARTZO VERMELHA', 'AREA QUARTZD', 'AREIA DE QUARTZ0 VERMELHA', 'P2 AREIA')
AREIA_PRIME = ('AREA PRIME', 'TESTE', 'AREIA PRIME')
AREIA_BRITADA = ('AR EIA ARTIF ClaL', 'AR EIA AR TIFICIAL', 'teste', 'AREIA ARTIFICIAL')
nome_pedido = [PEDRA_1, PO_PEDRA, BRITA_0, CIMENTO_CP2, CIMENTO_CP3, CIMENTO_CP5, AREIA_RIO, AREIA_QUARTZO, AREIA_PRIME, AREIA_BRITADA]
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
    CIMENTO_CP3: 'PED_CPIII40.png'
}




def valida_pedido(acabou_pedido=False):
    tentativa = 0
    #Força a abertura da tela de vinculação de item versus nota
    ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
    ahk.win_wait_active('Vinculação Itens da Nota', title_match_mode = 2, timeout= 20)
    time.sleep(0.6)

    #Coleta o texto do campo "item XML", que é o item a constar na nota fiscal, e com base nisso, trata o dado
    txt_itensXML = extrai_txt_img(imagem='item_nota.png',area_tela=(170, 400, 280, 30))
    print(F'Texto extraido do campo Itens XML: {txt_itensXML}') 

    #Indentifica qual o item que consta na extração.
    definiu_pedido = False
    img_pedido = 0
    for nome in nome_pedido: #Para cada item na lista.
        if definiu_pedido is True:
            break
        
        for item_pedido in nome: #Para cada item dentro das linhas.          
            #Pesquisa se o txt_itensXML bate com o item da lista atual
            if item_pedido in txt_itensXML:

                #Verificação do nome no mapeamento
                if nome in mapeamento_imagens:
                    img_pedido = mapeamento_imagens[nome]
                    print(F'\n--- O item: {item_pedido} é igual ao item extraido: {txt_itensXML}, imagem: {img_pedido}')
                    #print(F'--- Lista onde encontrou: {nome}\n')
                    #print(F'--- Procurando a imagem: {img_pedido}')
                    validou_itensXml = True
                else:
                    validou_itensXml = False
                    if img_pedido == 0:
                        exit(bot.alert(f'Texto não padronizado, verificar script, texto: {txt_itensXML.strip()}'))
                definiu_pedido = True
        
    #Caso não tenha encontrado o texto em nenhuma lista. 
    if validou_itensXml is False:
        exit(bot.alert(F'--- Não foi possivel encontrar: {txt_itensXML} em nenhuma lista.'))
        
#* --------------------------------- Pedidos Encontrados ---------------------------------
    while tentativa <= 2:
        ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
        time.sleep(0.4)
        vazio = '' 
        
        #* Validação para saber se encontrou em algum local, caso não encontre, exibe um erro.
        #*Tenta encontrar a imagem do pedido e salva as posições onde encontrar
        if tentativa >= 1:
            print('--- Baixando a lista dos pedidos')
            bot.click(744, 230) #Clica para descer o menu e exibir o resto das opções

        print(F'--- Tentando localizar {img_pedido}')
        posicoes = bot.locateAllOnScreen('img_pedidos/' + img_pedido, confidence= 0.92, grayscale=True, region=(0, 0, 850, 400))

        contagem = 0
        for pos in posicoes:
            print(pos)
            contagem += 1
        print(F'Encontrou em: {contagem} posições')
        
        if contagem == 0:
            bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
            marca_lancado(texto_marcacao= 'Erro_Pedido')
            vazio = False
            acabou_pedido = True
            return acabou_pedido

        #Verifica nas posições que encontrou
        posicoes = bot.locateAllOnScreen('img_pedidos/' + img_pedido, confidence= 0.92, grayscale=True, region=(0, 0, 850, 400))
        for pos in posicoes:  # Tenta em todos pedidos encontrados
            time.sleep(0.3)
            #Caso já esteja na segunda tentativa, passa a tela para o lado
            print(F'--- Tentativa: {tentativa}, achou o {txt_itensXML} na posição {pos}')
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
                
                #Confere se após clicar nos botões, ainda assim o campo ficou vazio.
                if verifica_ped_vazio(texto=txt_itensXML, pos=pos) is not True:
                    time.sleep(0.3)
                    print(F'--- Não ficou vazio, desmarcando pedido, tentativa {tentativa}')
                    bot.doubleClick(pos) # Clica novamente no mesmo pedido, para desmarcar
            else:
                print(F'--- Pedido validado, saindo do loop dos pedidos encontrados, valor do campo: {vazio}')
                break
        
        
        #Marcou o pedido, saindo dos loop
        if vazio is True: 
            break
        else:
            tentativa += 1
    else:
        if vazio is False:
            bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
            marca_lancado('Erro_Pedido')
            acabou_pedido = True
            return acabou_pedido
        
if __name__ == '__main__':
    valida_pedido()