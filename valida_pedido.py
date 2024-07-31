# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
#import cv2
#import pytesseract
from ahk import AHK
import pyautogui as bot
from colorama import Fore, Style
from funcoes import marca_lancado, procura_imagem, extrai_txt_img, verifica_ped_vazio

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
numero_nf = "965999"
transportador = "111594"
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
#Nome dos itens que constarem no "Itens XML"
PEDRA_1 = ('BRITA]','BRITA1', 'PEDRA 01', 'PEDRA DI', 'BRITADA 01', 'PEDRA 1', 'PEDRA BRITADA 01', 'PEDRAT', 'PEDRA BRITADA 1', 'BRITADA 1', 'BRITA 01', 'BRITA 1', 'BRITA NR "01"', 'BRITA O01')
PO_PEDRA = ('PO DE PEDRA', 'AREA INDUSTRIAL', 'INDUSTRIAL')
BRITA_0 = ('BRITA 0', 'PEDRISCO LIMPO', 'BRITAD™', 'BRITAO')
CIMENTO_CP3 = ('CP 111', 'teste')
CIMENTO_CP2 = ('CP II-E-40', '£-40', 'CIMENTO PORTLAND CP IIE-40 RS |', "CIMENTO PORTLAND CP I'E-40 RS.", "CIMENTO PORTLAND CP IE-40 RS")
AREIA_RIO = ('AREIA LAVADA MEDIA', 'AREIA MEDIA', 'ARE A LAVADA MEDIA', 'AREA LAVADA MEDIA', 'AREIA LAVADA')
CIMENTO_CP5 = ('CPV', 'V-ARI')
AREIA_QUARTZO = ('AREIA DE QUARTZO VERMELHA', 'AREA QUARTZD', 'AREIA DE QUARTZ0 VERMELHA', 'P2 AREIA', 'AREI|A DE QUARTZ0', 'AREIA VERMELHA')
AREIA_PRIME = ('AREA PRIME', 'AREIA PRIME')
AREIA_BRITADA = ('AR EIA ARTIF ClaL', 'AR EIA AR TIFICIAL', 'AREIA ARTIFICIAL')
PEDRISCO_MISTO = ('PEDRA MISTO', 'TESTE')
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


def valida_pedido(acabou_pedido=False):
    bot.PAUSE = 0.6
    tentativa = 0
    img_pedido = 0
    item_pedido = ''
    validou_itensXml = False

    #Aguarda a abertura da tela de vinculação de item versus nota
    ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
    
    #Aguarda aparecer o botão "confirma" para poder continuar o processo.
    while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec= True, limite_tentativa= 2, confianca= 0.74) is False:
        time.sleep(0.2)

    #Coleta o texto do campo "item XML", que é o item a constar na nota fiscal, e com base nisso, trata o dado
    txt_itensXML = extrai_txt_img(imagem='item_nota.png',area_tela=(170, 407, 280, 20))
    print(F'--- Texto extraido do campo Itens XML: {txt_itensXML}') 

    #Indentifica qual o item que consta na extração.
    for nome in nome_pedido: #Para cada item na lista.
        for item_pedido in nome: #Para cada item dentro das linhas.          
            #Pesquisa se o txt_itensXML bate com o item da lista atual
            if item_pedido in txt_itensXML:
                #Verificação do nome no mapeamento
                if (nome in mapeamento_imagens) and (validou_itensXml is False):
                    img_pedido = mapeamento_imagens[nome]
                    print(Fore.GREEN + F'--- O item: {item_pedido} é igual ao item extraido: {txt_itensXML}, procurando a imagem: {img_pedido}' + Style.RESET_ALL)
                    validou_itensXml = True
                    break

    #Caso não tenha encontrado o texto em nenhuma lista. 
    if validou_itensXml is False:
        print(F'--- Não foi possivel encontrar: {txt_itensXML} em nenhuma lista.')
        bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
        marca_lancado(texto_marcacao='Padronizar_Item')
        return acabou_pedido

#* --------------------------------- Pedidos Encontrados ---------------------------------
    vazio = False
    while (tentativa < 3) and (vazio is False):
        ahk.win_activate('Vinculação Itens da Nota', title_match_mode = 2)
        time.sleep(0.2)
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
            #print(pos)
            contagem += 1
        print(F'--- Encontrou em: {contagem} posições')

        if contagem == 0:
            print(F'--- Não encontrou {img_pedido}, saindo do processo.')
            bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
            marca_lancado(texto_marcacao= 'Pedido_Inexistente')
            vazio = False
            tentativa = 3
            acabou_pedido = True
            return acabou_pedido
        else:
            posicoes = bot.locateAllOnScreen('img_pedidos/' + img_pedido, confidence= 0.92, grayscale=True, region=(0, 0, 850, 400))

        #Verifica nas posições que encontrou
        for pos in posicoes:  # Tenta em todos pedidos encontrados
            print(Fore.CYAN + F'--- Tentativa: {tentativa}, achou o {txt_itensXML} na posição {pos}' + Style.RESET_ALL)
            bot.doubleClick(pos)  # Marca o pedido encontrado
            bot.click(procura_imagem(imagem='img_topcon/localizar.png'))
            vazio = verifica_ped_vazio(texto=txt_itensXML, pos=pos)
            print(F'--- Valor campo "item XML": {vazio}')
            if vazio is not True:
                bot.click(procura_imagem('img_topcon/vinc_xml_pedido.png',continuar_exec=True))
                vazio = verifica_ped_vazio(texto=txt_itensXML, pos=pos)
                print(F'--- Valor campo "vazio": {vazio}')
                if vazio is True:
                    tentativa = 3
                    break
                if procura_imagem('img_topcon/dife_valor.png', continuar_exec=True):
                    bot.press('ENTER')
                if procura_imagem('img_topcon/operacao_fiscal_configurada.png', continuar_exec=True):
                    bot.press('ENTER')
                
                #Confere se após clicar nos botões, ainda assim o campo ficou vazio.
                if verifica_ped_vazio(texto=txt_itensXML, pos=pos) is not True:
                    tentativa += 1
                    print(F'--- Não ficou vazio, desmarcando pedido, indo para proxima tentativa {tentativa}')
                    bot.doubleClick(pos) # Clica novamente no mesmo pedido, para desmarcar
            else:
                print(Fore.GREEN + F'--- Pedido validado, saindo do loop dos pedidos encontrados, valor do campo: {vazio}' + Style.RESET_ALL)
                return False

    else:
        if vazio is False:
            print(Fore.RED + '--- Acabou o pedido, marcando informação na planilha' + Style.RESET_ALL)
            bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
            marca_lancado('Erro_Pedido_' + img_pedido)
            acabou_pedido = True
            return acabou_pedido
        
if __name__ == '__main__':
    valida_pedido()