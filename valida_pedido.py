# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import cv2
import pytesseract
from ahk import AHK
from funcoes import marca_lancado, procura_imagem, extrai_txt_img
import pyautogui as bot

# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1.2  # Pausa padrão do bot
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"
transportador = "111594"
#tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''

def verifica_ped_vazio(texto, tentativa, pos):
    texto_xml = extrai_txt_img('valida_itensxml.png', area_tela=(168, 400, 250, 30)).strip()
    print(F'Item da nota: {texto}, texto que ainda ficou: {texto_xml}')
    if len(texto_xml) > 4:  # Verifica se ainda ficou item no "Itens pedido"
        print('Itens XML ainda tem informação! desmarcando pedido marcado')
        bot.doubleClick(pos)
        acabou_pedido = False
    else: #Caso fique vazio
        print('Itens XML ficou vazio! prosseguindo')
        bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
        bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg'))
        acabou_pedido = False
        tentativa = 5
    if acabou_pedido is True:  # Caso durante a verificação tenha ocorrido algum erro
        bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
        marca_lancado(texto_marcacao='Erro_Pedido')
    time.sleep(1)
    return tentativa

def valida_pedido(acabou_pedido = False):
    tentativa = 0
    # * -------------------------------------- VALIDAÇÃO DO PEDIDO --------------------------------------
    while tentativa <= 2:
        time.sleep(2)
        item_pedido = []
        
        # 1: Topo direto imagem, #2 inferior lado esquerdo
        texto = extrai_txt_img(imagem = 'item_nota.png', area_tela=(170, 400, 280, 30))      
        PEDRA_1 = ['PEDRA 01', 'PEDRA DI', 'BRITADA 01', 'PEDRA 1', 'PEDRA BRITADA 01']
        if texto in PEDRA_1:
            print('Contém PEDRA 1')
            item_pedido.append('PED_BRITA1.jpg')
        elif ('CPV' in texto):
            print('Contém CIMENTO CP V/5')
            item_pedido.append('PED_CIMENTOCPV.png')
        elif ('PEDRISCO LIMPO' in texto) or ('LAVAD' in texto):
            print('Contém PEDRISCO LIMPO')
            item_pedido.append('PED_BRITA0.jpg')
        elif ('AREIA PRIME' in texto) or ('AREA PRIME' in texto):
            print('Contém AREIA PRIME')
            item_pedido.append('PED_AREIAPRIME.png')
        elif ('E-40' in texto):
            print(F'Contém Cimento CP II E 40, texto coletado: {texto}')
            item_pedido.append('PED_CPIIE40.png')
        elif ('CP 1ll' in texto) or ('CP lll' in texto) or ('CP 111' in texto):
            print(F'Contém Cimento CP III, texto coletado: {texto}')
            item_pedido.append('PED_CPIII40.png')
        elif 'PO DE PEDRA' in texto:
            print('Contém PO DE PEDRA')
            item_pedido.append('PED_POPEDRA.png')
        elif 'QUARTZ' in texto:
            print('Contém AREIA QUARTZO')
            item_pedido.append('PED_AREIAFINA.png')
        else:
            exit(print(F'Texto não padronizado, verificar script, texto: {texto.strip()}'))

        posicoes = bot.locateAllOnScreen('img_pedidos/' + item_pedido[0], confidence=0.9, grayscale=True, region=(0, 0, 850, 321))
        for pos in posicoes:  # Tenta em todos pedidos encontrados
            print(F'Achou o {texto} na posição {pos}')
            bot.doubleClick(pos)
            bot.click(procura_imagem(imagem='img_topcon/localizar.png'))
            tentativa = verifica_ped_vazio(texto = texto, tentativa = tentativa, pos = pos)
            exit(print(tentativa))
            bot.click(procura_imagem('img_topcon/vinc_xml_pedido.png', continuar_exec=True, limite_tentativa= 2))
            if procura_imagem('img_topcon/dife_valor.png', continuar_exec=True, limite_tentativa= 2):
                bot.press('ENTER')
            if procura_imagem('img_topcon/operacao_fiscal_configurada.png', continuar_exec=True, limite_tentativa= 2):
                bot.press('ENTER')
            tentativa = verifica_ped_vazio(texto = texto, tentativa = tentativa, pos = pos)
            #TODO -
            #TODO -
            #TODO -
            #TODO -
            time.sleep(2)
        else:
            bot.click(744, 227, 1)
            if tentativa >= 2: #Caso já tenha realiza duas execuções
                bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
                marca_lancado('Erro_Pedido')
                acabou_pedido = True
                return acabou_pedido
            tentativa += 1
                 
valida_pedido()