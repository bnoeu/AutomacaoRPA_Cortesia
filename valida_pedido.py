# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import cv2
import pytesseract
from ahk import AHK
from funcoes import marca_lancado, procura_imagem
import pyautogui as bot

# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1.2  # Pausa padrão do bot
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"
transportador = "111594"
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''


def valida_pedido():
    tentativa = 0
    # * -------------------------------------- VALIDAÇÃO DO PEDIDO --------------------------------------
    while tentativa <= 2:
        time.sleep(1)
        item_pedido = []
        # 1: Topo direto imagem, #2 inferior lado esquerdo
        bot.screenshot('img_topcon/item_nota.png', region=(170, 400, 280, 30))
        print('--- Tirou print ----')
        img = cv2.imread('img_topcon/item_nota.png')
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        scale_percent = 180  # percent of original size
        width = int(img.shape[1] * scale_percent / 100)
        height = int(img.shape[0] * scale_percent / 100)
        dim = (width, height)
        img = cv2.resize(img, dim, interpolation=cv2.INTER_AREA)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        # cv2.imshow('T', thresh)
        # cv2.waitKey()
        texto = pytesseract.image_to_string(thresh, lang='eng', config='--psm 6').strip()

        if ('PEDRA 01' in texto) or ('PEDRA DI' in texto) or ('BRITADA 01' in texto):
            print('Contém PEDRA 1')
            item_pedido.append('PED_BRITA1.jpg')
        elif ('PEDRISCO LIMPO' in texto) or ('LAVAD' in texto):
            print('Contém PEDRISCO LIMPO')
            item_pedido.append('PED_BRITA0.jpg')
        elif 'AREIA PRIME' in texto:
            print('Contém AREIA PRIME')
            item_pedido.append('PED_AREIAPRIME.png')
        elif 'E-40' in texto:
            print('Contém Cimento CP II E 40')
            item_pedido.append('PED_CPIIE40.png')
        elif 'PO DE PEDRA' in texto:
            print('Contém PO DE PEDRA')
            item_pedido.append('PED_POPEDRA.png')
        elif 'QUARTZ' in texto:
            print('Contém AREIA QUARTZO')
            item_pedido.append('PED_AREIAFINA.png')
        else:
            exit(
                print(F'Texto não padronizado, verificar script, texto: {texto.strip()}'))

        posicoes = bot.locateAllOnScreen('img_pedidos/' + item_pedido[0], confidence=0.9, grayscale=True)
        for pos in posicoes:  # Tenta em todos pedidos encontrados
            print(F'Achou o {texto} na posição {pos}')
            bot.doubleClick(pos)
            bot.click(procura_imagem(imagem='img_topcon/localizar.png'))
            # 1: Topo direto imagem, #Erro_Pedi inferior lado esquerdo
            bot.screenshot('img_topcon/valida_itensxml.png', region=(168, 400, 250, 30))
            img = cv2.imread('img_topcon/valida_itensxml.png')
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
            texto_xml = pytesseract.image_to_string(thresh, lang='por', config='--psm 6').strip()
            print(F'Item da nota: {texto}, texto que ainda ficou: {texto_xml}')
            if len(texto_xml) > 2:  # Verifica se ainda ficou item no "Itens pedido"
                print('"Itens XML" ainda tem informação! desmarcando pedido marcado')
                bot.doubleClick(pos)
                acabou_pedido = False
            else:
                print('Itens XML ficou vazio! prosseguindo')
                bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
                bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg'))
                acabou_pedido = False
                tentativa = 5
                break
            if acabou_pedido is True:  # Caso durante a verificação tenha ocorrido algum erro
                bot.click(procura_imagem(imagem='img_topcon/bt_cancela.png'))
                marca_lancado(texto_marcacao='Erro_Pedido')
                # programa_principal()
            time.sleep(1)
        else:
            bot.click(744, 227, 1)
            if tentativa >= 2:
                exit('Não encontrou nenhum item na tela 6201 no campo "Pedidos", aumentar a sensibilidade')
            tentativa += 1