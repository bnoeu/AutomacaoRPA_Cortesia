# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import cv2
import pygetwindow as gw
import pytesseract
from ahk import AHK
from datetime import date
from funcoes import marca_lancado, procura_imagem, verifica_tela
from valida_pedido import valida_pedido
from coleta_planilha import coleta_planilha
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1.5  # Pausa padrão do bot
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
numero_nf = "965999"  # Valor para teste
transportador = "111594"  # Valor para teste
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"

#! Funções
def acoes_planilha():
    time.sleep(1)
    validou_xml = False
    while validou_xml is False:
        # * Trata os dados coletados em "dados_planilha"
        dados_planilha = coleta_planilha()
        chave_xml = dados_planilha[4].strip()
        # * -------------------------------------- Lançamento Topcon --------------------------------------
        bot.PAUSE = 1  # Pausa padrão do bot
        print('--- Abrindo TopCompras')
        ahk.win_activate('TopCompras')
        if ahk.win_is_active('TopCompras'):
            print('Tela compras está maximizada! Iniciando o programa')
        else:
            exit(bot.alert('Tela de Compras não abriu.'))
        # Processo de lançamento
        bot.press('F2')
        bot.press('F3', presses= 2, interval= 0.3)
        bot.click(558, 235)  # Clica dentro do campo para inserir a chave XML
        bot.write(chave_xml)
        bot.press('ENTER')
        ahk.win_wait_active('TopCompras')
        while procura_imagem(imagem='img_topcon/naorespondendo.png', limite_tentativa=3, continuar_exec=True) is not False:
            time.sleep(1)
            print('Aguardando topvoltar')
        tentativa = 0
        while tentativa < 10:
            # print(F'Tentativa: {tentativa}')
            if procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa=3, continuar_exec=True) is not False:
                bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa=2, continuar_exec=True))
                validou_xml is True
                return dados_planilha
            elif procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa=3, continuar_exec=True) is not False:
                print('--- Nota já lançada, marcando planilha!')
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Lancado_Manual')
                break
            elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', limite_tentativa=3, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Aguardando_SEFAZ')
                programa_principal()
            elif procura_imagem(imagem='img_topcon/chave_44digitos.png', limite_tentativa=3, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Chave_invalida')
                programa_principal()
            elif procura_imagem(imagem='img_topcon/transportador_incorreto.png', limite_tentativa=3, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Transportador_incorreto')
                programa_principal()
            tentativa += 1
        if tentativa >= 15:
            exit('Rodou 10 verificações e não achou nenhuma tela, aumentar o tempo')


# * ---------------------------------------------------------------------------------------------------
# *                                        Inicio do Programa
# * ---------------------------------------------------------------------------------------------------
def programa_principal():
    while True:  # ! Programa principal
        acabou_pedido = True
        while acabou_pedido is True:
            dados_planilha = acoes_planilha()
            cracha_mot = dados_planilha[0]
            silo1 = dados_planilha[1]
            silo2 = dados_planilha[2]
            filial_estoq = dados_planilha[3].split('-')
            filial_estoq = filial_estoq[0]
            if filial_estoq == '1001':
                centro_custo = 'VILA'
            elif filial_estoq == '1002':
                centro_custo = 'CACAPAVA'
            elif filial_estoq == '1003':
                centro_custo = 'BARUERI'
            elif filial_estoq == '1004':
                centro_custo = 'JAGUARE'
            elif filial_estoq == '1006':
                centro_custo = 'ATIBAIA'
            elif filial_estoq == '1008':
                centro_custo = 'MOGI'
            elif filial_estoq == '1005':
                centro_custo = 'SANTOS'
            else:
                exit(F'Filial de estoque não padronizada {filial_estoq}')
            chave_xml = dados_planilha[4]
            print(F'Crachá: {cracha_mot} Silo1: {silo1} Silo2: {silo2}, {filial_estoq}, {chave_xml}')
            time.sleep(1)
            acabou_pedido = valida_pedido(acabou_pedido=False)
            print(F'Chegou até aqui acabou pedido = {acabou_pedido}')
        # * -------------------------------------- PREENCHE DATA --------------------------------------
        bot.click(900, 201)  # Clica no campo filial de estoque
        bot.write(filial_estoq)
        # Confirma a informação da nova filial de estoque
        bot.press('ENTER', presses=2)
        bot.click(procura_imagem(imagem='img_topcon/data_operacao.jpg'))
        # bot.click(1006, 345)  # Campo data da operação
        hoje = date.today()
        hoje = hoje.strftime("%d%m%Y")  # dd/mm/YY
        bot.write(hoje, interval=0.10)
        bot.press('enter')
        time.sleep(10)
        ahk.win_wait_active('TopCompras')
        bot.write(centro_custo) # Altera o campo centro de custo, para o dado coletado
        bot.press('ENTER')
        time.sleep(2)
        while ahk.win_exists('Não está respondendo'):
            time.sleep(2)
        # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
        ahk.win_wait_active('TopCompras')
        bot.click(105, 515)  # Clica no campo "Valores Totais"
        time.sleep(1)
        bot.click(317, 897)  # Campo transportador
        time.sleep(2)
        procura_imagem(imagem='img_topcon/campo_re_0.png', limite_tentativa=20)
        print(F'--- PREENCHENDO TRANSPORTADOR: {cracha_mot}')
        time.sleep(1)
        bot.write(cracha_mot, interval=0.10)  # ID transportador
        bot.press('enter')
        #TODO --- Inserir a pesquisa do transportador_incorreto
        time.sleep(1)
        bot.press('enter')
        if bot.click(procura_imagem('img_topcon/campo_placa.png', continuar_exec=True)) is not False:
            bot.write('XXX0000')
            bot.press('ENTER')
        else:
            print('Achou o campo ou já está preenchido')
        # * -------------------------------------- Aba Pedido --------------------------------------
        bot.click(procura_imagem(
            imagem='img_topcon/produtos_servicos.png', limite_tentativa=8))
        bot.click(procura_imagem(imagem='img_topcon/botao_alterar.png',
                  limite_tentativa=8, area=(100, 839, 300, 400)))
        time.sleep(2.5)
        # 1: Topo direto imagem, #2 inferior lado esquerdo
        bot.screenshot('img_geradas/toneladas.png', region=(198, 167, 75, 25))
        print('--- Tirou print ----')
        # Verificação do texto da imagem.
        img = cv2.imread('img_geradas/toneladas.png')
        scale_percent = 140  # percent of original size
        width = int(img.shape[1] * scale_percent / 100)
        height = int(img.shape[0] * scale_percent / 100)
        dim = (width, height)
        img = cv2.resize(img, dim, interpolation=cv2.INTER_AREA)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        qtd_ton = pytesseract.image_to_string(thresh, config='--psm 6').strip()
        qtd_ton = qtd_ton.replace(",", ".")
        print(F'--- Texto coletado da quantidade: {qtd_ton}')
        time.sleep(1)
        bot.click(851, 443)  # Clica na linha para informar o primeiro silo
        if silo2 != '':  # realiza a divisão da quantidade de cimento
            qtd_ton = str((qtd_ton / 2))
            qtd_ton = qtd_ton.replace(".", ",")
            print(F'--- Foi informado dois silos, preenchendo... {silo1} e {silo2}, quantidade: {qtd_ton}')
            bot.write(silo1)
            bot.press('ENTER')
            bot.write(str(qtd_ton))
            bot.press('ENTER')
            bot.write(silo2)
            bot.press('ENTER')
            bot.write(str(qtd_ton))
            bot.press('ENTER')
        elif silo1 != '':
            print(F'--- Foi informado UM silo, preenchendo... {silo1}, quantidade: {qtd_ton}')
            qtd_ton = str(qtd_ton)
            qtd_ton = qtd_ton.replace(".", ",")
            bot.write(silo1)
            bot.press('ENTER')
            bot.write(str(qtd_ton))
            bot.press('ENTER')
        # * -------------------------------------- Conclusão lançamento --------------------------------------
        bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
        bot.press('pagedown')  # Conclui o lançamento
        while ahk.win_exists('Não está respondendo'):
            time.sleep(2)
        while procura_imagem(imagem='img_topcon/operacao_realizada.png', limite_tentativa=1000) is False:
                ahk.win_activate('TopCompras')
                time.sleep(0.5)
        bot.click(procura_imagem(imagem='img_topcon/operacao_realizada.png', limite_tentativa=1000))
        time.sleep(1)
        bot.press('ENTER')
        time.sleep(2)
        if bot.click(procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True, limite_tentativa= 8)) is not None:
                time.sleep(1)
                bot.press('S')
        ahk.win_wait('.pdf', title_match_mode= 2)
        ahk.win_activate('.pdf', title_match_mode= 2)
        ahk.win_close('pdf - Google Chrome', title_match_mode= 2)
        ahk.win_activate('Transmissão', title_match_mode=2)
        bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))
        time.sleep(5)
        exit('--- PROGRAMAR ESSA PARTE DO PROGRAMA.')


        
        tempo_final = time.time()
        tempo_final_seg = (tempo_final - tempo_inicio) / 60
        print(F'\n--- Lançamento concluido, tempo: {tempo_final_seg}')
        # * -------------------------------------- Marca planilha --------------------------------------
        marca_lancado(texto_marcacao='Lancado_RPA')
programa_principal()


# TODO --- Caso o pedido acabe, avisar ao Mateus
