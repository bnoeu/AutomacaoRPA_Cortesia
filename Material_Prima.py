# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
import pytesseract
from ahk import AHK
from datetime import date
from funcoes import marca_lancado, procura_imagem, extrai_txt_img
from acoes_planilha import valida_lancamento
from valida_pedido import valida_pedido
import pyautogui as bot
#import sqlite3

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
tempo_inicio = time.time()

chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\tesseract\tesseract.exe"
bot.PAUSE = 1.4

'''
#Cria a conexão com o banco de dados
con = sqlite3.connect("informacoes.db")

#Cursor para realizar comandos dentro do banco de dados
cur = con.cursor()

#Utilizando o cursor, executa a ação da criação da tabela informacoes, com as seguintes colunas: XML, CRACHA, TEMPO
#cur.execute("CREATE TABLE informacoes(xml, cracha, tempo)")
#! Continuar tutorial de banco de dados https://docs.python.org/3/library/sqlite3.html
exit()
'''

# * ---------------------------------------------------------------------------------------------------
# *                                        Inicio do Programa
# * ---------------------------------------------------------------------------------------------------
def programa_principal():
    while True:  # ! Programa principal.
        acabou_pedido = True
        while acabou_pedido is True: #Verifica se o pedido está valido.
            dados_planilha = valida_lancamento()
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
            elif filial_estoq == '1005':
                centro_custo = 'SANTOS'
            elif filial_estoq == '1032':
                centro_custo = 'TAMOIO'
            else:
                exit(F'Filial de estoque não padronizada {filial_estoq}')
            chave_xml = dados_planilha[4]
            print(F'Crachá: {cracha_mot} Silo1: {silo1} Silo2: {silo2}, {filial_estoq}, {chave_xml}')
            acabou_pedido = valida_pedido(acabou_pedido=False)
            print(F'Chegou até aqui acabou pedido = {acabou_pedido}')
            
#* -------------------------- PROSSEGUINDO COM O LANÇAMENTO DA NFE -------------------------- 
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2)
        bot.click(900, 201)  # Clica no campo filial de estoque
        bot.write(filial_estoq)

        # Confirma a informação da nova filial de estoque
        bot.press('ENTER', presses=1)
        time.sleep(1)

        bot.click(1006, 345)  # Campo data da operação
        hoje = date.today()
        hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
        bot.write(hoje)
        bot.press('enter')

        # Altera o campo centro de custo, para o dado coletado
        bot.write(centro_custo)

        # Aguarda aparecer o campo "cod_desc"
        print('--- Aguarda aparecer o campo cod_desc')
        while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True) is False:
            time.sleep(0.5)
            
        bot.press('ENTER')

        # Aguarda até SUMIR o campo "cod_desc"
        print('--- Aguarda até SUMIR o campo "cod_desc"')
        while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True) is not False:
            time.sleep(0.5)
            ahk.win_wait_active('TopCompras', title_match_mode= 2)
            ahk.win_activate('TopCompras', title_match_mode= 2)
        else: #Sumiu, continuando a execução
            time.sleep(0.5)

        # Clica no campo "Valores Totais"
        bot.doubleClick(105, 515, interval= 0.5)

        # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
        print(F'--- PREENCHENDO TRANSPORTADOR: {cracha_mot}')
        bot.click(317, 897)  # Campo transportador
        while procura_imagem(imagem='img_topcon/campo_re_0.png') is False:
            time.sleep(0.5)
        else:
            print('--- Campo RE habilitado, preenchendo.')
            
        #Preenche o campo do transportador e verifica se aconteceu algum erro.
        bot.write(cracha_mot)  # ID transportador
        bot.press('enter')
        time.sleep(1)

        #Verifica 15 vezes se apareceu a tela "transportador incorreto", caso apareça, tratar.
        if procura_imagem(imagem='img_topcon/transportador_incorreto.png', continuar_exec= True) is not False:
            print('--- Transportador incorreto!')
            bot.press('ENTER')
            marca_lancado(texto_marcacao='RE_incorreto')
            programa_principal()
        else:
            print('--- Transportador validado! Prosseguindo para validação da placa')

        # Verifica se o campo da placa ficou preenchido
        bot.press('enter')
        if procura_imagem('img_topcon/campo_placa.png', continuar_exec=True) is not False:
            bot.click(procura_imagem('img_topcon/campo_placa.png', continuar_exec=True))
            bot.write('XXX0000')
            bot.press('ENTER')
        else:
            print('--- Não achou o campo ou já está preenchido')


        # * -------------------------------------- Aba Pedido --------------------------------------
        bot.doubleClick(procura_imagem(imagem='img_topcon/produtos_servicos.png'))

        # Realiza a extração da quantidade de toneladas
        qtd_ton = extrai_txt_img(imagem='img_toneladas.png', area_tela=(895, 577, 70, 20)).strip()
        qtd_ton = qtd_ton.replace(",", ".")
        qtd_ton = float(qtd_ton)
        print(F'--- Texto coletado da quantidade: {qtd_ton}')
        
        #Clica no alterar para exibir a tela 
        bot.click(procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400)))

        #Verifica se abriu a tela com os detalhes do item que consta na NFE (Tela botão alterar)
        while procura_imagem(imagem='img_topcon/valor_cofins.png') is False:
            print('--- Aguardando aparecer a tela "Itens nota fiscal de compra" ')
            time.sleep(0.2)
        else:
            print('--- Apareceu a tela "Itens nota fiscal de compra" ')
            # Aguardando aparecer o botão de "confirma", para prosseguir com as ações.
            procura_imagem(imagem='img_topcon/confirma.png')
            print('--- Preenchendo SILO e quantidade')


        if (silo1 != '') or (silo2 != ''):
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
        else:
            print('--- Nenhum silo coletado, nota de agregado!')

        #Após preencher ou não os silos, clica para confirmar as informações. 
        bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
        while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True) is not False:
            print('--- Aguardando fechamento da tela do botão "Alterar" ')
            time.sleep(0.5)
        
        # Conclui o lançamento
        bot.press('pagedown')  # Conclui o lançamento

        # Espera até que o topcom volte a responder
        print('--- Aguardando TopCompras Retornar')
        while ahk.win_exists('Não está respondendo'):
            time.sleep(0.5)
        time.sleep(1)

        # Verifica se a tela "Deseja processar" apareceu, caso sim, procede para emissão da NFE.
        ahk.win_wait_active('TopCom', timeout=10, title_match_mode=2)
        ahk.win_activate('TopCom', title_match_mode=2)
        
        # Espera até aparecer a tela de operação realizada, e quando ela aparecer, clica no botão OK
        while procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec=True) is False:
            time.sleep(0.5)
        else:
            bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec=True))


        #Verifica se apareceu a tela de transferencia 
        if procura_imagem('img_topcon/transferencia.png', continuar_exec=True, limite_tentativa=4) is not False:
            if procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True, limite_tentativa=4) is not False:
                bot.click(procura_imagem('img_topcon/bt_sim.png',
                        continuar_exec=True, limite_tentativa=4))
                while True:  # Aguardar o .PDF
                    try:
                        ahk.win_wait('.pdf', title_match_mode=2, timeout=2)
                        time.sleep(0.5)
                    except TimeoutError:
                        print('Aguardando .PDF')
                    else:
                        ahk.win_activate('.pdf', title_match_mode=2)
                        ahk.win_close('pdf - Google Chrome', title_match_mode=2)
                        print('Fechou o PDF')
                        break
                time.sleep(1)
                ahk.win_activate('Transmissão', title_match_mode=2)
                bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))

        # * -------------------------------------- Marca planilha --------------------------------------
        marca_lancado(texto_marcacao='Lancado_RPA')
programa_principal()

# TODO --- Caso o pedido acabe, avisar ao Mateus
