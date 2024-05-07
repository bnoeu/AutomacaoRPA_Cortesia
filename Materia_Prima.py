# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EU6ahKCIVdxFjiB_rViPfN0Bo9SGYGReQ7VTqbKDjMXyLQ?e=QrTGT0

import os
import time
#import sqlite3
#import subprocess
import pytesseract
from ahk import AHK
import pyautogui as bot
from datetime import date
from colorama import Back, Style
from valida_pedido import valida_pedido
from acoes_planilha import valida_lancamento
from funcoes import marca_lancado, procura_imagem, extrai_txt_img

# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1
posicao_img = 0
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

# * ---------------------------------------------------------------------------------------------------
# *                                        Inicio do Programa
# * ---------------------------------------------------------------------------------------------------
def abre_topcon():
    #Primeiro força o fechamento das telas, para evitar erros de validações.
    ahk.win_kill('Segurança do Windows', title_match_mode= 2)
    ahk.win_kill('RemoteApp', title_match_mode= 2)
    ahk.win_kill('RemoteApp', title_match_mode= 2)
    ahk.win_kill('TopCon', title_match_mode= 2)
    
    os.startfile('RemoteApp-Cortesia.rdp')
    
    #Realiza o login no RDP, que deve utilizar as informações de login do usuario "CORTESIA\BARBARA.K"
    ahk.win_activate('Segurança do Windows')
    ahk.win_wait_active('Segurança do Windows')
    bot.click(procura_imagem(imagem='img_windows/txt_seguranca.png'))
    bot.write('C0rtesi@01') #Senha BARBARA.K
    bot.press('TAB', presses= 3, interval= 0.02)
    bot.press('ENTER')
    print('--- Login realizado no RemoteApp-Cortesia.rdp')
    
    #Realiza login no TopCon
    while procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png') is False: #Aguarda até aparecer o campo do servidor preenchido
        time.sleep(0.2)
    else:
        bot.click(procura_imagem(imagem='img_topcon/txt_ServidorAplicacao.png'))
        print('--- Tela de login do topcon aberta')
        bot.press('tab', presses= 2, interval= 0.005)
        bot.press('backspace')
        
        #Insere os dados de login do usuario BRUNO.S
        bot.write('BRUNO.S')
        bot.press('tab')
        bot.write('rockie')
        bot.press('tab')
        bot.press('enter')

    #Abre o modulo de compras e navega até a tela de lançamento
    bot.click(procura_imagem(imagem='img_topcon/icone_compras.png'))
    ahk.win_activate('TopCompras - Versão', title_match_mode= 2)
    bot.press('ENTER')
    exit()
    while procura_imagem(imagem='img_topcon/txt_interveniente.png', continuar_exec= True) is False:
        print('--- Aguardando modulo de compras abrir.')
    else:
        ahk.win_activate('TopCompras - Versão', title_match_mode= 2)
        #bot.click(procura_imagem(imagem='img_topcon/txt_interveniente.png'))
 
 
def programa_principal():
    while True:  # ! Programa principal.
        acabou_pedido = True
        tentativa = 0
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
            print(Back.GREEN + F'Crachá: {cracha_mot} Silo1: {silo1} Silo2: {silo2}, {filial_estoq}, {chave_xml}' + Style.RESET_ALL)
            acabou_pedido = valida_pedido(acabou_pedido=False)
            print(F'Chegou até aqui acabou pedido = {acabou_pedido}')
            
#* -------------------------- PROSSEGUINDO COM O LANÇAMENTO DA NFE -------------------------- 
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 5)
        time.sleep(0.4) 
        bot.click(900, 201)  # Clica no campo filial de estoque
        bot.write(filial_estoq)
        time.sleep(0.2)

        # Confirma a informação da nova filial de estoque
        bot.press('ENTER', presses=1)
        time.sleep(1.2)

        bot.click(1006, 345)  # Campo data da operação
        hoje = date.today()
        hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
        bot.write(hoje)
        bot.press('enter')       
        time.sleep(1)

        # Altera o campo centro de custo, para o dado coletado
        print(F'--- Trocando o centro de custo para {centro_custo}')
        bot.write(centro_custo)

        # Aguarda aparecer o campo "cod_desc"
        print('--- Aguarda aparecer o campo cod_desc')
        while procura_imagem(imagem='img_topcon/cod_desc.png', limite_tentativa=1, continuar_exec=True) is False:
            time.sleep(0.2)
        bot.press('ENTER')

        # Aguarda até SUMIR o campo "cod_desc"
        print('--- Aguarda até SUMIR o campo "cod_desc"')
        while procura_imagem(imagem='img_topcon/cod_desc.png', limite_tentativa=1, continuar_exec=True) is not False:
            time.sleep(0.2)
            ahk.win_wait_active('TopCompras', title_match_mode= 2)
            ahk.win_activate('TopCompras', title_match_mode= 2)

        # Clica no campo "Valores Totais"
        bot.doubleClick(105, 515, interval= 0.2)

        # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
        print(F'--- PREENCHENDO TRANSPORTADOR: {cracha_mot}')
        bot.click(317, 897)  # Campo transportador
        while procura_imagem(imagem='img_topcon/campo_re_0.png', continuar_exec= True) is False:
            time.sleep(0.2)
        else:
            print('--- Campo RE habilitado, preenchendo.')
            
        #Preenche o campo do transportador e verifica se aconteceu algum erro.
        bot.write(cracha_mot)  # ID transportador
        bot.press('enter', interval= 0.5)
        
        print('--- Aguardando validar o campo do transportador')
        if procura_imagem(imagem='img_topcon/transportador_incorreto.png', continuar_exec= True) is not False:
            print('--- Transportador incorreto!')
            bot.press('ENTER')
            marca_lancado(texto_marcacao='RE_incorreto')
            programa_principal()
        else:
            print('--- Transportador validado! Prosseguindo para validação da placa')
        bot.press('enter')

        # Verifica se o campo da placa ficou preenchido
        if procura_imagem('img_topcon/campo_placa.png', continuar_exec=True) is not False:
            bot.click(procura_imagem('img_topcon/campo_placa.png', continuar_exec=True))
            bot.write('XXX0000')
            bot.press('ENTER')
        else:
            print('--- Não achou o campo ou já está preenchido')


        # * -------------------------------------- Aba Pedido --------------------------------------
        bot.doubleClick(procura_imagem(imagem='img_topcon/produtos_servicos.png'))
        time.sleep(0.8)

        if '38953477000164' not in chave_xml: #Caso não tenha o CNPJ da Consmar
            # Realiza a extração da quantidade de toneladas
            qtd_ton = extrai_txt_img(imagem='img_toneladas.png', area_tela=(892, 577, 70, 20)).strip()
            qtd_ton = qtd_ton.replace(",", ".")
            qtd_ton = float(qtd_ton)
            print(F'--- Texto coletado da quantidade: {qtd_ton}')
            
            #Clica no alterar para exibir a tela 
            bot.click(procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400)))

            #Verifica se abriu a tela com os detalhes do item que consta na NFE (Tela botão alterar)
            while procura_imagem(imagem='img_topcon/valor_cofins.png', limite_tentativa= 12) is False:
                print('--- Aguardando aparecer a tela "Itens nota fiscal de compra" ')
                time.sleep(0.3)
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
                tentativa += 1
                print('--- Aguardando fechamento da tela do botão "Alterar" ')
                time.sleep(0.6)
                if tentativa > 10: #Executa o loop 10 vezes até dar erro.
                    exit(bot.alert('Apresentou algum erro.'))
            # TODO --- CASO O REMOTE APP DESCONECTE, RODAR O ABRE TOPCON
            
        # Conclui o lançamento
        bot.press('pagedown')  # Conclui o lançamento
        print('--- Aguardando TopCompras Retornar')
        while ahk.win_exists('Não está respondendo'):
            time.sleep(0.4)

        # Verifica se a tela "Deseja processar" apareceu, caso sim, procede para emissão da NFE.
        ahk.win_wait_active('TopCom', timeout=10, title_match_mode=2)
        ahk.win_activate('TopCom', title_match_mode=2)
        
        # Espera até aparecer a tela de operação realizada, e quando ela aparecer, clica no botão OK
        while procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec=True) is False:
            if procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa= 1, continuar_exec=True) is not False:
                print('--- Nota já lançada, marcando planilha!')
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Lancado_Manual')
                programa_principal()

        bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec=True))

        #Verifica se apareceu a tela de transferencia 
        if procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True, limite_tentativa= 10) is not False:
            bot.click(procura_imagem('img_topcon/bt_sim.png', continuar_exec=True))
            while True:  # Aguardar o .PDF
                try:
                    ahk.win_wait('.pdf', title_match_mode=2, timeout=2)
                    time.sleep(0.6)
                except TimeoutError:
                    print('--- Aguardando .PDF da transferencia')
                else:
                    ahk.win_activate('.pdf', title_match_mode=2)
                    ahk.win_close('pdf - Google Chrome', title_match_mode=2)
                    print('--- Fechou o PDF da transferencia')
                    break
            time.sleep(0.8)
            ahk.win_activate('Transmissão', title_match_mode=2)
            bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))

        # * -------------------------------------- Marca planilha --------------------------------------
        marca_lancado(texto_marcacao='Lancado_RPA')
#abre_topcon()
programa_principal()

# TODO --- Caso o pedido acabe, avisar ao Mateus
# TODO --- Caso NFE Faturada no final do mes, lançar com qual data? 
# TODO --- Chave_XML não está baixando via gerenciador, agora é tudo pelo OBJ