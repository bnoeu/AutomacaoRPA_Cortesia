# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EU6ahKCIVdxFjiB_rViPfN0Bo9SGYGReQ7VTqbKDjMXyLQ?e=QrTGT0
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EW_8FZwWFYVAol4MpV1GglkBJEaJaDx6cfuClnesIu60Ng?e=pveECF
# Debug db alltrips
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg

import os
import time
#import sqlite3
#import subprocess
import pytesseract
from ahk import AHK
import pyautogui as bot
from datetime import date
from colorama import Back, Style, Fore
from valida_pedido import valida_pedido
from acoes_planilha import valida_lancamento
from funcoes import marca_lancado, procura_imagem, extrai_txt_img

# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 1.5
posicao_img = 0
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

# * ---------------------------------------------------------------------------------------------------
# *                                        Inicio do Programa
# * ---------------------------------------------------------------------------------------------------

def programa_principal():
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
        elif filial_estoq == '1036':
            centro_custo = 'PERUS'
        else:
            exit(F'Filial de estoque não padronizada {filial_estoq}')
        chave_xml = dados_planilha[4]
        acabou_pedido = valida_pedido(acabou_pedido=False)

#* -------------------------- PROSSEGUINDO COM O LANÇAMENTO DA NFE -------------------------- 
    print('--- Preenchendo dados na tela principal do lançamento')
    
    time.sleep(1)
    while procura_imagem(imagem='img_topcon/produtos_servicos.png', continuar_exec= True) is False:
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
        try:
            ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 25)
        except TimeoutError:
            bot.alert(exit('Topcompras não encontrado'))
        time.sleep(0.2)

    bot.press('up')
    bot.write(filial_estoq)
    bot.press('TAB', presses= 2) # Confirma a informação da nova filial de estoque
    #* --- Alteração da data
    time.sleep(1)
    hoje = date.today()
    hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
    #dias_fatura = ['23', '29', '30', '31', '01']
    #data_NfeFaturada = extrai_txt_img(imagem='valida_itensxml.png', area_tela=(895, 299, 20, 20))
    #bot.write('22/06/2024')
    bot.press('ENTER')
    if procura_imagem(imagem='img_topcon/txt_NaoPermitidoData.png', continuar_exec=True, limite_tentativa= 12):
        print(Fore.RED + '--- Precisa mudar a data' + Style.RESET_ALL)
        bot.press('enter')
        bot.write(hoje)
        bot.press('enter')
        time.sleep(0.5)

    ''' #! Depreciado por: Não estava coletando a data de faturamento de forma correta, sendo assim o codigo acima substitui essa logica. 
    if data_NfeFaturada in dias_fatura:
        print(Back.RED + F'--- Data de faturamento menor que 20, data faturada: {data_NfeFaturada}' + Style.RESET_ALL)
        #bot.write(F'{data_NfeFaturada}' + '052024')
        #bot.write(hoje)
        bot.press('enter')
        time.sleep(0.5)
        if procura_imagem(imagem='img_topcon/txt_NaoPermitidoData.png', continuar_exec=True, limite_tentativa= 12):
            print(Fore.RED + '--- Precisa mudar a data' + Style.RESET_ALL)
            bot.press('enter')
            bot.write(hoje)
            bot.press('enter')
            time.sleep(0.5)
    else:
        print(F'--- Alterando a data para {hoje}, data NFE coletada {data_NfeFaturada}')
        bot.write(hoje)
        bot.press('enter')
        time.sleep(0.5)
    '''

    print(F'--- Trocando o centro de custo para {centro_custo}')
    bot.write(centro_custo)

    print('--- Aguarda aparecer o campo cod_desc')
    while procura_imagem(imagem='img_topcon/cod_desc.png', limite_tentativa= 50, continuar_exec=True) is False:
        time.sleep(0.2)
    else:
        print('--- Apareceu o campo COD_DESC')

    bot.press('ENTER') # Pressiona enter, e aguarda sumir o campo "cod_desc"
    print('--- Aguarda até SUMIR o campo "cod_desc"')
    while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True) is not False:
        time.sleep(0.1)

    bot.click(procura_imagem(imagem='img_topcon/txt_ValoresTotais.png', continuar_exec= True))
    # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
    print(F'--- Preenchendo transportador: {cracha_mot}')
    bot.click(procura_imagem(imagem='img_topcon/campo_000.png', continuar_exec= True))
    bot.press('tab')
    while procura_imagem(imagem='img_topcon/campo_re_0.png', continuar_exec= True) is False:
        time.sleep(0.1)
    else:
        print('--- Campo RE habilitado, preenchendo.')
        #Preenche o campo do transportador e verifica se aconteceu algum erro.
        bot.write(cracha_mot)  # ID transportador
        bot.press('enter')

    print('--- Aguardando validar o campo do transportador')
    if procura_imagem(imagem='img_topcon/transportador_incorreto.png', continuar_exec= True) is not False:
        print('--- Transportador incorreto!')
        bot.press('ENTER')
        bot.press('F2')
        marca_lancado(texto_marcacao='RE_Invalido')
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
        time.sleep(0.25)

 
    # * -------------------------------------- Aba Produtos e serviços --------------------------------------
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 25)    
    bot.doubleClick(procura_imagem(imagem='img_topcon/produtos_servicos.png'))
    #Aguarda até aparecer o botão "alterar"
    procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 100)
    
    if '38953477000164' in chave_xml: #Caso não tenha o CNPJ da Consmar
        exit(bot.alert('CNPJ da consmar, necessario scriptar'))
 
    #* Realiza a extração da quantidade de toneladas
    valor_escala = 200
    while True:
        while True:
            try:
                qtd_ton = extrai_txt_img(imagem='img_toneladas.png', area_tela=(892, 577, 70, 20), porce_escala= valor_escala).strip()
                qtd_ton = qtd_ton.replace(",", ".")
                qtd_ton = float(qtd_ton)
            except ValueError:
                valor_escala += 10
            else:
                print(F'--- Texto coletado da quantidade: {qtd_ton}, Valor escala: {valor_escala}')
                break

        print('--- Abrindo a tela "Itens nota fiscal de compra" ')
        bot.click(procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400)))
        while procura_imagem(imagem='img_topcon/valor_cofins.png', limite_tentativa= 1, continuar_exec= True) is False:
            print('--- Aguardando aparecer a tela "Itens nota fiscal de compra" ')

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
            bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
            break
            
        bot.click(procura_imagem(imagem='img_topcon/confirma.png'))            
        if procura_imagem(imagem='img_topcon/txt_ErroAtribuida.png', limite_tentativa = 12, continuar_exec = True) is False:
            print(Fore.GREEN + '--- Preenchimento completo, saindo do loop.' + Style.RESET_ALL)
            break
        else:
            print(Fore.RED + F'--- Falha, executando novamente a coleta das toneladas. Escala atual: {valor_escala}' + Style.RESET_ALL)
            valor_escala += 10
            while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True) is not False:
                bot.press('ENTER')
                bot.press('ESC')
                time.sleep(2)
            
        while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True) is not False:
            tentativa += 1
            print('--- Aguardando fechamento da tela do botão "Alterar" ')
            time.sleep(0.3)
            #TODO --- VerificaR se apareceu a tela "quantidade atribuida aos locais"

            if tentativa > 10: #Executa o loop 10 vezes até dar erro.
                exit(bot.alert('Apresentou algum erro.'))
        # TODO --- CASO O REMOTE APP DESCONECTE, RODAR O ABRE TOPCON

    # Conclui o lançamento
    bot.press('pagedown')  # Conclui o lançamento
    print('--- Enviado pagedown, aguardando tela de operação realizada')
    time.sleep(2)
    
    while procura_imagem(imagem='img_topcon/txt_inclui.png', continuar_exec= True, area= (852, 956, 1368, 1045)) is False:
        # Espera até aparecer a tela de operação realizada ou chave_invalida
        while procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec=True) is False:
            print('--- Aguardando a tela de operação realizada')
            if procura_imagem(imagem='img_topcon/chave_invalida.png', continuar_exec=True) is not False:
                print('--- Nota já lançada, marcando planilha!')
                bot.press('ENTER')
                bot.press('F2', presses = 2)
                marca_lancado(texto_marcacao='Lancado_Manual')
        else: #Caso encontre a opção
            print('--- Encontrou a tela de operação realizada, fechando e marcando a planilha')
            while True:
                if procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec= True) is not False:
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    ahk.win_wait_active('TopCompras', timeout=50, title_match_mode=2)
                    bot.click(procura_imagem(imagem='img_topcon/operacao_realizada.png'))
                    bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg'))
                    while procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True) is not False:
                        bot.click(procura_imagem(imagem='img_topcon/botao_ok.jpg'))
                else:
                    break
                
        #Verifica se apareceu a tela de transferencia 
        time.sleep(2)
        if procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True, limite_tentativa= 12, confianca= 0.7):
            print('--- Encontrou a tela do processo de transferencia')
            bot.click(procura_imagem('img_topcon/bt_sim.png', continuar_exec=True))
            while True:  # Aguardar o .PDF
                try:
                    ahk.win_wait('.pdf', title_match_mode=2, timeout= 15)
                except TimeoutError:
                    print('--- Aguardando .PDF da transferencia')
                else:
                    ahk.win_activate('.pdf', title_match_mode=2)
                    ahk.win_close('pdf - Google Chrome', title_match_mode=2)
                    print('--- Fechou o PDF da transferencia')
                    break
            time.sleep(0.4)
            ahk.win_activate('Transmissão', title_match_mode=2)
            bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))
            ahk.win_wait_active('TopCom', timeout=10, title_match_mode=2)
            ahk.win_activate('TopCom', title_match_mode=2)
    else:
        print('--- Lançamento concluido com sucesso')
    # * -------------------------------------- Marca planilha --------------------------------------
    marca_lancado(texto_marcacao='Lancado_RPA')
    return True


if __name__ == '__main__':
    while True:
        programa_principal()

# TODO --- Caso NFE Faturada no final do mes, lançar com qual data? 