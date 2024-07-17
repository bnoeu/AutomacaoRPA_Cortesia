# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
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
from valida_lancamento import valida_lancamento
from funcoes import marca_lancado, procura_imagem, extrai_txt_img, corrige_topcompras

# --- Definição de parametros
ahk = AHK()
bot.PAUSE = 0.5
posicao_img = 0
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

#* Funções

def processo_transferencia():
    if procura_imagem(imagem='img_topcon/txt_transferencia.png', continuar_exec= True, confianca= 0.7):
        print('--- Encontrou a tela de transferencia!')
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
                    time.sleep(0.4)
                    break
            
            # Fechando a tela de transmissão
            while ahk.win_exists('Transmissão', title_match_mode= 2):
                ahk.win_activate('Transmissão', title_match_mode=2)
                bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))
                
                ahk.win_wait_active('TopCom', timeout=10, title_match_mode=2)
                ahk.win_activate('TopCom', title_match_mode=2)
    else:
        print('--- Não foi gerada transferencia para essa nota fiscal.')

def abre_mercantil():
    # Inicia fechando o modulo de compras.
    ahk.win_close('TopCompras', title_match_mode=2)
    
    # Ativa o Topcon, e clica no topcompras, e executa a função para correção do nome.
    ahk.win_activate('TopCon', title_match_mode= 2)
    bot.click(procura_imagem(imagem='img_topcon/logo_topcompras.png'))
    time.sleep(2)
    corrige_topcompras()

    #Abre o TopCompras, e verifica se aparece a tela "interveniente"
    ahk.win_activate('TopCompras', title_match_mode= 2)
    if procura_imagem(imagem='img_topcon/botao_ok.jpg', continuar_exec= True):
        print('--- Encontrou a tela do interveniente, clicando no botão "OK"')
        bot.press('ENTER')
    else:
        print('--- Não exibiu a tela de interveniente.')
        
    #Navegando entre os menus para abrir a opção "Compras - Mercantil"
    ahk.win_activate('TopCompras', title_match_mode= 2)
    bot.press('ALT')
    bot.press('RIGHT', presses= 2, interval= 0.05)
    bot.press('DOWN', presses= 7, interval= 0.05)
    bot.press('ENTER')
    time.sleep(3)
    print(Fore.GREEN +  '--- TopCompras aberto!' + Style.RESET_ALL)

def finaliza_lancamento():
    lancamento_concluido = False
    tentativas_telas = 0
    while True:
        # Para manter o TopCompras aberto.
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 25)
        time.sleep(1)
        # 1. Caso chave invalida.  
        if procura_imagem(imagem='img_topcon/chave_invalida.png', continuar_exec=True) is not False:
            print('--- Nota já lançada, marcando planilha!')
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            marca_lancado(texto_marcacao='Lancado_Manual')
            break

        # 2. Caso operação realizada.
        if procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec= True) is not False:
            print('--- Encontrou a tela de operação realizada, fechando e marcando a planilha')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.click(procura_imagem(imagem='img_topcon/operacao_realizada.png'))
            bot.press('ENTER')
            time.sleep(2)
                    
        else:
            print('--- Não encontrou a tela "operação realizada" ')
            if procura_imagem(imagem='img_topcon/bt_obslancamento.png', continuar_exec= True) is not False:
                print('--- Encontrou o botão "OBS. Lancamento." encerrando loop das telas.')
                
                #Retorna a tela para o modo localizar
                bot.press('F2', presses = 2)
                if procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
                    print('--- Entrou no modo localizar, lançamento realmente concluido!')
                    lancamento_concluido = True
                    break
                else:
                    print(Fore.RED + '--- Não voltou para o modo localizar, alguma tela ainda deve estar aberta...\n' + Style.RESET_ALL)
                    
        # 3. Caso apareça "deseja imprimir o espelho da nota?"
        if procura_imagem(imagem='img_topcon/txt_espelhonota.png', continuar_exec=True) is not False:
            print('--- Apareceu a tela: deseja imprimir o espelho da nota?')
            bot.press('ENTER')
        
        # 4. Caso apareça tela "Espelho da nota fiscal"
        while ahk.win_exists('Espelho de Nota Fiscal', title_match_mode= 2):
            ahk.win_close('Espelho de Nota Fiscal', title_match_mode= 2)

        # 5. 
        processo_transferencia()
        time.sleep(1)
        # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        tentativas_telas += 1
        if tentativas_telas >= 5:
            abre_mercantil()
        
        if lancamento_concluido is True:
            time.sleep(2)
            bot.press('F2') # Aperta F2 para retornar a tela para o modo "Localizar"
            marca_lancado(texto_marcacao='Lancado_RPA')

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
        if acabou_pedido is False:
            os.system('cls')
            print(Fore.GREEN + F'--- Pedido validado, dados planilha: {dados_planilha}\n' + Style.RESET_ALL)

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
    print('--- Preenchendo filial de estoque')
    bot.write(filial_estoq)
    bot.press('TAB', presses= 2) # Confirma a informação da nova filial de estoque
    
    # Alteração da data
    hoje = date.today()
    hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
    #bot.write('13/07/2024')
    bot.press('ENTER')
    time.sleep(0.5)
    
    # Caso o sistema informe que a data deve ser maior/igual a data inserida acima.
    if procura_imagem('img_topcon/data_invalida.png', continuar_exec= True):
        print('--- Precisa mudar a data')
        bot.press('enter')          
        bot.write(hoje)
        bot.press('enter')
        time.sleep(0.2)

    try:
        ahk.win_wait('Topsys', title_match_mode= 2, timeout= 30)
    except TimeoutError:
        print('--- Tela de erro NÃO apareceu, continuando...')
    else:
        if ahk.win_exists('Topsys', title_match_mode= 2):
            ahk.win_activate('Topsys', title_match_mode= 2)
            print('--- Precisa mudar a data')
            bot.press('enter')          
            bot.write(hoje)
            bot.press('enter')
            time.sleep(0.2)

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
    while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True) is False:
        time.sleep(0.2)
    else:
        print('--- Apareceu o campo COD_DESC')
        bot.press('ENTER') # Pressiona enter, e aguarda sumir o campo "cod_desc"
        
    print('--- Aguarda até SUMIR o campo "cod_desc"')
    while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True) is not False:
        time.sleep(0.2)
    else:
        print('--- sumiu o campo "cod_desc" ')

    bot.click(procura_imagem(imagem='img_topcon/txt_ValoresTotais.png', continuar_exec= True))

    # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
    print(F'--- Preenchendo transportador: {cracha_mot}')
    bot.click(procura_imagem(imagem='img_topcon/campo_000.png', continuar_exec= True))
    time.sleep(1)
    bot.press('tab')
    time.sleep(1)
    while procura_imagem(imagem='img_topcon/campo_re_0.png', continuar_exec= True) is False:
        time.sleep(0.3)
    else:
        print('--- Campo RE habilitado, preenchendo.')
        time.sleep(1)
        # Preenche o campo do transportador e verifica se aconteceu algum erro.
        bot.write(cracha_mot)  # ID transportador
        time.sleep(1)
        bot.press('enter')
        time.sleep(1)

    print('--- Aguardando validar o campo do transportador')
    ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
    if procura_imagem(imagem='img_topcon/transportador_incorreto.png', continuar_exec= True) is not False:
        print('--- Transportador incorreto!')
        bot.press('ENTER')
        bot.press('F2')
        marca_lancado(texto_marcacao='RE_Invalido')
        programa_principal()
    else:
        print('--- Transportador validado! Prosseguindo para validação da placa')
        time.sleep(1)
        bot.press('enter')

    # Verifica se o campo da placa ficou preenchido
    ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
    if procura_imagem('img_topcon/campo_placa.png', continuar_exec=True) is not False:
        print('--- Encontrou o campo vazio, inserindo XXX0000')
        bot.click(procura_imagem('img_topcon/campo_placa.png', continuar_exec=True))
        bot.write('XXX0000')
        bot.press('ENTER')
    else:
        print('--- Não achou o campo ou já está preenchido')
        time.sleep(0.25)

    # * -------------------------------------- Aba Produtos e serviços --------------------------------------
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 25)
    print('--- Navegando para a aba Produtos e Servicos')  
    bot.doubleClick(procura_imagem(imagem='img_topcon/produtos_servicos.png'))
    time.sleep(2)
    # Aguarda até aparecer o botão "alterar"
    procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 100)
    
    if '38953477000164' in chave_xml: #Caso não tenha o CNPJ da Consmar
        exit(bot.alert('CNPJ da consmar, necessario scriptar'))
 
    #* Realiza a extração da quantidade de toneladas
    valor_escala = 200
    while True:
        while True: # Realiza a extração das toneladas.
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
        if ('SILO' in silo1) or ('SILO' in silo2):
            bot.click(851, 443)  # Clica na linha para informar o primeiro silo
            if silo2 != '':  # Realiza a divisão da quantidade de cimento
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
            elif silo1 != '' and silo2 == '':
                print(F'--- Foi informado UM silo, preenchendo... {silo1}, quantidade: {qtd_ton}')
                qtd_ton = str(qtd_ton)
                qtd_ton = qtd_ton.replace(".", ",")
                bot.write(silo1)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
        else: # Caso não tenha coletado nenhum silo.            
            if procura_imagem(imagem='img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True):
                print(Fore.RED + '--- Não foi informado nenhum SILO, porém a nota é de cimento!\n' + Style.RESET_ALL)
                bot.click(procura_imagem(imagem='img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True))
                bot.press('ESC')
                time.sleep(1)
                marca_lancado(texto_marcacao= 'Faltou_InfoSilo')
                programa_principal()
            else: # Caso realmente seja de agregado.
                print('--- Nota de agregado, continuando o processo!')
            
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
    print('--- Enviado pagedown, aguardando tela de operação realizada')
    bot.press('pagedown')  # Conclui o lançamento
    time.sleep(3)
    
    # Realiza todo o processo de finalização de lançamento.
    finaliza_lancamento()
    return True


if __name__ == '__main__':
    #abre_mercantil()
    #finaliza_lancamento()
    while True:
        programa_principal()

# TODO --- Caso NFE Faturada no final do mes, lançar com qual data? 