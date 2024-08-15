# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EW_8FZwWFYVAol4MpV1GglkBJEaJaDx6cfuClnesIu60Ng?e=pveECF
# Debug db alltrips
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg
# db_alltrips no paulo, apenas leitura
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg?e=RYfgcA

import os
import time
import platform
import pytesseract
import pyautogui as bot

from ahk import AHK
from colorama import Style, Fore
from datetime import date, timedelta
from valida_pedido import valida_pedido
from valida_lancamento import valida_lancamento
from abre_topcon import abre_mercantil, abre_topcon
from funcoes import marca_lancado, procura_imagem, extrai_txt_img, corrige_nometela

# --- Definição de parametros
ahk = AHK()
bot.LOG_SCREENSHOTS = True
bot.LOG_SCREENSHOTS_LIMIT = 5
posicao_img = 0
continuar = True
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"


def processo_transferencia():
    if procura_imagem(imagem='img_topcon/txt_transferencia.png', continuar_exec= True, limite_tentativa= 2, confianca= 0.74):
        print(Fore.GREEN + '\n--- Iniciando a função: processo transferencia ---' + Style.RESET_ALL)
        ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
        ahk.win_wait_active('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2, timeout= 30)
        
        if procura_imagem('img_topcon/deseja_processar.png', continuar_exec=True, confianca= 0.75):
            print('--- Encontrou a tela "Deseja processar NFE" ')
            while procura_imagem('img_topcon/bt_sim.png', continuar_exec=True, limite_tentativa= 3, confianca= 0.74):
                bot.click(procura_imagem('img_topcon/bt_sim.png', continuar_exec=True))
                
            contador_pdf = 0
            while True:  # Aguardar o .PDF
                time.sleep(0.2)
                try:
                    ahk.win_wait('.pdf', title_match_mode=2, timeout= 8)
                except TimeoutError:
                    if contador_pdf >= 15:
                        # Fechando a tela de transmissão
                        while ahk.win_exists('Transmissão', title_match_mode= 2):
                            ahk.win_activate('Transmissão', title_match_mode=2)
                            bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))
                            ahk.win_wait_close('Transmissão', title_match_mode=2, timeout= 15)
                            
                            ahk.win_wait_active('TopCompras', timeout=10, title_match_mode=2)
                            ahk.win_activate('TopCompras', title_match_mode=2)
                            return True

                    contador_pdf += 1
                    print('--- Aguardando .PDF da transferencia')
                    
                else:
                    ahk.win_activate('.pdf', title_match_mode=2)
                    ahk.win_close('pdf - Google Chrome', title_match_mode=2)
                    print('--- Fechou o PDF da transferencia')
                    time.sleep(0.5)
                    break
            
            # Fechando a tela de transmissão
            while ahk.win_exists('Transmissão', title_match_mode= 2):
                ahk.win_activate('Transmissão', title_match_mode=2)
                bot.click(procura_imagem(imagem='img_topcon/sair_tela.png'))
                
                ahk.win_wait_active('TopCompras', timeout=10, title_match_mode=2)
                ahk.win_activate('TopCompras', title_match_mode=2)
                return True
        else:
            print('--- Não encontrou a tela "Deseja processar NFE ainda')
        
def finaliza_lancamento(planilha_marcada = False, lancamento_concluido = False, realizou_transferencia = False, tentativas_telas = 0):
    print(Fore.GREEN + '\n--- Iniciando a função de finalização de lançamento, enviando PAGEDOWN ---' + Style.RESET_ALL)
    ahk.win_activate('TopCompras', title_match_mode=2)
    bot.press('pagedown')  # Conclui o lançamento
    
    while True:
        ahk.win_activate('TopCompras', title_match_mode=2) # Para manter o TopCompras aberto.

        if ahk.win_exists('CsjTb', title_match_mode= 2): # Caso apareça a tela de campo obrigatorio (Aparece quando não preencher nenhum campo.)
            ahk.win_close('CsjTb', title_match_mode= 2)
            abre_mercantil()
            print('--- Reabriu o mercantil, recomeçando o processo.')
            programa_principal()
        
        # 0. Verifica se ocorreu algo de transferencia
        realizou_transferencia = processo_transferencia()
        time.sleep(1)
        # 1. Caso chave invalida.  
        if procura_imagem(imagem='img_topcon/chave_invalida.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            print('--- Nota já lançada, marcando planilha!')
            bot.press('ENTER')
            bot.press('F2', presses = 2)
            marca_lancado(texto_marcacao='Lancado_Manual')
            break

        # 2. Caso operação realizada.
        if procura_imagem(imagem='img_topcon/operacao_realizada.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is not False:
            if planilha_marcada is False:
                print('--- Encontrou a tela de operação realizada, fechando e marcando a planilha')
                marca_lancado(texto_marcacao='Lancado_RPA')
                planilha_marcada = True
            
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.click(procura_imagem(imagem='img_topcon/operacao_realizada.png'))
            bot.press('ENTER')
            #ahk.win_wait_close('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode= 2, timeout= 15)
                    
        elif planilha_marcada is True: # Essa parte só pode rodar, se encontrar a opção "operação realizada"
            print('--- Não encontrou a tela "operação realizada", porém a planilha está marcada!')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            
            #Validando se já fecharam todas as telas.
            if procura_imagem(imagem='img_topcon/bt_obslancamento.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is not False:
                print(F'--- Encontrou o botão "OBS. Lancamento." encerrando loop das telas, valor do realizou transf: {realizou_transferencia}')
                if realizou_transferencia is True:
                    print('--- Realizou transferencia, reabrindo o modulo do topcompras para evitar erros.')
                    time.sleep(0.1)
                    abre_mercantil()
                else: # Segue o processo a baixo.
                    # Retorna a tela para o modo localizar
                    ahk.win_activate('TopCompras', title_match_mode= 2)
                    bot.press('F3', presses = 1)
                    bot.press('F2', presses = 1)
                    time.sleep(0.1)
                    
                    if procura_imagem(imagem='img_topcon/txt_localizar.png', continuar_exec= True, area= (852, 956, 1368, 1045)):
                        print('--- Entrou no modo localizar, lançamento realmente concluido!')
                        lancamento_concluido = True
                        return True
                    
        # 3. Caso apareça "deseja imprimir o espelho da nota?"
        if procura_imagem(imagem='img_topcon/txt_espelhonota.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            print('--- Apareceu a tela: deseja imprimir o espelho da nota?')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('ENTER')
            ahk.win_activate('Espelho de Nota Fiscal', title_match_mode= 2)
            ahk.win_wait('Espelho de Nota Fiscal', title_match_mode= 2, timeout= 30)
        
        # 4. Caso apareça tela "Espelho da nota fiscal"
        while ahk.win_exists('Espelho de Nota Fiscal', title_match_mode= 2):
            ahk.win_close('Espelho de Nota Fiscal', title_match_mode= 2)

        if lancamento_concluido is True:
            time.sleep(1)
            ahk.win_activate('TopCompras', title_match_mode= 2)
            bot.press('F2') # Aperta F2 para retornar a tela para o modo "Localizar"
            marca_lancado(texto_marcacao='Lancado_RPA')
        
        # Caso exceta o limite de tentativas, tenta fechar e abrir a tela de compras.
        if tentativas_telas >= 10:
            print(Fore.RED + F'--- Excedeu o limite de tentativas de encontrar as telas, reabrindo o TopCompras, tentativa: {tentativas_telas}' + Style.RESET_ALL)
            time.sleep(1)
            abre_mercantil()
            return False # Retorna False pois o lançamento não foi concluido
        else:
            time.sleep(tentativas_telas * 0.2)
            print(F'--- Não encontrou nenhuma das telas do processo finaliza lançamento, executando novamente, {tentativas_telas}, tempo pausa: {tentativas_telas * 0.5}')
            ahk.win_activate('TopCompras', title_match_mode= 2)
            tentativas_telas += 1
        
        # 6. Caso apareça o erro de vencimento
        if procura_imagem(imagem='img_topcon/txt_vencimento.PNG', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            ahk.win_activate('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
            print('--- Apareceu a tela de vencimento, alterando para +3 dias')
            bot.press('ENTER')
            ahk.win_wait_close('TopCompras (VM-CortesiaApli.CORTESIA.com)', title_match_mode=2)
            time.sleep(0.5)
            # Altera a data de vencimento para +3 dias
            bot.click(procura_imagem(imagem='img_topcon/bt_contasapagar.PNG'))
            bot.click(procura_imagem(imagem='img_topcon/bt_datavencimento.PNG', area= (419, 536, 811, 715)))
            data_vencimento = date.today() + timedelta(3)
            data_vencimento = data_vencimento.strftime("%d%m%y")
            bot.write(data_vencimento)
            bot.press('ENTER')
            time.sleep(1)
            bot.press('pagedown')  # Conclui o lançamento
    
         
def programa_principal():
    bot.pause = 0.1
    acabou_pedido = True
    tentativa = 0
    contagem_validalancamento = 0
    print('---------------------------------------------------------------------------------------------------')
    print('--- INICIANDO UM NOVO LANÇAMENTO DE NFE --- ')
    print('---------------------------------------------------------------------------------------------------')
    
    while acabou_pedido is True: #Verifica se o pedido está valido.
        while contagem_validalancamento < 2: # Necessario para coletar os erros de não encontrar a tela do topcon (Timeout)
            try:
                dados_planilha = valida_lancamento()
            except TimeoutError:
                print(Fore.CYAN + '--- VALIDA LANCAMENTO deu timeouterroor' + Style.RESET_ALL)
                abre_topcon()
                contagem_validalancamento += 1
            else:
                break
 
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
        
        # Prrocesso de validação do pedido.
        acabou_pedido = valida_pedido(acabou_pedido=False)
        if acabou_pedido is False:
            #os.system('cls')
            print('--- Pedido não validado!')
        else:
            print(Fore.GREEN + '\n--- Pedido validado, retornando para o programa principal' + Style.RESET_ALL)

#* -------------------------- PROSSEGUINDO COM O LANÇAMENTO DA NFE -------------------------- 
    print('--- Preenchendo dados na tela principal do lançamento')
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 10)
    time.sleep(0.2)
    
    while procura_imagem(imagem='img_topcon/produtos_servicos.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
        time.sleep(0.2)
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
        try:
            ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 10)
        except TimeoutError:
            bot.alert(exit('Topcompras não encontrado'))
        time.sleep(0.2)

    bot.press('up')
    print('--- Preenchendo filial de estoque')
    bot.write(filial_estoq)
    bot.press('TAB', presses= 2) # Confirma a informação da nova filial de estoque
    
    
    # Alteração da data
    print('--- Realizando validação/alteração da data')
    hoje = date.today()
    hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
    #bot.write('10/08/2024')
    bot.press('ENTER')
    
    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
    
    # Caso o sistema informe que a data deve ser maior/igual a data inserida acima.
    if procura_imagem('img_topcon/data_invalida.png', continuar_exec= True):
        print('--- Precisa mudar a data, inserindo a data de hoje')
        bot.press('enter')          
        bot.write(hoje)
        bot.press('enter')
        time.sleep(0.1)
        # Aguarda até o topcompras voltar a funcionar
        ahk.win_activate('TopCompras', title_match_mode= 2)
        ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)

    try:
        ahk.win_wait('Topsys', title_match_mode= 2, timeout= 3)
    except TimeoutError:
        pass
        #print('--- Tela de erro NÃO apareceu, continuando...')
    else:
        if ahk.win_exists('Topsys', title_match_mode= 2):
            ahk.win_activate('Topsys', title_match_mode= 2)
            print('--- Precisa mudar a data')
            bot.press('enter')          
            bot.write(hoje)
            bot.press('enter')
            time.sleep(0.1)

    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
    
    print(F'--- Trocando o centro de custo para {centro_custo}')
    bot.write(centro_custo)
    ahk.win_activate('TopCompras', title_match_mode= 2)
    print('--- Aguarda aparecer o campo cod_desc')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is False:
        if tentativa_cod_desc >= 100:
            print('--- Não foi possivel encontrar o campo cod_desc, reiniciando o processo.')
            time.sleep(0.5)
            abre_mercantil()
            return True
        else:
            # Aguarda até o topcompras voltar a funcionar
            ahk.win_activate('TopCompras', title_match_mode= 2)
            ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
            tentativa_cod_desc += 1 
    else:
        print(F'--- Apareceu o campo COD_DESC, tentativa: {tentativa_cod_desc} ')
        bot.press('ENTER') # Pressiona enter, e aguarda sumir o campo "cod_desc"
        
    print('--- Aguarda até SUMIR o campo "cod_desc"')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is not False:
        bot.click(procura_imagem(imagem='img_topcon/txt_ValoresTotais.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74))
        #print(F'--- Tentativa de aguardar sumir o cod_desc: {tentativa_cod_desc}')
        if tentativa_cod_desc >= 100:
            print('--- O campo cod_desc não sumiu, reiniciando o processo.')
            time.sleep(0.5)
            abre_mercantil()
            return True
        else:
            # Aguarda até o topcompras voltar a funcionar
            ahk.win_activate('TopCompras', title_match_mode= 2)
            ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
            tentativa_cod_desc += 1 
    else:
        print(F'--- sumiu o campo "cod_desc", tentativa: {tentativa_cod_desc}')

    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
    bot.click(procura_imagem(imagem='img_topcon/txt_ValoresTotais.png', continuar_exec= True))

    # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
    print(F'--- Preenchendo transportador: {cracha_mot}')
    ahk.win_activate('TopCompras', title_match_mode= 2)
    bot.click(procura_imagem(imagem='img_topcon/campo_000.png', continuar_exec= True))
    bot.press('tab')
    tentativa_achar_camp_re = 0
    while procura_imagem(imagem='img_topcon/campo_re_0.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
        print(F'Tentativa: {tentativa_achar_camp_re}')
        time.sleep(0.1)
        tentativa_achar_camp_re += 1
        if tentativa_achar_camp_re >= 10:
            print('--- Limite de tentativas de achar o campo "RE", reabrindo topcompras e reiniciando o processo.')
            time.sleep(0.5)
            abre_mercantil()
            return True
    else:
        print('--- Campo RE habilitado, preenchendo.')
        # Preenche o campo do transportador e verifica se aconteceu algum erro.
        bot.write(cracha_mot)  # ID transportador
        time.sleep(0.1)
        bot.press('enter')

    print('--- Aguardando validar o campo do transportador')
    ahk.win_activate('TopCompras', title_match_mode=2)
    if procura_imagem(imagem='img_topcon/transportador_incorreto.png', continuar_exec= True) is not False:
        print('--- Transportador incorreto!')
        bot.press('ENTER')
        bot.press('F2')
        marca_lancado(texto_marcacao='RE_Invalido')
        programa_principal()
    else:
        print('--- Transportador validado! Prosseguindo para validação da placa')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.press('enter')

    # Verifica se o campo da placa ficou preenchido
    time.sleep(0.5)
    if procura_imagem('img_topcon/campo_placa.png', confianca= 0.74, continuar_exec=True) is not False:
        print('--- Encontrou o campo vazio, inserindo XXX0000')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.click(procura_imagem('img_topcon/campo_placa.png', continuar_exec=True))
        bot.write('XXX0000')
        bot.press('ENTER')
        time.sleep(0.5)
    else:
        print('--- Não achou o campo ou já está preenchido')

    # * -------------------------------------- Aba Produtos e serviços --------------------------------------
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 15)
    print('--- Navegando para a aba Produtos e Servicos')
    tela_prod_servico = 0
    while procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 1, continuar_exec= True, confianca= 0.74) is False:
        if tela_prod_servico > 15:
            return True
        
        bot.click(procura_imagem(imagem='img_topcon/produtos_servicos.png', confianca= 0.74, limite_tentativa= 3, continuar_exec= True))
        # Aguarda até aparecer o botão "alterar"
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        tela_prod_servico += 1
    
    procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 16)
    
    if '38953477000164' in chave_xml: #Caso não tenha o CNPJ da Consmar
        finaliza_lancamento()
        return True
 
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
        while procura_imagem(imagem='img_topcon/valor_cofins.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
            print('--- Aguardando aparecer a tela "Itens nota fiscal de compra" ')
        
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        time.sleep(0.25)
        print('--- Preenchendo SILO e quantidade')
        if ('SILO' in silo1) or ('SILO' in silo2):
            bot.click(851, 443)  # Clica na linha para informar o primeiro silo
            if ('SILO' in silo1) and ('SILO' in silo2):  
                qtd_ton = str((qtd_ton / 2)) # Realiza a divisão da quantidade de cimento, pois será distribuido em dois silos!
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
            else:
                print(F'--- Foi informado UM silo, preenchendo... {silo1}, quantidade: {qtd_ton}')
                qtd_ton = str(qtd_ton)
                qtd_ton = qtd_ton.replace(".", ",")
                bot.write(silo1)
                bot.write(silo2)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
        else: # Caso não tenha coletado nenhum silo.            
            if procura_imagem(imagem='img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True):
                print(Fore.RED + '--- Não foi informado nenhum SILO, porém a nota é de cimento!' + Style.RESET_ALL)
                bot.click(procura_imagem(imagem='img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True))
                bot.press('ESC')
                time.sleep(0.1)
                marca_lancado(texto_marcacao= 'Faltou_InfoSilo')
                return True
            else: # Caso realmente seja de agregado.
                print('--- Nota de agregado, continuando o processo!')
            
            bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
            break
            
        bot.click(procura_imagem(imagem='img_topcon/confirma.png'))            
        if procura_imagem(imagem='img_topcon/txt_ErroAtribuida.png', limite_tentativa = 6, continuar_exec = True) is False:
            print(Fore.GREEN + '--- Preenchimento completo, saindo do loop.' + Style.RESET_ALL)
            break
        else:
            print(Fore.RED + F'--- Falha, executando novamente a coleta das toneladas. Escala atual: {valor_escala}' + Style.RESET_ALL)
            valor_escala += 10
            while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
                bot.press('ENTER')
                bot.press('ESC')
                time.sleep(0.1)
            
        while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            tentativa += 1
            print('--- Aguardando fechamento da tela do botão "Alterar" ')
            time.sleep(0.25)
            #TODO --- VerificaR se apareceu a tela "quantidade atribuida aos locais"

            if tentativa > 10: #Executa o loop 10 vezes até dar erro.
                exit(bot.alert('Apresentou algum erro.'))
    

    finaliza_lancamento() # Realiza todo o processo de finalização de lançamento.
    return True


if __name__ == '__main__':
    os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
    os.system('cls')
    
    if 'VLPTIC1Z9HD33' in platform.node(): # Verifica qual sistema está rodando o script
        bot.FAILSAFE = True
    else:
        bot.FAILSAFE = False
        
    #finaliza_lancamento()
    while True:
        try:
            programa_principal()
        except ValueError:
            os.system('taskkill /im AutoHotkey.exe /f /t')
            #exit(bot.alert('--- Erro por que não encontrou alguma das telas.'))
            abre_topcon()
            programa_principal()
        except TimeoutError:
            #exit(bot.alert('--- Erro por que não encontrou alguma das telas.'))
            abre_topcon()
            programa_principal()

# TODO --- Caso NFE Faturada no final do mes, lançar com qual data? 