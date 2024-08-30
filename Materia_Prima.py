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
import logging
import platform
import pytesseract
import pyautogui as bot

from ahk import AHK
from datetime import date
from valida_pedido import valida_pedido
from valida_lancamento import valida_lancamento
from abre_topcon import abre_mercantil, abre_topcon
from automacao.finaliza_lancamento import finaliza_lancamento
from utils.funcoes import marca_lancado, procura_imagem, extrai_txt_img


#* Definição de parametros
ahk = AHK()
posicao_img = 0
continuar = True
tempo_inicio = time.time()
bot.LOG_SCREENSHOTS = True  
bot.LOG_SCREENSHOTS_LIMIT = 5
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
          

def programa_principal():
    bot.PAUSE = 0.3
    acabou_pedido = False
    tentativa = 0

    print('---------------------------------------------------------------------------------------------------')
    print('--- INICIANDO UM NOVO LANÇAMENTO DE NFE --- ')
    print('---------------------------------------------------------------------------------------------------')
    while acabou_pedido is False: # Verifica se o pedido está valido.
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
        time.sleep(1)
    else:
        logging.info('--- Pedido validado, retornando para o programa principal' )

#* -------------------------- PROSSEGUINDO COM O LANÇAMENTO DA NFE -------------------------- 
    logging.info('--- Preenchendo dados na tela principal do lançamento')
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
    logging.info('--- Preenchendo filial de estoque')
    bot.write(filial_estoq)
    bot.press('TAB', presses= 2) # Confirma a informação da nova filial de estoque
    
    # Alteração da data
    logging.info('--- Realizando validação/alteração da data')
    hoje = date.today()
    hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
    #bot.write('10/08/2024')
    bot.press('ENTER')
    
    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
    
    # Caso o sistema informe que a data deve ser maior/igual a data inserida acima.
    if procura_imagem('img_topcon/data_invalida.png', continuar_exec= True):
        logging.warning('--- Precisa mudar a data, inserindo a data de hoje')
        bot.press('enter')          
        bot.write(hoje)
        bot.press('enter')
        time.sleep(0.25)
        # Aguarda até o topcompras voltar a funcionar
        ahk.win_activate('TopCompras', title_match_mode= 2)
        ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)

    try: # Aguarda a tela de erro do TopCon 
        ahk.win_wait('Topsys', title_match_mode= 2, timeout= 3)
    except TimeoutError:
        pass
    else:
        if ahk.win_exists('Topsys', title_match_mode= 2):
            ahk.win_activate('Topsys', title_match_mode= 2)
            logging.warning('--- Precisa mudar a data')
            bot.press('enter')          
            bot.write(hoje)
            bot.press('enter')
            time.sleep(0.25)

    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
    
    logging.info(F'--- Trocando o centro de custo para {centro_custo}')
    bot.write(centro_custo)
    ahk.win_activate('TopCompras', title_match_mode= 2)
    logging.info('--- Aguarda aparecer o campo cod_desc')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is False:
        if tentativa_cod_desc >= 100:
            logging.info('--- Não foi possivel encontrar o campo cod_desc, reiniciando o processo.')
            time.sleep(1)
            abre_mercantil()
            return True
        else:
            # Aguarda até o topcompras voltar a funcionar
            ahk.win_activate('TopCompras', title_match_mode= 2)
            ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
            tentativa_cod_desc += 1 
    else:
        logging.info(F'--- Apareceu o campo COD_DESC, tentativa: {tentativa_cod_desc} ')
        bot.press('ENTER') # Pressiona enter, e aguarda sumir o campo "cod_desc"
        
    logging.info('--- Aguarda até SUMIR o campo "cod_desc"')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is not False:
        bot.click(procura_imagem(imagem='img_topcon/txt_ValoresTotais.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74))
        #logging.info(F'--- Tentativa de aguardar sumir o cod_desc: {tentativa_cod_desc}')
        if tentativa_cod_desc >= 100:
            logging.info('--- O campo cod_desc não sumiu, reiniciando o processo.')
            time.sleep(1)
            abre_mercantil()
            return True
        else:
            # Aguarda até o topcompras voltar a funcionar
            ahk.win_activate('TopCompras', title_match_mode= 2)
            ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
            tentativa_cod_desc += 1 
    else:
        logging.info(F'--- sumiu o campo "cod_desc", tentativa: {tentativa_cod_desc}')

    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 100)
    bot.click(procura_imagem(imagem='img_topcon/txt_ValoresTotais.png', continuar_exec= True))

    # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
    logging.info(F'--- Preenchendo transportador: {cracha_mot}')
    ahk.win_activate('TopCompras', title_match_mode= 2)
    time.sleep(1)
    bot.click(procura_imagem(imagem='img_topcon/campo_000.png', continuar_exec= True))
    time.sleep(1)
    bot.press('tab')
    time.sleep(1)
    tentativa_achar_camp_re = 0
    while procura_imagem(imagem='img_topcon/campo_re_0.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
        logging.info(F'Tentativa: {tentativa_achar_camp_re}')
        time.sleep(0.25)
        tentativa_achar_camp_re += 1
        if tentativa_achar_camp_re >= 10:
            logging.info('--- Limite de tentativas de achar o campo "RE", reabrindo topcompras e reiniciando o processo.')
            time.sleep(1)
            abre_mercantil()
            return True
    else:
        logging.info('--- Campo RE habilitado, preenchendo.')
        # Preenche o campo do transportador e verifica se aconteceu algum erro.
        bot.write(cracha_mot)  # ID transportador
        time.sleep(0.25)
        bot.press('enter')

    logging.info('--- Aguardando validar o campo do transportador')
    ahk.win_activate('TopCompras', title_match_mode=2)
    if procura_imagem(imagem='img_topcon/transportador_incorreto.png', continuar_exec= True) is not False:
        logging.info('--- Transportador incorreto!')
        bot.press('ENTER')
        bot.press('F2')
        marca_lancado(texto_marcacao='RE_Invalido')
        programa_principal()
    else:
        logging.info('--- Transportador validado! Prosseguindo para validação da placa')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.press('enter')

    # Verifica se o campo da placa ficou preenchido
    time.sleep(1)
    if procura_imagem('img_topcon/campo_placa.png', confianca= 0.74, continuar_exec=True) is not False:
        logging.info('--- Encontrou o campo vazio, inserindo XXX0000')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.click(procura_imagem('img_topcon/campo_placa.png', continuar_exec=True))
        bot.write('XXX0000')
        bot.press('ENTER')
        time.sleep(1)
    else:
        logging.info('--- Não achou o campo ou já está preenchido')

    # * -------------------------------------- Aba Produtos e serviços --------------------------------------
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 15)
    logging.info('--- Navegando para a aba Produtos e Servicos')
    tela_prod_servico = 0
    while procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 1, continuar_exec= True, confianca= 0.74) is False:
        if tela_prod_servico > 15:
            return True
        
        bot.click(procura_imagem(imagem='img_topcon/produtos_servicos.png', confianca= 0.74, limite_tentativa= 3, continuar_exec= True))
        # Aguarda até aparecer o botão "alterar"
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        tela_prod_servico += 1
    
    
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
                logging.warning(F'--- Texto coletado da quantidade: {qtd_ton}, Valor escala: {valor_escala}')
                break

        logging.info('--- Abrindo a tela "Itens nota fiscal de compra" ')
        bot.click(procura_imagem(imagem='img_topcon/botao_alterar.png', area=(100, 839, 300, 400)))
        while procura_imagem(imagem='img_topcon/valor_cofins.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
            logging.info('--- Aguardando aparecer a tela "Itens nota fiscal de compra" ')
        
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        time.sleep(0.25)
        logging.info('--- Preenchendo SILO e quantidade')
        if ('SILO' in silo1) or ('SILO' in silo2):
            bot.click(851, 443)  # Clica na linha para informar o primeiro silo
            if ('SILO' in silo1) and ('SILO' in silo2):  
                qtd_ton = str((qtd_ton / 2)) # Realiza a divisão da quantidade de cimento, pois será distribuido em dois silos!
                qtd_ton = qtd_ton.replace(".", ",")
                logging.info(F'--- Foi informado dois silos, preenchendo... {silo1} e {silo2}, quantidade: {qtd_ton}')
                bot.write(silo1)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
                bot.write(silo2)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
            else:
                logging.info(F'--- Foi informado UM silo, preenchendo... {silo1}, quantidade: {qtd_ton}')
                qtd_ton = str(qtd_ton)
                qtd_ton = qtd_ton.replace(".", ",")
                bot.write(silo1)
                bot.write(silo2)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
        else: # Caso não tenha coletado nenhum silo.            
            if procura_imagem(imagem='img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True):
                logging.info('--- Não foi informado nenhum SILO, porém a nota é de cimento!' )
                bot.click(procura_imagem(imagem='img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True))
                bot.press('ESC')
                time.sleep(0.25)
                marca_lancado(texto_marcacao= 'Faltou_InfoSilo')
                return True
            else: # Caso realmente seja de agregado.
                logging.info('--- Nota de agregado, continuando o processo!')
            
            bot.click(procura_imagem(imagem='img_topcon/confirma.png'))
            break

        bot.click(procura_imagem(imagem='img_topcon/confirma.png'))            
        if procura_imagem(imagem='img_topcon/txt_ErroAtribuida.png', limite_tentativa = 6, continuar_exec = True) is False:
            logging.info('--- Preenchimento completo, saindo do loop.' )
            break
        else:
            logging.warning(F'--- Falha, executando novamente a coleta das toneladas. Escala atual: {valor_escala}' )
            valor_escala += 10
            while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
                bot.press('ENTER')
                bot.press('ESC')
                time.sleep(0.25)
            
        while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
            tentativa += 1
            logging.info('--- Aguardando fechamento da tela do botão "Alterar" ')
            time.sleep(0.25)
            #TODO --- VerificaR se apareceu a tela "quantidade atribuida aos locais"

            if tentativa > 10: #Executa o loop 10 vezes até dar erro.
                exit(bot.alert('Apresentou algum erro.'))
    
    finaliza_lancamento() # Realiza todo o processo de finalização de lançamento.
    return True


if __name__ == '__main__':
    logging.basicConfig(
        filename="automacao.log",
        filemode="a",
        encoding="utf-8",
        level=logging.INFO,
        format="{asctime} - {levelname} - {message}",
        style="{",
        )
    
    logging.info('\n------------------------ Iniciando um novo log ------------------------ ')
    os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
    os.system('cls')
    
    if 'VLPTIC1Z9HD33' in platform.node(): # Verifica qual sistema está rodando o script
        bot.FAILSAFE = True
    else:
        bot.FAILSAFE = False
        
    while True:
        try:
            programa_principal()
        #except (TimeoutError, OSError, ValueError):
        except (TimeoutError, ValueError, OSError):
            #os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
            time.sleep(5)
            abre_topcon()
            logging.critical("A execução principal acusou algum erro ( TimeoutError ), executando o script inteiro novamente.")

# TODO --- Caso NFE Faturada no final do mes, lançar com qual data?