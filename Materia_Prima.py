# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EW_8FZwWFYVAol4MpV1GglkBJEaJaDx6cfuClnesIu60Ng?e=pveECF
# Debug db alltrips
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg
# db_alltrips no paulo, apenas leitura
#  

import os
import time
import platform
import traceback
import pytesseract
import pyautogui as bot
from utils.funcoes import ahk as ahk
from datetime import date, datetime
from abre_topcon import abre_topcon
from valida_pedido import valida_pedido
from utils.enviar_email import enviar_email
from utils.configura_logger import get_logger
from valida_lancamento import valida_lancamento
from automacao.finaliza_lancamento import finaliza_lancamento
from utils.funcoes import marca_lancado, procura_imagem, extrai_txt_img


#* Definição de parametros
posicao_img = 0
continuar = True
qtd_notas_lancadas = 0
tempo_inicio = time.time()
bot.LOG_SCREENSHOTS = True  
bot.LOG_SCREENSHOTS_LIMIT = 5
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"

def valida_filial_estoque(filial_estoq = ""):
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
        exit(F'Filial de estoque não padronizada: {filial_estoq}')
    
    if centro_custo != "":
        return centro_custo

def verifica_horario():
    while True:
        hora_atual = datetime.now().time() # Obter o horário atual
        hora_inicio_pausa = datetime.strptime("00:00", "%H:%M").time() # Definir o horário de inicio de referência (02:00)
        hora_final_pausa = datetime.strptime("01:00", "%H:%M").time() # Definir o horário de inicio de referência (02:00)

        if hora_atual > hora_inicio_pausa and hora_atual < hora_final_pausa:
            logger.info(F'--- São: {hora_atual}, aguardando 1 hora para tentar novamente.')
            time.sleep(7200)
        else:
            return

def programa_principal():
    global qtd_notas_lancadas
    bot.PAUSE = 0.6
    acabou_pedido = False
    tentativa = 0

    logger.info("\n\n\n")
    logger.info('---------------------------------------------------------------------------------------------------')
    logger.info('--- INICIANDO UM NOVO LANÇAMENTO DE NFE --- ')
    logger.info('---------------------------------------------------------------------------------------------------')

    verifica_horario() # Confere o horario dessa execução.

    while acabou_pedido is False: # Realiza a validação do pedido
        dados_planilha = valida_lancamento() # Coleta e confere os dados do lançamento atual
        # Passa todos osdados parasuas variaveis.
        cracha_mot = dados_planilha[0]
        silo1 = dados_planilha[1]
        silo2 = dados_planilha[2]
        filial_estoq = dados_planilha[3].split('-') # Recebe por exemplo: ['1001', 'VILA PRUDENTE']
        filial_estoq = filial_estoq[0] # O dado é passado assim: ['1001', 'VILA PRUDENTE'], aqui formata para '1001'
        centro_custo = valida_filial_estoque(filial_estoq) # Realiza a validação da filial de estoque.
        chave_xml = dados_planilha[4]
        acabou_pedido = valida_pedido() # Verifica se o pedido está valido.

    #bot.alert('Teste!')
#* -------------------------- Continua o processo de lançamento da NFE -------------------------- 
    logger.info('--- Preenchendo dados na tela principal do lançamento')
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 10)

    #* Aguarda até aparecer o botão "Produtos e serviços", isso valida que fechou a tela de vinculação de pedido
    while procura_imagem(imagem='imagens/img_topcon/produtos_servicos.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
        time.sleep(0.4)
        ahk.win_activate('TopCompras', title_match_mode=2, detect_hidden_windows= True)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 10)

    logger.info('--- Preenchendo filial de estoque')
    bot.press('up')
    bot.write(filial_estoq)
    bot.press('TAB', presses= 2) # Confirma a informação da nova filial de estoque
    
    #* Alteração da data
    logger.info('--- Realizando validação/alteração da data')
    hoje = date.today()
    hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
    bot.press('ENTER')
    
    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 70)
    
    # Caso o sistema informe que a data deve ser maior/igual a data inserida acima.
    if procura_imagem('imagens/img_topcon/data_invalida.png', continuar_exec= True):
        logger.warning('--- Precisa mudar a data, inserindo a data de hoje')
        bot.press('enter')          
        bot.write(hoje)
        bot.press('enter')
        time.sleep(0.25)
        # Aguarda até o topcompras voltar a funcionar
        ahk.win_activate('TopCompras', title_match_mode= 2)
        ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 70)

    try: # Aguarda a tela de erro do TopCon 
        ahk.win_wait('Topsys', title_match_mode= 2, timeout= 3)
    except TimeoutError:
        pass
    else:
        if ahk.win_exists('Topsys', title_match_mode= 2):
            ahk.win_activate('Topsys', title_match_mode= 2)
            logger.warning('--- Precisa mudar a data')
            bot.press('enter')          
            bot.write(hoje)
            bot.press('enter')
            time.sleep(0.25)

    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 70)
    
    logger.info(F'--- Trocando o centro de custo para {centro_custo}')
    bot.write(centro_custo)
    ahk.win_activate('TopCompras', title_match_mode= 2)
    logger.info('--- Aguarda aparecer o campo cod_desc')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='imagens/img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is False:
        time.sleep(0.25)
        if tentativa_cod_desc >= 100:
            logger.info('--- Não foi possivel encontrar o campo cod_desc, reiniciando o processo.')
            time.sleep(0.5)
            abre_topcon()
            return True
        else: # Aguarda até o topcompras voltar a funcionar
            ahk.win_activate('TopCompras', title_match_mode= 2)
            ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 70)
            tentativa_cod_desc += 1 
    else:
        logger.info(F'--- Apareceu o campo COD_DESC, tentativa: {tentativa_cod_desc} ')
        bot.press('ENTER') # Pressiona enter, e aguarda sumir o campo "cod_desc"
        
    logger.info('--- Aguarda até SUMIR o campo "cod_desc"')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='imagens/img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is not False:
        bot.click(procura_imagem(imagem='imagens/img_topcon/txt_ValoresTotais.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74))
        #logger.info(F'--- Tentativa de aguardar sumir o cod_desc: {tentativa_cod_desc}')
        if tentativa_cod_desc >= 100:
            logger.info('--- O campo cod_desc não sumiu, reiniciando o processo.')
            time.sleep(0.5)
            abre_topcon()
            return True
        else: # Aguarda até o topcompras voltar a funcionar
            ahk.win_activate('TopCompras', title_match_mode= 2)
            ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 70)
            tentativa_cod_desc += 1 
    else:
        logger.info(F'--- sumiu o campo "cod_desc", tentativa: {tentativa_cod_desc}')

    # Aguarda até o topcompras voltar a funcionar
    ahk.win_activate('TopCompras', title_match_mode= 2)
    ahk.win_wait_active('TopCompras', title_match_mode= 2, timeout= 70)
    bot.click(procura_imagem(imagem='imagens/img_topcon/txt_ValoresTotais.png', continuar_exec= True))

    # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
    logger.info(F'--- Preenchendo transportador: {cracha_mot}')
    ahk.win_activate('TopCompras', title_match_mode= 2)
    time.sleep(0.5)
    bot.click(procura_imagem(imagem='imagens/img_topcon/campo_000.png', continuar_exec= True))
    time.sleep(0.5)
    bot.press('tab')
    time.sleep(0.5)
    tentativa_achar_camp_re = 0
    while procura_imagem(imagem='imagens/img_topcon/campo_re_0.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
        logger.info(F'Tentativa: {tentativa_achar_camp_re}')
        time.sleep(0.25)
        tentativa_achar_camp_re += 1
        if tentativa_achar_camp_re >= 10:
            logger.info('--- Limite de tentativas de achar o campo "RE", reabrindo topcompras e reiniciando o processo.')
            time.sleep(0.5)
            abre_topcon()
            return True
    else:
        logger.info('--- Campo RE habilitado, preenchendo.')
        # Preenche o campo do transportador e verifica se aconteceu algum erro.
        bot.write(cracha_mot)  # ID transportador
        time.sleep(0.25)
        bot.press('enter')

    logger.info('--- Aguardando validar o campo do transportador')
    ahk.win_activate('TopCompras', title_match_mode=2)
    if procura_imagem(imagem='imagens/img_topcon/transportador_incorreto.png', continuar_exec= True) is not False:
        logger.info('--- Transportador incorreto!')
        bot.press('ENTER')
        bot.press('F2')
        marca_lancado(texto_marcacao='RE_Invalido')
        programa_principal()
    else:
        logger.info('--- Transportador validado! Prosseguindo para validação da placa')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.press('enter')

    # Verifica se o campo da placa ficou preenchido
    time.sleep(0.25)
    if procura_imagem('imagens/img_topcon/campo_placa.png', confianca= 0.74, continuar_exec=True) is not False:
        logger.info('--- Encontrou o campo vazio, inserindo XXX0000')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.click(procura_imagem('imagens/img_topcon/campo_placa.png', continuar_exec=True))
        bot.write('XXX0000')
        bot.press('ENTER')
        time.sleep(0.25)
    else:
        logger.info('--- Não achou o campo ou já está preenchido')

    # * -------------------------------------- Aba Produtos e serviços --------------------------------------
    ahk.win_activate('TopCompras', title_match_mode=2)
    ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 15)
    logger.info('--- Navegando para a aba Produtos e Servicos')
    tela_prod_servico = 0
    while procura_imagem(imagem='imagens/img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 1, continuar_exec= True, confianca= 0.74) is False:
        if tela_prod_servico > 15:
            logger.error('--- Não encontrou a tela produtos e serviços')
            raise TimeoutError
        
        bot.click(procura_imagem(imagem='imagens/img_topcon/produtos_servicos.png', confianca= 0.74, limite_tentativa= 3, continuar_exec= True))
        # Aguarda até aparecer o botão "alterar"
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        logger.info(F'--- Tentativa de procurar PRODUTO E SERVIÇOS: {tela_prod_servico}')
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
                logger.debug(F'--- Texto coletado da quantidade: {qtd_ton}, Valor escala: {valor_escala}')
                break

        logger.info('--- Abrindo a tela "Itens nota fiscal de compra" ')
        bot.click(procura_imagem(imagem='imagens/img_topcon/botao_alterar.png', area=(100, 839, 300, 400)))
        while procura_imagem(imagem='imagens/img_topcon/valor_cofins.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
            logger.info('--- Aguardando aparecer a tela "Itens nota fiscal de compra" ')
        
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait_active('TopCompras', title_match_mode=2, timeout= 30)
        time.sleep(0.25)
        logger.info('--- Preenchendo SILO e quantidade')
        if ('SILO' in silo1) or ('SILO' in silo2):
            bot.click(851, 443)  # Clica na linha para informar o primeiro silo
            if ('SILO' in silo1) and ('SILO' in silo2):  
                qtd_ton = str((qtd_ton / 2)) # Realiza a divisão da quantidade de cimento, pois será distribuido em dois silos!
                qtd_ton = qtd_ton.replace(".", ",")
                logger.info(F'--- Foi informado dois silos, preenchendo... {silo1} e {silo2}, quantidade: {qtd_ton}')
                bot.write(silo1)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
                bot.write(silo2)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
            else:
                logger.info(F'--- Foi informado UM silo, preenchendo... {silo1}, quantidade: {qtd_ton}')
                qtd_ton = str(qtd_ton)
                qtd_ton = qtd_ton.replace(".", ",")
                bot.write(silo1)
                bot.write(silo2)
                bot.press('ENTER')
                bot.write(str(qtd_ton))
                bot.press('ENTER')
        else: # Caso não tenha coletado nenhum silo.            
            if procura_imagem(imagem='imagens/img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True):
                marca_lancado(texto_marcacao= 'Faltou_InfoSilo')
                logger.info('--- Não foi informado nenhum SILO, porém a nota é de cimento!' )
                ahk.win_activate('TopCompras', title_match_mode=2)
                bot.click(procura_imagem(imagem='imagens/img_topcon/txt_cimento.png', limite_tentativa= 1, continuar_exec= True))
                bot.press('ESC')
                time.sleep(0.25)
                return True
            else: # Caso realmente seja de agregado.
                logger.info('--- Nota de agregado, continuando o processo!')
            
            bot.click(procura_imagem(imagem='imagens/img_topcon/confirma.png'))
            break


        #* Confirma ou cancela os processo executados na tela "itens nota fiscal de compras "
        bot.click(procura_imagem(imagem='imagens/img_topcon/confirma.png'))       
        if procura_imagem(imagem='imagens/img_topcon/txt_ErroAtribuida.png', limite_tentativa = 6, continuar_exec = True) is False:
            logger.info('--- Preenchimento completo, saindo do loop.' )
            break
        else:
            logger.warning(F'--- Falha, executando novamente a coleta das toneladas. Escala atual: {valor_escala}' )
            valor_escala += 10
            while procura_imagem(imagem='imagens/img_topcon/confirma.png', continuar_exec=True, limite_tentativa= 1, confianca= 0.74) is not False:
                bot.press('ENTER')
                bot.press('ESC')
                time.sleep(0.25)
            
        while procura_imagem(imagem='imagens/img_topcon/confirma.png', continuar_exec=True) is not False:
            ahk.win_activate('TopCompras', title_match_mode=2)
            logger.info('--- Aguardando fechamento da tela do botão "Alterar" ')
            time.sleep(0.25)

            if procura_imagem(imagem='imagens/img_topcon/local_estoque_obrigatorio.png', continuar_exec=True):
                marca_lancado("Erro local estoque")
                ahk.win_activate('TopCompras', title_match_mode=2)
                bot.click(procura_imagem(imagem='imagens/img_topcon/botao_ok.jpg', continuar_exec=True))
                bot.click(procura_imagem(imagem='imagens/img_topcon/bt_cancela.png', continuar_exec=True))
                return False
                
            tentativa += 1
            if tentativa > 10: # Caso a tela não feche.
                exit(bot.alert('Não foi possivel fechar a tela "itens nota fiscal de compra" '))
        else:
            logger.warning('--- Não fechou a tela "itens nota fiscal de compras" ')
    
    finaliza_lancamento() # Realiza todo o processo de finalização de lançamento.
    qtd_notas_lancadas += 1
    print(F"Quantidade de NFS lançadas: {qtd_notas_lancadas}")
    return True


if __name__ == '__main__':
    enviar_email("brunobola2010@gmail.com", "RPA Cortesia iniciando nova execução", "Realizando uma nova execução da função {PROGRAMA_PRINCIPAL}!")
    logger = get_logger("automacao") # Obter logger configurado
    tentativa = 0

    os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
    os.system('cls')

    #* Verifica qual sistema está rodando o script
    if 'VLPTIC1Z9HD33' not in platform.node(): 
        bot.FAILSAFE = False

    #* Garante a primeira abertura do TopCon e TopCompras
    abre_topcon()

    while tentativa < 5:
        try:
            logger.info('--- Iniciando o Try-Catch do PROGRAMA PRINCIPAL')
            programa_principal()
            print(F'- Tentativa: {tentativa}')
        except Exception:
            enviar_email("brunobola2010@gmail.com", "RPA Cortesia apresentou erro", F"Erro coletado: \n {traceback.format_exc()}")
            logger.exception('--- A execução principal apresentou erro! Executando o script PROGRAMA PRINCIPAL novamente.')
            logger.info("--- Reiniciando o TopCon e TopCompras para garantir a execução correta.")
            tentativa += 1
            abre_topcon()
        except(KeyboardInterrupt) as e:
            enviar_email("brunobola2010@gmail.com", "RPA Cortesia apresentou erro", F"Execução pausada pelo usuario!{e}")
            exit(logger.critical("Execução pausada pelo usuario"))
            
# TODO --- Caso NFE Faturada no final do mes, lançar com qual data?