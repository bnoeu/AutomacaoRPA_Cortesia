# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#! Link da planilha
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EQx5PclDGRFGkweQjtb3QckByyAsqydfI5za0MTuO9tjXg
# Debug db alltrips
# https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=n9xgTPCH3Eg
# db_alltrips no paulo, apenas leitura
#  

#*'''
#* A resolução da maquina precisa ser: 1920 x 1080, com a aproximação em "150%"
#*'''

import os
import time
import platform
import traceback
from numpy import true_divide
import pytesseract
import pyautogui as bot
from datetime import date, datetime, timedelta
from utils.funcoes import ahk as ahk
from abre_topcon import main as abre_topcon
from utils.enviar_email import enviar_email
from utils.configura_logger import get_logger
from valida_pedido import main as valida_pedido
from valida_lancamento import valida_lancamento
from preenche_local import main as preenche_local
from finaliza_lancamento import finaliza_lancamento
from utils.funcoes import marca_lancado, procura_imagem, verifica_horario, ativar_janela

#* Definição de parametros
posicao_img = 0
continuar = True
qtd_notas_lancadas = 0
bot.LOG_SCREENSHOTS = True  
bot.LOG_SCREENSHOTS_LIMIT = 5
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
logger = get_logger("automacao", print_terminal= True) # Obter logger configurado

def calcula_tempo_processo(tempo_inicial):
    # Linha específica onde você quer medir o tempo
    end_time = time.time()
    elapsed_time = end_time - tempo_inicial
    medicao_minutos = elapsed_time / 60
    print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")
    exit(bot.alert("acabou"))

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
        marca_lancado(texto_marcacao= "Filial_nao_cadastrada")
        exit(F'Filial de estoque não padronizada: {filial_estoq}')
    
    if centro_custo != "":
        return centro_custo

def coleta_valida_dados():
    logger.info('--- Executando COLETA VALIDA DADOS ')
    #* Realiza a validação do pedido
    acabou_pedido = False

    while acabou_pedido is False: 
        dados_planilha = valida_lancamento() # Coleta e confere os dados do lançamento atual
        acabou_pedido = valida_pedido() # Verifica se o pedido está valido.
        time.sleep(0.2)
    else:
        print(dados_planilha)
        return dados_planilha


def formata_data_coletada(dados_copiados):
    data_copiada = dados_copiados.split(' ')[0]  # Pega apenas a parte da data
    print(F"Data copiada: {data_copiada}")
    # Converter a string para objeto datetime.date
    data_obj = datetime.strptime(data_copiada, "%d/%m/%y").date()

    # Obtém a data de amanhã como objeto date
    amanha_data = coleta_proximo_dia()

    # Comparação correta entre objetos date
    if data_obj >= amanha_data:
        print("--- A data coletada é do próximo dia! Alterando para a data atual.")
        return date.today().strftime("%d/%m/%y")  # Retorna a data atual formatada
    
    print("--- A data coletada é válida!")
    return data_obj.strftime("%d/%m/%y")  # Retorna a data coletada formatada

def coleta_proximo_dia():
    # Retorna a data de amanhã como objeto date
    return date.today() + timedelta(days=1)

def programa_principal():
    global qtd_notas_lancadas
    bot.PAUSE = 0.4

    #* Confere o horario dessa execução.
    verifica_horario()
    logger.info("\n\n\n")
    logger.info('---------------------------------------------------------------------------------------------------')
    logger.info('--- INICIANDO UM NOVO LANÇAMENTO DE NFE --- ')
    logger.info('---------------------------------------------------------------------------------------------------')

    #* Passa todos os dados para as suas variaveis.
    dados_planilha = coleta_valida_dados()
    tempo_inicial = time.time()
    silo1 = dados_planilha[1]
    silo2 = dados_planilha[2]
    chave_xml = dados_planilha[4]
    cracha_mot = dados_planilha[0]
    filial_estoq = dados_planilha[3].split('-') # Recebe por exemplo: ['1001', 'VILA PRUDENTE']
    filial_estoq = filial_estoq[0] # O dado é passado assim: ['1001', 'VILA PRUDENTE'], aqui formata para '1001'
    centro_custo = valida_filial_estoque(filial_estoq) # Realiza a validação da filial de estoque.
    data_formatada = formata_data_coletada(dados_planilha[8])


#* -------------------------- Continua o processo de lançamento da NFE -------------------------- 
    logger.info('--- Preenchendo dados na tela principal do lançamento')
    ativar_janela('TopCompras')

    #* Aguarda até aparecer o botão "Produtos e serviços", isso valida que fechou a tela de vinculação de pedido
    while procura_imagem(imagem='imagens/img_topcon/produtos_servicos.png', continuar_exec= True) is False:
        ativar_janela('TopCompras')

    logger.info('--- Preenchendo filial de estoque')
    bot.press('up')
    bot.write(filial_estoq)
    time.sleep(2)
    bot.press('TAB', presses= 1) # Confirma a informação da nova filial de estoque
    time.sleep(1.5)
    bot.press('TAB', presses= 1) # Confirma a informação da nova filial de estoque


    #* Alteração da data
    logger.info('--- Realizando validação/alteração da data')
    hoje = date.today()
    #hoje = hoje.strftime("%d%m%y")  # dd/mm/YY
    logger.info(F'--- Inserindo a data coletada: {data_formatada} e apertando ENTER')
    bot.write(data_formatada)
    bot.press('ENTER')
    time.sleep(1.5)
    ativar_janela('TopCompras', 70)

    # Caso o sistema informe que a data deve ser maior/igual a data inserida acima.
    logger.info('--- Verificando se apareceu data')
    if procura_imagem('imagens/img_topcon/data_invalida.png', continuar_exec= True, limite_tentativa= 2):
        logger.warning('--- Precisa mudar a data, inserindo a data de hoje!')
        #bot.alert("Apresentou tela erro")
        ahk.win_close("TopCompras (VM-CortesiaApli.CORTESIA.com)", title_match_mode= 2)
        time.sleep(0.5)        
        bot.write(hoje)
        bot.press('enter')
        # Aguarda até o topcompras voltar a funcionar
        ativar_janela('TopCompras', 70)
    else:
        logger.info('--- Não foi necessario alterar a data!')

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
            time.sleep(0.4)

    logger.info(F'--- Trocando o centro de custo para {centro_custo}')
    ativar_janela('TopCompras', 70)
    bot.write(centro_custo)
    ahk.win_activate('TopCompras', title_match_mode= 2)
    logger.info('--- Aguarda aparecer o campo cod_desc')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='imagens/img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is False:
        time.sleep(0.4)
        if tentativa_cod_desc >= 100:
            logger.info('--- Não foi possivel encontrar o campo cod_desc, reiniciando o processo.')
            time.sleep(0.5)
            abre_topcon()
            return True
        else: # Aguarda até o topcompras voltar a funcionar
            ativar_janela('TopCompras', 70)
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
            ativar_janela('TopCompras', 70)
            tentativa_cod_desc += 1 
    else:
        logger.info(F'--- sumiu o campo "cod_desc", tentativa: {tentativa_cod_desc}')

    # Aguarda até o topcompras voltar a funcionar
    ativar_janela('TopCompras', 70)
    bot.click(procura_imagem(imagem='imagens/img_topcon/txt_ValoresTotais.png', continuar_exec= True))

    # * -------------------------------------- VALIDAÇÃO TRANSPORTADOR --------------------------------------
    logger.info(F'--- Preenchendo transportador: {cracha_mot}')
    ahk.win_activate('TopCompras', title_match_mode= 2)
    time.sleep(0.25)
    bot.click(procura_imagem(imagem='imagens/img_topcon/campo_000.png'))
    time.sleep(0.5)
    bot.press('tab')
    time.sleep(0.5)
    tentativa_achar_camp_re = 0
    while procura_imagem(imagem='imagens/img_topcon/campo_re_0.png', continuar_exec= True, limite_tentativa= 1, confianca= 0.74) is False:
        logger.info(F'Tentativa: {tentativa_achar_camp_re}')
        time.sleep(0.4)
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
        time.sleep(0.5)
        bot.press('enter')

    logger.info('--- Aguardando validar o campo do transportador')
    ahk.win_activate('TopCompras', title_match_mode=2)
    if procura_imagem(imagem='imagens/img_topcon/transportador_incorreto.png', continuar_exec= True, limite_tentativa= 6) is not False:
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
    time.sleep(0.4)
    if procura_imagem('imagens/img_topcon/campo_placa.png', confianca= 0.74, continuar_exec=True) is not False:
        logger.info('--- Encontrou o campo vazio, inserindo XXX0000')
        ahk.win_activate('TopCompras', title_match_mode=2)
        bot.click(procura_imagem('imagens/img_topcon/campo_placa.png', continuar_exec=True))
        bot.write('XXX0000')
        bot.press('ENTER')
        time.sleep(0.4)
    else:
        logger.info('--- Não achou o campo ou já está preenchido')

    # * -------------------------------------- Aba Produtos e serviços --------------------------------------
    ativar_janela('TopCompras')
    logger.info('--- Navegando para a aba Produtos e Servicos')
    tela_prod_servico = 0
    while procura_imagem(imagem='imagens/img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 1, continuar_exec= True, confianca= 0.74) is False:
        
        if tela_prod_servico > 15:
            logger.error('--- Não encontrou a tela produtos e serviços')
            raise TimeoutError
        
        # Aguarda até aparecer o botão "alterar"
        bot.click(procura_imagem(imagem='imagens/img_topcon/produtos_servicos.png', confianca= 0.74, limite_tentativa= 3, continuar_exec= True))
        ativar_janela('TopCompras', 30)
        logger.info(F'--- Tentativa de procurar PRODUTO E SERVIÇOS: {tela_prod_servico}')
        tela_prod_servico += 1
    
    
    if '38953477000164' in chave_xml: #Caso não tenha o CNPJ da Consmar
        finaliza_lancamento()
        return True
    
    #* Realiza a extração da quantidade de toneladas
    preenche_local(silo1, silo2)

    #* Finaliza o processo de lançamento
    finaliza_lancamento() # Realiza todo o processo de finalização de lançamento.
    qtd_notas_lancadas += 1
    print(F"Quantidade de NFS lançadas: {qtd_notas_lancadas}")
    logger.info(F"Quantidade de NFS lançadas: {qtd_notas_lancadas}")

    # Valida a medição de tempo
    end_time = time.time()
    elapsed_time = end_time - tempo_inicial
    medicao_minutos = elapsed_time / 60
    print(f"Tempo decorrido: {medicao_minutos:.2f} segundos")
    logger.info(f"Tempo decorrido: {medicao_minutos:.2f} segundos")
    time.sleep(0.05)
    #exit(bot.alert(F"Lançamento concluido! \n Tempo que levou: {medicao_minutos:.2f}"))

    return True


def trata_erro(ultimo_erro, tentativa):
    last_trace = traceback.extract_tb(ultimo_erro.__traceback__)[-1]  # Última entrada do traceback
    arquivo_erro = os.path.basename(last_trace.filename) # Nome do arquivo

    # Captura o traceback completo
    erro_traceback = traceback.format_exc()
    erro_tipo = type(ultimo_erro).__name__
    erro_mensagem = str(ultimo_erro)

    # Mensagem detalhada do erro
    mensagem_erro = (
        f"Erro ocorrido durante a execução do RPA:\n\n"
        f"Tipo do erro: {erro_tipo}\n"
        f"Mensagem: {erro_mensagem}\n\n"
        f"Traceback:\n{erro_traceback}"
    )

    return arquivo_erro, mensagem_erro


def main(lancamento_realizado = False):
    verifica_horario() # Confere o horario dessa execução.
    
    if lancamento_realizado == False:
        if not abre_topcon():
            raise Exception("Falhou ao abrir o topcon")

    
    while programa_principal():
        logger.info(F"Lançamento realizado! Valor da variavel: {lancamento_realizado}")
        return True
    else:
        logger.info(F"Lançamento não foi realizado! Variavel: {lancamento_realizado}")
        lancamento_realizado = False
    

if __name__ == '__main__':
    tentativa = 0
    lancamento_realizado = False
    tempo_inicial = time.time()
    tempo_pausa = 600 # 10 minutos
    
    #* Realiza os processos inicias da execução da automação
    os.system('taskkill /im AutoHotkey.exe /f /t') # Encerra todos os processos do AHK
    os.system('cls')

    #* Verifica qual sistema está rodando o script
    if 'VLPTIC1Z9HD33' not in platform.node(): 
        bot.FAILSAFE = False

    while tentativa < 10:
        logger.info(F'--- Iniciando nova tentativa Nº {tentativa} o Try-Catch do PROGRAMA PRINCIPAL')
        try:
            if main(lancamento_realizado) is True:
                lancamento_realizado = True
                if tentativa > 1:
                    tentativa - 1
            else:
                lancamento_realizado = False
        except Exception as ultimo_erro:
            lancamento_realizado = False
            arquivo_erro, mensagem_erro = trata_erro(ultimo_erro, tentativa)

            #enviar_email("brunobola2010@gmail.com", F"[RPA Cortesia] Apresentou erro na task: {arquivo_erro}, tentativa: {tentativa}", F"Erro coletado: \n {mensagem_erro}")
            logger.exception(F'--- A execução principal apresentou erro! Tentativa: {tentativa}, Pausa anterior: {tempo_pausa}')

            #* Realiza as verificações antes da proxima tentativa
            verifica_horario()

            if tentativa >= 4: # Começa a pausar o script após a 5º execução
                logger.info(F"Pausando por: {tempo_pausa} segundos antes da proxima tentativa")
                time.sleep(tempo_pausa)
                tempo_pausa += (0.5 * tempo_pausa)

            if tentativa > 9:
                enviar_email(
                    "brunobola2010@gmail.com",
                    f"[RPA Cortesia] Erro catastrofico: {arquivo_erro}",
                    f"Erro coletado: \n{traceback.format_exc()}"
                )
                logger.critical(F'--- A execução principal apresentou erro! Executando o script principal novamente, tentativa: {tentativa}')
                break

            abre_topcon()
            tentativa += 1

        except(KeyboardInterrupt) as e:
            exit(logger.critical("Execução pausada pelo usuario"))
        else:
            tentativa = 0
    else:
        enviar_email(
            "brunobola2010@gmail.com",
            "[RPA Cortesia] Executou todas as tentativas",
            f"A execução principal executou todas as tentativas e quebrou\n {arquivo_erro} \n {mensagem_erro}"
        )
        logger.critical('--- A execução principal executou todas as tentativas')

        # Log do erro crítico no sistema
        logger.critical("A execução principal falhou com erro crítico.")
        logger.critical(mensagem_erro)
