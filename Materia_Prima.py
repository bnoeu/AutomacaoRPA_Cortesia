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


import time
import traceback
import pytesseract
import pyautogui as bot
from utils.funcoes import ahk as ahk
import asyncio
from abre_topcon import main as abre_topcon
from utils.enviar_email import enviar_email
from utils.configura_logger import get_logger
from preenche_local import main as preenche_local
from finaliza_lancamento import finaliza_lancamento
from utils.funcoes import (
    procura_imagem, verifica_horario, ativar_janela, 
    print_erro, matar_autohotkey, calcula_tempo_processo, trata_erro
)
from utils.validators import (
    valida_filial_estoque, formata_data_coletada,
    coleta_valida_dados
)
from automacao.topcompras_handler import (
    valida_transportador, preenche_data, preenche_filial_estoque
)

#* Definição de parametros
posicao_img = 0
continuar = True
bot.FAILSAFE = False
qtd_notas_lancadas = 0
bot.LOG_SCREENSHOTS = True  
bot.LOG_SCREENSHOTS_LIMIT = 5
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe" # pyrefly: ignore[bad-assignment]
logger = get_logger("automacao", print_terminal= True) # Obter logger configurado



def programa_principal():
    global qtd_notas_lancadas
    bot.PAUSE = 0.6

    #* Confere o horario dessa execução.
    verifica_horario()
    logger.info("\n\n\n")
    logger.info('---------------------------------------------------------------------------------------------------')
    logger.info('--- INICIANDO UM NOVO LANÇAMENTO DE NFE --- ')
    logger.info('---------------------------------------------------------------------------------------------------')

    #* Passa todos os dados para as suas variaveis.
    dados_planilha = coleta_valida_dados()
    #dados_planilha = ['8078', '', '', '1036-PERUS', '35250648302640001588550100005409241251434230', 'p4ozMaAb2_Q', '', '', '18/06/2025 18:25', '', '', '', '45826,89236', '', 'Versão AllTrips: 172']
    
    # Preenche dados do campo Tipo Doc
    ahk.win_activate('TopCompras', title_match_mode= 2)
    time.sleep(1)
    if procura_imagem(imagem='imagens/img_topcon/txt_15-NF_ELETRO.png', continuar_exec= True, confianca= 0.74) is False:
        bot.write('15', interval= 1)
        bot.press('enter', interval= 1)
        time.sleep(2)
    else:
        pass
        # Já está preenchido

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

    preenche_filial_estoque(filial_estoq= filial_estoq)

    preenche_data(data_formatada)

    logger.info(F'--- Trocando o centro de custo para {centro_custo}')
    ativar_janela('TopCompras', 70)
    bot.write(centro_custo)
    ahk.win_activate('TopCompras', title_match_mode= 2)
    logger.info('--- Aguarda aparecer o campo cod_desc')
    tentativa_cod_desc = 0
    while procura_imagem(imagem='imagens/img_topcon/cod_desc.png', continuar_exec=True, confianca= 0.74, limite_tentativa= 1) is False:
        time.sleep(0.2)
        if tentativa_cod_desc >= 100:
            logger.info('--- Não foi possivel encontrar o campo cod_desc, reiniciando o processo.')
            time.sleep(0.2)
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
        if tentativa_cod_desc >= 100:
            logger.info('--- O campo cod_desc não sumiu, reiniciando o processo.')
            time.sleep(0.25)
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

    if valida_transportador(cracha_mot) is False:
        logger.info('--- Falhou na validação do transportador, recomeçando o processo.')
        return False

    # * -------------------------------------- Aba Produtos e serviços --------------------------------------
    ativar_janela('TopCompras')
    logger.info('--- Navegando para a aba Produtos e Servicos')
    tela_prod_servico = 0
    while procura_imagem(imagem='imagens/img_topcon/botao_alterar.png', area=(100, 839, 300, 400), limite_tentativa= 1, continuar_exec= True, confianca= 0.74) is False:
        
        if tela_prod_servico > 25:
            logger.error('--- Não encontrou a tela produtos e serviços')
            raise TimeoutError
        
        # Aguarda até aparecer o botão "alterar"
        bot.click(procura_imagem(imagem='imagens/img_topcon/produtos_servicos.png', confianca= 0.74, limite_tentativa= 3, continuar_exec= True))
        ativar_janela('TopCompras', 30)
        logger.info(F'--- Tentativa de procurar PRODUTO E SERVIÇOS: {tela_prod_servico}')
        tela_prod_servico += 1
    
    '''
    if '38953477000164' in chave_xml: #Caso tenha o CNPJ da Consmar
        finaliza_lancamento()
        return True
    '''

    #* Realiza a extração da quantidade de toneladas
    preenche_local(silo1, silo2)

    #* Finaliza o processo de lançamento
    finaliza_lancamento(temp_inicial= tempo_inicial) # Realiza todo o processo de finalização de lançamento.
    qtd_notas_lancadas += 1
    print(F"Quantidade de NFS lançadas: {qtd_notas_lancadas}")
    logger.info(F"Quantidade de NFS lançadas: {qtd_notas_lancadas}")

    calcula_tempo_processo(tempo_inicial= tempo_inicial)

    return True




def main(lancamento_realizado = False):
    verifica_horario() # Confere o horario dessa execução.
    
    if lancamento_realizado is False:
        abre_topcon()
        pass

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
    arquivo_erro = ""
    mensagem_erro = ""


    #* Realiza os processos inicias da execução da automação
    print("--- Iniciando RPA! Realizando o fechamento do AHK!")
    asyncio.run(matar_autohotkey(nome_exec= "AutoHotkey.exe"))

    while tentativa < 10:
        logger.info(F'--- Iniciando nova tentativa Nº {tentativa} o Try-Catch do PROGRAMA PRINCIPAL')
        try:
            if main(lancamento_realizado) is True:
                lancamento_realizado = True
                if tentativa > 1:
                    tentativa - 1
            else:
                lancamento_realizado = False
        except (ValueError, RuntimeError, TimeoutError) as ultimo_erro:
            lancamento_realizado = False
            arquivo_erro, mensagem_erro = trata_erro(ultimo_erro, tentativa)
            caminho_imagem = print_erro()
            #enviar_email("brunobola2010@gmail.com", F"[RPA Cortesia] Apresentou erro na task: {arquivo_erro}, tentativa: {tentativa}", F"Erro coletado: \n {mensagem_erro}")
            logger.exception(F'--- A execução principal apresentou erro! Tentativa: {tentativa}, Pausa anterior: {tempo_pausa}, Erro: {mensagem_erro}')
            print(F'--- A execução principal apresentou erro! Tentativa: {tentativa}, Pausa anterior: {tempo_pausa}')

            #* Realiza as verificações antes da proxima tentativa
            verifica_horario()

            if tentativa >= 4: # Começa a pausar o script após a 5º execução
                logger.info(F"Pausando por: {tempo_pausa} segundos antes da proxima tentativa")
                time.sleep(tempo_pausa)
                tempo_pausa = min(int(tempo_pausa + (0.5 * tempo_pausa)), 3600)  # Limita a pausa máxima a 1 hora (3600 segundos)

            if tentativa > 9:
                enviar_email(
                    "brunobola2010@gmail.com",
                    f"[RPA Cortesia] Erro catastrofico: {arquivo_erro}",
                    f"Erro coletado: \n{traceback.format_exc()}"
                )
                logger.critical(F'--- A execução principal apresentou erro! Executando o script principal novamente, tentativa: {tentativa}')
                break

            tentativa += 1

        except(KeyboardInterrupt) as e:
            exit(logger.critical(f"Execução pausada pelo usuario: {e}"))
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
