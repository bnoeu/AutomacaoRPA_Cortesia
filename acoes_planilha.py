# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
# import cv2
# import pygetwindow as gw
import pytesseract
from ahk import AHK
from funcoes import marca_lancado, procura_imagem
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = False
tempo_inicio = time.time()
chave_xml, cracha_mot, silo2, silo1 = '', '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
bot.PAUSE = 1.3


def valida_lancamento():
    def coleta_planilha():
        bot.PAUSE = 0.4
        print('--- Abrindo planilha - COLETA_PLANILHA')
        ahk.win_activate('db_alltrips', title_match_mode= 2)
        time.sleep(0.3)

        #Verifica se já está no modo de edição, caso esteja, muda para o modo "exibição"
        if procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True) is False:
            time.sleep(0.3)
            print('--- Não está no modo exibição! Realizando alteração.')
                        
            #Espera até encontar o botão "Exibição" (Lapis bloqueado) e realiza um click nele
            print('--- Alterando para o modo exibição ')
            while procura_imagem(imagem='img_planilha/txt_exibicao.png', continuar_exec= True) is False:
                time.sleep(0.4)
                #Clica no botão da edição (lapis branco), para mostrar o dropdown com as opções
                bot.click(procura_imagem(imagem='img_planilha/txt_edicao.png', continuar_exec= True))
            else:
                print('--- Clicou no botão "Exibição" ')
                bot.click(procura_imagem(imagem='img_planilha/txt_exibicao.png', continuar_exec= True))
                
                #Aguarda sumir o botão
                while procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True) is False:
                    print('--- Aguardando entrar no modo de exibição ')
                    time.sleep(0.4)

            #Aguarda até aparecer o botão do modo "exibição"
            while procura_imagem(imagem='img_planilha/botao_exibicaoverde.png', continuar_exec=True) is False: #Aguarda enquanto não achar o botão
                time.sleep(0.3)
            else:
                print('--- Alterado para o modo exibição, continuando.')
        else: #Caso não esteja no modo "Edição"
            print('--- A planilha já está no modo "Exibição", continuando processo')

        #Altera o filtro para "vazio", para iniciar a coleta de dados.
        if procura_imagem(imagem='img_planilha/bt_filtro.png', continuar_exec=True, area= (1468, 400, 200, 200)) is not False:
            print('--- Já está filtrado, continuando!')
        else:
            print('--- Não está filtrado, executando o filtro!')
            bot.click(procura_imagem(imagem='img_planilha/bt_setabaixo.png', area=(1529, 459, 75, 75)))

            #Caso não apareça o botão "Selecionar tudo" clica em "limpar filtro" e executa tudo novamente.
            if procura_imagem(imagem='img_planilha/botao_selecionartudo.png', continuar_exec= True) is False:
                bot.click(procura_imagem(imagem='img_planilha/bt_limparFiltro.png'))
                coleta_planilha()
            else: #Se tudo estiver ok, prossegue aplicando o filtro nas notas vazias. 
                bot.click(procura_imagem(imagem='img_planilha/botao_selecionartudo.png'))
                bot.click(procura_imagem(imagem='img_planilha/bt_vazias.png'))
                bot.click(procura_imagem(imagem='img_planilha/bt_aplicar.png'))
                print('--- Filtrado pelas notas vazias!')

                #Aguarda aparecer o botão do filtro, para confirmar que está filtrado! 
                while procura_imagem(imagem='img_planilha/bt_filtro.png', limite_tentativa= 10, area= (1468, 400, 200, 200)) is False:
                    print('--- Aguardando o botão do filtro na coluna "Status" ')
                    time.sleep(0.4)
                else:
                    print('--- Filtro das notas vazias aplicado!')
        
        # * Coleta os dados da linha atual
        dados_planilha = []
        print('--- Copiando dados e formatando')
        time.sleep(1)
        
        #Clica na primeira linha (Campo RE), e pressiona seta para baixo
        bot.click(procura_imagem(imagem='img_planilha/titulo_re.png'))
        time.sleep(1)
        bot.press('DOWN')
        
        for n in range(0, 7, 1):  # Copia dados dos 6 campos
            while True:
                bot.hotkey('ctrl', 'c')
                if 'Recuperando' in ahk.get_clipboard():
                    time.sleep(0.2)
                else:
                    break
            dados_planilha.append(ahk.get_clipboard())
            bot.press('right')
        print(F'--- Dados copiados com sucesso: {dados_planilha}')
        tempo_coleta = time.time() - tempo_inicio
        tempo_coleta = tempo_coleta / 60
        print(F'\n Tempo que levou: {tempo_coleta:0f}')
        return dados_planilha
    
    #Realiza o processo de validação do lançamento.
    while True:
        # Trata os dados coletados em "dados_planilha"
        dados_planilha = coleta_planilha()
        chave_xml = dados_planilha[4].strip()
        if len(chave_xml) < 10:
            exit(bot.alert('chave_xml invalida'))
        
        # -------------------------------------- Lançamento Topcon --------------------------------------
        print('--- Abrindo TopCompras para iniciar o lançamento')
        ahk.win_activate('TopCompras', title_match_mode=2)
        ahk.win_wait('TopCompras', title_match_mode=2, timeout= 10)
        if ahk.win_is_active('TopCompras', title_match_mode=2):
            print('Tela compras está maximizada! Iniciando o programa')
        else:
            exit(bot.alert('Tela de Compras não abriu.'))
        # Processo de lançamento
        bot.press('F2')
        bot.press('F3', presses=2, interval=0.5)
        time.sleep(1.5)       
        bot.doubleClick(558, 235, interval = 0.5)  # Clica dentro do campo para inserir a chave XML
        bot.write(chave_xml)
        bot.press('ENTER')
        ahk.win_wait_active('TopCompras')
        tentativa = 0
        while tentativa < 10:
            time.sleep(0.3)                
            #Verifica quais das telas apareceu. 
            if procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.click(procura_imagem(imagem='img_topcon/botao_sim.jpg', limite_tentativa=1, continuar_exec=True))
                print('--- XML Validado, indo para verificação do pedido\n')
                return dados_planilha
            elif procura_imagem(imagem='img_topcon/chave_invalida.png', limite_tentativa= 1, continuar_exec=True) is not False:
                print('--- Nota já lançada, marcando planilha!')
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Lancado_Manual')
                break
            elif procura_imagem(imagem='img_topcon/naoencontrado_xml.png', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Aguardando_SEFAZ')
                break
            elif procura_imagem(imagem='img_topcon/chave_44digitos.png', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='Chave_invalida')
                break
            elif procura_imagem(imagem='img_topcon/nfe_cancelada.png', limite_tentativa= 1, continuar_exec=True) is not False:
                bot.press('ENTER')
                marca_lancado(texto_marcacao='NFE_CANCELADA')
                break
            
            # Verifica caso tenha travado e espera até que o topcom volte a responder
            print('--- Aguardando TopCompras Retornar')
            while ahk.win_exists('Não está respondendo'):
                time.sleep(0.3)
            tentativa += 1
        if tentativa >= 15:
            exit('Rodou 10 verificações e não achou nenhuma tela, aumentar o tempo')
