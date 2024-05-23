# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import time
from datetime import date
import pytesseract
#import cv2
from ahk import AHK
import pyautogui as bot
from funcoes import procura_imagem, extrai_txt_img, marca_lancado
from acoes_planilha import valida_lancamento
from valida_pedido import valida_pedido
'''
#*Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from pynfe.processamento.comunicacao import ComunicacaoSefaz
'''
#import winsound
#import pygetwindow as gw

# --- Definição de parametros
ahk = AHK()
posicao_img = 0  # Define a variavel para utilização global dela.
continuar = True
bot.FAILSAFE = True
acabou_pedido = ''
numero_nf = "965999"
transportador = "111594"
chave_xml, silo2, silo1 = '', '', ''
pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract-OCR\tesseract.exe"
time.sleep(0.8)
start = time.time()


#! Variavel de teste
silo1 = 'SILO 1'
filial_estoq = 'JAGUARE'
centro_custo = filial_estoq
cracha_mot = '112480'


''' #* Cria banco de dados
#Cria a conexão com o banco de dados
con = sqlite3.connect("informacoes.db")

#Cursor para realizar comandos dentro do banco de dados
cur = con.cursor()

#Utilizando o cursor, executa a ação da criação da tabela informacoes, com as seguintes colunas: XML, CRACHA, TEMPO
#cur.execute("CREATE TABLE informacoes(xml, cracha, tempo)")
#! Continuar tutorial de banco de dados https://docs.python.org/3/library/sqlite3.html
exit()
'''

''' #* Consulta as telas abertas
for telas in ahk.list_windows():
    print(telas.text)
'''


ahk.win_activate('TopCompras', title_match_mode= 2)
#ahk.win_activate('db_alltrips', title_match_mode= 2)
time.sleep(0.5)
#! Utilizado apenas para estar trechos de codigo.
#bot.click(procura_imagem(imagem='img_topcon/icone_topcon.png', continuar_exec=True))

'''
while True:
    bot.press('alt') #Ativa os atalhos
    #Clica no botão para alterar Edição / Exibição
    bot.press('z')
    bot.press('m')
    #Habilita a opção: Exibição
    bot.press('e')
'''

'''
certificado = "/certificado_nfe/certificado.pfx"
senha = '123456'
uf = 'sp'
homologacao = True

con = ComunicacaoSefaz(uf, certificado, senha, homologacao)
xml = con.status_servico('nfe')
print(xml.text)
'''
tentativa = 0
valor_escala = 200
while True:
    try:
        qtd_ton = extrai_txt_img(imagem='img_toneladas.png', area_tela=(892, 577, 70, 20), porce_escala= valor_escala).strip()
        qtd_ton = qtd_ton.replace(",", ".")
        qtd_ton = float(qtd_ton)
    except ValueError:
        valor_escala += 15
    else:
        print(F'--- Texto coletado da quantidade: {qtd_ton}')

    #* ----------------------- Parte "Itens nota fiscal de compra" -----------------------         
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
        break
        
    bot.click(procura_imagem(imagem='img_topcon/confirma.png'))            
    if procura_imagem(imagem='img_topcon/txt_ErroAtribuida.png', continuar_exec=True) is not False:
        bot.press('enter')
        bot.press('esc')
        time.sleep(2)
    else:
        break
        
    
while procura_imagem(imagem='img_topcon/confirma.png', continuar_exec=True) is not False:
    tentativa += 1
    print('--- Aguardando fechamento da tela do botão "Alterar" ')
    time.sleep(0.3)
    #TODO --- VerificaR se apareceu a tela "quantidade atribuida aos locais"

    if tentativa > 10: #Executa o loop 10 vezes até dar erro.
        exit(bot.alert('Apresentou algum erro.'))
# TODO --- CASO O REMOTE APP DESCONECTE, RODAR O ABRE TOPCON