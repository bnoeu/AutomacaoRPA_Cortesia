import logging
import pyautogui as bot


def verifica_finalizou_planilha(dados_planilha = [], chave_xml= ""):
    if len(dados_planilha[4]) < 7: 
        exit(bot.alert('chave XML invalida.'))
    if len(dados_planilha[6]) > 1: # Caso realmente esteja preenchido
        logging.warning(F'--- Realmente está na ultima chave: {chave_xml}, executando COPIA BANCO')
        return True
        
        
if __name__ == '__main__':
    dados_copiados = verifica_finalizou_planilha()