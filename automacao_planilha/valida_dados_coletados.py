# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

#import time
import logging
from ahk import AHK
import pyautogui as bot


# --- Definição de parametros
ahk = AHK()


def valida_dados_coletados(dados_planilha = []):
        if len(dados_planilha) < 6: # Verifica se coletou todos os campos
            return False 
        elif len(dados_planilha[0]) == 1: # Se o campo RE não estiver preenchido
            return False
        elif len(dados_planilha[3]) < 1: # Campo filial de estoque
            return False
        elif len(dados_planilha[4]) < 44: # Campo chave XML
            return False
        else:
            return True # Dados validados
        
if __name__ == '__main__':
    bot.FAILSAFE = True
    valida_dados_coletados()