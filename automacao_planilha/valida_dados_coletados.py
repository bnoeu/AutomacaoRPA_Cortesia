# -*- Criado por Bruno da Silva Santos. -*-
# Para utilização na Cortesia Concreto.

import pyautogui as bot
from utils.funcoes import marca_lancado

# --- Definição de parametros locais
from utils.funcoes import ahk as ahk

dados_planilha = ['Status', 'teste']

def valida_dados_coletados(dados_planilha = []):
    chave_xml = dados_planilha[4]
    tipo_chave = (chave_xml[20:22])

    if "Status" in dados_planilha:
        return False
    if len(dados_planilha) < 6: # Verifica se coletou todos os campos
        return False 
    elif len(dados_planilha[0]) == 1: # Se o campo RE não estiver preenchido
        return False
    elif len(dados_planilha[3]) < 1: # Campo filial de estoque
        return False
    elif len(dados_planilha[4]) < 42 and len(dados_planilha[4]) > 1: # Caso a chave tenha menos de 42 digitos, ela é invalida!
        marca_lancado('chave_invalida')
        return True 
    elif (len(dados_planilha[0]) < 4) or (len(dados_planilha[0]) == 5): # Caso o crachá inserido seja menor de 4 digitos ou de 5 digitos
        marca_lancado('RE_Invalido')
        return True
    elif "55" not in tipo_chave: # Verifica se o tipo da chave é de um CT-E (Verificando o digitos 21º e 22º)
        marca_lancado('chave_invalida')
        return True
    else:
        return True # Dados validados
        
if __name__ == '__main__':
    bot.FAILSAFE = True
    valida_dados_coletados()
