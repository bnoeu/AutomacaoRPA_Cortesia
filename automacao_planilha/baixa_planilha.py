import openpyxl
import pandas as pd
import time
import os
from onedrivedownloader import download as one_download

#* Parametros
planilha_debug = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bruno_silva_cortesiaconcreto_com_br/ETubFnXLMWREkm0e7ez30CMBnID3pHwfLgGWMHbLqk2l5A?rtime=jFhSykjw3Eg"

#* Funções

#! Essa função só deve ser executada não houver linhas com a coluna "Status" vazias no "producao"
def baixa_planilha():
    try:
        os.remove("db_alltrips/debug_db_alltrips.xlsx")
    except FileNotFoundError: 
        print('Arquivo db_alltrips.xlsx excluido.')
    except PermissionError:
        exit(print('--- Planilha aberta, fechar a planilha'))
    finally:
        print("Baixando a planilha")
        one_download(planilha_debug, filename="db_alltrips/debug_db_alltrips.xlsx")


def main():
    baixa_planilha()

if __name__ == '__main__':
    main()