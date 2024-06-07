import os
import pandas as pd
from onedrivedownloader import download as one_download

#! Funções
def baixa_planilha():
    try:
        os.remove("db_alltrips/db_alltrips.xlsx")
    except FileNotFoundError:
        print('Arquivo db_alltrips.xlsx já excluido.')
    except PermissionError:
        exit(print('--- Planilha aberta, fechar a planilha'))

    ln = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EW_8FZwWFYVAol4MpV1GglkBJEaJaDx6cfuClnesIu60Ng?e=pveECF"
    one_download(ln, filename="db_alltrips")

# Reorganiza o arquivo.
def reorganiza_arquivo(nome = ""):
    dataframe = pd.read_excel('db_alltrips/' + nome, date_format= "dd/mm/yyyy") #Insere o arquivo no dataframe
    dataframe['D. Insercao'] = pd.to_datetime(dataframe["D. Insercao"], errors= 'coerce', format= '%d/%m/%Y') # Converte a coluna para o tipo "datetime"
    #! Vou precisar deletar as linhas sem a data de insercao?!
    #dataframe = dataframe.dropna(subset=["D. Insercao"])
    dataframe = dataframe.sort_values(by = "D. Insercao") # Reordena o arquivo, com base na coluna "D. Insercao"
    dataframe.to_excel('db_alltrips/ordenado_' + nome, index= False)

def verifica_status_vazio(arquivo):
    df = pd.read_excel(arquivo) # Carrega o arquivo excel
    
    # Verificar se há valores vazios na coluna "Status"
    # Define o indice/posição da coluna status dentro da planilha
    status_col_index = 6
    coluna_status = df.iloc[0].tolist()
    print(coluna_status)
    df.loc[0, 'Status'] = "Bruno"
    coluna_status = df.iloc[0].tolist()
    print(coluna_status)    
    exit()
    
    # Procura pela primeira ocorrencia de um campo vazio na coluna "Status"
    ultima_linha_vazia = coluna_status[coluna_status.isna().index]
    print(ultima_linha_vazia)
    exit()
    
    
    if not ultima_linha_vazia.empty:
        ultima_linha = ultima_linha_vazia[0]
        print(F'Valor ultima linha: {ultima_linha}')
        linha_atual = df.iloc[1].tolist()
        print(linha_atual)
        
        
        
        exit()
        conta_linha = ultima_linha + 2 # +2 porque o DF é zero-based, sendo assim, começa a contar do zero.
        print(F'--- Existem linhas vazias: {linha_atual}')
        print(F'--- Ultima linha com campo status vazio: {conta_linha}')
    else:
        linha_atual = []
        conta_linha = -1
        print('--- Não existem linhas vazias, todas notas lançadas.')
    
    return linha_atual, conta_linha

arquivo = 'db_alltrips/ordenado_db_alltrips.xlsx'
verifica_status_vazio(arquivo)
exit()


#reorganiza_arquivo(nome = "db_alltrips.xlsx")



#Coloca a planilha dentro do "dataframe"
db_alltrips = pd.read_excel('db_alltrips/db_alltrips.xlsx', engine='openpyxl', dtype=object, date_format= "dd/mm/yyyy" )

