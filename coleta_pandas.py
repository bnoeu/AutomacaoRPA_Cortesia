import pandas as pd
from onedrivedownloader import download as one_download

'''
ln = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EW_8FZwWFYVAol4MpV1GglkBJEaJaDx6cfuClnesIu60Ng?e=pveECF"
one_download(ln, filename="db_alltrips")
'''

#Coloca a planilha dentro do "dataframe"
db_alltrips = pd.read_excel('db_alltrips/db_alltrips.xlsx', engine='openpyxl', dtype=object, date_format= "dd/mm/yyyy" )

#print(db_alltrips.loc[0, 1])
#Converte o dataframe numa lista, para pegar o ultimo valor
lista_alltrips = db_alltrips.values.tolist()
ultimo_registro = len(db_alltrips)
print(len(lista_alltrips))

#print(list(linha_atual))

for x in len(db_alltrips):
    linha_atual = db_alltrips.loc[x]
    if linha_atual['Status'] == 'Nan':
        print('Status vazio')
        break
    else:
        v_linha = linha_atual['Status']
        print(F'Valor da linha: {v_linha}')


exit()  
db_alltrips.at[ultimo_registro, 'Status'] = 'Bruno'
#db_alltrips.at[ultimo_registro, 'produtos'] = 'Bruno'
#db_alltrips.loc[2716:2716, ['Status']] = ['Teste_pandas1']
#print(type(db_alltrips))
#print(lista_alltrips[1])
db_alltrips.to_excel('db_alltrips/db_producao.xlsx', index= True)

#print(datetime.datetime.now())