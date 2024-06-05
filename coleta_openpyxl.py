import openpyxl
import time
from onedrivedownloader import download as one_download

# Baixa a planilha
ln = "https://cortesiaconcreto-my.sharepoint.com/:x:/g/personal/bi_cortesiaconcreto_com_br/EW_8FZwWFYVAol4MpV1GglkBJEaJaDx6cfuClnesIu60Ng?e=pveECF"
one_download(ln, filename="db_alltrips")
time.sleep(2)
# Carrega os arquivos
book = openpyxl.load_workbook("db_alltrips/db_alltrips.xlsx")
# Seleciona uma pagina
db_alltrips =  book['db_alltrips']
#Retorna o tamanho maximo da planilha
ultinha_linha = db_alltrips.max_row
# Imprimindo os dados de cada linha
for rows in db_alltrips.iter_rows(min_row= 7487, max_row = 7487):
    if rows[6].value == None:
        print(rows[0][1][2].value)
        print('Nota não lançada!')
    #print(rows[6].value)