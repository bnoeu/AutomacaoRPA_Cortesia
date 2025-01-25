import time
import sqlite3
import pandas as pd
from datetime import datetime


#* Converte a planilha para um bando de dados.
def planilha_para_banco(nome_planilha = "debug_db_alltrips"):
    # Caminhos dos arquivos
    excel_file = F"db_alltrips/{nome_planilha}.xlsx"  # Substitua pelo caminho da sua planilha
    db_file = F"db_alltrips/{nome_planilha}.sqlite"  # Substitua pelo caminho onde o banco será criado

    # Nome da tabela que será criada no banco
    table_name = nome_planilha  # Substitua pelo nome desejado

    try:
        # Carregar a planilha Excel em um DataFrame
        df = pd.read_excel(excel_file, sheet_name=0)  # Ajuste `sheet_name` se necessário

        # Conectar ou criar o banco SQLite
        conn = sqlite3.connect(db_file)
        
        # Salvar o DataFrame no banco como uma tabela
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        print(f"Planilha convertida para banco SQLite com sucesso! Arquivo: {db_file}")
        
    except Exception as e:
        print(f"Erro: {e}")
    finally:
        # Garantir que a conexão seja fechada
        if 'conn' in locals():
            conn.close()

def banco_para_planilha():
    """
    Converte uma tabela de um banco SQLite para uma planilha Excel.

    Parâmetros:
        db_file (str): Caminho para o banco de dados SQLite.
        table_name (str): Nome da tabela no banco de dados.
        output_file (str): Caminho para o arquivo Excel de saída.
    """

    db_file = "db_alltrips/debug_db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    conn = sqlite3.connect(db_file)
    output_file = "db_alltrips/saida_planilha.xlsx"  # Substitua pelo caminho do arquivo Excel de saída


    try:
        # Conectar ao banco SQLite
        conn = sqlite3.connect(db_file)
        
        # Ler a tabela para um DataFrame
        df = pd.read_sql_query("SELECT * FROM debug_db_alltrips", conn)
        
        # Exportar o DataFrame para uma planilha Excel
        df.to_excel(output_file, index=False)
        print(f"Banco convertido para planilha com sucesso! Arquivo: {output_file}")
    
    except Exception as e:
        print(f"Erro: {e}")
    
    finally:
        # Garantir que a conexão seja fechada
        if 'conn' in locals():
            conn.close()


def atualiza_linha(novos_dados = ()):
    # Retorna as linhas da tabela
    db_file = "db_alltrips/debug_db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    conn = sqlite3.connect(db_file)

    # Gera um cursor na base de dados
    cur = conn.cursor() 

    # Atualiza dados de alguma linha
    cur.execute("""
        UPDATE debug_db_alltrips
        SET [status] = ?
        WHERE [cod barras] = ?
    """, novos_dados)

    # Confirme as alterações
    conn.commit()

    # Fechando conexão
    cur.close()
    conn.close()


def insere_linha_debug(linha_atual = ()):
    print(F"Inserindo linha: {linha_atual}")
    # Retorna as linhas da tabela
    db_file = "db_alltrips/debug_db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    conn = sqlite3.connect(db_file)

    # Gera um cursor na base de dados
    cur = conn.cursor() 

    # Atualiza dados de alguma linha
    cur.execute("""
        INSERT INTO debug_db_alltrips
        (RE, [Local Estoque1], [Local Estoque2], [Filial Estoque], [Cod Barras], [__PowerAppsId__], Status, [D. Lancamento], [D. Insercao], posicao)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)     
    """, linha_atual)

    # Confirme as alterações
    conn.commit()

    # Fechando conexão
    cur.close()
    conn.close()

def coleta_banco_original():
    bd_original = "db_alltrips/db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    bd_debug = "db_alltrips/debug_db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    
    # Conexão com os bancos de dados
    conn_original = sqlite3.connect(bd_original)
    conn_debug = sqlite3.connect(bd_debug)

    cursor_original = conn_original.cursor()
    cursor_debug = conn_debug.cursor()


    tabelas = cursor_original.fetchall()
    print("Tabelas no banco de dados:", tabelas)

    exit()


    # Inserindo os novos usuários do banco 'debug_db_alltrips.sqlite' na tabela 'usuarios' do banco 'db_alltrips.sqlite'
    cursor_debug.execute("""
    INSERT INTO debug_db_alltrips (
        RE, [Local Estoque1], [Local Estoque2], [Filial Estoque], [Cod Barras], __PowerAppsId__, Status, [D. Lancamento], [D. Insercao], posicao
    )
    SELECT 
        RE, [Local Estoque1], [Local Estoque2], [Filial Estoque], [Cod Barras], __PowerAppsId__, Status, [D. Lancamento], [D. Insercao], posicao
    FROM lista_notas
    WHERE NOT EXISTS (
        SELECT * FROM debug_db_alltrips
        WHERE debug_db_alltrips.[Cod Barras] = lista_notas.[Cod Barras]
        AND debug_db_alltrips.[__PowerAppsId__] = lista_notas.[__PowerAppsId__]
    )
    """)

    # Confirmando as alterações
    cursor_debug.commit()

    # Verificando os registros finais na tabela 'usuarios'
    cursor_debug.execute("SELECT * FROM usuarios")
    resultados = cursor_debug.fetchall()

    print("Registros finais na tabela 'usuarios':")
    for linha in resultados:
        print(linha)


def remove_valor():
    # Retorna as linhas da tabela
    db_file = "db_alltrips/db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    conn = sqlite3.connect(db_file)

    # Gera um cursor na base de dados
    cur = conn.cursor() 

    # Inserir na tabela apenas se o registro não existir
    cur.execute("""
        DELETE from lista_notas
        WHERE [Local Estoque1] = 'SILO 2'
        
        """)

    conn.commit()  


def verifica_linha_existe(chave_xml = ""):
    # Consulta se uma linha já existe no banco de debug

    # Retorna as linhas da tabela
    db_file = "db_alltrips/debug_db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    conn = sqlite3.connect(db_file)

    # Gera um cursor na base de dados
    cur = conn.cursor() 

    cur.execute("""
    SELECT EXISTS(
        SELECT 1 FROM debug_db_alltrips
        WHERE [Cod Barras] = ?
    )
    """, (chave_xml,))

    existe = cur.fetchone()[0]  # Retorna 1 se o registro existir, 0 caso contrário

    if existe:
        print(F"O registro: {chave_xml} já existe.")
        return False
    else:
        print(F"O registro: {chave_xml} não existe.")
        return True


def consulta_linhas():
    """ Consulta as linhas atuais do banco de ORIGINAL, e insere no banco de debug
    """

    # Retorna as linhas da tabela
    db_file = "db_alltrips/debug_db_alltrips.sqlite"  # Substitua pelo caminho onde o banco será criado
    conn = sqlite3.connect(db_file)

    # Gera um cursor na base de dados
    cur = conn.cursor() 

    # Seleciona informações com o campo "STATUS" vazio e salva na variavel "TABELA"
    #tabela = cur.execute("SELECT * FROM debug_db_alltrips WHERE Status IS NULL")
    
    #* Consulta tudo no banco de debug
    tabela = cur.execute(
        """
            SELECT * 
            FROM debug_db_alltrips 
            WHERE DATE([d. lancamento]) = '2025-01-11'
            AND [status] = 'Lancado_Manual';
        """) 


    linha = tabela.fetchone()
    linha_atual = linha

    # Fechando conexão
    cur.close()
    conn.close()

    return linha_atual



def main():
    linha = consulta_linhas()
    return linha

if __name__ == '__main__':
    banco_para_planilha()
    exit()

    dados = ("TESTE_LANCAMENTO", "33250133051624000197550010003334201000000011", "VbrKV7MAtKY")

    atualiza_linha(dados)
    '''
    hoje = datetime.now()
    hoje_formatado = hoje.strftime('%d/%m/%Y')

    linha = main()
    print(linha)
    '''

'''
    #verifica_linha_existe()

    #remove_valor()
    #exit()
    # Mesmo para UM valor, precisa estar dentro de uma tupla
    #xml = (hoje_formatado, "35240333039223000979550010003644991606443464")

    atualiza_linha(chave_xml= xml)
    #consulta_linhas_original()
    #coleta_banco_original()

    #consulta_linhas()
    #planilha_para_banco()
    #planilha_para_banco(nome_planilha= "db_alltrips")
    #banco_para_planilha()
'''