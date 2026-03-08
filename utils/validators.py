# -*- Validadores de dados para o RPA Cortesia. -*-

from datetime import date, datetime, timedelta
from utils.funcoes import marca_lancado
from utils.configura_logger import get_logger

logger = get_logger("validadores", print_terminal=True)

# Mapeamento de filiais
FILIAIS = {
    '1001': 'VILA',
    '1002': 'CACAPAVA',
    '1003': 'BARUERI',
    '1004': 'JAGUARE',
    '1006': 'ATIBAIA',
    '1008': 'MOGI',
    '1005': 'SANTOS',
    '1032': 'TAMOIO',
    '1036': 'PERUS',
}


def valida_filial_estoque(filial_estoq: str) -> str:
    """
    Valida se a filial de estoque está cadastrada.
    
    Args:
        filial_estoq (str): Código da filial de estoque.
        
    Returns:
        str: Nome da filial formatado.
        
    Raises:
        ValueError: Se a filial não estiver cadastrada.
    """
    centro = FILIAIS.get(filial_estoq)
    if not centro:
        marca_lancado(texto_marcacao="Filial_nao_cadastrada")
        raise ValueError(f'Filial de estoque nao padronizada: {filial_estoq}')
    return centro


def coleta_proximo_dia() -> date:
    """
    Retorna a data de amanhã como objeto date.
    
    Returns:
        date: Data de amanhã.
    """
    return date.today() + timedelta(days=1)


def formata_data_coletada(dados_copiados: str) -> str:
    """
    Formata e valida a data coletada da planilha.
    Verifica se a data não é do próximo dia.
    
    Args:
        dados_copiados (str): Data copiada da planilha.
        
    Returns:
        str: Data formatada como DD/MM/YY.
    """
    # Verifica se a variável está vazia ou contém apenas espaços
    if not dados_copiados.strip():
        logger.info("--- Nenhuma data coletada. Usando a data atual.")
        return date.today().strftime("%d/%m/%y")

    data_copiada = dados_copiados.split(' ')[0]  # Pega apenas a parte da data
    logger.info(f"Data copiada: {data_copiada}, realizando a formatação")
    
    # Converter a string para objeto datetime.date
    data_obj = datetime.strptime(data_copiada, "%d/%m/%Y").date()

    # Obtém a data de amanhã como objeto date
    amanha_data = coleta_proximo_dia()

    # Comparação correta entre objetos date
    if data_obj >= amanha_data:
        logger.info("--- A data coletada é do próximo dia! Alterando para a data atual.")
        return date.today().strftime("%d/%m/%y")  # Retorna a data atual formatada
    
    logger.info("--- A data coletada é válida!")
    return data_obj.strftime("%d/%m/%y")  # Retorna a data coletada formatada


def coleta_valida_dados():
    """
    Coleta e valida dados da planilha do lançamento.
    Aguarda até que os dados sejam válidos.
    
    Returns:
        list: Lista de dados da planilha validados.
    """
    from valida_lancamento import valida_lancamento
    from valida_pedido import main as valida_pedido
    import time
    
    logger.info('--- Executando COLETA VALIDA DADOS ')
    acabou_pedido = False

    while acabou_pedido is False: 
        dados_planilha = valida_lancamento()  # Coleta e confere os dados do lançamento atual
        
        if dados_planilha is None:
            logger.warning('--- valida_lancamento() retornou None, tentando novamente')
            time.sleep(0.2)
            continue
        acabou_pedido = valida_pedido(dados_planilha[4], dados_planilha[15])  # Verifica se o pedido está válido
        time.sleep(0.2)
    else:
        logger.info(f"Dados coletados: {dados_planilha}")
        return dados_planilha
    

if __name__ == '__main__':
    coleta_valida_dados()