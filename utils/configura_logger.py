import logging
from datetime import datetime
import os

# Variável global para armazenar o nome do arquivo de log
LOG_FILE = None

# Definir o nível personalizado "SUCCESS"
SUCCESS_LEVEL_NUM = 25  # Deve ser um número único não usado por outros níveis
logging.addLevelName(SUCCESS_LEVEL_NUM, "SUCCESS")

# Criar a função para registrar logs no nível "SUCCESS"
def success(self, message, *args, **kwargs):
    if self.isEnabledFor(SUCCESS_LEVEL_NUM):
        self._log(SUCCESS_LEVEL_NUM, message, args, **kwargs)

# Adicionar o método à classe Logger
logging.Logger.success = success

# Função para configurar o logger
def get_logger(name):
    """
    Retorna um logger configurado com um arquivo de log compartilhado.
    """
    global LOG_FILE

    # Garantir que o diretório "logs" exista
    os.makedirs("logs", exist_ok=True)

    # Gerar o nome do arquivo de log apenas na primeira execução
    if LOG_FILE is None:
        horario_inicio = datetime.now()
        LOG_FILE = f"logs/automacao_D{horario_inicio.day}-{horario_inicio.month}__H{horario_inicio.hour}-{horario_inicio.minute}_.log"

    # Configurar o logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)  # Configura o nível mínimo de log

    # Evitar duplicação de handlers
    if not logger.hasHandlers():
        # Configurar handler para escrever em arquivo
        file_handler = logging.FileHandler(LOG_FILE, mode="a", encoding="utf-8")
        file_handler.setLevel(logging.INFO)
        formatter = logging.Formatter(
            fmt="{asctime} - {levelname} - {lineno} = {funcName} - {message}",
            style="{",
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger
