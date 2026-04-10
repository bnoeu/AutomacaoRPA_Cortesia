import time
import traceback
from .comunicacao_chat import msg_chat
from .enviar_email import enviar_email
from .configura_logger import get_logger
from .funcoes import (
    print_erro, trata_erro, verifica_horario
)

# Obter logger configurado (similar ao usado em Materia_Prima.py)
logger = get_logger("automacao", print_terminal=True)


def handle_exception(ultimo_erro, tentativa, tempo_pausa, arquivo_erro, mensagem_erro):
    """
    Trata exceções ocorridas durante o loop principal da automação, realizando logging, pausas progressivas,
    verificações de horário e envio de notificações por e-mail em casos críticos. Decide se o loop deve continuar
    ou ser interrompido com base no número de tentativas.

    Esta função é chamada dentro do bloco except do run_main_loop para centralizar o tratamento de erros,
    evitando repetição de código e facilitando manutenções.

    Args:
        ultimo_erro (Exception): O objeto de exceção capturado, contendo detalhes do erro ocorrido.
        tentativa (int): O número da tentativa atual de execução (inicia em 0 e incrementa a cada erro).
        tempo_pausa (int): O tempo atual de pausa em segundos antes da próxima tentativa (inicia em 600).
        arquivo_erro (str): Uma string vazia inicialmente; será atualizada com o nome do arquivo ou tarefa onde o erro ocorreu.
        mensagem_erro (str): Uma string vazia inicialmente; será atualizada com a mensagem detalhada do erro.

    Returns:
        tuple: Uma tupla contendo três elementos:
            - bool: True se o loop deve continuar, False se deve ser interrompido.
            - int: O valor atualizado de tentativa (incrementado em 1).
            - int: O valor atualizado de tempo_pausa (pode aumentar progressivamente).
    """
    
    arquivo_erro, mensagem_erro = trata_erro(ultimo_erro, tentativa)
    print_erro()
    logger.exception(F'--- A execução principal apresentou erro! Tentativa: {tentativa}, Pausa anterior: {tempo_pausa}, Erro: {mensagem_erro}')
    print(F'--- A execução principal apresentou erro! Tentativa: {tentativa}, Pausa anterior: {tempo_pausa}')

    #* Realiza as verificações antes da proxima tentativa
    verifica_horario()

    if tentativa >= 9:
        enviar_email(
            "brunobola2010@gmail.com",
            f"[RPA Cortesia] Erro catastrofico: {arquivo_erro}",
            f"Erro coletado: \n{traceback.format_exc()}"
        )
        msg_chat(f'A execução principal apresentou erro! Executando o script principal novamente, tentativa: {tentativa}')
        logger.critical(F'--- A execução principal apresentou erro! Executando o script principal novamente, tentativa: {tentativa}')
        return False, tentativa + 1, tempo_pausa  # Return False to indicate break

    elif tentativa >= 4:  # Começa a pausar o script após a 5º execução
        logger.info(F"Pausando por: {tempo_pausa} segundos antes da proxima tentativa")
        time.sleep(tempo_pausa)
        tempo_pausa = min(int(tempo_pausa + (0.5 * tempo_pausa)), 3600)  # Limita a pausa máxima a 1 hora (3600 segundos)


    return True, tentativa + 1, tempo_pausa  # Return True to continue


if __name__ == '__main__':
    # Teste simples da função com uma exceção simulada
    try:
        raise ValueError("Erro de teste para validação da função handle_exception")
    except Exception as e:
        # Chamada de teste com valores iniciais
        resultado = handle_exception(e, 6, 600, "", "")
        print(f"Resultado do teste: {resultado}")