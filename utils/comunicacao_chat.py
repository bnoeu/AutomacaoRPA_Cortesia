import requests


def msg_chat(teste, erro_critico=False):
    url = "https://chat.googleapis.com/v1/spaces/AAQAP0Y4j00/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=_Oohs5lpjeQc3DPmWYyZxAw1GsnktcFkCUzKBk77UlI"


    if erro_critico:
        titulo = '[NFE] APRESENTOU ERRO CRITICO ❌'
    else:
        titulo = '[NFE] Enviando comunicação para o chat...'

    payload = {
        "text": 
        f"""{titulo} \nULTIMO ERRO: {teste}"""
    }

    response = requests.post(url, json=payload)

    return response


if __name__ == '__main__':
    msg_chat("TESTE", erro_critico=False)