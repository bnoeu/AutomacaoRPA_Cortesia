import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from utils.funcoes import print_erro

def enviar_email(destinatario, assunto, mensagem):
    # Configurações do servidor de e-mail
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    remetente = "leigan0301@gmail.com"
    senha = "afnh dcfu jiya sang"

    # Criação da mensagem
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Adiciona o corpo da mensagem
    msg.attach(MIMEText(mensagem, 'plain'))

    # Caminho da imagem a ser anexada
    caminho_imagem = print_erro()
    #caminho_imagem = "imagens/img_geradas/erros/erro2024-11-27 16_04_56_935859.png"

    # Adicionar a imagem como anexo
    with open(caminho_imagem, "rb") as imagem:
        anexo = MIMEBase("application", "octet-stream")
        anexo.set_payload(imagem.read())

    # Codificar o anexo em base64
    encoders.encode_base64(anexo)

    # Adicionar cabeçalho com o nome do arquivo
    anexo.add_header("Content-Disposition", f"attachment; filename={caminho_imagem.split('/')[-1]}")

    # Anexar a imagem ao e-mail
    msg.attach(anexo)

    try:
        # Conecta ao servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls() # Usa TLS para segurança
        server.login(remetente, senha) # Login no servidor SMTP

        server.sendmail(remetente, destinatario, msg.as_string())
        print("E-mail enviado com sucesso!")

    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

    finally:
        server.quit() # Encerra a conexão com o servidor SMTP

# Exemplo de uso
if __name__ == '__main__':
    enviar_email("brunobola2010@gmail.com", "Assunto do E-mail", "Corpo do e-mail")