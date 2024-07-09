import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import time


# essa é a nossa função principal, como se fosse a main do C, mas para enviar e-mails
def send_email(subject, body, to_email, from_email, password, smtp_server, smtp_port):
    # o try funciona como se fosse um if, mas para erros, se tudo der certo, ele executa o try e da uma mensagem de sucesso, se der errado, o exception é executado e da uma mensagem de erro, como se fosse nosso else

    try:
        # essas duas variaveis configuram o servidor SMTP(pesquisar para entender melhor o que é) e a porta, e inicia a conexão TLS(pesquisar tmb)
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()

        # Faz login no servidor SMTP com o email do remetente
        server.login(from_email, password)

        # Cria uma mensagem de e-mail com todas as partes ( de quem o email esta vindo, para quem vai ser enviado, o assunto e etc...)
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        # Anexa o corpo da mensagem como texto simples ( como não estamos usando html ainda, o codigo fica menor)
        msg.attach(MIMEText(body, 'plain'))

        # Converte todo o texto da mensagem para uma string
        text = msg.as_string()

        # Envia o e-mail para o destinatario
        server.sendmail(from_email, to_email, text)
        # mensagem informando que o email foi enviado, a variavel to_email contem o email do destinatario que esta na tabela do excel, isso so ocorre se tudo der certo, se não, o exception acontece
        print(f"E-mail enviado para {to_email}")

        # Encerra a conexão com o servidor SMTP
        server.quit()
    except Exception as e:
        # Em caso de erro, exibe a mensagem de erro
        print(f"Erro ao enviar o e-mail para {to_email}: {e}")


# Configurações do e-mail
from_email = "daniel@redeabrigo.org"  # Endereço de e-mail do remetente
password = "XXXXX"  # Senha de aplicativo ( por favor me pedir a senha para os testes, deixei oculta pois o repositorio esta publico)
smtp_server = "smtp.gmail.com"  # Servidor SMTP do Gmail
smtp_port = 587  # Porta do servidor SMTP (587 para TLS)

# ESSA PARTE DO CODIGO ESTA EM TESTE

# Ler a tabela do Excel
emails_df = pd.read_excel('emails.xlsx')  # Lê os dados do arquivo Excel 'emails.xlsx' para recolher os dados

# Loop para enviar e-mails e reenvio após 10 dias, se necessário
while True:
    for index, row in emails_df.iterrows():

        to_email = row['Email']  # Lê o endereço de e-mail do destinatário
        subject = row['Assunto']  # Lê o assunto do e-mail
        body = row['Corpo']  # Lê o corpo do e-mail
        recebido = row['Recebido']  # Lê o status de recebimento do e-mail

        if recebido.lower() != 'sim':
            # Se o e-mail não foi marcado como recebido, envia o e-mail novamente
            send_email(subject, body, to_email, from_email, password, smtp_server, smtp_port)
            # Marca o e-mail como enviado na tabela
            emails_df.at[index, 'Recebido'] = 'Sim'
        else:
            # Se o e-mail já foi recebido, exibe uma mensagem no console
            print(f"E-mail já recebido por {to_email}")

    # Salvar o arquivo atualizado
    emails_df.to_excel('emails.xlsx', index=False)  # Salva as atualizações no arquivo Excel

    # Esperar 10 dias antes de verificar novamente
    print("Esperando 10 dias para reenviar e-mails não recebidos...")
    time.sleep(864000)  # Espera 864000 segundos (10 dias) antes de executar o loop novamente
