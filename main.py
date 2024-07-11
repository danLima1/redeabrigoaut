import smtplib  # Biblioteca para enviar e-mails usando o protocolo SMTP
from email.mime.multipart import MIMEMultipart  # Classe para criar e-mails com múltiplas partes (texto, anexos, etc.)
from email.mime.text import MIMEText  # Classe para criar objetos de texto para o e-mail
import pandas as pd  # Biblioteca para manipulação de dados, especialmente arquivos do Excel
import time  # Biblioteca para manipulação de tempo (pausas, etc.)
import schedule  # Biblioteca para agendamento de tarefas
import uuid  # Biblioteca para gerar identificadores únicos
from flask import Flask, request, jsonify  # Biblioteca para criar um servidor web

# Função para enviar e-mail
def enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port, from_name):
    try:
        # Inicia uma conexão com o servidor SMTP usando o contexto "with" para garantir que a conexão seja fechada corretamente
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Inicia a conexão TLS (Transport Layer Security) para segurança
            server.login(from_email, password)  # Faz login no servidor SMTP com as credenciais fornecidas

            # Cria uma mensagem de e-mail com múltiplas partes
            msg = MIMEMultipart()
            msg['From'] = f"{from_name} <{from_email}>"  # Define o remetente do e-mail
            msg['To'] = to_email  # Define o destinatário do e-mail
            msg['Subject'] = subject  # Define o assunto do e-mail
            msg.attach(MIMEText(body_html, 'html'))  # Anexa o corpo do e-mail no formato HTML

            server.sendmail(from_email, to_email, msg.as_string())  # Envia o e-mail
            print(f"E-mail enviado para {to_email}")  # Imprime uma mensagem de confirmação
    except Exception as e:
        print(f"Erro ao enviar o e-mail para {to_email}: {e}")  # Em caso de erro, imprime a mensagem de erro

# Função para ler um arquivo HTML
def ler_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()  # Lê e retorna o conteúdo do arquivo HTML

# Função para enviar e-mails iniciais
def enviar_emails_iniciais():
    from_email = "daniel@redeabrigo.org"  # E-mail do remetente
    from_name = "Daniel Mendes"  # Nome do remetente
    password = "XXXXXXXXX"  # Senha do e-mail do remetente
    smtp_server = "smtp.gmail.com"  # Servidor SMTP do Gmail
    smtp_port = 587  # Porta do servidor SMTP (587 para TLS)

    html_inicial = ler_html('modeloHtml/email1.html')  # Lê o conteúdo do arquivo HTML inicial

    # Lê a tabela do Excel com os dados dos e-mails
    emails_df = pd.read_excel('emails.xlsx', dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str, 'ConfirmationID': str})

    # Itera sobre cada linha do DataFrame
    for index, row in emails_df.iterrows():
        to_email = row['Email']  # Endereço de e-mail do destinatário
        subject = row['Assunto']  # Assunto do e-mail
        body = row['Corpo']  # Corpo do e-mail
        recebido = row['Recebido']  # Status de recebimento
        nome = row.get('Nome', 'Prezado(a)')  # Nome do destinatário (ou 'Prezado(a)' se não especificado)

        if pd.isna(to_email):  # Verifica se o endereço de e-mail está ausente (NaN)
            continue  # Pula para a próxima iteração

        # Gera um identificador único para o link de confirmação
        confirmation_id = str(uuid.uuid4())
        confirmation_link = f"http://localhost:5000/confirm?email={to_email}&id={confirmation_id}"

        # Variáveis para substituir no HTML do e-mail
        variables = {
            'Nome': nome,
            'nome_da_crianca': 'Maria',
            'presente': 'batman',
            'tamanho': 'M',
            'confirmation_link': confirmation_link
        }

        # Verifica se o e-mail já foi recebido
        if pd.isna(recebido) or recebido.lower() != 'sim':
            body_html = html_inicial
            # Substitui placeholders com os valores reais
            for key, value in variables.items():
                body_html = body_html.replace(f'{{{key}}}', value)

            # Envia o e-mail
            enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port, from_name)
            # Armazena o identificador único no DataFrame
            emails_df.at[index, 'ConfirmationID'] = confirmation_id
        else:
            print(f"E-mail já recebido por {to_email}")

    emails_df.to_excel('emails.xlsx', index=False)  # Salva as alterações no arquivo Excel

# Função para enviar e-mails de reenvio
def enviar_emails_reenvio():
    from_email = "daniel@redeabrigo.org"  # E-mail do remetente
    from_name = "Daniel Mendes"  # Nome do remetente
    password = "XXXXXXXX"  # Senha do e-mail do remetente
    smtp_server = "smtp.gmail.com"  # Servidor SMTP do Gmail
    smtp_port = 587  # Porta do servidor SMTP (587 para TLS)

    html_reenvio = ler_html('modeloHtml/email2.html')  # Lê o conteúdo do arquivo HTML de reenvio

    # Lê a tabela do Excel com os dados dos e-mails
    emails_df = pd.read_excel('emails.xlsx', dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str, 'ConfirmationID': str})

    # Itera sobre cada linha do DataFrame
    for index, row in emails_df.iterrows():
        to_email = row['Email']  # Endereço de e-mail do destinatário
        subject = "Lembrete: Confirmação de Recebimento"  # Assunto do e-mail de reenvio
        body = row['Corpo']  # Corpo do e-mail
        recebido = row['Recebido']  # Status de recebimento
        nome = row.get('Nome', 'Prezado(a)')  # Nome do destinatário (ou 'Prezado(a)' se não especificado)

        if pd.isna(to_email):  # Verifica se o endereço de e-mail está ausente (NaN)
            continue  # Pula para a próxima iteração

        confirmation_id = row['ConfirmationID']  # Usa o mesmo ID gerado anteriormente
        confirmation_link = f"http://localhost:5000/confirm?email={to_email}&id={confirmation_id}"

        # Variáveis para substituir no HTML do e-mail
        variables = {
            'Nome': nome,
            'nome_da_crianca': 'Maria',
            'presente': 'batman',
            'tamanho': 'M',
            'confirmation_link': confirmation_link
        }

        # Verifica se o e-mail já foi recebido
        if pd.isna(recebido) or recebido.lower() != 'sim':
            body_html = html_reenvio
            # Substitui placeholders com os valores reais
            for key, value in variables.items():
                body_html = body_html.replace(f'{{{key}}}', value)

            # Envia o e-mail
            enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port, from_name)
        else:
            print(f"E-mail já recebido por {to_email}")

    emails_df.to_excel('emails.xlsx', index=False)  # Salva as alterações no arquivo Excel

# Função principal que configura o agendamento
def main():
    # Agendar a função enviar_emails_reenvio para rodar a cada 10 dias
    schedule.every(10).days.do(enviar_emails_reenvio)
    enviar_emails_iniciais()  # Executa a função enviar_emails_iniciais imediatamente
    while True:
        schedule.run_pending()  # Executa as tarefas agendadas se houver
        time.sleep(1)  # Aguarda 1 segundo antes de verificar novamente

if __name__ == "__main__":
    main()  # Executa a função principal se o script for executado diretamente
