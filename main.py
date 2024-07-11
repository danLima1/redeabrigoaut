import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import time
import schedule
import uuid
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv
from flask import Flask, request, jsonify

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

app = Flask(__name__)


# Função para enviar e-mail
def enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port, from_name):
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(from_email, password)

            msg = MIMEMultipart()
            msg['From'] = f"{from_name} <{from_email}>"
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body_html, 'html'))

            server.sendmail(from_email, to_email, msg.as_string())
            print(f"E-mail enviado para {to_email}")
    except Exception as e:
        print(f"Erro ao enviar o e-mail para {to_email}: {e}")


# Função para ler um arquivo HTML
def ler_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()


# Função para enviar e-mails iniciais
def enviar_emails_iniciais():
    from_email = os.getenv("EMAIL")
    from_name = os.getenv("NAME")
    password = os.getenv("EMAIL_PASSWORD")
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    html_inicial = ler_html('modeloHtml/email1.html')

    emails_df = pd.read_excel('emails.xlsx',
                              dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str,
                                     'ConfirmationID': str, 'DataEnvioInicial': str})

    for index, row in emails_df.iterrows():
        to_email = row['Email']
        subject = row['Assunto']
        body = row['Corpo']
        recebido = row['Recebido']
        nome = row.get('Nome', 'Prezado(a)')

        if pd.isna(to_email):
            continue

        confirmation_id = str(uuid.uuid4())
        confirmation_link = f"http://localhost:5000/confirm?email={to_email}&id={confirmation_id}"

        variables = {
            'Nome': nome,
            'nome_da_crianca': 'Maria',
            'presente': 'batman',
            'tamanho': 'M',
            'confirmation_link': confirmation_link
        }

        if pd.isna(recebido) or recebido.lower() != 'sim':
            body_html = html_inicial
            for key, value in variables.items():
                body_html = body_html.replace(f'{{{key}}}', value)

            enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port, from_name)
            emails_df.at[index, 'ConfirmationID'] = confirmation_id
            emails_df.at[index, 'DataEnvioInicial'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        else:
            print(f"E-mail já recebido por {to_email}")

    emails_df.to_excel('emails.xlsx', index=False)


# Função para enviar e-mails de reenvio e remover campos Email e Nome após 7 dias
def enviar_emails_reenvio():
    from_email = os.getenv("EMAIL")
    from_name = os.getenv("NAME")
    password = os.getenv("EMAIL_PASSWORD")
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    html_reenvio = ler_html('modeloHtml/email2.html')

    emails_df = pd.read_excel('emails.xlsx',
                              dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str,
                                     'ConfirmationID': str, 'DataEnvioInicial': str})

    today = datetime.now()

    for index, row in emails_df.iterrows():
        to_email = row['Email']
        subject = "Lembrete: Confirmação de Recebimento"
        body = row['Corpo']
        recebido = row['Recebido']
        nome = row.get('Nome', 'Prezado(a)')
        data_envio_inicial = row.get('DataEnvioInicial')

        if pd.isna(to_email):
            continue

        confirmation_id = row['ConfirmationID']
        confirmation_link = f"http://localhost:5000/confirm?email={to_email}&id={confirmation_id}"

        variables = {
            'Nome': nome,
            'nome_da_crianca': 'Maria',
            'presente': 'batman',
            'tamanho': 'M',
            'confirmation_link': confirmation_link
        }

        if pd.isna(recebido) or recebido.lower() != 'sim':
            body_html = html_reenvio
            for key, value in variables.items():
                body_html = body_html.replace(f'{{{key}}}', value)

            enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port, from_name)
        else:
            print(f"E-mail já recebido por {to_email}")

        if data_envio_inicial and isinstance(data_envio_inicial, str):
            data_envio_inicial_dt = datetime.strptime(data_envio_inicial, '%Y-%m-%d %H:%M:%S')
            elapsed_time = today - data_envio_inicial_dt

            if elapsed_time > timedelta(days=7):
                print(f"Removendo o email e o nome de {to_email} da lista de e-mails")
                emails_df.at[index, 'Email'] = ''
                emails_df.at[index, 'Nome'] = ''
                emails_df.at[index, 'ConfirmationID'] = ''
                emails_df.at[index, 'DataEnvioInicial'] = ''

    emails_df.to_excel('emails.xlsx', index=False)


@app.route('/confirm', methods=['GET'])
def confirm_receipt():
    email = request.args.get('email')
    confirmation_id = request.args.get('id')

    emails_df = pd.read_excel('emails.xlsx',
                              dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str,
                                     'ConfirmationID': str, 'DataEnvioInicial': str})

    for index, row in emails_df.iterrows():
        if row['Email'] == email and row['ConfirmationID'] == confirmation_id:
            emails_df.at[index, 'Recebido'] = 'sim'
            emails_df.to_excel('emails.xlsx', index=False)
            return jsonify({"message": "Confirmação recebida com sucesso!"}), 200

    return jsonify({"message": "Confirmação inválida!"}), 400


# Função principal que configura o agendamento
def main():
    schedule.every(7).days.do(enviar_emails_reenvio)  # Agendar o reenvio a cada 7 dias
    print("Agendamento iniciado")
    enviar_emails_iniciais()
    while True:
        schedule.run_pending()
        time.sleep(1)


if __name__ == "__main__":
    from threading import Thread

    # Iniciar o servidor Flask em uma thread separada
    flask_thread = Thread(target=lambda: app.run(host='0.0.0.0', port=5000))
    flask_thread.start()

    # Executar a função principal de agendamento
    main()
