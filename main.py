import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import time


def send_email(subject, body, to_email, from_email, password, smtp_server, smtp_port):
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(from_email, password)

            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            server.sendmail(from_email, to_email, msg.as_string())
            print(f"E-mail enviado para {to_email}")
    except Exception as e:
        print(f"Erro ao enviar o e-mail para {to_email}: {e}")


def main():
    from_email = "daniel@redeabrigo.org"
    password = "XXXXXX"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    while True:
        emails_df = pd.read_excel('emails.xlsx')

        for index, row in emails_df.iterrows():
            to_email = row['Email']
            subject = row['Assunto']
            body = row['Corpo']
            recebido = row['Recebido']

            if pd.isna(to_email):
                continue

            if recebido.lower() != 'sim':
                send_email(subject, body, to_email, from_email, password, smtp_server, smtp_port)
                emails_df.at[index, 'Recebido'] = 'Sim'
            else:
                print(f"E-mail já recebido por {to_email}")

        emails_df.to_excel('emails.xlsx', index=False)
        print("Esperando 10 dias para reenviar e-mails não recebidos...")
        time.sleep(864000)


if __name__ == "__main__":
    main()
