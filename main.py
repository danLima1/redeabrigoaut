import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import time

def enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port):
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(from_email, password)

            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body_html, 'html'))

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
        emails_df = pd.read_excel('emails.xlsx', dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str})

        for index, row in emails_df.iterrows():
            to_email = row['Email']
            subject = row['Assunto']
            body = row['Corpo']
            recebido = row['Recebido']
            nome = row.get('Nome', 'Prezado(a)')

            if pd.isna(to_email):
                continue

            if recebido.lower() != 'sim':
                body_html = f'''
                 <html>
                <head>
                    <style>
                        body {{
                            font-family: Arial, sans-serif;
                            color: #333;
                            line-height: 1.6;
                        }}
                        .container {{
                            width: 80%;
                            margin: 0 auto;
                            padding: 20px;
                            border: 1px solid #ddd;
                            border-radius: 10px;
                            background-color: #f9f9f9;
                        }}
                        .header {{
                            background-color: #4CAF50;
                            color: white;
                            padding: 10px;
                            border-radius: 10px 10px 0 0;
                            text-align: center;
                        }}
                        .content {{
                            padding: 20px;
                        }}
                        .footer {{
                            margin-top: 20px;
                            text-align: center;
                            font-size: 0.9em;
                            color: #777;
                        }}
                    </style>
                </head>
                <body>
                    <div class="container">
                        <div class="header">
                            <h2>Rede Abrigo</h2>
                        </div>
                        <div class="content">
                            <p>Prezado(a) {nome},</p>
                            <p>{body}</p>
                        </div>
                        <div class="footer">
                            <p>Atenciosamente,</p>
                            <p><strong>Rede Abrigo</strong></p>
                        </div>
                    </div>
                </body>
                </html>
                '''
                enviar_email(subject, body_html, to_email, from_email, password, smtp_server, smtp_port)
                emails_df.at[index, 'Recebido'] = 'Sim'
            else:
                print(f"E-mail já recebido por {to_email}")

        emails_df.to_excel('emails.xlsx', index=False)
        print("Esperando 10 dias para reenviar e-mails não recebidos...")
        time.sleep(864000)

if __name__ == "__main__":
    main()
