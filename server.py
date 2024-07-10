from flask import Flask, request, render_template_string, url_for, redirect
import pandas as pd
import random

app = Flask(__name__)

# Função para preencher um e-mail aleatoriamente na planilha
def preencher_email_aleatorio(emails_df, user_email, user_nome):
    indices_emails_vazios = emails_df[emails_df['Email'].isna()].index.tolist()

    if indices_emails_vazios:
        indice_aleatorio = random.choice(indices_emails_vazios)
        emails_df.at[indice_aleatorio, 'Email'] = user_email
        emails_df.at[indice_aleatorio, 'Nome'] = user_nome
    else:
        # Adiciona uma nova linha caso não haja espaços vazios
        new_row = {'Email': user_email, 'Nome': user_nome, 'Assunto': '', 'Corpo': '', 'Recebido': ''}
        emails_df = emails_df.append(new_row, ignore_index=True)

    return emails_df

# Função para preencher o assunto e o corpo da planilha
def preencher_assunto_corpo(emails_df, assunto, corpo):
    indices_corpo_vazios = emails_df[emails_df['Corpo'].isna()].index.tolist()

    if indices_corpo_vazios:
        indice_aleatorio = random.choice(indices_corpo_vazios)
        emails_df.at[indice_aleatorio, 'Assunto'] = assunto
        emails_df.at[indice_aleatorio, 'Corpo'] = corpo
    else:
        # Adiciona uma nova linha caso não haja espaços vazios
        new_row = {'Email': '', 'Nome': '', 'Assunto': assunto, 'Corpo': corpo, 'Recebido': ''}
        emails_df = emails_df.append(new_row, ignore_index=True)

    return emails_df

# Rota para a página inicial
@app.route('/')
def home():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Bem-vindo</title>
        <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {
                background: linear-gradient(90deg, #9C27B0, #E040FB);
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
                font-family: Arial, sans-serif;
            }
            .container {
                background-color: white;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                padding: 30px;
                text-align: center;
            }
            .btn {
                background-color: #00C9FF;
                border: none;
                border-radius: 20px;
                padding: 10px;
                width: 100%;
                margin: 10px 0;
                color: white;
                text-transform: uppercase;
                font-weight: bold;
            }
            .btn:hover {
                background-color: #00B2E3;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h2>Bem-vindo à Rede Abrigo</h2>
            <p>Escolha uma das opções abaixo:</p>
            <a href="/form_doador" class="btn">Se inscreva como doador</a>
            <a href="/form_abrigos" class="btn">Inscreva seu abrigo</a>
        </div>
        <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    </body>
    </html>
    ''')

# Rota para exibir o formulário de doador
@app.route('/form_doador')
def form_doador():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Formulário de Doador</title>
        <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {
                background: linear-gradient(90deg, #9C27B0, #E040FB);
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
                font-family: Arial, sans-serif;
            }
            .container {
                background-color: white;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                padding: 30px;
                width: 350px;
                text-align: center;
            }
            .container h2 {
                margin-bottom: 20px;
            }
            .form-control {
                border-radius: 20px;
                padding: 10px;
                margin-bottom: 15px;
            }
            .btn-primary {
                background-color: #00C9FF;
                border: none;
                border-radius: 20px;
                padding: 10px;
                width: 100%;
                margin-top: 10px;
            }
            .btn-primary:hover {
                background-color: #00B2E3;
            }
            .footer {
                margin-top: 20px;
                color: #555;
            }
            .logo {
                width: 100px;
                margin-bottom: 20px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <img src="{{ url_for('static', filename='Cartinha-Acolhedora-300x146.png') }}" alt="Cartinha Acolhedora" class="logo">
            <h2>Insira seu Nome e E-mail</h2>
            <form method="POST" action="/submit_doador">
                <div class="form-group">
                    <input type="text" class="form-control" id="nome" name="nome" placeholder="Nome" required>
                </div>
                <div class="form-group">
                    <input type="email" class="form-control" id="email" name="email" placeholder="Email" required>
                </div>
                <button type="submit" class="btn btn-primary">Enviar</button>
            </form>
            <div class="footer">
                <p><strong>Cartinha Acolhedora</strong><br>
                Feito por Rede Abrigo</p>
            </div>
        </div>
        <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    </body>
    </html>
    ''')

# Rota para exibir o formulário para os abrigos
@app.route('/form_abrigos')
def form_abrigos():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Formulário de Abrigos</title>
        <link href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {
                background: linear-gradient(90deg, #9C27B0, #E040FB);
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
                font-family: Arial, sans-serif;
            }
            .container {
                background-color: white;
                border-radius: 10px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
                padding: 30px;
                width: 350px;
                text-align: center;
            }
            .container h2 {
                margin-bottom: 20px;
            }
            .form-control {
                border-radius: 20px;
                padding: 10px;
                margin-bottom: 15px;
            }
            .btn-primary {
                background-color: #00C9FF;
                border: none;
                border-radius: 20px;
                padding: 10px;
                width: 100%;
                margin-top: 10px;
            }
            .btn-primary:hover {
                background-color: #00B2E3;
            }
            .footer {
                margin-top: 20px;
                color: #555;
            }
            .logo {
                width: 100px;
                margin-bottom: 20px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <img src="{{ url_for('static', filename='Cartinha-Acolhedora-300x146.png') }}" alt="Cartinha Acolhedora" class="logo">
            <h2>Insira o Assunto e a Mensagem</h2>
            <form method="POST" action="/submit_abrigos">
                <div class="form-group">
                    <input type="text" class="form-control" id="assunto" name="assunto" placeholder="Assunto" required>
                </div>
                <div class="form-group">
                    <textarea class="form-control" id="corpo" name="corpo" placeholder="Mensagem" rows="4" required></textarea>
                </div>
                <button type="submit" class="btn btn-primary">Enviar</button>
            </form>
            <div class="footer">
                <p><strong>Cartinha Acolhedora</strong><br>
                Feito por Rede Abrigo</p>
            </div>
        </div>
        <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/@popperjs/core@2.5.4/dist/umd/popper.min.js"></script>
        <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
    </body>
    </html>
    ''')

# Rota para processar o formulário de doador
@app.route('/submit_doador', methods=['POST'])
def submit_doador():
    user_nome = request.form['nome']
    user_email = request.form['email']

    # Ler a planilha, garantindo que a coluna 'Email' e 'Nome' sejam tratadas como string
    emails_df = pd.read_excel('emails.xlsx', dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str})

    # Preencher o e-mail e nome aleatoriamente na planilha
    emails_df = preencher_email_aleatorio(emails_df, user_email, user_nome)

    # Salvar a planilha atualizada
    emails_df.to_excel('emails.xlsx', index=False)

    return f"E-mail {user_email} e nome {user_nome} adicionados à planilha com sucesso!"

# Rota para processar o formulário dos abrigos
@app.route('/submit_abrigos', methods=['POST'])
def submit_abrigos():
    assunto = request.form['assunto']
    corpo = request.form['corpo']

    # Ler a planilha, garantindo que as colunas 'Assunto' e 'Corpo' sejam tratadas como string
    emails_df = pd.read_excel('emails.xlsx', dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str})

    # Preencher o assunto e corpo aleatoriamente na planilha
    emails_df = preencher_assunto_corpo(emails_df, assunto, corpo)

    # Salvar a planilha atualizada
    emails_df.to_excel('emails.xlsx', index=False)

    return f"Assunto e mensagem adicionados à planilha com sucesso!"

if __name__ == "__main__":
    app.run(debug=True)
