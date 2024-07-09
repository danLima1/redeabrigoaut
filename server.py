from flask import Flask, request, render_template_string
import pandas as pd
import random

app = Flask(__name__)

# Função para preencher um e-mail aleatoriamente na planilha
def fill_random_email(emails_df, user_email):
    empty_email_indices = emails_df[emails_df['Email'].isna()].index.tolist()

    if empty_email_indices:
        random_index = random.choice(empty_email_indices)
        emails_df.at[random_index, 'Email'] = user_email
    else:
        print("Nenhuma célula vazia na coluna 'Email' para preencher.")

    return emails_df

# Rota para exibir o formulário
@app.route('/')
def index():
    return render_template_string('''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Formulário de E-mail</title>
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
            <h2>Insira seu E-mail</h2>
            <form method="POST" action="/submit">
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

# Rota para processar o formulário
@app.route('/submit', methods=['POST'])
def submit():
    user_email = request.form['email']

    # Ler a planilha, garantindo que a coluna 'Email' seja tratada como string
    emails_df = pd.read_excel('emails.xlsx', dtype={'Email': str, 'Assunto': str, 'Corpo': str, 'Recebido': str})

    # Preencher o e-mail aleatoriamente na planilha
    emails_df = fill_random_email(emails_df, user_email)

    # Salvar a planilha atualizada
    emails_df.to_excel('emails.xlsx', index=False)

    return f"E-mail {user_email} adicionado à planilha com sucesso!"

if __name__ == '__main__':
    app.run(debug=True)
