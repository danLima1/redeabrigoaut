from flask import Flask, request, jsonify
import pandas as pd

app = Flask(__name__)


@app.route('/confirm', methods=['GET'])
def confirm_email():
    email = request.args.get('email')
    confirmation_id = request.args.get('id')

    # Lê a tabela do Excel com os dados dos e-mails
    emails_df = pd.read_excel('emails.xlsx',
                              dtype={'Email': str, 'Nome': str, 'Assunto': str, 'Corpo': str, 'Recebido': str,
                                     'ConfirmationID': str})

    # Procura o e-mail e o ID de confirmação no DataFrame
    email_row = emails_df[(emails_df['Email'] == email) & (emails_df['ConfirmationID'] == confirmation_id)]

    if not email_row.empty:
        # Atualiza o status de recebimento para 'Sim'
        emails_df.loc[email_row.index, 'Recebido'] = 'Sim'

        # Salva as alterações no arquivo Excel
        emails_df.to_excel('emails.xlsx', index=False)

        return jsonify({"status": "success", "message": "E-mail confirmado com sucesso!"})
    else:
        return jsonify({"status": "error", "message": "Confirmação inválida!"})


if __name__ == "__main__":
    app.run(debug=True)
