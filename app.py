from flask import Flask, render_template, request
import imaplib
import email
from email.header import decode_header
import pandas as pd
import matplotlib.pyplot as plt
import os

app = Flask(__name__)

app.config['SESSION_COOKIE_SECURE'] = True  # Cookie só é enviado via HTTPS
app.config['SESSION_COOKIE_HTTPONLY'] = True  # Cookie não acessível via JavaScript
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'  # Protege contra alguns tipos de CSRF


# PESQUISA = "TESTE"  # O título de pesquisa para os emails

def fetch_email_data(EMAIL, PASSWORD, PESQUISA):
    # Configurações do email
    # EMAIL = os.getenv('EMAIL')
    # PASSWORD = os.getenv('PASSWORD')
    IMAP_SERVER = "imap.gmail.com"
    IMAP_PORT = 993

    if not EMAIL or not PASSWORD:
        raise ValueError("As variáveis de ambiente EMAIL e PASSWORD devem estar definidas.")

    try:
        # Conectar ao servidor IMAP
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        PESQUISA_EMAIL = f'(SUBJECT "{PESQUISA}")'
        # Buscar e-mails com o título especificado
        status, messages = mail.search(None, PESQUISA_EMAIL)

        if status != 'OK':
            raise ValueError("Erro ao buscar e-mails.")

        # Lista para armazenar os dados dos e-mails
        data = []

        for num in messages[0].split():
            try:
                status, msg_data = mail.fetch(num, "(RFC822)")
                if status != 'OK':
                    raise ValueError(f"Erro ao buscar e-mail com número {num}.")
                
                msg = email.message_from_bytes(msg_data[0][1])

                # Decodificar o assunto do e-mail
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding)

                # Obter a data do e-mail
                date = msg["Date"]

                # Adicionar a data e o assunto na lista
                data.append({"Date": date, "Subject": subject})

            except Exception as e:
                print(f"Erro ao processar o e-mail {num}: {e}")

        # Fechar a conexão com o servidor de e-mail
        mail.close()
        mail.logout()

        # Converter os dados em um DataFrame do pandas
        df = pd.DataFrame(data)

        # Converter a coluna 'Date' para datetime
        df["Date"] = pd.to_datetime(df["Date"], errors='coerce', utc=True)

        # Remover linhas onde a conversão falhou
        df = df.dropna(subset=["Date"])

        # Contar o número de chamados por dia
        daily_counts = df.groupby(df["Date"].dt.date)["Subject"].count()

        return df, daily_counts

    except Exception as e:
        print(f"Erro geral: {e}")
        return pd.DataFrame(), pd.Series()

def create_plot(daily_counts, PESQUISA):
    if daily_counts.empty:
        print("Nenhum dado disponível para plotar.")
        return
    
    plt.figure(figsize=(10, 6))
    daily_counts.plot(kind="bar", color="blue")
    plt.title(f"Quantidade de Chamados por Dia ({PESQUISA})")
    plt.xlabel("Data")
    plt.ylabel("Número de Chamados")
    plt.xticks(rotation=30)
    
    if not os.path.exists('static'):
        os.makedirs('static')
    
    plt.savefig('static/plot.png')
    plt.close()
    

@app.route('/', methods=['GET', 'POST'])
def pesquisa():
    if request.method == "POST":
        EMAIL = request.form.get('EMAIL')
        PASSWORD = request.form.get('PASSWORD')
        PESQUISA = request.form.get('PESQUISA')
        # Obter os dados do e-mail
        email_data, daily_counts = fetch_email_data(EMAIL, PASSWORD, PESQUISA)
        
        # Criar o gráfico
        create_plot(daily_counts, PESQUISA)

        # Renderizar o template com o gráfico e os dados
        return render_template('index.html', plot_url='static/plot.png', emails=email_data.to_dict(orient='records'), PESQUISA=PESQUISA)
    else:
        return render_template('form.html')


if __name__ == '__main__':
    context = ('cert.pem', 'key.pem')
    app.run(host='0.0.0.0', port=443, ssl_context=context, debug=True)