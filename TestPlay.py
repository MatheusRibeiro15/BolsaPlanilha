import requests
import yfinance as yf
import pandas as pd

# Obtém os 50 maiores ativos da B3 usando a API do brapi
url = "https://brapi.dev/api/quote/list?sort=volume&limit=25"
response = requests.get(url)

if response.status_code == 200:
    tickers_data = response.json()
    tickers = [stock["stock"] + ".SA" for stock in tickers_data["stocks"]]  # Adiciona ".SA" para ações brasileiras
else:
    print("Erro ao buscar a lista de ativos.")
    tickers = []  # Caso não consiga obter, mantém a lista vazia

# Lista para armazenar os dados dos ativos
data_list = []

# Obtendo os dados de cada ativo
for ticker in tickers:
    stock = yf.Ticker(ticker)
    history = stock.history(period="1y")  # Obtém 1 ano de histórico para cálculos

    if not history.empty:
        last_quote = history.iloc[-1]  # Última cotação disponível
        last_date = last_quote.name.to_pydatetime().replace(tzinfo=None)  # Removendo timezone

        # Preço de fechamento em diferentes períodos
        close_today = last_quote["Close"]
        close_week = history["Close"].iloc[-6] if len(history) > 6 else None  # 5 dias úteis atrás
        close_month = history["Close"].iloc[-22] if len(history) > 22 else None  # 21 dias úteis atrás
        close_year = history["Close"].iloc[0] if len(history) > 252 else None  # 1 ano atrás

        # Cálculo das variações percentuais
        rendimento_semanal = ((close_today - close_week) / close_week * 100) if close_week else None
        rendimento_mensal = ((close_today - close_month) / close_month * 100) if close_month else None
        rendimento_anual = ((close_today - close_year) / close_year * 100) if close_year else None

        data_list.append({
            "Ativo": ticker,
            "Data": last_date,
            "Abertura": last_quote["Open"],
            "Máxima": last_quote["High"],
            "Mínima": last_quote["Low"],
            "Fechamento": close_today,
            "Volume": last_quote["Volume"],
            "Rendimento Semanal (%)": rendimento_semanal,
            "Rendimento Mensal (%)": rendimento_mensal,
            "Rendimento Anual (%)": rendimento_anual
        })

# Criando um DataFrame
df = pd.DataFrame(data_list)

# Salvando os dados em uma planilha Excel
df.to_excel("top50_bolsa_brasileira.xlsx", index=False)

print("Arquivo 'top50_bolsa_brasileira.xlsx' gerado com sucesso!")
