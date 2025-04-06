import pandas as pd
import smtplib
from datetime import datetime, timedelta
from email.mime.text import MIMEText


# Minha planilha
arquivo = r"C:\Users\juanf\OneDrive\Documentos\excel\exceltestes.xlsx"            

# Em qual aba está minha planilha 
aba = "Planilha1"
df = pd.read_excel(arquivo, sheet_name=aba)

# Converte a coluna de data corretamente
df['Data de saída'] = pd.to_datetime(df['Data de saída'], errors='coerce', dayfirst=True)

# intervalo da semana atual (segunda a sexta)
hoje = datetime.today()
segunda = hoje - timedelta(days=hoje.weekday())
sexta = segunda + timedelta(days=4)

# Filtra os dados de segunda a sexta
df_semana = df[(df['Data de saída'] >= segunda) & (df['Data de saída'] <= sexta)]

# Agrupa por modelo e soma as quantidades
relatorio = df_semana.groupby('Modelo')['Q saída'].sum().reset_index()

# Remove modelos com quantidade 0
relatorio = relatorio[relatorio['Q saída'] > 0]

# Ordena por quantidade
relatorio = relatorio.sort_values(by='Q saída', ascending=False)

# Exibe no meu terminal
print(f"Relatório semanal dos painéis reparados, segue a lista: \n")

for _, row in relatorio.iterrows():
    print(f"{row['Modelo']} = {int(row['Q saída'])} gabinetes")
