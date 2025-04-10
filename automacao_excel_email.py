import pandas as pd
import smtplib
from datetime import datetime, timedelta
from email.mime.text import MIMEText

print("Vamos enviar um email com os dados do relatório semanal sobre os painéis de LED\n")

#criação de uma entrada de dados que precisa ser preenchida manualmente 
painel_pl2_5 = int(input("Digite a quantide de módulos do pl 2.5mm:\n"))

#criação de uma segunda entrada de dados que precisa ser preenchida manualmente 
painel_pl2_5c = int(input("Digite a quantide de módulos do pl 2.5mm C:\n"))

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



#---------------- Parte Email ----------------#

assunto = f"Itens devolvidos ao estoque do dia {segunda.strftime('%d/%m/%Y')} até o dia {sexta.strftime('%d/%m/%Y')}"

corpo_mensagem =  f"Segue relação de painéis devolvidos ao estoque do dia – \n\n" + "\n".join(
[f"{row['Modelo']} - {int(row['Q saída'])} gabinetes" for _, row in relatorio.iterrows()]),"\n \nAtenciosamente, Manutenção!"


destinatario = ["Email destino"]

#Armanezamento do login

remetente = "*********@gmail.com"
senha = "senha"

#funções da biblioteca 

msg = MIMEText(corpo_mensagem)
msg['Subject'] = assunto
msg['From'] = remetente
msg['To'] = ', '.join(destinatario)

#conexão com o servidor do gmail
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
  smtp_server.login(remetente, senha)
  smtp_server.sendmail(remetente, destinatario, msg.as_string())
print("Mensagem enviada!")
