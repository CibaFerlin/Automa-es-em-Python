import pandas as pd
import win32com.client as win32

# Caminho e aba da planilha
caminho_arquivo = r"C:\Users\cferlin\Downloads\E-mails_Cobrança de metas.xlsx"
aba = "Liderança + Liderado pendente"

# Lê os dados da aba correta
df = pd.read_excel(caminho_arquivo, sheet_name=aba)

# Inicia o Outlook
outlook = win32.Dispatch("outlook.application")

# Itera por cada linha (colaborador + gestor)
for index, row in df.iterrows():
    colaborador = row['Colaborador']
    email_colaborador = row['E-mail']
    gestor = row['Gestor']
    email_gestor = row['E-mail Gestor']

    # Corpo do e-mail
    corpo_email = f"""
    <p>Olá, <b>{colaborador}</b>!</p>

    <p>Identificamos que sua autoavaliação ainda não foi registrada no Seu Portal.</p>
    Por esse motivo, a liderança responsável (<b>{gestor}</b>) também não conseguiu concluir sua avaliação.

    <p>Reforçamos que esse processo é obrigatório e a <b>Avaliação de Performance Intermediária</b> está disponível até <b>o final da semana</b>.</p>

    <p><b>O que precisa ser feito:</b><br>
    • <a href="https://seuportal.br.viterra.online/PAD/SitePages/Home.aspx" target="_blank">Acesse o Seu Portal (PAD - Página Inicial)</a><br>
    • Registre a avaliação para cada meta inserida;<br>
    • Finalize na aba “Performance”, atribuindo uma nota de 1 a 5 referente ao período do 1º semestre de 2025.</p>

    <p>Abraços,<br><br>
    <b>Cibely Ferlin</b><br>
    Analista de Recursos Humanos Sr<br>
    Phone: +55 11 97498-7417<br>
    bunge.com<br>
    São Paulo, Brazil</p>
    """

    # Criação do e-mail
    mail = outlook.CreateItem(0)
    mail.To = email_colaborador
    mail.CC = email_gestor
    mail.Subject = "ATENÇÃO | Avaliação Intermediária de Performance - Realize a sua autoavaliação e finalize o processo"
    mail.HTMLBody = corpo_email
    mail.Display()  # Altere para .Send() se quiser enviar direto
