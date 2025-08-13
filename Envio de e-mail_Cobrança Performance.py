import pandas as pd
import win32com.client as win32

# Caminho e aba da planilha
caminho_arquivo = r"C:\Users\cferlin\Downloads\E-mails_Cobrança de metas.xlsx"
aba = "Liderança pendente"

# Lê os dados da aba correta
df = pd.read_excel(caminho_arquivo, sheet_name=aba)

# Agrupa os colaboradores por gestor
agrupado = df.groupby(['Gestor', 'E-mail Gestor'])

# Inicia o Outlook
outlook = win32.Dispatch('outlook.application')

# Itera por cada gestor e gera os e-mails
for (gestor, email), grupo in agrupado:
    # Monta a lista de colaboradores por área com bullet
    corpo_colaboradores = ""
    for area in grupo['Área'].unique():
        nomes = grupo[grupo['Área'] == area]['Colaborador'].tolist()
        lista_nomes = "<br>".join([f"• {nome}" for nome in nomes])
        corpo_colaboradores += f"<p><b>{area}</b><br>{lista_nomes}</p>"

    # Corpo do e-mail com a variável corpo_colaboradores inserida corretamente
    corpo_email = f"""
    <p>Olá, <b>{gestor}</b>!</p>

    <p>Identificamos que ainda não foi registrada a avaliação das pessoas listadas abaixo.</p>
    Reforçamos que esse processo é obrigatório e a <b>Avaliação de Performance Intermediária</b> permanece disponível no Seu Portal, com prazo de conclusão até <b>hoje</b>.

    <p><b>O que precisa ser feito:</b><br>
    • <a href="https://seuportal.br.viterra.online/PAD/SitePages/Home.aspx" target="_blank">Acesse o Seu Portal (PAD - Página Inicial)</a><br>
    • Registre a avaliação para cada meta inserida pelas pessoas do seu time;<br>
    • Finalize na aba “Performance”, atribuindo uma nota de 1 a 5 referente ao período do 1º semestre de 2025.</p>

    <p><b>Colaboradores sem avaliação da liderança:</b></p>
    {corpo_colaboradores}

    <p><i>Estagiários, aprendizes e colaboradores com menos de 6 meses de empresa não são obrigados a participar, mas estão convidados, caso queiram.</i></p>

    <p>Abraços,<br><br>
    <b>Cibely Ferlin</b><br>
    Analista de Recursos Humanos Sr<br>
    Phone: +55 11 97498-7417<br>
    bunge.com<br>
    São Paulo, Brazil</p>
    """

    # Criação do e-mail no Outlook
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = "ATENÇÃO | Avaliação Intermediária de Performance - Avalie a sua equipe"
    mail.HTMLBody = corpo_email
    mail.Display()  # Altere para mail.Send() se quiser enviar diretamente
