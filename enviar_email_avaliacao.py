import pandas as pd
import win32com.client as win32

# === Caminho da planilha ===
arquivo = r'C:\Users\cferlin\Downloads\verificacao_metas_corrigidas_formatada.xlsx'

# === Carregar e limpar ===
df = pd.read_excel(arquivo, sheet_name='Sheet1', dtype=str)
df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
print("✅ Planilha carregada com sucesso!")
print("📋 Colunas encontradas:", df.columns.tolist())

# === Filtrar linhas com campos essenciais preenchidos ===
colunas_obrigatorias = [
    'Colaborador', 'E-mail', 'Gestor', 'E-mail liderança',
    'Título Sistema', 'Descrição Sistema', 'Sugestão de Melhoria'
]
df_filtrado = df.dropna(subset=colunas_obrigatorias)
df_filtrado = df_filtrado[
    df_filtrado[colunas_obrigatorias].apply(
        lambda col: col.map(lambda x: str(x).strip() != '')
    ).all(axis=1)
]

# === Preencher campos de nota com string vazia ===
df_filtrado['Nota Autoavaliação'] = df_filtrado['Nota Autoavaliação'].fillna('')

print(f"🔎 Linhas válidas para e-mail: {len(df_filtrado)}")

if df_filtrado.empty:
    print("⚠️ Nenhuma linha com dados completos e sugestão de melhoria.")
    exit()

# === Agrupar colaboradores mesmo com campos vazios ===
colaboradores = df_filtrado.groupby([
    'Colaborador', 'E-mail', 'Gestor', 'E-mail liderança', 'Nota Autoavaliação'
])

# === Iniciar Outlook ===
outlook = win32.Dispatch("Outlook.Application")
portal_link = '<a href="http://arsassrv6apt02:84/PAD/SitePages/Home.aspx" target="_blank">Seu Portal</a>'
emails_gerados = 0

for (colab, email_colab, gestor, email_gestor, nota_auto), grupo in colaboradores:
    metas_texto = ""
    for _, linha in grupo.iterrows():
        metas_texto += f"""
        <li>
            <p><strong>{linha['Título Sistema']}</strong></p>
            <p>{linha['Descrição Sistema']}</p>
            <p><strong>Corrija a meta aplicando a seguinte observação:</strong> {linha['Sugestão de Melhoria']}</p>
        </li>
        """

    if str(nota_auto).strip() == "":
        trecho_avaliacao = f"""
        <p>E, até o momento, identificamos que você ainda não realizou sua <strong>Autoavaliação</strong> no sistema.</p>
        <p>Sendo assim, finalize este processo até o dia <strong>30/07/2025</strong> acessando o {portal_link}, corrijindo as metas abaixo aplicando a <strong>Metodologia SMART</strong>.</p>
        <strong>Lembre-se, metas precisam ser específicas, mensuráveis, atingíveis, relevantes e temporais</strong>
       <p> <strong>Atenção:</strong> Revise-as antes da avaliação intermediária.</p>
        <ul>{metas_texto}</ul>
        """
    else:
        trecho_avaliacao = f"""
        <p>E já identificamos que você já realizou sua autoavaliação.</p>
        <p>Contudo, há <strong>ajustes necessários</strong> que precisam ser realizados nas metas abaixo.</p>
        <strong>Lembre-se, metas precisam ser específicas, mensuráveis, atingíveis, relevantes e temporais</strong>
        <ul>{metas_texto}</ul>
        """

    corpo_html = f"""
    <div style="font-family:Segoe UI; font-size:15px; line-height:1.0">

    <p>Olá, <strong>{colab}</strong>,</p>

    <p>Estamos na fase intermediária da <strong>Avaliação de Performance</strong>, uma etapa essencial para o seu desenvolvimento individual.</p>

    {trecho_avaliacao}

    <hr style="border: none; border-top: 1px solid #ccc; margin: 20px 0;">
    <p><strong>{gestor}</strong>, você está em cópia deste e-mail para acompanhar o processo citado acima.</p>
    <hr style="border: none; border-top: 1px solid #ccc; margin: 20px 0;">

    <p>Dúvidas, estou à disposição.</p>
    <p>Abraços,</p>
    <p><strong>Time de Recursos Humanos</strong></p>

    </div>
    """

    try:
        mail = outlook.CreateItem(0)
        mail.To = email_colab
        mail.CC = email_gestor
        mail.Subject = "[ATENÇÃO] Correção de metas – Prazo até 30/07"
        mail.HTMLBody = corpo_html
        mail.Display()  # Apenas abre o e-mail para revisão antes do envio
        print(f"📨 E-mail preparado para {colab} ({email_colab})")
        emails_gerados += 1
    except Exception as e:
        print(f"❌ Erro ao montar e-mail para {colab}: {e}")

print(f"\n✅ Processo finalizado. Total de e-mails enviados: {emails_gerados}")
