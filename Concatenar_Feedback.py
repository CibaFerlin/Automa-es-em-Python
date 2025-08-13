import pandas as pd

# Carregar o arquivo Excel
df = pd.read_excel(r'C:\Users\cferlin\PycharmProjects\Avaliação Performance 2024\ComentáriosPAD.xlsx')

# Verificar se as colunas necessárias estão no DataFrame
if "Destinatário" in df.columns and "Comentário" in df.columns:
    # Agrupar pelos valores duplicados na coluna "Destinatário"
    df_concatenado = df.groupby("Destinatário", as_index=False).agg({
        "Comentário": lambda x: " | ".join(filter(pd.notnull, x))  # Concatenar os comentários
    })

    # Salvar o resultado em um novo arquivo Excel
    arquivo_saida = "resultado_concatenado.xlsx"
    df_concatenado.to_excel(arquivo_saida, index=False)
    print(f"Arquivo salvo como {arquivo_saida}")
else:
    print("As colunas 'Destinatário' e 'Comentários' não foram encontradas no arquivo Excel.")
