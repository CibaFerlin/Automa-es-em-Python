import pandas as pd

# Carregar o arquivo Excel
def processar_excel(arquivo_entrada, arquivo_saida):
    # Ler o arquivo Excel
    df = pd.read_excel(r'C:\Users\cferlin\Downloads\Metas_status_fechamento24.xlsx')

    # Certifique-se de que as colunas necessárias existem
    colunas_necessarias = ['Avaliação de Desempenho:Colaborador', 'Descrição', 'Status']
    for coluna in colunas_necessarias:
        if coluna not in df.columns:
            raise ValueError(f"A coluna {coluna} não foi encontrada no arquivo Excel.")

    # Concatenar 'Descrição' e 'Status' e numerar
    df['Descrição_Status'] = df['Descrição'] + '\n'' (' + df['Status'] + ')'

    # Agrupar por 'Avaliação de Desempenho:Colaborador' e concatenar com numeração
    def concatenar_metainformacoes(grupo):
        print(grupo)
        resultado_meta='\n\n'.join(f"{i+1}. {meta}" for i, meta in enumerate(grupo))
        return resultado_meta

    resultado = df.groupby('Avaliação de Desempenho:Colaborador')['Descrição_Status'].apply(concatenar_metainformacoes).reset_index()

    # Renomear as colunas para algo mais significativo
    resultado.columns = ['Colaborador', 'Metas Concatenadas']
    # Salvar o resultado em um novo arquivo Excel
    resultado.to_excel(arquivo_saida, index=False)
    print(f"Arquivo processado e salvo como {arquivo_saida}")

# Exemplo de uso
# Substitua 'entrada.xlsx' pelo caminho do seu arquivo de entrada e 'saida.xlsx' pelo caminho do arquivo de saída desejado
processar_excel(r'C:/Users/cferlin/Downloads/Metas_v15012025.xlsx', 'metas_concatenadas_fechamento24.xlsx')
