# import pandas as pd
# from faker import Faker

# def substituir_nomes_por_falsos(df, coluna_nomes):
#     """Função para substituir nomes por nomes falsos, garantindo consistência e armazenando o mapeamento."""
#     fake = Faker()  # Inicializa o gerador de nomes falsos
#     mapeamento_nomes = {}  # Dicionário para armazenar o mapeamento original -> falso

#     def gerar_nome_falso(nome_real):
#         """Gera um nome falso e garante que o mesmo nome real tenha o mesmo nome falso."""
#         if nome_real not in mapeamento_nomes:
#             mapeamento_nomes[nome_real] = fake.name()  # Cria um nome falso se ainda não existir no mapeamento
#         return mapeamento_nomes[nome_real]

#     # Substitui os nomes reais na coluna pela função gerar_nome_falso
#     df[coluna_nomes] = df[coluna_nomes].apply(lambda x: gerar_nome_falso(x) if pd.notnull(x) else x)
    
#     # Retorna o DataFrame atualizado e o mapeamento de nomes
#     return df, mapeamento_nomes

# df = pd.read_excel('prev.xlsx')

# # Exemplo de uso:
# df, mapeamento = substituir_nomes_por_falsos(df, 'Relator')
# print(mapeamento)  # Exibe o mapeamento de nomes reais para falsos
# df.to_excel('anonimizado.xlsx', index=False)


# import pandas as pd
# from faker import Faker

# def substituir_nomes_por_falsos(df, colunas_nomes, trocar_cidade=False):
#     """Função para substituir nomes por nomes falsos em várias colunas, garantindo consistência e armazenando o mapeamento."""
#     fake = Faker()  # Inicializa o gerador de nomes falsos
#     mapeamento_nomes = {}  # Dicionário para armazenar o mapeamento original -> falso

#     def gerar_nome_falso(nome_real):
#         """Gera um nome falso e garante que o mesmo nome real tenha o mesmo nome falso."""
#         if nome_real not in mapeamento_nomes:
#             mapeamento_nomes[nome_real] = fake.name()  # Cria um nome falso se ainda não existir no mapeamento
#         return mapeamento_nomes[nome_real]

#     def gerar_nome_cidade_falsa(cidade_real):
#         """Gera uma cidade falsa e garante que a mesma cidade tenha o mesmo nome falso."""
#         if cidade_real not in mapeamento_nomes:
#             mapeamento_nomes[cidade_real] = fake.city()  # Gera um nome de cidade falso
#         return mapeamento_nomes[cidade_real]

#     # Substitui os nomes reais nas colunas especificadas
#     for coluna in colunas_nomes:
#         if coluna in df.columns:
#             if trocar_cidade and coluna == "Unidade":  # Se for a coluna de cidade, troca com nome de cidade
#                 df[coluna] = df[coluna].apply(lambda x: gerar_nome_cidade_falsa(x) if pd.notnull(x) else x)
#             else:  # Para colunas normais (não cidades)
#                 df[coluna] = df[coluna].apply(lambda x: gerar_nome_falso(x) if pd.notnull(x) else x)
    
#     # Retorna o DataFrame atualizado e o mapeamento de nomes
#     return df, mapeamento_nomes

# def alterar_arquivo_excel(caminho_arquivo, colunas_nomes, caminho_novo_arquivo, trocar_cidade=False):
#     """Função para alterar múltiplas colunas em todas as abas de um arquivo Excel, preservando o formato de datas."""
#     # Carrega o arquivo Excel com todas as abas
#     xls = pd.ExcelFile(caminho_arquivo)
#     escritor = pd.ExcelWriter(caminho_novo_arquivo, engine='xlsxwriter')

#     mapeamentos_completos = {}  # Armazena o mapeamento de todas as abas
    
#     # Itera sobre cada aba (sheet)
#     for nome_aba in xls.sheet_names:
#         df = pd.read_excel(xls, sheet_name=nome_aba, parse_dates=True)  # Força pandas a interpretar datas corretamente
        
#         # Preserva o formato de datas removendo a parte do horário se não for necessária
#         for coluna in df.select_dtypes(include=['datetime']):
#             df[coluna] = df[coluna].dt.strftime('%m/%d/%Y')  # Formato MM/DD/YYYY

#         # Substitui os nomes nas colunas especificadas, incluindo cidades se necessário
#         df, mapeamento = substituir_nomes_por_falsos(df, colunas_nomes, trocar_cidade)
#         mapeamentos_completos[nome_aba] = mapeamento
        
#         # Salva o DataFrame alterado na nova planilha Excel
#         df.to_excel(escritor, sheet_name=nome_aba, index=False)
    
#     # Salva o arquivo Excel com as alterações
#     escritor.save()
    
#     print(f"Arquivo salvo como {caminho_novo_arquivo}")
#     return mapeamentos_completos

# # Exemplo de uso:
# caminho_arquivo = 'prev.xlsx'  # Caminho do arquivo original
# colunas_nomes = ['Relator', 'Observador', 'Responsável', 'Responsável2', 'Unidade']  # Inclua "Unidade" na lista de colunas
# caminho_novo_arquivo = 'prev_anonimizado.xlsx'  # Caminho do novo arquivo

# # Chama a função para alterar o arquivo Excel, incluindo a troca da coluna "Unidade" como cidade
# mapeamentos = alterar_arquivo_excel(caminho_arquivo, colunas_nomes, caminho_novo_arquivo, trocar_cidade=True)

import pandas as pd
from faker import Faker

def substituir_nomes_por_falsos(df, colunas_nomes, trocar_cidade=False):
    """Função para substituir nomes por nomes falsos em várias colunas, garantindo consistência e armazenando o mapeamento."""
    fake = Faker()  # Inicializa o gerador de nomes falsos
    mapeamento_nomes = {}  # Dicionário para armazenar o mapeamento original -> falso

    def gerar_nome_falso(nome_real):
        """Gera um nome falso e garante que o mesmo nome real tenha o mesmo nome falso."""
        if nome_real not in mapeamento_nomes:
            mapeamento_nomes[nome_real] = fake.name()  # Cria um nome falso se ainda não existir no mapeamento
        return mapeamento_nomes[nome_real]

    def gerar_nome_cidade_falsa(cidade_real):
        """Gera uma cidade falsa e garante que a mesma cidade tenha o mesmo nome falso."""
        if cidade_real not in mapeamento_nomes:
            mapeamento_nomes[cidade_real] = fake.city()  # Gera um nome de cidade falso
        return mapeamento_nomes[cidade_real]

    # Substitui os nomes reais nas colunas especificadas
    for coluna in colunas_nomes:
        if coluna in df.columns:
            if trocar_cidade and coluna == "Unidade":  # Se for a coluna de cidade, troca com nome de cidade
                df[coluna] = df[coluna].apply(lambda x: gerar_nome_cidade_falsa(x) if pd.notnull(x) else x)
            else:  # Para colunas normais (não cidades)
                df[coluna] = df[coluna].apply(lambda x: gerar_nome_falso(x) if pd.notnull(x) else x)
    
    # Retorna o DataFrame atualizado e o mapeamento de nomes
    return df, mapeamento_nomes

def alterar_arquivo_excel(caminho_arquivo, colunas_nomes, caminho_novo_arquivo, trocar_cidade=False):
    """Função para alterar múltiplas colunas em todas as abas de um arquivo Excel, preservando o formato de datas."""
    # Carrega o arquivo Excel com todas as abas
    xls = pd.ExcelFile(caminho_arquivo)
    escritor = pd.ExcelWriter(caminho_novo_arquivo, engine='xlsxwriter')

    mapeamentos_completos = {}  # Armazena o mapeamento de todas as abas
    
    # Itera sobre cada aba (sheet)
    for nome_aba in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=nome_aba, parse_dates=True)  # Força pandas a interpretar datas corretamente
        
        # Preserva o formato de datas removendo a parte do horário se não for necessária
        for coluna in df.select_dtypes(include=['datetime']):
            df[coluna] = df[coluna].dt.strftime('%m/%d/%Y')  # Formato MM/DD/YYYY

        # Substitui os nomes nas colunas especificadas, incluindo cidades se necessário
        df, mapeamento = substituir_nomes_por_falsos(df, colunas_nomes, trocar_cidade)
        mapeamentos_completos[nome_aba] = mapeamento
        
        # Salva o DataFrame alterado na nova planilha Excel
        df.to_excel(escritor, sheet_name=nome_aba, index=False)
    
    # Salva o arquivo Excel com as alterações
    escritor.close()  # Corrige o erro de save()
    
    print(f"Arquivo salvo como {caminho_novo_arquivo}")
    return mapeamentos_completos

# Exemplo de uso:
caminho_arquivo = 'prev.xlsx'  # Caminho do arquivo original
colunas_nomes = ['Relator', 'Observador', 'Responsável', 'Responsável2', 'Unidade']  # Lista de colunas que precisam ser alteradas
caminho_novo_arquivo = 'prev_anonimizado.xlsx'  # Caminho do novo arquivo

# Chama a função para alterar o arquivo Excel, incluindo a troca da coluna "Unidade" como cidade
mapeamentos = alterar_arquivo_excel(caminho_arquivo, colunas_nomes, caminho_novo_arquivo, trocar_cidade=True)
