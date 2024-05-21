import warnings
import os
import glob
import pandas as pd
import numpy as np

warnings.filterwarnings("ignore", category=FutureWarning)

entrada_pasta = r"C:\Users\josuesilva\Documents\2. Documentos originais\DEFENSAS\ANEXO I\BR-163"
saida = r"C:\Users\josuesilva\Desktop\analise_barreiras.xlsx"

# Funções definidas anteriormente
def concat_with_e(x):
    return '&'.join(x.fillna('').astype(str))

def process(x, y):
    x = x.split('&')
    y = y.split('&')
    index = y.index('X')
    return x[index].replace(':', '')

def fix_values(x):
    return x.replace('&', '')

# Listar todos os arquivos Excel na pasta
arquivos = glob.glob(os.path.join(entrada_pasta, "*.xls"))

# Inicializar uma lista para armazenar os DataFrames processados
df_list = []

# Processar cada arquivo Excel
for arquivo in arquivos:
    df = pd.read_excel(arquivo)

    # Tratamento de dados
    df = df.drop([0, 1, 2])
    df = df[1:]
    df = df.iloc[:, 1:]
    df = df.drop(df.columns[[9, 10]], axis=1)

    df.reset_index(drop=True, inplace=True)
    df.columns = range(len(df.columns))

    df.columns = df.iloc[0]
    df = df[1:]
    df.reset_index(drop=True, inplace=True)

    novo_nome = 'vazio'
    df = df.rename(columns={np.nan: novo_nome})

    # Agrupa os indices em elementos com 3 linhas cada
    df['group'] = np.floor(df.index / 3)

    # Coloca um & em atributos nulos
    result = df.groupby('group').agg(concat_with_e).reset_index(drop=True)
    result = result.drop(columns=['group'], errors='ignore')

    positions = []

    # Identifica as colunas que possuem os atributos com múltipla escolha
    for j, colum_name in enumerate(result.columns):
        if 'vazio' in colum_name:
            positions.append(j - 1)

    drop = []
    for pos in positions:
        result.iloc[:, pos] = result.apply(lambda row: process(row[pos], row[pos+1]), axis=1)
        drop.append(pos+1)

    result = result.drop(df.columns[drop], axis=1)

    colunas = result.columns
    for index in colunas:
        result[index] = result.apply(lambda row: fix_values(row[index]), axis=1)

    # Adiciona o DataFrame processado à lista
    df_list.append(result)

# Concatenar todos os DataFrames em um único DataFrame
df_final = pd.concat(df_list, ignore_index=True)

# Salvar o DataFrame final em um arquivo Excel
df_final.to_excel(saida, index=False)

print("\nAnálise concluída com sucesso!")
