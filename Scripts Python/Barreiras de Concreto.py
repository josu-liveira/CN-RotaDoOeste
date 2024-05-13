import pandas as pd
import numpy as np


diretorio = (r'C:\Users\josuesilva\Documents\2. Documentos originais\DEFENSAS\ANEXO I\BR-163\BR-163_Barreira de concreto.xls')

df = pd.read_excel(diretorio)


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

# Inicio das funções
def concat_with_e(x):
    return '&'.join(x.fillna('').astype(str))


def process(x, y):

    x = x.split('&')

    y = y.split('&')

    #print(x, y)

    index = y.index('X')

    return x[index].replace(':','')


def fix_values(x):

    return x.replace('&','')

# Agrupa os indices em elementos com 3 linhas cada
df['group'] = np.floor(df.index / 3)


# Coloca um & em atributos nulos
result = df.groupby('group').agg(concat_with_e).reset_index(drop=True)

result = result.drop(columns=['group'], errors='ignore')

positions = []

# identifica as colunas que possuem os atributos com multipla escolha
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

#display(result)

nome_do_arquivo = 'teste_novão.xlsx'  # Nome do arquivo
result.to_excel(nome_do_arquivo, index=False)

print("\nANÁLISE CONCLUÍDA COM SUCESSO!")