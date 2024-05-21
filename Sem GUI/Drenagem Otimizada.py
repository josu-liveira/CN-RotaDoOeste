import os
import glob
import pandas as pd

entrada_pasta = r"C:\Users\josuesilva\Documents\2. Documentos originais\Fichas Monitoração de Drenagem"
saida = r"C:\Users\josuesilva\Desktop\drenagem_analise.xlsx"

# Listar todos os arquivos Excel na pasta
arquivos = glob.glob(os.path.join(entrada_pasta, "*.xlsx"))

# Inicializar uma lista para armazenar os DataFrames processados
df_list = []

# Processar cada arquivo Excel
for arquivo in arquivos:
    df = pd.read_excel(arquivo)
    
    data = {
        "ID": df.iloc[4, 2],
        "EXT": df.iloc[5, 2],
        "KM_INI": df.iloc[4, 7],
        "KM_FIM": df.iloc[5, 7],
        "TIPO": df.iloc[9, 2],
        "FORMA": df.iloc[9, 7],
        "LAT_MONT": df.iloc[7, 2],
        "LONG_MONT": df.iloc[8, 2],
        "DIM_MONT": df.iloc[11, 2],
        "LAD_MON": df.iloc[12, 2],
        "EST_MONT": df.iloc[13, 2],
        "MAT_MONT": df.iloc[14, 2],
        "CONSERVA_MONT": df.iloc[15, 2],
        "OK_MONT": df.iloc[17, 3],
        "LIMP_MONT": df.iloc[18, 3],
        "ASSO_MONT": df.iloc[19, 3],
        "AFOG_MONT": df.iloc[20, 3],
        "OBS_MONT": df.iloc[21, 2],
        "TESTA_DAN_MONT": df.iloc[17, 8],
        "TUB_MONT": df.iloc[18, 8],
        "CX_MONT": df.iloc[19, 8],
        "EROSAO_MONT": df.iloc[20, 8],
        "TRINCA_MONT": df.iloc[21, 8],
        "TAMPA_DAN_MONT": df.iloc[22, 8],
        "LAT_JUS": df.iloc[7, 7],
        "LONG_JUS": df.iloc[8, 7],
        "DIM_JUS": df.iloc[11, 7],
        "LAD_JUS": df.iloc[12, 7],
        "EST_JUS": df.iloc[13, 7],
        "MAT_JUS": df.iloc[14, 7],
        "CONSERVA_JUS": df.iloc[15, 7],
        "OK_JUS": df.iloc[17, 4],
        "LIMP_JUS": df.iloc[18, 4],
        "ASSO_JUS": df.iloc[19, 4],
        "AFOG_JUS": df.iloc[20, 4],
        "OBS_JUS": df.iloc[21, 2],
        "TESTA_DAN_JUS": df.iloc[17, 9],
        "TUB_JUS": df.iloc[18, 9],
        "CX_JUS": df.iloc[19, 9],
        "EROSAO_JUS": df.iloc[20, 9],
        "TRINCA_JUS": df.iloc[21, 9],
        "TAMPA_DAN_JUS": df.iloc[22, 9],
    }
    
    # Converter o dicionário em um DataFrame e adicioná-lo à lista
    data_df = pd.DataFrame([data])
    df_list.append(data_df)

# Concatenar todos os DataFrames em um único DataFrame
df_final = pd.concat(df_list, ignore_index=True)

# Salvar o DataFrame final em um arquivo Excel
df_final.to_excel(saida, index=False)

print("\nAnálise concluída com sucesso!")