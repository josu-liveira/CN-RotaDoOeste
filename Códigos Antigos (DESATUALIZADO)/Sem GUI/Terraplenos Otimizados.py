import os
import glob
import pandas as pd

entrada_pasta = r"C:\Users\josuesilva\Documents\2. Documentos originais\Fichas Terraplenos\EXCEL"
saida = r"C:\Users\josuesilva\Desktop\terraplenos_analise.xlsx"

# Listar todos os arquivos Excel na pasta
arquivos = glob.glob(os.path.join(entrada_pasta, "*.xlsx"))

# Inicializar uma lista para armazenar os DataFrames processados
df_list = []

# Processar cada arquivo Excel
for arquivo in arquivos:
    df = pd.read_excel(arquivo)

    # Iterar sobre os valores
    desc_dist_acst = ', '.join([str(val) for val in df.iloc[10, 2:].values.flatten() if pd.notna(val)]) if not df.iloc[10, 2:].isnull().all().all() else ''
    obs_gerais = ', '.join([str(val) for val in df.iloc[35:38, 0:].values.flatten() if pd.notna(val)]) if not df.iloc[35:38, 0:].isnull().all().all() else ''

    data = {
        # DADOS GERAIS
        "ID": df.iloc[4, 1],
        "TIPO": df.iloc[3, 1],
        "LOCAL": df.iloc[3, 4],
        "TRECHO": df.iloc[4, 4],
        "RODOVIA": df.iloc[3, 6],
        "DT_INSP": df.iloc[4, 6].strftime('%d/%m/%Y'),
        
        # CADASTRAMENTO
        "KM_INI": df.iloc[6, 1],
        "KM_FIN": df.iloc[7, 1],
        "SENT": df.iloc[6, 4],
        "CORD_INI": df.iloc[6, 6],
        "CORD_FIN": df.iloc[7, 6],
        
        # DADOS GEOMÉTRICOS DO TERRAPLENO
        "EXT": df.iloc[9, 1],
        "ALT_GEO": df.iloc[9, 4],
        "INCLIN": df.iloc[9, 6],
        "DIST_ACST": df.iloc[10, 1],
        "DESC_DIST_ACST": desc_dist_acst,
        
        # CARACTERÍSTICAS GERAIS
        "TIPO_TRPL": df.iloc[12, 1],
        "TIPO_RELEV": df.iloc[12, 6],
        "VGT": df.iloc[13, 1],
        "DENS_VGT": df.iloc[13, 6],

        # DADOS ESTRUTURA DE CONTENCAO
        "TIPO_CON_1": df.iloc[16, 1],
        "*******1": df.iloc[16, 2],
        "TIPO_CON_2": df.iloc[16, 3],
        "*******2": df.iloc[16, 4],
        "TIPO_CON_3": df.iloc[16, 5],
        "TIPO_CON_4": df.iloc[16, 6],
        "EXT_CON1": df.iloc[17, 1],
        "*******3": df.iloc[17, 2],
        "EXT_CON2": df.iloc[17, 3],
        "*******4": df.iloc[17, 4],
        "EXT_CON3": df.iloc[17, 5],
        "EXT_CON4": df.iloc[17, 6],
        "ALT_CON1": df.iloc[18, 1],
        "*******5": df.iloc[18, 2],
        "ALT_CON2": df.iloc[18, 3],
        "*******6": df.iloc[18, 4],
        "ALT_CON3": df.iloc[18, 5],
        "ALT_CON4": df.iloc[18, 6],
        "ANC_CON1": df.iloc[19, 1],
        "*******7": df.iloc[19, 2],
        "ANC_CON2": df.iloc[19, 3],
        "*******8": df.iloc[19, 4],
        "ANC_CON3": df.iloc[19, 5],
        "ANC_CON4": df.iloc[19, 6],
        "ELMT_CON1": df.iloc[20, 1],
        "*******9": df.iloc[20, 2],
        "ELMT_CON2": df.iloc[20, 3],
        "*******10": df.iloc[20, 4],
        "ELMT_CON3": df.iloc[20, 5],
        "ELMT_CON4": df.iloc[20, 6],

        # DRENAGEM
        "DRN_SUP": df.iloc[22, 1],
        "COND_DRN_SUP": df.iloc[22, 5],
        "DRN_SUB": df.iloc[23, 1],
        "TIPO_DRN": df.iloc[23, 4],
        "COND_DRN_SUB": df.iloc[23, 6],

        # CONDIÇÕES GERAIS DE SATURAÇÃO
        "PRES_AGUA": df.iloc[25, 1],
        
        # TIPOS DE INSTABILIZAÇÃO
        "TIPO_OCRR": df.iloc[27, 1],
        
        # DIAGNÓSTICO
        "CAUSAS_PROV": df.iloc[29, 1],
        "PASS_AMBI": df.iloc[30, 1],
        "DESC_PASS_AMBI": df.iloc[30, 2],
        
        # GRAVIDADE DA SITUAÇÃO
        "NVL_RSC": df.iloc[32, 5],
        "OUTR_ELMNT": df.iloc[33, 5],
        
        # OBSERVAÇÕES GERAIS
        "OBS_GERAIS": obs_gerais,
    }

    # Converter o dicionário para um DataFrame e adicioná-lo à lista
    data_df = pd.DataFrame([data])
    df_list.append(data_df)

# Concatenar todos os DataFrames em um único DataFrame
df_final = pd.concat(df_list, ignore_index=True)

# Salvar o DataFrame final em um arquivo Excel
df_final.to_excel(saida, index=False)

print("\nAnálise concluída com sucesso!")