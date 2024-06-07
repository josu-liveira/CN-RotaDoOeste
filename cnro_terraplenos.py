import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import time
import threading


# Variável global para armazenar os arquivos selecionados
arquivos_selecionados = []

def selecionar_arquivos():
    global arquivos_selecionados
    arquivos = filedialog.askopenfilenames(
        title="Selecione os arquivos Excel",
        filetypes=[("Arquivos Excel", ".xlsx"), ("All files", ".*")]
    )
    if arquivos:
        contador_arquivos.set(f"{len(arquivos)} arquivo(s) importado(s)")
        arquivos_selecionados = list(arquivos)

def selecionar_pasta():
    global arquivos_selecionados
    pasta = filedialog.askdirectory(title="Selecione a pasta contendo os arquivos Excel")
    if pasta:
        arquivos = [os.path.join(pasta, arquivo) for arquivo in os.listdir(pasta) if arquivo.endswith('.xlsx')]
        contador_arquivos.set(f"{len(arquivos)} arquivo(s) importado(s)")
        arquivos_selecionados = arquivos

def selecionar_diretorio():
    diretorio_destino = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx"), ("All files", "*")],
        title="Selecione o diretório de destino",
    )
    if diretorio_destino:
        sucesso = exportar_analise(diretorio_destino)
        if sucesso:
            messagebox.showinfo("Sucesso", f'Arquivo salvo em "{diretorio_destino}".')
        else:
            messagebox.showwarning("Exportação Falhou", "Você precisa realizar uma análise antes de exportar os resultados.")

def calcular_eta(tempo_inicial, idx, total_arquivos):
    tempo_atual = time.time()
    tempo_passado = tempo_atual - tempo_inicial
    tempo_medio_por_arquivo = tempo_passado / idx if idx > 0 else 0
    arquivos_rest = total_arquivos - idx
    eta_segundos = tempo_medio_por_arquivo * arquivos_rest

    horas = int(eta_segundos / 3600)
    minutos = int((eta_segundos % 3600) / 60)
    segundos = int(eta_segundos % 60)
    return f"{horas:02d}:{minutos:02d}:{segundos:02d}"

def analisar_arquivos():
    global arquivos_selecionados
    if arquivos_selecionados:
        dados = []
        total_arquivos = len(arquivos_selecionados)
        progresso['maximum'] = total_arquivos
        tempo_inicial = time.time()

        for idx, arquivo in enumerate(arquivos_selecionados, start=1):
            progresso.set(idx / total_arquivos)
            root.update_idletasks()
            try:
                nome_arquivo = os.path.basename(arquivo)
                lbl_arquivo_atual.configure(text=f"Processando: {nome_arquivo}")
                df = pd.read_excel(arquivo)

                desc_dist_acst = ', '.join([str(val) for val in df.iloc[10, 2:].values.flatten() if pd.notna(val)]) if not df.iloc[10, 2:].isnull().all().all() else ''
                obs_gerais = ', '.join([str(val) for val in df.iloc[35:38, 0:].values.flatten() if pd.notna(val)]) if not df.iloc[35:38, 0:].isnull().all().all() else ''

                data = {
                    # DADOS GERAIS
                    "ID": df.iloc[4, 1],
                    "TIPO": df.iloc[3, 1],
                    "LOCAL": df.iloc[3, 4],
                    "TRECHO": df.iloc[4, 4],
                    "RODOVIA": df.iloc[3, 6],
                    "DT_INSP": df.iloc[4, 6],
                    
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

                combined_data = {**data}
                dados.append(combined_data)

                # Calcular e exibir ETA
                eta = calcular_eta(tempo_inicial, idx, total_arquivos)
                lbl_eta.configure(text=f"Tempo estimado: {eta}")
            except Exception as e:
                print(f"Erro ao processar o arquivo {arquivo}: {e}")

        global df_final
        df_final = pd.DataFrame(dados)
        lbl_arquivo_atual.configure(text="")
        progresso.set(0)
        messagebox.showinfo("Sucesso", "Análise concluída com sucesso!\n\nExporte para seu diretório.")

def iniciar_analise():
    thread = threading.Thread(target=analisar_arquivos)
    thread.start()

def exportar_analise(diretorio_destino):
    if 'df_final' in globals():
        try:
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(diretorio_destino), exist_ok=True)
            df_final.to_excel(diretorio_destino, index=False)
            return True
        except Exception as e:
            print(e)
            return False
    else:
        return False

def mostrar_sobre():
    messagebox.showinfo("Sobre", "Versão v1.4\n\nBy Josué\n\nSe houver dúvidas, não exite em me mandar um Teams")

def abrir_ajuda():
    import webbrowser
    webbrowser.open("https://github.com/josu-liveira/cnro-fichas")

# Criar a janela principal
ctk.set_appearance_mode("dark")  # Modo de aparência
ctk.set_default_color_theme("blue")  # Tema de cores

root = ctk.CTk()
root.title("Linear terraplenos - Monitoração 2024")
root.geometry("755x600")

# Frame principal
frame_principal = ctk.CTkFrame(root)
frame_principal.pack(pady=20, padx=20, fill="both", expand=True)

# Frame de importação
frame_importacao = ctk.CTkFrame(frame_principal)
frame_importacao.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

lbl_importacao = ctk.CTkLabel(frame_importacao, text="Importação", font=("Arial", 16, "bold"))
lbl_importacao.pack(pady=10)

btn_selecionar = ctk.CTkButton(frame_importacao, text="Arquivo único", width=200, height=40, command=selecionar_arquivos)
btn_selecionar.pack(pady=10)

btn_selecionar_pasta = ctk.CTkButton(frame_importacao, text="Importar pasta", width=200, height=40, command=selecionar_pasta)
btn_selecionar_pasta.pack(pady=10)

# Frame de análise
frame_analise = ctk.CTkFrame(frame_principal)
frame_analise.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

lbl_analise = ctk.CTkLabel(frame_analise, text="Análise", font=("Arial", 16, "bold"))
lbl_analise.pack(pady=10)

btn_analisar = ctk.CTkButton(frame_analise, text="Iniciar análise", width=200, height=40, command=iniciar_analise)
btn_analisar.pack(pady=10, padx=16)

# Frame de exportação
frame_exportacao = ctk.CTkFrame(frame_principal)
frame_exportacao.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

lbl_exportacao = ctk.CTkLabel(frame_exportacao, text="Exportação", font=("Arial", 16, "bold"))
lbl_exportacao.pack(pady=10)

btn_exportar = ctk.CTkButton(frame_exportacao, text="Exportar resultados", width=200, height=40, command=selecionar_diretorio)
btn_exportar.pack(pady=10)

# Contador de arquivos
frame_contador = ctk.CTkFrame(frame_principal)
frame_contador.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

contador_arquivos = ctk.StringVar()
contador_arquivos.set("Nenhum arquivo foi importado")
lbl_contador = ctk.CTkLabel(frame_contador, textvariable=contador_arquivos, font=("Arial", 14, "bold"), anchor='w')  
lbl_contador.pack(pady=40)

# Barra de progresso
frame_progresso = ctk.CTkFrame(frame_principal)
frame_progresso.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

lbl_progresso = ctk.CTkLabel(frame_progresso, text="Progresso", font=("Arial", 16, "bold"))
lbl_progresso.pack(pady=5)

progresso = ctk.CTkProgressBar(frame_progresso, width=300)
progresso.pack(pady=10, padx=70)

lbl_arquivo_atual = ctk.CTkLabel(frame_progresso, text="", font=("Arial", 12))
lbl_arquivo_atual.pack(pady=5)

# ETA
frame_eta = ctk.CTkFrame(frame_principal)
frame_eta.grid(row=2, column=1, padx=10, pady=10, sticky="nsew")

lbl_eta = ctk.CTkLabel(frame_eta, text="Tempo estimado: ", font=("Arial", 14, "bold"))
lbl_eta.pack(pady=40)

# Botões de Ajuda e Sobre
frame_info = ctk.CTkFrame(root)
frame_info.pack(pady=10, padx=10, fill="both", expand=True)

btn_ajuda = ctk.CTkButton(frame_info, text="Dúvidas?", width=100, height=30, command=abrir_ajuda)
btn_ajuda.pack(side="left", padx=20, pady=10)

btn_sobre = ctk.CTkButton(frame_info, text="Sobre", width=100, height=30, command=mostrar_sobre)
btn_sobre.pack(side="right", padx=20, pady=10)

# Executar o loop principal da interface
root.mainloop()