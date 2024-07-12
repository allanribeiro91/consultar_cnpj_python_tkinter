import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkthemes import ThemedTk
from consultar_cnpj import consultar_cnpj, processar_consulta, extrair_dados, extrair_cnaes_secundarios
import pandas as pd
from datetime import datetime, timedelta
import json

def validar_planilha(file_path):
    try:
        df = pd.read_excel(file_path)
        if "CNPJ" not in df.columns:
            messagebox.showerror("Erro", "A planilha não contém a coluna 'CNPJ'.")
            return None
        return df
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        return None

def consulta_lote():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = validar_planilha(file_path)
        if df is None:
            return
        
        num_cnpjs = len(df)
        data_inicio = datetime.now()
        previsao_encerramento = data_inicio + timedelta(seconds=num_cnpjs * 5)
        
        # Calcular o tempo total em horas, minutos e segundos
        tempo_total = previsao_encerramento - data_inicio
        horas, resto = divmod(tempo_total.total_seconds(), 3600)
        minutos, segundos = divmod(resto, 60)

        messagebox.showinfo("Informação",
                            f"Nº de CNPJs do arquivo: {num_cnpjs}\n"
                            f"Data de início: {data_inicio.strftime('%d/%m/%Y %H:%M:%S')}\n"
                            f"Previsão de encerramento: {previsao_encerramento.strftime('%d/%m/%Y %H:%M:%S')}\n"
                            f"Tempo total para processar as consultas: {int(horas):02d}:{int(minutos):02d}:{int(segundos):02d}")

        progress['maximum'] = num_cnpjs

        resultados = []
        cnaes_secundarios = []

        for index, row in df.iterrows():
            cnpj = row["CNPJ"]
            resultado = consultar_cnpj(cnpj)
            dados = extrair_dados(resultado)
            cnaes = extrair_cnaes_secundarios(resultado)

            row_result = {
                "CNPJ": cnpj,
                "Data Processamento": datetime.now(),
                "Resultado": processar_consulta(resultado),
                **dados
            }
            resultados.append(row_result)

            for cnae in cnaes:
                cnaes_secundarios.append({
                    "CNPJ": cnpj,
                    "CNAE Secundário - Código": cnae["codigo"],
                    "CNAE Secundário - Nome": cnae["nome"]
                })
            
            progress['value'] += 1
            root.update_idletasks()
            time.sleep(5)

        df_resultados = pd.DataFrame(resultados)
        df_cnaes_secundarios = pd.DataFrame(cnaes_secundarios)

        writer = pd.ExcelWriter(file_path.replace(".xlsx", "_resultado.xlsx"), engine='xlsxwriter')
        df_resultados.to_excel(writer, sheet_name='Resultados', index=False)
        df_cnaes_secundarios.to_excel(writer, sheet_name='CNAEs Secundários', index=False)
        writer.close()

        messagebox.showinfo("Sucesso", "Consulta concluída. Resultados salvos no arquivo.")
        progress['value'] = 0
    else:
        messagebox.showwarning("Aviso", "Por favor, selecione um arquivo.")

def exibir_resultado(result):
    result_text.delete(1.0, tk.END)
    if result:
        result_text.insert(tk.END, json.dumps(result, indent=4))
    else:
        result_text.insert(tk.END, "Erro na consulta")

def consulta_individual():
    cnpj = entry_cnpj.get()
    if cnpj:
        result = consultar_cnpj(cnpj)
        exibir_resultado(result)
    else:
        messagebox.showwarning("Aviso", "Por favor, insira um CNPJ.")
# Configuração da interface Tkinter com ttkthemes
root = ThemedTk(theme="arc")  # Pode escolher outros temas como "breeze", "radiance", "yaru", etc.

# Definir o ícone da janela
# root.iconbitmap('caminho_para_seu_icone.ico')

root.title("Consulta de CNPJ")
root.geometry("540x460")

# Criação dos frames para organizar a interface
frame_top = ttk.Frame(root)
frame_top.grid(row=0, column=0, columnspan=3, pady=10, padx=10, sticky="ew")

frame_middle = ttk.Frame(root)
frame_middle.grid(row=1, column=0, columnspan=3, pady=10, padx=10, sticky="ew")

frame_bottom = ttk.Frame(root)
frame_bottom.grid(row=2, column=0, columnspan=3, pady=10, padx=10, sticky="ew")

# Configuração dos widgets
label_titulo = ttk.Label(frame_top, text="CONSULTA DE CNPJ", font=("Consolas", 16))
label_titulo.pack()

label_cnpj = ttk.Label(frame_middle, text="CNPJ:", font=("Calibri", 12))
label_cnpj.grid(row=0, column=0, sticky=tk.E, padx=5)
entry_cnpj = ttk.Entry(frame_middle, font=("Calibri", 12))
entry_cnpj.grid(row=0, column=1, padx=5)
button_individual = ttk.Button(frame_middle, text="Consultar", command=consulta_individual)
button_individual.grid(row=0, column=2, padx=5)

label_resultado = ttk.Label(frame_bottom, text="Resultado", font=("Calibri", 12))
label_resultado.pack()
result_text = tk.Text(frame_bottom, height=15, width=70, font=("Calibri", 10))
result_text.pack(padx=10, pady=10)
button_lote = ttk.Button(frame_bottom, text="Consultar em Lote", command=consulta_lote)
button_lote.pack(pady=10)

progress = ttk.Progressbar(frame_bottom, orient='horizontal', mode='determinate')
progress.pack(padx=10, pady=10, fill=tk.X)

root.mainloop()