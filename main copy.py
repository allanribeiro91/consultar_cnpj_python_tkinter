import time
import tkinter as tk
import TKinterModernThemes as TKMT
from tkinter import filedialog, messagebox, ttk
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
        previsao_encerramento = data_inicio + timedelta(seconds=num_cnpjs * 10)

        messagebox.showinfo("Informação",
                            f"Nº de CNPJs do arquivo: {num_cnpjs}\n"
                            f"Data de início: {data_inicio.strftime('%d/%m/%Y %H:%M:%S')}\n"
                            f"Previsão de encerramento: {previsao_encerramento.strftime('%d/%m/%Y %H:%M:%S')}")

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
            time.sleep(10)

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

# Configuração da interface Tkinter
root = tk.Tk()
root.title("Consulta de CNPJ")
root.geometry("600x450")

# Configuração dos widgets
label_titulo = tk.Label(root, text="CONSULTA DE CNPJ", font=("Consolas", 16))
label_cnpj = tk.Label(root, text="CNPJ:", font=("Calibri", 12))
entry_cnpj = tk.Entry(root, font=("Helvetica", 12))
button_individual = tk.Button(root, text="Consultar", command=consulta_individual, bg="green", fg="white", font=("Helvetica", 12))
button_lote = tk.Button(root, text="Consultar em Lote", command=consulta_lote, bg="green", fg="white", font=("Helvetica", 12))
label_resultado = tk.Label(root, text="Resultado", font=("Helvetica", 12))
result_text = tk.Text(root, height=15, width=70, font=("Helvetica", 10))
progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')

# Posicionamento dos widgets
label_titulo.grid(row=0, column=0, columnspan=3, pady=10)
label_cnpj.grid(row=1, column=0, sticky=tk.E, padx=10)
entry_cnpj.grid(row=1, column=1, padx=10)
button_individual.grid(row=1, column=2, padx=10)
button_lote.grid(row=2, column=1, pady=10)
label_resultado.grid(row=3, column=0, columnspan=3)
result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10)
progress.grid(row=5, column=0, columnspan=3, padx=10, pady=10, sticky=tk.W+tk.E)

root.mainloop()
