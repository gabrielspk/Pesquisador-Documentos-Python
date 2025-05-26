import chardet
from tkinter import messagebox
import tkinter as tk
import os


#função responsável na identificação do enconding
def detectar_encoding(file_path, tamanho=1000):
    with open(file_path, 'rb') as f:
        rawdata = f.read(tamanho)
    resultado = chardet.detect(rawdata)
    return resultado['encoding'] or 'utf-8'  #fallback se não detectar

#função responsável por pegar os documentos do textbox para realizar a pesquisa dos documentos
def carregar_documentos(documentos_textbox): 
    if not documentos_textbox.get("1.0", tk.END).strip():
        messagebox.showerror("Erro", "Por favor, insira os documentos no campo de texto.")  
        return
    
    documentos = documentos_textbox.get("1.0", tk.END).strip().split("\n") #declarando a variável documentos e pegando com quebra de linha os mesmos e armazenando
    
    return [documento.strip() for documento in documentos if documento.strip()] #retornando uma lista de documentos já ordenados

def incrementar_nome_arquivo(nome_arquivo): #função responsável por iterar no salvamento de um novo arquivo para diferentes relatórios
    count = 1
    while os.path.exists(nome_arquivo): #enquanto a nomenclatura atual existir, será adicionado uma numeração única
        nome_arquivo = f"RelatorioProcessamento({count}).xlsx"
        count += 1
    return nome_arquivo

def limpar_textbox(log_textbox):
    log_textbox.config(state='normal')
    log_textbox.delete('1.0', 'end')
    log_textbox.config(state='disabled')

def inserir_log(log_textbox, mensagem):
    log_textbox.config(state='normal')
    log_textbox.insert(tk.END, mensagem)
    log_textbox.see(tk.END)
    log_textbox.config(state='disabled')