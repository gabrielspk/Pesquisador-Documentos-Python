import os
import glob
import openpyxl
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

#função responsável por manipular e carregar os documentos que serão carregados no textbox
def carregar_documentos(documentos_textbox):
    documentos = documentos_textbox.get("1.0", tk.END).strip().split("\n")
    return [documento.strip() for documento in documentos if documento.strip()] #retorna a lista de string formatada

#define a função pesquisar documentos, que vai iterar e percorrer cada arquivo do diretório selecionado
def pesquisar_documentos(diretorio_atual, extensoes_descartadas, validacao_nomenclatura, documentos):
    documentos_encontrados = {}
    documentos_nao_encontrados = set(documentos)
    qtde_arquivos_pesquisados = 0 #variável para contabilizar a quantidade de arquivos percorridos

    #Percorre cada arquivo do diretório que será selecionado
    for root, _, files in os.walk(diretorio_atual):
        for file in files:
            #valida se o arquivo não possui a extensão, e alguma sequência de caracteres que não deve ser incluída no laço
            if not file.endswith(extensoes_descartadas) and file not in validacao_nomenclatura:
                for documento in documentos_nao_encontrados.copy():
                    if documento in open(os.path.join(root, file), errors="ignore").read():
                        if documento not in documentos_encontrados:
                            documentos_encontrados[documento] = [] #inicia uma nova chave com o documento encontrado
                        documentos_encontrados[documento].append(os.path.join(file)) #atribui a nomenclatura do arquivo como um valor da chave documento.
                        documentos_nao_encontrados.remove(documento) #retira o documento da tupla de documentos não encontrados
                qtde_arquivos_pesquisados += 1

    return documentos_encontrados, documentos_nao_encontrados, qtde_arquivos_pesquisados

#define a função aonde será criado o relatório Excel dos documentos pesquisados
def criar_relatorio(documentos_encontrados, documentos_nao_encontrados, qtde_arquivos_pesquisados, nome_arquivo):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    #define o cabeçalho padrão do arquivo .xlsx
    sheet['A1'] = 'Documentos em arquivos'
    sheet['B1'] = 'Nomenclatura arquivo'
    sheet['C1'] = 'Não encontrado'

    #define a contagem das linhas de docs encontrados e não encontrados
    linha_docs_encontrados = 1
    linha_docs_nao_encontrados = 1

    #percorre cada valores no dicionário de documentos encontrados para adiciona-los ao relatório .xlsx
    for documento, arquivos in documentos_encontrados.items():
        linha_docs_encontrados += 1 #cada vez que passar no laço será iterado +1 na linha por exemplo A1,A2,A3
        sheet.cell(row=linha_docs_encontrados, column=1, value=documento)
        for i, arquivo in enumerate(arquivos, start=1):
            sheet.cell(row=linha_docs_encontrados, column=2, value=arquivo)

    #percorre cada documento não encontrado, para adiciona-los ao relatório .xlsx
    for documento in documentos_nao_encontrados:
        linha_docs_nao_encontrados += 1
        sheet.cell(row=linha_docs_nao_encontrados, column=3).value = documento

    #chama a função "incrementar_nome_arquivo" para iterar sobre cada arquivo com a nomenclatura igual
    nome_arquivo = incrementar_nome_arquivo(nome_arquivo)
    try:
        workbook.save(nome_arquivo)
        messagebox.showinfo("Relatório Gerado", f"Relatório gerado com sucesso: {nome_arquivo}") #traz mensagem de sucesso ao gerar o relatório
    except PermissionError:
        messagebox.showerror("Erro", "Não foi possível gerar o relatório, pois está aberto por outro programa.")#tratamento de exceção caso já esteja aberto

#função responsável por iterar caso já haja arquivos salvos no diretório com a mesma nomenclatura
def incrementar_nome_arquivo(nome_arquivo):
    count = 1
    while os.path.exists(nome_arquivo): #enquanto o nome do arquivo existir, será somado mais um ao final da contagem
        nome_arquivo = f"RelatorioProcessamento({count}).xlsx"
        count += 1
    return nome_arquivo

#define a função para o botão aonde será selecionado o diretório
def btn_selecionar_diretorio():
    diretorio = filedialog.askdirectory()
    diretorio_entry.delete(0, tk.END) #removendo o diretório anterior para a nova entrada
    diretorio_entry.insert(0, diretorio) #inserindo o novo diretório selecionado pelo usuário

#define a função do botão pesquisar aonde vai chamar as demais funções
def btn_pesquisar():
    diretorio_atual = diretorio_entry.get()
    if not diretorio_atual:
        messagebox.showerror("Erro", "Por favor, selecione um diretório.")
        return
    
    extensoes_descartadas = (".fpl", ".zip", ".ini", ".pdf")#define um valor fixo das extensões descartadas
    validacao_nomenclatura = glob.glob('Retorno')#define um valor fixo do tipo de sequencia de string a não ser validado na nomenclatura
    documentos = carregar_documentos(documentos_textbox)#pega os arquivos usando a função carregar_documentos dentro do textbox

    documentos_encontrados, documentos_nao_encontrados, qtde_arquivos_pesquisados = pesquisar_documentos(diretorio_atual, extensoes_descartadas, validacao_nomenclatura, documentos)

    #cria o relatório com base dos documentos encontrados, não eonctrados e a quantidade dos arquivos pesquisados
    criar_relatorio(documentos_encontrados, documentos_nao_encontrados, qtde_arquivos_pesquisados, "RelatorioProcessamento.xlsx")

    #insere todos os dados dentro da textbox de log
    tempo_execucao = time.time() - inicio
    log_textbox.insert(tk.END, f'Tempo de execução: {tempo_execucao:.4f} segundos\n')
    log_textbox.insert(tk.END, f'Documentos encontrados: {documentos_encontrados}\n')
    log_textbox.insert(tk.END, f'Documentos não encontrados: {documentos_nao_encontrados}\n')
    log_textbox.insert(tk.END, f'Quantidade de arquivos pesquisados: {qtde_arquivos_pesquisados}\n\n')

def limpar_log():
    log_textbox.delete('1.0', tk.END)

# Criar a janela
janela = tk.Tk()
janela.title("Pesquisa de Documentos")
janela.configure(bg="#E5E8E8")
janela.resizable(False,False)

# Criar os widgets
diretorio_frame = tk.Frame(janela)
diretorio_frame.pack(padx=10, pady=5)
diretorio_frame.configure(bg="#E5E8E8")


diretorio_label = tk.Label(diretorio_frame, font=("Helvetica Neue",12), text="Diretório:")
diretorio_label.grid(row=0, column=0, padx=5, pady=5)
diretorio_label.configure(bg="#E5E8E8")

diretorio_entry = tk.Entry(diretorio_frame, width=50)
diretorio_entry.grid(row=0, column=1, padx=5, pady=5)

selecionar_documentos_button = tk.Button(diretorio_frame, text="Selecionar", command=btn_selecionar_diretorio)
selecionar_documentos_button.grid(row=0, column=2, padx=5, pady=5)

documentos_frame = tk.Frame(janela)
documentos_frame.pack(padx=10, pady=5)
documentos_frame.configure(bg="#E5E8E8")

documentos_label = tk.Label(documentos_frame, font=("Helvetica",12), text="Documentos (um embaixo do outro)")
documentos_label.grid(row=0, column=0, padx=5, pady=5)
documentos_label.configure(bg="#E5E8E8")

documentos_textbox = scrolledtext.ScrolledText(documentos_frame, width=50, height=5)
documentos_textbox.grid(row=1, column=0, padx=5, pady=5)

pesquisar_button = tk.Button(janela, text="Pesquisar", command=btn_pesquisar)
pesquisar_button.pack(padx=10, pady=5)

log_frame = tk.Frame(janela)
log_frame.pack(padx=10, pady=5)
log_frame.configure(bg="#E5E8E8")

log_label = tk.Label(log_frame, font=("Helvetica Neue", 12), text="Log:")
log_label.grid(row=0, column=0, padx=5, pady=5)
log_label.configure(bg="#E5E8E8")

log_textbox = scrolledtext.ScrolledText(log_frame, width=50, height=10)
log_textbox.grid(row=1, column=0, padx=5, pady=5)
log_textbox.configure(bg="#F2F4F4")

limpar_button = tk.Button(log_frame, text="Limpar", command=limpar_log)
limpar_button.grid(row=2, column=0, padx=5, pady=5)

#Inicia o loop da interface gráfica
inicio = time.time()
janela.mainloop()