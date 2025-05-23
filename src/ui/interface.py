import tkinter as tk
from ttkbootstrap import Style
from tkinter import filedialog, messagebox, scrolledtext, ttk
import time
from core.extracao import busca_reversa
from core.pesquisa import executar_pesquisa_documentos
from core.analise import processar_validacao
from utils.utilitarios import carregar_documentos
from utils.utilitarios import limpar_log

# Variável global para armazenar o diretório selecionado
diretorio_selecionado = None

def iniciar_interface():
    global diretorio_selecionado

    janela = tk.Tk()
    janela.title("Pesquisa de Documentos")
    janela.geometry("600x600")
    janela.resizable(False, False)

    # Estilo ttkbootstrap
    style = Style(theme='flatly')

    # Notebook para abas
    notebook = ttk.Notebook(janela)
    notebook.pack(pady=10, expand=True, fill='both')

    ####### ABA 1 - Pesquisar Documentos #######
    aba1 = ttk.Frame(notebook)
    notebook.add(aba1, text='Pesquisar Documentos')

    # Frame para seleção do diretório
    diretorio_frame = tk.Frame(aba1)
    diretorio_frame.pack(pady=20)

    diretorio_label = tk.Label(diretorio_frame, text="Diretório:", font=("Montserrat", 8, 'bold'))
    diretorio_label.grid(row=0, column=0, padx=5)

    entry_diretorio_aba1 = tk.Entry(diretorio_frame, width=50)
    entry_diretorio_aba1.grid(row=0, column=1, padx=5)

    selecionar_diretorio_button = tk.Button(
        diretorio_frame,
        text="Selecionar",
        width=10,
        command=lambda: btn_selecionar_diretorio(entry_diretorio_aba1)
    )
    selecionar_diretorio_button.grid(row=0, column=2, padx=5)

    documentos_label = tk.Label(
        aba1,
        text="Insira os documentos (um embaixo do outro)",
        font=("Montserrat", 11, 'bold')
    )
    documentos_label.pack(pady=5)

    documentos_textbox = scrolledtext.ScrolledText(aba1, width=50, height=10)
    documentos_textbox.pack()

    var_fpl = tk.IntVar()
    checkBoxFpl = tk.Checkbutton(
        aba1,
        text='Pesquisar em arquivos .FPL',
        font=("Montserrat", 8, 'bold'),
        variable=var_fpl,
        onvalue=1,
        offvalue=0
    )
    checkBoxFpl.pack(pady=5)

    def btn_pesquisar():
        if not diretorio_selecionado:
            messagebox.showerror("Erro", "Por favor, selecione um diretório.")
            return

        documentos = carregar_documentos(documentos_textbox)
        validar_fpl = var_fpl.get() == 1

        executar_pesquisa_documentos(
            diretorio_atual=diretorio_selecionado,
            validar_fpl=validar_fpl,
            documentos=documentos,
            log_callback=lambda msg: log_textbox.insert(tk.END, msg)
        )

    pesquisar_button = tk.Button(
        aba1,
        text="Pesquisar",
        width=11,
        command=btn_pesquisar
    )
    pesquisar_button.pack(pady=5, padx=5)

    log_frame = tk.Frame(aba1)
    log_frame.pack(padx=5, pady=5)

    log_textbox = scrolledtext.ScrolledText(log_frame, width=60, height=10)
    log_textbox.pack()

    limpar_log_button = tk.Button(
        log_frame,
        text="Limpar Log",
        width=10,
        command=lambda: limpar_log(log_textbox)
    )
    limpar_log_button.pack(pady=5)


    ####### ABA 2 - Análise de relatórios #######
    aba2 = ttk.Frame(notebook)
    notebook.add(aba2, text='Análise de relatórios')

    # Configurar colunas para redimensionamento relativo
    aba2.grid_columnconfigure(0, weight=1)
    aba2.grid_columnconfigure(1, weight=3)
    aba2.grid_columnconfigure(2, weight=1)

    # Título centralizado
    nova_funcionalidade_label = tk.Label(
        aba2,
        text="Análise e validação de relatórios",
        font=("Montserrat", 12, 'bold')
    )
    nova_funcionalidade_label.grid(row=0, column=0, columnspan=3, pady=10, padx=3, sticky="nsew")

    label_descricao_aba2 = tk.Label(
        aba2,
        text=(
            "Selecione o diretório com os relatórios extraídos do ga, para que seja extraído os dados, "
            "validados e gerados planilhas ordenadas para análise e validação."
        ),
        wraplength=575,
        justify="left",
        font=("Roboto", 10),
        bg="#f0f0f0",
        fg="#333333"
    )
    label_descricao_aba2.grid(row=1, column=0, columnspan=3, pady=3, padx=2, sticky="n")

    label_diretorio_aba2 = tk.Label(aba2, text="Diretório:", font=("Montserrat", 8, 'bold'))
    label_diretorio_aba2.grid(row=2, column=0, padx=2, pady=10, sticky="e")

    entry_diretorio_aba2 = tk.Entry(aba2, width=50)
    entry_diretorio_aba2.grid(row=2, column=1, padx=5, pady=10, sticky="we")

    button_selecionar_diretorio_aba2 = tk.Button(
        aba2,
        text="Selecionar",
        width=10,
        command=lambda: btn_selecionar_diretorio(entry_diretorio_aba2)
    )
    button_selecionar_diretorio_aba2.grid(row=2, column=2, padx=5, pady=10, sticky="w")

    log_textbox2 = scrolledtext.ScrolledText(aba2, width=90, height=20)
    log_textbox2.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="n")

    def atualizar_log_validacao(mensagens):
        log_textbox2.delete('1.0', tk.END)
        for msg in mensagens:
            log_textbox2.insert(tk.END, msg + "\n")

    def btn_validar():
        if not diretorio_selecionado:
            messagebox.showerror("Erro", "Por favor, selecione um diretório.")
            return

        resultados, mensagens = processar_validacao(diretorio_selecionado)

        if resultados is None:
            atualizar_log_validacao(mensagens)
            return

        atualizar_log_validacao(mensagens)

    validar_button2 = tk.Button(
        aba2,
        text="Validar",
        width=11,
        command=btn_validar
    )
    validar_button2.grid(row=4, column=0, columnspan=3, pady=10)

    ####### ABA 3 - Pesquisa Reversa #######
    aba3 = ttk.Frame(notebook)
    notebook.add(aba3, text='Pesquisa Reversa')

    label_aba3 = tk.Label(aba3, text="Pesquisa Reversa", font=("Montserrat", 12, 'bold'))
    label_aba3.grid(row=0, column=0, columnspan=3, pady=10, padx=2, sticky="n")

    label_descricao_aba3 = tk.Label(
        aba3,
        text=(
            "Declare a posição dentro dos arquivos que deseja pesquisar e o comprimento da sequência "
            "de caracteres, para que seja pesquisado e extraído de arquivo a arquivo no seu diretório selecionado."
        ),
        wraplength=580,
        justify="left",
        font=("Roboto", 10),
        bg="#f0f0f0",
        fg="#333333"
    )
    label_descricao_aba3.grid(row=1, column=0, columnspan=3, pady=3, padx=15, sticky="n")

    label_diretorio_aba3 = tk.Label(aba3, text="Diretório:", font=("Montserrat", 8, 'bold'))
    label_diretorio_aba3.grid(row=2, column=0, padx=2, pady=10, sticky="e")

    entry_diretorio_aba3 = tk.Entry(aba3, width=50)
    entry_diretorio_aba3.grid(row=2, column=1, padx=2, pady=10, sticky="we")

    button_selecionar_diretorio_aba3 = tk.Button(
        aba3,
        text="Selecionar",
        width=10,
        command=lambda: btn_selecionar_diretorio(entry_diretorio_aba3)
    )
    button_selecionar_diretorio_aba3.grid(row=2, column=2, padx=5, pady=10, sticky="w")

    tk.Label(aba3, text="Posição:", font=("Montserrat", 8, 'bold')).grid(row=3, column=0, padx=5, pady=2, sticky="e")
    entry_posicao = tk.Entry(aba3, width=8)
    entry_posicao.grid(row=3, column=1, padx=2, pady=5, sticky="w")

    tk.Label(aba3, text="Comprimento:", font=("Montserrat", 8, 'bold')).grid(row=4, column=0, padx=5, pady=5, sticky="e")
    entry_comprimento = tk.Entry(aba3, width=8)
    entry_comprimento.grid(row=4, column=1, padx=2, pady=5, sticky="w")

    log_text = scrolledtext.ScrolledText(aba3, width=60, height=15)
    log_text.grid(row=7, column=0, columnspan=3, padx=10, pady=10, sticky="n")

    def btn_pesquisa_reversa():
        if not diretorio_selecionado:
            messagebox.showerror("Erro", "Por favor, selecione um diretório.")
            return

        EXTENSOES_IGNORADAS_PESQUISA = (".fpl", ".zip", ".ini", ".pdf", ".xlsx")

        posicao = entry_posicao.get()
        comprimento = entry_comprimento.get()

        log_text.delete(1.0, tk.END)  # Limpa o log
        documentosExtraidos = busca_reversa(diretorio_selecionado, posicao, comprimento, EXTENSOES_IGNORADAS_PESQUISA)

        for doc in documentosExtraidos:
            log_text.insert(tk.END, doc + "\n")

    pesquisa_reversa_button = tk.Button(
        aba3,
        text="Pesquisar",
        width=12,
        command=btn_pesquisa_reversa
    )
    pesquisa_reversa_button.grid(row=5, column=0, columnspan=3, pady=10, sticky="n")

    pesquisa_reversa_label = tk.Label(
        aba3,
        text="Informações extraídas dos arquivos pesquisados",
        font=("Montserrat", 11, 'bold')
    )
    pesquisa_reversa_label.grid(row=6, column=0, columnspan=3, pady=2, sticky="n")

    ####### Funções auxiliares para selecionar diretório #######
    def btn_selecionar_diretorio(entry):
        global diretorio_selecionado
        diretorio_selecionado = filedialog.askdirectory()
        if diretorio_selecionado:
            entry.delete(0, tk.END)
            entry.insert(0, diretorio_selecionado)

    janela.mainloop()


if __name__ == "__main__":
    iniciar_interface()