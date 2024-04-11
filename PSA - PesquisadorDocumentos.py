import os
import glob
import openpyxl
import time

diretorioAtual = "C:\\Users\\gabriel.ferreira\\Downloads" #armazena o diretorio que vai ser utilizado

os.chdir(diretorioAtual) #alterando para o diretório selecionado

#diretorio aonde o arquivo com os documentos a serem pesquisados está
diretorioPesquisaDocumentos = "C:\\Users\\gabriel.ferreira\\3D Objects\\Projetos Processamento\\INSIRA AQUI OS DOCUMENTOS A SEREM PESQUISADOS.txt"

#abrindo o diretorio o arquivo em modo de leitura
diretorioPesquisaDocumentos = open(diretorioPesquisaDocumentos, 'r')

#variável que armazena o que não vai ser validado se o arquivo tiver determinado nome na nomenc (ex: retorno)
validacaoNomenclatura = glob.glob('Retorno')

#nessa váriavel, é armazenada todas as extensões dos arquivos que não queremos usar na pesquisa
extensoesDescartadas = (".fpl", ".zip", ".ini", ".pdf")

#varíavel aonde vai ser armazenados os documentos
documentos = []

#dicionário onde vai ser armazenado o documento (chave) e em qual arquivo foi encontrado (valor)
documentosEncontrados = {}

#laço de repetição aonde vai ser pesquisado cada documento dentro do arquivo de texto
for documentoInserido in diretorioPesquisaDocumentos.readlines():
    documentoInserido = documentoInserido.rstrip() #rstrip usado para não fazer quebra de linha na string
    documentos.append(documentoInserido) #inserindo cada documento passado no laço dentro da variável documentos

documentosNãoEncontrados = set(documentos) #variável aonde vai ser armazenado primeiramente todos os documentos, de forma imutável
qtdeArquivosPesquisados = 0 #variável para realizar a contagem de arquivos pesquisados

#Inicio do cálculo de tempo em que a execução é executada
inicio = time.time() 

#laço de repetição que pesquisa em cada arquivo presente no diretório parametrizado
for root, _, files in os.walk(diretorioAtual):
    for file in files:
        if not file.endswith(extensoesDescartadas) and file not in validacaoNomenclatura:
            for documento in documentosNãoEncontrados.copy():
                if documento in open(os.path.join(root, file), errors="ignore").read():
                    if documento not in documentosEncontrados:
                        documentosEncontrados[documento] = []
                    documentosEncontrados[documento].append(os.path.join(file))
                    documentosNãoEncontrados.remove(documento)
                    print(documentosEncontrados)
            qtdeArquivosPesquisados += 1

#iniciando a planilha no excel
workbook = openpyxl.Workbook()
sheet = workbook.active

#iniciando a planilha do relatório declarando os cabeçalhos
sheet['A1'] = 'Documentos em arquivos'
sheet['B1'] = 'Nomenclatura arquivo'
sheet['C1'] = 'Não encontrado'

#comentado essa linha de código, pois encontrei uma solução melho mas não descartei
'''for row, (documento, arquivos) in enumerate(documentosEncontrados.items(), start=2):
    print(f'Documento {documento} encontrado nos arquivos: {", ".join(arquivos)}. Pesquisa realizada em {qtdeArquivosPesquisados} arquivos')
    for arquivo in enumerate(arquivos, start=1):
        sheet.cell(row=row, column=1).value = documento

for row, documento in enumerate(documentosNãoEncontrados, start=2):
    print(f'documento {documento} não encontrado em arquivos. Pesquisa realizada em {qtdeArquivosPesquisados} arquivos')
    sheet.cell(row=row, column=2).value = documento'''

#Iniciando variáveis para iterar em cada linha na planilha
linhaDocsEncontrados = 1
linhaDocsNaoEncontrados = 1

#Laço para escrever em cada linha do excel as informações
for documento, arquivos in documentosEncontrados.items():
    print(f'Documento {documento} encontrado nos arquivos: {", ".join(arquivos)} Pesquisa realizada em {qtdeArquivosPesquisados}')
    linhaDocsEncontrados += 1
    sheet.cell(row=linhaDocsEncontrados, column=1, value=documento)

    for i, arquivo in enumerate(arquivos,start=1):
        sheet.cell(row=linhaDocsEncontrados, column=2, value=arquivo)

#Laço para escrever em cada linha do excel as informações que não foram encontrado em arquivos
for documento in documentosNãoEncontrados:
    linhaDocsNaoEncontrados += 1
    print(f"Documento {documento} não encontrado em arquivo. Pesquisado em {qtdeArquivosPesquisados} arquivos")
    sheet.cell(row=linhaDocsNaoEncontrados, column=3).value = documento

#finalizando o contabilizador de tempo da execução
fim = time.time()

#exibindo o tempo que foi percorrido na execução da aplicação
print ('Tempo de execucao: %.4f' % (fim - inicio))

#declarando o nome do arquivo .xlsx que vai ser salvo os documentos
nomeArquivo = "RelatorioProcessamento.xlsx"
count = 1 #contador para iterar caso o arquivo já exista, para que não haja duplicidade na pasta

#enquanto existir arquivo, vai ser adicionado uma numeração na nomenclatura para que não haja duplicidade de nomenclatura.
while os.path.exists(nomeArquivo):
    nomeArquivo = f"RelatorioProcessamento({count}).xlsx"
    count += 1

#tratamento de erro e exceção caso o arquivo excel já esteja aberto
try:
    workbook.save(nomeArquivo)
except PermissionError:
    print(f"Não foi possível gerar o relatório, pois está aberto por outro programa.")