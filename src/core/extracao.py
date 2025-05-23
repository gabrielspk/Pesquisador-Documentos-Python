import os

def busca_reversa(diretorio_atual, posicao, comprimento, nomenclaturaDescartada):
    documentos_extraidos = []
    linha_cabecalho = 0

    try:
        posicao = int(posicao)
        comprimento = int(comprimento)

        for root, _, files in os.walk(diretorio_atual):
            for file in files:
                if not file.endswith(nomenclaturaDescartada):
                    caminho_arquivo = os.path.join(root, file)

                    with open(caminho_arquivo, 'r', encoding='utf-8', errors="ignore") as f:
                        for linha in f:
                            if file.endswith(".csv") or file.endswith(".CSV"):
                                    if linha_cabecalho == 0:
                                        linha_cabecalho += 1
                                        continue
                                    tabulador = linha.split(";")
                                    if posicao >= len(tabulador):
                                        documentos_extraidos.append("Posição declarada fora do range")
                                    else:
                                        nova_posicao = posicao - 1 if posicao > 1 else 0
                                        documento = tabulador[nova_posicao][:comprimento]
                                        documentos_extraidos.append(documento)
                            else:
                                nova_posicao = posicao - 1 if posicao > 1 else 0
                                documento = linha[nova_posicao:nova_posicao + comprimento]
                                documentos_extraidos.append(documento)
        return documentos_extraidos
    except ValueError:
        return ["Erro: Posição e comprimento precisam ser números inteiros."]