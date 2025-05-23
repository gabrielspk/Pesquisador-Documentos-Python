import pandas as pd
import os

def criar_dataframe(file_path): #função responsável pela criação do dataframe
    #criando dataframe
    df = pd.read_excel(file_path)
    
    #aplicando filtro para que todas nomenclaturas fiquem em mínusculo
    df['Nome Arquivo'] = df['Nome Arquivo'].str.lower()

    df['Processo'] = df['Processo'].str.lower()

    #aplicando um filtro para retirar o que vai ser desnecessário para a validação. Por exemplo: retirando tudo que tem ret*
    df = df.loc[~df['Processo'].str.contains(r"redecard_envioarscliente_rem.*|ret.*|reprog.*|enviopos.*", na=False)]

    df['Nome Arquivo'] = df['Nome Arquivo'].str.replace('xlsx', '')

    qtde_registros_concatenados = 0
    qtde_registros_originais = 0

    if df['Nome Arquivo'].str.contains(r"concaten.*").any(): #utilizando método any() para caso haja qualquer arquivo concaten, entre nessa validação.
        #aplicando um filtro para pegar tudo que tem concaten, porém retirando o que tem .fpl, para validar se tudo que foi concatenado bate com os originais
        df_registros_concatenados = df[df['Nome Arquivo'].str.contains(r"concaten.*") & ~df['Nome Arquivo'].str.contains(r'\.fpl')]
        
        qtde_registros_concatenados = df_registros_concatenados['Registros'].sum() #faz a soma somente dos registros concatenados
        
        print(f"Quantidade registros concatenados: {qtde_registros_concatenados}")

        #df_registros_originais = df[df['Nome Arquivo'].str.contains(r'import_padrao.*')]
        #aplicando um filtro para pegar somente os arquivos originais padrão API.
        df_registros_originais = df[df['Status'].str.contains('Concatenado')]

        qtde_registros_originais = df_registros_originais['Registros'].sum() #realizando a soma dos registros originais.
        
        print(f"Quantidade registros originais: {qtde_registros_originais}")

    #classificando os arquivos pela nomenclatura de A a Z
    df = df.sort_values(by='Nome Arquivo', ascending=True)


    #df['Registros JALL'] = df['Registros JALL'].apply(lambda x: int(x) if pd.notnull(x) else x)

    #verificando caso o primeiro registro termine com .fpl, só comece a validação após não houver mais o mesmo.
    while not df.empty and df.iloc[0]['Nome Arquivo'].endswith('.fpl'):
            df = df.iloc[1:] #pula a linha atual

    print(df)

    return df, qtde_registros_concatenados, qtde_registros_originais

def validar_quantidades(dataframe): #função responsável por validar as quantidades
    divergenciasPlanilha = []
    arquivosValidados = []

    #iniciando laço de repetição para percorrer todo o dataframe do processamento
    for i in range(0, len(dataframe)):
        original = dataframe.iloc[i] #declarando variável do arquivo original
        hora = dataframe.iloc[i] #declarando variável para pegar a data e hora do processamento
        if i < len(dataframe) - 1: #validação para acessar os arquivos de validação corretamente e não exceder o len da planilha
            validacao = dataframe.iloc[i + 1]
        
        #validando se a nomenclatura do arquivo original NÃO termina com fpl e se o arquivo de validação (subsequente) termina com .fpl
        if not original['Nome Arquivo'].endswith('.fpl') and validacao['Nome Arquivo'].endswith('.fpl'):
            if original['Registros'] != validacao['Registros']: #verificando se há divergência de valor
                divergenciasPlanilha.append({ #inserindo os valores dentro da lista declarada para montagem da planilha de divergência
                    'Nome Original': original['Nome Arquivo'],
                    'Quantidade Original': original['Registros'],
                    'Nome Validacao': validacao['Nome Arquivo'],
                    'Quantidade Validacao': validacao['Registros'],
                    'Hora Processamento': hora['Data']
                })

            else:
                 arquivosValidados.append({ #Caso não haja divergência, armazenando os valores validados da outra lista declarada
                      'NomeArquivo': original['Nome Arquivo'],
                      'Quantidade': original['Registros']
                 })
                 arquivosValidados.append({
                      'NomeArquivo': validacao['Nome Arquivo'],
                      'Quantidade': validacao['Registros']
                 })
        #Validação para qualquer arquivo que esteja no status entregar ou erro, ser adicionado a planilha de divergência para análise
        if original['Status'] == "Entregar" or original['Status'] == "Erro":
             divergenciasPlanilha.append({
                    'Nome Original': original['Nome Arquivo'],
                    'Quantidade Original': original['Registros'],
                    'Nome Validacao': original['Nome Arquivo'],
                    'Quantidade Validacao': original['Registros'],
                    'Hora Processamento': hora['Data'],
                    'Status': original['Status']
                })

    #VERIFICAR     
    divergenciasPlanilhaConvertida = "\n\n".join(
    [f"{item.get('Nome Original', 'N/A')} / {item.get('Quantidade Original', 'N/A')}\n"
     f"{item.get('Nome Validacao', 'Não disponível')} / {item.get('Quantidade Validacao', 'N/A')}\n"
     f"Horário: {item.get('Hora Processamento', 'N/A')}".strip() for item in divergenciasPlanilha]
)

    return pd.DataFrame(divergenciasPlanilha), pd.DataFrame(arquivosValidados), divergenciasPlanilhaConvertida

def concatenar_arquivos(caminhoArquivos, caminhoRelatorio): #função responsável por contenar arquivos dentro do diretório
    listarArquivos = os.listdir(caminhoArquivos) #listando os arquivos no diretório
    
    #listando os arquivos que começam com "Arquivos Processados"
    listarCaminhoEArquivos = [caminhoArquivos + '\\' + arquivo for arquivo in listarArquivos if arquivo.startswith('Arquivos Processados')]

    dfConcatenado = pd.DataFrame()

    for arquivo in listarCaminhoEArquivos:
        dados = pd.read_excel(arquivo)
        dfConcatenado = dfConcatenado._append(dados)

    dfConcatenado.to_excel(caminhoRelatorio)

    return dfConcatenado

def processar_validacao(diretorio_atual):

    mensagens = []
    caminhoRelatorio = os.path.join(diretorio_atual, "RelatorioConcatenado.xlsx")
    listaArquivosDiretorio = os.listdir(diretorio_atual)
    listaArquivosProcessados = [arquivo for arquivo in listaArquivosDiretorio if arquivo.startswith("Arquivos Processados")]

    if not listaArquivosProcessados:
        mensagens.append("Nenhum arquivo processado encontrado no diretório.")
        return None, mensagens

    nomenclaturaArquivoProcessado = listaArquivosProcessados[0]
    planilhaOrdenada = os.path.join(diretorio_atual, "RelatorioOrdenado(VERIFICAR).xlsx")

    if len(listaArquivosProcessados) > 1:
        dfConcatenado = concatenar_arquivos(diretorio_atual, caminhoRelatorio)
        dfConcatenado, qtde_registros_concatenados, qtde_registros_Originais = criar_dataframe(caminhoRelatorio)
        dfConcatenado.to_excel(planilhaOrdenada, index=False)
        divergenciasPlanilha, arquivosValidados, divergenciasPlanilhaConvertida = validar_quantidades(dfConcatenado)
    else:
        df, qtde_registros_concatenados, qtde_registros_Originais = criar_dataframe(os.path.join(diretorio_atual, nomenclaturaArquivoProcessado))
        df.to_excel(planilhaOrdenada, index=False)
        divergenciasPlanilha, arquivosValidados, divergenciasPlanilhaConvertida = validar_quantidades(df)

    #Preparar mensagens para log conforme resultado
    if qtde_registros_concatenados and qtde_registros_Originais != 0:
        mensagens.append(f"Qtde Arquivos originais: {qtde_registros_Originais}")
        mensagens.append(f"Quantidade arquivos concatenados: {qtde_registros_concatenados}")

        if divergenciasPlanilha.empty:
            if qtde_registros_concatenados == qtde_registros_Originais:
                mensagens.append("Não foi encontrado divergências dentre os arquivos validados. Confira o relatório 'RelatorioOrdenado(VERIFICAR).xlsx' para validação.")
            else:
                mensagens.append("Consulte as planilhas para analisar a diferença dentre original e concatenados.")
        else:
            mensagens.append("Foram encontradas as seguintes divergências entre arquivos abaixo:")
            mensagens.append(divergenciasPlanilhaConvertida)
            mensagens.append("Consulte o relatório de divergência para mais informações.")
    else:
        if not divergenciasPlanilha.empty:
            mensagens.append("Foram encontradas as seguintes divergências entre arquivos abaixo:")
            mensagens.append(divergenciasPlanilhaConvertida)
            mensagens.append("Consulte a planilha RelatorioDivergencias.xlsx para mais informações.")
        else:
            mensagens.append("Não foi encontrado nenhuma divergência dentre os arquivos validados. Consulte o relatório RelatorioOrdenado(VERIFICAR).xlsx para verificação.")

    # Retorna dados relevantes e mensagens para exibir
    resultados = {
        "qtde_concatenados": qtde_registros_concatenados,
        "qtde_originais": qtde_registros_Originais,
        "divergencias_df": divergenciasPlanilha,
        "divergencias_texto": divergenciasPlanilhaConvertida,
    }

    #Salva planilhas de divergência
    planilhaDivergencias = os.path.join(diretorio_atual, "RelatorioDivergencias.xlsx")
    divergenciasPlanilha.to_excel(planilhaDivergencias, index=False)

    return resultados, mensagens
