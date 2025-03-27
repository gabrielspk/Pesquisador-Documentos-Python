**Pesquisador de Documentos - Validação e Extração de Dados**

📌 **Descrição**

Este projeto é uma ferramenta desenvolvida em Python para automação de tarefas relacionadas a pesquisa, análise e extração de informações dentro de documentos e relatórios. Ele oferece três funcionalidades principais:

Pesquisa de Documentos: Busca documentos em um diretório com base em uma lista fornecida e gera um relatório com os arquivos encontrados e não encontrados.

Análise de Relatórios: Valida planilhas extraídas do GA, verificando discrepâncias entre arquivos concatenados e originais, gerando relatórios de divergências.

Pesquisa Reversa: Extrai dados de arquivos .TXT com base em posições e comprimentos definidos pelo usuário.

⚙️ **Funcionalidades**

🔍 1. Pesquisa de Documentos
Realiza uma busca recursiva em um diretório.

Identifica quais documentos foram encontrados e quais não foram.

Gera um relatório com os resultados da busca.

📌 Como usar:

Selecione o diretório onde os arquivos estão localizados.

Insira os documentos a serem pesquisados (um por linha).

Clique em Pesquisar.

O resultado será exibido no log e salvo em um relatório no diretório selecionado.

📊 2. **Análise de Relatórios**

Ordena e valida dados extraídos do GA.

Compara remessas e FPL para identificar divergências.

Verifica arquivos concatenados para garantir consistência.

Gera duas planilhas: uma com as divergências encontradas e outra com os dados ordenados.

📌 Como usar:

Selecione o diretório contendo os relatórios do GA.

Clique em Validar.

O resultado será salvo no diretório, destacando possíveis falhas no processamento.

📌 Detalhes:

Arquivos com status ENTREGAR ou ERRO são adicionados na planilha de divergências.

Remoção de registros específicos (RET*, REPROG*, ENVIOPOS*) para uma análise mais precisa.

📝 3. **Pesquisa Reversa**

Extrai informações de arquivos .TXT com formato padronizado e também de arquivos .CSV tabulados.

Permite definir manualmente a posição e o comprimento da informação a ser extraída.

📌 Como usar:

Selecione o diretório dos arquivos.

Defina a posição inicial da informação.

Defina o comprimento da informação.

Clique em Pesquisar.

O resultado será salvo em um arquivo no diretório escolhido.

🛠️ **Tecnologias Utilizadas**

Python

Pandas (manipulação e validação de dados)

Tkinter (interface gráfica)

Openpyxl (manipulação de arquivos Excel)

