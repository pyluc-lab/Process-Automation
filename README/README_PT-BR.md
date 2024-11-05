# Análise de Vendas

Este projeto realiza a análise de vendas, automatizando processos e gerando relatórios de desempenho. O sistema utiliza três bases de dados em formato Excel e gera arquivos de saída organizados em diretórios específicos.

## Visão Geral

Ele utiliza três arquivos Excel que estão localizados na pasta **Bases de Dados**:

- **Sales**: Contém os dados de vendas das lojas.
- **Emails**: Inclui as informações de contato dos gerentes das lojas.
- **Stores**: Contém informações sobre as lojas.

***IMPORTANTE***: No arquivo "**Emails**", troque os valores da coluna "**E-mail**" para um endereço de e-mail que você possa testar o script.

Após a execução do script, os arquivos gerados são salvos na pasta **Backup Arquivos Lojas**. Para cada loja, uma subpasta será criada (caso ainda não exista) e, dentro dela, um arquivo `.xlsx` com as informações da loja. Além disso, dois arquivos adicionais serão salvos na mesma pasta:

- **Ranking Anual**: Um arquivo que contém o ranking de faturamento anual das lojas.
- **Ranking do Dia**: Um arquivo com o ranking de faturamento do dia em análise.

## Outlook

É de extrema importância ter o aplicativo do Outlook instalado no computador e aberto como administrador quando for rodar o script.

## Pré-requisitos

Antes de executar o código, certifique-se de que você possui os seguintes arquivos e dependências:

1. **Arquivos de Dados**:

   - Três arquivos Excel na pasta **Bases de Dados**:
     - Um arquivo com os dados de vendas, incluindo as colunas:
       - `Store ID`
       - `Total Value`
       - `Date`
     - Um arquivo com informações de e-mail das lojas, incluindo:
       - `Store`
       - `Manager`
       - `E-mail`
     - Um arquivo com informações das lojas, incluindo:
       - `Store ID`
       - `Store`
2. **Diretório de Backup**:

   - A pasta **Backup Arquivos Lojas** será criada automaticamente pelo script para armazenar os relatórios.
3. **Dependências**:

   - Certifique-se de ter as seguintes bibliotecas Python instaladas:
     - `pandas`
     - `openpyxl` (para manipulação de arquivos Excel)
     - `pywin32` (para interação com o Outlook)
     - `logging`

## Como Executar

1. Certifique-se de ter o Python instalado em seu sistema.
2. Clone este repositório ou faça o download dos arquivos.
3. Abra um terminal e navegue até o diretório do projeto.
4. Execute o arquivo principal com o comando:
   ```bash
   python nome_do_arquivo.py
   ```

## Estrutura do Código

O código é organizado nas seguintes funções principais:

* `main()`: Função principal que executa o fluxo de trabalho do programa de análise de vendas.
* `envia_email()`: Envia e-mails personalizados para os gerentes das lojas.
* `cria_rankings()`: Cria rankings de faturamento anual e diário das lojas.
* `email_diretoria()`: Envia um e-mail para a diretoria com os resultados de faturamento.
