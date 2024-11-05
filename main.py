import pandas as pd
import pathlib
import win32com.client as win32
import pythoncom
import time
import os
import logging

def configura_log():
    """
    Configura o sistema de logging para o aplicativo.

    Esta função define as configurações do logging, incluindo o nome do arquivo
    onde os logs serão salvos, o nível mínimo de mensagens a serem registradas 
    e a codificação do arquivo.

    O arquivo de log é nomeado como 'app.log' e registra mensagens a partir do nível 
    INFO e acima.

    Returns:
        None
    """
    try:
        # Definindo o diretório de logs
        log_dir = 'Logs'
        
        # Criando a pasta "Logs" se ela não existir
        os.makedirs(log_dir, exist_ok=True)

        # Definindo o nome do arquivo de log com o caminho completo
        log_file = os.path.join(log_dir, 'app.log')

        # Validação do nome do arquivo
        if not isinstance(log_file, str) or not log_file.endswith('.log'):
            raise ValueError("O nome do arquivo deve ser uma string terminando com '.log'.")

        # Validação do nível de log
        log_level = logging.INFO
        if log_level not in [logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL]:
            raise ValueError("Nível de log inválido.")

        # Validação da codificação
        encoding = 'UTF-8'
        if not isinstance(encoding, str) or not encoding.isascii():
            raise ValueError("A codificação deve ser uma string ASCII válida.")

        logging.basicConfig(
            filename=log_file,
            level=log_level,
            encoding=encoding
        )
        logging.INFO(f"Arquivo de log criado em: {log_file}")

    except Exception as e:
        logging.error(f"Erro ao configurar log: {e}") # Capturando qualquer erro que possa aparecer



def configura_outlook():
    """
    Configura e inicializa o Outlook.

    Esta função inicializa uma nova instância do Outlook. 

    Mensagens de log são registradas para indicar se a 
    inicialização foi bem-sucedida ou se ocorreram erros.

    Returns:
        outlook: Objeto da aplicação Outlook se a inicialização for bem-sucedida, 
                 caso contrário, retorna None.
    """
    # Validação de ambiente
    if not os.name == 'nt':
        logging.error("Esta função é suportada apenas em sistemas Windows.")
        return None

    # Inicializar Outlook
    try:
        pythoncom.CoInitialize()  # Inicializa o COM
        outlook = win32.Dispatch('outlook.application')
        logging.info("Outlook iniciado com sucesso.")
        return outlook
    except Exception as e:
        logging.error(f"Erro ao inicializar o Outlook: {e}")
        return None


def formata_valores(valor):
    """
    Formata um valor numérico como uma string em formato monetário.

    Parâmetros:
    valor: O valor numérico a ser formatado.

    Retorna:
    str: O valor formatado como uma string com duas casas decimais e separadores de milhar.
    """
    return f'{valor:,.2f}'


def define_cores(valor_1, valor_2):
    """
    Define a cor com base na comparação entre dois valores.

    Parâmetros:
    valor_1: O primeiro valor a ser comparado.
    valor_2: O segundo valor a ser comparado.

    Retorna:
    str: 'green' se valor_1 for maior ou igual a valor_2; 'red' caso contrário.

    """
    return 'green' if valor_1 >= valor_2 else 'red'


def carrega_dfs(caminho=None):
    """
    Carrega DataFrames a partir de arquivos Excel e CSV.

    Esta função tenta ler três arquivos: um arquivo Excel com dados de vendas, 
    um arquivo Excel com endereços de e-mail e um arquivo CSV com dados de lojas. 
    Se algum dos arquivos não for encontrado, a função captura a exceção e 
    registra uma mensagem de erro.

    Parâmetros:
    caminhos (dict): Um dicionário com os caminhos dos arquivos a serem carregados.
                        Se None, utiliza os caminhos padrão.

    Retorna:
        tuple: Um tupla contendo três DataFrames (vendas, emails, lojas).
               Se ocorrer um erro ao carregar os arquivos, retorna (None, None, None).
    """
    # Definindo os caminhos dos arquivos

    if caminho is None:
        caminhos = {
            'vendas': r'Bases de Dados/Sales.xlsx',
            'emails': r'Bases de Dados/Emails.xlsx',
            'lojas': r'Bases de Dados/Stores.csv'
        }
        
    # Inicializando dicionário para armazenar DataFrames
    dataframes = {}

    # Tentando carregar os DataFrames
    for nome, caminho in caminhos.items():
        if not os.path.exists(caminho):
            logging.error(f'Arquivo não encontrado: {caminho}')
            return None, None, None
        
        try:
            if nome in ['vendas', 'emails']:
                dataframes[nome] = pd.read_excel(caminho, engine='openpyxl')
            elif nome == 'lojas':
                dataframes[nome] = pd.read_csv(caminho, encoding='latin1', sep=';')
        except Exception as e:
            logging.error(f'Erro ao carregar o arquivo {caminho}: {e}')
            return None, None, None

    return dataframes['vendas'], dataframes['emails'], dataframes['lojas']


def merge_dfs(sales, stores, coluna_referencia):
    """
    Mescla dois DataFrames com base em uma coluna de referência.

    Esta função combina os DataFrames `sales` e `stores` utilizando a 
    coluna especificada como referência. Se um dos DataFrames estiver vazio 
    ou se a coluna de referência não existir em um dos DataFrames, uma 
    mensagem de erro é registrada e a função retorna None.

    Args:
        sales (DataFrame): DataFrame contendo os dados de vendas.
        stores (DataFrame): DataFrame contendo os dados das lojas.
        coluna_referencia (str): Nome da coluna pela qual os DataFrames devem ser mesclados.

    Returns:
        DataFrame: O DataFrame resultante da mesclagem, ou None em caso de erro.
    
    Raises:
        ValueError: Se os argumentos não forem DataFrames válidos ou se a coluna de referência não existir.
    """

    # Validação dos DataFrames
    if not isinstance(sales, pd.DataFrame) or not isinstance(stores, pd.DataFrame):
        logging.error('Erro: Um ou mais argumentos não são DataFrames válidos.')
        return None

    if sales.empty or stores.empty:
        logging.error('Erro: Um ou mais DataFrames estão vazios.')
        return None

    # Verificação da coluna de referência
    if coluna_referencia not in sales.columns:
        logging.error(f"Erro: a coluna '{coluna_referencia}' não existe no DataFrame de sales.")
        return None

    if coluna_referencia not in stores.columns:
        logging.error(f"Erro: a coluna '{coluna_referencia}' não existe no DataFrame de stores.")
        return None

    # Tentativa de mesclagem dos DataFrames
    try:
        merged_df = sales.merge(stores, on=coluna_referencia)
    except Exception as e:
        logging.error(f'Erro inesperado durante a mesclagem: {e}')
        return None
    
    return merged_df


def cria_tabelas_para_lojas(stores, sales, coluna_referencia):
    """
    Cria um dicionário de DataFrames separados para cada loja.

    Esta função verifica se os DataFrames `stores` e `sales` estão vazios e se a coluna 
    de referência existe em ambos. Em seguida, ela cria um dicionário onde cada chave é 
    o nome de uma loja e o valor é um DataFrame contendo as vendas correspondentes a essa loja.

    Args:
        stores (DataFrame): DataFrame contendo os dados das lojas.
        sales (DataFrame): DataFrame contendo os dados das vendas.
        coluna_referencia (str): Nome da coluna pela qual as vendas serão agrupadas por loja.

    Returns:
        dict: Um dicionário onde as chaves são os nomes das lojas e os valores são DataFrames 
              com as vendas correspondentes. Retorna um dicionário vazio em caso de erro ou 
              se as validações não forem atendidas.
    """

    # Validação dos DataFrames
    if not isinstance(stores, pd.DataFrame) or not isinstance(sales, pd.DataFrame):
        logging.error('Erro: Um ou mais argumentos não são DataFrames válidos.')
        return {}
    
    if stores.empty or sales.empty:
        logging.error('Erro: Um ou mais DataFrames estão vazios.')
        return {}

    # Verificação da coluna de referência
    if coluna_referencia not in stores.columns:
        logging.error(f"Erro: A coluna '{coluna_referencia}' não existe no DataFrame 'stores'.")
        return {}
    
    if coluna_referencia not in sales.columns:
        logging.error(f"Erro: A coluna '{coluna_referencia}' não existe no DataFrame 'sales'.")
        return {}

    dict_lojas = {}
    try:
        for loja in stores[coluna_referencia].unique():  # Usar unique() para evitar duplicação
            dict_lojas[loja] = sales[sales[coluna_referencia] == loja]
    except Exception as e:
        logging.error(f'Erro inesperado ao criar tabelas para lojas: {e}')
        return {}
    
    return dict_lojas


def cria_indicador_dia(sales):
    """
    Cria um indicador representando o dia mais recente nas vendas.

    Esta função verifica se o DataFrame `sales` está vazio e, em seguida, 
    identifica o dia mais recente presente na coluna 'Date'. Se a coluna não existir, 
    ou se ocorrer qualquer outro erro durante o processamento, uma mensagem de erro é registrada.

    Args:
        sales (DataFrame): DataFrame contendo os dados de vendas, incluindo uma coluna 'Date'.

    Returns:
        datetime: O dia mais recente encontrado na coluna 'Date', ou None em caso de erro.
    
    Raises:
        ValueError: Se a coluna 'Date' não puder ser convertida para o tipo datetime.
    """

    # Validação do DataFrame
    if sales is None or sales.empty:
        logging.error("Erro: O DataFrame 'sales' está vazio.")
        return None
    
    # Verificação da coluna 'Date'
    if 'Date' not in sales.columns:
        logging.error("Erro: A coluna 'Date' não existe no DataFrame 'sales'.")
        return None

    try:
        # Garantir que a coluna 'Date' é do tipo datetime
        sales['Date'] = pd.to_datetime(sales['Date'], errors='coerce')
        dia_indicador = sales['Date'].max()
        
        # Verificar se dia_indicador é um valor válido
        if pd.isnull(dia_indicador):
            logging.error("Erro: Não foi possível determinar um dia válido a partir da coluna 'Date'.")
            return None

    except Exception as e:
        logging.error(f"Erro inesperado ao criar dia indicador: {e}")
        return None

    return dia_indicador


def pega_caminho_backup():
    """
    Obtém o caminho do diretório de backup.

    Esta função verifica se o diretório 'Backup Arquivos Lojas' existe. 
    Se o diretório não existir, cria a pasta.

    Returns:
        Path: O objeto Path representando o diretório de backup.
    """
    caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

    # Verificação da existência do diretório
    if not caminho_backup.is_dir():
        try:
            # Cria a nova pasta, se não existir
            caminho_backup.mkdir(parents=True, exist_ok=True)
            logging.info(f"Pasta: '{caminho_backup}' criada com sucesso.")
        except Exception as e:
            logging.error(f"Erro ao criar a pasta '{caminho_backup}': {e}")
            return None

    return caminho_backup


def lista_nomes_doBackup(caminho_backup):
    """
    Lista os nomes dos arquivos no diretório de backup.

    Esta função verifica se o caminho de backup é válido, se o diretório existe e 
    se é um diretório válido. Em seguida, retorna uma lista com os nomes dos arquivos 
    contidos no diretório. Se houver algum erro, uma mensagem de erro é registrada.

    Args:
        caminho_backup (Path): O caminho do diretório de backup.

    Returns:
        list: Uma lista contendo os nomes dos arquivos no diretório de backup, ou uma lista vazia em caso de erro.
    """
    # Verificação se o caminho de backup é None
    if caminho_backup is None:
        logging.error("Erro: O caminho de backup é None.")
        return []
    
    # Verificação da existência do diretório
    if not caminho_backup.exists():
        logging.error(f"Erro: O diretório '{caminho_backup}' não existe.")
        return []
    
    # Verificação se é um diretório válido
    if not caminho_backup.is_dir():
        logging.error(f"Erro: '{caminho_backup}' não é um diretório válido.")
        return []
    
    try:
        # Listando os arquivos no diretório
        arquivos_pasta_backup = caminho_backup.iterdir()
        lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup if arquivo.is_file()]
    except Exception as e:
        logging.error(f'Erro inesperado ao listar arquivos: {e}')
        return []
    
    return lista_nomes_backup


def cria_pastas_para_lojas(dict_lojas, lista_nomes_backup, caminho_backup):
    """
    Cria pastas para lojas no diretório de backup.

    Esta função verifica se o caminho de backup é válido e, em seguida, 
    cria uma nova pasta para cada loja que ainda não existe no diretório 
    de backup. Se ocorrer algum erro durante a criação da pasta, uma 
    mensagem de erro é registrada.

    Args:
        dict_lojas (dict): Um dicionário onde as chaves são os nomes das lojas.
        lista_nomes_backup (list): Uma lista contendo os nomes das pastas existentes no diretório de backup.
        caminho_backup (Path): O caminho do diretório de backup onde as pastas serão criadas.

    Returns:
        None: A função não retorna nada, mas registra mensagens sobre o status da criação das pastas.
    """
    # Verificação do caminho de backup
    if caminho_backup is None:
        logging.error("Erro: O caminho de backup é None.")
        return
    
    if not caminho_backup.exists():
        logging.error(f"Erro: O diretório '{caminho_backup}' não existe.")
        return

    if not caminho_backup.is_dir():
        logging.error(f"Erro: '{caminho_backup}' não é um diretório válido.")
        return

    # Criação das pastas para as lojas
    for loja in dict_lojas:
        if loja not in lista_nomes_backup:
            nova_pasta = caminho_backup / loja
            try:
                # Cria a nova pasta, se não existir
                nova_pasta.mkdir(parents=True, exist_ok=True)
                logging.info(f"Pasta: '{nova_pasta}' criada com sucesso.")
            except Exception as e:
                logging.error(f"Erro ao criar a pasta '{nova_pasta}': {e}")


def salva_excel_para_cada_loja(dict_lojas, dia_indicador, caminho_backup):
    """
    Salva um arquivo Excel para cada loja no diretório de backup.

    Esta função verifica se o caminho de backup é válido e, em seguida, 
    salva um arquivo Excel para cada loja no dicionário `dict_lojas`. 
    O nome do arquivo é baseado na data do indicador e no nome da loja. 
    Se o DataFrame da loja estiver vazio, uma mensagem de erro é registrada, 
    e o arquivo não é salvo. Qualquer erro durante o salvamento também é tratado.

    Args:
        dict_lojas (dict): Um dicionário onde as chaves são os nomes das lojas 
                           e os valores são DataFrames com as vendas correspondentes.
        dia_indicador (datetime): O dia indicador que será usado para nomear os arquivos.
        caminho_backup (Path): O caminho do diretório de backup onde os arquivos serão salvos.

    Returns:
        None: A função não retorna nada, mas registra mensagens sobre o status do salvamento dos arquivos.
    """
    # Verificação do caminho de backup
    if caminho_backup is None:
        logging.error("Erro: O caminho de backup é None.")
        return
    
    if not caminho_backup.exists() or not caminho_backup.is_dir():
        logging.error(f"Erro: O diretório '{caminho_backup}' não existe ou não é um diretório válido.")
        return

    # Salvar arquivos Excel para cada loja
    for loja, df in dict_lojas.items():
        nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
        local_arquivo = caminho_backup / loja / nome_arquivo
        
        try:
            if df.empty:
                logging.error(f"Erro: O DataFrame para a loja '{loja}' está vazio. Não será salvo.")
                continue
            
            # Certifica que o diretório da loja existe
            (caminho_backup / loja).mkdir(parents=True, exist_ok=True)
            
            # Salvar DataFrame como arquivo Excel
            df.to_excel(local_arquivo, index=False)
            logging.info(f"Arquivo '{local_arquivo}' salvo com sucesso.")
        
        except Exception as e:
            logging.error(f"Erro ao salvar o arquivo '{local_arquivo}': {e}")


def metas():
    """
    Define e retorna as metas anuais e diárias de faturamento e produtos.

    A função cria dois dicionários: um para as metas anuais e outro para as metas diárias.
    Em caso de erro durante a criação dos dicionários, um log de erro é gerado.

    Returns:
    tuple: Um tupla contendo dois dicionários:
        - dict_metas_ano: Metas anuais (faturamento, quantidade de produtos e ticket médio).
        - dict_metas_dia: Metas diárias (faturamento, quantidade de produtos e ticket médio).
    """
    try:    
        dict_metas_ano = {
            'meta_faturamento_ano': 1500000,
            'meta_qtd_produtos_ano': 120,
            'meta_ticketMedio_ano': 440
        }

        dict_metas_dia = {
            'meta_faturamento_dia': 800,
            'meta_qtd_produtos_dia': 3,
            'meta_ticketMedio_dia': 400
        }
        
        return dict_metas_ano, dict_metas_dia

    except Exception as e:
        logging.error(f'Erro: Não foi possível criar metas: {e}')
        return {}, {}  # Retorna dicionários vazios em caso de erro


def calcula_indicadores(dict_lojas, dia_indicador):
    """
    Calcula indicadores de vendas para múltiplas lojas em um dia específico e no ano.

    Para cada loja no dicionário fornecido, a função calcula o faturamento, a diversidade de produtos 
    e o ticket médio, tanto para o ano quanto para um dia específico. Em caso de erro durante o cálculo,
    mensagens de log são geradas.

    Args:
        dict_lojas (dict): Um dicionário onde cada chave é o nome de uma loja e o valor é um DataFrame 
                           contendo os dados de vendas, incluindo as colunas 'Date', 'Total Value' e 'Product'.
        dia_indicador (str): A data no formato 'YYYY-MM-DD' para a qual os indicadores diários 
                             devem ser calculados.

    Returns:
        dict: Um dicionário com os indicadores para cada loja, contendo:
            - 'indicadores_ano': Outro dicionário com faturamento total, quantidade de produtos e ticket médio do ano.
            - 'indicadores_dia': Outro dicionário com faturamento total, quantidade de produtos e ticket médio do dia.
    
    Logs:
        Registra mensagens de erro caso os DataFrames estejam vazios ou se as colunas esperadas não existirem.
    """
    indicadores_todas_lojas = {}

    # Calculando indicadores
    for loja, vendas_loja in dict_lojas.items():
        if vendas_loja.empty:
            logging.error(f"Erro: O DataFrame para a loja '{loja}' está vazio. Indicadores não calculados.")
            continue

        vendas_loja_dia = vendas_loja.loc[vendas_loja['Date'] == dia_indicador]

        # Faturamento
        faturamento_ano = vendas_loja['Total Value'].sum()
        faturamento_dia = vendas_loja_dia['Total Value'].sum() if not vendas_loja_dia.empty else 0

        # Diversidade de produtos
        qtd_produtos_ano = len(vendas_loja['Product'].unique())
        qtd_produtos_dia = len(vendas_loja_dia['Product'].unique()) if not vendas_loja_dia.empty else 0

        # ticket médio ano
        try:
            valor_venda = vendas_loja.groupby('Sale Code').sum(numeric_only=True)
            ticket_medio_ano = valor_venda['Total Value'].mean() if not valor_venda.empty else 0
        except KeyError:
            logging.error(f"Erro: A coluna 'Total Value' não existe para a loja '{loja}'.")
            ticket_medio_ano = 0
        
        # ticket médio dia
        try:
            valor_venda_dia = vendas_loja_dia.groupby('Sale Code').sum(numeric_only=True)
            ticket_medio_dia = valor_venda_dia['Total Value'].mean() if not valor_venda_dia.empty else 0
        except KeyError:
            logging.error(f"Erro: A coluna 'Total Value' não existe para a loja '{loja}' no dia {dia_indicador}.")
            ticket_medio_dia = 0

        indicadores_ano = {
            'faturamento_ano': faturamento_ano,
            'qtd_produtos_ano': qtd_produtos_ano,
            'ticket_medio_ano': ticket_medio_ano
        }

        indicadores_dia = {
            'faturamento_dia': faturamento_dia,
            'qtd_produtos_dia': qtd_produtos_dia,
            'ticket_medio_dia': ticket_medio_dia
        }

        indicadores_todas_lojas[loja] = {
            'indicadores_ano': indicadores_ano,
            'indicadores_dia': indicadores_dia
        }

    return indicadores_todas_lojas


def envia_email(outlook, dict_lojas, dia_indicador, caminho_backup, emails):
    """
    Envia um e-mail com indicadores de desempenho de vendas para cada loja, baseado em um dia específico.

    A função calcula metas e indicadores de vendas para as lojas e envia um e-mail personalizado 
    para os gerentes de cada loja, incluindo dados de faturamento, quantidade de produtos e ticket médio, 
    tanto para o dia quanto para o ano.

    Args:
        outlook (object): Um objeto do Outlook para criar e enviar e-mails.
        dict_lojas (dict): Um dicionário onde cada chave é o nome de uma loja e o valor é um DataFrame 
                           contendo os dados de vendas, incluindo colunas 'Date', 'Total Value' e 'Product'.
        dia_indicador (datetime): A data para a qual os indicadores devem ser calculados.
        caminho_backup (str): O caminho para o diretório onde as planilhas de backup estão armazenadas.
        emails (DataFrame): Um DataFrame contendo informações sobre as lojas, gerentes e e-mails.

    Returns:
        None: A função não retorna valor, mas envia e-mails e gera logs sobre o sucesso ou falha do envio.

    Logs:
        Registra informações sobre o sucesso ou falha do envio de e-mails, além de erros ao calcular metas e indicadores.
    """

    try:
        # unpacking das metas 
        dict_metas_ano, dict_metas_dia = metas()

        # metas ano
        meta_faturamento_ano = formata_valores(dict_metas_ano['meta_faturamento_ano'])
        meta_qtd_produtos_ano = dict_metas_ano['meta_qtd_produtos_ano']
        meta_ticketMedio_ano = formata_valores(dict_metas_ano['meta_ticketMedio_ano'])

        # metas dia
        meta_faturamento_dia = formata_valores(dict_metas_dia['meta_faturamento_dia'])
        meta_qtd_produtos_dia = dict_metas_dia['meta_qtd_produtos_dia']
        meta_ticketMedio_dia = formata_valores(dict_metas_dia['meta_ticketMedio_dia'])

        # unpacking indicadores
        indicadores_todas_lojas = calcula_indicadores(dict_lojas, dia_indicador)    
    except Exception as e:
        logging.error(f'Erro: Não foi possível obter metas e/ou indicadores: {e}')
        return None
    
    for nome_loja in indicadores_todas_lojas.keys():
        try:
            # Verificando se a loja existe no DataFrame
            if nome_loja in emails['Store'].values:
                nome = emails.loc[emails['Store'] == nome_loja, 'Manager'].values[0]  
                email_destino = emails.loc[emails['Store'] == nome_loja, 'E-mail'].values[0]

                # Validações
                if not nome or not email_destino:
                    logging.error(f"Dados inválidos para a loja '{nome_loja}': Manager ou e-mail não encontrado.")
                    continue

                mail = outlook.CreateItem(0)
                mail.To = email_destino
                mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Store {nome_loja}'

                # pegando os indicadores ano/dia da loja atual
                indicadores_ano = indicadores_todas_lojas[nome_loja]['indicadores_ano']
                indicadores_dia = indicadores_todas_lojas[nome_loja]['indicadores_dia']

                # Verifica se os indicadores estão presentes
                if 'faturamento_ano' not in indicadores_ano or 'faturamento_dia' not in indicadores_dia:
                    logging.error(f"Indicadores faltando para a loja '{nome_loja}'.")
                    continue

                # indicadores ano
                faturamento_ano = formata_valores(indicadores_ano['faturamento_ano'])
                qtd_produtos_ano = indicadores_ano['qtd_produtos_ano']
                ticket_medio_ano = formata_valores(indicadores_ano['ticket_medio_ano'])

                # indicadores dia
                faturamento_dia = formata_valores(indicadores_dia['faturamento_dia'])
                qtd_produtos_dia = indicadores_dia['qtd_produtos_dia']
                ticket_medio_dia = formata_valores(indicadores_dia['ticket_medio_dia'])

                # definindo cores dos indicadores da loja atual
                cor_fat_dia = define_cores(faturamento_dia, meta_faturamento_dia)
                cor_fat_ano = define_cores(faturamento_ano, meta_faturamento_ano)
                cor_qtde_dia = define_cores(qtd_produtos_dia, meta_qtd_produtos_dia)
                cor_qtde_ano = define_cores(qtd_produtos_ano, meta_qtd_produtos_ano)
                cor_ticket_dia = define_cores(ticket_medio_dia, meta_ticketMedio_dia)
                cor_ticket_ano = define_cores(ticket_medio_ano, meta_ticketMedio_ano)

                mail.HTMLBody = f'''
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
            </head>
            <body style="font-family: Arial, sans-serif; background-color: #f4f4f4; color: #333; margin: 0; padding: 20px;">
                <h1 style="color: #007BFF;">Bom dia, {nome}</h1>

                <p style="line-height: 1.6;">O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {nome_loja}</strong> foi:</p>

                <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fff; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);">
                    <tr>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd; background-color: #007BFF; color: #fff;">Indicador</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Valor Dia</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Meta Dia</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Cenário Dia</th>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Faturamento</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{faturamento_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_faturamento_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_fat_dia};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Diversidade de Produtos</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{qtd_produtos_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_qtd_produtos_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_qtde_dia};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Ticket Médio</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{ticket_medio_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_ticketMedio_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_ticket_dia};">◙</td>
                    </tr>
                </table>

                <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fff; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);">
                    <tr>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd; background-color: #007BFF; color: #fff;">Indicador</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Valor Ano</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Meta Ano</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Cenário Ano</th>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Faturamento</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{faturamento_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_faturamento_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_fat_ano};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Diversidade de Produtos</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{qtd_produtos_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_qtd_produtos_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_qtde_ano};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Ticket Médio</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{ticket_medio_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_ticketMedio_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_ticket_ano};">◙</td>
                    </tr>
                </table>

                <p style="line-height: 1.6;">Segue em anexo a planilha com todos os dados para mais detalhes.</p>
                <p style="line-height: 1.6;">Qualquer dúvida, estou à disposição.</p>

                <p style="margin-top: 20px; font-size: 0.9em; color: #666;">Atenciosamente,<br>Lucas</p>
            </body>
            </html>
        
'''
                
                attachment = pathlib.Path.cwd() / caminho_backup / nome_loja / f'{dia_indicador.month}_{dia_indicador.day}_{nome_loja}.xlsx'
                # importante colocar como str() para evitar erros
                mail.Attachments.Add(str(attachment))

                mail.Send()
                logging.info(f'Email da Loja {nome_loja} enviado para {nome} - {email_destino}')

                # Esperar para evitar conflitos
                time.sleep(0.2)

        except Exception as e:
            logging.error(f"Erro ao enviar email para a loja '{nome_loja}': {e}")


def cria_rankings(dia_indicador, sales):
    """
    Cria rankings de faturamento anual e diário das lojas e salva em arquivos Excel.

    Esta função agrupa os dados de vendas (sales) por loja para calcular o faturamento total
    anual e o faturamento do dia especificado. Os resultados são salvos em dois arquivos Excel 
    separados.

    Args:
        dia_indicador (datetime): O dia para o qual o ranking diário é calculado.
        sales (DataFrame): Um DataFrame contendo dados de vendas, incluindo colunas
                           para 'Store', 'Total Value' e 'Date'.

    Returns:
        tuple: Uma tupla contendo dois DataFrames:
            - faturamento_lojas_ano (DataFrame): Ranking anual das lojas baseado no faturamento total.
            - faturamento_lojas_dia (DataFrame): Ranking diário das lojas baseado no faturamento do dia especificado.

    Raises:
        KeyError: Se as colunas 'Store', 'Total Value' ou 'Date' não forem encontradas no DataFrame de sales.
        Exception: Qualquer outro erro que ocorra durante o processamento.

    Logs:
        Registra erros relacionados à ausência de colunas necessárias e alerta se não houver vendas para o dia indicado.
    """
    try:
        # Validações iniciais
        required_columns = ['Store', 'Total Value', 'Date']
        for column in required_columns:
            if column not in sales.columns:
                logging.error(f"Erro: A coluna '{column}' não encontrada no DataFrame de sales.")
                return None, None

        # Ranking anual
        faturamento_lojas = sales.groupby('Store')[['Total Value']].sum(numeric_only=True)
        faturamento_lojas_ano = faturamento_lojas.sort_values(by='Total Value', ascending=False)
        nome_arquivo_ano = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Anual.xlsx'
        faturamento_lojas_ano.to_excel(fr"Backup Arquivos Lojas/{nome_arquivo_ano}")  # Salvando ranking anual

        # Ranking diário
        vendas_dia = sales.loc[sales['Date'] == dia_indicador]
        if vendas_dia.empty:
            logging.warning(f"Nenhuma venda encontrada para o dia {dia_indicador.strftime('%d/%m/%Y')}.")
            return faturamento_lojas_ano, pd.DataFrame(columns=['Store', 'Total Value'])

        faturamento_lojas_dia = vendas_dia.groupby('Store')[['Total Value']].sum(numeric_only=True)
        faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Total Value', ascending=False)
        nome_arquivo_dia = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx'
        faturamento_lojas_dia.to_excel(fr"Backup Arquivos Lojas/{nome_arquivo_dia}")  # Salvando ranking diário

        return faturamento_lojas_ano, faturamento_lojas_dia

    except Exception as e:
        logging.error(f"Erro inesperado ao criar ranking: {e}")
        return None, None
    

def email_diretoria(outlook, faturamento_lojas_ano, faturamento_lojas_dia, dia_indicador, emails, caminho_backup):
    """
    Envia um e-mail para a diretoria com os resultados de faturamento do dia e do ano.

    Esta função cria um e-mail formatado em HTML contendo os melhores e piores desempenhos de lojas 
    em termos de faturamento, além de anexar os arquivos de ranking anual e diário.

    Args:
        outlook (object): Um objeto do Outlook para criar e enviar e-mails.
        faturamento_lojas_ano (DataFrame): DataFrame contendo o ranking anual de lojas baseado no faturamento.
        faturamento_lojas_dia (DataFrame): DataFrame contendo o ranking diário de lojas baseado no faturamento do dia.
        dia_indicador (datetime): O dia para o qual os resultados são reportados.
        emails (DataFrame): DataFrame contendo informações de e-mail das lojas, incluindo a coluna 'E-mail'.
        caminho_backup (str): O caminho onde os arquivos de ranking estão salvos.

    Raises:
        IndexError: Se não for possível encontrar os dados necessários no DataFrame de emails.
        FileNotFoundError: Se os arquivos de ranking não forem encontrados no caminho especificado.
        Exception: Qualquer outro erro que ocorra durante o envio do e-mail.

    Returns:
        None: A função não retorna valor, mas envia um e-mail e gera logs sobre o sucesso ou falha do envio.

    Logs:
        Registra erros relacionados à ausência de dados de e-mail e arquivos de ranking.
    """

    try:
        # Verifica se existem dados de e-mail para a diretoria
        if 'CEO' not in emails['Store'].values:
            logging.error("Erro: Dados de e-mail da diretoria não encontrados.")
            return
        
        destinatario = emails.loc[emails['Store'] == 'CEO', 'E-mail'].values[0]
        
        # Validações dos DataFrames de faturamento
        if faturamento_lojas_ano.empty or faturamento_lojas_dia.empty:
            logging.error("Erro: Um ou ambos os DataFrames de faturamento estão vazios.")
            return

        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = f'Ranking do dia {dia_indicador.day}/{dia_indicador.month}'
        
        # Formatação do corpo do e-mail
        melhor_loja_dia = faturamento_lojas_dia.index[0]
        pior_loja_dia = faturamento_lojas_dia.index[-1]
        melhor_loja_ano = faturamento_lojas_ano.index[0]
        pior_loja_ano = faturamento_lojas_ano.index[-1]
        
        mail.HTMLBody = f'''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Relatório de Faturamento</title>
</head>
<body>
    <h2>Prezados, bom dia</h2>

    <h3>Resultados do Dia:</h3>
    <p><strong>Melhor loja do dia em faturamento:</strong> Loja <span style="font-weight: bold;">{melhor_loja_dia}</span> com faturamento R$<span style="font-weight: bold;">{faturamento_lojas_dia.iloc[0, 0]:,.2f}</span></p>
    <p><strong>Pior loja do dia em faturamento:</strong> Loja <span style="font-weight: bold;">{pior_loja_dia}</span> com faturamento R$<span style="font-weight: bold;">{faturamento_lojas_dia.iloc[-1, 0]:,.2f}</span></p>

    <h3>Resultados do Ano:</h3>
    <p><strong>Melhor loja do ano em faturamento:</strong> Loja <span style="font-weight: bold;">{melhor_loja_ano}</span> com faturamento R$<span style="font-weight: bold;">{faturamento_lojas_ano.iloc[0, 0]:,.2f}</span></p>
    <p><strong>Pior loja do ano em faturamento:</strong> Loja <span style="font-weight: bold;">{pior_loja_ano}</span> com faturamento R$<span style="font-weight: bold;">{faturamento_lojas_ano.iloc[-1, 0]:,.2f}</span></p>

    <p>Segue em anexo os rankings do ano e do dia de todas as lojas.</p>

    <p>Qualquer dúvida estou à disposição.</p>

    <p>Att.,<br>Lucas</p>
</body>
</html>
'''

        # Anexando arquivos com validação
        for tipo in ['Ranking_Anual', 'Ranking_Dia']:
            attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_{tipo}.xlsx'
            if not attachment.exists():
                logging.error(f"Erro: Arquivo {attachment} não encontrado.")
                continue  # Pula para o próximo arquivo

            mail.Attachments.Add(str(attachment))

        mail.Send()
        logging.info('E-mail da diretoria enviado com sucesso.')
    
    except IndexError as e:
        logging.error(f"Erro: Não foi possível encontrar os dados para o envio do e-mail. Detalhe: {e}")
    except FileNotFoundError as e:
        logging.error(f"Erro: Um ou mais arquivos de ranking não foram encontrados. Detalhe: {e}")
    except Exception as e:
        logging.error(f"Erro inesperado ao enviar e-mail: {e}")


def main():
    """
    Função principal que executa o fluxo de trabalho do programa de análise de vendas.

    Esta função realiza as seguintes etapas:
    1. Carrega os DataFrames de vendas, e-mails e informações das lojas.
    2. Mescla os DataFrames de vendas e lojas usando o identificador da loja.
    3. Cria um dicionário com as informações de cada loja.
    4. Gera um indicador do dia com base nas vendas.
    5. Obtém o caminho para o backup.
    6. Lista os nomes do backup.
    7. Cria pastas para cada loja no diretório de backup.
    8. Salva um arquivo Excel contendo as vendas de cada loja.
    9. Envia e-mails com o relatório de vendas para os gerentes de loja.
    10. Cria rankings de faturamento para o ano e para o dia.
    11. Envia um e-mail com os rankings para a diretoria.

    Raises:
        FileNotFoundError: Se um arquivo necessário não for encontrado durante o processo.
        KeyError: Se houver um erro ao acessar os dados nos DataFrames.
        Exception: Levanta uma exceção para qualquer erro inesperado durante a execução do programa.

    Returns:
        None: A função não retorna valor, mas realiza ações como enviar e-mails e gerar logs sobre o processo.
    """
    try:
        # Configura logging
        configura_log()

        # Configura Outlook
        outlook = configura_outlook()

        # Carrega os DataFrames
        vendas, emails, lojas = carrega_dfs()

        # Mescla os DataFrames de vendas e lojas
        vendas = merge_dfs(vendas, lojas, 'Store ID')

        # Cria um dicionário com as informações de cada loja
        dict_lojas = cria_tabelas_para_lojas(lojas, vendas, 'Store')

        # Cria o indicador do dia
        dia_indicador = cria_indicador_dia(vendas)

        # Obtém o caminho para o backup
        caminho_backup = pega_caminho_backup()

        # Lista nomes do backup
        lista_nomes_backup = lista_nomes_doBackup(caminho_backup)

        # Cria pastas para cada loja no diretório de backup
        cria_pastas_para_lojas(dict_lojas, lista_nomes_backup, caminho_backup)

        # Salva um arquivo Excel para cada loja
        salva_excel_para_cada_loja(dict_lojas, dia_indicador, caminho_backup)

        # Envia os e-mails com relatórios de vendas
        envia_email(outlook, dict_lojas, dia_indicador, caminho_backup, emails)

        # Cria rankings de faturamento
        faturamento_lojas_ano, faturamento_lojas_dia = cria_rankings(dia_indicador, vendas)

        # Espera para evitar conflitos
        time.sleep(0.3)

        # Envia e-mail com os rankings para a diretoria
        email_diretoria(outlook, faturamento_lojas_ano, faturamento_lojas_dia, dia_indicador, emails, caminho_backup)

    except FileNotFoundError as e:
        logging.error(f"Arquivo não encontrado: {e}")
    except KeyError as e:
        logging.error(f"Erro ao acessar dados: {e}")
    except Exception as e:
        logging.error(f"Ocorreu um erro durante a execução do programa: {e}")


# Executar a função principal
try:
    if __name__ == "__main__":
        main()
except Exception as e:
    logging.critical(f"Ocorreu um erro durante a execução do programa: {e}")