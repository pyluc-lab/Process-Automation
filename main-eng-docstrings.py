import pandas as pd
import pathlib
import win32com.client as win32
import pythoncom
import time
import os
import logging

def configura_log():
    """
    Configures the logging system for the application.

    This function sets up the logging configurations, including the name of the file
    where logs will be saved, the minimum level of messages to be logged, 
    and the file encoding.

    The log file is named 'app.log' and records messages from the INFO level and above.

    Returns:
        None
    """
    try:
        # Defining the log directory
        log_dir = 'Logs'
        
        # Creating the "Logs" folder if it doesn't exist
        os.makedirs(log_dir, exist_ok=True)

        # Defining the log file name with the full path
        log_file = os.path.join(log_dir, 'app.log')

        # Validating the file name
        if not isinstance(log_file, str) or not log_file.endswith('.log'):
            raise ValueError("The file name must be a string ending with '.log'.")

        # Validating the log level
        log_level = logging.INFO
        if log_level not in [logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL]:
            raise ValueError("Invalid log level.")

        # Validating the encoding
        encoding = 'UTF-8'
        if not isinstance(encoding, str) or not encoding.isascii():
            raise ValueError("The encoding must be a valid ASCII string.")

        logging.basicConfig(
            filename=log_file,
            level=log_level,
            encoding=encoding
        )
        logging.INFO(f"Log file created at: {log_file}")

    except Exception as e:
        logging.error(f"Error configuring log: {e}") # Catching any errors that may occur

def configura_outlook():
    """
    Configures and initializes Outlook.

    This function initializes a new instance of Outlook.

    Log messages are recorded to indicate whether the initialization was successful 
    or if any errors occurred.

    Returns:
        outlook: Outlook application object if initialization is successful, 
                 otherwise, returns None.
    """
    # Environment validation
    if not os.name == 'nt':
        logging.error("This function is supported only on Windows systems.")
        return None

    # Initialize Outlook
    try:
        pythoncom.CoInitialize()  # Initializes COM
        outlook = win32.Dispatch('outlook.application')
        logging.info("Outlook started successfully.")
        return outlook
    except Exception as e:
        logging.error(f"Error initializing Outlook: {e}")
        return None


def formata_valores(valor):
    """
    Formats a numeric value as a string in monetary format.

    Parameters:
    valor: The numeric value to be formatted.

    Returns:
    str: The value formatted as a string with two decimal places and thousand separators.
    """
    return f'{valor:,.2f}'


def define_cores(valor_1, valor_2):
    """
    Defines the color based on the comparison between two values.

    Parameters:
    valor_1: The first value to be compared.
    valor_2: The second value to be compared.

    Returns:
    str: 'green' if valor_1 is greater than or equal to valor_2; 'red' otherwise.
    """
    return 'green' if valor_1 >= valor_2 else 'red'

def carrega_dfs(caminho=None):
    """
    Loads DataFrames from Excel and CSV files.

    This function attempts to read three files: an Excel file with sales data, 
    an Excel file with email addresses, and a CSV file with store data. 
    If any of the files are not found, the function catches the exception and 
    logs an error message.

    Parameters:
    caminho (dict): A dictionary with the file paths to be loaded.
                    If None, uses the default paths.

    Returns:
        tuple: A tuple containing three DataFrames (sales, emails, stores).
               If there is an error loading the files, returns (None, None, None).
    """
    # Defining the file paths

    if caminho is None:
        caminhos = {
            'vendas': r'Bases de Dados/Sales.xlsx',
            'emails': r'Bases de Dados/Emails.xlsx',
            'lojas': r'Bases de Dados/Stores.csv'
        }
        
    # Initializing dictionary to store DataFrames
    dataframes = {}

    # Trying to load the DataFrames
    for nome, caminho in caminhos.items():
        if not os.path.exists(caminho):
            logging.error(f'File not found: {caminho}')
            return None, None, None
        
        try:
            if nome in ['vendas', 'emails']:
                dataframes[nome] = pd.read_excel(caminho, engine='openpyxl')
            elif nome == 'lojas':
                dataframes[nome] = pd.read_csv(caminho, encoding='latin1', sep=';')
        except Exception as e:
            logging.error(f'Error loading the file {caminho}: {e}')
            return None, None, None

    return dataframes['vendas'], dataframes['emails'], dataframes['lojas']


def merge_dfs(sales, stores, coluna_referencia):
    """
    Merges two DataFrames based on a reference column.

    This function combines the `sales` and `stores` DataFrames using the 
    specified column as the reference. If one of the DataFrames is empty 
    or if the reference column does not exist in one of the DataFrames, an 
    error message is logged, and the function returns None.

    Args:
        sales (DataFrame): DataFrame containing the sales data.
        stores (DataFrame): DataFrame containing the store data.
        coluna_referencia (str): Name of the column by which the DataFrames should be merged.

    Returns:
        DataFrame: The resulting merged DataFrame, or None in case of error.
    
    Raises:
        ValueError: If the arguments are not valid DataFrames or if the reference column does not exist.
    """

    # DataFrame validation
    if not isinstance(sales, pd.DataFrame) or not isinstance(stores, pd.DataFrame):
        logging.error('Error: One or more arguments are not valid DataFrames.')
        return None

    if sales.empty or stores.empty:
        logging.error('Error: One or more DataFrames are empty.')
        return None

    # Checking the reference column
    if coluna_referencia not in sales.columns:
        logging.error(f"Error: the column '{coluna_referencia}' does not exist in the sales DataFrame.")
        return None

    if coluna_referencia not in stores.columns:
        logging.error(f"Error: the column '{coluna_referencia}' does not exist in the stores DataFrame.")
        return None

    # Attempting to merge the DataFrames
    try:
        merged_df = sales.merge(stores, on=coluna_referencia)
    except Exception as e:
        logging.error(f'Unexpected error during merging: {e}')
        return None
    
    return merged_df


def cria_tabelas_para_lojas(stores, sales, coluna_referencia):
    """
    Creates a dictionary of DataFrames for each store.

    This function checks if the `stores` and `sales` DataFrames are empty and if the reference 
    column exists in both. Then, it creates a dictionary where each key is the name of a store, 
    and the value is a DataFrame containing the sales corresponding to that store.

    Args:
        stores (DataFrame): DataFrame containing the store data.
        sales (DataFrame): DataFrame containing the sales data.
        coluna_referencia (str): Name of the column by which sales will be grouped by store.

    Returns:
        dict: A dictionary where the keys are the store names and the values are DataFrames 
              with the corresponding sales. Returns an empty dictionary in case of error or 
              if the validations are not met.
    """

    # DataFrame validation
    if not isinstance(stores, pd.DataFrame) or not isinstance(sales, pd.DataFrame):
        logging.error('Error: One or more arguments are not valid DataFrames.')
        return {}
    
    if stores.empty or sales.empty:
        logging.error('Error: One or more DataFrames are empty.')
        return {}

    # Checking the reference column
    if coluna_referencia not in stores.columns:
        logging.error(f"Error: The column '{coluna_referencia}' does not exist in the 'stores' DataFrame.")
        return {}
    
    if coluna_referencia not in sales.columns:
        logging.error(f"Error: The column '{coluna_referencia}' does not exist in the 'sales' DataFrame.")
        return {}

    dict_lojas = {}
    try:
        for loja in stores[coluna_referencia].unique():  # Use unique() to avoid duplication
            dict_lojas[loja] = sales[sales[coluna_referencia] == loja]
    except Exception as e:
        logging.error(f'Unexpected error while creating tables for stores: {e}')
        return {}
    
    return dict_lojas


def cria_indicador_dia(sales):
    """
    Creates an indicator representing the most recent day in the sales data.

    This function checks if the `sales` DataFrame is empty, and then 
    identifies the most recent day present in the 'Date' column. If the column does not exist, 
    or if any other error occurs during processing, an error message is logged.

    Args:
        sales (DataFrame): DataFrame containing the sales data, including a 'Date' column.

    Returns:
        datetime: The most recent day found in the 'Date' column, or None in case of error.
    
    Raises:
        ValueError: If the 'Date' column cannot be converted to datetime type.
    """

    # DataFrame validation
    if sales is None or sales.empty:
        logging.error("Error: The 'sales' DataFrame is empty.")
        return None
    
    # Checking the 'Date' column
    if 'Date' not in sales.columns:
        logging.error("Error: The 'Date' column does not exist in the 'sales' DataFrame.")
        return None

    try:
        # Ensuring the 'Date' column is of datetime type
        sales['Date'] = pd.to_datetime(sales['Date'], errors='coerce')
        dia_indicador = sales['Date'].max()
        
        # Check if dia_indicador is a valid value
        if pd.isnull(dia_indicador):
            logging.error("Error: Could not determine a valid date from the 'Date' column.")
            return None

    except Exception as e:
        logging.error(f"Unexpected error while creating day indicator: {e}")
        return None

    return dia_indicador


def pega_caminho_backup():
    """
    Retrieves the backup directory path.

    This function checks if the 'Backup Arquivos Lojas' directory exists. 
    If the directory does not exist, it creates the folder.

    Returns:
        Path: The Path object representing the backup directory.
    """
    caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

    # Checking the existence of the directory
    if not caminho_backup.is_dir():
        try:
            # Creates the new folder if it doesn't exist
            caminho_backup.mkdir(parents=True, exist_ok=True)
            logging.info(f"Folder: '{caminho_backup}' created successfully.")
        except Exception as e:
            logging.error(f"Error creating the folder '{caminho_backup}': {e}")
            return None

    return caminho_backup


def lista_nomes_doBackup(caminho_backup):
    """
    Lists the filenames in the backup directory.

    This function checks if the backup path is valid, if the directory exists, 
    and if it is a valid directory. It then returns a list with the filenames 
    in the directory. If there is any error, an error message is logged.

    Args:
        caminho_backup (Path): The backup directory path.

    Returns:
        list: A list containing the filenames in the backup directory, or an empty list in case of error.
    """
    # Checking if the backup path is None
    if caminho_backup is None:
        logging.error("Error: The backup path is None.")
        return []
    
    # Checking if the directory exists
    if not caminho_backup.exists():
        logging.error(f"Error: The directory '{caminho_backup}' does not exist.")
        return []
    
    # Checking if it is a valid directory
    if not caminho_backup.is_dir():
        logging.error(f"Error: '{caminho_backup}' is not a valid directory.")
        return []
    
    try:
        # Listing the files in the directory
        arquivos_pasta_backup = caminho_backup.iterdir()
        lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup if arquivo.is_file()]
    except Exception as e:
        logging.error(f'Unexpected error while listing files: {e}')
        return []
    
    return lista_nomes_backup


def cria_pastas_para_lojas(dict_lojas, lista_nomes_backup, caminho_backup):
    """
    Creates folders for stores in the backup directory.

    This function checks if the backup path is valid, and then 
    creates a new folder for each store that does not already exist in the backup directory. 
    If an error occurs during folder creation, an error message is logged.

    Args:
        dict_lojas (dict): A dictionary where the keys are store names.
        lista_nomes_backup (list): A list containing the names of existing folders in the backup directory.
        caminho_backup (Path): The backup directory path where the folders will be created.

    Returns:
        None: The function does not return anything but logs messages about the folder creation status.
    """
    # Checking the backup path
    if caminho_backup is None:
        logging.error("Error: The backup path is None.")
        return
    
    if not caminho_backup.exists():
        logging.error(f"Error: The directory '{caminho_backup}' does not exist.")
        return

    if not caminho_backup.is_dir():
        logging.error(f"Error: '{caminho_backup}' is not a valid directory.")
        return

    # Creating folders for the stores
    for loja in dict_lojas:
        if loja not in lista_nomes_backup:
            nova_pasta = caminho_backup / loja
            try:
                # Create the new folder if it does not exist
                nova_pasta.mkdir(parents=True, exist_ok=True)
                logging.info(f"Folder: '{nova_pasta}' created successfully.")
            except Exception as e:
                logging.error(f"Error creating the folder '{nova_pasta}': {e}")


def salva_excel_para_cada_loja(dict_lojas, dia_indicador, caminho_backup):
    """
    Saves an Excel file for each store in the backup directory.

    This function checks if the backup path is valid, and then saves an Excel file 
    for each store in the `dict_lojas` dictionary. The filename is based on the indicator 
    date and the store name. If the store's DataFrame is empty, an error message is logged, 
    and the file is not saved. Any errors during the saving process are also handled.

    Args:
        dict_lojas (dict): A dictionary where the keys are store names 
                           and the values are DataFrames with the corresponding sales.
        dia_indicador (datetime): The indicator date that will be used to name the files.
        caminho_backup (Path): The backup directory path where the files will be saved.

    Returns:
        None: The function does not return anything but logs messages about the file saving status.
    """
    # Checking the backup path
    if caminho_backup is None:
        logging.error("Error: The backup path is None.")
        return
    
    if not caminho_backup.exists() or not caminho_backup.is_dir():
        logging.error(f"Error: The directory '{caminho_backup}' does not exist or is not a valid directory.")
        return

    # Saving Excel files for each store
    for loja, df in dict_lojas.items():
        nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
        local_arquivo = caminho_backup / loja / nome_arquivo
        
        try:
            if df.empty:
                logging.error(f"Error: The DataFrame for store '{loja}' is empty. It will not be saved.")
                continue
            
            # Ensure the store directory exists
            (caminho_backup / loja).mkdir(parents=True, exist_ok=True)
            
            # Save DataFrame as an Excel file
            df.to_excel(local_arquivo, index=False)
            logging.info(f"File '{local_arquivo}' saved successfully.")
        
        except Exception as e:
            logging.error(f"Error saving the file '{local_arquivo}': {e}")

def metas():
    """
    Define and returns the annual and daily revenue and product goals.

    The function creates two dictionaries: one for the annual goals and another for the daily goals.
    In case of an error during the creation of the dictionaries, an error log is generated.

    Returns:
    tuple: A tuple containing two dictionaries:
        - dict_metas_ano: Annual goals (revenue, product quantity, and average ticket).
        - dict_metas_dia: Daily goals (revenue, product quantity, and average ticket).
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
        logging.error(f'Error: Could not create goals: {e}')
        return {}, {}  # Returns empty dictionaries in case of an error


def calcula_indicadores(dict_lojas, dia_indicador):
    """
    Calculates sales indicators for multiple stores on a specific day and for the year.

    For each store in the provided dictionary, the function calculates revenue, product diversity, 
    and the average ticket, both for the year and for a specific day. In case of an error during the calculation,
    error log messages are generated.

    Args:
        dict_lojas (dict): A dictionary where each key is a store's name, and the value is a DataFrame 
                           containing sales data, including the columns 'Date', 'Total Value', and 'Product'.
        dia_indicador (str): The date in 'YYYY-MM-DD' format for which daily indicators 
                             should be calculated.

    Returns:
        dict: A dictionary with the indicators for each store, containing:
            - 'indicadores_ano': Another dictionary with total revenue, product quantity, and average ticket for the year.
            - 'indicadores_dia': Another dictionary with total revenue, product quantity, and average ticket for the day.
    
    Logs:
        Registers error messages if the DataFrames are empty or if the expected columns do not exist.
    """
    indicadores_todas_lojas = {}

    # Calculating indicators
    for loja, vendas_loja in dict_lojas.items():
        if vendas_loja.empty:
            logging.error(f"Error: The DataFrame for the store '{loja}' is empty. Indicators not calculated.")
            continue

        vendas_loja_dia = vendas_loja.loc[vendas_loja['Date'] == dia_indicador]

        # Revenue
        faturamento_ano = vendas_loja['Total Value'].sum()
        faturamento_dia = vendas_loja_dia['Total Value'].sum() if not vendas_loja_dia.empty else 0

        # Product diversity
        qtd_produtos_ano = len(vendas_loja['Product'].unique())
        qtd_produtos_dia = len(vendas_loja_dia['Product'].unique()) if not vendas_loja_dia.empty else 0

        # Average ticket year
        try:
            valor_venda = vendas_loja.groupby('Sale Code').sum(numeric_only=True)
            ticket_medio_ano = valor_venda['Total Value'].mean() if not valor_venda.empty else 0
        except KeyError:
            logging.error(f"Error: The column 'Total Value' does not exist for the store '{loja}'.")
            ticket_medio_ano = 0
        
        # Average ticket day
        try:
            valor_venda_dia = vendas_loja_dia.groupby('Sale Code').sum(numeric_only=True)
            ticket_medio_dia = valor_venda_dia['Total Value'].mean() if not valor_venda_dia.empty else 0
        except KeyError:
            logging.error(f"Error: The column 'Total Value' does not exist for the store '{loja}' on day {dia_indicador}.")
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
    Sends an email with sales performance indicators for each store, based on a specific day.

    The function calculates sales goals and indicators for the stores and sends a personalized 
    email to each store's manager, including revenue, product quantity, and average ticket data, 
    both for the day and for the year.

    Args:
        outlook (object): An Outlook object for creating and sending emails.
        dict_lojas (dict): A dictionary where each key is a store's name, and the value is a DataFrame 
                           containing sales data, including 'Date', 'Total Value', and 'Product' columns.
        dia_indicador (datetime): The date for which the indicators should be calculated.
        caminho_backup (str): The path to the directory where the backup spreadsheets are stored.
        emails (DataFrame): A DataFrame containing store information, managers, and email addresses.

    Returns:
        None: The function does not return a value but sends emails and logs the success or failure of sending.

    Logs:
        Logs information about the success or failure of sending emails, as well as errors in calculating goals and indicators.
    """

    try:
        # unpacking the goals 
        dict_metas_ano, dict_metas_dia = metas()

        # annual goals
        meta_faturamento_ano = formata_valores(dict_metas_ano['meta_faturamento_ano'])
        meta_qtd_produtos_ano = dict_metas_ano['meta_qtd_produtos_ano']
        meta_ticketMedio_ano = formata_valores(dict_metas_ano['meta_ticketMedio_ano'])

        # daily goals
        meta_faturamento_dia = formata_valores(dict_metas_dia['meta_faturamento_dia'])
        meta_qtd_produtos_dia = dict_metas_dia['meta_qtd_produtos_dia']
        meta_ticketMedio_dia = formata_valores(dict_metas_dia['meta_ticketMedio_dia'])

        # unpacking indicators
        indicadores_todas_lojas = calcula_indicadores(dict_lojas, dia_indicador)    
    except Exception as e:
        logging.error(f'Error: Could not obtain goals and/or indicators: {e}')
        return None
    
    for nome_loja in indicadores_todas_lojas.keys():
        try:
            # Checking if the store exists in the DataFrame
            if nome_loja in emails['Store'].values:
                nome = emails.loc[emails['Store'] == nome_loja, 'Manager'].values[0]  
                email_destino = emails.loc[emails['Store'] == nome_loja, 'E-mail'].values[0]

                # Validations
                if not nome or not email_destino:
                    logging.error(f"Invalid data for store '{nome_loja}': Manager or email not found.")
                    continue

                mail = outlook.CreateItem(0)
                mail.To = email_destino
                mail.Subject = f'OnePage Day {dia_indicador.day}/{dia_indicador.month} - Store {nome_loja}'

                # getting the year's and day's indicators for the current store
                indicadores_ano = indicadores_todas_lojas[nome_loja]['indicadores_ano']
                indicadores_dia = indicadores_todas_lojas[nome_loja]['indicadores_dia']

                # Checking if the indicators are present
                if 'faturamento_ano' not in indicadores_ano or 'faturamento_dia' not in indicadores_dia:
                    logging.error(f"Missing indicators for store '{nome_loja}'.")
                    continue

                # yearly indicators
                faturamento_ano = formata_valores(indicadores_ano['faturamento_ano'])
                qtd_produtos_ano = indicadores_ano['qtd_produtos_ano']
                ticket_medio_ano = formata_valores(indicadores_ano['ticket_medio_ano'])

                # daily indicators
                faturamento_dia = formata_valores(indicadores_dia['faturamento_dia'])
                qtd_produtos_dia = indicadores_dia['qtd_produtos_dia']
                ticket_medio_dia = formata_valores(indicadores_dia['ticket_medio_dia'])

                # defining the colors of the current store's indicators
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
                <h1 style="color: #007BFF;">Good morning, {nome}</h1>

                <p style="line-height: 1.6;">The result from yesterday <strong>({dia_indicador.day}/{dia_indicador.month})</strong> at <strong>Store {nome_loja}</strong> was:</p>

                <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fff; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);">
                    <tr>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd; background-color: #007BFF; color: #fff;">Indicator</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Value Day</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Goal Day</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Scenario Day</th>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Revenue</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{faturamento_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_faturamento_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_fat_dia};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Product Diversity</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{qtd_produtos_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_qtd_produtos_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_qtde_dia};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Average Ticket</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{ticket_medio_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_ticketMedio_dia}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_ticket_dia};">◙</td>
                    </tr>
                </table>

                <table style="width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fff; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);">
                    <tr>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd; background-color: #007BFF; color: #fff;">Indicator</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Value Year</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Goal Year</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">Scenario Year</th>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Revenue</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{faturamento_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_faturamento_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_fat_ano};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Product Diversity</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{qtd_produtos_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_qtd_produtos_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_qtde_ano};">◙</td>
                    </tr>
                    <tr>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">Average Ticket</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{ticket_medio_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd;">{meta_ticketMedio_ano}</td>
                        <td style="padding: 12px; text-align: center; border: 1px solid #ddd; color: {cor_ticket_ano};">◙</td>
                    </tr>
                </table>

                <p style="line-height: 1.6;">Attached is the spreadsheet with all the data for further details.</p>
                <p style="line-height: 1.6;">If you have any questions, feel free to reach out.</p>

                <p style="margin-top: 20px; font-size: 0.9em; color: #666;">Best regards,<br>Lucas</p>
            </body>
            </html>
        
'''
                
                attachment = pathlib.Path.cwd() / caminho_backup / nome_loja / f'{dia_indicador.month}_{dia_indicador.day}_{nome_loja}.xlsx'
                # It's important to use str() to avoid errors
                mail.Attachments.Add(str(attachment))

                mail.Send()
                logging.info(f'Email for Store {nome_loja} sent to {nome} - {email_destino}')

                # Wait to avoid conflicts
                time.sleep(0.2)

        except Exception as e:
            logging.error(f"Error sending email for store '{nome_loja}': {e}")


def cria_rankings(dia_indicador, sales):
    """
    Creates annual and daily revenue rankings for stores and saves them to Excel files.

    This function groups the sales data (sales) by store to calculate the total annual revenue
    and the revenue for the specified day. The results are saved in two separate Excel files.

    Args:
        dia_indicador (datetime): The day for which the daily ranking is calculated.
        sales (DataFrame): A DataFrame containing sales data, including columns
                           for 'Store', 'Total Value', and 'Date'.

    Returns:
        tuple: A tuple containing two DataFrames:
            - faturamento_lojas_ano (DataFrame): Annual ranking of stores based on total revenue.
            - faturamento_lojas_dia (DataFrame): Daily ranking of stores based on revenue for the specified day.

    Raises:
        KeyError: If the columns 'Store', 'Total Value', or 'Date' are not found in the sales DataFrame.
        Exception: Any other errors that occur during processing.

    Logs:
        Logs errors related to missing necessary columns and warns if no sales are found for the specified day.
    """
    try:
        # Initial validations
        required_columns = ['Store', 'Total Value', 'Date']
        for column in required_columns:
            if column not in sales.columns:
                logging.error(f"Error: The column '{column}' was not found in the sales DataFrame.")
                return None, None

        # Annual ranking
        faturamento_lojas = sales.groupby('Store')[['Total Value']].sum(numeric_only=True)
        faturamento_lojas_ano = faturamento_lojas.sort_values(by='Total Value', ascending=False)
        nome_arquivo_ano = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Anual.xlsx'
        faturamento_lojas_ano.to_excel(fr"Backup Arquivos Lojas/{nome_arquivo_ano}")  # Saving annual ranking

        # Daily ranking
        vendas_dia = sales.loc[sales['Date'] == dia_indicador]
        if vendas_dia.empty:
            logging.warning(f"No sales found for the day {dia_indicador.strftime('%d/%m/%Y')}.")
            return faturamento_lojas_ano, pd.DataFrame(columns=['Store', 'Total Value'])

        faturamento_lojas_dia = vendas_dia.groupby('Store')[['Total Value']].sum(numeric_only=True)
        faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Total Value', ascending=False)
        nome_arquivo_dia = f'{dia_indicador.month}_{dia_indicador.day}_Ranking_Dia.xlsx'
        faturamento_lojas_dia.to_excel(fr"Backup Arquivos Lojas/{nome_arquivo_dia}")  # Saving daily ranking

        return faturamento_lojas_ano, faturamento_lojas_dia

    except Exception as e:
        logging.error(f"Unexpected error when creating ranking: {e}")
        return None, None

def email_diretoria(outlook, faturamento_lojas_ano, faturamento_lojas_dia, dia_indicador, emails, caminho_backup):
    """
    Sends an email to the management team with the daily and annual revenue results.

    This function creates an HTML-formatted email containing the best and worst store performances 
    in terms of revenue, in addition to attaching the annual and daily ranking files.

    Args:
        outlook (object): An Outlook object to create and send emails.
        faturamento_lojas_ano (DataFrame): DataFrame containing the annual store ranking based on revenue.
        faturamento_lojas_dia (DataFrame): DataFrame containing the daily store ranking based on daily revenue.
        dia_indicador (datetime): The day for which the results are reported.
        emails (DataFrame): DataFrame containing email information for stores, including the 'E-mail' column.
        caminho_backup (str): The path where the ranking files are saved.

    Raises:
        IndexError: If the necessary data cannot be found in the emails DataFrame.
        FileNotFoundError: If the ranking files are not found in the specified path.
        Exception: Any other error that occurs during the email sending process.

    Returns:
        None: The function does not return a value, but sends an email and generates logs regarding the success or failure of the sending.

    Logs:
        Logs errors related to missing email data and ranking files.
    """

    try:
        # Check if email data for the management team exists
        if 'CEO' not in emails['Store'].values:
            logging.error("Error: Management email data not found.")
            return
        
        destinatario = emails.loc[emails['Store'] == 'CEO', 'E-mail'].values[0]
        
        # Validate the revenue DataFrames
        if faturamento_lojas_ano.empty or faturamento_lojas_dia.empty:
            logging.error("Error: One or both revenue DataFrames are empty.")
            return

        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = f'Revenue Ranking for {dia_indicador.day}/{dia_indicador.month}'
        
        # Format the body of the email
        melhor_loja_dia = faturamento_lojas_dia.index[0]
        pior_loja_dia = faturamento_lojas_dia.index[-1]
        melhor_loja_ano = faturamento_lojas_ano.index[0]
        pior_loja_ano = faturamento_lojas_ano.index[-1]
        
        mail.HTMLBody = f'''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <title>Revenue Report</title>
</head>
<body>
    <h2>Dear all, good morning</h2>

    <h3>Results for the Day:</h3>
    <p><strong>Best store of the day in revenue:</strong> Store <span style="font-weight: bold;">{melhor_loja_dia}</span> with revenue R$<span style="font-weight: bold;">{faturamento_lojas_dia.iloc[0, 0]:,.2f}</span></p>
    <p><strong>Worst store of the day in revenue:</strong> Store <span style="font-weight: bold;">{pior_loja_dia}</span> with revenue R$<span style="font-weight: bold;">{faturamento_lojas_dia.iloc[-1, 0]:,.2f}</span></p>

    <h3>Results for the Year:</h3>
    <p><strong>Best store of the year in revenue:</strong> Store <span style="font-weight: bold;">{melhor_loja_ano}</span> with revenue R$<span style="font-weight: bold;">{faturamento_lojas_ano.iloc[0, 0]:,.2f}</span></p>
    <p><strong>Worst store of the year in revenue:</strong> Store <span style="font-weight: bold;">{pior_loja_ano}</span> with revenue R$<span style="font-weight: bold;">{faturamento_lojas_ano.iloc[-1, 0]:,.2f}</span></p>

    <p>The annual and daily rankings for all stores are attached.</p>

    <p>If you have any questions, feel free to reach out.</p>

    <p>Best regards,<br>Lucas</p>
</body>
</html>
'''

        # Attaching files with validation
        for tipo in ['Ranking_Anual', 'Ranking_Dia']:
            attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_{tipo}.xlsx'
            if not attachment.exists():
                logging.error(f"Error: File {attachment} not found.")
                continue  # Skip to the next file

            mail.Attachments.Add(str(attachment))

        mail.Send()
        logging.info('Management email sent successfully.')
    
    except IndexError as e:
        logging.error(f"Error: Could not find data for sending the email. Detail: {e}")
    except FileNotFoundError as e:
        logging.error(f"Error: One or more ranking files not found. Detail: {e}")
    except Exception as e:
        logging.error(f"Unexpected error while sending email: {e}")


def main():
    """
    Main function that runs the sales analysis workflow.

    This function performs the following steps:
    1. Loads the sales, emails, and store information DataFrames.
    2. Merges the sales and store DataFrames using the store identifier.
    3. Creates a dictionary with the information of each store.
    4. Generates a daily indicator based on sales.
    5. Obtains the backup path.
    6. Lists the backup file names.
    7. Creates folders for each store in the backup directory.
    8. Saves an Excel file containing the sales for each store.
    9. Sends emails with the sales report to store managers.
    10. Creates revenue rankings for the year and the day.
    11. Sends an email with the rankings to the management team.

    Raises:
        FileNotFoundError: If a required file is not found during the process.
        KeyError: If there is an error accessing data in the DataFrames.
        Exception: Raises any other unexpected errors during the execution of the program.

    Returns:
        None: The function does not return a value, but performs actions such as sending emails and generating logs about the process.
    """
    try:
        # Configure logging
        configura_log()

        # Set up Outlook
        outlook = configura_outlook()

        # Load the DataFrames
        vendas, emails, lojas = carrega_dfs()

        # Merge the sales and store DataFrames
        vendas = merge_dfs(vendas, lojas, 'Store ID')

        # Create a dictionary with the information of each store
        dict_lojas = cria_tabelas_para_lojas(lojas, vendas, 'Store')

        # Create the daily indicator
        dia_indicador = cria_indicador_dia(vendas)

        # Get the path for the backup
        caminho_backup = pega_caminho_backup()

        # List the backup file names
        lista_nomes_backup = lista_nomes_doBackup(caminho_backup)

        # Create folders for each store in the backup directory
        cria_pastas_para_lojas(dict_lojas, lista_nomes_backup, caminho_backup)

        # Save an Excel file for each store
        salva_excel_para_cada_loja(dict_lojas, dia_indicador, caminho_backup)

        # Send emails with the sales reports
        envia_email(outlook, dict_lojas, dia_indicador, caminho_backup, emails)

        # Create revenue rankings
        faturamento_lojas_ano, faturamento_lojas_dia = cria_rankings(dia_indicador, vendas)

        # Wait to avoid conflicts
        time.sleep(0.3)

        # Send the email with the rankings to the management team
        email_diretoria(outlook, faturamento_lojas_ano, faturamento_lojas_dia, dia_indicador, emails, caminho_backup)

    except FileNotFoundError as e:
        logging.error(f"File not found: {e}")
    except KeyError as e:
        logging.error(f"Error accessing data: {e}")
    except Exception as e:
        logging.error(f"An error occurred during the program execution: {e}")


# Run the main function
try:
    if __name__ == "__main__":
        main()
except Exception as e:
    logging.critical(f"An error occurred during the program execution: {e}")
