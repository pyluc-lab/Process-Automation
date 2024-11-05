# Sales Analysis

This project performs sales analysis, automating processes and generating performance reports. The system uses three Excel database files and generates output files organized in specific directories.

## Overview

It uses three Excel files located in the **Bases de Dados** folder:

- **Sales**: Contains sales data from the stores.
- **Emails**: Includes contact information for the store managers.
- **Stores**: Contains information about the stores.

***IMPORTANT*** : In the " **Emails** " file, replace the values in the " **E-mail** " column with an email address that you can use to test the script.

After the script runs, the generated files are saved in the **Backup Arquivos Lojas** folder. For each store, a subfolder will be created (if it does not already exist), and within it, a `.xlsx` file containing the store's information. Additionally, two other files will be saved in the same folder:

- **Ranking Anual**: A file that contains the annual sales ranking of the stores.
- **Ranking do Dia**: A file with the sales ranking for the day in analysis.

## Outlook

It is essential to have the Outlook application installed on your computer and open it as an administrator when running the script.

## Prerequisites

Before executing the code, make sure you have the following files and dependencies:

1. **Data Files**:

   - Three Excel files in the **Bases de Dados** folder:
     - A file with sales data, including the columns:
       - `Store ID`
       - `Total Value`
       - `Date`
     - A file with email information for the stores, including:
       - `Store`
       - `Manager`
       - `Email`
     - A file with information about the stores, including:
       - `Store ID`
       - `Store`
2. **Backup Directory**:

   - The **Backup Arquivos Lojas** folder will be automatically created by the script to store the reports.
3. **Dependencies**:

   - Make sure you have the following Python libraries installed:
     - `pandas`
     - `openpyxl` (for handling Excel files)
     - `pywin32` (for interacting with Outlook)
     - `logging`

## How to Run

1. Make sure you have Python installed on your system.
2. Clone this repository or download the files.
3. Open a terminal and navigate to the project directory.
4. Run the main file with the command:
   ```bash
   python filename.py
   ```

## Code Structure

The code is organized into the following main functions:

* `main()`: The main function that executes the workflow of the sales analysis program.
* `envia_email()`: Sends personalized emails to the store managers.
* `cria_rankings()`: Creates annual and daily sales rankings for the stores.
* `email_diretoria()`: Sends an email to management with the sales results
