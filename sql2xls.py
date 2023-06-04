import pandas as pd
import sqlite3

# Funzione per estrarre i dati di una tabella e restituire un DataFrame
def extract_table_data(connection, table_name):
    query = f"SELECT * FROM {table_name}"
    df = pd.read_sql_query(query, connection)
    return df

# Funzione per salvare i dati in un file Excel
def save_to_excel(dataframes, output_file):
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    for table_name, df in dataframes.items():
        df.to_excel(writer, sheet_name=table_name, index=False)
    writer.save()

# Connessione al database SQL
connection = sqlite3.connect('database.db')  # Sostituisci con il percorso del tuo file SQL

# Estrai i nomi delle tabelle dal database
cursor = connection.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
table_names = cursor.fetchall()
table_names = [name[0] for name in table_names]

# Estrai i dati di tutte le tabelle
dataframes = {}
for table_name in table_names:
    df = extract_table_data(connection, table_name)
    dataframes[table_name] = df

# Chiudi la connessione al database
connection.close()

# Salva i dati in un file Excel
save_to_excel(dataframes, 'output.xlsx')  # Sostituisci con il percorso e nome del file Excel di output desiderato
