import sys
import pandas as pd
import sqlite3
from datetime import datetime

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
    writer._save()

# Ottieni il nome del file SQL dall'argomento di input
if len(sys.argv) < 2:
    print("Usage: python sql_to_excel.py <input_sql_file>")
    sys.exit(1)

sql_file = sys.argv[1]

# Connessione al database SQL
connection = sqlite3.connect(sql_file)

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

# Genera il nome del file Excel di output con la libreria datetime
output_file = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".xlsx"

# Salva i dati in un file Excel
save_to_excel(dataframes, output_file)

print(f"I dati sono stati salvati nel file Excel: {output_file}")
