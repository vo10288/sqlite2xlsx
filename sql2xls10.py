import sys
import pandas as pd
from pysqlite3 import dbapi2 as sqlite3
#import sqlite3
from datetime import datetime

# Funzione per estrarre i dati di una tabella e restituire un DataFrame
def extract_table_data(connection, table_name):
    query = f"SELECT * FROM {table_name}"
    df = pd.read_sql_query(query, connection)
    return df

# Funzione per ottenere i nomi dei campi di una tabella
def get_table_fields(connection, table_name):
    cursor = connection.cursor()
    cursor.execute(f"PRAGMA table_info({table_name})")
    fields = cursor.fetchall()
    field_names = [field[1] for field in fields]
    return field_names

# Funzione per salvare i dati in un file Excel
def save_to_excel(dataframes, output_file):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    for table_name, data in dataframes.items():
        df = data['data']
        fields = data['fields']
        df.to_excel(writer, sheet_name=table_name, index=False)
        worksheet = writer.sheets[table_name]
        for i, field in enumerate(fields):
            worksheet.write(0, i, field)
    writer.close()

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

# Estrai i dati e i nomi dei campi di tutte le tabelle
dataframes = {}
for table_name in table_names:
    df = extract_table_data(connection, table_name)
    fields = get_table_fields(connection, table_name)
    dataframes[table_name] = {'data': df, 'fields': fields}

# Chiudi la connessione al database
connection.close()

# Genera il nome del file Excel di output con la libreria datetime
output_file = datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".xlsx"

# Salva i dati in un file Excel con fogli di lavoro separati per ogni tabella
save_to_excel(dataframes, output_file)

# Stampa i dati estratti
for table_name, data in dataframes.items():
    print(f"Tabella: {table_name}")
    print(f"Campi: {', '.join(data['fields'])}")
    print(data['data'])
    print()

print(f"I dati sono stati salvati nel file Excel: {output_file}")
