import pandas as pd
import pyodbc

cols = []

type_cols = {

}

df = pd.read_excel('Disparo_imagem.xlsx', usecols=cols, dtype=type_cols, engine='openpyxl')

df[''] = pd.to_datetime(df[''], errors='coerce', infer_datetime_format=True)
df[''] = pd.to_datetime(df[''], errors='coerce', infer_datetime_format=True)

df[''] = df[''].dt.strftime('%Y-%m-%d')
df[''] = df[''].dt.strftime('%Y-%m-%d')

SERVER = ''
DATABASE = ''
USERNAME = ''
PASSWORD = ''

connection_string = f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};'

conn = pyodbc.connect(connection_string)

table = ''

insert_query = f"""
INSERT INTO {table} ()
VALUES (?, ?, ?, ?, ?, ?, ?, ?)
"""

cursor = conn.cursor()

for index, row in df.iterrows():
    nome = row['']
    email = row['']
    nome_upper = row['']
    aniversario = row['']
    admissao = row['a']
    aniversario_vida = str(row['']) if pd.notna(row['']) else None
    aniversario_empresa = str(row['']) if pd.notna(row['']) else None
    status = str(row['']) if pd.notna(row['']) else None

    print(f"Inserindo:")
    
    cursor.execute(insert_query, nome, email, aniversario, admissao, nome_upper, aniversario_vida, aniversario_empresa, status)

conn.commit()
conn.close()

print(f'Dados inseridos na tabela {table} com sucesso.')
