import pandas as pd
import pyodbc

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

cols = ['nome', 'email', 'aniversario', 'admissao', 'Nome', 'Aniversario-de-vida', 'Aniversario-Tempo-Empresa', 'STATUS']

type_cols = {
    'nome': str,
    'email': str,
    'aniversario': str,
    'admissao': str,
    'Nome': str,
    'Aniversario-de-vida': str,
    'Aniversario-Tempo-Empresa': str,
    'STATUS': str
}

df = pd.read_excel('Disparo_imagem.xlsx', usecols=cols, dtype=type_cols, engine='openpyxl')

df['aniversario'] = pd.to_datetime(df['aniversario'], errors='coerce', infer_datetime_format=True)
df['admissao'] = pd.to_datetime(df['admissao'], errors='coerce', infer_datetime_format=True)

df['aniversario'] = df['aniversario'].dt.strftime('%Y-%m-%d')
df['admissao'] = df['admissao'].dt.strftime('%Y-%m-%d')

SERVER = ''
DATABASE = ''
USERNAME = ''
PASSWORD = ''

connection_string = f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};'

conn = pyodbc.connect(connection_string)

table = ''

insert_query = f"""
INSERT INTO {table} (NOME, EMAIL, ANIVERSARIO, ADMISSAO, NOME_UPPER, ANIVERSARIO_VIDA, ANIVERSARIO_EMPRESA, STATUS)
VALUES (?, ?, ?, ?, ?, ?, ?, ?)
"""

cursor = conn.cursor()

for index, row in df.iterrows():
    nome = row['nome']
    email = row['email']
    nome_upper = row['Nome']
    aniversario = row['aniversario']
    admissao = row['admissao']
    aniversario_vida = str(row['Aniversario-de-vida']) if pd.notna(row['Aniversario-de-vida']) else None
    aniversario_empresa = str(row['Aniversario-Tempo-Empresa']) if pd.notna(row['Aniversario-Tempo-Empresa']) else None
    status = str(row['STATUS']) if pd.notna(row['STATUS']) else None

    print(f"Inserindo: {nome}, {email}, {aniversario}, {admissao}, {nome_upper}, {aniversario_vida}, {aniversario_empresa}, {status}")
    
    cursor.execute(insert_query, nome, email, aniversario, admissao, nome_upper, aniversario_vida, aniversario_empresa, status)

conn.commit()
conn.close()

print(f'Dados inseridos na tabela {table} com sucesso.')
