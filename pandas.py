import pandas as pd
import pyodbc

# Configurar o pandas para exibir todas as colunas e linhas
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

# Definindo as colunas a serem usadas e seus tipos
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

# Lendo o arquivo Excel
df = pd.read_excel('Disparo_imagem.xlsx', usecols=cols, dtype=type_cols, engine='openpyxl')

# Convertendo colunas de datas para datetime, permitindo formatos mistos e tratando erros
df['aniversario'] = pd.to_datetime(df['aniversario'], errors='coerce', infer_datetime_format=True)
df['admissao'] = pd.to_datetime(df['admissao'], errors='coerce', infer_datetime_format=True)

# Convertendo as colunas de data para o formato aceito pelo SQL Server
df['aniversario'] = df['aniversario'].dt.strftime('%Y-%m-%d')
df['admissao'] = df['admissao'].dt.strftime('%Y-%m-%d')

# Configuração da URL de conexão com o SQL Server
SERVER = ''
DATABASE = ''
USERNAME = ''
PASSWORD = ''

# URL de conexão com o SQL Server
connection_string = f'DRIVER={{SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD};'

# Criar a engine de conexão com o SQL Server
conn = pyodbc.connect(connection_string)

# Nome da tabela onde os dados serão inseridos
table = ''

# Criar a query de inserção (sem a coluna ID)
insert_query = f"""
INSERT INTO {table} (NOME, EMAIL, ANIVERSARIO, ADMISSAO, NOME_UPPER, ANIVERSARIO_VIDA, ANIVERSARIO_EMPRESA, STATUS)
VALUES (?, ?, ?, ?, ?, ?, ?, ?)
"""

# Inserir os dados na tabela SQL Server
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

    # Debugging: Imprimir os valores antes de inserir
    print(f"Inserindo: {nome}, {email}, {aniversario}, {admissao}, {nome_upper}, {aniversario_vida}, {aniversario_empresa}, {status}")
    
    cursor.execute(insert_query, nome, email, aniversario, admissao, nome_upper, aniversario_vida, aniversario_empresa, status)

# Commit as mudanças e fechar a conexão
conn.commit()
conn.close()

print(f'Dados inseridos na tabela {table} com sucesso.')
