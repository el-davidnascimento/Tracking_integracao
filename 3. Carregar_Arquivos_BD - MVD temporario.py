import psycopg2
import pandas as pd


################# TRAKING ###############################

# Caminho do arquivo Excel
caminho_arquivo = r'G:\Meu Drive\Dados\Traking Integracao\8. Integração\8.2. Integração Jardim América\8.2.3. Bases\8.2.3.3. Consolidado das Bases\consolidado - temporario.xlsx'

# Conexão com o banco de dados PostgreSQL
conn = psycopg2.connect(
    host="localhost",
    port=5433,
    database="Multiverso",
    user="postgres",
    password="Multiverso@Educa"
)

# Ler os dados do arquivo Excel
dados_excel = pd.read_excel(caminho_arquivo)

# Cursor
cur = conn.cursor()

nome_tabela = 'traking_integracao'

# Iterar sobre as linhas do DataFrame e inserir os valores na tabela do PostgreSQL
for _, linha in dados_excel.iterrows():
    valores = tuple(linha)
    cur.execute("INSERT INTO traking_integracao VALUES ( %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)", valores)


# Executar o comando DELETE para remover registros duplicados
cur.execute("""
    DELETE FROM traking_integracao
    WHERE (atividade, processo, "status_de_realizaÇÃo", Setor, "data") NOT IN (
      SELECT atividade, processo, "status_de_realizaÇÃo", Setor, MAX("data")
      FROM traking_integracao
      GROUP BY atividade, processo, "status_de_realizaÇÃo", Setor
    )
""")


# Commit e fechamento da conexão
conn.commit()
cur.close()
conn.close()