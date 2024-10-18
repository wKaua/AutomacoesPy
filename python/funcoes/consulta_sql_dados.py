from sqlalchemy import create_engine
import pandas as pd

def consulta_sql_dados(consulta_sql):
    usuario = 'joao.ruiz'
    senha = 'pSLOCgBgy5tB'
    servidor = 'sodreproducao.public.608eda075a62.database.windows.net,3342'
    banco_de_dados = 'toxicologico'
    conexao_sql = f'mssql+pyodbc://{usuario}:{senha}@{servidor}/{banco_de_dados}?driver=ODBC+Driver+17+for+SQL+Server'
    try:
        engine = create_engine(conexao_sql)
        df = pd.read_sql(consulta_sql, engine)
        return df
    except Exception as e:
        print(f'Ocorreu um erro ao executar a consulta SQL: {e}')
        return None  
