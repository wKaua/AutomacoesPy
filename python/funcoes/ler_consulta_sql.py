def ler_consulta_sql(arquivo_sql):
    sql_base = r'C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Felipe Corradi\Sistem\sql\queries\;.sql'
    caminho_sql = sql_base.replace(';', arquivo_sql)
    with open(caminho_sql, 'r', encoding='utf-8') as file:
        return file.read()