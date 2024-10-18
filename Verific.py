import pandas as pd
from datetime import datetime

file_path = r'C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Diretoria\Joao Ruiz\Controladoria\Dados Laboratórios\Testes\Base Teste Comparativo.xlsx'
df = pd.read_excel(file_path)
data_atual = datetime.today().strftime('%Y-%m-%d')
lista_laboratorios = ['lab142','lab96', 'lab74', 'lab316']
df_novos = pd.DataFrame({
    'LABORATÓRIOS': lista_laboratorios,
    'ENDEREÇO': [None] * len(lista_laboratorios),  
    'ORIGEM': [None] * len(lista_laboratorios),    
    'NÚMERO': [None] * len(lista_laboratorios),
    'BAIRRO': [None] * len(lista_laboratorios),
    'CEP': [None] * len(lista_laboratorios),
    'DATA_CRIAÇÃO': [None] * len(lista_laboratorios),
    'DATA_ATUALIZAÇÃO': [None] * len(lista_laboratorios),
    'DATA_EXCLUSÃO': [None] * len(lista_laboratorios),
    'CIDADE': [None] * len(lista_laboratorios),
    'ESTADO': [None] * len(lista_laboratorios),
    'TELEFONE': [None] * len(lista_laboratorios),
    'TELEFONE2': [None] * len(lista_laboratorios),
    'PREÇO': [None] * len(lista_laboratorios),
    'EMAIL': [None] * len(lista_laboratorios),
    'SITUAÇÃO': [None] * len(lista_laboratorios)})
df['DATA_ATUALIZAÇÃO'] = data_atual
df_novos['DATA_ATUALIZAÇÃO'] =  data_atual 
df_existentes = df[df['LABORATÓRIOS'].isin(lista_laboratorios)]
df_novos = df_novos[~df_novos['LABORATÓRIOS'].isin(df_existentes['LABORATÓRIOS'])]
df.loc[df['LABORATÓRIOS'].isin(lista_laboratorios), 'DATA_ATUALIZAÇÃO'] = data_atual
df_novos['DATA_CRIAÇÃO'] = data_atual
df_atualizado = pd.concat([df, df_novos], ignore_index=True)
df_atualizado['chave_unica'] = (df_atualizado['LABORATÓRIOS'].astype(str) + '-' +
                                df_atualizado['ENDEREÇO'].fillna('').astype(str) + '-' +
                                df_atualizado['ORIGEM'].fillna('').astype(str))
unidades_excluidas = df_atualizado[~df_atualizado['LABORATÓRIOS'].isin(lista_laboratorios)]
df_atualizado.loc[unidades_excluidas.index, 'DATA_EXCLUSÃO'] = data_atual
df_atualizado['DUPLICADO'] = df_atualizado.duplicated(subset=['LABORATÓRIOS', 'ENDEREÇO'], keep=False)
df_atualizado = df_atualizado.drop(columns=['chave_unica'])
df_atualizado.to_excel(file_path, index=False)
print("Unidades Excluídas:")
print(unidades_excluidas[['LABORATÓRIOS', 'ENDEREÇO', 'DATA_EXCLUSÃO']])
print("\nLaboratórios Duplicados:\n")
print(df_atualizado[['LABORATÓRIOS', 'ENDEREÇO', 'DUPLICADO']])
