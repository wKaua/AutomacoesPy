import re
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

file = r"C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Diretoria\Joao Ruiz\Py\Codigos_Dados_Laboratorios\Dados_Laboratórios.xlsx"
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
nav = webdriver.Chrome(options=options)
nav.get("https://exames.contraprova.com.br/buscacep?category=3")
element = WebDriverWait(nav, 4).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="top"]/body/div/div[5]'))
)
nav.execute_script("arguments[0].parentNode.removeChild(arguments[0]);", element)

OpcEstado = WebDriverWait(nav, 4).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="myApp"]/section/div/select[1]'))
)
html_content = nav.page_source
nav.quit()
pattern = r'var marcadores = (\[.*?\]);'
match = re.search(pattern, html_content, re.DOTALL)
if match:
    marcadores_json = match.group(1)
    marcadores = json.loads(marcadores_json)    
    print("Variável 'marcadores' extraída e salva em 'marcadores.json'")
else:
    print("Variável 'marcadores' não encontrada no conteúdo HTML.")
dados_processados = []
for item in marcadores:
    laboratorio = item.get('titulo', '')
    endereco = f"{item.get('rua', '')}"
    bairro = item.get('bairro', '')
    cidade = item.get('cidade', '')
    estado = item.get('estado', '')
    telefone = item.get('telefone', '')
    preco = item.get('preco', '')
    cep = ''  
    email = ''  
    dados_processados.append({
        'LABORATÓRIO': laboratorio,
        'ENDEREÇO': endereco,
        'NÚMERO': item.get('numero', ''),
        'BAIRRO': bairro,
        'CIDADE': cidade,
        'ESTADO': estado,
        'TELEFONE': telefone,
        'PREÇO': preco,
        'CEP': cep,
        'EMAIL': email,
        'ORIGEM': 'CONTRAPROVA'
    })
df_novos_dados = pd.DataFrame(dados_processados)
try:
    df_existente = pd.read_excel(file, engine='openpyxl')
    df_combinado = pd.concat([df_existente, df_novos_dados], ignore_index=True)
except FileNotFoundError:
    df_combinado = df_novos_dados
df_combinado.to_excel(file, index=False, engine='openpyxl')
print('Dados extraídos e salvos com sucesso!')
