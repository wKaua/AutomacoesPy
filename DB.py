from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
import re

file = r"C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Diretoria\Joao Ruiz\Py\Codigos_Dados_Laboratorios\Dados_Laboratórios.xlsx"

estado_siglas = {
    'Acre': 'AC',
    'Alagoas': 'AL',
    'Amapá': 'AP',
    'Amazonas': 'AM',
    'Bahia': 'BA',
    'Ceará': 'CE',
    'Distrito Federal': 'DF',
    'Espírito Santo': 'ES',
    'Goiás': 'GO',
    'Maranhão': 'MA',
    'Mato Grosso': 'MT',
    'Mato Grosso do Sul': 'MS',
    'Minas Gerais': 'MG',
    'Pará': 'PA',
    'Paraíba': 'PB',
    'Paraná': 'PR',
    'Pernambuco': 'PE',
    'Piauí': 'PI',
    'Rio de Janeiro': 'RJ',
    'Rio Grande do Norte': 'RN',
    'Rio Grande do Sul': 'RS',
    'Rondônia': 'RO',
    'Roraima': 'RR',
    'Santa Catarina': 'SC',
    'São Paulo': 'SP',
    'Sergipe': 'SE',
    'Tocantins': 'TO'
}
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
nav = webdriver.Chrome(options=options)
url_page = "https://dbtoxicologico.com.br/redes-de-coletas?"
nav.get(url_page)

wait = WebDriverWait(nav, 10)

def obter_opcoes(xpath_opcao):
    opcoes = []
    try:      
        select_element = nav.find_element(By.XPATH, xpath_opcao)
        options = select_element.find_elements(By.TAG_NAME, 'option')
        for option in options:
            opcoes.append(option.text)
    except Exception as e:
        print(f"Erro ao obter opções")
    return opcoes

xpath_estado = '//*[@id="address-state"]'
xpath_cidade = '//*[@id="address-city"]'

lista_estado = obter_opcoes(xpath_estado)
estado_remove = ["Selecione o estado"]
lista_estado = [estado for estado in lista_estado if estado not in estado_remove]
lista_estado = ['Alagoas']
print(f'Lista de estados:\n{lista_estado}')

estado_cidade_dict = {}

for estado in lista_estado:
    sigla_estado = estado_siglas.get(estado, estado)
    nav.find_element(By.XPATH, xpath_estado).click()
    nav.find_element(By.XPATH, f'//select[@id="address-state"]/option[text()="{estado}"]').click()
    time.sleep(1)
    lista_cidade = obter_opcoes(xpath_cidade)
    estado_cidade_dict[sigla_estado] = lista_cidade
nav.quit()

estados = []
cidades = []

for estado, lista_cidade in estado_cidade_dict.items():
    for cidade in lista_cidade:
        estados.append(estado)
        cidades.append(cidade)

df = pd.DataFrame({
    'Estado': estados,
    'Cidade': cidades
})

print(f'Terminou a primeira parte do código\nExibindo a planilha:\n{df}')

print(f'Começando a segunda parte do código....')
HEADERS = {'User-Agent': 'Mozilla/5.0'}
resultados = []
def separar_endereco(endereco):
    match = re.match(r'^(.*),\s*(\d+)\.\s*Bairro:\s*(.*)$', endereco)
    if match:
        rua = match.group(1).strip()
        numero = match.group(2).strip() if match.group(2) else 'N/A'
        bairro = match.group(3).strip()
    else:
        rua = endereco
        numero = 'N/A'
        bairro = 'N/A'
    return rua, numero, bairro

def buscar_dados_cidade(estado, cidade):
    base_url = f"https://www.dbtoxicologico.com.br/redes-de-coletas?estado={estado}&cidade={{}}"
    cidade_formatada = quote(cidade, safe='')
    url = base_url.format(cidade_formatada)
    response = requests.get(url, headers=HEADERS)
    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        ul_box_networks = soup.find("ul", class_="box-networks")
        if ul_box_networks:
            for item in ul_box_networks.find_all("li"):
                dados = item.text.strip().split('\n')

                unidade = 'N/A'
                telefone = 'N/A'
                email = 'N/A'
                endereco = 'N/A'

                if len(dados) > 0:
                    unidade = dados[0]  
                if len(dados) > 1:
                    telefone = dados[1]  
                if len(dados) > 2:
                    email = dados[2]  
                if len(dados) > 3:
                    endereco = dados[3]  
                rua, numero, bairro = separar_endereco(endereco)

                resultado = {
                    'LABORATÓRIO': unidade,
                    'ENDEREÇO': rua,
                    'NÚMERO': numero,
                    'BAIRRO': bairro,
                    'CIDADE': cidade,
                    'ESTADO': estado,
                    'TELEFONE':telefone,
                    'EMAIL': email,
                    'ORIGEM' : 'DB'
                }
                resultados.append(resultado)
        else:
            resultado = {
                'LABORATÓRIO': 'N/A',
                'ENDEREÇO': 'N/A',
                'NÚMERO': 'N/A',
                'BAIRRO': 'N/A',
                'CIDADE': cidade,
                'ESTADO': estado,
                'TELEFONE': 'Não foi possível encontrar informações',
                'EMAIL':'N/A',
                'ORIGEM' : 'DB'
            }
            resultados.append(resultado)
    else:
        resultado = {        
            'LABORATÓRIO': 'N/A',
            'ENDEREÇO': 'N/A',
            'NÚMERO': 'N/A',
            'BAIRRO': 'N/A',
            'CIDADE': cidade,
            'ESTADO': estado,
            'TELEFONE': f'Erro: {response.status_code}',
            'EMAIL':'N/A',
            'ORIGEM' : 'DB'
        }
        resultados.append(resultado)
for index, row in df.iterrows():
    estado = row['Estado']
    cidade = row['Cidade']
    print(f'Estado: {estado}\nCidade: {cidade}')
    buscar_dados_cidade(estado, cidade)
df_resultados = pd.DataFrame(resultados)
try:
    df_existente = pd.read_excel(file, engine='openpyxl')
    df_combinado = pd.concat([df_resultados, df_existente], ignore_index=True)
except FileNotFoundError:
    df_combinado = df_resultados
df_combinado.to_excel(file, index=False, engine='openpyxl')

print(f"Dados salvos em {file}")

