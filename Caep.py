# Código faz a busca das informações e salva no mesmo arquivo 

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from bs4 import BeautifulSoup
import pandas as pd
import os

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
nav = webdriver.Chrome(options=options)
nav.get("https://caeptox.com.br/comprar-exame/motorista")

wait = WebDriverWait(nav, 10)

def obter_opcoes(xpath_drop):
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_drop)))
    dropdown.click()
    opcoes = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//*[@role="option"]')))
    lista_opcoes = [opcao.text.strip() for opcao in opcoes if opcao.text.strip()]
    time.sleep(1)
    return lista_opcoes

def extract_lab_info(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    lab_info_divs = soup.find_all('div', class_='lab-info')
    lab_price_divs = soup.find_all('div', class_='lab-price')
    lab_infos = []
    
    for lab_info_div, lab_price_div in zip(lab_info_divs, lab_price_divs):
        try:
            lab_name = lab_info_div.find('h4', class_='lab-name').text.strip()
            address = lab_info_div.find_all('p', class_='mat-line')[0].text.strip()            
            city_state = lab_info_div.find_all('p', class_='mat-line')[1].text.strip()
            city = city_state[:-4]
            state = city_state[-2:]
            phone = lab_info_div.find('a', class_='phone').text.strip()
            price = lab_price_div.find('p', class_='price').text.strip()
        except (IndexError, AttributeError) as e:
            print(f"Erro ao extrair dados: {e}")
            continue
        
        lab_info = {
            'LABORATÓRIO': lab_name,
            'ENDEREÇO': address,
            'CIDADE': city,
            'ESTADO': state,
            'TELEFONE': phone,
            'PREÇO': price,
            'ORIGEM': 'CAEP'
        }
        lab_infos.append(lab_info)
    
    return lab_infos

xpath_estado = '//*[@id="mat-select-0"]/div/div[1]'
lista_estado =  obter_opcoes(xpath_estado)
xpath_cidade = '//*[@id="mat-select-1"]/div/div[1]'
estado_remove = ["Selecione"]
lista_estados = [estado for estado in lista_estado if estado not in estado_remove]
lista_estados = ['Acre']
all_lab_infos = []
folder_path = r"C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Diretoria\Joao Ruiz\Py\Codigos_Dados_Laboratorios\Dados_Laboratórios.xlsx"

nav.refresh()

for estado in lista_estados:
    try:
        element_estado = WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mat-select-0"]')))
        nav.execute_script("arguments[0].click();", element_estado)
        nav.find_element(By.XPATH, f'//mat-option/span[contains(text(),"{estado}")]').click()
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao selecionar estado: {e}")
        continue

    lista_cidade = obter_opcoes(xpath_cidade)
    time.sleep(5)
    print(f'\n{estado}\nOpções de cidades: {lista_cidade}')
    primeira_cidade = lista_cidade[0]
    
    for cidade in lista_cidade:
        print(f'Cidade: {cidade}')
        try:
            if cidade != primeira_cidade:
                nav.find_element(By.XPATH, '//*[@id="mat-select-1"]/div/div[1]/span').click()
                nav.find_element(By.XPATH, f'//mat-option/span[contains(text(),"{cidade}")]').click()
                time.sleep(3)
            else:
                nav.find_element(By.XPATH, f'//mat-option/span[contains(text(),"{cidade}")]').click()
                time.sleep(3)
        except Exception as e:
            print(f"Erro ao selecionar cidade: {e}")
            continue
        html_content = nav.page_source
        lab_infos = extract_lab_info(html_content)
        all_lab_infos.extend(lab_infos)

df_novos_dados = pd.DataFrame(all_lab_infos)
try:
    df_existente = pd.read_excel(folder_path, engine='openpyxl')
    df_combinado = pd.concat([df_existente, df_novos_dados], ignore_index=True)
except FileNotFoundError:
    df_combinado = df_novos_dados
df_combinado.to_excel(folder_path, index=False, engine='openpyxl')

print(f'Dados extraídos e salvos em {folder_path}')

    