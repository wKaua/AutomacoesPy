import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains

nav = webdriver.Chrome()
nav.get("https://labet.com.br/")
element = nav.find_element(By.XPATH, '//*[@id="app"]/div/div[1]')
nav.execute_script("arguments[0].parentNode.removeChild(arguments[0]);", element)
wait = WebDriverWait(nav, 10)

EncontarLaboratorio = nav.find_element(By.XPATH, '//*[@id="btnsTemplateContainer"]/div[3]/div/div[1]/a/span')
EncontarLaboratorio.click()
time.sleep(2)

def selecionar_opcao(xpath_dropdown, opcao_text):
    try:
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_dropdown)))
        dropdown.click()
        time.sleep(1)

# Digita a opção
        input_field = ActionChains(nav)
        input_field.move_to_element(dropdown)
        input_field.click()
        input_field.send_keys(opcao_text)
        input_field.perform()
        time.sleep(2)  
# Seleciona a opção
        opcao = wait.until(EC.element_to_be_clickable((By.XPATH, f'//div[@role="option" and contains(., "{opcao_text}")]')))
        opcao.click()
        time.sleep(2)
    except Exception as e:
        scroll_dropdown(xpath_dropdown)
        return False
    return True

def escolher_cidade():
    try:
        EscolherCidades = nav.find_element(By.XPATH, '//*[@id="mapSearch"]/div/div[4]/div[1]/div/div[2]/button')
        EscolherCidades.click()
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao escolher cidade")

def scroll_dropdown(dropdown_xpath):
    try:
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
        dropdown.click()
        time.sleep(1)
        dropdown_menu = wait.until(EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "v-menu__content theme--light menuable__content__active")]')))
        last_height = nav.execute_script("return arguments[0].scrollHeight", dropdown_menu)
        while True:
            nav.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", dropdown_menu)
            time.sleep(1)
            new_height = nav.execute_script("return arguments[0].scrollHeight", dropdown_menu)
            if new_height == last_height:
                break
            last_height = new_height
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao rolar dropdown: {dropdown_xpath}.")

def obter_opcoes(xpath_dropdown):
    scroll_dropdown(xpath_dropdown)
    opcoes = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//div[@role="option"]')))
    return [opcao.text.strip() for opcao in opcoes if opcao.text.strip()]

extracted_data = []

def processar_dados():
    html_content = nav.page_source
    soup = BeautifulSoup(html_content, 'html.parser')
    data_blocks = soup.find_all('div', class_='d-flex flex-column ma-4 px-3 container-cards v-card v-card--flat v-card--link v-sheet theme--light')
    for block in data_blocks:
        lab_name = block.find('p', class_='card-title ma-0').get_text(strip=True)
        address = block.find('p', class_='text-infos').find('span').get_text(strip=True)
        city_state = block.find('p', class_='text-infos').find_all('span')[1].get_text(strip=True)
        price = block.find('div', class_='d-flex').get_text(strip=True).replace('R$', '').replace(',', '.').strip()
        extracted_data.append({
            'Laboratório': lab_name,
            'Endereço': address,
            'Cidade e Estado': city_state,
            'Valor': price
        })

def reiniciar_busca():
    try:
        Nova_Busca = nav.find_element(By.XPATH, '//*[@id="mapSearch"]/div/div[4]/button/span')
        Nova_Busca.click()
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao reiniciar a busca")

complete = False
while not complete:
    try:
        escolher_cidade()
        time.sleep(2)
        lista_estados = obter_opcoes('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[2]/div/div/div')
        estado_remove = ["Acre"]
        lista_estados = [estado for estado in lista_estados if estado not in estado_remove]
        print(f"Estados encontrados: {lista_estados}")
        for estado in lista_estados:
            try:
                print(f"Selecionando estado: {estado}")
                if not selecionar_opcao('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[2]/div/div/div', estado):
                    continue
                lista_cidades = obter_opcoes('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[3]/div/div/div')
                print(f"Lista de cidades do estado {estado}: {lista_cidades}")
                for cidade in lista_cidades:
                    try:
                        print(f"Selecionando cidade: {cidade}")
                        if not selecionar_opcao('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[3]/div/div/div', cidade):
                            continue
                        lista_bairros = obter_opcoes('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[4]/div/div/div')
                        print(f"Lista de bairros para a cidade {cidade}: {lista_bairros}")
                        for bairro in lista_bairros:
                            try:
                                print(f"Selecionando bairro: {bairro}")
                                if not selecionar_opcao('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[4]/div/div/div', bairro):
                                    continue
                                processar_dados()
                            except Exception as e:
                                print(f"Erro ao selecionar o bairro: {bairro}.")
                                continue
                            reiniciar_busca()
                            escolher_cidade()
                    except Exception as e:
                        print(f"Erro ao selecionar a cidade: {cidade}.")
                        scroll_dropdown('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[3]/div/div/div')
                        continue
                reiniciar_busca()
                escolher_cidade()
            except Exception as e:
                print(f"Erro ao selecionar o estado: {estado}.")
                scroll_dropdown('//*[@id="mapSearch"]/div/div[4]/div[2]/div/div[2]/div/div/div')
                continue
        complete = True
    except Exception as e:
        print(f"Erro geral: {e}")
        break

df = pd.DataFrame(extracted_data)
df.to_excel('CidadesLabet.xlsx', index=False)
print("Dados extraídos e salvos com sucesso!")
