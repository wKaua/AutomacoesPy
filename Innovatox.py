import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from bs4 import BeautifulSoup

# Função para extrair informações de uma div específica
def extrair_informacoes_div(div):
    # Extrai o conteúdo do div
    row_content = div.get_attribute('outerHTML')

    # Usa BeautifulSoup para processar o conteúdo extraído
    soup = BeautifulSoup(row_content, 'html.parser')

    # Extrai as informações desejadas
    laboratorio = soup.find('div', class_='card-title').text.strip()
    endereco_parts = soup.find('p', style='min-height: 130px;').contents
    endereco = endereco_parts[0].strip()
    bairro = endereco_parts[2].strip().split(' - ')[0]
    preco = soup.find('h4').text.strip().split('R$')[1].strip()

    # Tenta extrair o telefone, se presente
    telefone = ''
    cep = ''
    for part in endereco_parts:
        if 'Telefone' in part:
            telefone = part.strip().split('Telefone ')[1]
        if 'CEP' in part:
            cep = part.strip().split('CEP ')[1]

    # Retorna um dicionário com as informações extraídas formatadas
    return {
        'Laboratório': laboratorio,
        'Endereço': endereco,
        'Bairro': bairro,
        'CEP': cep,
        'Telefone': telefone,
        'Preço': preco
    }

# Inicializa o WebDriver
driver = webdriver.Chrome()

# Abre a página
driver.get("https://itox.com.br/vendas/0/CNH")

# Espera a página carregar
time.sleep(3)

# Lê os dados da planilha Excel
df_excel = pd.read_excel(r'D:\Projetos\DataTox\dadosconsult\inovatox.xlsx')

# Variável para armazenar o estado atualmente selecionado
estado_atual = None

# Itera sobre os dados da planilha Excel
for index, row in df_excel.iterrows():
    estado = row['estado']
    cidade = row['cidade']

    # Verifica se o estado atual é diferente do estado a ser selecionado
    if estado != estado_atual:
        # Clica no dropdown para selecionar estado
        dropdown_estado = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-home-layout/div/div/div[3]/div/div/app-cadastro-vendas/div/form/div/div[1]/div/div[1]/div/div/div[1]/div[2]/ng-select/div'))
        )
        dropdown_estado.click()

        # Espera as opções aparecerem e clica no estado atual
        option_estado = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, f'//span[contains(text(), "{estado}")]'))
        )
        option_estado.click()

        # Atualiza o estado atual selecionado
        estado_atual = estado

    # Espera o input estar disponível e clica nele
    input_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//input[@aria-autocomplete="list"]'))
    )
    input_element.click()

    # Clica no dropdown para selecionar a cidade
    dropdown_cidade = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/app-home-layout/div/div/div[3]/div/div/app-cadastro-vendas/div/form/div/div[1]/div/div[1]/div/div/div[2]/div[2]/ng-select/div'))
    )
    dropdown_cidade.click()

    # Espera as opções aparecerem e clica na cidade atual
    option_city = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, f'//span[contains(text(), "{cidade}")]'))
    )
    option_city.click()

    # Espera todas as divs com a classe "card gutter-b" estarem disponíveis
    card_divs = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.XPATH, '//div[@class="card gutter-b"]'))
    )

    # Inicializa uma lista para armazenar todos os dados extraídos
    todos_os_dados = []

    # Itera sobre todas as divs encontradas e extrai as informações de cada uma delas
    for card_div in card_divs:
        data = extrair_informacoes_div(card_div)
        data['Cidade'] = cidade  # Adiciona a cidade ao dicionário de dados
        todos_os_dados.append(data)

    # Converte a lista de dados em um DataFrame pandas
    df = pd.DataFrame(todos_os_dados)

    # Imprime as informações extraídas formatadas
    for idx, row_data in df.iterrows():
        print(f"Estado: {estado}")
        print(f"Cidade: {cidade}")
        print(f"Laboratório: {row_data['Laboratório']}")
        print(f"Endereço: {row_data['Endereço']}")
        print(f"Bairro: {row_data['Bairro']}")
        print(f"CEP: {row_data['CEP']}")
        print(f"Telefone: {row_data['Telefone']}")
        print(f"Preço: {row_data['Preço']}")
        print()

    # Pausa para evitar problemas de carregamento
    time.sleep(2)

# Fecha o WebDriver
driver.quit()

