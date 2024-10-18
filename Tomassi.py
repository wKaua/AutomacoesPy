import csv
import requests
from bs4 import BeautifulSoup

# URL da página que contém os links
url = "https://tommasi.com.br/laboratorios/"

# Cabeçalhos da requisição HTTP para evitar bloqueio por agentes de usuário
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

# Fazer a requisição HTTP para obter o conteúdo da página
response = requests.get(url, headers=headers)
if response.status_code == 200:
    page_content = response.content

    # Analisar o conteúdo HTML com BeautifulSoup
    soup = BeautifulSoup(page_content, 'html.parser')

    # Encontrar todas as <div> com a classe "box-unidades"
    div_elements = soup.find_all('div', class_='box-unidades')

    if div_elements:
        # Lista para armazenar os links coletados
        link_urls = []

        # Iterar sobre cada <div> e coletar links dos <ul> dentro dela
        for div_element in div_elements:
            ul_elements = div_element.find_all('ul', class_='list-unidade')
            for ul in ul_elements:
                links = ul.find_all('a', href=True)
                for link in links:
                    link_urls.append(link['href'])

        # Lista para armazenar os dados dos laboratórios
        lab_data = []

        # Iterar sobre cada link e coletar as informações desejadas
        for link_url in link_urls:
            lab_response = requests.get(link_url, headers=headers)
            if lab_response.status_code == 200:
                lab_page_content = lab_response.content
                lab_soup = BeautifulSoup(lab_page_content, 'html.parser')

                # Extrair o nome do laboratório
                lab_name = lab_soup.find('h1', class_='elementor-heading-title elementor-size-default')
                if lab_name:
                    lab_name = lab_name.text.strip()
                else:
                    lab_name = "Nome do laboratório não encontrado"

                # Extrair os detalhes do laboratório
                lab_details = lab_soup.find('div', class_='detail-unidade')
                if lab_details:
                    lab_details = lab_details.text.strip()
                else:
                    lab_details = "Detalhes do laboratório não encontrados"

                # Adicionar os dados coletados à lista
                lab_data.append([lab_name, lab_details])
            else:
                print(f"Falha ao acessar a página do laboratório. Status code: {lab_response.status_code}")

        # Caminho para salvar o arquivo CSV
        csv_file_path = r"C:\Users\jhsruiz\OneDrive - Laboratorio Morales LTDA\Área de Trabalho\Marcus\Inteligência de Mercado\Dados\laboratorios.csv"

        # Escrever os dados coletados no arquivo CSV
        with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["Laboratório", "Detalhes"])  # Cabeçalho do CSV
            writer.writerows(lab_data)

        print(f"Dados salvos no arquivo CSV em: {csv_file_path}")

    else:
        print("Não foi possível encontrar as <div> com a classe 'box-unidades'.")
else:
    print(f"Falha ao acessar a página. Status code: {response.status_code}")
