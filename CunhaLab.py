import requests
from bs4 import BeautifulSoup
import csv
import os

# Cabeçalho para simular um navegador
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

def get_links_from_divs(url, div_class):
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    soup = BeautifulSoup(response.content, 'html.parser')
    divs = soup.find_all('div', class_=div_class)
    links = [a['href'] for div in divs for a in div.find_all('a', href=True)]
    return links

def get_lab_info(url):
    response = requests.get(url, headers=HEADERS)
    response.raise_for_status()
    soup = BeautifulSoup(response.content, 'html.parser')
    
    lab_name_tag = soup.find('h1', class_='product_title entry-title')
    price_tag = soup.find('span', class_='woocommerce-Price-amount amount')
    address_tag = soup.find('div', class_='woocommerce-product-details__short-description')
    address_tag2 = soup.find('div', class_='woocommerce-Tabs-panel woocommerce-Tabs-panel--description panel entry-content wc-tab')
    
    lab_name = lab_name_tag.get_text(strip=True) if lab_name_tag else 'N/A'
    price = price_tag.get_text(strip=True) if price_tag else 'N/A'
    address = address_tag.get_text(strip=True) if address_tag else 'N/A'
    address2 = address_tag2.get_text(strip=True) if address_tag2 else 'N/A'
    
    return {
        'Laboratório': lab_name,
        'Preço': price,
        'Endereço': address,
        'Endereço2': address2
    }

def main():
    main_url = 'https://cunhalab.com.br/rede-coleta/'
    category_links = get_links_from_divs(main_url, 'single-product-category')

    all_lab_info = []

    for link in category_links:
        sub_links = get_links_from_divs(link, 'nv-card-content-wrapper')
        
        for sub_link in sub_links:
            info = get_lab_info(sub_link)
            all_lab_info.append(info)
    
    # Caminho para salvar o arquivo CSV
    csv_file_path = r"C:\Users\wkasouto\OneDrive\OneDrive - Laboratorio Morales LTDA\Diretoria\Joao Ruiz\Controladoria\Dados Laboratórios\CunhaLab\CunhaLab_Teste.xlsx"
    
    # Verifica se o diretório existe, se não, cria o diretório
    os.makedirs(os.path.dirname(csv_file_path), exist_ok=True)

    # Escreve os dados em um arquivo CSV
    with open(csv_file_path, mode='w', newline='', encoding='utf-8') as csv_file:
        fieldnames = ['Laboratório', 'Preço', 'Endereço', 'Endereço2']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        
        writer.writeheader()
        writer.writerows(all_lab_info)
    
    print(f'Dados salvos em: {csv_file_path}')

if __name__ == '__main__':
    main()
