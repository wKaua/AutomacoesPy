import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

# Dados do site
DOMAIN = 'https://www.exametoxicologico.com.br/'
URLS = [
    'exames-toxicologicos-acre/',
    'exames-toxicologicos-alagoas/',
    'exames-toxicologicos-amapa/',
    'exames-toxicologicos-amazonas/',
    'exames-toxicologicos-bahia/',
    'exames-toxicologicos-ceara/',
    'exames-toxicologicos-distrito-federal/',
    'exames-toxicologicos-espirito-santo/',
    'exames-toxicologicos-goias/',
    'exames-toxicologicos-maranhao/',
    'exames-toxicologicos-mato-grosso/',
    'exames-toxicologicos-mato-grosso-do-sul/',
    'exames-toxicologicos-minas-gerais/',
    'exames-toxicologicos-para/',
    'exames-toxicologicos-paraiba/',
    'exames-toxicologicos-parana/',
    'exames-toxicologicos-pernambuco/',
    'exames-toxicologicos-piaui/',
    'exames-toxicologicos-rio-de-janeiro/',
    'exames-toxicologicos-rio-grande-do-norte/',
    'exames-toxicologicos-rio-grande-do-sul/',
    'exames-toxicologicos-rondonia/',
    'exames-toxicologicos-roraima/',
    'exames-toxicologicos-santa-catarina/',
    'exames-toxicologicos-sao-paulo/',
    'exames-toxicologicos-sergipe/',
    'exames-toxicologicos-tocantins/'
]
HEADERS = {'User-Agent': 'Mozilla/5.0'}
PARAMETERS = {}

# Função para obter o conteúdo da página
def get_page_content(url, headers, parameters):
    response = requests.get(url, headers=headers, params=parameters)
    response.raise_for_status()  # Verifica se a requisição foi bem sucedida
    soup = BeautifulSoup(response.text, 'html.parser')
    return soup

# Função para extrair o valor numérico do texto do preço
def extrair_preco(preco_text):
    match = re.search(r'R\$ (\d+,\d+)', preco_text)
    if match:
        return float(match.group(1).replace(',', '.'))  # Convertendo para float
    else:
        return None

# Cria uma nova planilha Excel
workbook = Workbook()
sheet = workbook.active
sheet.title = "Laboratorios"

# Adiciona os cabeçalhos das colunas
sheet.append([
    "Laboratorio", "Email", "Endereco", "Bairro", "Cidade",
    "Estado", "CEP", "Origem", "Data", "Preço", "Telefone"
])

# Data da execução do código
data_execucao = datetime.now().strftime("%d/%m/%Y")

def process_url(url):
    try:
        soup = get_page_content(url, HEADERS, PARAMETERS)
        links = soup.find_all('a', class_='city__link main-color')
        links_list = [link.get('href') for link in links]
        return links_list
    except Exception as e:
        print(f"Erro ao processar a URL {url}: {e}")
        return []

# Usando ThreadPoolExecutor para realizar as requisições em paralelo
with ThreadPoolExecutor(max_workers=10) as executor:
    future_to_url = {executor.submit(process_url, f"{DOMAIN}{url_path}"): url_path for url_path in URLS}
    all_links = []
    for future in as_completed(future_to_url):
        url_path = future_to_url[future]
        try:
            links = future.result()
            all_links.extend(links)
        except Exception as e:
            print(f"Erro ao processar a URL {url_path}: {e}")

def process_laboratory(link):
    try:
        pagina = get_page_content(link, HEADERS, None)

        # Encontra todas as tags <a> com a classe especificada
        laboratory_links = pagina.find_all('a', class_='laboratory__title__link main-color')

        # Encontra todos os endereços, cidades, telefones e preços na página
        enderecos = pagina.find_all('p', class_='laboratory-address__content__item public-place')
        cidades = pagina.find_all('p', class_='laboratory-address__content__item city-state')
        telefones = pagina.find_all('a', class_='telephone__link laboratory-telephone-link main-color')
        precos = pagina.find_all('p', class_='laboratory__purchase__price-desc')

        results = []
        for i, laboratory_link in enumerate(laboratory_links):
            href = laboratory_link.get('href')
            texto = laboratory_link.text.strip()
            endereco_text = enderecos[i].text.strip() if i < len(enderecos) else "N/A"
            cidade_text = cidades[i].text.strip() if i < len(cidades) else "N/A"
            telefone_text = telefones[i].text.strip() if i < len(telefones) else "N/A"
            preco_text = extrair_preco(precos[i].text.strip()) if i < len(precos) else None
            
            # Extrai o estado do texto da cidade
            estado_text = cidade_text.split('-')[-1].strip() if '-' in cidade_text else ""
            
            # Extrai o bairro do texto do endereço
            bairro_text = endereco_text.split('-')[-1].strip() if '-' in endereco_text else ""

            results.append([
                texto, "", endereco_text.split('-')[0].strip(), bairro_text, cidade_text.split('-')[0].strip(), estado_text,
                "", "Pardini", data_execucao, preco_text, telefone_text
            ])
        return results
    except Exception as e:
        print(f"Erro ao processar o laboratório {link}: {e}")
        return []

# Usando ThreadPoolExecutor para processar laboratórios em paralelo
with ThreadPoolExecutor(max_workers=10) as executor:
    future_to_link = {executor.submit(process_laboratory, link): link for link in all_links}
    for future in as_completed(future_to_link):
        link = future_to_link[future]
        try:
            results = future.result()
            for result in results:
                sheet.append(result)
        except Exception as e:
            print(f"Erro ao processar o laboratório {link}: {e}")

# Formata a coluna de Preço como número
for cell in sheet['J']:
    cell.number_format = '#,##0.00'

# Salva a planilha Excel
workbook.save("dados_pardini.xlsx")
