import re
import pandas as pd
from docx import Document
import requests
from bs4 import BeautifulSoup
import time


# Função para extrair nomes das revistas de referências no Word
def extract_journals_from_docx(file_path):
    document = Document(file_path)
    journals = []

    for paragraph in document.paragraphs:
        text = paragraph.text
        # Regex para capturar o nome da revista dentro da referência
        match = re.search(r"(?:\.\s)?([A-Za-z&.,\s]+)(?:, v\.|$)", text)  # Ajustado para capturar nomes de revistas
        if match:
            journals.append(match.group(1).strip())

    return list(set(journals))  # Remove duplicados


def get_scimago_ranking(journal_name):
    base_url = "https://www.scimagojr.com/journalrank.php"
    query = '+'.join(journal_name.split())
    url = base_url + query

    print(f"Pesquisando por: {journal_name}")
    print(f"URL: {url}")
    
    response = requests.get(url)
    if response.status_code != 200:
        print(f"Erro ao acessar SCImago para: {journal_name}")
        return None, None

    soup = BeautifulSoup(response.text, "html.parser")
    
    # Verificar se há resultados de busca
    results = soup.find_all("div", class_="search_results")  # Procurar resultados
    if not results:
        print(f"Nenhum resultado encontrado para: {journal_name}")
        return None, None

    # Obter o primeiro link de revista
    first_result = results[0].find("a")
    if not first_result:
        print(f"Nenhum detalhe encontrado para: {journal_name}")
        return None, None

    # Acessar página de detalhes da revista
    details_url = "https://www.scimagojr.com/" + first_result['href']
    print(f"URL de detalhes: {details_url}")

    details_response = requests.get(details_url)
    if details_response.status_code != 200:
        print(f"Erro ao acessar detalhes para: {journal_name}")
        return None, None

    details_soup = BeautifulSoup(details_response.text, "html.parser")

    try:
        # Buscar SJR Indicator e Quartile
        sjr_value = details_soup.find("div", class_="cell", text="SJR indicator").find_next_sibling("div").text.strip()
        quartile = details_soup.find("div", class_="cell", text="Quartile").find_next_sibling("div").text.strip()
        return sjr_value, quartile
    except AttributeError:
        print(f"Erro ao extrair informações para: {journal_name}")
        return None, None


# Função principal
def main(docx_file):
    # Extrair nomes das revistas do arquivo .docx
    journals = extract_journals_from_docx(docx_file)
    data = []

    # Buscar rankings no SCImago
    for journal in journals:
        sjr, quartile = get_scimago_ranking(journal)
        data.append({
            "Journal Name": journal,
            "SJR Indicator": sjr,
            "Quartile": quartile
        })
        time.sleep(2)  # Aguarde 2 segundos entre as requisições para evitar bloqueios

    # Salvar os resultados em um arquivo CSV
    df = pd.DataFrame(data)
    df.to_csv("journal_rankings.csv", index=False)
    print("Resultados salvos em 'journal_rankings.csv'.")


# Executar o script
#if __name__ == "__main__":
#    main("referencias.docx")  # Substitua pelo nome do seu arquivo Word


get_scimago_ranking("Quarterly Journal of Economics")