import re
import pandas as pd
from docx import Document
import requests
from bs4 import BeautifulSoup

# Função para extrair nomes das revistas de referências no Word
def extract_journals_from_docx(file_path):
    document = Document(file_path)
    journals = []

    for paragraph in document.paragraphs:
        text = paragraph.text
        match = re.search(r"([A-Za-z&., ]+?), v\.", text)  # Padrão para nomes de revistas
        if match:
            journals.append(match.group(1).strip())

    return list(set(journals))  # Remove duplicados

# Função para buscar o ranking no SCImago
def get_scimago_ranking(journal_name):
    base_url = "https://www.scimagojr.com/journalsearch.php?q="
    query = '+'.join(journal_name.split())
    url = base_url + query

    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")
    results = soup.find("div", {"class": "search_results"})

    if results:
        journal_info = results.find("a")  # Primeira entrada da pesquisa
        if journal_info:
            details_url = "https://www.scimagojr.com/" + journal_info['href']
            details_response = requests.get(details_url)
            details_soup = BeautifulSoup(details_response.text, "html.parser")

            # Buscar o SJR Indicator e o Quartil
            sjr_value = details_soup.find("div", text="SJR indicator").find_next_sibling("div").text
            quartile = details_soup.find("div", text="Quartile").find_next_sibling("div").text
            return sjr_value, quartile

    return None, None  # Caso não encontre a revista

# Função principal
def main(docx_file):
    # Extrair nomes das revistas
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

    # Salvar em um arquivo CSV
    df = pd.DataFrame(data)
    df.to_csv("journal_rankings.csv", index=False)
    print("Resultados salvos em 'journal_rankings.csv'.")

# Executar o script
if __name__ == "__main__":
    main("referencias.docx")  # Substitua pelo nome do seu arquivo Word
