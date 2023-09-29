from bs4 import BeautifulSoup
import requests
import pandas as pd
import re


def fetch_articles(keyword, num_pages=10):
    titles = []
    abstracts = []
    authors = []
    years = []

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    for i in range(num_pages):
        start = i * 10
        url = f"https://scholar.google.com/scholar?start={start}&q={keyword}"
        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            print(response.status_code)
            print(f"Failed to retrieve page {i + 1}")
            continue

        soup = BeautifulSoup(response.content, 'html.parser')

        titles += [item.text.replace("HTML", "").strip() for item in soup.find_all('h3', {'class': 'gs_rt'})]
        abstracts += [item.text for item in soup.find_all('div', {'class': 'gs_rs'})]
        for item in soup.find_all('div', {'class': 'gs_a'}):
            text = item.get_text()
            author_year = text.split(' - ')

            if len(author_year) > 1:
                authors.append(author_year[0])
                year = ''.join(filter(str.isdigit, author_year[1]))
                years.append(year if year else 'NA')
            else:
                authors.append('NA')
                years.append('NA')

    data = {'Title': titles, 'Abstract': abstracts, 'Authors': authors, 'Year': years}
    df = pd.DataFrame(data)
    return df


def clean_text(text):
    if pd.isnull(text):
        return text
    # Keeping only alphanumeric characters and spaces
    return ''.join(e for e in text if e.isalnum() or e.isspace())


def clean_dataframe(df):
    for column in df.columns:
        df[column] = df[column].apply(clean_text)
    return df


def extract_year(authors_text):
    if pd.isnull(authors_text):
        return 'NA'
    # Find the last occurrence of a 4-digit number in the authors text
    years = re.findall(r'\b\d{4}\b', authors_text)
    return years[-1] if years else 'NA'


df = fetch_articles(
    "(antarct* OR ''southern ocean'' OR ''ross sea'' OR ''amundsen sea'' OR ''weddell sea'') AND NOT (candida OR ‘‘except antarctica’’ OR ‘‘not antarctica’’)",
    num_pages=10)
df_cleaned = clean_dataframe(df)
df_cleaned['Year'] = df_cleaned['Authors'].apply(extract_year)
df_cleaned.to_csv(r'D:\python\ATS\DownloadCorpusBuilder\files\Data frames\Scholar.csv')
