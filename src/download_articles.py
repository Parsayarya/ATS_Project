import os
import requests
import pandas as pd

def xlsx_to_csv(xlsx_file: str, csv_file: str) -> None:
    """
    Convert a specified Excel file to CSV format.
    
    :param xlsx_file: The path to the input Excel file.
    :type xlsx_file: str
    :param csv_file: The path to the output CSV file.
    :type csv_file: str
    :return: None
    """
    data_xls = pd.read_excel(xlsx_file, index_col=None)
    data_xls.to_csv(csv_file, encoding='utf-8', index=False)

def download_articles(csv_input_path: str, docx_output_path: str) -> None:
    """
    Download articles listed in a specified CSV file and save them to a specified directory.
    
    :param csv_input_path: The path to the input CSV file containing article URLs.
    :type csv_input_path: str
    :param docx_output_path: The directory path where downloaded articles should be saved.
    :type docx_output_path: str
    :return: None
    """
    df = pd.read_csv(csv_input_path)
    for index, row in df.iterrows():
        url = row['URL']
        response = requests.get(url)
        file_name = url.split('/')[-1]
        with open(f"{docx_output_path}/{file_name}", 'wb') as file:
            file.write(response.content)

if __name__ == "__main__":
    # Example usage:
    # xlsx_to_csv('input.xlsx', 'output.csv')
    # download_articles('output.csv', 'articles/')