from src.download_articles import download_articles
from src.convert import convert_pdf_to_docx
from src.corpus_builder import build_corpus
from src.delete_extras import first_level_cleaning
from src.download_science_papers import download_papers

def main():
    """
    Main function to demonstrate the usage of the refactored scripts.
    
    In this function, we import and use the functions from the refactored scripts.
    This serves as an example of how the refactored scripts can be used.
    """
    # Example usage of download_articles
    download_articles('input.csv', 'articles/')
    
    # Example usage of convert_pdf_to_docx
    convert_pdf_to_docx('pdfs/', 'docx/')
    
    # Example usage of build_corpus
    build_corpus('docx/', 'corpus.txt')
    
    # Example usage of first_level_cleaning
    first_level_cleaning('directory/', 10)
    
    # Example usage of download_papers
    download_papers(['paper1_url', 'paper2_url'], 'papers/')

if __name__ == "__main__":
    main()