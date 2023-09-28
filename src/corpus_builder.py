import os
import re
import pandas as pd
import docx
import string

def from_roman_numeral(numeral: str) -> int:
    """
    Convert a Roman numeral to an integer.

    :param numeral: The Roman numeral string to convert.
    :type numeral: str
    :return: The integer equivalent of the Roman numeral.
    :rtype: int
    """
    value_map = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    value = 0
    last_digit_value = 0

    for roman_digit in reversed(numeral):
        digit_value = value_map[roman_digit]
        if digit_value < last_digit_value:
            value -= digit_value
        else:
            value += digit_value
        last_digit_value = digit_value

    return value

def build_corpus(docx_folder_path: str, corpus_output_path: str) -> None:
    """
    Build a text corpus from all DOCX files in the specified folder.

    This function iterates over all files in the specified folder, and for each file with a .docx extension,
    it extracts the text content and appends it to the corpus.

    :param docx_folder_path: The path to the folder containing DOCX files.
    :type docx_folder_path: str
    :param corpus_output_path: The path to the output file where the corpus will be saved.
    :type corpus_output_path: str
    :return: None
    """
    corpus_text = ""
    for file_name in os.listdir(docx_folder_path):
        if file_name.endswith('.docx'):
            docx_file_path = os.path.join(docx_folder_path, file_name)
            doc = docx.Document(docx_file_path)
            for paragraph in doc.paragraphs:
                corpus_text += paragraph.text + "\n"

    with open(corpus_output_path, 'w', encoding='utf-8') as file:
        file.write(corpus_text)

if __name__ == "__main__":
    # Example usage:
    # build_corpus('docx/', 'corpus.txt')