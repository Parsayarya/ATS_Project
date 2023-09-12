import os
import docx
import re
import string
import pandas as pd


def get_text(filename):
    """
    Extract text from a DOCX file.

    :param str filename: The path to the DOCX file.
    :return: The extracted text as a single string.
    :rtype: str
    """
    doc = docx.Document(filename)
    full_text = [para.text for para in doc.paragraphs]
    return '\n'.join(full_text)


def clean_text(text):
    """
    Clean and preprocess text.

    :param str text: The input text to be cleaned.
    :return: The cleaned text.
    :rtype: str
    """
    # Convert text to lowercase
    text = text.lower()
    # Remove text enclosed in square brackets
    text = re.sub(r'\[.*?\]', ' ', text)
    # Remove punctuation characters
    text = re.sub(r'[%s]' % re.escape(string.punctuation), ' ', text)
    # Remove words containing digits
    text = re.sub(r'\w*\d\w*', ' ', text)
    # Replace special character '�' with a space
    text = text.replace('�', ' ')
    return text


def build_text_list(text_folder):
    """
    Build a list of cleaned texts from a folder of DOCX files.

    :param str text_folder: The path to the folder containing DOCX files.
    :return: A list of cleaned text strings.
    :rtype: list[str]
    """
    text_list = []
    for filename in os.listdir(text_folder):
        file_path = os.path.join(text_folder, filename)
        try:
            # Attempt to clean and append the text
            cleaned_text = re.sub(' {2,}', ' ', clean_text(get_text(file_path).replace('\n', '').strip()))
            text_list.append(cleaned_text)
        except Exception as e:
            # If an exception occurs (e.g., file not unreadable), append 0
            text_list.append('0')
    return text_list


def create_dataframe(text_folder_IP, text_folder_WP, ip_excel_path, wp_excel_path):
    """
    Create a consolidated DataFrame from text and Excel data.

    This function reads text files from two directories and Excel files containing metadata.
    It then associates the text data with the metadata and returns a consolidated DataFrame.

    :param str text_folder_IP: Path to the folder containing IP text files
    :param str text_folder_WP: Path to the folder containing WP text files
    :param str ip_excel_path: Path to the Excel file containing IP metadata
    :param str wp_excel_path: Path to the Excel file containing WP metadata
    :return: Consolidated DataFrame containing text data and associated metadata
    :rtype: pd.DataFrame
    """
    # Get lists of texts from the specified IP and WP directories
    text_list_IP = build_text_list(text_folder_IP)
    text_list_WP = build_text_list(text_folder_WP)

    # Read the Excel files into pandas dataframes and sort them
    df_IP = pd.read_excel(ip_excel_path).sort_values(by=['Meeting', 'No.'])
    df_WP = pd.read_excel(wp_excel_path).sort_values(by=['Meeting', 'No.'])

    # Add the text data as a new column in both dataframes
    df_IP['Text'] = text_list_IP
    df_WP['Text'] = text_list_WP

    # Concatenate the dataframes and remove rows with 'Text' column equal to '0'
    df = pd.concat([df_IP, df_WP])
    df = df[df['Text'] != '0']

    return df


# Paths to your folders and files
text_folder_IP = r'D:/python/ATS/DownloadCorpusBuilder/files/docxOutputs/IP/'
text_folder_WP = r'D:/python/ATS/DownloadCorpusBuilder/files/docxOutputs/WP/'
ip_excel_path = r'D:/python/ATS/DownloadCorpusBuilder/files/csvInputs/listofpapersI.xlsx'
wp_excel_path = r'D:/python/ATS/DownloadCorpusBuilder/files/csvInputs/listofpapersW.xlsx'

# Call the function and get the final dataframe
df = create_dataframe(text_folder_IP, text_folder_WP, ip_excel_path, wp_excel_path)
# print(df.head(10), '\n\n\n', df.describe(), '\n\n\n', df.info(),'\n\n\n\n\n\n',df[df['Text'] == '0'])

# text_folder_IP = r'D:/python/ATS/DownloadCorpusBuilder/files/docxOutputs/IP/'
# text_folder_WP = r'D:/python/ATS/DownloadCorpusBuilder/files/docxOutputs/WP/'
# text_list_IP = build_text_list(text_folder_IP)
# text_list_WP = build_text_list(text_folder_WP)
# df_IP = pd.read_excel(r'D:/python/ATS/DownloadCorpusBuilder/files/csvInputs/listofpapersI.xlsx')
# df_WP = pd.read_excel(r'D:/python/ATS/DownloadCorpusBuilder/files/csvInputs/listofpapersW.xlsx')
# df_IP = df_IP.sort_values(by=['Meeting', 'No.'])
# df_WP = df_WP.sort_values(by=['Meeting', 'No.'])
# df_IP['Text'] = text_list_IP
# df_WP['Text'] = text_list_WP
# frames = [df_IP, df_WP]
# df = pd.concat(frames)
# df = df.drop(df[df['Text'] == '0'].index)
# print(df.head(10), '\n\n\n', df.describe(), '\n\n\n', df.info())
