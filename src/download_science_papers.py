import requests
import os

def download_papers(paper_urls: list, output_directory: str) -> None:
    """
    Download a list of science papers given their URLs.

    :param paper_urls: A list of URLs of the science papers to download.
    :type paper_urls: list
    :param output_directory: The directory where the downloaded papers should be saved.
    :type output_directory: str
    :raises FileNotFoundError: If the specified output directory does not exist.
    :return: None
    """

    # Check if output directory exists
    if not os.path.exists(output_directory):
        raise FileNotFoundError(f"The specified output directory does not exist: {output_directory}")

    for url in paper_urls:
        response = requests.get(url)
        file_name = url.split('/')[-1]
        file_path = os.path.join(output_directory, file_name)
        with open(file_path, 'wb') as file:
            file.write(response.content)

if __name__ == "__main__":
    # Example usage:
    # download_papers(['url1', 'url2', 'url3'], 'papers/')