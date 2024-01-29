# ATS_Topic Modelling Scripts

This repository contains a collection of scripts used for various tasks in a research project. The tasks include downloading articles, converting PDF files to DOCX format using OCR, building a text corpus from DOCX files, deleting extra files, and downloading science papers.

This repository also contains information about the models used and the corresponding results in the Modelling folder.

## Setup

1. Clone this repository to your local machine.
2. Navigate to the project directory.
3. Install the required dependencies using pip:

```bash
pip install -r requirements.txt
```

## Usage

The `main.py` script demonstrates how to use the functions from the refactored scripts. You can run `main.py` to see the scripts in action, or import the functions from the scripts in your own Python files.

Here's an example of how you might use the `download_articles` function from the `download_articles.py` script:

```python
from src.download_articles import download_articles

# Download articles listed in input.csv and save them to the articles/ directory
download_articles('input.csv', 'articles/')
```

## License

This project is licensed under the terms of the Research Use License. See the `LICENSE` file for details.
