## Document Generation Tool

This repository contains a Python-based tool for generating personalized documents for employees using predefined templates and data from multiple file types. The tool leverages OpenAI's GPT-4 model for extracting values from documents and supports .txt, .docx, and .pdf files. The final documents are zipped and made available for download via a Streamlit web interface.

### Features

- **Template Parsing:** Extracts keys from a DOCX template.
- **Document Reading:** Reads text from .txt, .docx, and .pdf files.
- **Data Extraction:** Uses OpenAI's GPT-4 to extract key-value pairs from documents.
- **CSV Integration:** Merges extracted values with data from a CSV file.
- **Template Filling:** Populates the template with extracted values.
- **Zipping Files:** Bundles the generated documents into a ZIP file.
- **Streamlit Interface:** Provides a user-friendly web interface for selecting document types and employees, and downloading the results.

### Requirements

- Python 3.11
- Required Python libraries: `os`, `re`, `json`, `fitz`, `pandas`, `docx`, `openai`, `streamlit`, `zipfile`, `datetime`

### Installation

1. Clone the repository.
2. Install the required Python libraries using pip:
    ```bash
    pip install -r requirements.txt
    ```

### Usage

1. **Prepare your template:**
    - Ensure your DOCX template file uses delimiters [key] for the placeholders you want to populate.
  
2. **Organize your data:**
    - Place your .txt, .docx, and .pdf files in their respective folders.
    - Ensure your CSV file is formatted correctly and located in the specified path.

3. **Run the application:**
    - Use the Streamlit command to run the application:
      ```bash
      streamlit run app.py
      ```
    - Navigate to the Streamlit interface in your browser.
    - Select the document type, employees, and initiate document creation.
    - Download the resulting ZIP file containing the generated documents.

### Code Description

#### Import Libraries
```python
import os
import re
import json
import fitz
import pandas as pd
from docx import Document
from openai import OpenAI
import streamlit as st
import zipfile
import datetime
```

#### Initialize OpenAI Client
```python
client = OpenAI()
```

#### Functions

- **extract_keys_from_template(template_docx_path)**
  - Extracts keys from a DOCX template enclosed in double curly braces.
  
- **read_txt_file(filepath)**
  - Reads text from a .txt file.
  
- **read_docx_file(filepath)**
  - Reads text from a .docx file.
  
- **read_pdf_file(filepath)**
  - Reads text from a .pdf file.
  
- **read_documents_from_folders(folders)**
  - Reads documents from specified folders.
  
- **extract_values_from_csv(csv_path, member, opcija)**
  - Extracts values from a CSV file based on member and option criteria.
  
- **extract_values_from_docs(keys, documents, csv_values)**
  - Uses OpenAI's GPT-4 to extract values from documents and integrates CSV values.
  
- **create_filled_template(template_docx_path, values_dict, output_docx_path)**
  - Populates the DOCX template with extracted values and saves the output.
  
- **zip_specific_files(files_to_zip, zip_name)**
  - Zips specified files into a ZIP archive.

#### Main Function
```python
def main(template_path, folders, base_folder, fixed_folder, members, opcija, podela):
    # Main logic for extracting keys, reading documents, extracting values, and creating filled templates.
    # Zips the final documents and provides a download link via Streamlit.
```

#### Streamlit Interface
```python
if __name__ == "__main__":
    # Streamlit UI for document selection and generation.
```

### Contributing

Contributions are welcome! Please fork the repository and create a pull request with your changes.

### License

This project is licensed under the MIT License. See the LICENSE file for details.
