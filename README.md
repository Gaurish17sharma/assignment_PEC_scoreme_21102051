# assignment_PEC_scoreme_21102051


# PDF Table Extraction with PyMuPDF

## Overview
This script extracts tables from a PDF document using PyMuPDF (also known as Fitz) and saves the extracted tables into an Excel file. It ensures that the extracted data is clean by removing any non-printable characters.

## Features
- Extracts tables from a PDF file using PyMuPDF's `find_tables` function.
- Cleans extracted text to remove illegal or non-printable characters.
- Saves extracted tables into an Excel file with separate sheets for each page.
- Uses `pandas` and `openpyxl` to handle data processing and Excel file creation.

## Requirements
Make sure you have the following Python libraries installed:

```sh
pip install pymupdf pandas openpyxl
```

## Usage
### 1. Modify the script
Update the `pdf_file` and `output_file` variables in the script to specify your input PDF and output Excel file paths:

```python
pdf_file = "./test6.pdf"  # Replace with your actual file path
output_file = "./extracted_tables2.xlsx"
extract_tables_from_pymupdf(pdf_file, output_file)
```

### 2. Run the script
Execute the script in your terminal or command prompt:

```sh
python script.py
```

### 3. Output
The extracted tables will be saved in an Excel file at the specified path. Each page containing tables will be stored as a separate sheet.

## Function Details
### `clean_cell(value)`
- Cleans cell values by removing non-printable characters.
- Keeps numerical values unchanged.
- Converts other types to strings if necessary.

### `extract_tables_from_pymupdf(pdf_path, output_excel)`
- Opens a PDF file and extracts tables using `find_tables(strategy='text')`.
- Cleans the extracted data using `clean_cell`.
- Saves the tables to an Excel file with separate sheets for each page.

## Notes
- Ensure that your PDF contains properly structured tables for accurate extraction.
- The script may not work well with complex or heavily formatted tables.
- If no tables are found, a warning message is displayed.

## License
This script is open-source and free to use and modify.

