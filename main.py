import pymupdf
import pandas as pd
import re

def clean_cell(value):
    """ Remove illegal characters and ensure the cell is a valid string. """
    if isinstance(value, str):
        # Remove non-printable characters
        return re.sub(r'[\x00-\x1F\x7F-\x9F]', '', value).strip()  
    elif isinstance(value, (int, float)):  
        return value
    elif value is None:
        return ""
    return str(value)

def extract_tables_from_pymupdf(pdf_path, output_excel):
    """
    Extract tables from a PDF using PyMuPDF and save them to an Excel file.
    :param pdf_path: Path to the input PDF file
    :param output_excel: Path to the output Excel file
    """
    extracted_tables = []
    
    with pymupdf.open(pdf_path) as doc:
        for i, page in enumerate(doc):
            tables = page.find_tables(strategy='text')
            for table in tables.tables:
                df = table.to_pandas()
                df = df.applymap(clean_cell)
                extracted_tables.append((f"Page_{i+1}", df))
    
    if extracted_tables:
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for sheet_name, df in extracted_tables:
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=True)
        print(f"Extraction complete. Tables saved to {output_excel}")
    else:
        print("No tables found in the PDF.")

# Example usage
pdf_file = "./test3.pdf"
output_file = "./output.xlsx"
extract_tables_from_pymupdf(pdf_file, output_file)
