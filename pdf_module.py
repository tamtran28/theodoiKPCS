import pdfplumber
import pandas as pd
from docx import Document

def pdf_to_tables(file):
    """Trích xuất bảng từ PDF."""
    tables = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            raw = page.extract_tables()
            for t in raw:
                if not t:
                    continue
                header, *rows = t
                df = pd.DataFrame(rows, columns=header)
                tables.append(df)
    return tables


def word_to_tables(file):
    """Trích xuất bảng từ file Word (DOCX)."""
    doc = Document(file)
    tables = []

    for tbl in doc.tables:
        data = []
        for row in tbl.rows:
            data.append([cell.text.strip() for cell in row.cells])

        header, *rows = data
        df = pd.DataFrame(rows, columns=header)
        tables.append(df)

    return tables
