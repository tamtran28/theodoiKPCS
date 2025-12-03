import pdfplumber
import pandas as pd

def pdf_to_tables(file):
    tables = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            raw_tables = page.extract_tables()
            for t in raw_tables:
                if not t:
                    continue
                header, *body = t
                df = pd.DataFrame(body, columns=header)
                tables.append(df)
    return tables
