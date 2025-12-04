import pdfplumber
import pandas as pd

def clean_header(cols):
    """Làm sạch header bị xuống dòng hoặc None."""
    out = []
    for c in cols:
        if c is None:
            out.append("")
        else:
            out.append(str(c).replace("\n", " ").strip())
    return out


def pdf_to_tables(file):
    tables = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            raw_tables = page.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5
            })

            for tbl in raw_tables:
                if not tbl or len(tbl) < 2:
                    continue

                header_raw = tbl[0]
                header = clean_header(header_raw)

                rows = tbl[1:]

                df = pd.DataFrame(rows, columns=header)

                # bỏ cột rỗng
                df = df.loc[:, ~(df.columns == "")]

                # bỏ dòng rác
                df = df.dropna(how="all")

                if len(df.columns) < 2:
                    continue

                tables.append(df)

    return tables
