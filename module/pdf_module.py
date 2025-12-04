from docx import Document
import pandas as pd

def word_to_tables(file):
    """Trích bảng Word CHUẨN, xử lý được:
       - merge ô
       - nhiều dòng header
       - dòng tiêu đề kiểu 'Chính sách, quy định…'
    """
    doc = Document(file)
    tables = []

    for tbl in doc.tables:
        data = []

        for row in tbl.rows:
            row_text = []
            for cell in row.cells:
                # lấy toàn bộ paragraph trong cell
                txt = "\n".join([p.text.strip() for p in cell.paragraphs])
                txt = txt.replace("\n", " ").strip()
                row_text.append(txt)

            # bỏ dòng rỗng
            if all(x == "" for x in row_text):
                continue

            data.append(row_text)

        # Nếu dòng đầu không phải header thật → bỏ
        if len(data) < 2:
            continue

        # Chuẩn hóa header (lấy dòng dài nhất)
        header = max(data[:3], key=lambda r: sum(len(c) for c in r))  
        rows = data[data.index(header)+1:]

        # Trim cột rỗng
        while len(header)>0 and header[-1] == "":
            header = header[:-1]
            for r in rows:
                if len(r)>0: r.pop()

        df = pd.DataFrame(rows, columns=header)
        tables.append(df)

    return tables
