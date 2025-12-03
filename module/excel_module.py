from io import BytesIO
import pandas as pd

def save_to_excel(df):
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer


def merge_kien_nghi(file_main, file_new):
    # Đọc file Excel
    df_main = pd.read_excel(file_main)
    df_new = pd.read_excel(file_new)

    # Tìm STT lớn nhất
    if "STT" in df_main.columns:
        max_stt = pd.to_numeric(df_main["STT"], errors="coerce").max()
        if pd.isna(max_stt):
            max_stt = 0
    else:
        max_stt = 0

    # Tạo STT mới
    df_new["STT"] = range(int(max_stt) + 1, int(max_stt) + 1 + len(df_new))

    # Ghép file
    df_out = pd.concat([df_main, df_new], ignore_index=True)

    # Xuất
    buffer = BytesIO()
    df_out.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer
