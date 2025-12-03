from io import BytesIO

def save_to_excel(df):
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer
