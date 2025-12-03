# ocr_module.py
import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
from PyPDF2 import PdfReader
from io import BytesIO
import re
import tempfile


# ---- ĐỌC WORD ----
def read_word(file_obj):
    doc = Document(file_obj)
    return "\n".join(p.text for p in doc.paragraphs)


# ---- ĐỌC PDF TEXT ----
def read_pdf(file_obj):
    """
    file_obj: file-like object (BytesIO hoặc UploadedFile)
    """
    reader = PdfReader(file_obj)
    text = ""
    for page in reader.pages:
        try:
            txt = page.extract_text()
            if txt:
                text += txt + "\n"
        except Exception:
            pass
    return text


# ---- OCR ẢNH ----
def ocr_image(file_obj):
    """
    file_obj: UploadedFile hoặc file-like
    """
    img = Image.open(file_obj).convert("RGB")
    img_np = np.array(img)

    gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)[1]

    text = pytesseract.image_to_string(gray, lang="vie")
    return text


# ---- OCR PDF SCAN ----
def ocr_pdf(file_bytes: bytes):
    """
    file_bytes: bytes của PDF (vd: uploaded.getvalue())
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file_bytes)
        tmp.flush()
        pages = convert_from_path(tmp.name, dpi=250)

    result = ""
    for page in pages:
        img_np = np.array(page)
        gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
        gray = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)[1]
        result += pytesseract.image_to_string(gray, lang="vie") + "\n"

    return result


# ---- TÁCH MỤC "KIẾN NGHỊ" THÀNH CÁC KIẾN NGHỊ RIÊNG ----
def extract_kien_nghi(text: str):
    """
    Tìm mục 'Kiến nghị' trong báo cáo, sau đó tách thành từng kiến nghị
    đánh số 1., 2., 3., ...
    """
    lower = text.lower()
    start_idx = lower.find("kiến nghị")
    if start_idx == -1:
        return []

    section = text[start_idx:]

    # tách theo dòng bắt đầu bằng số + . hoặc )
    parts = re.split(r"\n\s*\d+[\.\)]\s+", section)

    results = []
    for p in parts:
        p = p.strip()
        if len(p) <= 10:
            continue
        if p.lower().startswith("kiến nghị"):
            continue
        results.append(p)

    return results
