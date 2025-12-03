# ocr_module.py
import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
from PyPDF2 import PdfReader

# ==== ĐỌC WORD ====
def read_word(file):
    doc = Document(file)
    return "\n".join(p.text for p in doc.paragraphs)

# ==== ĐỌC PDF TEXT ====
def read_pdf(file):
    pdf = PdfReader(file)
    text = ""
    for page in pdf.pages:
        try:
            text += page.extract_text() + "\n"
        except:
            pass
    return text

# ==== OCR ẢNH ====
def ocr_image(file):
    img = Image.open(file).convert("RGB")
    img_np = np.array(img)

    gray = cv2.cvtColor(img_np, cv2.COLOR_BGR2GRAY)
    gray = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)[1]

    text = pytesseract.image_to_string(gray, lang='vie')
    return text

# ==== OCR PDF SCAN ====
def ocr_pdf(file):
    pages = convert_from_path(file, dpi=250)
    result = ""

    for page in pages:
        img = np.array(page)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY)[1]
        text = pytesseract.image_to_string(gray, lang='vie')
        result += text + "\n"

    return result


# ==== TÁCH KIẾN NGHỊ ====
def extract_kien_nghi(text):
    t = text.lower()
    start = t.find("kiến nghị")
    if start == -1:
        return []

    section = text[start:]

    import re
    parts = re.split(r"\n\s*\d+[\.\)]\s+", section)

    results = []
    for p in parts:
        p = p.strip()
        if len(p) > 10 and not p.lower().startswith("kiến nghị"):
            results.append(p)
    return results
