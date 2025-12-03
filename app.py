# app.py
import streamlit as st
from module.ocr_module import read_word, read_pdf, ocr_image, ocr_pdf, extract_kien_nghi
from module.extract_module import create_excel, merge_kien_nghi

st.set_page_config(page_title="Công cụ Kiến nghị Kiểm toán", layout="wide")

st.title("📋 Công cụ Kiến Nghị Kiểm Toán – Streamlit Cloud")
st.write("Hỗ trợ OCR tiếng Việt, tạo file kiến nghị mới & import kiến nghị vào file chính.")


# ==========================
# 1) TẠO FILE KIẾN NGHỊ MỚI
# ==========================

st.header("📝 1. Tạo file kiến nghị từ báo cáo")

uploaded = st.file_uploader("Tải báo cáo (PDF, DOCX, JPG, PNG):", 
                            type=["pdf", "docx", "jpg", "jpeg", "png"])

if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    st.info("Đang xử lý...")

    # OCR / TEXT
    if ext in ["jpg", "png", "jpeg"]:
        text = ocr_image(uploaded)
    elif ext == "pdf":
        t = read_pdf(uploaded)
        if len(t.strip()) < 20:
            st.warning("PDF scan → OCR...")
            text = ocr_pdf(uploaded)
        else:
            text = t
    elif ext == "docx":
        text = read_word(uploaded)

    st.subheader("📌 Preview nội dung:")
    st.text_area("Văn bản trích xuất", text[:3000], height=200)

    kien_nghi = extract_kien_nghi(text)

    st.success(f"Tìm thấy {len(kien_nghi)} kiến nghị.")

    if kien_nghi:
        excel_new = create_excel(kien_nghi)
        st.download_button(
            label="⬇ Tải file Excel kiến nghị mới",
            data=excel_new.getvalue(),
            file_name="kien_nghi_moi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ==========================
# 2) IMPORT KIẾN NGHỊ VÀO FILE CHÍNH
# ==========================

st.header("➕ 2. Thêm kiến nghị vào file KPCS chính")

file_main = st.file_uploader("File KPCS chính:", type=["xlsx"], key="main")
file_add  = st.file_uploader("File kiến nghị mới:", type=["xlsx"], key="add")

if file_main and file_add:
    if st.button("Thực hiện import"):
        result = merge_kien_nghi(file_main, file_add)
        st.success("Đã import thành công!")

        st.download_button(
            label="⬇ Tải file KPCS sau khi import",
            data=result.getvalue(),
            file_name="KPCS_updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
