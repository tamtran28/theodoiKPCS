# app.py
import streamlit as st
from module.ocr_module import read_word, read_pdf, ocr_image, ocr_pdf, extract_kien_nghi
from module.extract_module import create_excel, merge_kien_nghi

import streamlit as st
from io import BytesIO


# ==============================
# CẤU HÌNH GIAO DIỆN
# ==============================
st.set_page_config(page_title="Công cụ Kiến nghị Kiểm toán", layout="wide")
st.title("📋 Công cụ Kiến nghị Kiểm toán")
st.write("• Tạo file kiến nghị từ báo cáo (DOCX / PDF / Ảnh, hỗ trợ OCR tiếng Việt)"
         "\n• Import file kiến nghị vào file KPCS chính và tự tính thời hạn hoàn thành.")


# =======================================================
# 1) TẠO FILE KIẾN NGHỊ MỚI TỪ BÁO CÁO
# =======================================================
st.header("📝 1. Tạo file kiến nghị từ báo cáo")

uploaded = st.file_uploader(
    "Tải báo cáo (PDF, DOCX, JPG, PNG):",
    type=["pdf", "docx", "jpg", "jpeg", "png"],
    key="bao_cao"
)

# ------ INPUT THÔNG TIN CHUNG ------
st.subheader("🔧 Thông tin chung áp dụng cho TẤT CẢ kiến nghị")
doi_tuong = st.text_input("Đối tượng được KT:")
so_van_ban = st.text_input("Số văn bản:")
ngay_ban_hanh = st.text_input("Ngày, tháng, năm ban hành (mm/dd/yyyy):")

if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    st.info("⏳ Đang xử lý báo cáo...")

    text = ""

    # ====== XỬ LÝ ẢNH ======
    if ext in ["jpg", "jpeg", "png"]:
        text = ocr_image(uploaded)

    # ====== XỬ LÝ PDF ======
    elif ext == "pdf":
        file_bytes = uploaded.getvalue()

        # Thử đọc PDF text
        try:
            text_try = read_pdf(BytesIO(file_bytes))
        except:
            text_try = ""

        # Nếu text PDF rỗng → OCR scan
        if not text_try or len(text_try.strip()) < 20:
            st.warning("PDF có thể là bản scan → đang OCR tiếng Việt…")
            text = ocr_pdf(file_bytes)
        else:
            text = text_try

    # ====== XỬ LÝ DOCX ======
    elif ext == "docx":
        text = read_word(uploaded)

    # Hiển thị preview
    st.subheader("📌 Preview văn bản trích xuất")
    st.text_area("Nội dung (đã cắt bớt nếu quá dài):", text[:3000], height=250)

    # Tách kiến nghị
    kien_nghi_list = extract_kien_nghi(text)
    st.success(f"🔍 Đã tìm được {len(kien_nghi_list)} kiến nghị.")

    # Tạo Excel kiến nghị
    if kien_nghi_list and st.button("📦 Tạo file Excel kiến nghị mới"):
        excel_bytes = create_excel(
            kien_nghi_list=kien_nghi_list,
            doi_tuong=doi_tuong,
            so_van_ban=so_van_ban,
            ngay_ban_hanh=ngay_ban_hanh
        )
        st.download_button(
            label="⬇ Tải file Excel kiến nghị mới",
            data=excel_bytes.getvalue(),
            file_name="kien_nghi_moi.xlsx",
            mime=("application/vnd.openxmlformats-officedocument."
                  "spreadsheetml.sheet")
        )


# =======================================================
# 2) IMPORT KIẾN NGHỊ VÀO FILE KPCS CHÍNH
# =======================================================
st.header("➕ 2. Thêm kiến nghị vào file KPCS chính")

file_main = st.file_uploader("File KPCS chính (.xlsx):", type=["xlsx"], key="main")
file_add = st.file_uploader("File kiến nghị mới (.xlsx):", type=["xlsx"], key="new")

if file_main and file_add:
    if st.button("🔁 Import kiến nghị vào file chính"):

        # Chi tiết rất quan trọng: Reset con trỏ
        file_main.seek(0)
        file_add.seek(0)

        merged_bytes = merge_kien_nghi(file_main, file_add)
        st.success("✅ Đã import kiến nghị vào file KPCS chính.")

        st.download_button(
            label="⬇ Tải file KPCS sau khi import",
            data=merged_bytes.getvalue(),
            file_name="KPCS_updated.xlsx",
            mime=("application/vnd.openxmlformats-officedocument."
                  "spreadsheetml.sheet")
        )
