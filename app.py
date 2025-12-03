import streamlit as st
from io import BytesIO

from module.ocr_module import read_word, read_pdf, ocr_image, ocr_pdf, extract_kien_nghi
from module.extract_module import create_excel, merge_kien_nghi

st.set_page_config(page_title="Công cụ Kiến nghị Kiểm toán", layout="wide")

st.title("📋 Công cụ Kiến nghị Kiểm toán")
st.write(
    "- Tạo file kiến nghị từ báo cáo (DOCX / PDF / Ảnh, OCR tiếng Việt)\n"
    "- Import kiến nghị mới vào file KPCS chính\n"
    "- Tự động tính Thời hạn hoàn thành = Ngày ban hành + Mức độ ưu tiên (tháng)\n"
    "- Cột 'Kiến nghị' chỉ lấy đoạn bắt đầu từ 'Đề nghị'"
)

# =====================================================
# 1) TẠO FILE KIẾN NGHỊ MỚI
# =====================================================
st.header("📝 1. Tạo file kiến nghị từ báo cáo")

uploaded = st.file_uploader(
    "Tải báo cáo (PDF, DOCX, JPG, PNG):",
    type=["pdf", "docx", "jpg", "jpeg", "png"],
)

st.subheader("🔧 Thông tin chung cho TẤT CẢ kiến nghị")
doi_tuong = st.text_input("Đối tượng được KT:")
so_van_ban = st.text_input("Số văn bản:")
ngay_ban_hanh = st.text_input("Ngày, tháng, năm ban hành (mm/dd/yyyy):")

if uploaded:
    ext = uploaded.name.split(".")[-1].lower()
    st.info("⏳ Đang xử lý báo cáo...")

    text = ""
    file_bytes = uploaded.getvalue()

    # ========== Ảnh ==========
    if ext in ["jpg", "jpeg", "png"]:
        text = ocr_image(uploaded)

    # ========== PDF ==========
    elif ext == "pdf":
        try:
            text_try = read_pdf(BytesIO(file_bytes))
        except:
            text_try = ""

        if not text_try or len(text_try.strip()) < 20:
            st.warning("PDF scan → OCR tiếng Việt...")
            text = ocr_pdf(file_bytes)
        else:
            text = text_try

    # ========== DOCX ==========
    elif ext == "docx":
        text = read_word(uploaded)

    st.subheader("📌 Preview văn bản trích xuất")
    st.text_area("Văn bản OCR:", text[:3000], height=250)

    kien_nghi_list = extract_kien_nghi(text)
    st.success(f"🔍 Đã tìm được {len(kien_nghi_list)} kiến nghị.")

    if kien_nghi_list and st.button("📦 Tạo file Excel kiến nghị mới"):
        excel_file = create_excel(
            kien_nghi_list=kien_nghi_list,
            doi_tuong=doi_tuong,
            so_van_ban=so_van_ban,
            ngay_ban_hanh=ngay_ban_hanh
        )

        st.download_button(
            "⬇ Tải file kiến nghị mới",
            data=excel_file.getvalue(),
            file_name="kien_nghi_moi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# =====================================================
# 2) IMPORT KIẾN NGHỊ VÀO FILE CHÍNH
# =====================================================
st.header("➕ 2. Thêm kiến nghị vào file KPCS chính")

file_main = st.file_uploader("File KPCS chính:", type=["xlsx"], key="main")
file_new = st.file_uploader("File kiến nghị mới:", type=["xlsx"], key="add")

if file_main and file_new:
    if st.button("🔁 Import vào file chính"):
        file_main.seek(0)
        file_new.seek(0)

        merged_bytes = merge_kien_nghi(file_main, file_new)

        st.success("✅ Import thành công!")

        st.download_button(
            "⬇ Tải file KPCS sau khi import",
            data=merged_bytes.getvalue(),
            file_name="KPCS_updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
