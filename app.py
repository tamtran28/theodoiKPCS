# app.py
import streamlit as st
from module.ocr_module import read_word, read_pdf, ocr_image, ocr_pdf, extract_kien_nghi
from module.extract_module import create_excel, merge_kien_nghi

import streamlit as st
from io import BytesIO


st.set_page_config(page_title="Theo dõi KPCS", layout="wide")
st.title("📋 Công cụ tách PDF → Excel Kiến nghị")

tab_pdf, tab_excel = st.tabs([
    "📄 1. Tách bảng từ PDF",
    "📝 2. Map & Xuất Excel"
])

# ===================== TAB 1 =========================
with tab_pdf:
    st.header("📄 1. Tách bảng từ PDF")

    pdf_file = st.file_uploader("Chọn file PDF:", type=["pdf"])

    if pdf_file:
        st.info("⏳ Đang đọc PDF...")
        tables = pdf_to_tables(pdf_file)

        st.success(f"Đã tìm thấy {len(tables)} bảng.")

        for idx, df in enumerate(tables):
            with st.expander(f"Bảng #{idx} (cột: {len(df.columns)})"):
                st.dataframe(df)

        st.subheader("🔗 Chọn bảng tóm tắt & chi tiết")

        summary_idx = st.selectbox(
            "Bảng Tóm tắt",
            options=list(range(len(tables)))
        )
        detail_idx = st.selectbox(
            "Bảng Chi tiết",
            options=list(range(len(tables)))
        )

        st.session_state["summary_df"] = tables[summary_idx]
        st.session_state["detail_df"] = tables[detail_idx]

        st.success("Đã lưu bảng. Sang TAB 2 để xuất Excel.")


# ===================== TAB 2 =========================
with tab_excel:
    st.header("📝 2. Map cột & Xuất Excel")

    if "summary_df" not in st.session_state:
        st.warning("⚠ Chưa có dữ liệu. Bạn cần dùng TAB 1 trước.")
        st.stop()

    summary_df = st.session_state["summary_df"]
    detail_df = st.session_state["detail_df"]

    sum_cols = list(summary_df.columns)
    det_cols = list(detail_df.columns)

    st.subheader("🧩 Map bảng TÓM TẮT")
    map_summary = {
        "ten_phat_hien": st.selectbox("Tên phát hiện", sum_cols),
        "anh_huong": st.selectbox("Ảnh hưởng", sum_cols),
        "xep_rr": st.selectbox("Xếp hạng rủi ro", sum_cols),
        "xep_ks": st.selectbox("Xếp hạng kiểm soát", sum_cols),
        "so_luong": st.selectbox("Số lượng chi tiết", sum_cols),
    }

    st.subheader("🧩 Map bảng CHI TIẾT")
    map_detail = {
        "phat_hien_nn": st.selectbox("Phát hiện & Nguyên nhân", det_cols),
        "anh_huong": st.selectbox("Ảnh hưởng", det_cols),
        "kien_nghi": st.selectbox("Kiến nghị", det_cols),
        "y_kien": st.selectbox("Ý kiến đơn vị", det_cols),
    }

    block_col = st.selectbox(
        "Cột chứa block thông tin (Kế hoạch, Người thực hiện…)",
        ["(Không chọn)"] + det_cols
    )
    if block_col == "(Không chọn)":
        block_col = None

    if st.button("📦 Xuất Excel kiến nghị"):
        df_out = build_output_df(summary_df, detail_df, map_summary, map_detail, block_col)

        st.dataframe(df_out)

        excel_bytes = save_to_excel(df_out)

        st.download_button(
            "⬇ Tải file Excel",
            excel_bytes.getvalue(),
            file_name="kien_nghi.xlsx"
        )
