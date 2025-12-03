import streamlit as st
import pandas as pd
from io import BytesIO

from module.pdf_module import pdf_to_tables, word_to_tables
from module.mapping_module import build_output_df
from module.excel_module import save_to_excel


st.set_page_config(page_title="Tách PDF/WORD → Excel", layout="wide")
st.title("📋 Công cụ tách PDF/WORD → Excel Kiến nghị")

tab_pdf, tab_excel = st.tabs([
    "📄 1. Tách bảng từ PDF/WORD",
    "📝 2. Map & Xuất Excel"
])


# ===================== TAB 1 =========================
with tab_pdf:
    st.header("📄 1. Tách bảng từ PDF hoặc WORD")

    file = st.file_uploader("Chọn file:", type=["pdf", "docx"])

    if file:
        ext = file.name.lower().split(".")[-1]

        st.info("⏳ Đang đọc bảng...")

        if ext == "pdf":
            tables = pdf_to_tables(file)

        elif ext == "docx":
            tables = word_to_tables(file)

        st.success(f"Đã tìm thấy {len(tables)} bảng.")

        for i, df in enumerate(tables):
            with st.expander(f"BẢNG #{i}"):
                st.dataframe(df)

        summary_idx = st.selectbox("Chọn bảng TÓM TẮT", list(range(len(tables))))
        detail_idx = st.selectbox("Chọn bảng CHI TIẾT", list(range(len(tables))))

        st.session_state["summary_df"] = tables[summary_idx]
        st.session_state["detail_df"] = tables[detail_idx]


# ===================== TAB 2 =========================
with tab_excel:
    st.header("📝 2. Map & Xuất Excel")

    if "summary_df" not in st.session_state:
        st.warning("⚠ Chưa có dữ liệu.")
        st.stop()

    summary_df = st.session_state["summary_df"]
    detail_df = st.session_state["detail_df"]

    sum_cols = list(summary_df.columns)
    det_cols = list(detail_df.columns)

    map_summary = {
        "ten_phat_hien": st.selectbox("Tên phát hiện (Tóm tắt)", sum_cols),
        "anh_huong": st.selectbox("Ảnh hưởng (Tóm tắt)", sum_cols),
        "xep_rr": st.selectbox("Xếp hạng rủi ro", sum_cols),
        "xep_ks": st.selectbox("Xếp hạng kiểm soát", sum_cols),
        "so_luong": st.selectbox("Số lượng chi tiết", sum_cols),
    }

    map_detail = {
        "phat_hien_nn": st.selectbox("Phát hiện & Nguyên nhân", det_cols),
        "anh_huong": st.selectbox("Ảnh hưởng (chi tiết)", det_cols),
        "kien_nghi": st.selectbox("Kiến nghị", det_cols),
        "y_kien": st.selectbox("Ý kiến đơn vị", det_cols),
    }

    block_col = st.selectbox("Cột chứa Kế hoạch / Người duyệt / Ngày hoàn thành", ["(none)"] + det_cols)
    if block_col == "(none)":
        block_col = None

    if st.button("📦 Xuất Excel"):
        df_out = build_output_df(summary_df, detail_df, map_summary, map_detail, block_col)

        st.dataframe(df_out)

        excel_bytes = save_to_excel(df_out)

        st.download_button("⬇ Tải Excel", excel_bytes.getvalue(), "kien_nghi.xlsx")
