# extract_module.py
from io import BytesIO
from datetime import datetime

import openpyxl
from openpyxl import Workbook, load_workbook
from dateutil.relativedelta import relativedelta

from extract_fields import extract_all_fields


# ---- TÍNH THỜI HẠN HOÀN THÀNH = NGÀY BAN HÀNH + X THÁNG ----
def calc_deadline(ngay_ban_hanh_str, uu_tien_str):
    """
    ngay_ban_hanh_str: 'mm/dd/yyyy' (chuỗi)
    uu_tien_str: số tháng (chuỗi/ số), ví dụ '2'
    """
    try:
        if not ngay_ban_hanh_str or not uu_tien_str:
            return ""
        dt = datetime.strptime(str(ngay_ban_hanh_str), "%m/%d/%Y")
        months = int(uu_tien_str)
        dt2 = dt + relativedelta(months=months)
        return dt2.strftime("%m/%d/%Y")
    except Exception:
        return ""


# ---- TẠO FILE EXCEL KIẾN NGHỊ MỚI ----
def create_excel(kien_nghi_list, doi_tuong, so_van_ban, ngay_ban_hanh):
    wb = Workbook()
    ws = wb.active
    ws.title = "KPCS"

    # Header theo mẫu (có thể chỉnh lại cho khớp file chuẩn của bạn)
    columns = [
        "STT",
        "Đối tượng được KT",
        "Số văn bản",
        "Ngày, tháng, năm ban hành (mm/dd/yyyy)",
        "Tên Đoàn kiểm toán",
        "Số hiệu rủi ro",
        "Số hiệu kiểm soát",
        "Nghiệp vụ (R0)",
        "Quy trình/hoạt động con (R1)",
        "Tên phát hiện (R2)",
        "Chi tiết phát hiện (R3)",
        "Dẫn chiếu",
        "Mô tả chi tiết phát hiện",
        "CIF Khách hàng/bút toán",
        "Tên khách hàng",
        "Loại KH",
        "Số phát hiện/số mẫu chọn",
        "Dư nợ sai phạm  (Triệu đồng)",
        "Số tiền tổn thất  (Triệu đồng)",
        "Số tiền cần thu hồi (Triệu đồng)",
        "Trách nhiệm trực tiếp",
        "Trách nhiệm quản lý",
        "Xếp hạng rủi ro",
        "Xếp hạng kiểm soát",
        "Nguyên nhân",
        "Ảnh hưởng",
        "Kiến nghị",
        "Loại/nhóm nguyên nhân",
        "Loại/nhóm ảnh hưởng",
        "Loại/nhóm kiến nghị",
        "Chủ thể kiến nghị",
        "Kế hoạch thực hiện",
        "Trách nhiệm thực hiện",
        "Đơn vị thực hiện KPCS",
        "ĐVKD, AMC, Hội sở",
        "Người phê duyệt",
        "Ý kiến của đơn vị",
        "Mức độ ưu tiên hành động",
        "Thời hạn hoàn thành (mm/dd/yyyy)",
        "Đã khắc phục",
        "Ngày đã KPCS (mm/dd/yyyy)",
        "CBKT",
    ]

    # ghi header
    for col_idx, name in enumerate(columns, start=1):
        ws.cell(1, col_idx, name)

    # để tiện: map tên → index
    col_index = {name: i + 1 for i, name in enumerate(columns)}

    # ghi từng kiến nghị
    for i, kn in enumerate(kien_nghi_list, start=2):
        ws.cell(i, col_index["STT"], i - 1)
        ws.cell(i, col_index["Đối tượng được KT"], doi_tuong or "")
        ws.cell(i, col_index["Số văn bản"], so_van_ban or "")
        ws.cell(i, col_index["Ngày, tháng, năm ban hành (mm/dd/yyyy)"], ngay_ban_hanh or "")
        ws.cell(i, col_index["Kiến nghị"], kn)

        # tách thêm các trường (nếu trong text có)
        fields = extract_all_fields(kn)
        if fields["nguyen_nhan"]:
            ws.cell(i, col_index["Nguyên nhân"], fields["nguyen_nhan"])
        if fields["uu_tien"]:
            ws.cell(i, col_index["Mức độ ưu tiên hành động"], fields["uu_tien"])
            # nếu có ngày ban hành + ưu tiên → tính luôn deadline
            dl = calc_deadline(ngay_ban_hanh, fields["uu_tien"])
            if dl:
                ws.cell(i, col_index["Thời hạn hoàn thành (mm/dd/yyyy)"], dl)
        if fields["nguoi_thuc_hien"]:
            ws.cell(i, col_index["Trách nhiệm thực hiện"], fields["nguoi_thuc_hien"])
        if fields["nguoi_phe_duyet"]:
            ws.cell(i, col_index["Người phê duyệt"], fields["nguoi_phe_duyet"])
        if fields["ngay_hoan_thanh"]:
            # nếu text có ngày hoàn thành riêng → cũng có thể ghi vào
            ws.cell(i, col_index["Thời hạn hoàn thành (mm/dd/yyyy)"], fields["ngay_hoan_thanh"])

    out = BytesIO()
    wb.save(out)
    return out


# ---- IMPORT KIẾN NGHỊ MỚI VÀO FILE KPCS CHÍNH ----
def merge_kien_nghi(file_main, file_new):
    wb_main = load_workbook(file_main)
    ws_main = wb_main.active

    wb_new = load_workbook(file_new)
    ws_new = wb_new.active

    # đọc header của file chính
    header = {}
    for c in range(1, ws_main.max_column + 1):
        name = ws_main.cell(1, c).value
        if name:
            header[name] = c

    # tìm cột theo tên (tương đối linh hoạt)
    def find_col_by_keyword(keyword_list):
        for name, idx in header.items():
            lower = str(name).lower()
            if all(k in lower for k in keyword_list):
                return idx
        return None

    col_ngay_bh = find_col_by_keyword(["ngày", "ban hành"])
    col_uu_tien = find_col_by_keyword(["mức độ ưu tiên"])
    col_deadline = find_col_by_keyword(["thời hạn", "hoàn thành"])

    # nếu thiếu cột deadline → thêm mới
    if col_deadline is None:
        col_deadline = ws_main.max_column + 1
        ws_main.cell(1, col_deadline, "Thời hạn hoàn thành (mm/dd/yyyy)")

    # ghép từng dòng kiến nghị mới
    for row in ws_new.iter_rows(min_row=2, values_only=True):
        new_row_idx = ws_main.max_row + 1
        # copy toàn bộ giá trị sang
        for col_idx, val in enumerate(row, start=1):
            ws_main.cell(new_row_idx, col_idx, val)

        # tính deadline nếu có đủ dữ liệu
        ngay_bh_val = ws_main.cell(new_row_idx, col_ngay_bh).value if col_ngay_bh else None
        uu_tien_val = ws_main.cell(new_row_idx, col_uu_tien).value if col_uu_tien else None
        dl = calc_deadline(ngay_bh_val, uu_tien_val)
        if dl:
            ws_main.cell(new_row_idx, col_deadline, dl)

    out = BytesIO()
    wb_main.save(out)
    return out
