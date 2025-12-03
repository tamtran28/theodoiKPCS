# extract_module.py
from io import BytesIO
import openpyxl
from openpyxl import Workbook, load_workbook
from datetime import datetime
from dateutil.relativedelta import relativedelta


# ==== TÍNH DEADLINE ====
def calc_deadline(date_str, priority):
    try:
        if not date_str or not priority:
            return ""
        dt = datetime.strptime(date_str, "%m/%d/%Y")
        months = int(priority)
        return (dt + relativedelta(months=months)).strftime("%m/%d/%Y")
    except:
        return ""


# ==== TẠO FILE EXCEL MỚI KIẾN NGHỊ ====
def create_excel(kien_nghi_list):
    wb = Workbook()
    ws = wb.active
    ws.title = "KPCS"

    columns = [
        "STT","Đối tượng được KT","Số văn bản","Ngày ban hành",
        "Tên Đoàn kiểm toán","Số hiệu rủi ro","Số hiệu kiểm soát",
        "Nghiệp vụ (R0)","Quy trình (R1)","Tên phát hiện (R2)",
        "Chi tiết phát hiện (R3)","Dẫn chiếu","Mô tả phát hiện",
        "CIF","Tên khách hàng","Loại KH","Số phát hiện/mẫu chọn",
        "Dư nợ sai phạm","Số tiền tổn thất","Số tiền cần thu hồi",
        "Trách nhiệm trực tiếp","Trách nhiệm quản lý",
        "Xếp hạng rủi ro","Xếp hạng kiểm soát",
        "Nguyên nhân","Ảnh hưởng","Kiến nghị",
        "Loại nguyên nhân","Loại ảnh hưởng","Loại kiến nghị",
        "Chủ thể kiến nghị","Kế hoạch thực hiện",
        "Trách nhiệm thực hiện","Đơn vị thực hiện KPCS",
        "ĐVKD/AMC/Hội sở","Người phê duyệt","Ý kiến đơn vị",
        "Mức độ ưu tiên","Thời hạn hoàn thành",
        "Đã khắc phục","Ngày đã KPCS","CBKT"
    ]

    for col_index, col_name in enumerate(columns, start=1):
        ws.cell(1, col_index, col_name)

    for i, kn in enumerate(kien_nghi_list, start=2):
        ws.cell(i, 1, i - 1)
        ws.cell(i, 27, kn)

        for col in range(2, len(columns) + 1):
            if col != 27:
                ws.cell(i, col, "")

    out = BytesIO()
    wb.save(out)
    return out


# ==== IMPORT KIẾN NGHỊ VÀO FILE CHÍNH ====
def merge_kien_nghi(file_main, file_new):

    wb_main = load_workbook(file_main)
    ws_main = wb_main.active

    wb_new = load_workbook(file_new)
    ws_new = wb_new.active

    # Lấy header
    header = {ws_main.cell(1, c).value: c for c in range(1, ws_main.max_column + 1)}

    # Tìm cột
    col_ngay_ban_hanh = None
    col_uu_tien = None
    col_deadline = None

    for name, idx in header.items():
        if not name:
            continue
        n = name.lower()
        if "ngày ban hành" in n:
            col_ngay_ban_hanh = idx
        if "mức độ ưu tiên" in n:
            col_uu_tien = idx
        if "thời hạn hoàn thành" in n:
            col_deadline = idx

    # Nếu thiếu deadline → thêm mới
    if not col_deadline:
        col_deadline = ws_main.max_column + 1
        ws_main.cell(1, col_deadline, "Thời hạn hoàn thành (mm/dd/yyyy)")

    # Ghép từng dòng
    for row in ws_new.iter_rows(min_row=2, values_only=False):
        new_values = [c.value for c in row]

        new_row = ws_main.max_row + 1
        for col_idx, val in enumerate(new_values, start=1):
            ws_main.cell(new_row, col_idx, val)

        # Lấy giá trị input
        ngay = ws_main.cell(new_row, col_ngay_ban_hanh).value
        uu_tien = ws_main.cell(new_row, col_uu_tien).value

        deadline = calc_deadline(ngay, uu_tien)

        ws_main.cell(new_row, col_deadline, deadline)

    out = BytesIO()
    wb_main.save(out)
    return out
