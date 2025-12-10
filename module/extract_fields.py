import re


def _extract_field(text: str, label: str):
    pattern = rf"{label}\s*[:\-–]\s*(.*?)(?=\n[A-ZÀ-Ỵ]|$)"
    m = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    return m.group(1).strip() if m else ""


def extract_all_fields(text: str):
    nguyen_nhan = _extract_field(text, "Nguyên nhân")
    uu_tien_raw = _extract_field(text, "Mức độ ưu tiên hành động")
    nguoi_th = _extract_field(text, "Người chịu trách nhiệm thực hiện")
    nguoi_phe_duyet = _extract_field(text, "Người phê duyệt")
    ngay_hoan_thanh = _extract_field(text, "Ngày hoàn thành")

    # Lấy số tháng
    uu_tien = ""
    if uu_tien_raw:
        m = re.search(r"(\d+)\s*tháng", uu_tien_raw, flags=re.IGNORECASE)
        uu_tien = m.group(1) if m else uu_tien_raw

    return {
        "nguyen_nhan": nguyen_nhan,
        "uu_tien": uu_tien,
        "nguoi_thuc_hien": nguoi_th,
        "nguoi_phe_duyet": nguoi_phe_duyet,
        "ngay_hoan_thanh": ngay_hoan_thanh,
    }

# # extract_fields.py
# import re


# def _extract_field(text: str, label: str):
#     """
#     Lấy phần nội dung sau 'label:' cho đến trước dòng bắt đầu bằng chữ in hoa khác.
#     Ví dụ: label = 'Nguyên nhân'
#     """
#     pattern = rf"{label}\s*[:\-–]\s*(.*?)(?=\n[A-ZÀ-ỴÝÊÂĂĐ]|$)"
#     m = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
#     if m:
#         return m.group(1).strip()
#     return ""


# def extract_all_fields(kien_nghi_text: str):
#     """
#     Trả về dict các trường được tách ra; nếu không có thì là chuỗi rỗng.
#     """
#     nguyen_nhan = _extract_field(kien_nghi_text, "Nguyên nhân")
#     uu_tien_raw = _extract_field(kien_nghi_text, "Mức độ ưu tiên hành động")
#     nguoi_th = _extract_field(kien_nghi_text, "Người chịu trách nhiệm thực hiện")
#     nguoi_phe_duyet = _extract_field(kien_nghi_text, "Người phê duyệt")
#     ngay_hoan_thanh = _extract_field(kien_nghi_text, "Ngày hoàn thành")

#     # chuẩn hóa ưu tiên: lấy số tháng nếu có dạng "03 tháng", "2 tháng"
#     uu_tien = ""
#     if uu_tien_raw:
#         m = re.search(r"(\d+)\s*tháng", uu_tien_raw, flags=re.IGNORECASE)
#         if m:
#             uu_tien = m.group(1)
#         else:
#             uu_tien = uu_tien_raw

#     return {
#         "nguyen_nhan": nguyen_nhan,
#         "uu_tien": uu_tien,
#         "nguoi_thuc_hien": nguoi_th,
#         "nguoi_phe_duyet": nguoi_phe_duyet,
#         "ngay_hoan_thanh": ngay_hoan_thanh,
#     }
