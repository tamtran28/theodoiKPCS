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
