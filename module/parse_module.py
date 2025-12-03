import re

def parse_block_info(text: str):
    if text is None:
        text = ""
    txt = str(text)

    def find(label):
        pattern = rf"{label}\s*[:\-–]\s*(.*?)(?=\n[A-ZÀ-Ỵ]|$)"
        m = re.search(pattern, txt, flags=re.IGNORECASE | re.DOTALL)
        return m.group(1).strip() if m else ""

    ke_hoach = find("Kế hoạch thực hiện")
    nguoi_th = find("Người chịu trách nhiệm thực hiện")
    nguoi_duyet = find("Người phê duyệt")
    ngay_ht = find("Ngày hoàn thành")
    uu_tien_raw = find("Mức độ ưu tiên")

    # Chỉ lấy số tháng
    uu_tien = ""
    if uu_tien_raw:
        m = re.search(r"(\d+)\s*tháng", uu_tien_raw)
        uu_tien = m.group(1) if m else uu_tien_raw

    return {
        "Mức độ ưu tiên": uu_tien,
        "Kế hoạch thực hiện": ke_hoach,
        "Người chịu trách nhiệm thực hiện": nguoi_th,
        "Người phê duyệt": nguoi_duyet,
        "Ngày hoàn thành": ngay_ht,
    }
