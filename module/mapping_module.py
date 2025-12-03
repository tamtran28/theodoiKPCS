from module.parse_module import parse_block_info
import pandas as pd

def build_output_df(summary_df, detail_df, map_summary, map_detail, block_col):

    n = min(len(summary_df), len(detail_df))
    rows = []

    for i in range(n):
        s = summary_df.iloc[i]
        d = detail_df.iloc[i]

        row = {
            "Tên phát hiện": s.get(map_summary["ten_phat_hien"], ""),
            "Ảnh hưởng (tóm tắt)": s.get(map_summary["anh_huong"], ""),
            "Xếp hạng rủi ro": s.get(map_summary["xep_rr"], ""),
            "Xếp hạng kiểm soát": s.get(map_summary["xep_ks"], ""),
            "Số lượng chi tiết phát hiện": s.get(map_summary["so_luong"], ""),

            "Phát hiện & Nguyên nhân": d.get(map_detail["phat_hien_nn"], ""),
            "Ảnh hưởng (chi tiết)": d.get(map_detail["anh_huong"], ""),
            "Kiến nghị": d.get(map_detail["kien_nghi"], ""),
            "Ý kiến đơn vị được kiểm toán": d.get(map_detail["y_kien"], ""),

            "Mức độ ưu tiên": "",
            "Kế hoạch thực hiện": "",
            "Người chịu trách nhiệm thực hiện": "",
            "Người phê duyệt": "",
            "Ngày hoàn thành": "",
        }

        if block_col:
            info = parse_block_info(d.get(block_col, ""))
            for k, v in info.items():
                if k in row:
                    row[k] = v

        rows.append(row)

    return pd.DataFrame(rows)
