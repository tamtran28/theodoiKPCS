def create_excel(kien_nghi_list, doi_tuong, so_van_ban, ngay_ban_hanh):
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

    # Ghi header
    for col_index, col_name in enumerate(columns, start=1):
        ws.cell(1, col_index, col_name)

    # Ghi dữ liệu
    for i, kn in enumerate(kien_nghi_list, start=2):

        ws.cell(i, 1, i - 1)                 # STT
        ws.cell(i, 2, doi_tuong or "")       # Đối tượng KT
        ws.cell(i, 3, so_van_ban or "")      # Số văn bản
        ws.cell(i, 4, ngay_ban_hanh or "")   # Ngày ban hành
        ws.cell(i, 27, kn)                   # Kiến nghị nội dung

        # Các cột khác để trống
        for col in range(5, len(columns) + 1):
            if col != 27:
                ws.cell(i, col, "")

    out = BytesIO()
    wb.save(out)
    return out
