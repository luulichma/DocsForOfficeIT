# -*- coding: utf-8 -*-
"""
Module biên soạn PHỤ LỤC TẬP I:
- Phụ lục 1: Bảng đối chiếu thể thức chuẩn theo Nghị định 30/2020/NĐ-CP
- Phụ lục 2: Bảng tổng hợp phím tắt thông dụng (Windows, Word, Excel)
- Phụ lục 3: Cẩm nang sơ cứu 15 lỗi thường gặp nhất
"""

from docx_helper import (
    add_heading_1, add_heading_2, add_heading_3, add_p,
    add_styled_table, add_callout
)
from docx.shared import Mm

def build_appendix_1(doc):
    add_heading_1(doc, "PHỤ LỤC TẬP I: CẨM NANG TRA CỨU NHANH DÀNH CHO DÂN CÔNG SỞ")
    
    # PHỤ LỤC 1
    add_heading_2(doc, "Phụ lục 1: Bảng đối chiếu quy cách các thành phần thể thức theo Nghị định số 30/2020/NĐ-CP")
    add_p(doc, "Căn cứ pháp lý: Phụ lục I - Thể thức và kỹ thuật trình bày văn bản hành chính (Ban hành kèm theo Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ).")
    
    headers_nd30 = ["Thành phần thể thức", "Cỡ chữ", "Kiểu chữ", "Vị trí trình bày", "Tính chất"]
    rows_nd30 = [
        ["Quốc hiệu: CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", "12 - 13", "In hoa, Đứng, ĐẬM", "Góc trên bên phải, dòng 1", "Bắt buộc"],
        ["Tiêu ngữ: Độc lập - Tự do - Hạnh phúc", "13 - 14", "Chữ thường, Đứng, ĐẬM", "Ngay dưới Quốc hiệu, có đường kẻ bằng độ dài dòng chữ", "Bắt buộc"],
        ["Tên cơ quan chủ quản (nếu có)", "12 - 13", "In hoa, Đứng", "Góc trên bên trái, dòng 1", "Bắt buộc"],
        ["Tên cơ quan, tổ chức ban hành", "12 - 13", "In hoa, Đứng, ĐẬM", "Dưới cơ quan chủ quản, có đường kẻ bằng 1/3 - 1/2 độ dài", "Bắt buộc"],
        ["Số và ký hiệu văn bản", "13", "Chữ và số, Đứng", "Ngay dưới tên cơ quan ban hành", "Bắt buộc"],
        ["Địa danh và ngày tháng năm", "13 - 14", "Chữ thường, NGHIÊNG", "Dưới Tiêu ngữ, căn giữa theo Tiêu ngữ", "Bắt buộc"],
        ["Tên loại văn bản (Thông báo, Tờ trình...)", "14 - 15", "In hoa, Đứng, ĐẬM", "Đặt ở giữa trang", "Bắt buộc"],
        ["Trích yếu nội dung (văn bản có tên loại)", "13 - 14", "Chữ thường, Đứng, ĐẬM", "Đặt ngay dưới tên loại văn bản", "Bắt buộc"],
        ["Trích yếu công văn hành chính", "12 - 13", "Chữ thường, Đứng", "Đặt dưới số ký hiệu văn bản, có đường kẻ ngắn", "Bắt buộc"],
        ["Tên cơ quan nhận (Kính gửi:...)", "13 - 14", "Chữ thường, Đứng", "Sau trích yếu, căn lề trái", "Bắt buộc"],
        ["Nội dung văn bản (Phần thân)", "13 - 14", "Chữ thường, Đứng", "Căn đều 2 lề (Ctrl+J), thụt đầu dòng 1 - 1.27cm", "Bắt buộc"],
        ["Chức vụ người ký (GIÁM ĐỐC...)", "13 - 14", "In hoa, Đứng, ĐẬM", "Góc dưới bên phải", "Bắt buộc"],
        ["Họ và tên người ký", "13 - 14", "Chữ thường, Đứng, ĐẬM", "Dưới chức vụ, cách 3 - 4cm chừa ký đóng dấu", "Bắt buộc"],
        ["Từ 'Nơi nhận:'", "12", "Chữ thường, NGHIÊNG, ĐẬM", "Góc dưới bên trái, ngang hàng chức vụ người ký", "Bắt buộc"],
        ["Danh sách nơi nhận cụ thể", "11", "Chữ thường, Đứng", "Dưới từ 'Nơi nhận:', gạch đầu dòng sát lề", "Bắt buộc"],
        ["Số trang văn bản", "13 - 14", "Chữ số Ả Rập, Đứng", "Góc giữa lề trên trang giấy (Trang 1 không hiển thị)", "Bắt buộc"]
    ]
    col_widths_nd30 = [Mm(45), Mm(18), Mm(32), Mm(45), Mm(20)]
    add_styled_table(doc, headers_nd30, rows_nd30, col_widths_nd30)

    # PHỤ LỤC 2
    add_heading_2(doc, "Phụ lục 2: Bảng tổng hợp các phím tắt 'vàng' cho dân văn phòng")
    add_p(doc, "Thành thạo các phím tắt này sẽ giúp bạn tăng tốc độ làm việc lên gấp 3 đến 5 lần và giảm tối đa cảm giác mỏi tay khi sử dụng chuột liên tục.")

    headers_keys = ["Tổ hợp phím tắt", "Ứng dụng", "Tác dụng thực chiến"]
    rows_keys = [
        ["Ctrl + S", "Toàn hệ thống", "LƯU TỆP NGAY LẬP TỨC (Nên bấm 5 phút/lần để phòng mất điện)."],
        ["Ctrl + C", "Toàn hệ thống", "Sao chép (Copy) đối tượng đang chọn vào bộ nhớ tạm."],
        ["Ctrl + X", "Toàn hệ thống", "Cắt (Cut) đối tượng đang chọn để di chuyển sang vị trí mới."],
        ["Ctrl + V", "Toàn hệ thống", "Dán (Paste) đối tượng từ bộ nhớ tạm ra vị trí con trỏ."],
        ["Ctrl + Z", "Toàn hệ thống", "HOÀN TÁC (Undo) - Lấy lại ngay thao tác vừa lỡ tay làm sai."],
        ["Ctrl + Y", "Toàn hệ thống", "Làm lại (Redo) thao tác vừa bị Undo."],
        ["Ctrl + A", "Toàn hệ thống", "Chọn toàn bộ (bôi đen toàn bộ văn bản hoặc toàn bộ bảng tính)."],
        ["Ctrl + P", "Toàn hệ thống", "Mở cửa sổ In ấn tài liệu (Print Preview)."],
        ["Ctrl + F", "Toàn hệ thống", "Tìm kiếm từ khóa hoặc dữ liệu trong tài liệu."],
        ["Ctrl + H", "Word / Excel", "Tìm kiếm và Thay thế hàng loạt (Replace)."],
        ["Alt + Tab", "Windows", "Chuyển đổi qua lại tức thì giữa các cửa sổ ứng dụng đang mở."],
        ["Windows + D", "Windows", "Thu nhỏ toàn bộ cửa sổ để quay về màn hình chính Desktop."],
        ["Windows + E", "Windows", "Mở nhanh cửa sổ quản lý tệp tin File Explorer."],
        ["Ctrl + B", "Word / Excel", "Bật / tắt in đậm (Bold)."],
        ["Ctrl + I", "Word / Excel", "Bật / tắt in nghiêng (Italic)."],
        ["Ctrl + U", "Word / Excel", "Bật / tắt gạch chân (Underline)."],
        ["Ctrl + J", "Word", "CĂN ĐỀU HAI BÊN LỀ (Bắt buộc cho văn bản hành chính)."],
        ["Ctrl + E", "Word", "Căn chính giữa dòng (dùng cho Tiêu đề, Chữ ký)."],
        ["Ctrl + Enter", "Word", "Ngắt trang sang trang mới ngay lập tức (Page Break)."],
        ["F2", "Excel", "SỬA NỘI DUNG Ô mà không làm mất dữ liệu cũ."],
        ["Ctrl + Shift + L", "Excel", "Bật / tắt tính năng Lọc dữ liệu tự động (Filter)."],
        ["Alt + =", "Excel", "Tự động tính tổng nhanh (AutoSum)."]
    ]
    col_widths_keys = [Mm(35), Mm(28), Mm(97)]
    add_styled_table(doc, headers_keys, rows_keys, col_widths_keys)

    # PHỤ LỤC 3
    add_heading_2(doc, "Phụ lục 3: Cẩm nang sơ cứu 15 sự cố văn phòng thường gặp nhất")
    
    add_callout(doc, "BỘ QUY TẮC XỬ LÝ NHANH SỰ CỐ MÁY TÍNH DÀNH CHO NGƯỜI MỚI", [
        "1. Lỡ tay xóa nhầm một đoạn văn bản hoặc xóa nhầm 1 cột bảng tính: ĐỪNG HOẢNG SỢ! Hãy bấm ngay tổ hợp phím 'Ctrl + Z' một hoặc vài lần, mọi thứ sẽ quay trở lại nguyên vẹn.",
        "2. Máy tính bị biến mất thanh công cụ Ribbon trong Word/Excel: Nhấp đúp chuột trái vào bất kỳ tên thẻ nào (như thẻ Home hoặc thẻ Insert), thanh công cụ sẽ tự động ghim trở lại.",
        "3. Gõ tiếng Việt chữ 'D' biến thành 'Đ', gõ 'A' biến thành 'Â': Do UniKey đang để chế độ gõ tiếng Việt nhưng bạn lại muốn gõ tiếng Anh. Bấm tổ hợp 'Ctrl + Shift' để chuyển UniKey sang chữ E.",
        "4. Bàn phím số bên tay phải không gõ được: Nhìn xem đèn 'Num Lock' có sáng không. Nhấn phím 'Num Lock' 1 lần để bật bàn phím số.",
        "5. Ô Excel hiển thị lỗi #DIV/0!: Do công thức đang chia cho số 0 hoặc chia cho một ô rỗng. Kiểm tra lại ô mẫu số trong phép chia.",
        "6. Ô Excel hiển thị lỗi #VALUE!: Trong phép tính có chứa một ô là chữ. Hãy kiểm tra các ô tính xem có khoảng trắng hoặc chữ cái nào không.",
        "7. Bảng Word bị rớt 1-2 dòng sang trang mới: Vào Layout > Margins > Custom Margins > Giảm lề trên/dưới từ 2.0cm xuống 1.8cm, hoặc giảm khoảng cách đoạn Spacing After từ 6pt xuống 4pt, văn bản sẽ co gọn trong trang.",
        "8. Đang gõ văn bản tự nhiên các chữ phía sau bị 'nuốt chửng' mất dần: Do bạn vô tình bấm phím 'Insert' (chế độ gõ đè). Hãy nhấn phím 'Insert' trên bàn phím 1 lần nữa để tắt chế độ gõ đè.",
        "9. Tệp Word mở lên bị lỗi phông chữ loằng ngoằng không đọc được: Do văn bản cũ được gõ bằng phông .VnTime (Bảng mã TCVN3 cũ) hoặc VNI-Times. Hãy dùng công cụ chuyển mã của UniKey (Ctrl + Shift + F6) để chuyển toàn bộ về phông Unicode chuẩn Times New Roman.",
        "10. Quên chưa bấm Save mà bị sập nguồn máy tính: Khi bật lại Word/Excel, nhìn sang thanh bên trái màn hình mục 'Document Recovery' (Khôi phục tài liệu), chọn tệp tự lưu gần nhất để khôi phục lại bài.",
        "11. Con chuột bị đơ hoặc không di chuyển được: Thử rút đầu cắm USB của chuột cắm sang một cổng USB khác trên thùng máy; nếu là chuột không dây, hãy kiểm tra và thay pin mới.",
        "12. Không tìm thấy tệp tài liệu vừa tải về: Bấm phím 'Windows + E' mở File Explorer > Nhấp chuột vào thư mục 'Downloads' ở danh mục bên trái.",
        "13. Tệp bị thông báo 'The file is open in another program' không cho xóa hoặc đổi tên: Tệp đó đang được mở trong Word hoặc Excel. Hãy kiểm tra thanh Taskbar phía dưới và bấm đóng tệp đó lại trước khi xóa.",
        "14. Chữ trong ô Excel bị tràn sang ô bên cạnh mà không nhìn thấy hết: Nhấp chuột chọn ô đó rồi bấm nút 'Wrap Text' trên thẻ Home.",
        "15. In văn bản bị mờ nhạt: Kiểm tra xem máy in có bị hết mực không, hoặc vào cài đặt máy in xem có đang vô tình bật chế độ 'Draft / Eco Mode' (In tiết kiệm mực) không."
    ], callout_type="warning")
