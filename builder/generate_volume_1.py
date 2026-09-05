# -*- coding: utf-8 -*-
"""
Chương trình xuất bản TẬP I:
KỸ NĂNG MÁY TÍNH NỀN TẢNG, SOẠN THẢO VĂN BẢN HÀNH CHÍNH VÀ BẢNG TÍNH VĂN PHÒNG
(Windows + Microsoft Word + Microsoft Excel)
"""

import os
import sys

# Đảm bảo in tiếng Việt ra console không bị lỗi mã hóa
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Thêm thư mục hiện tại vào sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

from docx_helper import (
    init_document, add_cover_page, add_heading_1, add_heading_2,
    add_heading_3, add_p, add_callout, add_styled_table
)
from content_part1 import build_part_1
from content_part2 import build_part_2
from content_part3 import build_part_3
from content_appendix1 import build_appendix_1

def generate_volume_1():
    print(">>> Đang khởi tạo tài liệu Tập I...")
    doc = init_document()
    
    # 1. Trang bìa
    add_cover_page(
        doc,
        volume_number="I",
        volume_title="KỸ NĂNG MÁY TÍNH NỀN TẢNG, SOẠN THẢO HÀNH CHÍNH VÀ BẢNG TÍNH VĂN PHÒNG",
        subtitle="Cẩm nang thực hành cầm tay chỉ việc: Windows — Microsoft Word — Microsoft Excel\nChuẩn hóa theo Nghị định số 30/2020/NĐ-CP của Chính phủ",
        author="Nguyễn Thế Chiến"
    )
    
    # 2. Lời nói đầu & Hướng dẫn học tập
    add_heading_1(doc, "LỜI NÓI ĐẦU & HƯỚNG DẪN SỬ DỤNG TÀI LIỆU")
    add_p(doc, "Chào mừng bạn đến với bộ tài liệu thực hành tin học văn phòng dành cho người mới bắt đầu. Bộ tài liệu này được biên soạn với phương châm cốt lõi: 'Cầm tay chỉ việc - Không lý thuyết suông - Học đến đâu làm được việc ngay đến đó'.")
    add_p(doc, "Trong bối cảnh chuyển đổi số và hiện đại hóa hành chính công hiện nay, việc làm chủ máy tính, tự tay soạn thảo được một văn bản hành chính đúng chuẩn pháp lý và lập được các bảng tính theo dõi công việc là yêu cầu bắt buộc đối với mỗi cán bộ, công chức, viên chức và người lao động văn phòng.")
    
    add_callout(doc, "CẤU TRÚC 7 BƯỚC CỦA MỖI BÀI HỌC", [
        "1. Mục tiêu bài học: Giúp bạn biết rõ kết quả cụ thể mình sẽ đạt được.",
        "2. Điều kiện chuẩn bị: Thiết bị, phần mềm và môi trường thực hành.",
        "3. Thao tác từng bước: Hướng dẫn chi tiết từng cú nhấp chuột (trái/phải/đúp), từng phím tắt, chú giải nút lệnh song ngữ (Tiếng Anh kèm Tiếng Việt).",
        "4. Kết quả mong đợi: Tiêu chuẩn thẩm mỹ và kỹ thuật của sản phẩm hoàn thành.",
        "5. Lỗi thường gặp và cách khắc phục: Cảnh báo trước các 'cái bẫy' người mới hay mắc phải.",
        "6. Bài thực hành tự rèn luyện: Bộ bài tập với dữ liệu giả lập thực tế để tự luyện.",
        "7. Bảng tiêu chí tự kiểm tra (Checklist): Bảng đánh giá để bạn tự chấm điểm mức độ hoàn thành của mình."
    ], callout_type="note")
    
    add_p(doc, "Lời khuyên dành cho bạn: Đừng đọc tài liệu như một cuốn sách lý thuyết! Hãy mở máy tính lên, đặt tài liệu bên cạnh hoặc chia đôi màn hình, và thực hành từng bước một. Nếu gặp sự cố, hãy bình tĩnh tra cứu mục 'Lỗi thường gặp' và 'Phụ lục 3: Cẩm nang sơ cứu' ở cuối sách. Chúc bạn có một hành trình học tập đầy hứng khởi và gặt hái nhiều kỹ năng bổ ích!")

    # 3. Mục lục tổng quát
    add_heading_1(doc, "MỤC LỤC TẬP I")
    toc_headers = ["Phần / Bài học", "Nội dung trọng tâm"]
    toc_rows = [
        ["PHẦN 1", "KỸ NĂNG MÁY TÍNH NỀN TẢNG & WINDOWS VĂN PHÒNG"],
        ["Bài 1", "Làm quen phần cứng, thao tác điều khiển chuột, bàn phím và quản lý cửa sổ"],
        ["Bài 2", "Làm chủ bộ gõ tiếng Việt UniKey và nguyên tắc gõ văn bản chuẩn"],
        ["Bài 3", "Quản lý thư mục, tệp tin khoa học và lưu trữ dữ liệu an toàn trên ổ D"],
        ["Bài 4", "Trình duyệt Web, tải dữ liệu, xuất tệp PDF và gửi Email công sở"],
        ["PHẦN 2", "SOẠN THẢO VĂN BẢN HÀNH CHÍNH CHUẨN VỚI MICROSOFT WORD"],
        ["Bài 5", "Thiết lập trang văn bản chuẩn Nghị định 30/2020/NĐ-CP trước khi soạn thảo"],
        ["Bài 6", "Kỹ thuật căn chỉnh các khối thể thức văn bản hành chính bằng Bảng ẩn viền"],
        ["Bài 7", "Trình bày nội dung văn bản, phân cấp mục và thao tác đoạn văn chuyên nghiệp"],
        ["Bài 8", "Bố trí phần Ký duyệt, Nơi nhận và Bảng biểu số liệu trong văn bản"],
        ["Bài 9", "Đánh số trang theo Nghị định 30, kiểm tra hoàn thiện và In ấn / Xuất PDF"],
        ["Bài 10", "Dự án thực hành Word: Soạn thảo 03 văn bản hành chính (Công văn, Thông báo, Tờ trình)"],
        ["PHẦN 3", "QUẢN LÝ & XỬ LÝ BẢNG TÍNH VĂN PHÒNG VỚI MICROSOFT EXCEL"],
        ["Bài 11", "Làm quen giao diện Excel và Nguyên tắc nhập liệu an toàn (Chữ, Số, Ngày tháng)"],
        ["Bài 12", "Định dạng bảng biểu chuyên nghiệp, Kẻ khung và Định dạng tiền tệ"],
        ["Bài 13", "Các phép tính số học và Các hàm thống kê cơ bản nhất (SUM, AVERAGE, MIN, MAX, COUNT)"],
        ["Bài 14", "Hàm điều kiện IF và Định dạng cảnh báo màu tự động (Conditional Formatting)"],
        ["Bài 15", "Sắp xếp, Lọc dữ liệu văn phòng (Filter) và Cố định dòng tiêu đề (Freeze Panes)"],
        ["Bài 16", "Căn chỉnh trang in bảng tính Excel lên khổ giấy A4 vừa khít"],
        ["Bài 17", "Dự án thực hành Excel: Xây dựng Sổ theo dõi hồ sơ & Bảng thanh toán văn phòng phẩm"],
        ["PHỤ LỤC", "CẨM NANG TRA CỨU NHANH DÀNH CHO DÂN CÔNG SỞ"],
        ["Phụ lục 1", "Bảng đối chiếu quy cách thể thức theo Nghị định số 30/2020/NĐ-CP"],
        ["Phụ lục 2", "Bảng tổng hợp các phím tắt 'vàng' cho dân văn phòng (Windows, Word, Excel)"],
        ["Phụ lục 3", "Cẩm nang sơ cứu 15 sự cố văn phòng thường gặp nhất"]
    ]
    add_styled_table(doc, toc_headers, toc_rows)
    doc.add_page_break()

    # 4. Biên soạn Phần 1
    print(">>> Đang biên soạn Phần 1 (Nền tảng Windows)...")
    build_part_1(doc)
    doc.add_page_break()

    # 5. Biên soạn Phần 2
    print(">>> Đang biên soạn Phần 2 (Microsoft Word Hành chính)...")
    build_part_2(doc)
    doc.add_page_break()

    # 6. Biên soạn Phần 3
    print(">>> Đang biên soạn Phần 3 (Microsoft Excel Bảng tính)...")
    build_part_3(doc)
    doc.add_page_break()

    # 7. Biên soạn Phụ lục
    print(">>> Đang biên soạn Phụ lục Tập I...")
    build_appendix_1(doc)

    # 8. Lưu tệp
    output_path = os.path.join(os.path.dirname(current_dir), "Tap_I_Ky_Nang_Nen_Tang_Word_Excel.docx")
    print(f">>> Đang lưu tệp tại: {output_path}")
    doc.save(output_path)
    print(">>> XUẤT BẢN TẬP I THÀNH CÔNG!")
    return output_path

if __name__ == "__main__":
    generate_volume_1()
