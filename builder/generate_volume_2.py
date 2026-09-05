# -*- coding: utf-8 -*-
"""
Chương trình xuất bản TẬP II:
SOẠN THẢO, QUẢN LÝ BẢNG TÍNH VÀ CỘNG TÁC TRỰC TUYẾN TRÊN GOOGLE WORKSPACE
(Google Docs + Google Sheets trên nền tảng Web Windows)
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
from content_vol2_part1 import build_vol2_part_1
from content_vol2_part2 import build_vol2_part_2
from content_vol2_appendix import build_vol2_appendix

def generate_volume_2():
    print(">>> Đang khởi tạo tài liệu Tập II...")
    doc = init_document()
    
    # 1. Trang bìa
    add_cover_page(
        doc,
        volume_number="II",
        volume_title="SOẠN THẢO, QUẢN LÝ BẢNG TÍNH VÀ CỘNG TÁC TRỰC TUYẾN TRÊN GOOGLE WORKSPACE",
        subtitle="Cẩm nang thực hành cầm tay chỉ việc: Google Docs — Google Sheets\nGiải pháp làm việc nhóm, chia sẻ dữ liệu và chuyển đổi số văn phòng thời gian thực",
        author="Nguyễn Thế Chiến"
    )
    
    # 2. Lời nói đầu & Hướng dẫn học tập
    add_heading_1(doc, "LỜI NÓI ĐẦU & HƯỚNG DẪN HỌC TẬP TẬP II")
    add_p(doc, "Chào mừng bạn đến với Tập II của bộ giáo trình tin học văn phòng thực chiến. Nếu như Tập I đã giúp bạn xây dựng nền móng vững chắc với hệ điều hành Windows, Microsoft Word và Microsoft Excel truyền thống, thì Tập II sẽ đưa bạn bước vào không gian làm việc số hiện đại với Google Docs và Google Sheets.")
    add_p(doc, "Xu hướng làm việc công sở hiện đại đòi hỏi tốc độ xử lý nhanh, khả năng phối hợp liên phòng ban và cộng tác từ xa. Bạn sẽ không còn phải gửi từng tệp đính kèm qua lại rồi lo lắng xem ai đang giữ bản mới nhất. Toàn bộ công việc được thực hiện trên đám mây, tự động lưu từng giây, bảo đảm an toàn và bảo mật thông tin tuyệt đối.")
    
    add_callout(doc, "ĐIỂM KHÁC BIỆT CỐT LÕI CỦA TẬP II", [
        "1. Kế thừa kỹ năng chuẩn thể thức: Toàn bộ quy định về lề A4, phông Times New Roman và bảng ẩn viền từ Nghị định 30/2020/NĐ-CP ở Tập I được áp dụng trọn vẹn trên nền tảng web.",
        "2. Không còn nỗi lo mất bài: Cơ chế Auto-save tự động lưu liên tục lên Google Drive.",
        "3. Làm việc nhóm không xung đột: Nắm vững kỹ thuật phân quyền (Xem/Nhận xét/Chỉnh sửa), chế độ Đề xuất (Suggesting), Lịch sử phiên bản (Version History) và Bộ lọc xem riêng (Filter Views).",
        "4. Cầm tay chỉ việc: Mỗi bài học đều có đủ 7 bước sư phạm, giải thích rõ thao tác chuột/phím và đánh dấu vị trí ảnh chụp minh họa."
    ], callout_type="note")
    
    add_p(doc, "Hãy mở trình duyệt web Chrome, đăng nhập tài khoản Google và bắt đầu trải nghiệm sức mạnh vượt trội của văn phòng số!")

    # 3. Mục lục tổng quát
    add_heading_1(doc, "MỤC LỤC TẬP II")
    toc_headers = ["Phần / Bài học", "Nội dung trọng tâm"]
    toc_rows = [
        ["PHẦN 1", "SOẠN THẢO VÀ CỘNG TÁC VĂN BẢN TRỰC TUYẾN VỚI GOOGLE DOCS"],
        ["Bài 1", "Khởi tạo Google Docs và Cơ chế lưu trữ đám mây tự động (Auto-save)"],
        ["Bài 2", "Soạn thảo và Định dạng thể thức văn bản hành chính trên Google Docs"],
        ["Bài 3", "Làm việc nhóm, Phân quyền (Xem/Nhận xét/Sửa) và Cộng tác thời gian thực"],
        ["Bài 4", "Lịch sử phiên bản (Version History) và Khôi phục tài liệu khi bị sửa nhầm"],
        ["Bài 5", "Dự án thực hành Docs: Phối hợp soạn thảo và góp ý Dự thảo Kế hoạch tuần"],
        ["PHẦN 2", "BẢNG TÍNH TRỰC TUYẾN VÀ PHÂN TÍCH SỐ LIỆU VỚI GOOGLE SHEETS"],
        ["Bài 6", "Khởi tạo Google Sheets, Nhập liệu và Định dạng bảng biểu trực tuyến"],
        ["Bài 7", "Áp dụng công thức tính toán và Hàm văn phòng trên Google Sheets (SUM, IF...)"],
        ["Bài 8", "Tạo Bộ lọc xem riêng (Filter Views) và Khóa bảo vệ dải ô dữ liệu"],
        ["Bài 9", "Dự án thực hành Sheets: Xây dựng Bảng phân công & Theo dõi tiến độ phòng ban"],
        ["PHỤ LỤC", "CẨM NANG CỘNG TÁC SỐ VÀ BẢO MẬT DỮ LIỆU ĐÁM MÂY"],
        ["Phụ lục 1", "Bảng tổng hợp phím tắt quyền năng trên Google Docs và Google Sheets"],
        ["Phụ lục 2", "Quy tắc vàng bảo đảm an toàn và bảo mật dữ liệu công vụ trên đám mây"]
    ]
    add_styled_table(doc, toc_headers, toc_rows)
    doc.add_page_break()

    # 4. Biên soạn Phần 1 (Google Docs)
    print(">>> Đang biên soạn Phần 1 (Google Docs)...")
    build_vol2_part_1(doc)
    doc.add_page_break()

    # 5. Biên soạn Phần 2 (Google Sheets)
    print(">>> Đang biên soạn Phần 2 (Google Sheets)...")
    build_vol2_part_2(doc)
    doc.add_page_break()

    # 6. Biên soạn Phụ lục
    print(">>> Đang biên soạn Phụ lục Tập II...")
    build_vol2_appendix(doc)

    # 7. Lưu tệp
    output_path = os.path.join(os.path.dirname(current_dir), "Tap_II_Google_Docs_Google_Sheets.docx")
    print(f">>> Đang lưu tệp tại: {output_path}")
    doc.save(output_path)
    print(">>> XUẤT BẢN TẬP II THÀNH CÔNG!")
    return output_path

if __name__ == "__main__":
    generate_volume_2()
