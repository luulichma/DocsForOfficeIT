# BỘ GIÁO TRÌNH TIN HỌC VĂN PHÒNG THỰC CHIẾN (DÀNH CHO NGƯỜI MỚI BẮT ĐẦU)

**Người biên soạn:** Nguyễn Thế Chiến  
**Lưu hành:** Nội bộ — Năm 2026  

Bộ tài liệu hướng dẫn từng bước (cầm tay chỉ việc) về **Windows, Microsoft Word, Microsoft Excel, Google Docs và Google Sheets** dành cho người chưa thành thạo máy tính, cán bộ công chức, viên chức và nhân viên văn phòng.

---

## 📁 CÁC TỆP TÀI LIỆU DOCX HOÀN CHỈNH ĐÃ XUẤT BẢN

Tài liệu được đóng gói thành **02 tập DOCX độc lập**, định dạng chuẩn A4, phông chữ Times New Roman, tuân thủ nghiêm ngặt **Nghị định số 30/2020/NĐ-CP** ngày 05/3/2020 của Chính phủ về công tác văn thư:

### 1. [Tập I: Kỹ năng Nền tảng, Soạn thảo Hành chính và Bảng tính Văn phòng](./Tap_I_Ky_Nang_Nen_Tang_Word_Excel.docx)
*   **Phần 1: Kỹ năng máy tính nền tảng & Windows văn phòng** (Bài 1 - 4: Thao tác chuột, bàn phím, gõ UniKey Telex, quản lý tệp thư mục ổ D, trình duyệt web, xuất PDF, gửi email công sở).
*   **Phần 2: Soạn thảo văn bản hành chính chuẩn với Microsoft Word** (Bài 5 - 10: Khổ A4, lề chuẩn NĐ 30, kỹ thuật dùng bảng ẩn viền căn Quốc hiệu - Tiêu ngữ - Chữ ký - Nơi nhận, thụt lề 1cm, căn đều 2 bên Ctrl+J, đánh số trang từ trang 2, thực hành 3 văn bản: Công văn, Thông báo, Tờ trình).
*   **Phần 3: Quản lý & xử lý bảng tính văn phòng với Microsoft Excel** (Bài 11 - 17: Nhập liệu số/chữ/ngày tháng, định dạng bảng, tiền tệ VNĐ, công thức SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, hàm điều kiện IF, tô màu cảnh báo Conditional Formatting, lọc Filter, cố định dòng Freeze Panes, in A4 vừa khít 1 trang, thực hành Sổ theo dõi hồ sơ & Bảng chi phí văn phòng phẩm).
*   **Phụ lục Tập I:** Bảng đối chiếu thể thức NĐ 30/2020/NĐ-CP, Bảng phím tắt vàng, Cẩm nang sơ cứu 15 lỗi thường gặp.

### 2. [Tập II: Soạn thảo, Quản lý Bảng tính và Cộng tác Trực tuyến trên Google Workspace](./Tap_II_Google_Docs_Google_Sheets.docx)
*   **Phần 1: Soạn thảo & cộng tác văn bản trực tuyến với Google Docs** (Bài 1 - 5: Lệnh docs.new, cơ chế Auto-save, chuyển giao Word <-> Docs, căn lề A4 chuẩn NĐ 30 trên web, phân quyền Xem/Nhận xét/Sửa, bình luận @email, đề xuất sửa Suggesting, lịch sử phiên bản Version History, thực hành Dự thảo Kế hoạch tuần).
*   **Phần 2: Bảng tính trực tuyến & phân tích số liệu với Google Sheets** (Bài 6 - 9: Lệnh sheets.new, định dạng vùng Việt Nam, hàm thông minh Smart Suggestions, bộ lọc xem riêng Filter Views không làm phiền người khác, khóa bảo vệ ô công thức Protect ranges, thực hành Bảng phân công tiến độ công việc).
*   **Phụ lục Tập II:** Bảng phím tắt Google Docs/Sheets, 5 nguyên tắc bảo mật và an toàn dữ liệu công vụ đám mây.

---

## 🛠 HỆ THỐNG MÃ NGUỒN TẠO SINH TÀI LIỆU (BUILDER)

Toàn bộ tài liệu được sinh tự động bằng Python (`python-docx`) bảo đảm tính thẩm mỹ, nhất quán và dễ dàng chỉnh sửa hoặc tái xuất bản:

*   `builder/docx_helper.py`: Thư viện định dạng kiểu dáng (Margins A4, Colors, Callout boxes, Tables, Checklists, Page numbers).
*   `builder/content_part1.py`: Nội dung Phần 1 Tập I.
*   `builder/content_part2.py`: Nội dung Phần 2 Tập I (Word Hành chính).
*   `builder/content_part3.py`: Nội dung Phần 3 Tập I (Excel Bảng tính).
*   `builder/content_appendix1.py`: Nội dung Phụ lục Tập I.
*   `builder/generate_volume_1.py`: Chương trình xuất bản Tập I.
*   `builder/content_vol2_part1.py`: Nội dung Phần 1 Tập II (Google Docs).
*   `builder/content_vol2_part2.py`: Nội dung Phần 2 Tập II (Google Sheets).
*   `builder/content_vol2_appendix.py`: Nội dung Phụ lục Tập II.
*   `builder/generate_volume_2.py`: Chương trình xuất bản Tập II.

### Lệnh tái tạo tài liệu khi cần:
```powershell
py "builder/generate_volume_1.py"
py "builder/generate_volume_2.py"
```