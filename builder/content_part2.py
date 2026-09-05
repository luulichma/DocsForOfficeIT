# -*- coding: utf-8 -*-
"""
Module biên soạn PHẦN 2: SOẠN THẢO VĂN BẢN HÀNH CHÍNH CHUẨN VỚI MICROSOFT WORD
Theo quy định tại Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư.
Dành cho Tập I.
"""

from docx_helper import (
    add_heading_1, add_heading_2, add_heading_3, add_p,
    add_step, add_callout, add_screenshot_placeholder,
    add_styled_table, add_checklist_table
)

def build_part_2(doc):
    # PHẦN 2
    add_heading_1(doc, "PHẦN 2: SOẠN THẢO VĂN BẢN HÀNH CHÍNH CHUẨN VỚI MICROSOFT WORD")
    add_p(doc, "Phần 2 là trái tim của kỹ năng văn phòng trong cơ quan nhà nước, tổ chức và doanh nghiệp. Toàn bộ các bài học được thiết kế bám sát tuyệt đối quy định pháp lý hiện hành tại Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư (Phụ lục I: Thể thức và kỹ thuật trình bày văn bản hành chính). Người học sẽ nắm vững từ việc thiết lập lề trang, font chữ, sử dụng bảng ẩn viền để căn các khối thể thức, đến việc hoàn thiện và in ấn 03 loại văn bản thông dụng nhất: Công văn, Thông báo và Tờ trình.")

    # BÀI 5
    add_heading_2(doc, "Bài 5: Thiết lập trang văn bản chuẩn Nghị định 30/2020/NĐ-CP trước khi soạn thảo")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cách cài đặt đơn vị đo trong Word từ Inch (inch) sang Centimet (cm).")
    add_p(doc, "• Thiết lập chính xác khổ giấy A4 và thông số căn lề 4 phía đúng quy định pháp luật.")
    add_p(doc, "• Thiết lập phông chữ Times New Roman, cỡ chữ, khoảng cách dòng và khoảng cách đoạn văn bản đạt chuẩn.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Mở phần mềm Microsoft Word (bản 2016, 2019, 2021 hoặc Microsoft 365).")
    add_p(doc, "• Nhấp chuột chọn 'Blank document' (Tài liệu trống) để bắt đầu một trang mới.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "5.1", "Đổi đơn vị đo thước kẻ sang Centimeters (cm)", [
        "Vào thẻ File (Tệp) ở góc trên bên trái > nhấp chuột chọn 'Options' (Tùy chọn) ở dưới đáy.",
        "Cửa sổ Word Options hiện ra, nhấp chuột trái vào mục 'Advanced' (Nâng cao) ở danh mục bên trái.",
        "Cuộn chuột xuống tìm nhóm mục 'Display' (Hiển thị).",
        "Tại dòng 'Show measurements in units of' (Hiển thị đơn vị đo bằng): Nhấp chuột chọn 'Centimeters'.",
        "Nhấp chuột vào nút 'OK' ở đáy cửa sổ để lưu lại."
    ])

    add_callout(doc, "CĂN CỨ PHÁP LÝ BẮT BUỘC: NGHỊ ĐỊNH SỐ 30/2020/NĐ-CP (PHỤ LỤC I)", [
        "Khổ giấy: Khổ A4 (kích thước 210 mm x 297 mm), định hướng trang đứng (Portrait).",
        "Quy định căn lề trang văn bản (Điểm 1 Mục II Phần I Phụ lục I):",
        "  • Lề trên (Top): Cách mép trên từ 20 đến 25 mm (2.0 cm - 2.5 cm).",
        "  • Lề dưới (Bottom): Cách mép dưới từ 20 đến 25 mm (2.0 cm - 2.5 cm).",
        "  • Lề trái (Left): Cách mép trái từ 30 đến 35 mm (3.0 cm - 3.5 cm) - chừa rộng để đóng ghim/đóng bìa hồ sơ.",
        "  • Lề phải (Right): Cách mép phải từ 15 đến 20 mm (1.5 cm - 2.0 cm).",
        "Phông chữ (Font): Phông chữ tiếng Việt Times New Roman, bộ mã ký tự Unicode (tiêu chuẩn TCVN 6909:2001)."
    ], callout_type="legal")

    add_step(doc, "5.2", "Cài đặt Khổ giấy A4 và Căn lề chuẩn (Margins)", [
        "Trên thanh công cụ Ribbon, nhấp chuột trái vào thẻ 'Layout' (Bố trí) hoặc 'Page Layout'.",
        "Mục Khổ giấy: Nhấp vào nút 'Size' (Kích cỡ) > chọn đúng 'A4 (21 x 29.7 cm)'.",
        "Mục Căn lề: Nhấp vào nút 'Margins' (Lề) > chọn dòng cuối cùng là 'Custom Margins...' (Lề tùy chỉnh).",
        "Hộp thoại Page Setup hiện ra, bạn nhập chính xác các thông số tối ưu được khuyến nghị dùng trong toàn bộ cơ quan:",
        "  • Top (Trên): 2 cm",
        "  • Bottom (Dưới): 2 cm",
        "  • Left (Trái): 3 cm",
        "  • Right (Phải): 2 cm (hoặc 1.5 cm)",
        "Bấm nút 'OK' để hoàn tất. (Nếu muốn mọi văn bản mới sau này đều dùng lề này, bấm nút 'Set As Default' ở góc dưới bên trái)."
    ])

    add_screenshot_placeholder(doc, "Hộp thoại Page Setup trong Word: Cài đặt Top 2cm, Bottom 2cm, Left 3cm, Right 2cm, Paper size A4, khoanh đỏ nút OK.")

    add_step(doc, "5.3", "Cài đặt Giãn dòng và Khoảng cách đoạn văn", [
        "Nhấp chuột trái vào thẻ 'Home' > Trong nhóm 'Paragraph' (Đoạn văn), nhấp chuột vào dấu mũi tên nhỏ ở góc dưới bên phải nhóm Paragraph để mở hộp thoại Paragraph.",
        "Tại mục 'Line spacing' (Khoảng cách giữa các dòng): Nhấp chọn 'Multiple' và gõ số '1.3' (hoặc chọn 1.15 đến 1.35 theo quy định).",
        "Tại mục 'Spacing' (Khoảng cách giữa các đoạn):",
        "  • Before (Đoạn trước): 0 pt (hoặc 3 pt)",
        "  • After (Đoạn sau): 4 pt (hoặc 6 pt) để các đoạn văn tách bạch thoáng đãng, không bị dính sát vào nhau.",
        "Bấm nút 'OK'."
    ])

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Trang Word trắng tinh đã sẵn sàng với khổ A4, lề trái 3cm thoáng đãng để đóng ghim, lề phải và trên dưới 2cm cân đối, phông chữ Times New Roman sẵn sàng để gõ nội dung.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI SAI KHỔ GIẤY KHI IN ẤN", [
        "Lỗi in ra bị mất chữ dưới đáy trang hoặc khoảng trắng đáy quá lớn: Do phần mềm Word mặc định để khổ giấy 'Letter' (khổ chuẩn của Mỹ) thay vì 'A4'. Khổ Letter ngắn hơn và rộng hơn A4. Cách khắc phục: Luôn vào Layout > Size > Chọn lại đúng 'A4' ngay từ đầu.",
        "Lỗi gõ số lề bị báo lỗi 'Number must be between...': Do máy tính chưa chuyển sang đơn vị cm, khi bạn gõ số 2 máy hiểu là 2 inch (bằng hơn 5cm) nên báo quá rộng. Hãy xem lại Bước 5.1 để đổi đơn vị đo sang cm."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở một tệp Word mới. Thực hiện đổi đơn vị đo sang cm, cài đặt Margins (Top: 2.0 cm, Bottom: 2.0 cm, Left: 3.0 cm, Right: 1.5 cm), chọn Paper Size A4. Gõ một đoạn văn bản ngắn 3 dòng và lưu tệp vào thư mục 'D:\\HOC_TAP_VAN_PHONG\\01_VAN_BAN_WORD\\Bai5_Thiet_lap_trang.docx'.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Đơn vị đo trên thước kẻ hiển thị theo cm", "Các số hiển thị trên thanh thước là 1, 2, 3... cm"),
        ("Khổ giấy chọn đúng A4 (21 x 29.7 cm)", "Trong mục Layout > Size hiển thị dấu chọn ở A4"),
        ("Căn lề chuẩn 4 phía: Trái 3cm, Phải 1.5-2cm, Trên 2cm, Dưới 2cm", "Khoảng cách lề trên thước kẻ đúng kích thước"),
        ("Phông chữ mặc định là Times New Roman", "Ô Font chữ hiển thị Times New Roman")
    ])

    # BÀI 6
    add_heading_2(doc, "Bài 6: Kỹ thuật căn chỉnh các khối thể thức văn bản hành chính bằng Bảng ẩn viền")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cấu tạo chính xác của phần đầu văn bản hành chính theo Nghị định 30/2020/NĐ-CP.")
    add_p(doc, "• Nắm vững kỹ thuật dùng Bảng 2 cột ẩn đường viền (Table) để căn Quốc hiệu - Tiêu ngữ song song tuyệt đối với Tên cơ quan - Số ký hiệu mà không bao giờ bị xô lệch.")
    add_p(doc, "• Kẻ đường chỉ phân cách có độ dài chuẩn xác dưới Tiêu ngữ và Tên cơ quan.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Tiếp tục sử dụng trang văn bản đã thiết lập khổ A4 và lề chuẩn ở Bài 5.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "QUY ĐỊNH THỂ THỨC PHẦN ĐẦU VĂN BẢN (NGHỊ ĐỊNH 30/2020/NĐ-CP)", [
        "1. Tên cơ quan ban hành (Góc trên bên trái):",
        "  • Cơ quan chủ quản (nếu có): Chữ in hoa, cỡ 12 - 13, kiểu đứng (ví dụ: BỘ NỘI VỤ hoặc UBND THÀNH PHỐ HÀ NỘI).",
        "  • Cơ quan ban hành: Chữ in hoa, cỡ 12 - 13, kiểu đứng, IN ĐẬM (ví dụ: VĂN PHÒNG BỘ hoặc SỞ KẾ HOẠCH VÀ ĐẦU TƯ).",
        "  • Có đường kẻ ngang bên dưới, nét liền, độ dài bằng 1/3 đến 1/2 độ dài dòng chữ.",
        "2. Số và ký hiệu văn bản (Nằm ngay dưới tên cơ quan): Cỡ 13, kiểu đứng (ví dụ: Số: 125/VP-TH hoặc Số: 45/QĐ-UBND).",
        "3. Quốc hiệu và Tiêu ngữ (Góc trên bên phải):",
        "  • Dòng trên: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM' - Chữ in hoa, cỡ 12 - 13, kiểu đứng, IN ĐẬM.",
        "  • Dòng dưới: 'Độc lập - Tự do - Hạnh phúc' - Chữ in thường (chữ cái đầu viết hoa), cỡ 13 - 14, kiểu đứng, IN ĐẬM, giữa các từ có gạch nối và khoảng cách.",
        "  • Đường kẻ ngang bên dưới Tiêu ngữ: Nét liền, có độ dài bằng độ dài dòng chữ 'Độc lập - Tự do - Hạnh phúc'.",
        "4. Địa danh và ngày tháng ban hành (Dưới Quốc hiệu - Tiêu ngữ): Cỡ 13 - 14, kiểu chữ NGHIÊNG (ví dụ: Hà Nội, ngày 05 tháng 9 năm 2026)."
    ], callout_type="legal")

    add_step(doc, "6.1", "Tạo khung bố trí bằng Bảng 2 cột 1 hàng (Table)", [
        "Vào thẻ Insert (Chèn) trên thanh Ribbon > Nhấp chuột vào nút 'Table' (Bảng).",
        "Rê chuột chọn lưới 2 cột và 1 hàng (bảng 2 x 1) > Nhấp chuột trái để chèn bảng vào đầu trang.",
        "Nhấp chuột vào ô bên trái: Dùng để nhập Cơ quan ban hành và Số ký hiệu.",
        "Nhấp chuột vào ô bên phải: Dùng để nhập Quốc hiệu, Tiêu ngữ và Địa danh ngày tháng."
    ])

    add_step(doc, "6.2", "Nhập nội dung và Căn chỉnh phông chữ từng ô", [
        "Tại ô bên trái: Nhập tên cơ quan chủ quản, tên cơ quan ban hành, và số ký hiệu. Chọn toàn bộ chữ trong ô và bấm tổ hợp phím 'Ctrl + E' để căn giữa ô.",
        "Tại ô bên phải: Nhập Quốc hiệu, Tiêu ngữ, và Địa danh ngày tháng. Chọn toàn bộ chữ trong ô và bấm 'Ctrl + E' để căn giữa ô.",
        "Bôi đen từng dòng để chỉnh đúng cỡ chữ: Cỡ chữ Quốc hiệu (12-13pt đậm), Tiêu ngữ (13-14pt đậm), Địa danh ngày tháng (13-14pt nghiêng)."
    ])

    add_step(doc, "6.3", "Vẽ đường kẻ chỉ phân cách chuẩn xác", [
        "Cách kẻ chuẩn nhất: Vào thẻ Insert > Nhấp nút 'Shapes' (Hình khối) > Chọn hình đường thẳng nét liền (Line).",
        "Nhấn giữ phím Shift trên bàn phím, đồng thời nhấn giữ chuột trái kéo từ trái sang phải ngay dưới dòng Tiêu ngữ để vẽ một đường thẳng nằm ngang hoàn hảo (giữ Shift giúp đường thẳng không bao giờ bị méo lệch).",
        "Vào thẻ 'Shape Format' xuất hiện trên thanh công cụ: Mục 'Shape Outline' chọn màu Đen (Black), mục 'Weight' chọn nét dày 1 pt hoặc 1.5 pt.",
        "Làm tương tự một đường kẻ ngắn hơn (bằng 1/3 dòng chữ) đặt dưới tên cơ quan ban hành bên trái."
    ])

    add_step(doc, "6.4", "Ẩn đường viền bảng (Tạo sự liền mạch cho trang giấy)", [
        "Nhấp chuột vào biểu tượng dấu cộng 4 hướng ở góc trên bên trái của bảng để chọn toàn bộ bảng.",
        "Trên thanh công cụ, nhấp chuột vào thẻ 'Table Design' (Thiết kế bảng).",
        "Nhấp vào nút 'Borders' (Viền) > Chọn dòng 'No Border' (Không viền). Toàn bộ đường viền đen của bảng sẽ biến mất, chỉ để lại nội dung được căn thẳng tắp hai bên một cách hoàn hảo."
    ])

    add_screenshot_placeholder(doc, "Bảng 2 cột ẩn viền: Bên trái là Tên cơ quan và Số hiệu, bên phải là Quốc hiệu, Tiêu ngữ kèm đường kẻ chuẩn và Địa danh ngày tháng.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Phần đầu văn bản hành chính được bố trí cân xứng tuyệt đối, cơ quan ban hành và Quốc hiệu ngang hàng nhau, đường kẻ thẳng hàng tăm tắp, đúng chuẩn 100% Nghị định 30/2020/NĐ-CP.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI DÙNG PHÍM TAB HOẶC PHÍM CÁCH ĐỂ CĂN THỂ THỨC", [
        "Lỗi kinh điển của người mới: Gõ tên cơ quan bên trái, sau đó bấm phím Cách (Space) hoặc phím Tab hàng chục lần để đẩy chữ Quốc hiệu sang bên phải. Hậu quả: Khi đổi phông chữ hoặc in ấn, chữ sẽ bị xô lệch lung tung, rớt dòng, méo mó văn bản.",
        "Cách sửa triệt để: Xóa bỏ hoàn toàn cách làm cũ. BẮT BUỘC dùng kỹ thuật chèn Bảng 2 cột ẩn viền (Insert Table 2x1) như đã hướng dẫn ở Bước 6.1."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Tạo một tài liệu Word mới, sử dụng bảng 2x1 ẩn viền để thiết lập phần đầu văn bản cho cơ quan giả lập sau:")
    add_p(doc, "• Cơ quan chủ quản: ỦY BAN NHÂN DÂN THÀNH PHỐ HÀ NỘI")
    add_p(doc, "• Cơ quan ban hành: SỞ THÔNG TIN VÀ TRUYỀN THÔNG (Chữ in hoa, đậm)")
    add_p(doc, "• Số ký hiệu: Số: 128/STTTT-CNTT")
    add_p(doc, "• Quốc hiệu - Tiêu ngữ đầy đủ theo quy định")
    add_p(doc, "• Địa danh: Hà Nội, ngày 15 tháng 10 năm 2026")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Sử dụng bảng 2x1 ẩn viền hoàn toàn", "Không còn nhìn thấy nét viền đen của bảng trên trang giấy"),
        ("Tên cơ quan và Quốc hiệu ngang hàng nhau", "Hai khối nội dung cân đối, thẳng hàng tuyệt đối"),
        ("Đường kẻ dưới Tiêu ngữ có độ dài bằng độ dài dòng chữ", "Đường kẻ thẳng ngang, màu đen rõ nét, không bị méo lệch"),
        ("Địa danh ngày tháng in nghiêng cỡ 13-14pt", "Dòng ngày tháng hiển thị đúng chữ nghiêng trang trọng")
    ])

    # BÀI 7
    add_heading_2(doc, "Bài 7: Trình bày Nội dung văn bản, Phân cấp mục và Thao tác đoạn văn chuyên nghiệp")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Trình bày Tên loại văn bản và Trích yếu nội dung đúng chuẩn.")
    add_p(doc, "• Cài đặt tính năng tự động thụt đầu dòng đoạn văn bản (First Line Indent) từ 1cm - 1.27cm.")
    add_p(doc, "• Nắm vững kỹ thuật căn đều hai bên lề (Justify) bằng phím tắt Ctrl + J.")
    add_p(doc, "• Biết cách phân cấp các điều khoản, đề mục (I, 1, a, gạch đầu dòng) đúng hệ thống văn bản nhà nước.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Mở tệp văn bản đã có phần đầu chuẩn ở Bài 6 để tiếp tục soạn phần thân nội dung.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "QUY ĐỊNH VỀ TÊN LOẠI VĂN BẢN VÀ TRÍCH YẾU NỘI DUNG", [
        "Đối với văn bản có tên loại (Nghị quyết, Quyết định, Kế hoạch, Thông báo, Tờ trình, Báo cáo...):",
        "  • Tên loại văn bản: Đặt ở giữa trang, chữ in hoa, cỡ 14 - 15, IN ĐẬM (ví dụ: THÔNG BÁO hoặc TỜ TRÌNH).",
        "  • Trích yếu nội dung: Đặt ngay dưới tên loại, chữ in thường, cỡ 13 - 14, kiểu đứng, IN ĐẬM (ví dụ: Về việc phân công nhiệm vụ công tác quý IV năm 2026).",
        "Đối với Công văn hành chính (Không có tên loại riêng):",
        "  • Trích yếu nội dung được đặt ngay dưới Số ký hiệu (ở góc trái): Chữ in thường, cỡ 12 - 13, kiểu đứng (ví dụ: V/v phối hợp tổ chức lớp đào tạo kỹ năng số). Dưới trích yếu có đường kẻ ngắn bằng 1/3 đến 1/2 độ dài dòng chữ."
    ], callout_type="legal")

    add_step(doc, "7.1", "Cài đặt Tự động thụt đầu dòng (First Line Indent)", [
        "Không dùng phím Space (cách) để thụt đầu dòng. Hãy thiết lập tự động bằng Paragraph:",
        "Bôi đen toàn bộ các đoạn văn bản trong phần nội dung (hoặc bấm Ctrl + A).",
        "Vào thẻ Home > Bấm vào mũi tên mở hộp thoại 'Paragraph'.",
        "Tại mục 'Indentation' (Thụt lề) > nhìn sang ô 'Special' (Đặc biệt): Nhấp chuột chọn dòng 'First line' (Dòng đầu tiên).",
        "Tại ô 'By': Nhập vào '1 cm' (hoặc 1.27 cm / 0.5 inch).",
        "Bấm nút 'OK'. Toàn bộ các dòng đầu tiên của mỗi đoạn văn sẽ tự động lùi vào trong một khoảng cách chuẩn xác."
    ])

    add_step(doc, "7.2", "Căn đều hai bên lề văn bản (Justify - Phím tắt Ctrl + J)", [
        "Văn bản hành chính nhà nước bắt buộc phải căn đều cả mép trái lẫn mép phải (không để mép phải bị so le lởm chởm).",
        "Bôi đen các đoạn văn bản cần căn lề.",
        "Nhấn tổ hợp phím tắt 'Ctrl + J' (hoặc bấm vào biểu tượng nút Justify trên thanh công cụ thẻ Home). Mọi dòng chữ sẽ được dóng thẳng tắp ở cả 2 mép lề giấy."
    ])

    add_step(doc, "7.3", "Quy tắc phân cấp các đề mục trong văn bản hành chính", [
        "Khi văn bản có nhiều mục, Nghị định 30/2020/NĐ-CP quy định thứ tự phân cấp chuẩn mực như sau:",
        "  • Cấp 1 - Phần, Chương: Dùng chữ số La Mã (I, II, III...). Chữ in hoa, cỡ 13 - 14, kiểu đứng, IN ĐẬM.",
        "  • Cấp 2 - Mục, Điều: Dùng chữ số Ả Rập kèm dấu chấm (1., 2., 3...). Chữ in thường, cỡ 13 - 14, kiểu đứng, IN ĐẬM.",
        "  • Cấp 3 - Khoản: Dùng chữ cái thường kèm dấu ngoặc đơn [a), b), c)...]. Chữ in thường, cỡ 13 - 14, kiểu đứng.",
        "  • Cấp 4 - Điểm, Ý nhỏ: Dùng dấu gạch đầu dòng (-) hoặc dấu cộng (+). Chữ in thường, cỡ 13 - 14, kiểu đứng."
    ])

    add_step(doc, "7.4", "Kỹ thuật ngắt trang chủ động (Page Break - Phím tắt Ctrl + Enter)", [
        "Khi muốn chuyển nội dung sang trang mới (ví dụ khi hết một chương hoặc muốn đẩy bảng biểu sang trang sau): Tuyệt đối không bấm phím Enter liên tục nhiều lần.",
        "Cách làm đúng: Đặt con trỏ soạn thảo tại vị trí muốn ngắt > Nhấn tổ hợp phím 'Ctrl + Enter'. Văn bản phía sau sẽ lập tức nhảy sang trang mới sạch sẽ, không bao giờ bị xô lệch khi trang trước có chỉnh sửa thêm bớt chữ."
    ])

    add_screenshot_placeholder(doc, "Thao tác thụt đầu dòng tự động trong hộp thoại Paragraph (Special: First Line By 1cm) và nút căn đều hai bên Justify (Ctrl + J).")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Các đoạn văn bản ngay ngắn, đầu dòng lùi đều 1cm, hai mép lề thẳng tắp; hệ thống đề mục I, 1, a) rõ ràng, khoa học.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "CÁC LỖI THƯỜNG GẶP TRONG PHẦN NỘI DUNG", [
        "Lỗi dùng phím Space gõ 4-5 lần để thụt đầu dòng: Hậu quả là các dòng trên và dòng dưới lùi vào không đều nhau, nhìn rất thiếu chuyên nghiệp. Cách sửa: Xóa các dấu cách thừa và dùng tính năng First line indent.",
        "Lỗi chữ bị giãn cách quá xa trên dòng (răng thưa) khi bấm Ctrl + J: Xảy ra khi bạn bấm phím 'Shift + Enter' (ngắt dòng mềm) thay vì phím 'Enter' (kết thúc đoạn văn). Cách sửa: Xóa ký tự ngắt dòng đó và nhấn phím Enter thông thường."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Soạn thảo phần thân nội dung của một văn bản Kế hoạch với các yêu cầu kỹ thuật:")
    add_p(doc, "1. Tên loại: 'KẾ HOẠCH' (Cỡ 14 đậm, giữa trang). Trích yếu: 'Về việc triển khai chuyển đổi số cơ quan năm 2026' (Cỡ 13 đậm, giữa trang).")
    add_p(doc, "2. Tạo mục 'I. MỤC ĐÍCH, YÊU CẦU' (chữ hoa đậm). Đoạn văn bên dưới thụt lề đầu dòng 1cm, căn đều 2 bên (Ctrl + J).")
    add_p(doc, "3. Tạo các mục con: 1. Mục đích; 2. Yêu cầu; bên trong có các điểm a), b) gạch đầu dòng.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Nội dung được căn đều hai bên mép lề (Justify)", "Các dòng văn bản thẳng đều ở cả lề trái và lề phải"),
        ("Đoạn văn tự động thụt lùi đầu dòng 1cm", "Dòng đầu tiên của mỗi đoạn lùi vào chính xác 1cm"),
        ("Phân cấp đề mục đúng thứ tự (I > 1 > a > -)", "Cấu trúc văn bản mạch lạc, dễ theo dõi"),
        ("Không có khoảng trắng thừa và không ngắt trang bằng phím Enter", "Dùng phím Ctrl + Enter khi cần sang trang mới")
    ])

    # BÀI 8
    add_heading_2(doc, "Bài 8: Bố trí phần Ký duyệt, Nơi nhận và Bảng biểu số liệu trong văn bản")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Hiểu rõ quy định thể thức phần Nơi nhận và phần Quyền hạn, Chức vụ chữ ký theo Nghị định 30/2020/NĐ-CP.")
    add_p(doc, "• Áp dụng thành thạo kỹ thuật Bảng 2 cột ẩn viền để bố trí phần đuôi văn bản hoàn mỹ.")
    add_p(doc, "• Chèn, định dạng và căn chỉnh bảng biểu số liệu tổng hợp ngay trong văn bản Word.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Mở tệp văn bản đang soạn thảo để hoàn thiện phần cuối trang.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "QUY ĐỊNH THỂ THỨC PHẦN CUỐI VĂN BẢN (NGHỊ ĐỊNH 30/2020/NĐ-CP)", [
        "1. Nơi nhận (Góc dưới bên trái):",
        "  • Từ 'Nơi nhận:': Cỡ chữ 12, kiểu chữ NGHIÊNG, IN ĐẬM, có dấu hai chấm.",
        "  • Danh sách nơi nhận bên dưới: Cỡ chữ 11, kiểu chữ đứng, gạch đầu dòng (-) thụt sát lề trái.",
        "  • Dòng cuối cùng của nơi nhận luôn có: '- Lưu: VT, [Tên đơn vị soạn thảo] (số lượng bản lưu).'",
        "2. Chức vụ, Họ tên và Chữ ký (Góc dưới bên phải):",
        "  • Quyền hạn (nếu ký thay): TM. (Thay mặt), KT. (Ký thay), TL. (Thừa lệnh), TUQ. (Thừa ủy quyền) - Chữ in hoa, cỡ 13 - 14, IN ĐẬM.",
        "  • Chức vụ của người ký (GIÁM ĐỐC, CHỦ TỊCH...): Chữ in hoa, cỡ 13 - 14, IN ĐẬM.",
        "  • Khoảng trống để ký tên và đóng dấu: Chừa khoảng 4 - 5 dòng trống (khoảng 3 - 4 cm).",
        "  • Họ và tên người ký: Chữ in thường (viết hoa chữ cái đầu), cỡ 13 - 14, kiểu đứng, IN ĐẬM, căn giữa so với chức vụ."
    ], callout_type="legal")

    add_step(doc, "8.1", "Bố trí phần Ký và Nơi nhận bằng Bảng 2 cột ẩn viền", [
        "Đặt con trỏ ở cuối văn bản, nhấn Enter tạo dòng mới.",
        "Vào thẻ Insert > Table > Chèn bảng 2 cột 1 hàng (Table 2x1).",
        "Căn chỉnh độ rộng cột: Kéo đường ranh giới giữa 2 cột sang bên trái một chút, sao cho cột bên trái (Nơi nhận) chiếm khoảng 40% độ rộng, cột bên phải (Chữ ký) chiếm khoảng 60% độ rộng trang.",
        "Ô bên trái: Gõ 'Nơi nhận:', nhấn Enter, gõ danh sách các cơ quan nhận, cỡ chữ 11.",
        "Ô bên phải: Bấm Ctrl + E để căn giữa ô. Gõ Chức vụ (in hoa, đậm, cỡ 13-14) > Nhấn Enter 4 lần để tạo khoảng trống ký tên > Gõ Họ và tên người ký (chữ thường, đậm, cỡ 13-14).",
        "Ẩn viền bảng: Chọn toàn bộ bảng > Vào thẻ Table Design > Borders > Chọn 'No Border'."
    ])

    add_step(doc, "8.2", "Chèn và Định dạng bảng biểu số liệu trong văn bản", [
        "Vào thẻ Insert > Table > Chọn số cột và số dòng cần thiết cho bảng số liệu.",
        "Nhập nội dung vào các ô: Sử dụng phím 'Tab' để di chuyển nhanh sang ô bên phải tiếp theo; khi ở ô cuối cùng của bảng, nhấn phím 'Tab' sẽ tự động tạo thêm 1 dòng mới.",
        "Gộp nhiều ô thành một ô (Merge Cells): Bôi đen các ô cần gộp > Nhấp chuột phải chọn 'Merge Cells'.",
        "Căn chỉnh dữ liệu trong ô:",
        "  • Dòng tiêu đề cột: Bôi đen toàn bộ dòng đầu > Vào thẻ Home chọn In đậm (Ctrl + B), Căn giữa (Ctrl + E). Vào thẻ Layout của Table > Nhấp vào nút 'Align Center' (căn chính giữa ô cả chiều dọc lẫn chiều ngang).",
        "  • Cột số thứ tự, ngày tháng, mã hiệu: Căn giữa.",
        "  • Cột tên gọi, nội dung diễn giải: Căn lề trái.",
        "  • Cột số tiền, số lượng: Căn lề phải."
    ])

    add_screenshot_placeholder(doc, "Bố cục hoàn chỉnh phần chân văn bản: Nơi nhận bên trái, Chức vụ và Họ tên người ký bên phải, ẩn toàn bộ viền bảng.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Phần chữ ký và nơi nhận cân đối tuyệt đối; bảng biểu trong văn bản có tiêu đề in đậm căn giữa, số liệu dóng thẳng hàng, không bị tràn ra ngoài lề trang giấy.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI BẢNG BỊ TRÀN RA NGOÀI MÉP GIẤY", [
        "Khi chèn bảng có nhiều cột, bảng có thể bị phình to và tràn qua lề phải trang giấy khiến khi in ra bị mất dữ liệu.",
        "Cách sửa nhanh trong 1 giây: Nhấp chuột vào bảng > Thẻ Layout của Table xuất hiện trên thanh Ribbon > Nhấp vào nút 'AutoFit' > Chọn dòng 'AutoFit to Window' (Tự động co vừa khít độ rộng trang giấy)."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Hoàn thiện phần chân của một văn bản hành chính giả lập:")
    add_p(doc, "1. Nơi nhận: Như điều 3; Giám đốc Sở (để b/c); Các phòng chuyên môn; Lưu: VT, CNTT (03b).")
    add_p(doc, "2. Chữ ký: KT. GIÁM ĐỐC / PHÓ GIÁM ĐỐC / (để trống 4 dòng) / Nguyễn Văn An.")
    add_p(doc, "3. Chèn một bảng thống kê 4 cột 3 dòng gồm: STT, Tên thiết bị, Số lượng, Đơn vị tính; nhập dữ liệu giả lập cho 2 mặt hàng máy tính và máy in.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Nơi nhận và Chữ ký được bố trí bằng Bảng ẩn viền", "Không bị xô lệch khi mở trên các máy tính khác"),
        ("Cỡ chữ nơi nhận đúng quy định (11-12pt)", "Chữ Nơi nhận in nghiêng đậm, danh sách chữ đứng"),
        ("Chức vụ và Họ tên người ký thẳng hàng chính giữa", "Có đủ khoảng cách 3-4cm để ký tên và đóng dấu đỏ"),
        ("Bảng biểu nằm gọn gàng bên trong lề giấy A4", "Không có cột nào bị tràn qua mép lề phải")
    ])

    # BÀI 9
    add_heading_2(doc, "Bài 9: Đánh số trang theo Nghị định 30, Kiểm tra hoàn thiện và In ấn / Xuất PDF")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Đánh số trang văn bản đúng tuyệt đối theo quy định của Nghị định 30/2020/NĐ-CP (Đánh ở giữa mép trên trang giấy, không hiển thị số trang tại trang 1).")
    add_p(doc, "• Sử dụng công cụ Tìm kiếm & Thay thế (Ctrl + H) để tự động xóa sạch các khoảng trắng thừa do gõ sai.")
    add_p(doc, "• Thiết lập chế độ In hai mặt (Print on Both Sides) và xử lý lỗi trang trắng thừa.")
    add_p(doc, "• Xuất văn bản sang tệp PDF sắc nét, sẵn sàng gửi phát hành hoặc ký số.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Tệp văn bản hành chính dài từ 2 trang trở lên.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "QUY ĐỊNH PHÁP LUẬT VỀ ĐÁNH SỐ TRANG VĂN BẢN (NGHỊ ĐỊNH 30/2020/NĐ-CP)", [
        "Vị trí: Đặt canh giữa theo chiều ngang trong phần lề trên của văn bản.",
        "Số trang: Được đánh từ số 1, bằng chữ số Ả Rập, cỡ chữ 13 - 14, kiểu chữ đứng.",
        "ĐIỀU KHOẢN ĐẶC BIỆT QUAN TRỌNG: Không hiển thị số trang thứ nhất (Trang 1 không có số; từ trang 2 trở đi mới hiển thị số 2, số 3...)."
    ], callout_type="legal")

    add_step(doc, "9.1", "Cài đặt Đánh số trang chuẩn Nghị định 30 trong Word", [
        "Vào thẻ Insert trên thanh công cụ Ribbon.",
        "Nhấp vào nút 'Page Number' (Số trang) > Chọn 'Top of Page' (Đỉnh trang) > Chọn kiểu 'Plain Number 2' (Số đơn giản nằm ở chính giữa mép trên).",
        "Lúc này con trỏ đang ở khu vực Header (Đầu trang). Nhìn lên thanh công cụ thẻ 'Header & Footer', tìm nhóm 'Options' và NHẤP CHUỘT TÍCH CHỌN VÀO Ô: 'Different First Page' (Khác biệt trang đầu tiên).",
        "Ngay lập tức, số trang tại trang thứ nhất sẽ tự động biến mất, và từ trang thứ hai sẽ hiển thị số 2, số 3 đúng chuẩn 100% theo quy định.",
        "Nhấp đúp chuột vào giữa trang văn bản (hoặc bấm nút Close Header and Footer) để quay lại soạn thảo bình thường."
    ])

    add_step(doc, "9.2", "Xóa sạch toàn bộ khoảng trắng thừa bằng công cụ Replace (Ctrl + H)", [
        "Trong quá trình gõ, người mới thường vô tình bấm 2-3 dấu cách liên tiếp giữa các từ. Hãy dùng mẹo tự động hóa sau:",
        "Nhấn tổ hợp phím 'Ctrl + H' để mở cửa sổ Find and Replace (Tìm kiếm và Thay thế).",
        "Tại ô 'Find what' (Tìm gì): Nhấn phím cách 2 lần (gõ 2 khoảng trắng).",
        "Tại ô 'Replace with' (Thay thế bằng): Nhấn phím cách 1 lần (gõ 1 khoảng trắng duy nhất).",
        "Nhấp chuột vào nút 'Replace All' (Thay thế tất cả). Word sẽ tự động quét và gom toàn bộ khoảng trắng thừa lại thành đúng 1 dấu cách duy nhất trên toàn văn bản. Lặp lại bấm nút này đến khi Word báo '0 replacements'."
    ])

    add_step(doc, "9.3", "Cài đặt In ấn văn bản 2 mặt chuẩn xác", [
        "Nhấn tổ hợp phím 'Ctrl + P' (hoặc vào File > Print) để mở chế độ in.",
        "Xem trước bản in (Print Preview) ở bên phải màn hình: Kiểm tra từng trang xem có dòng chữ nào bị rớt lẻ loi sang trang cuối không.",
        "Chọn máy in tại mục 'Printer'.",
        "Cài đặt in hai mặt: Tại mục 'Print One Sided' (In một mặt), nhấp chuột chọn 'Print on Both Sides' (In trên cả hai mặt) > Chọn kiểu lật trang 'Flip pages on long edge' (Lật trang theo cạnh dài - cách mở sổ/văn bản thông thường).",
        "Bấm nút lớn 'Print' để bắt đầu in."
    ])

    add_callout(doc, "CÁCH XÓA TRANG TRẮNG THỪA Ở CUỐI VĂN BẢN", [
        "Hiện tượng: Văn bản chỉ có 2 trang nhưng máy in lại đẩy ra trang thứ 3 trắng tinh, hoặc trên màn hình hiển thị thêm 1 trang trắng vô ích ở cuối.",
        "Nguyên nhân: Do có các dấu xuống dòng (Enter) thừa ở cuối trang, hoặc sau bảng biểu có 1 đoạn văn ẩn.",
        "Cách khắc phục:",
        "  1. Bấm vào thẻ Home > Nhấp vào biểu tượng nút Show/Hide ¶ (hình chữ P ngược) để nhìn thấy toàn bộ ký tự ẩn.",
        "  2. Cuộn chuột xuống trang trắng cuối cùng, bạn sẽ thấy các biểu tượng ¶ màu xám.",
        "  3. Đặt con trỏ chuột ở cuối trang trắng đó và nhấn phím 'Backspace' hoặc 'Delete' liên tục cho đến khi trang trắng biến mất hoàn toàn."
    ], callout_type="warning")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Tài liệu được đánh số trang ở giữa phía trên (trang 1 ẩn số, trang 2 hiện số 2); văn bản sạch sẽ không còn khoảng trắng đôi; in ấn 2 mặt chuẩn mực hoặc xuất thành tệp PDF hoàn hảo.")

    add_heading_3(doc, "5. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở một văn bản dài từ 3 trang. Thực hiện đánh số trang theo chuẩn Nghị định 30 (Top of page > Plain number 2 > Tích chọn Different First Page). Sử dụng Ctrl + H để làm sạch khoảng trắng thừa. Cuối cùng, vào File > Export > Xuất văn bản ra tệp PDF.")

    add_heading_3(doc, "6. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Số trang đặt ở mép trên chính giữa (Header)", "Số trang nằm đúng vị trí quy định của NĐ 30"),
        ("Trang đầu tiên KHÔNG hiển thị số trang", "Trang 1 hoàn toàn không có số; trang 2 hiện số 2"),
        ("Không có khoảng trắng kép giữa các từ", "Kiểm tra bằng Ctrl + H báo 0 replacements"),
        ("Không còn trang trắng thừa ở cuối tài liệu", "Số trang thực tế trên màn hình khớp với số trang in ra")
    ])

    # BÀI 10
    add_heading_2(doc, "Bài 10: Dự án thực hành tổng hợp Word - Soạn thảo 03 Văn bản hành chính thực tế")
    
    add_p(doc, "Bài học này là dự án thực hành lớn giúp bạn xâu chuỗi toàn bộ kỹ năng từ Bài 5 đến Bài 9. Bạn sẽ tự tay tạo ra 03 mẫu văn bản hành chính phổ biến nhất trong cơ quan nhà nước theo dữ liệu giả lập chuẩn tắc dưới đây.")

    add_heading_3(doc, "Dự án 1: Soạn thảo Công văn hành chính gửi Sở Thông tin & Truyền thông")
    add_p(doc, "Yêu cầu: Soạn thảo một Công văn của UBND Quận X gửi Sở Thông tin và Truyền thông đề nghị hỗ trợ chuyên môn về triển khai chữ ký số công vụ.")
    
    add_callout(doc, "DỮ LIỆU GIẢ LẬP ĐỂ THỰC HÀNH CÔNG VĂN", [
        "• Cơ quan chủ quản: ỦY BAN NHÂN DÂN QUẬN HOÀNG MAI",
        "• Đơn vị ban hành: VĂN PHÒNG HĐND VÀ UBND (In hoa, đậm)",
        "• Số ký hiệu: Số: 452/VP-CNTT",
        "• Địa danh ngày tháng: Hoàng Mai, ngày 05 tháng 9 năm 2026",
        "• Trích yếu nội dung (nằm dưới số hiệu): V/v hướng dẫn kỹ thuật triển khai chữ ký số chuyên dùng công vụ.",
        "• Nơi nhận kính gửi: Kính gửi: Sở Thông tin và Truyền thông thành phố Hà Nội.",
        "• Nội dung: Thực hiện Kế hoạch số 88/KH-UBND về ứng dụng công nghệ thông tin trong hoạt động cơ quan nhà nước; để phục vụ tốt công tác gửi nhận văn bản điện tử và ký số văn bản hành chính; Văn phòng HĐND và UBND quận kính đề nghị Sở Thông tin và Truyền thông hỗ trợ cử cán bộ kỹ thuật hướng dẫn tập huấn nghiệp vụ cho cán bộ, công chức của quận... Thời gian dự kiến: Thứ Sáu, ngày 18 tháng 9 năm 2026. Địa điểm: Hội trường tầng 3, trụ sở UBND quận.",
        "• Nơi nhận: Như kính gửi; Chủ tịch UBND quận (để b/c); Lưu: VT, CNTT (02b).",
        "• Người ký: CHÁNH VĂN PHÒNG / (khoảng trống ký) / Trần Quốc Toản"
    ], callout_type="note")

    add_heading_3(doc, "Dự án 2: Soạn thảo Thông báo kết luận cuộc họp giao ban")
    add_p(doc, "Yêu cầu: Soạn thảo một Thông báo truyền đạt kết luận của Lãnh đạo cơ quan tại cuộc họp định kỳ.")
    
    add_callout(doc, "DỮ LIỆU GIẢ LẬP ĐỂ THỰC HÀNH THÔNG BÁO", [
        "• Cơ quan ban hành: SỞ LAO ĐỘNG - THƯƠNG BINH VÀ XÃ HỘI TỈNH BẮC GIANG",
        "• Số ký hiệu: Số: 89/TB-SLĐTBXH, ngày 10 tháng 9 năm 2026",
        "• Tên loại và Trích yếu: THÔNG BÁO (cỡ 14 đậm giữa) / Kết luận của Giám đốc Sở tại cuộc họp giao ban công tác tháng 9 năm 2026 (cỡ 13 đậm giữa).",
        "• Nội dung: Ngày 09/9/2026, tại Phòng họp số 1, Giám đốc Sở đã chủ trì cuộc họp giao ban. Sau khi nghe báo cáo và ý kiến thảo luận, Giám đốc Sở kết luận như sau:",
        "  1. Văn phòng Sở: Tiếp tục đôn đốc các phòng hoàn thành báo cáo số liệu trước ngày 15 hàng tháng; rà soát thiết bị phòng họp trực tuyến.",
        "  2. Phòng Kế hoạch - Tài chính: Khẩn trương hoàn thiện hồ sơ dự toán trang thiết bị công nghệ thông tin phục vụ chuyển đổi số.",
        "  3. Các phòng chuyên môn: Nghiêm túc thực hiện việc tiếp nhận và xử lý hồ sơ hành chính trên Hệ thống Một cửa điện tử đúng thời hạn quy định.",
        "• Nơi nhận: Ban Giám đốc Sở; Các phòng thuộc Sở; Lưu: VT, VP.",
        "• Người ký: TL. GIÁM ĐỐC / CHÁNH VĂN PHÒNG / Nguyễn Thị Mai Lan"
    ], callout_type="note")

    add_heading_3(doc, "Dự án 3: Soạn thảo Tờ trình xin phê duyệt kinh phí")
    add_p(doc, "Yêu cầu: Soạn thảo một Tờ trình có lồng ghép bảng biểu số liệu dự toán kinh phí nâng cấp máy tính.")
    
    add_callout(doc, "DỮ LIỆU GIẢ LẬP ĐỂ THỰC HÀNH TỜ TRÌNH", [
        "• Tên loại & Trích yếu: TỜ TRÌNH / Về việc phê duyệt kinh phí mua sắm trang thiết bị tin học phục vụ công tác văn phòng năm 2026.",
        "• Kính gửi: Giám đốc Sở Kế hoạch và Đầu tư.",
        "• Nội dung: Căn cứ tình hình thực tế hiện nay, hệ thống máy vi tính tại bộ phận Tiếp nhận và Trả kết quả đã xuống cấp, tốc độ xử lý chậm... Văn phòng Sở kính trình Giám đốc Sở xem xét, phê duyệt danh mục kinh phí mua sắm như sau:",
        "  (Chèn bảng biểu dự toán 4 cột: STT | Thiết bị | Số lượng | Thành tiền dự kiến)",
        "  1 | Bộ máy vi tính để bàn cấu hình cao | 03 bộ | 45.000.000 đồng",
        "  2 | Máy quét văn bản tài liệu 2 mặt tự động | 01 chiếc | 12.000.000 đồng",
        "  Tổng cộng kinh phí dự kiến: 57.000.000 đồng (Bằng chữ: Năm mươi bảy triệu đồng chẵn).",
        "• Kính trình Giám đốc Sở xem xét, quyết định./.",
        "• Chữ ký: CHÁNH VĂN PHÒNG / Vũ Đình Trọng"
    ], callout_type="note")

    add_heading_3(doc, "Tiêu chí đánh giá bài thực hành tổng hợp Word")
    add_checklist_table(doc, [
        ("Cả 3 văn bản đều thiết lập lề trang chuẩn Nghị định 30 (Trái 3cm, Trên/Dưới 2cm, Phải 1.5-2cm)", "Khổ giấy A4 đứng, căn lề chuẩn mực"),
        ("Phần đầu và phần chân dùng đúng bảng ẩn viền Table", "Không bị xô lệch, đường kẻ thẳng đẹp"),
        ("Nội dung được căn đều 2 bên (Ctrl+J) và thụt đầu dòng tự động (First Line 1cm)", "Bố cục đoạn văn ngay ngắn, chuyên nghiệp"),
        ("Bảng biểu trong Tờ trình được căn chỉnh đẹp, số tiền căn lề phải", "Bảng biểu nằm vừa khít trong trang giấy"),
        ("Đánh số trang từ trang 2 ở giữa mép trên văn bản", "Đúng quy định pháp lý hiện hành")
    ])
