# -*- coding: utf-8 -*-
"""
Module biên soạn PHẦN 3: QUẢN LÝ & XỬ LÝ BẢNG TÍNH VĂN PHÒNG VỚI MICROSOFT EXCEL
Dành cho Tập I.
"""

from docx_helper import (
    add_heading_1, add_heading_2, add_heading_3, add_p,
    add_step, add_callout, add_screenshot_placeholder,
    add_styled_table, add_checklist_table
)

def build_part_3(doc):
    # PHẦN 3
    add_heading_1(doc, "PHẦN 3: QUẢN LÝ & XỬ LÝ BẢNG TÍNH VĂN PHÒNG VỚI MICROSOFT EXCEL")
    add_p(doc, "Bảng tính Microsoft Excel là công cụ không thể thiếu để theo dõi công việc, quản lý hồ sơ, kiểm kê tài sản và tính toán chi phí văn phòng. Phần 3 được thiết kế từng bước đơn giản, trực quan, giúp người mới bắt đầu xóa bỏ nỗi sợ công thức tính toán, nắm vững bản chất của ô và dòng/cột, sử dụng thành thạo các hàm thiết yếu nhất (SUM, AVERAGE, MIN, MAX, COUNT, IF), lọc dữ liệu thông minh và đặc biệt là làm chủ kỹ thuật căn chỉnh in ấn bảng tính vừa khít trên một trang giấy A4.")

    # BÀI 11
    add_heading_2(doc, "Bài 11: Làm quen giao diện Excel và Nguyên tắc nhập liệu an toàn")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Hiểu rõ khái niệm Cột (Column), Dòng (Row), Địa chỉ Ô (Cell) và Vùng dữ liệu (Range).")
    add_p(doc, "• Phân biệt được 3 kiểu dữ liệu cốt lõi: Dạng chữ (Văn bản), Dạng số và Dạng Ngày tháng.")
    add_p(doc, "• Biết cách sửa nội dung ô bằng phím tắt F2 mà không bị xóa mất dữ liệu cũ.")
    add_p(doc, "• Thực hiện thành thạo thao tác chèn thêm hoặc xóa bớt dòng, cột trong bảng tính.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Khởi động phần mềm Microsoft Excel > Nhấp chuột chọn 'Blank workbook' (Sổ tính trống).")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "11.1", "Khám phá cấu trúc bảng tính Excel", [
        "Cột (Columns): Được ký hiệu bằng các chữ cái in hoa (A, B, C, D...) chạy dọc từ trên xuống dưới.",
        "Dòng (Rows): Được ký hiệu bằng các chữ số (1, 2, 3, 4...) chạy ngang từ trái sang phải.",
        "Ô (Cell): Là giao điểm giữa một cột và một dòng. Tên ô (Địa chỉ ô) được ghép bởi Tên Cột đứng trước và Tên Dòng đứng sau. Ví dụ: Ô nằm ở Cột B và Dòng 5 có địa chỉ là B5.",
        "Thanh công thức (Formula Bar): Nằm ngay phía trên bảng tính, hiển thị chính xác nội dung thực tế hoặc công thức đang nằm bên trong ô được chọn."
    ])

    add_step(doc, "11.2", "Nguyên tắc vàng phân biệt 3 kiểu dữ liệu (Quy tắc tự canh lề của Excel)", [
        "Khi bạn gõ dữ liệu vào một ô và nhấn Enter, Excel sẽ tự động nhận diện kiểu dữ liệu qua cách căn lề mặc định:",
        "1. Dữ liệu Văn bản (Chữ): Excel TỰ ĐỘNG CĂN LỀ TRÁI. Ví dụ: 'Hà Nội', 'Nguyễn Văn A', 'Máy in HP'.",
        "2. Dữ liệu Dạng Số: Excel TỰ ĐỘNG CĂN LỀ PHẢI. Ví dụ: 50, 1500000. Dữ liệu này có thể cộng trừ nhân chia được.",
        "3. Dữ liệu Ngày tháng (Date): Excel cũng TỰ ĐỘNG CĂN LỀ PHẢI. Định dạng ngày tháng tại Việt Nam là: Ngày/Tháng/Năm (dd/mm/yyyy). Ví dụ: 05/09/2026.",
        "DẤU HIỆU CẢNH BÁO NGUY HIỂM: Nếu bạn nhập số tiền (ví dụ 100000) hoặc ngày tháng mà thấy nó tự động dạt sang LỀ TRÁI, nghĩa là Excel đang hiểu đó là một dòng chữ (Text). Bạn sẽ không thể cộng trừ hay tính toán được trên ô này!"
    ])

    add_step(doc, "11.3", "Kỹ thuật sửa ô bằng phím F2 (Tránh gõ đè làm mất dữ liệu)", [
        "Sai lầm phổ biến: Nhấp chuột vào ô rồi gõ luôn khiến toàn bộ nội dung cũ biến mất sạch sẽ.",
        "Cách sửa đúng: Nhấp chuột chọn ô cần sửa > NHẤN PHÍM F2 TRÊN BÀN PHÍM. Con trỏ soạn thảo sẽ nhấp nháy ở cuối ô, cho phép bạn dùng phím mũi tên di chuyển đến vị trí cần sửa để xóa bớt hoặc thêm chữ.",
        "Hoặc nhấp đúp chuột trái trực tiếp vào vị trí chữ cần sửa trong ô."
    ])

    add_step(doc, "11.4", "Chèn thêm hoặc Xóa bớt Dòng và Cột", [
        "Chèn thêm 1 dòng mới: Nhấp chuột phải vào con số đại diện cho dòng (ở mép ngoài cùng bên trái bảng) > Chọn 'Insert'. Một dòng trắng mới sẽ được chèn lên phía trên dòng đang chọn.",
        "Chèn thêm 1 cột mới: Nhấp chuột phải vào chữ cái đại diện cho cột (ở mép trên cùng bảng) > Chọn 'Insert'. Một cột trắng mới sẽ được chèn sang bên trái cột đang chọn.",
        "Xóa dòng hoặc cột: Nhấp chuột phải vào tên Dòng hoặc tên Cột muốn bỏ > Chọn 'Delete'."
    ])

    add_screenshot_placeholder(doc, "Giao diện bảng tính Excel: Hộp địa chỉ Name Box hiển thị ô B5, thanh Formula Bar và minh họa dữ liệu Chữ căn trái, Số căn phải.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Học viên nhận diện chính xác địa chỉ bất kỳ ô nào trên màn hình, nhập số và ngày tháng đúng cách (tự dạt sang phải), biết dùng F2 để sửa nhanh dữ liệu.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI NHẬP NGÀY THÁNG BỊ ĐẢO NGƯỢC THÀNH THÁNG/NGÀY", [
        "Hiện tượng: Bạn gõ 15/09/2026 thì ô dạt sang phải (đúng), nhưng gõ 05/09/2026 thì máy lại hiểu là ngày 09 tháng 5 (bị đảo tháng trước ngày sau do máy tính dùng cài đặt kiểu Mỹ mm/dd/yyyy).",
        "Cách khắc phục: Bôi đen cột ngày tháng > Nhấp chuột phải chọn 'Format Cells' > Chọn mục 'Date' > Tại ô 'Locale (location)' nhấp chuột chọn 'Vietnamese' > Chọn kiểu hiển thị ngày tháng năm rõ ràng (dd/mm/yyyy)."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở Excel, tạo một bảng tính tại Sheet 1. Nhập các dữ liệu sau vào các ô tương ứng:")
    add_p(doc, "• Ô A1: 'SỔ THEO DÕI HỒ SƠ CƠ QUAN' (Dữ liệu chữ)")
    add_p(doc, "• Ô A3: 'Mã hồ sơ' | Ô B3: 'Họ tên công dân' | Ô C3: 'Ngày nộp' | Ô D3: 'Số tiền lệ phí'")
    add_p(doc, "• Ô A4: 'HS-001' | Ô B4: 'Trần Văn Bình' | Ô C4: '10/09/2026' | Ô D4: '50000'")
    add_p(doc, "Quan sát xem ô C4 và D4 có tự động căn lề phải không. Dùng phím F2 sửa 'Trần Văn Bình' thành 'Trần Văn An'.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Phân biệt đúng cột, dòng và đọc chính xác địa chỉ ô", "Đọc đúng ô A1, B3, C4..."),
        ("Dữ liệu số và ngày tháng tự động căn mép phải", "Các con số dạt sang bên phải ô một cách tự nhiên"),
        ("Biết dùng phím F2 để chỉnh sửa nội dung trong ô", "Sửa được từ ngữ mà không làm mất toàn bộ nội dung cũ"),
        ("Chèn và xóa dòng/cột thành thạo bằng chuột phải", "Thêm bớt dòng cột nhanh chóng đúng vị trí")
    ])

    # BÀI 12
    add_heading_2(doc, "Bài 12: Định dạng bảng biểu chuyên nghiệp, Kẻ khung và Định dạng tiền tệ")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Kẻ đường viền khung bảng (Borders) sắc nét để khi in ra không bị trang giấy trắng trơn.")
    add_p(doc, "• Gộp nhiều ô thành một ô tiêu đề lớn (Merge & Center) và Ngắt dòng tự động trong ô (Wrap Text).")
    add_p(doc, "• Định dạng hiển thị số tiền có dấu chấm phân cách hàng nghìn (ví dụ: 1.500.000 đ) chuẩn mực kế toán Việt Nam.")
    add_p(doc, "• Điều chỉnh độ rộng cột và độ cao dòng vừa vặn với nội dung chữ.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Mở tệp Excel đã nhập dữ liệu ở Bài 11 để bắt đầu trang trí bảng biểu chuyên nghiệp.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "LƯU Ý CỐT TỬ: CÁC ĐƯỜNG LƯỚI XÁM TRONG EXCEL KHÔNG ĐƯỢC IN RA", [
        "Nhiều người mới nhìn thấy trên màn hình có các đường kẻ ô màu xám thì lầm tưởng rằng in ra giấy sẽ có sẵn bảng kẻ. THỰC TẾ: Đó chỉ là lưới ảo (Gridlines) để bạn nhìn vị trí ô. Nếu không thực hiện thao tác Kẻ viền (Borders), khi in ra giấy tài liệu của bạn sẽ là một trang giấy trắng tinh chỉ có chữ!",
        "Vì vậy: BẮT BUỘC phải thực hiện thao tác Kẻ khung viền (Borders) cho mọi bảng biểu."
    ], callout_type="warning")

    add_step(doc, "12.1", "Kẻ khung viền cho bảng tính (Borders)", [
        "Dùng chuột trái kéo bôi đen toàn bộ vùng bảng tính cần kẻ viền (từ ô góc trên bên trái đến ô góc dưới bên phải).",
        "Trên thanh công cụ, vào thẻ 'Home' > Tìm nhóm công cụ 'Font'.",
        "Nhấp chuột vào dấu mũi tên cạnh biểu tượng hình ô vuông có 4 phần (nút Borders).",
        "Nhấp chọn dòng 'All Borders' (Tất cả đường viền). Ngay lập tức, toàn bộ bảng tính sẽ được kẻ các đường viền đen đậm rõ nét."
    ])

    add_step(doc, "12.2", "Gộp ô tiêu đề (Merge & Center) và Tự động ngắt dòng (Wrap Text)", [
        "Gộp ô tiêu đề: Muốn dòng chữ 'SỔ THEO DÕI HỒ SƠ' nằm chính giữa bảng từ cột A đến cột F:",
        "  1. Kéo chuột bôi đen các ô từ A1 đến F1.",
        "  2. Nhấp chuột vào nút 'Merge & Center' (Gộp và Căn giữa) trong nhóm Alignment trên thẻ Home. Các ô sẽ gom lại thành 1 ô duy nhất và chữ tự động căn giữa.",
        "Ngắt dòng trong ô (Wrap Text): Khi tiêu đề một cột quá dài (ví dụ: 'Ngày hoàn thành thực tế') làm tràn sang ô bên cạnh:",
        "  1. Nhấp chuột chọn ô đó.",
        "  2. Nhấp vào nút 'Wrap Text' trên thanh công cụ. Chữ sẽ tự động rớt xuống dòng thứ 2 bên trong chính ô đó, giúp cột không bị kéo dài quá mức."
    ])

    add_step(doc, "12.3", "Định dạng số tiền có dấu phân cách hàng nghìn chuẩn Việt Nam", [
        "Bôi đen cột chứa số tiền (ví dụ cột D).",
        "Nhấp chuột phải vào vùng bôi đen > Chọn 'Format Cells' (Định dạng ô).",
        "Tại thẻ 'Number' > Chọn danh mục 'Number' ở cột bên trái.",
        "Tại ô 'Decimal places' (Số chữ số thập phân sau dấu phẩy): Chỉnh về số '0'.",
        "TÍCH CHỌN VÀO Ô: 'Use 1000 Separator (,)' (Sử dụng dấu phân cách hàng nghìn).",
        "Bấm 'OK'. Ngay lập tức, số '50000' sẽ được hiển thị chuyên nghiệp thành '50,000' hoặc '50.000'."
    ])

    add_step(doc, "12.4", "Tô màu tiêu đề bảng và Chỉnh độ rộng cột tự động", [
        "Tô màu nền tiêu đề: Bôi đen dòng tiêu đề các cột (dòng 3) > Bấm vào biểu tượng 'Thùng sơn đổ màu' (Fill Color) trên thẻ Home > Chọn màu xanh dương nhạt hoặc xám nhạt trang nhã > Bấm nút In đậm 'B' (Ctrl + B).",
        "Co giãn độ rộng cột tự động: Đưa con trỏ chuột lên thanh chữ cái (A, B, C...) ở mép ranh giới giữa 2 cột sao cho chuột biến thành hình mũi tên 2 chiều đen <|> > NHẤP ĐÚP CHUỘT TRÁI. Cột sẽ tự động co giãn vừa khít với dòng chữ dài nhất."
    ])

    add_screenshot_placeholder(doc, "Bảng tính Excel sau khi định dạng: Tiêu đề gộp ô Merge & Center, nền tiêu đề cột màu xanh nhạt, kẻ khung All Borders và số tiền phân cách hàng nghìn.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Bảng tính có hình thức chỉn chu, tiêu đề nổi bật, các dòng tiêu đề cột rõ ràng, các con số tài chính có dấu phân cách hàng nghìn giúp quan sát số tiền hàng triệu, hàng tỷ không bao giờ bị nhầm lẫn.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI Ô BỊ HIỂN THỊ DÃY DẤU THĂNG (###)", [
        "Hiện tượng: Sau khi định dạng số tiền hoặc ngày tháng, ô biến thành các dấu '###' hoặc '#######'.",
        "Nguyên nhân: Cột quá hẹp, không đủ chỗ trống để hiển thị hết các con số dài.",
        "Cách khắc phục cực dễ: Đưa chuột lên vạch ranh giới cột ở trên cùng và NHẤP ĐÚP CHUỘT TRÁI (hoặc kéo cột rộng ra sang bên phải). Các con số sẽ lập tức hiển thị đầy đủ."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở một bảng tính mới, tạo 'BẢNG THEO DÕI VĂN PHÒNG PHẨM QUÝ III':")
    add_p(doc, "1. Hàng 1: Tiêu đề lớn, gộp ô từ A1 đến E1, cỡ chữ 15 đậm, căn giữa.")
    add_p(doc, "2. Hàng 2: Các tiêu đề cột: STT | Tên văn phòng phẩm | Đơn vị tính | Số lượng | Đơn giá.")
    add_p(doc, "3. Nhập dữ liệu cho 3 mặt hàng: Giấy A4 Double A (5 ram, giá 75000), Bút bi Thiên Long (20 chiếc, giá 4500), Kẹp bướm (10 hộp, giá 18000).")
    add_p(doc, "4. Kẻ All Borders cho toàn bộ bảng; cột Đơn giá định dạng có dấu phân cách hàng nghìn; tô màu nền dòng tiêu đề cột.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Bảng tính được kẻ viền All Borders sắc nét", "Toàn bộ các ô dữ liệu có đường viền đen bao quanh"),
        ("Tiêu đề lớn được gộp ô Merge & Center chính giữa", "Tiêu đề nằm cân xứng ngay giữa bảng tính"),
        ("Cột đơn giá hiển thị dấu phân cách hàng nghìn (75,000)", "Dễ đọc, không bị nhầm lẫn số chữ số 0"),
        ("Không có ô nào bị hiển thị lỗi ###", "Tất cả các cột đều đủ độ rộng hiển thị trọn vẹn số liệu")
    ])

    # BÀI 13
    add_heading_2(doc, "Bài 13: Các phép tính số học và Các hàm thống kê cơ bản nhất")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Hiểu bản chất công thức trong Excel: Bắt buộc phải bắt đầu bằng dấu bằng (=).")
    add_p(doc, "• Thực hiện thành thạo các phép toán số học: Cộng (+), Trừ (-), Nhân (*), Chia (/).")
    add_p(doc, "• Kéo sao chép công thức tự động trong tích tắc bằng nút góc dưới ô (Fill Handle).")
    add_p(doc, "• Làm chủ 5 hàm thống kê văn phòng phổ biến nhất: SUM, AVERAGE, MIN, MAX, COUNT/COUNTA.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Sử dụng tiếp bảng tính Văn phòng phẩm đã nhập ở Bài 12 để tính tiền và thống kê.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "QUY TẮC VÀNG KHI VIẾT CÔNG THỨC VÀ HÀM TRONG EXCEL", [
        "1. Luôn gõ dấu bằng (=) đầu tiên. Nếu không có dấu =, Excel sẽ coi công thức là một dòng chữ bình thường và không tính toán.",
        "2. Tên hàm không phân biệt chữ hoa hay chữ thường (gõ =sum hay =SUM đều đúng như nhau).",
        "3. Dùng địa chỉ ô chứ không gõ con số chết: Phải viết '=D3*E3' chứ không viết '=5*75000'. Nhờ vậy khi thay đổi số lượng ở ô D3, số tiền tự động cập nhật ngay lập tức!"
    ], callout_type="tip")

    add_step(doc, "13.1", "Các phép toán số học và Kỹ thuật kéo công thức tự động (Fill Handle)", [
        "Tính Thành tiền (Cột F = Số lượng * Đơn giá):",
        "  1. Nhấp chuột vào ô F3 (dòng mặt hàng đầu tiên).",
        "  2. Gõ dấu bằng: =",
        "  3. Nhấp chuột trái vào ô D3 (Số lượng) > gõ dấu sao: * > nhấp chuột trái vào ô E3 (Đơn giá). Trên màn hình hiện: =D3*E3",
        "  4. Nhấn phím Enter. Kết quả thành tiền hiển thị ngay lập tức.",
        "Kéo công thức cho các dòng dưới (Fill Handle):",
        "  1. Nhấp chuột chọn lại ô F3.",
        "  2. Rê chuột đến góc dưới cùng bên phải của ô F3, con trỏ chuột sẽ chuyển từ hình dấu cộng màu trắng rỗng sang hình DẤU CỘNG MÀU ĐEN NHỎ (+).",
        "  3. Nhấn giữ chuột trái và kéo thẳng xuống các dòng dưới, rồi thả tay ra (hoặc NHẤP ĐÚP CHUỘT TRÁI vào dấu cộng đen). Toàn bộ các dòng còn lại sẽ được tính toán tự động trong 0.1 giây!"
    ])

    add_step(doc, "13.2", "Hàm tính Tổng cộng: SUM", [
        "Ý nghĩa: Cộng tổng toàn bộ các con số trong một vùng ô.",
        "Cú pháp: =SUM(ô_bắt_đầu : ô_kết_thúc)",
        "Ví dụ tính tổng tiền cả bảng (từ F3 đến F5):",
        "  1. Tại ô F6 (dưới cùng cột Thành tiền), gõ: =SUM(F3:F5)",
        "  2. Nhấn Enter để xem kết quả tổng kinh phí.",
        "Mẹo cực nhanh (AutoSum): Đặt chuột tại ô F6 > Nhấn tổ hợp phím 'Alt + =' (hoặc bấm nút AutoSum Σ trên thẻ Home) > Excel tự chọn vùng > Nhấn Enter."
    ])

    add_step(doc, "13.3", "Hàm tính Giá trị trung bình: AVERAGE", [
        "Ý nghĩa: Tính trung bình cộng của một dãy số.",
        "Cú pháp: =AVERAGE(vùng_dữ_liệu)",
        "Ví dụ tính đơn giá trung bình của các mặt hàng: Tại ô cần tính gõ: =AVERAGE(E3:E5) > Nhấn Enter."
    ])

    add_step(doc, "13.4", "Hàm tìm Giá trị Lớn nhất (MAX) và Nhỏ nhất (MIN)", [
        "Hàm MAX: Tìm con số cao nhất trong vùng. Cú pháp: =MAX(E3:E5) (tìm mặt hàng có đơn giá đắt nhất).",
        "Hàm MIN: Tìm con số thấp nhất trong vùng. Cú pháp: =MIN(E3:E5) (tìm mặt hàng có đơn giá rẻ nhất)."
    ])

    add_step(doc, "13.5", "Hàm Đếm số lượng: COUNT và COUNTA", [
        "Hàm COUNT: Đếm xem trong dãy có bao nhiêu ô chứa DỮ LIỆU SỐ. Cú pháp: =COUNT(D3:D5).",
        "Hàm COUNTA: Đếm tất cả các ô CÓ CHỨA NỘI DUNG (cả chữ lẫn số, không rỗng). Cú pháp: =COUNTA(B3:B5) (rất hay dùng để đếm xem có bao nhiêu người nộp hồ sơ hoặc bao nhiêu mặt hàng đã nhập)."
    ])

    add_screenshot_placeholder(doc, "Vị trí nút dấu cộng đen nhỏ Fill Handle ở góc ô F3 và cách viết công thức =SUM(F3:F5) tính tổng tiền.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Học viên tự tay lập được các phép nhân chia tính thành tiền, tự động tính tổng kinh phí bằng hàm SUM, biết kéo công thức Fill Handle để làm việc nhanh gấp 10 lần làm thủ công.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "CÁC LỖI TÍNH TOÁN CẦN TRÁNH", [
        "Lỗi #VALUE!: Xảy ra khi công thức thực hiện phép tính trên ô chứa chữ (ví dụ lấy ô chứa chữ 'Không có' nhân với ô chứa số 50000). Cách sửa: Kiểm tra lại các ô trong phép tính xem có ô nào bị gõ nhầm chữ hoặc dấu cách không.",
        "Lỗi #NAME?: Xảy ra khi bạn gõ sai chính tả tên hàm (ví dụ gõ nhầm thành '=SUN' thay vì '=SUM'). Cách sửa: Nhấp vào ô bấm F2 và sửa lại đúng tên hàm tiếng Anh."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở rộng bảng tính Văn phòng phẩm ở Bài 12 với 5 dòng mặt hàng:")
    add_p(doc, "1. Dùng phép nhân =Số_lượng * Đơn_giá và kéo Fill Handle để tính Thành tiền cho cả 5 mặt hàng.")
    add_p(doc, "2. Tại dòng cuối cùng, dùng hàm =SUM(...) tính Tổng tiền phải thanh toán.")
    add_p(doc, "3. Ở phía dưới bảng, dùng các hàm: =AVERAGE(...) tính Đơn giá trung bình; =MAX(...) tìm Thành tiền lớn nhất; =MIN(...) tìm Thành tiền nhỏ nhất; =COUNTA(...) đếm tổng số mặt hàng.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Mọi công thức đều bắt đầu bằng dấu bằng (=)", "Công thức tính ra kết quả số chuẩn xác"),
        ("Áp dụng đúng kỹ thuật kéo công thức Fill Handle", "Tính toán đồng loạt cho nhiều dòng trong vài giây"),
        ("Sử dụng thành thạo hàm tính tổng =SUM()", "Dòng tổng cộng tính đúng toàn bộ các khoản mục"),
        ("Phân biệt đúng hàm AVERAGE, MIN, MAX và COUNT/COUNTA", "Thống kê đầy đủ các chỉ số văn phòng cơ bản")
    ])

    # BÀI 14
    add_heading_2(doc, "Bài 14: Hàm điều kiện IF và Định dạng cảnh báo màu tự động (Conditional Formatting)")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Hiểu bản chất tư duy logic của hàm điều kiện IF trong quản lý công việc.")
    add_p(doc, "• Viết được hàm IF đơn giản để tự động đánh giá: Đạt / Không đạt; Đúng hạn / Quá hạn; Cần bổ sung.")
    add_p(doc, "• Sử dụng công cụ Conditional Formatting để tự động tô màu Đỏ cảnh báo cho các hồ sơ quá hạn hoặc số tiền vượt ngân sách.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Mở tệp Sổ theo dõi hồ sơ đã tạo ở Bài 11.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "CẤU TRÚC HÀM IF - 'NẾU THÌ' TRONG EXCEL", [
        "Hàm IF có đúng 3 thành phần ngăn cách nhau bởi dấu phẩy (hoặc chấm phẩy tùy cài đặt máy):",
        "Cú pháp: =IF(Điều_kiện_kiểm_tra, Giá_trị_nếu_ĐÚNG, Giá_trị_nếu_SAI)",
        "Quy tắc bắt buộc: Nếu kết quả trả về là chữ (văn bản), bạn BẮT BUỘC PHẢI ĐẶT TRONG DẤU NGOẶC KÉP \" \".",
        "Ví dụ: =IF(C4<=D4, \"Đúng hạn\", \"Quá hạn\")",
        "(Giải nghĩa: Nếu Ngày hoàn thành C4 nhỏ hơn hoặc bằng Hạn trả kết quả D4, hãy điền chữ 'Đúng hạn', ngược lại nếu trễ hạn thì điền chữ 'Quá hạn')."
    ], callout_type="tip")

    add_step(doc, "14.1", "Áp dụng hàm IF để tự động phân loại hồ sơ", [
        "Thêm một cột mới có tên: 'Đánh giá tiến độ' tại cột E.",
        "Tại ô E4, gõ công thức:",
        "  =IF(C4<=D4, \"Đúng hạn\", \"Quá hạn\")",
        "Nhấn Enter. Excel sẽ so sánh 2 ngày tháng và tự động điền chữ.",
        "Rê chuột vào góc dưới ô E4 kéo Fill Handle xuống cho các hồ sơ còn lại."
    ])

    add_step(doc, "14.2", "Cài đặt Tự động tô màu đỏ cảnh báo hồ sơ Quá hạn (Conditional Formatting)", [
        "Bôi đen toàn bộ cột Đánh giá tiến độ (từ E4 đến E10).",
        "Trên thanh công cụ thẻ Home > Tìm nhóm 'Styles' > Nhấp vào nút 'Conditional Formatting' (Định dạng có điều kiện).",
        "Rê chuột vào dòng 'Highlight Cells Rules' (Quy tắc làm nổi bật ô) > Nhấp chọn dòng 'Text that Contains...' (Ô chứa đoạn chữ...).",
        "Cửa sổ hiện ra, tại ô bên trái bạn gõ đúng từ: Quá hạn.",
        "Tại ô bên phải, chọn kiểu màu: 'Light Red Fill with Dark Red Text' (Nền đỏ nhạt, chữ đỏ đậm).",
        "Bấm nút 'OK'. Ngay lập tức, bất kỳ ô nào có chữ 'Quá hạn' sẽ tự động rực lên màu đỏ cảnh báo. Nếu sau này ngày tháng thay đổi thành đúng hạn, màu đỏ sẽ tự động biến mất!"
    ])

    add_screenshot_placeholder(doc, "Menu Conditional Formatting > Highlight Cells Rules > Text that Contains và các ô có chữ Quá hạn được tô màu đỏ nổi bật.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Bảng theo dõi hồ sơ trở nên thông minh: tự động đưa ra đánh giá tiến độ mà không cần người dùng tự gõ tay từng dòng, các trường hợp vi phạm thời hạn được tô màu cảnh báo trực quan.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI DẤU PHÂN CÁCH ĐỐI SỐ (PHẨY HAY CHẤM PHẨY)", [
        "Khi gõ công thức =IF(C4<=D4, \"Đúng hạn\", \"Quá hạn\") máy tính báo lỗi cú pháp 'There's a problem with this formula'.",
        "Nguyên nhân: Một số máy tính tại Việt Nam cài đặt dấu phân cách công thức là DẤU CHẤM PHẨY (;) thay vì DẤU PHẨY (,).",
        "Cách khắc phục: Hãy thử thay toàn bộ dấu phẩy trong công thức thành dấu chấm phẩy: =IF(C4<=D4; \"Đúng hạn\"; \"Quá hạn\") rồi nhấn Enter."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Tạo bảng điểm kiểm tra nghiệp vụ văn phòng gồm các cột: STT, Họ và tên, Điểm kiểm tra (thang điểm 10).")
    add_p(doc, "1. Nhập điểm cho 5 cán bộ (từ 4 đến 9 điểm).")
    add_p(doc, "2. Tạo cột 'Kết quả': Dùng hàm IF sao cho nếu Điểm >= 5 thì ghi 'Đạt', ngược lại ghi 'Học lại'.")
    add_p(doc, "3. Dùng Conditional Formatting tô màu Vàng cho chữ 'Đạt' và màu Đỏ cho chữ 'Học lại'.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Viết đúng cú pháp hàm IF có chữ trong dấu ngoặc kép \" \"", "Hàm trả về kết quả chữ chính xác"),
        ("Xác định đúng dấu phân cách đối số của máy tính (, hoặc ;)", "Công thức không bị báo lỗi cú pháp"),
        ("Cài đặt thành công tô màu cảnh báo tự động", "Các trường hợp cảnh báo đổi màu trực quan, chuyên nghiệp")
    ])

    # BÀI 15
    add_heading_2(doc, "Bài 15: Sắp xếp, Lọc dữ liệu văn phòng và Cố định dòng tiêu đề")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cách Cố định dòng tiêu đề (Freeze Panes) để khi cuộn trang xuống sâu không bị mất tiêu đề cột.")
    add_p(doc, "• Bật tính năng Lọc tự động (Filter) để lọc nhanh danh sách hồ sơ theo cán bộ phụ trách hoặc theo trạng thái.")
    add_p(doc, "• Sắp xếp dữ liệu tăng dần hoặc giảm dần (theo ngày tháng, theo số tiền hoặc theo bảng chữ cái A-Z).")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Bảng dữ liệu có từ 10 dòng trở lên.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "15.1", "Cố định dòng tiêu đề khi cuộn trang (Freeze Panes)", [
        "Vấn đề: Khi bảng tính dài 100 dòng, cuộn chuột xuống dưới sẽ không còn nhìn thấy dòng tiêu đề cột, không biết cột nào là số tiền, cột nào là ngày tháng.",
        "Cách cố định dòng tiêu đề trên cùng:",
        "  1. Vào thẻ 'View' trên thanh công cụ Ribbon.",
        "  2. Tìm nút 'Freeze Panes' (Cố định khung).",
        "  3. Nhấp chọn dòng 'Freeze Top Row' (Cố định dòng trên cùng).",
        "Giờ đây, bạn thoải mái cuộn chuột xuống dưới, dòng tiêu đề vẫn 'dính chặt' ở mép trên màn hình cực kỳ tiện lợi!"
    ])

    add_step(doc, "15.2", "Bật tính năng Lọc dữ liệu tự động (Filter)", [
        "Nhấp chuột vào một ô bất kỳ bên trong dòng tiêu đề của bảng.",
        "Vào thẻ 'Data' trên thanh công cụ > Nhấp chuột vào biểu tượng hình chiếc phễu lớn: nút 'Filter' (Phím tắt: Ctrl + Shift + L).",
        "Quan sát các ô tiêu đề: Mỗi cột sẽ xuất hiện thêm một hình mũi tên chỉ xuống nhỏ ở góc phải.",
        "Cách lọc dữ liệu: Nhấp chuột vào mũi tên ở cột 'Cán bộ thụ lý' > Bỏ dấu tích ở ô '(Select All)' > Chỉ tích chọn vào tên đồng chí 'Nguyễn Văn An' > Bấm OK. Toàn bộ hồ sơ của các cán bộ khác sẽ tạm thời ẩn đi, chỉ hiện duy nhất hồ sơ của đồng chí An.",
        "Cách bỏ lọc (hiển thị lại toàn bộ): Nhấp lại vào biểu tượng chiếc phễu trên cột đó > Chọn 'Clear Filter From...'"
    ])

    add_step(doc, "15.3", "Sắp xếp dữ liệu (Sort)", [
        "Sắp xếp theo thứ tự thời gian hoặc số tiền:",
        "  • Muốn xếp các hồ sơ có ngày nộp cũ nhất lên đầu: Nhấp vào mũi tên lọc ở cột Ngày nộp > Chọn 'Sort Oldest to Newest' (Từ cũ nhất đến mới nhất).",
        "  • Muốn xếp số tiền từ cao nhất đến thấp nhất: Nhấp vào mũi tên lọc ở cột Số tiền > Chọn 'Sort Largest to Smallest' (Từ lớn đến bé)."
    ])

    add_screenshot_placeholder(doc, "Biểu tượng nút Freeze Panes trên thẻ View và các mũi tên lọc Filter hình tam giác trên dòng tiêu đề bảng tính.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Người học tự tin làm chủ các bảng số liệu hàng trăm dòng; lọc ra đúng thông tin lãnh đạo yêu cầu chỉ trong 3 giây; dòng tiêu đề luôn hiển thị rõ ràng.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI SẮP XẾP LÀM XÁO TRỘN BẢNG DỮ LIỆU", [
        "Nguy cơ: Nếu bạn chỉ bôi đen 1 cột duy nhất rồi bấm Sắp xếp (Sort), cột đó sẽ bị đảo lộn trong khi các cột khác đứng yên, khiến 'họ tên của người này ghép nhầm với số tiền của người khác'!",
        "Quy tắc an toàn: Luôn đặt con trỏ chuột vào dòng tiêu đề và dùng tính năng Sort qua nút Filter, Excel sẽ tự động đảo cả hàng đi cùng nhau, đảm bảo dữ liệu luôn chính xác 100%."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở Sổ theo dõi hồ sơ cơ quan (có ít nhất 8 dòng dữ liệu với 3 cán bộ phụ trách khác nhau):")
    add_p(doc, "1. Thực hiện Freeze Panes để cố định dòng tiêu đề.")
    add_p(doc, "2. Bật Filter (Ctrl + Shift + L). Thực hiện lọc ra toàn bộ các hồ sơ có trạng thái 'Quá hạn'.")
    add_p(doc, "3. Bỏ lọc, sau đó sắp xếp cột Ngày nộp theo thứ tự ngày tăng dần.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Dòng tiêu đề được cố định cố định khi cuộn trang", "Cuộn chuột xuống dòng 50 vẫn nhìn thấy rõ tiêu đề cột"),
        ("Bật tắt thành thạo nút Filter bằng phím tắt Ctrl + Shift + L", "Biểu tượng mũi tên tam giác xuất hiện/ẩn đi nhanh chóng"),
        ("Lọc chính xác dữ liệu theo điều kiện mong muốn", "Bảng tính chỉ hiển thị đúng các dòng thỏa mãn bộ lọc"),
        ("Sắp xếp toàn bộ bảng dữ liệu không bị lệch dòng", "Dữ liệu trên từng dòng luôn đi liền với nhau")
    ])

    # BÀI 16
    add_heading_2(doc, "Bài 16: Căn chỉnh trang in bảng tính Excel lên khổ giấy A4 vừa khít")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Xóa bỏ vĩnh viễn 'nỗi ám ảnh kinh hoàng' khi in Excel bị mất cột hoặc cột bị rớt lẻ loi sang trang khác.")
    add_p(doc, "• Thiết lập bảng tính nằm vừa vặn trọn vẹn trên 1 trang giấy A4 (theo chiều Ngang hoặc Dọc).")
    add_p(doc, "• Cài đặt tính năng tự động lặp lại dòng tiêu đề cột trên mọi trang in khi in bảng biểu dài.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Một bảng tính Excel hoàn chỉnh có nhiều cột.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "16.1", "Chọn hướng giấy và Khổ giấy chuẩn trong Excel", [
        "Vào thẻ 'Page Layout' trên thanh công cụ Ribbon.",
        "Mục 'Size': Nhấp chọn đúng khổ 'A4'.",
        "Mục 'Orientation' (Hướng giấy):",
        "  • Nếu bảng có ít cột (dưới 6 cột): Chọn 'Portrait' (Hướng giấy Đứng).",
        "  • Nếu bảng có nhiều cột (từ 7 cột trở lên): Chọn 'Landscape' (Hướng giấy Ngang) để có không gian rộng rãi."
    ])

    add_step(doc, "16.2", "Kỹ thuật 'Ép' toàn bộ các cột nằm vừa khít 1 trang (Fit All Columns on One Page)", [
        "Nhấn phím tắt 'Ctrl + P' để vào cửa sổ xem trước bản in (Print Preview).",
        "Nhìn xuống dòng dưới cùng của danh mục cài đặt bên trái, nhấp chuột vào mục 'No Scaling' (Không co giãn).",
        "Nhấp chọn dòng: 'Fit All Columns on One Page' (Co toàn bộ cột vừa khít trên một trang).",
        "Quan sát khung xem trước bên phải: Ngay lập tức, Excel sẽ tự động thu nhỏ tỷ lệ phần trăm vừa khít để tất cả các cột hiển thị trọn vẹn trên bề ngang trang giấy A4. Bạn sẽ không bao giờ lo bị rớt 1-2 cột sang trang thứ hai nữa!"
    ])

    add_step(doc, "16.3", "Cài đặt Lặp lại dòng tiêu đề khi in nhiều trang (Rows to repeat at top)", [
        "Khi bảng tính dài 5-10 trang, từ trang thứ 2 trở đi sẽ không có dòng tiêu đề, người đọc không biết cột nào là gì. Hãy xử lý như sau:",
        "1. Quay lại bảng tính, vào thẻ 'Page Layout' > Nhấp vào nút 'Print Titles' (Tiêu đề in).",
        "2. Hộp thoại Page Setup hiện ra ở thẻ 'Sheet'.",
        "3. Nhấp chuột vào ô: 'Rows to repeat at top' (Dòng cần lặp lại ở mép trên).",
        "4. Dùng chuột nhấp chọn vào dòng tiêu đề bảng trên màn hình (ví dụ dòng 3 hoặc dòng 2:3).",
        "5. Bấm nút 'OK'. Khi in ra, dù bảng dài 50 trang thì trang nào cũng tự động có sẵn dòng tiêu đề chuẩn mực ở đầu trang!"
    ])

    add_screenshot_placeholder(doc, "Cửa sổ Print Preview trong Excel: Lựa chọn Fit All Columns on One Page và hộp thoại Page Setup lặp lại dòng Rows to repeat at top.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Bản in Excel ra giấy A4 đẹp như in từ phần mềm chuyên nghiệp: bảng nằm cân đối ở giữa trang, đủ cột, không bị méo lệch, trang nào cũng có tiêu đề rõ ràng.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "BẢNG IN RA BỊ DẠT SANG MỘT BÊN MÉP GIẤY", [
        "Hiện tượng: Bảng tính in ra bị lệch hẳn sang góc trên bên trái, bên phải để thừa một khoảng trắng lớn mất thẩm mỹ.",
        "Cách khắc phục: Nhấn Ctrl + P > Nhấp vào dòng 'Page Setup' ở đáy menu bên trái > Chọn thẻ 'Margins' > Nhìn xuống dưới cùng mục 'Center on page' (Căn giữa trang) và TÍCH CHỌN VÀO Ô 'Horizontally' (Theo chiều ngang) > Bấm OK. Bảng sẽ tự động nhảy vào chính giữa trang giấy."
    ], callout_type="tip")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở bảng tính có nhiều cột. Thực hiện:")
    add_p(doc, "1. Chọn hướng giấy Ngang (Landscape), khổ giấy A4.")
    add_p(doc, "2. Vào chế độ In (Ctrl + P) và chọn Fit All Columns on One Page.")
    add_p(doc, "3. Cài đặt căn giữa trang theo chiều ngang (Center on page Horizontally).")
    add_p(doc, "4. Cài đặt Print Titles lặp lại dòng tiêu đề cột.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Khổ giấy chọn đúng A4 và hướng giấy phù hợp (Ngang/Đứng)", "Bản in hiển thị rõ ràng cân đối"),
        ("Toàn bộ các cột nằm gọn gàng trên bề ngang trang giấy", "Không có cột nào bị đẩy sang trang in riêng lẻ"),
        ("Dòng tiêu đề được lặp lại trên mọi trang in", "Mở trang 2, trang 3 trên Print Preview đều có tiêu đề cột"),
        ("Bảng tính được căn giữa trang theo chiều ngang", "Khoảng cách lề trái và lề phải của bản in bằng nhau")
    ])

    # BÀI 17
    add_heading_2(doc, "Bài 17: Dự án thực hành tổng hợp Excel - Xây dựng 02 Bảng tính công sở hoàn chỉnh")
    
    add_p(doc, "Bài học này là dự án thực hành lớn giúp bạn làm chủ toàn bộ kỹ năng Excel: từ lập cấu trúc, nhập liệu, định dạng bảng, viết hàm tính toán thống kê đến thiết lập trang in hoàn chỉnh.")

    add_heading_3(doc, "Dự án 1: Xây dựng Sổ theo dõi tiếp nhận và giải quyết hồ sơ công việc")
    add_p(doc, "Mục đích: Giúp cán bộ văn phòng quản lý toàn bộ hồ sơ tiếp nhận của công dân/đơn vị gửi đến, tự động cảnh báo tiến độ xử lý.")
    
    add_callout(doc, "THÔNG SỐ DỰ ÁN 1 (SỔ THEO DÕI HỒ SƠ)", [
        "1. Tiêu đề lớn: 'SỔ THEO DÕI TIẾP NHẬN VÀ GIẢI QUYẾT HỒ SƠ HÀNH CHÍNH NĂM 2026' (Merge & Center A1:H1, nền xanh đậm, chữ trắng, cỡ 14).",
        "2. Cấu trúc 8 cột (Hàng 3):",
        "   • Cột A: STT (Căn giữa, 1, 2, 3...)",
        "   • Cột B: Mã hồ sơ (HS-01, HS-02...)",
        "   • Cột C: Tên tổ chức/cá nhân nộp",
        "   • Cột D: Nội dung tóm tắt",
        "   • Cột E: Cán bộ thụ lý (Nguyễn Văn An, Lê Thị Hoa, Trần Đình Trọng)",
        "   • Cột F: Ngày tiếp nhận (Định dạng dd/mm/yyyy)",
        "   • Cột G: Hạn trả kết quả (Định dạng dd/mm/yyyy)",
        "   • Cột H: Trạng thái xử lý",
        "3. Yêu cầu tính toán & Cảnh báo:",
        "   • Dùng hàm IF tại Cột H để xác định: nếu Ngày trả thực tế trễ hơn thì ghi 'Quá hạn', ngược lại ghi 'Đúng hạn'.",
        "   • Dùng Conditional Formatting tô màu Đỏ nổi bật cho các ô 'Quá hạn'.",
        "   • Dùng hàm COUNTA đếm tổng số hồ sơ đã tiếp nhận trong tháng.",
        "   • Bật tính năng Filter để sẵn sàng lọc hồ sơ theo từng cán bộ khi Lãnh đạo yêu cầu."
    ], callout_type="note")

    add_heading_3(doc, "Dự án 2: Xây dựng Bảng tổng hợp chi phí văn phòng phẩm & vật tư cơ quan")
    add_p(doc, "Mục đích: Lập bảng dự trù kinh phí chi tiết, tính toán thành tiền tự động và căn chỉnh in ấn chuẩn A4 để trình Lãnh đạo ký duyệt.")
    
    add_callout(doc, "THÔNG SỐ DỰ ÁN 2 (BẢNG KÊ CHI PHÍ VĂN PHÒNG PHẨM)", [
        "1. Tiêu đề: 'BẢNG TỔNG HỢP THANH TOÁN KINH PHÍ VĂN PHÒNG PHẨM QUÝ III/2026'.",
        "2. Cấu trúc bảng gồm 7 cột:",
        "   • Cột A: STT | Cột B: Danh mục vật tư | Cột C: Đơn vị tính | Cột D: Số lượng | Cột E: Đơn giá | Cột F: Thành tiền | Cột G: Ghi chú.",
        "3. Dữ liệu giả lập cho 5 mặt hàng: Giấy in A4, Bút ký lãnh đạo, Mực máy in Canon, Bìa còng lưu hồ sơ, Ghim dập.",
        "4. Yêu cầu kỹ thuật:",
        "   • Cột Thành tiền: Viết công thức =Số lượng * Đơn giá và kéo Fill Handle.",
        "   • Định dạng phân cách hàng nghìn cho Đơn giá và Thành tiền.",
        "   • Dòng Tổng cộng: Dùng hàm =SUM(...) tính tổng kinh phí toàn bảng.",
        "   • Dưới bảng tính các chỉ số: Chi phí trung bình (=AVERAGE), Khoản chi lớn nhất (=MAX).",
        "   • Thiết lập trang in: Chọn khổ A4 ngang, Fit All Columns on One Page, căn giữa trang Horizontally, sẵn sàng in ra giấy trình ký."
    ], callout_type="note")

    add_heading_3(doc, "Tiêu chí đánh giá bài thực hành tổng hợp Excel")
    add_checklist_table(doc, [
        ("Bảng biểu được kẻ All Borders, định dạng phông chữ Times New Roman đồng bộ", "Giao diện sạch sẽ, chuyên nghiệp"),
        ("Nhập đúng kiểu dữ liệu ngày tháng và số tiền (tự động căn lề phải)", "Không có số nào bị hiểu nhầm thành chữ"),
        ("Áp dụng chính xác hàm SUM, AVERAGE, MAX và hàm điều kiện IF", "Số liệu tự động nhảy chính xác khi thay đổi đầu vào"),
        ("Tự động tô màu cảnh báo hồ sơ Quá hạn bằng Conditional Formatting", "Màu sắc trực quan, phục vụ tốt công tác chỉ đạo"),
        ("Bản in Print Preview hiển thị vừa khít trên 1 trang A4, cân đối", "Không bị mất cột, sẵn sàng in ấn chất lượng cao")
    ])
