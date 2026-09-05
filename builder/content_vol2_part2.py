# -*- coding: utf-8 -*-
"""
Module biên soạn PHẦN 2 TẬP II:
BẢNG TÍNH TRỰC TUYẾN VÀ PHÂN TÍCH SỐ LIỆU VỚI GOOGLE SHEETS
"""

import os

from docx_helper import (
    add_heading_1, add_heading_2, add_heading_3, add_p,
    add_step, add_callout, add_screenshot_placeholder,
    add_styled_table, add_checklist_table, add_image
)

IMG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "images")

def build_vol2_part_2(doc):
    add_heading_1(doc, "PHẦN 2: BẢNG TÍNH TRỰC TUYẾN VÀ PHÂN TÍCH SỐ LIỆU VỚI GOOGLE SHEETS")
    add_p(doc, "Google Sheets (Google Trang tính) là công cụ bảng tính đám mây xuất sắc, giúp các cơ quan, đơn vị giải quyết triệt để bài toán: 'Cả phòng cùng cần nhập số liệu báo cáo nhưng mỗi người lại lưu một tệp Excel riêng lẻ trên máy mình'. Với Google Sheets, chỉ cần một đường liên kết duy nhất, 10 đến 50 cán bộ có thể đồng thời vào nhập dữ liệu mà không sợ bị đè bài. Phần này sẽ hướng dẫn bạn từ các bước nhập liệu cơ bản, viết hàm thông minh, đến việc làm chủ 2 'vũ khí bí mật' tối thượng của Google Sheets: Bộ lọc xem riêng (Filter Views) và Khóa bảo vệ dải ô công thức.")

    # BÀI 6
    add_heading_2(doc, "Bài 6: Khởi tạo Google Sheets, Nhập liệu và Định dạng bảng biểu trực tuyến")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cách mở một trang tính Google Sheets mới toanh trong 1 giây bằng lệnh 'sheets.new'.")
    add_p(doc, "• Thiết lập định dạng ngày tháng Việt Nam (dd/mm/yyyy) và tiền tệ chuẩn.")
    add_p(doc, "• Kẻ khung viền, tô màu tiêu đề bảng và điều chỉnh độ rộng cột trên giao diện web.")
    add_p(doc, "• Biết cách mở và chỉnh sửa trực tiếp tệp Excel (.xlsx) ngay trên Google Sheets.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Máy tính có kết nối mạng Internet và trình duyệt Google Chrome.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "6.1", "Khởi tạo bảng tính mới siêu tốc", [
        "Mở Google Chrome > Nhấp chuột vào thanh địa chỉ > Gõ chữ: 'sheets.new' rồi nhấn phím Enter.",
        "Một bảng tính mới trắng tinh sẽ mở ra ngay lập tức!",
        "Đổi tên bảng tính: Nhấp chuột vào ô chữ 'Bảng tính chưa có tiêu đề' ở góc trên cùng bên trái > Gõ tên: 'Theo_doi_Tien_do_Cong_tac_2026' > Nhấn Enter.",
        "Cơ chế tự lưu: Bất kỳ con số nào bạn gõ vào ô đều được Google tự động lưu lên đám mây trong vòng 0.5 giây."
    ])

    add_step(doc, "6.2", "Cài đặt Vùng & Ngôn ngữ Việt Nam (Để ngày tháng không bị đảo ngược)", [
        "Mẹo rất quan trọng cho người mới: Để Google Sheets hiểu đúng ngày trước tháng sau (dd/mm/yyyy):",
        "  1. Vào menu 'Tệp' (File) > Chọn 'Cài đặt' (Settings).",
        "  2. Tại thẻ 'Chung' > Mục 'Ngôn ngữ' (Locale): Nhấp chọn đúng 'Việt Nam'.",
        "  3. Mục 'Múi giờ' (Time zone): Chọn '(GMT+07:00) Giờ Việt Nam (Hà Nội)'.",
        "  4. Nhấp nút 'Lưu và tải lại' (Save and reload). Bây giờ bạn thoải mái gõ ngày tháng chuẩn Việt Nam mà không sợ bị nhảy nhầm sang kiểu Mỹ."
    ])

    add_step(doc, "6.3", "Định dạng Bảng biểu, Kẻ khung và Định dạng tiền tệ", [
        "Kẻ khung viền: Bôi đen vùng bảng tính > Nhấp chuột vào biểu tượng nút 'Đường viền' (Borders) có 4 ô vuông trên thanh công cụ > Nhấp chọn biểu tượng 'Tất cả đường viền' (All borders).",
        "Gộp ô tiêu đề: Bôi đen các ô tiêu đề > Nhấp chuột vào biểu tượng nút 'Hợp nhất ô' (Merge cells).",
        "Ngắt dòng trong ô: Chọn ô cần ngắt dòng > Nhấp vào nút 'Xuống dòng văn bản' (Text wrapping) > Chọn kiểu 'Xuống dòng' (Wrap).",
        "Định dạng tiền tệ phân cách hàng nghìn:",
        "  1. Bôi đen cột số tiền.",
        "  2. Vào menu 'Định dạng' (Format) > 'Số' (Number) > Chọn 'Số' (hoặc 'Tùy chỉnh đơn vị tiền tệ' > chọn 'Đồng Việt Nam ₫')."
    ])

    add_step(doc, "6.4", "Mở và Làm việc với tệp Excel trên Google Sheets", [
        "Khi đồng nghiệp gửi cho bạn một tệp Excel (.xlsx):",
        "  1. Trong Google Sheets, vào menu 'Tệp' > Chọn 'Mở' (Open) - Phím tắt Ctrl + O.",
        "  2. Nhấp vào thẻ 'Tải lên' (Upload) ở góc phải > Kéo tệp Excel từ máy tính thả vào khung hoặc bấm 'Duyệt' để chọn tệp.",
        "  3. Tệp Excel sẽ được mở trực tiếp trên Google Sheets. Bạn có thể tiếp tục tính toán và chia sẻ cho cả phòng cùng làm việc chung."
    ])

    add_image(doc, os.path.join(IMG_DIR, "vol2_img5_sheets_settings_borders.png"), "Giao diện Google Sheets: Menu Cài đặt vùng Việt Nam (Settings Locale Vietnam) và thanh công cụ kẻ viền All borders, gộp ô Merge.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Bảng tính Google Sheets được cấu hình chuẩn mực múi giờ và định dạng Việt Nam, bảng kẻ khung rõ nét, số liệu tài chính có dấu chấm/phẩy phân cách hàng nghìn chuyên nghiệp.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI GÕ NGÀY THÁNG BỊ HIỂU THÀNH CHỮ DO SAI ĐỊNH DẠNG", [
        "Hiện tượng: Bạn gõ 25/12/2026 nhưng ngày tháng lại dạt sang bên trái (bị hiểu là chữ).",
        "Nguyên nhân: Do Bảng tính chưa được thiết lập Vùng là 'Việt Nam' nên máy đang áp dụng chuẩn Mỹ (bắt buộc tháng phải đứng trước ngày, tối đa chỉ có tháng 12).",
        "Cách xử lý: Làm theo Bước 6.2 ở trên: Vào Tệp > Cài đặt > Đổi Vùng sang 'Việt Nam' và bấm Lưu & tải lại."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "1. Dùng lệnh 'sheets.new' tạo một bảng tính mới. Đổi tên thành 'Kinh_phi_Hoat_dong_2026'.")
    add_p(doc, "2. Vào Tệp > Cài đặt kiểm tra xem Vùng đã là Việt Nam chưa.")
    add_p(doc, "3. Nhập tiêu đề và kẻ khung bảng biểu gồm 5 cột: STT, Nội dung chi, Ngày chi, Số tiền, Người tạm ứng. Nhập thử 3 dòng dữ liệu và định dạng cột Số tiền có phân cách hàng nghìn.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Tạo bảng tính mới nhanh chóng bằng lệnh sheets.new", "Trang tính mở ra tức thì trên trình duyệt"),
        ("Thiết lập Vùng là Việt Nam để ngày tháng hiển thị chuẩn dd/mm/yyyy", "Dữ liệu ngày tháng tự động căn lề phải chuẩn xác"),
        ("Kẻ khung All borders cho bảng tính", "Các ô dữ liệu có đường viền đen bao quanh rõ nét"),
        ("Định dạng số tiền có dấu phân cách hàng nghìn", "Dễ đọc số tiền hàng triệu, không bị nhầm chữ số 0")
    ])

    # BÀI 7
    add_heading_2(doc, "Bài 7: Áp dụng công thức tính toán và Hàm văn phòng trên Google Sheets")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Viết thành thạo các hàm tính toán cốt lõi: SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF trên Google Sheets.")
    add_p(doc, "• Tận dụng tính năng Trí tuệ nhân tạo gợi ý công thức thông minh (Smart Suggestions) để làm việc nhanh hơn.")
    add_p(doc, "• Sao chép công thức tự động cho cả cột chỉ bằng 1 thao tác.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Sử dụng tiếp bảng tính kinh phí đã lập ở Bài 6.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "7.1", "Tận dụng tính năng Tự động gợi ý công thức thông minh của Google", [
        "Google Sheets cực kỳ thông minh trong việc đoán ý người dùng:",
        "Khi bạn đứng ở ô cuối cùng dưới một cột số tiền và vừa gõ dấu '=', Google Sheets sẽ tự động hiện một hộp gợi ý màu xám nhạt: '=SUM(D4:D8)'.",
        "Nếu công thức gợi ý đúng như ý bạn muốn: Bạn chỉ cần NHẤN PHÍM TAB (hoặc phím Enter) trên bàn phím. Toàn bộ công thức sẽ tự động điền vào ô mà bạn không cần phải gõ tay từng chữ cái!"
    ])

    add_step(doc, "7.2", "Viết các hàm thống kê văn phòng chuẩn mực", [
        "Các hàm trên Google Sheets có cú pháp và cách hoạt động giống 100% so với Microsoft Excel:",
        "  • Hàm tính Tổng: =SUM(D4:D10) (Nhấn Enter để tính tổng kinh phí).",
        "  • Hàm tính Trung bình: =AVERAGE(D4:D10) (Tính chi phí bình quân mỗi hoạt động).",
        "  • Hàm tìm Lớn nhất / Nhỏ nhất: =MAX(D4:D10) và =MIN(D4:D10).",
        "  • Hàm đếm: =COUNTA(B4:B10) (Đếm xem có bao nhiêu khoản mục công việc đã nhập)."
    ])

    add_step(doc, "7.3", "Áp dụng hàm điều kiện IF đánh giá tiến độ", [
        "Cú pháp chuẩn: =IF(Điều_kiện, \"Kết_quả_nếu_Đúng\", \"Kết_quả_nếu_Sai\")",
        "Ví dụ kiểm tra tiến độ thanh toán:",
        "  Tại ô F4 gõ: =IF(E4=\"Đã chi\", \"Hoàn thành\", \"Chờ duyệt\")",
        "Kéo công thức xuống: Rê chuột vào góc dưới bên phải ô F4 sao cho con trỏ biến thành dấu cộng đen nhỏ (+) > Nhấn giữ chuột kéo xuống dưới (hoặc nhấp đúp chuột trái vào dấu cộng đen)."
    ])

    add_image(doc, os.path.join(IMG_DIR, "vol2_img6_sheets_smart_suggestions.png"), "Tính năng Smart Suggestions trên Google Sheets tự động gợi ý công thức =SUM(...) khi gõ dấu =.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Bảng tính tự động tính toán tổng số liệu nhanh chóng; học viên biết tận dụng phím Tab để nhận gợi ý công thức thông minh của Google.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI DẤU PHÂN CÁCH TRONG CÔNG THỨC TRÊN GOOGLE SHEETS", [
        "Khi máy tính cài Vùng Việt Nam, một số tài khoản Google Sheets quy định dấu phân cách giữa các đối số trong hàm là DẤU CHẤM PHẨY (;).",
        "Nếu gõ =IF(A1>5, \"Đạt\", \"Trượt\") bị báo lỗi #ERROR!, bạn hãy đổi dấu phẩy thành dấu chấm phẩy: =IF(A1>5; \"Đạt\"; \"Trượt\") là sẽ chạy ngon lành."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Tạo bảng thống kê chấm điểm thi đua cán bộ gồm: Họ tên, Điểm chuyên cần, Điểm hoàn thành nhiệm vụ, Điểm tổng cộng.")
    add_p(doc, "1. Dùng hàm =SUM(...) tính Tổng điểm cho 5 cán bộ.")
    add_p(doc, "2. Dùng hàm =AVERAGE(...) tính điểm bình quân toàn phòng.")
    add_p(doc, "3. Thêm cột 'Xếp loại': Dùng hàm =IF(...) sao cho nếu Tổng điểm >= 80 thì ghi 'Khen thưởng', ngược lại ghi 'Đạt yêu cầu'.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Sử dụng thành thạo phím Tab để nhận công thức gợi ý tự động", "Tăng tốc độ nhập công thức gấp 3 lần"),
        ("Áp dụng chính xác các hàm SUM, AVERAGE, COUNTA", "Số liệu thống kê tự động cập nhật khi đổi dữ liệu"),
        ("Viết đúng hàm điều kiện IF phân loại kết quả", "Kết quả trả về đúng theo từng điều kiện logic")
    ])

    # BÀI 8
    add_heading_2(doc, "Bài 8: Tạo Bộ lọc xem riêng (Filter Views) và Khóa bảo vệ dải ô dữ liệu")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Hiểu được vấn đề 'Xung đột khi nhiều người cùng lọc dữ liệu' trên bảng tính online.")
    add_p(doc, "• Biết cách tạo Chế độ xem bộ lọc riêng (Filter Views): Bạn lọc theo ý bạn, đồng nghiệp lọc theo ý đồng nghiệp mà không hề làm rối màn hình của nhau!")
    add_p(doc, "• Biết cách Khóa bảo vệ ô dữ liệu (Protect sheets and ranges) để chỉ mình bạn được sửa công thức, người khác chỉ được nhập liệu.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Bảng tính Google Sheets có nhiều người cùng tham gia chỉnh sửa.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_callout(doc, "VẤN ĐỀ 'KINH HOÀNG' KHI BẬT FILTER THÔNG THƯỜNG TRÊN GOOGLE SHEETS", [
        "Nếu bạn bấm nút Filter thông thường (phễu) và lọc ra 'Chỉ xem hồ sơ của Nguyễn Văn An', thì NGAY LẬP TỨC trên màn hình máy tính của tất cả các đồng nghiệp và Sếp đang mở bảng tính đó cũng bị giật nảy và chỉ còn thấy hồ sơ của An!",
        "Điều này gây phiền toái cực lớn khi 5-10 người cùng làm việc một lúc.",
        "GIẢI PHÁP ĐỈNH CAO: BẮT BUỘC PHẢI DÙNG 'CHẾ ĐỘ XEM BỘ LỌC' (FILTER VIEWS) ĐỂ LỌC RIÊNG CHO MỘT MÌNH BẠN XEM!"
    ], callout_type="warning")

    add_step(doc, "8.1", "Tạo Bộ lọc xem riêng (Filter Views) - Lọc không làm phiền người khác", [
        "Bôi đen bảng dữ liệu cần lọc (kể cả dòng tiêu đề).",
        "Vào menu 'Dữ liệu' (Data) trên thanh công cụ.",
        "Rê chuột vào mục 'Chế độ xem bộ lọc' (Filter views) > Nhấp chọn 'Tạo chế độ xem bộ lọc mới' (Create new filter view).",
        "Quan sát màn hình: Viền bảng tính sẽ đổi sang MÀU ĐEN SẪM với thanh thông báo màu xám ở trên cùng. Đây là dấu hiệu bạn đang ở trong không gian lọc riêng của mình!",
        "Đặt tên cho bộ lọc: Nhấp vào ô chữ 'Tên:' trên thanh xám > Gõ tên bạn muốn (ví dụ: 'Hồ sơ của An' hoặc 'Việc cần làm gấp').",
        "Thực hiện lọc dữ liệu như bình thường. Toàn bộ thao tác lọc này CHỈ DUY NHẤT MÌNH BẠN NHÌN THẤY, màn hình của sếp và đồng nghiệp vẫn giữ nguyên trạng thái đầy đủ 100%!",
        "Thoát chế độ xem riêng: Bấm vào dấu nhân (X) ở góc trên bên phải thanh màu xám."
    ])

    add_step(doc, "8.2", "Khóa bảo vệ dải ô công thức (Protect ranges) - Tránh bị người khác xóa nhầm", [
        "Để tránh việc giao bảng tính cho người khác nhập số liệu mà họ lại vô tình nhấp vào cột Thành tiền/Tổng cộng rồi xóa mất công thức:",
        "1. Dùng chuột bôi đen dải ô chứa công thức cần bảo vệ (ví dụ cột F và dòng Tổng cộng).",
        "2. Nhấp chuột phải vào vùng bôi đen > Rê chuột xuống chọn 'Xem các hành động khác đối với ô' > Chọn 'Bảo vệ dải ô' (Protect range) - hoặc vào menu Dữ liệu > Trang tính và dải ô được bảo vệ.",
        "3. Bảng bên phải hiện ra, nhấp nút 'Đặt quyền' (Set permissions).",
        "4. Chọn chế độ: 'Chỉ bạn' (Only you) mới có quyền chỉnh sửa dải ô này; hoặc chọn 'Hiển thị cảnh báo khi chỉnh sửa dải ô này'.",
        "5. Bấm nút 'Đã xong' (Done). Bây giờ, đồng nghiệp có thể thoải mái nhập tên và số lượng ở cột ngoài, nhưng khi họ chạm vào ô công thức thì hệ thống sẽ khóa lại, không cho phép xóa!"
    ])

    add_image(doc, os.path.join(IMG_DIR, "vol2_img7_sheets_filter_view_protect.png"), "Giao diện viền đen sẫm của Chế độ xem bộ lọc (Filter view) và thanh Khóa bảo vệ dải ô (Protect ranges).")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Bảng tính trực tuyến được quản trị an toàn, chuyên nghiệp: mọi người cùng truy cập nhưng lọc dữ liệu độc lập không gây xung đột, công thức tính toán được khóa bảo vệ an toàn tuyệt đối.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI QUÊN BẤM NÚT X ĐỂ THOÁT FILTER VIEW", [
        "Nhiều người sau khi lọc xong thấy màn hình có viền đen thì tưởng bảng tính bị lỗi giao diện.",
        "Khắc phục: Chỉ cần nhấp chuột vào dấu 'X' nhỏ ở góc trên bên phải của thanh xám, bảng tính sẽ quay về giao diện bình thường."
    ], callout_type="tip")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở một bảng dữ liệu danh sách cán bộ hoặc hồ sơ công việc:")
    add_p(doc, "1. Tạo một Chế độ xem bộ lọc mới (Data > Filter views > Create new filter view). Đặt tên là 'Loc_Can_bo_Hoan_thanh'. Thực hiện lọc theo trạng thái Hoàn thành. Thoát ra bằng nút X.")
    add_p(doc, "2. Chọn cột Tổng cộng/Thành tiền, thực hiện thao tác Khóa bảo vệ dải ô (Protect range) và đặt quyền 'Chỉ bạn' có quyền sửa.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Tạo thành công Chế độ xem bộ lọc riêng (Filter View)", "Màn hình chuyển sang khung viền sẫm màu với thanh tiêu đề riêng"),
        ("Đặt tên rõ ràng cho từng chế độ xem bộ lọc", "Có thể mở lại bộ lọc nhanh chóng bất cứ lúc nào"),
        ("Khóa bảo vệ thành công dải ô công thức", "Tài khoản khác không thể tự ý sửa đổi công thức đã khóa")
    ])

    # BÀI 9
    add_heading_2(doc, "Bài 9: Dự án thực hành Sheets - Xây dựng Bảng phân công và theo dõi tiến độ công việc")
    
    add_p(doc, "Dự án này là bài tập lớn kết thúc Tập II, giúp bạn thiết kế một 'Hệ thống quản lý công việc số' thực thụ dành cho phòng ban của mình trên Google Sheets, có phân công nhiệm vụ, tự động theo dõi thời hạn và bảo vệ dữ liệu.")

    add_heading_3(doc, "Đặc tả yêu cầu của Dự án Quản lý tiến độ phòng ban")
    add_callout(doc, "CẤU TRÚC HỆ THỐNG THEO DÕI CÔNG VIỆC TRỰC TUYẾN", [
        "1. Tên bảng tính: 'BẢNG THEO DÕI TIẾN ĐỘ THỰC HIỆN NHIỆM VỤ PHÒNG BAN NĂM 2026'.",
        "2. Cấu trúc 8 cột dữ liệu:",
        "   • Cột A: STT",
        "   • Cột B: Nhiệm vụ / Công việc cụ thể",
        "   • Cột C: Người chủ trì (Cán bộ phụ trách)",
        "   • Cột D: Ngày giao việc (dd/mm/yyyy)",
        "   • Cột E: Thời hạn hoàn thành (Deadline - dd/mm/yyyy)",
        "   • Cột F: Ngày hoàn thành thực tế",
        "   • Cột G: Tình trạng (Đang làm / Đã xong / Trễ hạn)",
        "   • Cột H: Ghi chú / Đánh giá của Lãnh đạo",
        "3. Yêu cầu kỹ thuật quản trị:",
        "   • Dùng tính năng Hộp kiểm (Checkbox) hoặc Danh sách chọn thả xuống (Data Validation) cho cột G: Chọn nhanh giữa các trạng thái 'Đang làm', 'Đã xong', 'Trễ hạn'.",
        "   • Định dạng có điều kiện (Conditional Formatting): Tự động tô màu Đỏ nhạt cho dòng có chữ 'Trễ hạn', màu Xanh lá nhạt cho chữ 'Đã xong'.",
        "   • Dùng hàm COUNTIF đếm xem mỗi cán bộ đang nhận bao nhiêu đầu việc và có bao nhiêu việc đã hoàn thành.",
        "   • Khóa dải ô chứa công thức thống kê ở đáy bảng (Protect range).",
        "   • Tạo sẵn 2 Filter Views: Một bộ lọc cho 'Nhiệm vụ đang làm' và một bộ lọc riêng cho từng cán bộ."
    ], callout_type="note")

    add_heading_3(doc, "Tiêu chí đánh giá bài thực hành tổng hợp Google Sheets")
    add_checklist_table(doc, [
        ("Bảng tính thiết lập đúng Vùng Việt Nam, ngày tháng và tiền tệ chuẩn mực", "Dữ liệu hiển thị đồng bộ, không bị lỗi căn lề"),
        ("Tự động tô màu trực quan các trạng thái công việc (Conditional Formatting)", "Dễ dàng theo dõi việc gấp, việc trễ hạn"),
        ("Sử dụng thành thạo Filter Views để nhiều người cùng làm việc không xung đột", "Phát huy tối đa sức mạnh cộng tác trực tuyến"),
        ("Khóa bảo vệ thành công các vùng dữ liệu và công thức trọng yếu", "Bảo đảm an toàn tuyệt đối cho số liệu của cơ quan")
    ])
