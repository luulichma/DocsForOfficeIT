# -*- coding: utf-8 -*-
"""
Module biên soạn PHẦN 1 TẬP II:
SOẠN THẢO VÀ CỘNG TÁC VĂN BẢN TRỰC TUYẾN VỚI GOOGLE DOCS
"""

from docx_helper import (
    add_heading_1, add_heading_2, add_heading_3, add_p,
    add_step, add_callout, add_screenshot_placeholder,
    add_styled_table, add_checklist_table
)

def build_vol2_part_1(doc):
    add_heading_1(doc, "PHẦN 1: SOẠN THẢO VÀ CỘNG TÁC VĂN BẢN TRỰC TUYẾN VỚI GOOGLE DOCS")
    add_p(doc, "Google Docs (Google Tài liệu) là công cụ soạn thảo văn bản trực tuyến hàng đầu thế giới chạy trực tiếp trên trình duyệt web. Khác với Word truyền thống, Google Docs cho phép nhiều người cùng mở một văn bản và chỉnh sửa đồng thời trong cùng một giây, loại bỏ hoàn toàn việc phải gửi qua gửi lại các tệp đính kèm qua email hoặc Zalo. Phần này sẽ hướng dẫn bạn làm chủ môi trường làm việc số, áp dụng chuẩn thể thức văn bản hành chính Việt Nam trên nền tảng web và phối hợp làm việc nhóm trơn tru.")

    # BÀI 1
    add_heading_2(doc, "Bài 1: Khởi tạo Google Docs và Cơ chế lưu trữ đám mây tự động")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cách mở Google Docs từ trình duyệt web Chrome hoặc Edge bằng tài khoản Google (Gmail).")
    add_p(doc, "• Hiểu bản chất cơ chế Tự động lưu (Auto-save) lên Google Drive: Không bao giờ phải bấm Ctrl + S.")
    add_p(doc, "• Chuyển đổi hai chiều mượt mà: Tải tệp Word (.docx) từ máy tính lên Google Docs để sửa, và tải ngược tệp Docs về máy tính dưới dạng Word hoặc PDF.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Máy tính kết nối mạng Internet.")
    add_p(doc, "• Trình duyệt Google Chrome (khuyến nghị) và một tài khoản Gmail cá nhân hoặc hòm thư công vụ Google Workspace.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "1.1", "Khởi tạo một văn bản Google Docs mới", [
        "Mở trình duyệt Google Chrome.",
        "Cách 1 (Nhanh nhất thế giới): Nhấp chuột vào thanh địa chỉ trên cùng của trình duyệt, gõ chữ: 'docs.new' rồi nhấn Enter. Một trang soạn thảo mới toanh sẽ hiện ra ngay lập tức!",
        "Cách 2: Truy cập trang web 'docs.google.com' > Nhấp chuột trái vào ô có dấu cộng nhiều màu 'Tài liệu trống' (Blank document).",
        "Đổi tên tệp văn bản: Nhìn lên góc trên cùng bên trái, nhấp chuột trái vào ô chữ 'Tài liệu không có tiêu đề' > Gõ tên tài liệu (ví dụ: 'Du_thao_Ke_hoach_tuan_2026') > Nhấn Enter."
    ])

    add_step(doc, "1.2", "Hiểu cơ chế 'Tự động lưu' (Auto-save) lên đám mây", [
        "Trong Google Docs, BẠN KHÔNG CẦN VÀ CŨNG KHÔNG THỂ BẤM LỆNH SAVE!",
        "Ngay khi bạn gõ từng chữ cái, Google Docs sẽ tự động lưu lại tức thì lên bộ nhớ đám mây Google Drive.",
        "Quan sát trạng thái lưu: Nhìn cạnh tên tài liệu ở góc trên bên trái, bạn sẽ thấy biểu tượng hình đám mây nhỏ có dấu tích xanh v kèm dòng chữ: 'Đã lưu vào Drive' (Saved to Drive). Dù mất điện đột ngột hay mất mạng, toàn bộ từng từ bạn vừa gõ đều được bảo toàn nguyên vẹn 100%."
    ])

    add_step(doc, "1.3", "Tải tệp Word từ máy lên Docs và Xuất tệp Docs về máy", [
        "Đưa tệp Word từ máy tính lên Google Docs:",
        "  1. Truy cập 'drive.google.com'.",
        "  2. Bấm nút '+ Mới' (+ New) ở góc trên bên trái > Chọn 'Tải tệp lên' (File upload).",
        "  3. Chọn tệp Word (.docx) từ Ổ D trên máy tính > Bấm Open. Sau khi tải xong, nhấp đúp vào tệp để mở và chỉnh sửa trực tiếp trên Google Docs.",
        "Tải tệp từ Google Docs về máy tính:",
        "  1. Trong cửa sổ Google Docs, vào menu 'Tệp' (File) ở góc trên bên trái.",
        "  2. Rê chuột vào mục 'Tải xuống' (Download).",
        "  3. Muốn chỉnh sửa tiếp trên máy: Chọn 'Microsoft Word (.docx)'.",
        "  4. Muốn gửi ban hành/trình ký: Chọn 'Tài liệu PDF (.pdf)'."
    ])

    add_screenshot_placeholder(doc, "Giao diện Google Docs: Ô đổi tên tài liệu góc trên bên trái, biểu tượng đám mây Đã lưu vào Drive và menu Tệp > Tải xuống > Microsoft Word / PDF.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Học viên nắm vững cách tạo văn bản mới bằng lệnh 'docs.new', tự tin làm việc mà không sợ mất bài khi cúp điện, trao đổi tệp Word và Docs qua lại không gặp trở ngại.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI TÀI LIỆU BỊ MẤT KẾT NỐI MẠNG KHI ĐANG GÕ", [
        "Hiện tượng: Đang soạn thảo thì hiện thông báo 'Đang thử kết nối lại...' kèm đám mây màu xám gạch chéo.",
        "Nguyên nhân: Mạng Wifi hoặc dây mạng cơ quan bị chập chờn.",
        "Cách xử lý: BẠN VẪN CỨ TIẾP TỤC GÕ BÌNH THƯỜNG! Google Docs có bộ nhớ đệm trên trình duyệt. Ngay khi có mạng trở lại (dù 5 phút sau), toàn bộ các đoạn bạn vừa gõ sẽ tự động được đồng bộ lên Google Drive mà không mất một chữ nào."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "1. Mở trình duyệt, gõ 'docs.new' để tạo một tài liệu mới.")
    add_p(doc, "2. Đặt tên tài liệu là 'Thuc_hanh_Docs_Bai1'. Gõ họ tên, đơn vị công tác và quan sát biểu tượng đám mây tự lưu.")
    add_p(doc, "3. Vào Tệp > Tải xuống > Microsoft Word (.docx) để lưu tệp về thư mục 'D:\\HOC_TAP_VAN_PHONG\\01_VAN_BAN_WORD'.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Mở được Google Docs từ trình duyệt web thành thạo", "Trang tài liệu mới mở ra sẵn sàng soạn thảo"),
        ("Đổi tên tài liệu rõ ràng, không để 'Tài liệu không có tiêu đề'", "Tên tệp hiển thị đúng ở góc trên bên trái"),
        ("Nhận biết trạng thái đã lưu vào Drive qua biểu tượng đám mây", "Hiểu cơ chế không cần bấm nút Save"),
        ("Tải được tệp về máy tính dưới định dạng .docx và .pdf", "Tệp lưu thành công trong thư mục máy tính")
    ])

    # BÀI 2
    add_heading_2(doc, "Bài 2: Soạn thảo và Định dạng thể thức văn bản hành chính trên Google Docs")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Thiết lập khổ giấy A4 và căn lề trang chuẩn Nghị định 30/2020/NĐ-CP trên Google Docs.")
    add_p(doc, "• Cài đặt phông chữ Times New Roman, cỡ chữ, giãn dòng 1.3 và thụt lề đoạn văn.")
    add_p(doc, "• Sử dụng Bảng ẩn viền trên giao diện web để căn Quốc hiệu - Tiêu ngữ và Chữ ký chuẩn xác.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Mở tài liệu Google Docs mới.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "2.1", "Thiết lập Khổ giấy A4 và Căn lề trang trên Google Docs", [
        "Trên thanh menu, nhấp chuột trái vào menu 'Tệp' (File) > Chọn dòng cuối cùng: 'Thiết lập trang' (Page setup).",
        "Trong cửa sổ hiện ra:",
        "  • Mục Hướng giấy: Chọn 'Đứng' (Portrait).",
        "  • Mục Kích thước giấy: Nhấp chuột chọn đúng 'A4 (210 x 297 mm)'.",
        "  • Mục Lề (đơn vị cm hoặc inch tùy cài đặt ngôn ngữ Google):",
        "    + Trên (Top): 2 cm (hoặc 0.79 inch)",
        "    + Dưới (Bottom): 2 cm (hoặc 0.79 inch)",
        "    + Trái (Left): 3 cm (hoặc 1.18 inch) - chuẩn để đóng ghim hồ sơ.",
        "    + Phải (Right): 2 cm (hoặc 0.79 inch)",
        "Nhấp nút 'Đặt làm mặc định' (Set as default) nếu muốn các tài liệu sau tự áp dụng, sau đó bấm 'OK'."
    ])

    add_step(doc, "2.2", "Định dạng Phông chữ, Giãn dòng và Căn đều 2 bên", [
        "Chọn phông chữ: Nhấp vào ô phông chữ trên thanh công cụ (thường đang hiển thị Arial) > Chọn 'Times New Roman'.",
        "Chỉnh cỡ chữ: Chọn cỡ 13 hoặc 14.",
        "Giãn dòng: Nhấp vào biểu tượng 'Dãn cách dòng và đoạn' trên thanh công cụ > Chọn 'Dãn cách tùy chỉnh...' > Nhập giãn dòng '1.3' và Khoảng cách đoạn sau là '4' hoặc '6' pt > Bấm Áp dụng.",
        "Căn đều hai bên: Bấm tổ hợp phím 'Ctrl + Shift + J' (phím tắt căn đều Justify trên Google Docs)."
    ])

    add_step(doc, "2.3", "Bố trí Quốc hiệu - Tiêu ngữ bằng Bảng ẩn viền trên Web", [
        "Vào menu 'Chèn' (Insert) > 'Bảng' (Table) > Chọn lưới 2 cột 1 hàng (2x1).",
        "Ô trái: Gõ Tên cơ quan ban hành (in hoa đậm cỡ 12-13) và Số ký hiệu.",
        "Ô phải: Gõ Quốc hiệu (in hoa đậm cỡ 12-13), Tiêu ngữ (đậm cỡ 13-14), Địa danh ngày tháng (nghiêng).",
        "Ẩn viền bảng trên Google Docs:",
        "  1. Bôi đen cả 2 ô trong bảng.",
        "  2. Nhấp chuột phải > Chọn 'Thuộc tính bảng' (Table properties).",
        "  3. Khung bên phải hiện ra, nhấp vào mục 'Màu sắc' (Color) > Tại ô 'Đường viền bảng' (Table border) nhấp chọn độ dày '0 pt' (hoặc chọn màu Trắng).",
        "Toàn bộ đường viền đen biến mất hoàn toàn, phần đầu văn bản đẹp hoàn mỹ!"
    ])

    add_screenshot_placeholder(doc, "Cửa sổ Thiết lập trang (Page setup) trên Google Docs: Khổ giấy A4, lề Trái 3cm, Trên/Dưới/Phải 2cm và bảng thuộc tính Table properties ẩn viền 0pt.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Văn bản trên Google Docs có giao diện thanh lịch, chuẩn mực theo Nghị định 30/2020/NĐ-CP, không khác biệt so với soạn trên Microsoft Word chuyên nghiệp.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI KHÔNG TÌM THẤY PHÔNG TIMES NEW ROMAN TRÊN GOOGLE DOCS", [
        "Hiện tượng: Mở danh sách phông chữ nhưng cuộn mãi không thấy tên 'Times New Roman'.",
        "Cách khắc phục: Nhấp vào ô chọn phông chữ > Nhấp vào dòng đầu tiên: 'Phông chữ khác' (More fonts) có dấu cộng lớn > Gõ 'Times New Roman' vào ô tìm kiếm > Tích chọn vào nó và bấm OK. Phông sẽ được ghim vĩnh viễn vào danh sách của bạn."
    ], callout_type="tip")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở một tài liệu Google Docs mới. Thiết lập lề A4 chuẩn NĐ 30. Dùng bảng 2x1 ẩn viền để soạn phần đầu cho một văn bản giả lập của UBND Xã/Phường.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Thiết lập đúng khổ A4 và lề trang 4 phía trên Google Docs", "Lề trái 3cm, lề trên dưới 2cm chuẩn xác"),
        ("Sử dụng phông chữ Times New Roman cỡ 13-14pt", "Văn bản hiển thị đúng phông chữ hành chính"),
        ("Ẩn hoàn toàn viền bảng bằng thiết lập 0 pt", "Không còn thấy nét viền đen của bảng trên tài liệu")
    ])

    # BÀI 3
    add_heading_2(doc, "Bài 3: Làm việc nhóm, Phân quyền và Cộng tác thời gian thực trên Google Docs")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cách Chia sẻ tài liệu (Share) cho đồng nghiệp và phân quyền chính xác: Xem, Nhận xét hoặc Chỉnh sửa.")
    add_p(doc, "• Để lại bình luận (Comment) trực tiếp trên từng câu từ và tag tên đồng nghiệp (@email) để giao việc.")
    add_p(doc, "• Sử dụng Chế độ Đề xuất chỉnh sửa (Suggesting) tương đương tính năng Track Changes trong Word.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Địa chỉ email Google (Gmail) của một đồng nghiệp cùng phòng để thực hành chia sẻ.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "3.1", "Chia sẻ tài liệu và Phân quyền truy cập chính xác", [
        "Ở góc trên cùng bên phải màn hình Google Docs, nhấp chuột vào nút lớn màu xanh: 'Chia sẻ' (Share).",
        "Thêm người nhận: Gõ chính xác địa chỉ email của đồng nghiệp vào ô 'Thêm người và nhóm'.",
        "PHÂN QUYỀN TRUY CẬP (Rất quan trọng): Nhìn sang ô bên phải cạnh tên người nhận, nhấp chuột chọn 1 trong 3 mức quyền:",
        "  1. 'Người xem' (Viewer): Chỉ được đọc tài liệu, không thể sửa chữa hay bình luận gì (dùng khi gửi tài liệu chính thức để đối tác tham khảo).",
        "  2. 'Người nhận xét' (Commenter): Được đọc và được viết bình luận góp ý, nhưng không được tự ý sửa nội dung gốc (rất hay dùng khi gửi Lãnh đạo hoặc các phòng chuyên môn góp ý dự thảo).",
        "  3. 'Người chỉnh sửa' (Editor): Có toàn quyền gõ, sửa, xóa như chủ tài liệu (dùng khi phối hợp cùng 1 nhóm trực tiếp soạn thảo).",
        "Bấm nút 'Gửi' (Send). Đồng nghiệp sẽ nhận được email thông báo kèm đường liên kết mở tài liệu ngay lập tức."
    ])

    add_step(doc, "3.2", "Để lại Bình luận (Comment) và Tag tên giao việc (@email)", [
        "Cách tạo bình luận: Dùng chuột bôi đen cụm từ hoặc đoạn văn bản cần góp ý > Nhấp vào biểu tượng dấu cộng nhỏ 'Thêm nhận xét' xuất hiện ở mép phải trang giấy (hoặc bấm tổ hợp phím Ctrl + Alt + M).",
        "Nhập nội dung góp ý vào khung nhận xét.",
        "Mẹo Tag tên giao việc: Gõ ký tự '@' kèm email của đồng nghiệp (ví dụ: '@lan.nguyen@gmail.com Đề nghị đồng chí rà soát lại số liệu kinh phí mục này nhé').",
        "Bấm nút 'Nhận xét' (Comment). Hệ thống sẽ gửi ngay một thông báo riêng đến hộp thư của người được tag."
    ])

    add_step(doc, "3.3", "Sử dụng Chế độ Đề xuất chỉnh sửa (Suggesting)", [
        "Khi được giao nhiệm vụ duyệt bài nhưng không muốn xóa thẳng tay làm mất câu chữ ban đầu của người soạn:",
        "Nhìn lên góc trên bên phải thanh công cụ (dưới nút Chia sẻ), bạn sẽ thấy biểu tượng chiếc bút chì có chữ 'Chỉnh sửa' (Editing).",
        "Nhấp chuột vào đó và CHUYỂN SANG: 'Đề xuất' (Suggesting).",
        "Bây giờ, mọi thao tác gõ thêm chữ sẽ hiển thị bằng màu xanh lá cây; mọi từ bạn xóa sẽ bị một đường gạch ngang ở giữa chứ không biến mất.",
        "Người chủ trì văn bản khi mở bài ra chỉ cần bấm nút Dấu tích xanh (Chấp nhận đề xuất) hoặc Dấu X (Từ chối đề xuất) là xong!"
    ])

    add_screenshot_placeholder(doc, "Cửa sổ Chia sẻ phân quyền (Người xem, Nhận xét, Chỉnh sửa) và thanh công cụ chuyển đổi giữa chế độ Chỉnh sửa (Editing) và Đề xuất (Suggesting).")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Cả nhóm có thể cùng mở một bản dự thảo, cùng nhìn thấy con trỏ màu của nhau đang gõ trực tiếp trên màn hình, bình luận trao đổi rôm rả ngay trên văn bản mà không cần họp hành phức tạp.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "LỖI CHIA SẺ LIÊN KẾT BỊ KHÓA QUYỀN TRUY CẬP", [
        "Hiện tượng: Bạn copy link gửi qua Zalo cho đồng nghiệp, đồng nghiệp bấm vào thì bị báo 'Bạn cần quyền truy cập' kèm nút Yêu cầu quyền.",
        "Nguyên nhân: Mặc định Google khóa tài liệu ở chế độ 'Bị hạn chế' (Restricted).",
        "Cách khắc phục: Bấm nút 'Chia sẻ' > Nhìn xuống mục 'Quyền truy cập chung' > Đổi từ 'Bị hạn chế' thành: 'Bất kỳ ai có đường liên kết' (Anyone with the link) > Chọn quyền 'Người xem' hoặc 'Người chỉnh sửa' > Bấm 'Sao chép đường liên kết' rồi gửi lại."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "1. Tạo một tài liệu Google Docs có tiêu đề: 'Dự thảo Thông báo lịch trực tuần'.")
    add_p(doc, "2. Bấm Chia sẻ tài liệu cho địa chỉ email của một người bạn/đồng nghiệp với quyền 'Người nhận xét'.")
    add_p(doc, "3. Bôi đen một dòng chữ, bấm phím Ctrl + Alt + M để để lại một lời bình luận.")
    add_p(doc, "4. Chuyển chế độ sang 'Đề xuất' (Suggesting) và sửa thử 2 từ trong văn bản để xem các vết sửa màu xanh.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Nắm vững sự khác biệt giữa 3 quyền: Xem, Nhận xét, Chỉnh sửa", "Phân đúng quyền cho đúng đối tượng"),
        ("Tạo được bình luận Comment bằng phím tắt Ctrl + Alt + M", "Khung bình luận hiển thị gắn chặt vào đoạn chữ"),
        ("Sử dụng thành thạo chế độ Đề xuất chỉnh sửa (Suggesting)", "Các câu chữ thêm bớt hiển thị rõ dấu vết duyệt bài")
    ])

    # BÀI 4
    add_heading_2(doc, "Bài 4: Lịch sử phiên bản và Khôi phục tài liệu khi bị sửa nhầm (Version History)")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cách mở tính năng Lịch sử phiên bản (Version History) để xem lại toàn bộ quá trình sửa đổi của tài liệu từ lúc mới tạo.")
    add_p(doc, "• Biết chính xác ai đã sửa nội dung gì, vào ngày giờ nào (mỗi người một màu sắc riêng).")
    add_p(doc, "• Tự tin khôi phục lại phiên bản cũ nguyên vẹn trong 1 giây khi đồng nghiệp vô tình xóa hỏng tài liệu.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Sử dụng tài liệu đã có nhiều người cùng chỉnh sửa hoặc đã qua vài ngày soạn thảo.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "4.1", "Mở cửa sổ Lịch sử phiên bản (Version History)", [
        "Vào menu 'Tệp' (File) trên thanh công cụ.",
        "Rê chuột vào mục 'Lịch sử phiên bản' (Version history).",
        "Nhấp chuột chọn 'Xem lịch sử phiên bản' (See version history) - Phím tắt: Ctrl + Alt + Shift + H.",
        "(Hoặc nhấp chuột thẳng vào dòng chữ nhỏ cạnh biểu tượng đám mây: 'Lần chỉnh sửa gần đây nhất là...' ở đầu trang)."
    ])

    add_step(doc, "4.2", "Kiểm tra ai đã sửa nội dung gì và Đặt tên phiên bản quan trọng", [
        "Quan sát cột danh sách bên phải màn hình: Toàn bộ các mốc thời gian sửa đổi (ví dụ: Hôm nay 09:30, Hôm qua 15:20) được liệt kê chi tiết.",
        "Mỗi người cộng tác sẽ được Google gán một MÀU SẮC RIÊNG (ví dụ: Bạn màu xanh lá, đồng nghiệp màu tím).",
        "Khi nhấp chuột vào một mốc thời gian, toàn bộ các chữ mà người đó vừa thêm hoặc xóa sẽ được tô màu tương ứng trên màn hình văn bản.",
        "Đặt tên cho các mốc quan trọng (Ví dụ: 'Bản nộp Giám đốc duyệt lần 1'): Nhấp chuột vào dấu 3 chấm cạnh mốc thời gian đó > Chọn 'Đặt tên cho phiên bản này' (Name this version)."
    ])

    add_step(doc, "4.3", "Khôi phục lại phiên bản cũ (Lấy lại bài nguyên vẹn)", [
        "Khi phát hiện ai đó lỡ tay xóa mất một phần nội dung quan trọng:",
        "1. Mở Lịch sử phiên bản.",
        "2. Nhấp chuột chọn mốc thời gian trước thời điểm bị xóa (nơi văn bản vẫn còn đầy đủ đẹp đẽ).",
        "3. Nhìn lên góc trên cùng bên trái màn hình, nhấp chuột vào nút lớn màu xanh: 'Khôi phục phiên bản này' (Restore this version).",
        "4. Bấm 'Khôi phục' để xác nhận. Ngay lập tức, văn bản sẽ quay ngược thời gian trở về trạng thái hoàn hảo như cũ!"
    ])

    add_screenshot_placeholder(doc, "Cửa sổ Lịch sử phiên bản (Version history): Cột mốc thời gian bên phải, mã màu của từng người và nút lớn Khôi phục phiên bản này.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Người dùng hoàn toàn xóa bỏ nỗi sợ bị người khác phá hỏng bài; nắm trong tay 'cỗ máy thời gian' để truy vết công việc và cứu dữ liệu trong nháy mắt.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "NỖI SỢ 'KHÔI PHỤC BẢN CŨ SẼ MẤT LUÔN CÁC ĐOẠN VỪA GÕ'", [
        "Nhiều người e ngại khi bấm 'Khôi phục phiên bản này' thì những đoạn vừa mới gõ hôm nay sẽ bị biến mất vĩnh viễn.",
        "SỰ THẬT: Google Docs lưu trữ tất cả! Khi bạn khôi phục bản cũ, bản hiện tại vẫn được lưu lại như một mốc lịch sử. Bất cứ lúc nào bạn cũng có thể mở lại Lịch sử phiên bản để quay lại thời điểm mới nhất."
    ], callout_type="tip")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "1. Mở một tài liệu Google Docs, gõ đoạn văn bản thứ nhất. Đặt tên phiên bản là 'Mốc 1 - Khởi tạo'.")
    add_p(doc, "2. Gõ tiếp đoạn văn bản thứ hai, sau đó bôi đen xóa sạch đoạn thứ nhất đi.")
    add_p(doc, "3. Mở Lịch sử phiên bản (Ctrl + Alt + Shift + H), chọn mốc 'Mốc 1 - Khởi tạo' và bấm nút 'Khôi phục phiên bản này' để lấy lại toàn bộ đoạn văn bản đã bị xóa.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Mở được Lịch sử phiên bản bằng menu Tệp hoặc phím tắt", "Cột lịch sử thời gian hiện ra bên phải màn hình"),
        ("Nhận diện được ai đã chỉnh sửa dựa vào mã màu sắc", "Phân biệt rõ ràng người sửa và nội dung sửa"),
        ("Thực hiện khôi phục thành công văn bản về trạng thái cũ", "Nội dung cũ được phục hồi nguyên vẹn 100%")
    ])

    # BÀI 5
    add_heading_2(doc, "Bài 5: Dự án thực hành Docs - Phối hợp soạn thảo và góp ý Dự thảo Kế hoạch tuần")
    
    add_p(doc, "Bài học này mô phỏng quy trình làm việc thực tế tại một cơ quan: Lãnh đạo giao cho nhóm 2-3 chuyên viên cùng nhau soạn thảo một bản Kế hoạch công tác tuần trên Google Docs, sau đó phân công nhận xét và hoàn thiện bản cuối.")

    add_heading_3(doc, "Tình huống thực hành nhóm giả lập")
    add_callout(doc, "NỘI DUNG DỰ THẢO KẾ HOẠCH CÔNG TÁC TUẦN", [
        "1. Thiết lập trang: Khổ A4, lề chuẩn NĐ 30, phông Times New Roman 13pt.",
        "2. Bảng ẩn viền phần đầu: VĂN PHÒNG SỞ TÀI CHÍNH / KẾ HOẠCH CÔNG TÁC TUẦN 38 (Từ ngày 14/9 đến ngày 20/9/2026).",
        "3. Phân công thực hành nhóm (hoặc tự đóng 2 vai):",
        "   • Người thứ nhất: Đóng vai Người soạn thảo chính - gõ khung kế hoạch từ Thứ Hai đến Thứ Sáu (Lịch họp giao ban, Lịch tiếp công dân, Lịch kiểm tra cơ sở...). Sau đó bấm nút Chia sẻ với quyền 'Người nhận xét' cho đồng nghiệp.",
        "   • Người thứ hai (hoặc mở cửa sổ ẩn danh): Đóng vai Lãnh đạo phòng duyệt bài - dùng tính năng Bình luận (Comment) để yêu cầu bổ sung người chủ trì cuộc họp ngày Thứ Ba; chuyển sang chế độ Đề xuất (Suggesting) để chỉnh sửa lại địa điểm họp ngày Thứ Năm từ Hội trường A sang Hội trường B.",
        "   • Người thứ nhất mở lại bài: Đọc các bình luận, phản hồi lại 'Đã tiếp thu', bấm nút Chấp nhận đề xuất (Dấu tích xanh), và vào menu Tệp > Tải xuống > Xuất tệp PDF hoàn chỉnh để trình ký."
    ], callout_type="note")

    add_heading_3(doc, "Tiêu chí đánh giá bài thực hành tổng hợp Google Docs")
    add_checklist_table(doc, [
        ("Văn bản thiết lập đúng lề A4 chuẩn Nghị định 30 trên web", "Giao diện cân đối, trang nhã"),
        ("Sử dụng thành thạo bảng ẩn viền để căn phần đầu văn bản", "Không bị lệch khi mở trên các trình duyệt khác nhau"),
        ("Thực hiện đầy đủ quy trình Chia sẻ, Nhận xét và Đề xuất duyệt bài", "Thể hiện rõ tinh thần làm việc cộng tác số"),
        ("Xuất bản thành công tệp PDF sạch đẹp để phát hành", "Tệp PDF tải về máy hoàn chỉnh, không lỗi font")
    ])
