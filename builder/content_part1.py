# -*- coding: utf-8 -*-
"""
Module biên soạn PHẦN 1: KỸ NĂNG MÁY TÍNH NỀN TẢNG & WINDOWS VĂN PHÒNG
Dành cho Tập I.
"""

from docx_helper import (
    add_heading_1, add_heading_2, add_heading_3, add_p,
    add_step, add_callout, add_screenshot_placeholder,
    add_styled_table, add_checklist_table
)

def build_part_1(doc):
    # PHẦN 1
    add_heading_1(doc, "PHẦN 1: KỸ NĂNG MÁY TÍNH NỀN TẢNG & WINDOWS VĂN PHÒNG")
    add_p(doc, "Phần mở đầu này trang bị toàn bộ những thao tác cơ bản nhất khi tiếp xúc với máy tính. Người học sẽ nắm vững cách điều khiển chuột, gõ bàn phím chuẩn tiếng Việt, tổ chức cây thư mục lưu trữ tài liệu ngăn nắp, và các kỹ năng mạng thiết yếu (duyệt web, tải tài liệu, gửi email công sở an toàn).")

    # BÀI 1
    add_heading_2(doc, "Bài 1: Làm quen phần cứng, thao tác điều khiển cốt lõi và quản lý cửa sổ Windows")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Phân biệt và thực hiện thành thạo 4 thao tác chuột: nhấp chuột trái, nhấp chuột phải, nhấp đúp và kéo thả.")
    add_p(doc, "• Nắm vững các phím chức năng quan trọng nhất trên bàn phím phục vụ soạn thảo và điều khiển.")
    add_p(doc, "• Biết cách bật/tắt máy tính đúng quy trình kỹ thuật, mở ứng dụng và chuyển đổi qua lại giữa các cửa sổ làm việc nhanh chóng.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Một máy tính để bàn hoặc máy tính xách tay (Laptop) cài hệ điều hành Windows (Windows 10 hoặc Windows 11).")
    add_p(doc, "• Chuột máy tính (khuyến khích người mới bắt đầu dùng chuột rời có con lăn thay vì Touchpad trên laptop để dễ thao tác).")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "1.1", "Thành thạo 4 thao tác điều khiển chuột cơ bản", [
        "Nhấp chuột trái (Left-click): Dùng ngón trỏ bấm dứt khoát 1 lần vào nút bên trái chuột. Tác dụng: Chọn một mục, đặt con trỏ soạn thảo tại vị trí muốn gõ, hoặc bấm vào các nút lệnh trên thanh công cụ.",
        "Nhấp đúp chuột trái (Double-click): Dùng ngón trỏ bấm 2 lần liên tiếp thật nhanh vào nút chuột trái. Tác dụng: Mở một tệp tin, mở một thư mục, hoặc khởi động một phần mềm từ màn hình chính.",
        "Nhấp chuột phải (Right-click): Dùng ngón giữa bấm 1 lần vào nút bên phải chuột. Tác dụng: Mở 'Bảng chọn ngữ cảnh' (Menu ngữ cảnh) chứa các lệnh hữu ích như Sao chép (Copy), Dán (Paste), Đổi tên (Rename), Xem thuộc tính (Properties).",
        "Kéo thả chuột (Drag and Drop): Đưa con trỏ chuột đến đối tượng, nhấn giữ chặt nút chuột trái, đồng thời di chuyển chuột đến vị trí mới rồi thả tay ra. Tác dụng: Di chuyển tệp tin vào thư mục, hoặc bôi đen chọn một đoạn văn bản."
    ])
    
    add_screenshot_placeholder(doc, "Hình minh họa cấu tạo chuột máy tính: Nút chuột trái, Con lăn ở giữa, Nút chuột phải và vị trí đặt các ngón tay chuẩn.")

    add_step(doc, "1.2", "Nhận biết các phím chức năng cốt lõi trên bàn phím", [
        "Phím Enter: Xuống dòng mới khi soạn thảo văn bản, hoặc tương đương lệnh Đồng ý/Xác nhận (OK).",
        "Phím Backspace (nằm phía trên phím Enter, có hình mũi tên sang trái): Xóa ký tự nằm ngay TRƯỚC (bên trái) con trỏ soạn thảo.",
        "Phím Delete (Del): Xóa ký tự nằm ngay SAU (bên phải) con trỏ soạn thảo, hoặc xóa tệp tin đang chọn.",
        "Phím Space (Thanh phím cách dài nhất ở đáy bàn phím): Tạo một khoảng trắng giữa hai từ. Quy tắc vàng: Giữa hai từ chỉ gõ đúng duy nhất 1 lần phím cách.",
        "Phím Shift: Nhấn giữ Shift đồng thời gõ một chữ cái để viết hoa chữ đó (ví dụ: giữ Shift + gõ 'a' được 'A'); hoặc gõ ký tự nằm ở hàng trên của phím số (như @, #, $, %, &).",
        "Phím Caps Lock: Bật/tắt chế độ viết hoa toàn bộ. Khi đèn Caps Lock sáng, mọi chữ cái gõ ra đều là chữ in hoa.",
        "Các phím điều hướng (Mũi tên Lên, Xuống, Trái, Phải): Di chuyển con trỏ soạn thảo từng ký tự hoặc từng dòng mà không làm mất văn bản."
    ])

    add_step(doc, "1.3", "Bật/Tắt máy tính đúng kỹ thuật và Quản lý cửa sổ làm việc", [
        "Bật máy: Nhấn nút nguồn (Power) trên thùng máy tính (PC) hoặc góc trên bàn phím (Laptop) 1 lần và chờ màn hình khởi động Windows.",
        "Tắt máy an toàn: Bấm chuột trái vào nút Start (hình cửa sổ Windows ở góc dưới cùng bên trái màn hình) > Nhấp vào biểu tượng Nguồn (Power) > Chọn Shut down. Tuyệt đối không bấm tắt công tắc ổ điện hoặc rút dây nguồn đột ngột vì sẽ gây hỏng ổ cứng và mất dữ liệu.",
        "Mở ứng dụng: Nhấp chuột trái vào ô Tìm kiếm (biểu tượng kính lúp cạnh nút Start), gõ tên ứng dụng muốn mở (ví dụ: gõ 'Word' hoặc 'Excel'), rồi nhấn phím Enter.",
        "Chuyển đổi nhanh giữa các cửa sổ ứng dụng: Nhấn giữ phím Alt, sau đó nhấn phím Tab (tổ hợp phím Alt + Tab) để thấy danh sách các ứng dụng đang mở, nhả phím để chọn ứng dụng mong muốn. Hoặc nhấp chuột vào biểu tượng ứng dụng tương ứng trên thanh tác vụ (Taskbar) ở đáy màn hình."
    ])

    add_screenshot_placeholder(doc, "Giao diện thanh tác vụ Taskbar của Windows, vị trí nút Start, ô tìm kiếm và cách bấm Shut down tắt máy chuẩn.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Người học tự tin cầm chuột, thao tác dứt khoát không bị run tay, điều khiển con trỏ chính xác vào các biểu tượng nhỏ; mở được đồng thời Word và Excel, chuyển qua lại mượt mà bằng phím Alt + Tab.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "CÁC LỖI KINH ĐIỂN CỦA NGƯỜI MỚI BẮT ĐẦU", [
        "Lỗi nhấp đúp chuột quá chậm: Nếu nhấp 2 lần cách nhau quá 1 giây, Windows sẽ hiểu là bạn muốn đổi tên tệp thay vì mở tệp. Cách sửa: Luyện tập bấm 2 tiếng 'cạch cạch' thật dứt khoát và liền nhau.",
        "Lỗi bị vô tình di chuyển chuột khi đang nhấp: Khi bấm nút chuột trái, ngón tay đè mạnh làm con chuột xê dịch khiến đối tượng bị kéo thả nhầm sang chỗ khác. Cách sửa: Giữ yên cổ tay và thân chuột khi bấm ngón trỏ.",
        "Lỗi viết hoa toàn bộ văn bản ngoài ý muốn: Do vô tình chạm vào phím Caps Lock. Cách sửa: Quan sát góc trên bên phải bàn phím, nếu thấy đèn phím Caps Lock đang sáng, hãy nhấn phím Caps Lock 1 lần nữa để tắt đi."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "1. Thực hành rê chuột vào 5 biểu tượng khác nhau trên màn hình Desktop, nhấp chuột trái để chọn, sau đó nhấp chuột phải để xem menu ngữ cảnh.")
    add_p(doc, "2. Mở ứng dụng Notepad (Sổ tay ghi chú đơn giản): Bấm Start > Gõ 'Notepad' > Enter. Luyện gõ họ và tên của bạn, số điện thoại, địa chỉ cơ quan, kết hợp sử dụng phím Shift để viết hoa chữ cái đầu và phím Backspace để xóa khi gõ nhầm.")
    add_p(doc, "3. Mở tiếp ứng dụng Máy tính (Calculator): Bấm Start > Gõ 'Calc' > Enter. Thực hành nhấn Alt + Tab để chuyển đổi qua lại giữa Notepad và Calculator 5 lần.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Cầm chuột đúng tư thế, nhấp đúp mở được tệp dứt khoát", "Tệp hoặc phần mềm mở ra ngay sau 2 lần nhấp"),
        ("Phân biệt và sử dụng đúng phím Backspace (xóa lùi) và Delete (xóa tới)", "Xóa chính xác ký tự mong muốn mà không xóa nhầm"),
        ("Sử dụng thành thạo tổ hợp phím Alt + Tab để chuyển cửa sổ", "Chuyển đổi qua lại giữa 2 phần mềm đang mở không cần dùng chuột"),
        ("Thực hiện tắt máy tính đúng trình tự qua nút Start > Shut down", "Máy tính lưu dữ liệu và tắt an toàn, không bị tắt đột ngột")
    ])

    # BÀI 2
    add_heading_2(doc, "Bài 2: Làm chủ bộ gõ tiếng Việt UniKey và nguyên tắc gõ văn bản chuẩn")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Hiểu bản chất và chọn đúng Bảng mã (Unicode) cùng Kiểu gõ (Telex) trên phần mềm UniKey.")
    add_p(doc, "• Nắm vững quy tắc gõ chữ cái có dấu và vị trí bỏ dấu thanh tiếng Việt.")
    add_p(doc, "• Tự khắc phục 100% các lỗi gõ tiếng Việt thường gặp: gõ bị mất dấu, gõ chữ hoa bị biến dạng, hoặc gõ trang web/mật khẩu bị dính chữ tiếng Việt.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Máy tính đã cài đặt phần mềm gõ tiếng Việt UniKey (hoặc EVKey).")
    add_p(doc, "• Nhận biết biểu tượng UniKey ở góc dưới cùng bên phải màn hình (khay hệ thống gần đồng hồ): Chữ V (màu đỏ) là đang bật chế độ gõ tiếng Việt; Chữ E (màu xanh) là đang ở chế độ gõ tiếng Anh.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "2.1", "Cài đặt thông số chuẩn trong bảng điều khiển UniKey", [
        "Tìm biểu tượng UniKey: Nhấp chuột trái vào dấu mũi tên hướng lên (^) ở góc dưới bên phải màn hình (khay hệ thống).",
        "Nếu thấy chữ 'E', nhấp chuột trái 1 lần vào đó để chuyển thành chữ 'V' (bật gõ tiếng Việt).",
        "Nhấp đúp chuột trái vào biểu tượng chữ V để mở Bảng điều khiển chính của UniKey.",
        "Tại mục 'Bảng mã': Nhấp chuột trái chọn đúng dòng 'Unicode' (Đây là quy định bắt buộc của Nhà nước theo Nghị định 30/2020/NĐ-CP và tiêu chuẩn quốc gia).",
        "Tại mục 'Kiểu gõ': Chọn 'Telex' (kiểu gõ phổ biến và nhanh nhất hiện nay).",
        "Bấm chuột trái vào nút 'Đóng' để lưu cài đặt (không bấm nút 'Kết thúc' vì sẽ tắt hẳn UniKey)."
    ])

    add_screenshot_placeholder(doc, "Bảng điều khiển UniKey với 2 thông số bắt buộc: Bảng mã là 'Unicode' và Kiểu gõ là 'Telex', khoanh đỏ nút Đóng.")

    add_step(doc, "2.2", "Học thuộc bảng quy tắc gõ tiếng Việt kiểu Telex", [
        "Cách tạo chữ cái có dấu tiếng Việt: Gõ lặp lại chữ cái đó hoặc thêm chữ cái tương ứng:",
        "  • aa = â   |   aw = ă   |   ee = ê   |   oo = ô   |   ow = ơ   |   uw = ư   |   dd = đ",
        "Cách bỏ dấu thanh tiếng Việt: Dùng các phím chữ cái đại diện cho dấu đặt ngay sau từ:",
        "  • Phím s = Dấu Sắc (ví dụ: cas = cá)",
        "  • Phím f = Dấu Huyền (ví dụ: caf = cà)",
        "  • Phím r = Dấu Hỏi (ví dụ: car = cả)",
        "  • Phím x = Dấu Ngã (ví dụ: cax = cã)",
        "  • Phím j = Dấu Nặng (ví dụ: caj = cạ)",
        "  • Phím z = Xóa dấu đã gõ (ví dụ: đang gõ 'cá', bấm z sẽ quay về 'ca')"
    ])

    add_step(doc, "2.3", "Quy tắc vàng đặt dấu câu trong văn bản chuẩn", [
        "Quy tắc 1: Luôn gõ hết toàn bộ các chữ cái trong từ rồi mới gõ phím dấu thanh ở cuối cùng (ví dụ gõ từ 'hoàng': gõ h-o-a-n-g rồi mới gõ f). Thói quen này giúp bộ gõ xử lý chính xác và không bị lỗi dấu.",
        "Quy tắc 2: Các dấu ngắt câu như dấu chấm (.), dấu phẩy (,), dấu hai chấm (:), dấu chấm phẩy (;), dấu chấm than (!), dấu hỏi (?) PHẢI ĐƯỢC GÕ DÍNH LIỀN vào ký tự đứng trước nó, sau đó BẮT BUỘC phải gõ một dấu cách (phím Space) rồi mới viết từ tiếp theo.",
        "Ví dụ ĐÚNG: 'Cộng hòa xã hội chủ nghĩa Việt Nam, Độc lập - Tự do - Hạnh phúc.'",
        "Ví dụ SAI: 'Cộng hòa xã hội chủ nghĩa Việt Nam ,Độc lập - Tự do - Hạnh phúc .' (bị thừa dấu cách trước dấu phẩy và thiếu dấu cách sau dấu phẩy)."
    ])

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Người học gõ văn bản tiếng Việt mượt mà, đúng dấu, không bị dính chữ, dấu câu đặt chuẩn mực theo quy tắc hành chính.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "XỬ LÝ SỰ CỐ GÕ TIẾNG VIỆT", [
        "Lỗi gõ chữ có dấu biến thành ký tự lạ (như: ddaay, ddang): Do UniKey đang ở chế độ tiếng Anh (biểu tượng chữ E). Bấm tổ hợp phím 'Ctrl + Shift' hoặc 'Alt + Z' (phím tắt chuyển nhanh) để đưa về chữ V.",
        "Lỗi gõ tiếng Việt bị nhảy dấu hoặc mất chữ khi gõ trên trình duyệt/thanh tìm kiếm: Nhấp chuột phải vào biểu tượng UniKey > chọn dòng 'Bật kiểm tra chính tả' để BỎ dấu tích ở mục này đi.",
        "Lỗi mở 2 phần mềm gõ tiếng Việt cùng lúc: Máy tính vừa chạy UniKey vừa chạy EVKey sẽ làm xung đột bàn phím, gõ một chữ ra hai chữ. Hãy kiểm tra khay hệ thống và tắt bớt một phần mềm."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "Mở ứng dụng Word hoặc Notepad, thực hành gõ chính xác đoạn văn bản sau (chú ý chữ hoa, dấu thanh và khoảng cách sau dấu câu):")
    add_p(doc, "'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\nNghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ quy định về công tác văn thư. Văn bản hành chính phải được trình bày đúng quy chuẩn, sử dụng phông chữ tiếng Việt Times New Roman, bộ mã ký tự Unicode dựng sẵn.'", italic=True)

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("UniKey được thiết lập đúng Bảng mã Unicode và Kiểu gõ Telex", "Biểu tượng chữ V hiển thị ở góc khay hệ thống"),
        ("Gõ được toàn bộ các nguyên âm tiếng Việt (ă, â, đ, ê, ô, ơ, ư)", "Chữ cái hiển thị đúng mẫu tự tiếng Việt chuẩn"),
        ("Dấu phẩy và dấu chấm dính sát từ phía trước, cách 1 khoảng trắng với từ phía sau", "Không có lỗi khoảng trắng dị thường quanh dấu câu")
    ])

    # BÀI 3
    add_heading_2(doc, "Bài 3: Quản lý thư mục, tệp tin khoa học và lưu trữ dữ liệu an toàn")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Hiểu rõ cấu trúc ổ đĩa máy tính và quy tắc 'sống còn' không lưu tài liệu quan trọng trên màn hình Desktop hoặc ổ C.")
    add_p(doc, "• Biết cách tự tạo cây thư mục công việc theo năm/tháng/phòng ban rõ ràng, khoa học.")
    add_p(doc, "• Thực hiện thành thạo các thao tác: Tạo mới, Đổi tên, Sao chép (Copy), Cắt (Cut), Dán (Paste) và Xóa tệp.")
    add_p(doc, "• Phân biệt rõ ràng giữa lệnh Lưu đè (Save) và Lưu thành tệp mới (Save As) để không bao giờ bị mất bài.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Mở công cụ quản lý tệp tin File Explorer trên Windows bằng cách: Nhấp chuột trái vào biểu tượng hình chiếc cặp vàng trên thanh Taskbar, hoặc nhấn tổ hợp phím tắt 'Windows + E'.")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "3.1", "Khám phá cấu trúc ổ đĩa và quy tắc bảo toàn dữ liệu", [
        "Ổ đĩa C (Local Disk C:): Là nơi chứa hệ điều hành Windows và các phần mềm cài đặt. Khu vực màn hình chính (Desktop) và thư mục Tải về (Downloads) cũng nằm trên ổ C. NGUYÊN TẮC BẢO MẬT: Không được lưu trữ tài liệu làm việc lâu dài trên Desktop hay ổ C, vì khi máy tính bị virus hoặc cài lại Windows, toàn bộ dữ liệu ở đây sẽ bị xóa sạch.",
        "Ổ đĩa D (Data D:) hoặc Ổ E: Là nơi an toàn để chứa tài liệu công việc cá nhân và cơ quan. Khi cài lại máy, dữ liệu trên ổ D vẫn được giữ nguyên vẹn 100%."
    ])

    add_callout(doc, "QUY TẮC ĐẶT TÊN TỆP VÀ THƯ MỤC CÔNG SỞ", [
        "Nên đặt tên tệp rõ ràng, có cấu trúc: [NămThángNgày]_[TênLoại]_[NộiDung] (Ví dụ: 20260905_CV_SoKHDT_XinKinhPhi.docx).",
        "Có thể viết tiếng Việt không dấu hoặc có dấu rõ ràng, dùng dấu gạch dưới (_) hoặc gạch nối (-) thay cho khoảng cách để tệp không bị lỗi khi gửi qua mạng.",
        "Tuyệt đối không đặt tên tệp chung chung như: 'van ban.docx', 'bai 1.docx', 'tai lieu moi.docx' vì chỉ sau 1 tuần bạn sẽ không nhớ bên trong chứa nội dung gì."
    ], callout_type="tip")

    add_step(doc, "3.2", "Tạo thư mục mới và Đổi tên thư mục", [
        "Nhấp đúp chuột trái vào biểu tượng ổ đĩa D: để mở ổ D.",
        "Tạo thư mục mới: Nhấp chuột phải vào một khoảng trống màu trắng bất kỳ trong thư mục > rê chuột chọn dòng 'New' > nhấp chuột trái chọn 'Folder'. Một thư mục mới hiện ra với tên mặc định 'New folder'.",
        "Gõ ngay tên thư mục mong muốn (ví dụ: 'TAI LIEU CONG TAC 2026'), sau đó nhấn phím Enter để hoàn tất.",
        "Đổi tên thư mục/tệp khi cần: Nhấp chuột phải vào thư mục/tệp cần đổi tên > chọn 'Rename' (hoặc chọn tệp rồi nhấn phím F2 trên bàn phím) > Gõ tên mới > Nhấn Enter."
    ])

    add_step(doc, "3.3", "Thao tác Sao chép (Copy), Cắt (Cut), Dán (Paste) và Xóa", [
        "Sao chép (Nhân đôi tệp - Bản gốc vẫn còn nguyên): Nhấp chuột phải vào tệp cần sao chép > Chọn 'Copy' (hoặc nhấn phím tắt Ctrl + C). Sau đó đi vào thư mục muốn lưu > Nhấp chuột phải vào khoảng trống > Chọn 'Paste' (hoặc nhấn phím tắt Ctrl + V).",
        "Di chuyển (Cắt tệp sang chỗ mới - Chỗ cũ không còn): Nhấp chuột phải vào tệp > Chọn 'Cut' (hoặc nhấn phím tắt Ctrl + X). Sau đó đi vào thư mục mới > Nhấp chuột phải > Chọn 'Paste' (Ctrl + V).",
        "Xóa tệp vào thùng rác: Nhấp chuột phải vào tệp > Chọn 'Delete' (hoặc bấm phím Delete trên bàn phím). Tệp sẽ được chuyển tạm thời vào Thùng rác (Recycle Bin) và có thể khôi phục lại nếu lỡ xóa nhầm.",
        "Khôi phục tệp bị xóa nhầm: Ra màn hình Desktop > Nhấp đúp mở 'Recycle Bin' > Nhấp chuột phải vào tệp cần lấy lại > Chọn 'Restore'."
    ])

    add_step(doc, "3.4", "Phân biệt cốt lõi: Save (Lưu cập nhật) và Save As (Lưu bản mới)", [
        "Lệnh Save (Phím tắt Ctrl + S): Dùng để lưu lại những thay đổi vừa gõ vào chính tệp tin đang mở. Ví dụ: Bạn đang soạn văn bản đã có tên, cứ sau 3-5 phút bạn bấm Ctrl + S một lần để phòng mất điện đột ngột.",
        "Lệnh Save As (Phím tắt F12 trong Office): Dùng khi bạn muốn nhân bản tài liệu thành một tệp tin mới với tên khác hoặc lưu sang thư mục khác, trong khi tệp cũ vẫn giữ nguyên trạng. Đây là cách tuyệt vời khi dùng một văn bản mẫu cũ để sửa thành văn bản mới mà không sợ làm hỏng mẫu gốc."
    ])

    add_screenshot_placeholder(doc, "Cửa sổ File Explorer: Cây thư mục bên trái (This PC, Ổ C, Ổ D), nút New Folder và thanh đường dẫn địa chỉ phía trên.")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Người học tạo được hệ thống thư mục 2-3 cấp ngăn nắp trong ổ D; biết copy/paste tệp tin thành thạo bằng cả chuột và phím tắt; không bao giờ bị mất tài liệu khi làm việc.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "CÁC NGUY CƠ MẤT TỆP TIN VÀ CÁCH PHÒNG TRÁNH", [
        "Lỗi lưu xong không biết tệp nằm ở đâu: Khi bấm Save trong Word/Excel, nhiều người chỉ nhìn thấy nút Save là bấm ngay mà không để ý phía trên đang lưu vào thư mục nào. Cách phòng tránh: Luôn nhìn lên thanh địa chỉ trên cùng của cửa sổ Save As để chọn đúng Ổ D > Thư mục công tác của mình trước khi bấm nút Save.",
        "Cách tìm lại tệp bị 'thất lạc': Mở File Explorer (Windows + E) > Nhấp chuột trái vào ô 'Search' ở góc trên cùng bên phải cửa sổ > Gõ một vài từ khóa có trong tên tệp hoặc nội dung tệp > Nhấn Enter để Windows tự động rà soát toàn bộ máy tính."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "1. Mở ổ đĩa D trên máy tính. Tạo một thư mục cha có tên: 'HOC_TAP_VAN_PHONG'.")
    add_p(doc, "2. Bên trong thư mục 'HOC_TAP_VAN_PHONG', tạo tiếp 3 thư mục con: '01_VAN_BAN_WORD', '02_BANG_TINH_EXCEL', '03_TAI_LIEU_THAM_KHAO'.")
    add_p(doc, "3. Mở phần mềm Notepad, gõ dòng chữ 'Đây là tài liệu thử nghiệm lưu trữ'. Bấm phím F12 (Save As), chọn đường dẫn lưu vào thư mục '03_TAI_LIEU_THAM_KHAO' với tên tệp 'Thu_nghiem.txt'.")
    add_p(doc, "4. Thực hành dùng lệnh Copy (Ctrl + C) và Paste (Ctrl + V) để nhân bản tệp 'Thu_nghiem.txt' sang thư mục '01_VAN_BAN_WORD'.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Thư mục được tạo đúng vị trí trong ổ D (không tạo ngoài Desktop)", "Đường dẫn thư mục bắt đầu bằng D:\\HOC_TAP_VAN_PHONG"),
        ("Tên thư mục và tệp tin rõ ràng, không chứa ký tự cấm", "Tên tệp không chứa ký tự: / \\ : * ? \" < > |"),
        ("Sử dụng thành thạo phím tắt Ctrl + C (sao chép) và Ctrl + V (dán)", "Tệp tin được nhân đôi sang thư mục đích thành công"),
        ("Nắm vững thao tác bấm Ctrl + S định kỳ khi soạn thảo", "Dữ liệu được cập nhật liên tục, không bị mất")
    ])

    # BÀI 4
    add_heading_2(doc, "Bài 4: Trình duyệt Web, Tải/Xuất tệp PDF và Giao tiếp Email công sở an toàn")
    
    add_heading_3(doc, "1. Mục tiêu bài học")
    add_p(doc, "Sau khi hoàn thành bài học này, bạn sẽ:")
    add_p(doc, "• Biết cách sử dụng trình duyệt Google Chrome hoặc Microsoft Edge để tra cứu thông tin và tải tài liệu về máy.")
    add_p(doc, "• Hiểu rõ định dạng tệp PDF là gì, cách xuất văn bản ra tệp PDF để gửi lãnh đạo duyệt hoặc ký số.")
    add_p(doc, "• Thành thạo quy trình gửi Email công vụ: Đặt tiêu đề chuẩn, viết nội dung lịch sự, đính kèm tệp văn bản chính xác.")
    add_p(doc, "• Nhận diện và phòng tránh các nguy cơ lừa đảo, mã độc qua email công sở.")

    add_heading_3(doc, "2. Điều kiện chuẩn bị")
    add_p(doc, "• Máy tính có kết nối mạng Internet.")
    add_p(doc, "• Đã có tài khoản Email (hòm thư điện tử cá nhân hoặc hòm thư công vụ cơ quan cấp).")

    add_heading_3(doc, "3. Thao tác từng bước (Cầm tay chỉ việc)")
    
    add_step(doc, "4.1", "Sử dụng trình duyệt web và Tải tài liệu về máy an toàn", [
        "Mở trình duyệt: Nhấp đúp chuột vào biểu tượng Google Chrome hoặc Microsoft Edge trên màn hình.",
        "Truy cập một trang web: Nhấp chuột vào thanh địa chỉ ở trên cùng của trình duyệt, gõ địa chỉ trang web (ví dụ: thuvienphapluat.vn hoặc google.com) rồi nhấn Enter.",
        "Tải một tệp văn bản từ mạng về: Khi tìm thấy liên kết tải tài liệu (thường có chữ 'Tải về' hoặc biểu tượng mũi tên chỉ xuống), nhấp chuột trái 1 lần vào đó.",
        "Tìm tệp vừa tải về: Sau khi tải xong, tệp sẽ nằm mặc định trong thư mục 'Downloads' của máy tính. Mở File Explorer > bấm chuột vào mục 'Downloads' ở cột bên trái để lấy tệp. Di chuyển (Cut/Paste) ngay tệp này sang Ổ D để bảo quản."
    ])

    add_step(doc, "4.2", "Định dạng PDF và Cách xuất tệp văn bản sang PDF", [
        "Định dạng PDF là gì?: PDF (Portable Document Format) là định dạng tài liệu số cố định. Khi xuất văn bản sang PDF, tài liệu sẽ không bị nhảy font, không bị lệch trang khi mang sang máy tính khác hoặc mở trên điện thoại. Đây là định dạng chuẩn bắt buộc dùng để gửi văn bản chính thức cho đối tác hoặc trình ký điện tử.",
        "Cách xuất tệp PDF từ Microsoft Word: Trong cửa sổ Word đang mở văn bản > Vào menu 'File' > Chọn 'Export' > Bấm nút 'Create PDF/XPS' > Chọn thư mục lưu trên Ổ D > Bấm nút 'Publish' (hoặc dùng menu File > Save As > ở mục Save as type chọn 'PDF')."
    ])

    add_screenshot_placeholder(doc, "Thao tác xuất tệp PDF trong Microsoft Word: Menu File > Export > Nút Create PDF/XPS Document.")

    add_step(doc, "4.3", "Quy chuẩn soạn thảo và gửi Email công sở", [
        "Đăng nhập hòm thư: Mở trình duyệt web > truy cập gmail.com (hoặc cổng email nội bộ của cơ quan) > Đăng nhập tên và mật khẩu.",
        "Bấm nút 'Soạn thư' (Compose) có dấu cộng lớn ở góc trên bên trái.",
        "Mục 'Đến' (To): Gõ chính xác địa chỉ email của người nhận (ví dụ: nguyenvanan@hanoi.gov.vn). Kiểm tra từng ký tự, chỉ cần sai một chữ cái thư sẽ không đến nơi.",
        "Mục 'Chủ đề / Tiêu đề' (Subject) - BẮT BUỘC PHẢI CÓ: Đặt tiêu đề rõ ràng, bao quát nội dung. Công thức chuẩn: [Tên cơ quan/bộ phận] - [Nội dung văn bản/công việc]. Ví dụ: 'Văn phòng UBND - Báo cáo tổng hợp số liệu quý III/2026'. Tuyệt đối không để trống tiêu đề.",
        "Phần thân thư: Phải có 3 phần chuẩn tắc:",
        "  1. Lời chào trang trọng: 'Kính gửi: Đồng chí Trưởng phòng...' hoặc 'Kính gửi: Ban Giám đốc...'",
        "  2. Nội dung vắn tắt: Trình bày ngắn gọn mục đích gửi thư (1-3 câu).",
        "  3. Lời kết và Chữ ký: 'Trân trọng cảm ơn! / Kính báo cáo!' kèm Họ tên, Chức vụ, Đơn vị công tác, Số điện thoại liên hệ.",
        "Đính kèm tệp (Attach files): Nhấp chuột trái vào biểu tượng chiếc kẹp ghim giấy ở hàng công cụ dưới đáy khung soạn thư > Cửa sổ hiện ra, duyệt đến Ổ D > Chọn tệp văn bản (Word hoặc PDF) cần gửi > Bấm 'Open'. Chờ thanh tiến trình tải tệp lên hoàn tất.",
        "Kiểm tra lại lần cuối và bấm nút 'Gửi' (Send)."
    ])

    add_callout(doc, "QUY TẮC AN TOÀN SỐ CÔNG SỞ: PHÒNG TRÁNH MÃ ĐỘC VÀ LỪA ĐẢO", [
        "Tuyệt đối không mở các email có tiêu đề trúng thưởng, đe dọa khóa tài khoản ngân hàng, hoặc yêu cầu cung cấp mật khẩu.",
        "Cảnh giác tối đa với các tệp đính kèm có đuôi: .exe, .scr, .bat, .vbs, hoặc tệp nén .zip không rõ người gửi. Khi bấm vào các tệp này, máy tính sẽ bị nhiễm virus tống tiền (ransomware) mã hóa toàn bộ dữ liệu.",
        "Chỉ mở tệp đính kèm khi bạn biết rõ người gửi là ai và đã được thông báo trước về nội dung công việc."
    ], callout_type="warning")

    add_heading_3(doc, "4. Kết quả mong đợi")
    add_p(doc, "Học viên tự tải được tài liệu mẫu từ internet, chuyển đổi văn bản sang định dạng PDF chuẩn mực và tự tin soạn một bức thư điện tử công vụ hoàn chỉnh, văn minh.")

    add_heading_3(doc, "5. Lỗi thường gặp và cách khắc phục")
    add_callout(doc, "CÁC LỖI EMAIL PHỔ BIẾN", [
        "Lỗi quên đính kèm tệp: Thư viết 'Kính gửi đồng chí tệp báo cáo đính kèm' nhưng bấm Gửi luôn mà quên bấm kẹp ghim. Cách xử lý: Nếu lỡ gửi nhầm, lập tức bấm nút 'Hoàn tác' (Undo) xuất hiện ở góc dưới màn hình trong vòng 10 giây; hoặc soạn ngay một thư ngắn xin lỗi và đính kèm lại tệp.",
        "Lỗi gửi tệp quá dung lượng cho phép: Email thông thường chỉ cho phép đính kèm tệp dưới 25 MB. Nếu gửi tài liệu hoặc ảnh quá nặng, Gmail sẽ tự động chuyển thành liên kết Google Drive."
    ], callout_type="warning")

    add_heading_3(doc, "6. Bài thực hành tự rèn luyện")
    add_p(doc, "1. Mở phần mềm Word, tạo một văn bản ngắn ghi 'Báo cáo công tác tuần'. Lưu tệp vào ổ D, sau đó xuất ra tệp PDF có tên '20260905_Bao_cao_tuan.pdf'.")
    add_p(doc, "2. Đăng nhập vào hòm thư điện tử cá nhân/công vụ của bạn. Soạn một email gửi cho một đồng nghiệp trong phòng (hoặc gửi cho chính địa chỉ email của bạn để kiểm tra):")
    add_p(doc, "   • Tiêu đề: '[Bộ phận Tổng hợp] - Gửi dự thảo Báo cáo công tác tuần'")
    add_p(doc, "   • Thân thư: Soạn đúng lời chào, nội dung giới thiệu tệp và chữ ký liên hệ.")
    add_p(doc, "   • Đính kèm: Chọn tệp PDF vừa xuất ở Bước 1.")
    add_p(doc, "   • Bấm Gửi và kiểm tra hộp thư 'Đã gửi' (Sent) để xác nhận thư đã đi thành công.")

    add_heading_3(doc, "7. Tiêu chí tự kiểm tra")
    add_checklist_table(doc, [
        ("Biết tìm và di chuyển tệp từ thư mục Downloads về Ổ D", "Tài liệu tải về được cất gọn trong ổ D an toàn"),
        ("Xuất thành công tệp văn bản từ Word sang định dạng PDF", "Tệp .pdf tạo ra mở được mượt mà, định dạng giữ nguyên"),
        ("Email có đầy đủ Tiêu đề, Lời chào, Thân thư và Chữ ký", "Thư mang tính chuyên nghiệp, văn minh công sở"),
        ("Đính kèm đúng tệp tài liệu trước khi bấm nút Gửi", "Tệp đính kèm hiển thị rõ dung lượng và tên tệp trong thư")
    ])
