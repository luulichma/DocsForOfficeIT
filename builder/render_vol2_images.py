# -*- coding: utf-8 -*-
"""
Kịch bản tạo sinh 7 ảnh chụp màn hình minh họa thực tế cho TẬP II
(Google Docs & Google Sheets) bằng công nghệ Playwright headless.
Các ảnh có độ phân giải cao (Retina 2x), giao diện tiếng Việt chuẩn Google Workspace,
kèm chú thích đóng khung đỏ (Annotations) và số thứ tự chỉ dẫn sư phạm.
"""

import os
import sys

# Đảm bảo in tiếng Việt ra console không lỗi
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from playwright.sync_api import sync_playwright

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "images")
os.makedirs(OUTPUT_DIR, exist_ok=True)

COMMON_CSS = """
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
  body { background: #f8f9fa; padding: 20px; display: flex; justify-content: center; align-items: center; }
  .window { background: #ffffff; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); border: 1px solid #dadce0; overflow: hidden; }
  .browser-bar { background: #edf2f7; padding: 8px 16px; display: flex; align-items: center; gap: 8px; border-bottom: 1px solid #e2e8f0; }
  .dot { width: 10px; height: 10px; border-radius: 50%; }
  .dot-red { background: #ef4444; } .dot-yellow { background: #f59e0b; } .dot-green { background: #10b981; }
  .url-bar { background: #ffffff; border-radius: 20px; padding: 4px 16px; font-size: 12px; color: #64748b; flex: 1; margin-left: 12px; border: 1px solid #cbd5e1; }
  
  /* Callout annotations */
  .callout-box { position: absolute; border: 2px solid #dc2626; border-radius: 4px; background: rgba(220, 38, 38, 0.08); pointer-events: none; }
  .step-badge { position: absolute; top: -12px; left: -12px; background: #dc2626; color: #ffffff; font-weight: bold; font-size: 12px; width: 22px; height: 22px; border-radius: 50%; display: flex; align-items: center; justify-content: center; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
  .arrow-tag { position: absolute; background: #dc2626; color: white; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; white-space: nowrap; }
</style>
"""

def generate_all_vol2_images():
    print(">>> Bắt đầu tạo sinh bộ 7 ảnh chụp thực tế cho Tập II...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, channel='chrome')
        context = browser.new_context(device_scale_factor=2)
        page = context.new_page()

        # --------------------------------------------------------------------
        # ẢNH 1: Google Docs - Đổi tên, Tự lưu Drive và Tải xuống Word/PDF
        # --------------------------------------------------------------------
        print("  -> Đang tạo vol2_img1_docs_autosave.png...")
        html_1 = f"""
        {COMMON_CSS}
        <div class="window" style="width: 850px;">
          <div class="browser-bar">
            <div class="dot dot-red"></div><div class="dot dot-yellow"></div><div class="dot dot-green"></div>
            <div class="url-bar">https://docs.google.com/document/d/1A2b3C.../edit</div>
          </div>
          <!-- Google Docs Header -->
          <div style="padding: 12px 20px; border-bottom: 1px solid #dadce0; display: flex; align-items: flex-start; gap: 16px; position: relative;">
            <!-- Docs Icon -->
            <div style="background: #1a73e8; color: white; font-weight: bold; font-size: 20px; width: 38px; height: 38px; border-radius: 4px; display: flex; align-items: center; justify-content: center; margin-top: 4px;">≡</div>
            
            <div style="flex: 1;">
              <!-- Title & Cloud status -->
              <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 6px; position: relative;">
                <div style="font-size: 17px; font-weight: 500; color: #202124; padding: 2px 8px; border-radius: 4px; border: 1px dashed #1a73e8; background: #e8f0fe;">
                  Du_thao_Ke_hoach_tuan_2026
                </div>
                <div style="display: flex; align-items: center; gap: 6px; font-size: 13px; color: #5f6368;">
                  <span style="color: #1a73e8; font-size: 16px;">☁✓</span>
                  <span>Đã lưu vào Drive</span>
                </div>
                <!-- Callout 1 -->
                <div class="callout-box" style="top: -6px; left: -6px; width: 380px; height: 38px;">
                  <div class="step-badge">1</div>
                  <div class="arrow-tag" style="top: -24px; left: 10px;">Tự đặt tên & Trạng thái tự lưu đám mây</div>
                </div>
              </div>
              
              <!-- Menu bar -->
              <div style="display: flex; gap: 16px; font-size: 13.5px; color: #202124;">
                <span style="font-weight: 600; color: #1a73e8; background: #e8f0fe; padding: 2px 8px; border-radius: 4px;">Tệp</span>
                <span>Chỉnh sửa</span><span>Xem</span><span>Chèn</span><span>Định dạng</span><span>Công cụ</span><span>Tiện ích</span>
              </div>
            </div>
            
            <!-- Share button -->
            <button style="background: #c2e7ff; color: #001d35; border: none; padding: 9px 20px; border-radius: 20px; font-weight: 600; font-size: 14px; display: flex; align-items: center; gap: 8px;">
              🔒 Chia sẻ
            </button>
          </div>

          <!-- File Menu Dropdown (Active) -->
          <div style="position: relative; height: 320px; background: #f8f9fa; padding: 10px 60px;">
            <div style="position: absolute; top: 0px; left: 70px; background: white; border: 1px solid #dadce0; border-radius: 6px; box-shadow: 0 4px 16px rgba(0,0,0,0.18); width: 260px; font-size: 13px; z-index: 10;">
              <div style="padding: 8px 16px; color: #3c4043;">Mới...</div>
              <div style="padding: 8px 16px; color: #3c4043;">Mở (Ctrl+O)</div>
              <div style="padding: 8px 16px; color: #3c4043; background: #e8f0fe; font-weight: 600; color: #1a73e8; display: flex; justify-content: space-between;">
                <span>Tải xuống</span> <span>▶</span>
              </div>
              <div style="height: 1px; background: #dadce0; margin: 4px 0;"></div>
              <div style="padding: 8px 16px; color: #3c4043;">Thiết lập trang...</div>
              <div style="padding: 8px 16px; color: #3c4043;">In (Ctrl+P)</div>
            </div>

            <!-- Submenu Download -->
            <div style="position: absolute; top: 60px; left: 325px; background: white; border: 1px solid #dadce0; border-radius: 6px; box-shadow: 0 4px 16px rgba(0,0,0,0.18); width: 280px; font-size: 13px; z-index: 20;">
              <div style="padding: 10px 16px; color: #1a73e8; font-weight: 600; background: #f0f7ff;">
                📄 Microsoft Word (.docx)
              </div>
              <div style="padding: 10px 16px; color: #1a73e8; font-weight: 600; background: #f0f7ff; border-top: 1px solid #e8eaed;">
                📕 Tài liệu PDF (.pdf)
              </div>
              <div style="padding: 8px 16px; color: #3c4043;">Định dạng OpenDocument (.odt)</div>
              <div style="padding: 8px 16px; color: #3c4043;">Văn bản thuần túy (.txt)</div>

              <!-- Callout 2 -->
              <div class="callout-box" style="top: 0px; left: 0px; width: 100%; height: 86px;">
                <div class="step-badge">2</div>
                <div class="arrow-tag" style="top: 25px; right: -150px;">Tải về máy dưới dạng Word hoặc PDF</div>
              </div>
            </div>
          </div>
        </div>
        """
        page.set_content(html_1)
        page.locator(".window").screenshot(path=os.path.join(OUTPUT_DIR, "vol2_img1_docs_autosave.png"))

        # --------------------------------------------------------------------
        # ẢNH 2: Google Docs - Hộp thoại Thiết lập trang (A4, Lề cm theo NĐ 30)
        # --------------------------------------------------------------------
        print("  -> Đang tạo vol2_img2_docs_page_setup.png...")
        html_2 = f"""
        {COMMON_CSS}
        <div class="window" style="width: 620px;">
          <!-- Modal Header -->
          <div style="padding: 16px 24px; border-bottom: 1px solid #dadce0; font-size: 18px; font-weight: 600; color: #202124;">
            Thiết lập trang
          </div>
          <!-- Modal Body -->
          <div style="padding: 20px 24px; font-size: 13.5px; color: #3c4043;">
            <div style="margin-bottom: 18px;">
              <div style="font-weight: 600; margin-bottom: 8px; color: #202124;">Hướng giấy:</div>
              <label style="margin-right: 24px;"><input type="radio" checked style="accent-color: #1a73e8;"> Đứng (Portrait)</label>
              <label><input type="radio" style="accent-color: #1a73e8;"> Ngang (Landscape)</label>
            </div>

            <!-- Paper Size -->
            <div style="margin-bottom: 20px; position: relative;">
              <div style="font-weight: 600; margin-bottom: 6px; color: #202124;">Kích thước giấy:</div>
              <div style="border: 1px solid #dadce0; border-radius: 4px; padding: 8px 12px; width: 260px; font-weight: 500; display: flex; justify-content: space-between; background: #f8f9fa;">
                <span>A4 (210 mm x 297 mm)</span> <span>▼</span>
              </div>
              <div class="callout-box" style="top: 22px; left: -4px; width: 270px; height: 42px;">
                <div class="step-badge">1</div>
                <div class="arrow-tag" style="top: 8px; right: -155px;">Bắt buộc chọn đúng khổ A4</div>
              </div>
            </div>

            <!-- Margins -->
            <div style="position: relative; margin-bottom: 24px;">
              <div style="font-weight: 600; margin-bottom: 8px; color: #202124;">Lề (Centimét):</div>
              <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; max-width: 320px;">
                <div>
                  <div style="font-size: 12px; color: #5f6368;">Trên:</div>
                  <input type="text" value="2.0" style="width: 100%; border: 1px solid #dadce0; border-radius: 4px; padding: 6px 10px; font-weight: bold; color: #1a73e8;">
                </div>
                <div>
                  <div style="font-size: 12px; color: #5f6368;">Dưới:</div>
                  <input type="text" value="2.0" style="width: 100%; border: 1px solid #dadce0; border-radius: 4px; padding: 6px 10px; font-weight: bold; color: #1a73e8;">
                </div>
                <div>
                  <div style="font-size: 12px; color: #5f6368;">Trái (Đóng ghim):</div>
                  <input type="text" value="3.0" style="width: 100%; border: 1px solid #dadce0; border-radius: 4px; padding: 6px 10px; font-weight: bold; color: #1a73e8;">
                </div>
                <div>
                  <div style="font-size: 12px; color: #5f6368;">Phải:</div>
                  <input type="text" value="2.0" style="width: 100%; border: 1px solid #dadce0; border-radius: 4px; padding: 6px 10px; font-weight: bold; color: #1a73e8;">
                </div>
              </div>
              <div class="callout-box" style="top: 22px; left: -6px; width: 335px; height: 115px;">
                <div class="step-badge">2</div>
                <div class="arrow-tag" style="top: 40px; right: -195px;">Lề chuẩn Nghị định 30/2020/NĐ-CP</div>
              </div>
            </div>

            <!-- Footer Buttons -->
            <div style="display: flex; justify-content: space-between; align-items: center; border-top: 1px solid #dadce0; padding-top: 16px; margin-top: 10px;">
              <button style="background: none; border: 1px solid #dadce0; color: #1a73e8; font-weight: 600; padding: 8px 16px; border-radius: 4px; font-size: 13px;">
                Đặt làm mặc định
              </button>
              <div style="display: flex; gap: 12px;">
                <button style="background: none; border: none; color: #5f6368; font-weight: 600; padding: 8px 16px; font-size: 13px;">Hủy</button>
                <button style="background: #1a73e8; color: white; border: none; font-weight: 600; padding: 8px 24px; border-radius: 4px; font-size: 13px;">OK</button>
              </div>
            </div>
          </div>
        </div>
        """
        page.set_content(html_2)
        page.locator(".window").screenshot(path=os.path.join(OUTPUT_DIR, "vol2_img2_docs_page_setup.png"))

        # --------------------------------------------------------------------
        # ẢNH 3: Google Docs - Chia sẻ Phân quyền & Chế độ Đề xuất (Suggesting)
        # --------------------------------------------------------------------
        print("  -> Đang tạo vol2_img3_docs_share_suggesting.png...")
        html_3 = f"""
        {COMMON_CSS}
        <div class="window" style="width: 780px;">
          <!-- Top Tool Bar: Mode Switcher -->
          <div style="background: #f1f3f4; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #dadce0;">
            <div style="font-size: 14px; font-weight: 600; color: #3c4043;">Chế độ thao tác tài liệu:</div>
            
            <!-- Mode dropdown button -->
            <div style="position: relative;">
              <div style="background: #ffffff; border: 1px solid #188038; border-radius: 4px; padding: 6px 14px; font-size: 13.5px; font-weight: 600; color: #188038; display: flex; align-items: center; gap: 8px;">
                <span>✏️ Đề xuất (Suggesting)</span> <span>▼</span>
              </div>
              <div class="callout-box" style="top: -4px; left: -4px; width: 104%; height: 120%;">
                <div class="step-badge">1</div>
                <div class="arrow-tag" style="top: 5px; left: -210px;">Bật chế độ duyệt bài (Track Changes)</div>
              </div>
            </div>
          </div>

          <!-- Document view with Suggestion -->
          <div style="padding: 24px; background: #ffffff; border-bottom: 1px solid #e2e8f0; display: flex; gap: 20px;">
            <div style="flex: 1; font-size: 14px; line-height: 1.6; border: 1px solid #f1f3f4; padding: 16px; border-radius: 4px;">
              Cuộc họp giao ban tuần 38 sẽ diễn ra vào lúc 08h00 Thứ Ba tại 
              <span style="color: #d93025; text-decoration: line-through; background: #fce8e6; padding: 0 4px;">Hội trường A</span> 
              <span style="color: #188038; font-weight: bold; background: #e6f4ea; padding: 0 4px;">Phòng họp số 2 trụ sở cơ quan</span>.
            </div>
            
            <!-- Comment Card -->
            <div style="width: 250px; background: #e6f4ea; border: 1px solid #ceead6; border-radius: 6px; padding: 10px; font-size: 12.5px;">
              <div style="font-weight: bold; color: #137333; margin-bottom: 4px;">Nguyễn Thế Chiến</div>
              <div style="color: #3c4043; margin-bottom: 8px;">Đã đề xuất thay thế "Hội trường A" bằng "Phòng họp số 2"</div>
              <div style="display: flex; gap: 8px;">
                <button style="background: #188038; color: white; border: none; border-radius: 3px; padding: 4px 10px; font-size: 11px; font-weight: bold;">✓ Chấp nhận</button>
                <button style="background: white; border: 1px solid #dadce0; border-radius: 3px; padding: 4px 10px; font-size: 11px;">✕ Từ chối</button>
              </div>
            </div>
          </div>

          <!-- Share Modal Section -->
          <div style="padding: 20px 24px; background: #f8f9fa;">
            <div style="font-size: 15px; font-weight: bold; color: #202124; margin-bottom: 12px;">Phân quyền chia sẻ tài liệu:</div>
            <div style="background: white; border: 1px solid #dadce0; border-radius: 6px; padding: 14px 18px; display: flex; align-items: center; justify-content: space-between; position: relative;">
              <div style="display: flex; align-items: center; gap: 12px;">
                <div style="background: #1a73e8; color: white; font-weight: bold; border-radius: 50%; width: 34px; height: 34px; display: flex; align-items: center; justify-content: center; font-size: 14px;">L</div>
                <div>
                  <div style="font-weight: 600; font-size: 13.5px; color: #202124;">Lê Thị Mai Lan (Chuyên viên Tổng hợp)</div>
                  <div style="font-size: 12px; color: #5f6368;">mailan.le@hanoi.gov.vn</div>
                </div>
              </div>

              <!-- Permission Dropdown -->
              <div style="border: 1px solid #1a73e8; border-radius: 4px; padding: 6px 12px; font-size: 13px; font-weight: bold; color: #1a73e8; background: #e8f0fe;">
                Người nhận xét (Commenter) ▼
              </div>

              <div class="callout-box" style="top: 14px; right: 10px; width: 220px; height: 38px;">
                <div class="step-badge">2</div>
                <div class="arrow-tag" style="top: -24px; right: 10px;">Chọn: Xem / Nhận xét / Chỉnh sửa</div>
              </div>
            </div>
          </div>
        </div>
        """
        page.set_content(html_3)
        page.locator(".window").screenshot(path=os.path.join(OUTPUT_DIR, "vol2_img3_docs_share_suggesting.png"))

        # --------------------------------------------------------------------
        # ẢNH 4: Google Docs - Lịch sử phiên bản (Version History) & Khôi phục
        # --------------------------------------------------------------------
        print("  -> Đang tạo vol2_img4_docs_version_history.png...")
        html_4 = f"""
        {COMMON_CSS}
        <div class="window" style="width: 820px; height: 420px; display: flex; flex-direction: column;">
          <!-- Top Bar -->
          <div style="background: #202124; color: white; padding: 12px 20px; display: flex; justify-content: space-between; align-items: center;">
            <div style="display: flex; align-items: center; gap: 14px;">
              <span style="font-size: 18px; cursor: pointer;">←</span>
              <span style="font-weight: 500; font-size: 15px;">Lịch sử phiên bản: Du_thao_Ke_hoach_tuan_2026</span>
            </div>
            
            <!-- Restore Button -->
            <div style="position: relative;">
              <button style="background: #1a73e8; color: white; border: none; font-weight: bold; font-size: 13.5px; padding: 8px 18px; border-radius: 4px;">
                Khôi phục phiên bản này
              </button>
              <div class="callout-box" style="top: -4px; left: -4px; width: 104%; height: 120%;">
                <div class="step-badge">1</div>
                <div class="arrow-tag" style="top: 4px; left: -190px;">Bấm khôi phục lại bản cũ hoàn hảo</div>
              </div>
            </div>
          </div>

          <!-- Main Split View -->
          <div style="display: flex; flex: 1; background: #e8eaed;">
            <!-- Document Canvas -->
            <div style="flex: 1; padding: 24px; display: flex; justify-content: center; overflow: hidden;">
              <div style="background: white; width: 440px; height: 100%; box-shadow: 0 2px 8px rgba(0,0,0,0.15); padding: 24px; font-size: 13px; line-height: 1.6;">
                <div style="font-weight: bold; text-align: center; margin-bottom: 12px;">KẾ HOẠCH CÔNG TÁC TUẦN 38</div>
                <div>1. Thứ Hai: Giao ban Lãnh đạo Sở lúc 08h00.</div>
                <div style="background: #c2e7ff; padding: 2px 4px; border-radius: 2px;">2. Thứ Ba: Tiếp công dân định kỳ tại Trụ sở (Cán bộ: Nguyễn Thế Chiến).</div>
                <div>3. Thứ Tư: Kiểm tra chuyển đổi số tại các đơn vị cơ sở.</div>
              </div>
            </div>

            <!-- Version History Sidebar -->
            <div style="width: 280px; background: white; border-left: 1px solid #dadce0; padding: 16px; position: relative;">
              <div style="font-size: 14px; font-weight: bold; color: #202124; margin-bottom: 14px;">Lịch sử chỉnh sửa</div>
              
              <!-- Selected Version -->
              <div style="background: #e8f0fe; border-radius: 6px; padding: 12px; margin-bottom: 10px; border-left: 4px solid #1a73e8;">
                <div style="font-weight: bold; font-size: 13px; color: #1a73e8;">Hôm nay, 09:30</div>
                <div style="font-size: 12px; color: #5f6368; margin-top: 4px; display: flex; align-items: center; gap: 6px;">
                  <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #1a73e8;"></span>
                  <span>Nguyễn Thế Chiến (Bản nộp duyệt)</span>
                </div>
              </div>

              <!-- Older Version -->
              <div style="padding: 12px; border-radius: 6px; color: #5f6368; font-size: 13px;">
                <div style="font-weight: 500;">Hôm qua, 15:45</div>
                <div style="font-size: 12px; color: #70757a; margin-top: 4px; display: flex; align-items: center; gap: 6px;">
                  <span style="display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #ea4335;"></span>
                  <span>Lê Thị Mai Lan</span>
                </div>
              </div>

              <div class="callout-box" style="top: 45px; left: 8px; width: 260px; height: 68px;">
                <div class="step-badge">2</div>
                <div class="arrow-tag" style="top: 20px; left: -180px;">Mã màu truy vết từng người sửa</div>
              </div>
            </div>
          </div>
        </div>
        """
        page.set_content(html_4)
        page.locator(".window").screenshot(path=os.path.join(OUTPUT_DIR, "vol2_img4_docs_version_history.png"))

        # --------------------------------------------------------------------
        # ẢNH 5: Google Sheets - Cài đặt Vùng Việt Nam & Kẻ khung All Borders
        # --------------------------------------------------------------------
        print("  -> Đang tạo vol2_img5_sheets_settings_borders.png...")
        html_5 = f"""
        {COMMON_CSS}
        <div class="window" style="width: 780px;">
          <!-- Sheets Toolbar -->
          <div style="background: #f9fbfd; padding: 10px 18px; border-bottom: 1px solid #dadce0; display: flex; align-items: center; gap: 14px;">
            <div style="background: #188038; color: white; font-weight: bold; width: 32px; height: 32px; border-radius: 4px; display: flex; align-items: center; justify-content: center; font-size: 18px;">⊞</div>
            <div style="font-size: 16px; font-weight: 500; color: #202124;">Theo_doi_Tien_do_Cong_tac_2026</div>
          </div>

          <!-- Formula and Formatting Bar -->
          <div style="padding: 8px 18px; background: white; border-bottom: 1px solid #e0e0e0; display: flex; align-items: center; gap: 16px; position: relative;">
            <div style="display: flex; gap: 8px; align-items: center;">
              <span style="font-weight: bold; font-size: 14px; padding: 4px 8px; border: 1px solid #dadce0; border-radius: 4px;">B</span>
              <span style="font-style: italic; font-size: 14px; padding: 4px 8px; border: 1px solid #dadce0; border-radius: 4px;">I</span>
            </div>

            <!-- Borders and Merge button highlight -->
            <div style="display: flex; gap: 8px; align-items: center; position: relative;">
              <div style="background: #e6f4ea; border: 1px solid #188038; border-radius: 4px; padding: 4px 10px; font-size: 13px; font-weight: bold; color: #137333; display: flex; align-items: center; gap: 6px;">
                <span>田 Tất cả đường viền (All borders)</span>
              </div>
              <div style="background: #f1f3f4; border: 1px solid #dadce0; border-radius: 4px; padding: 4px 10px; font-size: 13px; color: #3c4043;">
                <span>⇋ Gộp ô (Merge)</span>
              </div>

              <div class="callout-box" style="top: -4px; left: -4px; width: 104%; height: 120%;">
                <div class="step-badge">1</div>
                <div class="arrow-tag" style="top: -24px; left: 10px;">Thao tác kẻ khung viền và gộp ô</div>
              </div>
            </div>
          </div>

          <!-- Settings Modal View (Overlay) -->
          <div style="padding: 20px; background: #f8f9fa; display: flex; justify-content: center;">
            <div style="background: white; border: 1px solid #dadce0; border-radius: 8px; width: 460px; box-shadow: 0 4px 16px rgba(0,0,0,0.12); padding: 18px 24px; position: relative;">
              <div style="font-size: 16px; font-weight: bold; color: #202124; margin-bottom: 14px; border-bottom: 1px solid #dadce0; padding-bottom: 8px;">
                Cài đặt cho bảng tính này
              </div>
              
              <div style="margin-bottom: 14px;">
                <div style="font-size: 12.5px; font-weight: 600; color: #5f6368; margin-bottom: 4px;">Vùng (Locale):</div>
                <div style="border: 1px solid #188038; border-radius: 4px; padding: 7px 12px; font-size: 13.5px; font-weight: bold; color: #137333; background: #e6f4ea; display: flex; justify-content: space-between;">
                  <span>Việt Nam</span> <span>▼</span>
                </div>
              </div>

              <div style="margin-bottom: 18px;">
                <div style="font-size: 12.5px; font-weight: 600; color: #5f6368; margin-bottom: 4px;">Múi giờ (Time zone):</div>
                <div style="border: 1px solid #dadce0; border-radius: 4px; padding: 7px 12px; font-size: 13px; color: #202124;">
                  (GMT+07:00) Giờ Hà Nội
                </div>
              </div>

              <div class="callout-box" style="top: 48px; left: 18px; width: 424px; height: 60px;">
                <div class="step-badge">2</div>
                <div class="arrow-tag" style="top: 15px; right: -195px;">Chọn Việt Nam để ngày hiển thị dd/mm/yyyy</div>
              </div>

              <div style="text-align: right; border-top: 1px solid #dadce0; padding-top: 12px;">
                <button style="background: #188038; color: white; border: none; font-weight: bold; padding: 7px 18px; border-radius: 4px; font-size: 13px;">
                  Lưu và tải lại
                </button>
              </div>
            </div>
          </div>
        </div>
        """
        page.set_content(html_5)
        page.locator(".window").screenshot(path=os.path.join(OUTPUT_DIR, "vol2_img5_sheets_settings_borders.png"))

        # --------------------------------------------------------------------
        # ẢNH 6: Google Sheets - Tự động gợi ý công thức (Smart Suggestions)
        # --------------------------------------------------------------------
        print("  -> Đang tạo vol2_img6_sheets_smart_suggestions.png...")
        html_6 = f"""
        {COMMON_CSS}
        <div class="window" style="width: 760px;">
          <div class="browser-bar">
            <div class="dot dot-red"></div><div class="dot dot-yellow"></div><div class="dot dot-green"></div>
            <div class="url-bar">https://docs.google.com/spreadsheets/d/1X9y.../edit</div>
          </div>

          <!-- Sheets Grid with Formula Popup -->
          <div style="padding: 24px; background: white;">
            <table style="width: 100%; border-collapse: collapse; font-size: 13px; border: 1px solid #dadce0;">
              <thead>
                <tr style="background: #e6f4ea; color: #137333; font-weight: bold; text-align: center;">
                  <th style="border: 1px solid #dadce0; padding: 8px; width: 50px;">A</th>
                  <th style="border: 1px solid #dadce0; padding: 8px;">B</th>
                  <th style="border: 1px solid #dadce0; padding: 8px; width: 120px;">C</th>
                  <th style="border: 1px solid #dadce0; padding: 8px; width: 140px;">D</th>
                </tr>
                <tr style="background: #f8f9fa; font-weight: 600; text-align: center;">
                  <td style="border: 1px solid #dadce0; padding: 6px; background: #eaedf1;">1</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: left;">Nội dung chi hoạt động</td>
                  <td style="border: 1px solid #dadce0; padding: 6px;">Số lượng</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: right;">Thành tiền (đ)</td>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: center; background: #eaedf1;">2</td>
                  <td style="border: 1px solid #dadce0; padding: 6px;">Mua giấy in A4 cho phòng</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: center;">5 ram</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: right;">375.000</td>
                </tr>
                <tr>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: center; background: #eaedf1;">3</td>
                  <td style="border: 1px solid #dadce0; padding: 6px;">Thay mực máy in Canon</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: center;">2 hộp</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: right;">650.000</td>
                </tr>
                <tr>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: center; background: #eaedf1;">4</td>
                  <td style="border: 1px solid #dadce0; padding: 6px;">Nước uống hội nghị quý III</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: center;">10 bình</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: right;">500.000</td>
                </tr>
                <!-- Active cell with Smart Suggestion -->
                <tr>
                  <td style="border: 1px solid #dadce0; padding: 6px; text-align: center; background: #eaedf1;">5</td>
                  <td style="border: 1px solid #dadce0; padding: 6px; font-weight: bold; text-align: right;" colspan="2">Tổng cộng kinh phí:</td>
                  <td style="border: 2px solid #1a73e8; padding: 6px; text-align: left; position: relative; background: #ffffff;">
                    <span style="font-weight: bold; color: #1a73e8;">=</span>
                    
                    <!-- AI Suggestion Bubble -->
                    <div style="position: absolute; top: 34px; left: -80px; background: #202124; color: white; border-radius: 6px; padding: 8px 14px; font-size: 12.5px; box-shadow: 0 4px 14px rgba(0,0,0,0.25); z-index: 100; white-space: nowrap;">
                      <span style="color: #8ab4f8; font-weight: bold;">=SUM(D2:D4)</span>
                      <span style="color: #9aa0a6; margin-left: 10px;">Nhấn <b>Tab</b> để chấp nhận</span>
                    </div>

                    <div class="callout-box" style="top: 26px; left: -90px; width: 280px; height: 48px;">
                      <div class="step-badge">1</div>
                      <div class="arrow-tag" style="top: 45px; left: 20px;">Trí tuệ nhân tạo tự động gợi ý công thức SUM</div>
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
        """
        page.set_content(html_6)
        page.locator(".window").screenshot(path=os.path.join(OUTPUT_DIR, "vol2_img6_sheets_smart_suggestions.png"))

        # --------------------------------------------------------------------
        # ẢNH 7: Google Sheets - Bộ lọc xem riêng (Filter Views) & Khóa ô bảo vệ
        # --------------------------------------------------------------------
        print("  -> Đang tạo vol2_img7_sheets_filter_view_protect.png...")
        html_7 = f"""
        {COMMON_CSS}
        <div class="window" style="width: 840px;">
          <!-- Filter View Dark Banner (Signature of Google Sheets Filter View) -->
          <div style="background: #202124; color: white; padding: 8px 20px; display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #5f6368; position: relative;">
            <div style="display: flex; align-items: center; gap: 12px; font-size: 13.5px;">
              <span style="color: #8ab4f8; font-weight: bold;">Tên bộ lọc:</span>
              <div style="background: #3c4043; border-radius: 4px; padding: 3px 12px; font-weight: 600; border: 1px dashed #8ab4f8;">
                Việc_can_lam_gap (Xem riêng của An)
              </div>
              <span style="color: #9aa0a6; font-size: 12px;">(Thao tác lọc này không làm ảnh hưởng người khác)</span>
            </div>
            <div style="color: #bdc1c6; font-size: 18px; font-weight: bold; cursor: pointer;">✕</div>

            <div class="callout-box" style="top: 2px; left: 4px; width: 620px; height: 38px;">
              <div class="step-badge">1</div>
              <div class="arrow-tag" style="top: -24px; left: 20px;">Dải viền đen sẫm: Bạn đang ở Chế độ xem bộ lọc riêng!</div>
            </div>
          </div>

          <!-- Split View: Table + Protect Side Panel -->
          <div style="display: flex; background: white;">
            <!-- Filtered Table -->
            <div style="flex: 1; padding: 18px;">
              <table style="width: 100%; border-collapse: collapse; font-size: 12.5px; border: 1px solid #dadce0;">
                <thead>
                  <tr style="background: #f1f3f4; font-weight: bold;">
                    <th style="border: 1px solid #dadce0; padding: 6px;">Nhiệm vụ</th>
                    <th style="border: 1px solid #dadce0; padding: 6px;">Người phụ trách ▼</th>
                    <th style="border: 1px solid #dadce0; padding: 6px;">Tình trạng ▼</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style="border: 1px solid #dadce0; padding: 6px;">Dự thảo Kế hoạch tuần 38</td>
                    <td style="border: 1px solid #dadce0; padding: 6px;">Nguyễn Văn An</td>
                    <td style="border: 1px solid #dadce0; padding: 6px; color: #c5221f; font-weight: bold;">Trễ hạn (Gấp)</td>
                  </tr>
                  <tr>
                    <td style="border: 1px solid #dadce0; padding: 6px;">Chuẩn bị tài liệu giao ban</td>
                    <td style="border: 1px solid #dadce0; padding: 6px;">Nguyễn Văn An</td>
                    <td style="border: 1px solid #dadce0; padding: 6px; color: #137333; font-weight: bold;">Đang xử lý</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <!-- Protect Sheets & Ranges Side Panel -->
            <div style="width: 290px; background: #f8f9fa; border-left: 1px solid #dadce0; padding: 14px; position: relative;">
              <div style="font-size: 13.5px; font-weight: bold; color: #202124; margin-bottom: 12px; display: flex; align-items: center; gap: 6px;">
                <span>🔒 Trang tính và dải ô được bảo vệ</span>
              </div>

              <div style="background: white; border: 1px solid #dadce0; border-radius: 6px; padding: 10px; margin-bottom: 12px; font-size: 12.5px;">
                <div style="font-weight: 600; color: #1a73e8; margin-bottom: 4px;">Dải ô công thức D2:D20</div>
                <div style="color: #5f6368; font-size: 11.5px;">Quyền chỉnh sửa: <b>Chỉ bạn</b> (Người khác bị khóa)</div>
              </div>

              <button style="width: 100%; background: #1a73e8; color: white; border: none; font-weight: 600; padding: 7px; border-radius: 4px; font-size: 12.5px;">
                + Thêm trang tính hoặc dải ô
              </button>

              <div class="callout-box" style="top: 10px; left: 4px; width: 280px; height: 110px;">
                <div class="step-badge">2</div>
                <div class="arrow-tag" style="top: 40px; left: -195px;">Khóa bảo vệ không cho người khác xóa hàm</div>
              </div>
            </div>
          </div>
        </div>
        """
        page.set_content(html_7)
        page.locator(".window").screenshot(path=os.path.join(OUTPUT_DIR, "vol2_img7_sheets_filter_view_protect.png"))

        browser.close()
    print(">>> HOÀN THÀNH TẠO SINH 7 ẢNH CHỤP CHO TẬP II!")

if __name__ == "__main__":
    generate_all_vol2_images()
