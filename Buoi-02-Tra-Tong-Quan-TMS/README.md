# Buổi 2 – Tự tạo tool tra tổng quan TMS bằng Gemini

Giảng viên: Ngô Hoàng Thuỳ Hương · Zalo 0328 878 297

Slide bài giảng: [Buoi2_Slide.pdf](Buoi2_Slide.pdf) · Quay về [mục lục khóa học](../README.md)

> 🆘 **Bị kẹt?** Vào [**TRỢ GIÚP PROMPT**](Tro-Giup-Prompt/README.md): có **bộ trợ giúp** (4 file mẫu chạy được, tải 1 file zip), prompt chuẩn để Gemini sửa đúng chỗ, và cách đọc log để sửa lỗi.

> ⚠️ **Tắt dịch tự động của trình duyệt** (bấm biểu tượng 文A trên thanh địa chỉ › *Hiển thị bản gốc*). Trang này viết sẵn tiếng Việt; bản dịch máy làm sai chữ, ví dụ "Chưa bấm Enter" bị dịch thành "Nhấn Enter".

Mục tiêu: học viên tự ghi kịch bản tra tổng quan 1 MST trên SAP GUI (FMCACOV), rồi ra lệnh cho Gemini để có bot tra tổng quan cho nhiều MST, ghi kết quả ra Excel. Không cần biết lập trình.

---

## Quy trình 5 bước

1. **Chuẩn bị**: bật Scripting (Alt+F12 › Options › Accessibility & Scripting › Scripting › tick *Enable scripting*, bỏ tick 2 ô *Notify when…*), có Excel, tạo thư mục **BOT_TONG_QUAN** trên Desktop.
2. **Ghi kịch bản**: Alt+F12 › *Script Recording and Playback…* › *More* (chọn chỗ lưu) › *Record* › làm tay trọn 1 MST trên FMCACOV › *Stop*.
3. **Ra lệnh Gemini**: làm theo mục **"Bước 3 – Cách gửi cho Gemini"** ngay bên dưới.
4. **Ráp và chạy thử**: lưu đúng tên, đúng encoding, thử với 3–5 MST.
5. **Sửa lỗi và nâng cấp**: dùng prompt sửa lỗi, mỗi lần xin một tính năng.

> **Bảo mật:** không bao giờ dán user hoặc mật khẩu TMS lên Gemini.

---

## Bước 3 – Cách gửi cho Gemini (hướng dẫn thao tác)

> 📖 Đây là hướng dẫn để đọc. **Chỉ copy khung xám có dòng 👉 COPY KHUNG NÀY.**

**Nhớ 1 câu: dán kịch bản trước, dán prompt sau, rồi Gửi.**

| Bước | Làm gì |
|---|---|
| 1 | Chuột phải file kịch bản ghi › **Open with › Notepad** › **Ctrl+A**, **Ctrl+C**. |
| 2 | Vào **gemini.google.com** › **New chat** › **Ctrl+V**. *Chưa Gửi.* |
| 3 | **Shift+Enter** 2 lần › copy **khung Prompt 1** bên dưới › **Ctrl+V**. |
| 4 | Bấm **Gửi**. |

> Lưu ý: không bấm đúp file kịch bản (bấm đúp là chạy). Trong Gemini, **Enter là gửi luôn**, xuống dòng thì **Shift+Enter**.

**Sau khi Gemini trả lời:** lưu file **.vbs** (mục *Lưu code thành file* bên dưới). Rồi copy **khung Prompt 2**, dán vào **cùng cuộc trò chuyện** đó, Gửi, lưu tiếp 3 file.

> Code bị cắt giữa chừng? Gõ: "Gửi lại toàn bộ file trong 1 khối code, không lược."

---

## Prompt 1 – Bot **.vbs**

👉 **COPY KHUNG NÀY dán vào Gemini, ngay dưới kịch bản đã dán** (bấm nút copy góc phải khung):

```
[VAI TRÒ] Bạn là lập trình viên VBScript chuyên SAP GUI Scripting.
[BỐI CẢNH] Tôi là cán bộ thuế, dùng TMS (SAP GUI) trên máy cơ quan. Máy KHÔNG cài được Python, KHÔNG chạy macro VBA. Chỉ có: wscript (VBScript), PowerShell 5, Excel.
[KỊCH BẢN GHI] Phần tôi dán ở PHÍA TRÊN tin nhắn này là kịch bản tôi ghi bằng Script Recording khi tra tổng quan 1 MST trên màn FMCACOV. MST trong kịch bản là MST tôi gõ lúc ghi (ở dòng có .text = "...").
[VIỆC CẦN LÀM] Viết file BOT_TONG_QUAN.vbs chạy cho NHIỀU MST:
1. Đọc MST_CHAY.txt cùng thư mục: mỗi dòng 1 MST, bỏ dòng trống, bỏ MST trùng. MST hợp lệ: 10 số, 10 số-3 số, hoặc 12 số.
2. Kết nối SAP GUI đang mở bằng GetObject("SAPGUI"), dùng phiên đầu tiên. Chưa mở SAP hoặc chưa bật scripting thì báo lỗi dễ hiểu rồi thoát.
3. Mỗi MST: làm lại đúng các bước trong kịch bản, thay MST lúc ghi bằng MST đang chạy. Chờ SAP xử lý xong (session.Busy) trước khi đọc, chờ tối đa 30 giây.
4. Đọc lưới kết quả (GuiGridView): lấy HẾT dòng (cuộn firstVisibleRow để tải đủ) và hết cột.
5. Ghi KET_QUA_TONG_QUAN.xlsx: sheet "Ket qua" (cột A = MST, sau đó các cột của lưới), sheet "Tien do" (STT, MST, tên NNT, trạng thái OK / KHONG CO DU LIEU / LOI + chữ ở sbar). Lưu file sau MỖI MST để dừng giữa chừng vẫn còn dữ liệu.
6. Gặp popup lạ (wnd[1], wnd[2]...): đóng bằng sendVKey 12, quay về FMCACOV, ghi LOI cho MST đó rồi chạy tiếp MST sau, không dừng cả bot.
7. Trước mỗi MST: thấy file STOP_TONG_QUAN.txt thì lưu Excel và dừng.
8. Ghi LOG_TONG_QUAN.txt (giờ, MST, bước, lỗi) và PROGRESS_TONG_QUAN.txt (dòng/tổng|MST|trạng thái) để app đọc tiến độ.
[RÀNG BUỘC] Toàn bộ code chỉ dùng ký tự ASCII: không dấu tiếng Việt, kể cả chú thích và MsgBox. Có Option Explicit. Mỗi hàm có 1 dòng chú thích. Không dùng thư viện ngoài.
[TRẢ VỀ] Một khối code HOÀN CHỈNH, không lược đoạn nào. Sau code, liệt kê các ID findById tôi cần kiểm tra lại trên máy mình.
```

## Prompt 2 – App giao diện + 2 file **.bat**

👉 **COPY KHUNG NÀY dán vào Gemini**, gửi tiếp trong **cùng cuộc trò chuyện** với Prompt 1:

```
[VIỆC CẦN LÀM] Viết thêm 3 file cho bot trên, để cùng thư mục:
A. APP_TONG_QUAN.ps1 – giao diện PowerShell 5 WinForms:
- Ô dán danh sách MST (Ctrl+V), nút "Chon file .txt", kéo thả file; đếm MST hợp lệ.
- Ô "Tu dong" / "Den dong" (để trống Den dong = chạy hết) để hôm sau chạy tiếp.
- Nút LAY TONG QUAN: hỏi xác nhận số MST, ghi MST_CHAY.txt (ASCII), xóa file STOP, rồi chạy: wscript //nologo BOT_TONG_QUAN.vbs
- Nút DUNG BOT: tạo STOP_TONG_QUAN.txt. Đóng cửa sổ cũng dừng bot.
- Dòng trạng thái đọc PROGRESS_TONG_QUAN.txt mỗi 1 giây.
B. CHAY.bat – mở app bằng: powershell -NoProfile -ExecutionPolicy Bypass -STA -File APP_TONG_QUAN.ps1. Báo lỗi dễ hiểu nếu thiếu file .vbs hoặc .ps1.
C. DUNG.bat – tạo file STOP_TONG_QUAN.txt.
[RÀNG BUỘC] Chữ trên giao diện viết không dấu. Không cần quyền admin.
[TRẢ VỀ] 3 khối code riêng, mỗi khối ghi rõ tên file và cách lưu.
```

## Lưu code thành file – từng bước

**Việc làm trước (1 lần): bật hiện đuôi file**

- Windows 11: mở thư mục bất kỳ › **View** (Xem) › **Show** (Hiện) › tick **File name extensions**.
- Windows 10: mở thư mục bất kỳ › tab **View** › tick **File name extensions**.
- Tạo thư mục **BOT_TONG_QUAN** trên Desktop: chuột phải ra màn hình › New › Folder › gõ tên › Enter.

**Lưu từng file (làm 4 lần)**

1. Bấm phím **Windows**, gõ **notepad**, bấm **Enter**.
2. Sang Gemini, ở góc trên bên phải khối code, bấm nút **Copy** (biểu tượng 2 ô vuông). Đừng bôi đen bằng chuột.
3. Quay lại Notepad, bấm **Ctrl+V**. Kéo xuống cuối xem code còn đủ.
4. Bấm **Ctrl+Shift+S** (File › Save as). Bên trái chọn **Desktop**, bấm đúp mở thư mục **BOT_TONG_QUAN**.
5. Ô **Save as type**: chọn **All files (\*.\*)**.
6. Ô **File name**: gõ đúng tên theo bảng dưới.
7. Ô **Encoding** (cạnh nút Save): chọn theo bảng, rồi bấm **Save**.
8. Bấm **Ctrl+N** mở trang trắng, làm lại từ bước 2 cho file tiếp theo.

| Gõ vào ô File name | Lấy code từ | Ô Encoding chọn |
|---|---|---|
| **BOT_TONG_QUAN.vbs** | Prompt 1 | ANSI |
| **APP_TONG_QUAN.ps1** | Prompt 2 – khối A | UTF-8 with BOM |
| **CHAY.bat** | Prompt 2 – khối B | ANSI |
| **DUNG.bat** | Prompt 2 – khối C | ANSI |

**Kiểm tra:** thư mục có đúng 4 file, đuôi **.vbs** **.ps1** **.bat**, không có **.txt**.
- Lỡ có **.txt**: chuột phải file › **Rename** › xóa **.txt** ở cuối › Enter › Yes.
- Lỡ sai encoding: mở lại file bằng Notepad › Ctrl+Shift+S › chọn lại Encoding › Save › Yes để ghi đè.

## Chạy thử với 3–5 MST

1. Chuột phải từng file › **Properties** › nếu dưới cùng có ô **Unblock** thì tick › **OK**.
2. Mở SAP, đăng nhập TMS, về màn hình chính.
3. Bấm đúp **CHAY.bat** (không bấm đúp file **.vbs**). Cửa sổ app hiện ra.
4. Copy 3–5 MST thử (mỗi dòng 1 MST, có 1 MST sai), bấm **Dan clipboard** hoặc Ctrl+V vào ô danh sách.
5. Bấm **LAY TONG QUAN** › **Yes**. Không đụng chuột vào SAP và Excel khi bot chạy.
6. Xong, mở **KET_QUA_TONG_QUAN.xlsx** trong thư mục, so 1 MST với tra tay trên FMCACOV.

## Prompt sửa lỗi

| Bước | Làm gì |
|---|---|
| 1 | Mở **LOG_TONG_QUAN.txt** bằng Notepad › **Ctrl+End** › bôi đen khoảng 20 dòng cuối › **Ctrl+C**. |
| 2 | Vào **đúng cuộc trò chuyện Gemini đã tạo bot** › **Ctrl+V**. |
| 3 | **Shift+Enter** 2 lần › copy **khung dưới đây** › **Ctrl+V** › gõ thông báo lỗi (nếu có) › **Gửi**. |

👉 **COPY KHUNG NÀY dán vào Gemini, ngay dưới đoạn log đã dán:**

```
Bot của tôi bị lỗi khi chạy. Phần tôi dán ở PHÍA TRÊN là 20 dòng cuối của LOG_TONG_QUAN.txt.
Thông báo lỗi Windows hiện ra (nếu có) tôi gõ ở cuối tin nhắn này.
Hãy:
1) Giải thích nguyên nhân trong 2–3 câu dễ hiểu.
2) Chỉ sửa hàm liên quan, trả về NGUYÊN hàm đã sửa, không đổi phần khác.
3) Nói rõ tôi thay đoạn đó vào chỗ nào trong file.
Thông báo lỗi:
```

## Lỗi thường gặp

| Hiện tượng | Thường do | Làm gì |
|---|---|---|
| "The control could not be found by id" | ID trên máy mình khác, hoặc có popup bất ngờ | Ghi lại riêng bước đó, dán cho Gemini: "thay ID cũ bằng ID này" |
| "Invalid character" khi chạy .vbs | Code có chữ có dấu | "Viết lại toàn bộ, chỉ dùng ký tự ASCII" |
| Báo không thấy SAPGUI | Chưa mở SAP hoặc chưa bật Scripting | Mở SAP, bật Scripting |
| Mỗi MST chỉ có khoảng 20–30 dòng | Lưới chưa cuộn để tải hết | "Cuộn firstVisibleRow để đọc đủ RowCount" |
| Bot đứng yên, không báo gì | Popup đang chờ bấm hoặc SAP chờ lâu | "Thêm đóng popup và giới hạn 30 giây mỗi bước" |
| Không lưu được file kết quả | File Excel kết quả đang mở | Đóng file Excel rồi chạy lại |
| CHAY.bat mở xong tắt ngay | Windows chặn file tải về | Properties › Unblock cho cả 4 file |

## Nâng cấp từng bước

Mỗi lần xin một tính năng, chạy thử xong mới xin tiếp. Nếu mở cuộc trò chuyện mới, dán kèm code hiện tại và nhắc "giữ nguyên ràng buộc ASCII".

- **Chọn tiểu mục:** "Thêm ô Tiểu mục trên app (vd 1601, 1602, 1649; * = tất cả), ghi ra TIEU_MUC_CHAY.txt; bot chỉ lấy dòng có tiểu mục đó."
- **Sheet tổng hợp:** "Thêm sheet Tong hop: cộng số tiền theo MST và tiểu mục."
- **Tách theo CQT:** "Tách sheet riêng cho từng cơ quan thuế, lấy từ cột cơ quan thu."
- **Khóa theo user:** "Bot chỉ chạy khi session.Info.User đúng user tôi cho phép, sai thì báo và thoát."
- **Nhớ lựa chọn:** "App nhớ tiểu mục và khoảng dòng lần trước, lưu vào file .ini cùng thư mục."

## Bài tập

Ghi kịch bản FMCACOV cho 1 MST, dùng Prompt 1 và Prompt 2 để có đủ 4 file, chạy với 5 MST. Đạt khi:

- [ ] Có file Excel kết quả cho đủ 5 MST
- [ ] Bấm DUNG BOT giữa chừng, Excel vẫn còn dữ liệu
- [ ] 1 MST sai: bot ghi LOI rồi chạy tiếp
- [ ] Số liệu 1 MST khớp với tra tay
- [ ] Tự sửa được 1 lỗi bằng prompt sửa lỗi

**Mình ghi. Gemini viết. Mình kiểm.**
