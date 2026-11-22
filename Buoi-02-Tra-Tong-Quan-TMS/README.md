# Buổi 2 – Tự tạo tool tra tổng quan TMS bằng Gemini

Giảng viên: Ngô Hoàng Thuỳ Hương · Zalo 0328 878 297

Slide bài giảng: [Buoi2_Slide.pdf](Buoi2_Slide.pdf) · Quay về [mục lục khóa học](../README.md)

> 🆘 **Bị kẹt?** Vào [**TRỢ GIÚP PROMPT**](Tro-Giup-Prompt/README.md): có **bộ trợ giúp** (4 file mẫu chạy được, tải 1 file zip), prompt chuẩn để Gemini sửa đúng chỗ, và cách đọc log để sửa lỗi.

> ⚠️ **Tắt dịch tự động của trình duyệt** (bấm biểu tượng 文A trên thanh địa chỉ › *Hiển thị bản gốc*). Trang này viết sẵn tiếng Việt; bản dịch máy làm sai chữ, ví dụ "Chưa bấm Enter" bị dịch thành "Nhấn Enter".

Mục tiêu: học viên tự ghi kịch bản tra tổng quan 1 MST trên SAP GUI (FMCACOV), rồi ra lệnh cho Gemini để có bot tra tổng quan cho nhiều MST, ghi kết quả ra Excel. Không cần biết lập trình.

---

## Quy trình 5 bước

1. **Chuẩn bị**: bật Scripting (Alt+F12 › Options › Accessibility & Scripting › Scripting › tick *Enable scripting*, bỏ tick 2 ô *Notify when…*), có Excel, tạo thư mục `BOT_TONG_QUAN` trên Desktop.
2. **Ghi kịch bản**: Alt+F12 › *Script Recording and Playback…* › *More* (chọn chỗ lưu) › *Record* › làm tay trọn 1 MST trên FMCACOV › *Stop*.
3. **Ra lệnh Gemini**: làm theo mục **"Bước 3 – Cách gửi cho Gemini"** ngay bên dưới.
4. **Ráp và chạy thử**: lưu đúng tên, đúng encoding, thử với 3–5 MST.
5. **Sửa lỗi và nâng cấp**: dùng prompt sửa lỗi, mỗi lần xin một tính năng.

> **Bảo mật:** trước khi dán lên Gemini, thay MST thật bằng MST mẫu `0100000000`. Không dán tên, địa chỉ, số tiền của NNT. Không bao giờ dán user hoặc mật khẩu TMS.

---

## Bước 3 – Cách gửi cho Gemini (hướng dẫn thao tác)

> 📖 Phần này là **hướng dẫn để đọc và làm theo**. **KHÔNG dán phần này vào Gemini.**
> Thứ dán vào Gemini chỉ là **khung xám** có dòng 👉 *"COPY KHUNG NÀY"* ở bên dưới.

Làm **đúng thứ tự** (dán prompt trước, copy kịch bản sau – nếu copy kịch bản trước thì khi copy prompt, kịch bản sẽ bị ghi đè mất):

1. Mở **gemini.google.com**, đăng nhập Google, bấm **New chat**. Có chế độ *Pro* hoặc *Thinking* thì chọn.
2. Kéo xuống khung xám **Prompt 1** bên dưới, bấm nút copy (góc phải khung), dán vào ô chat Gemini. **Chưa bấm Enter, chưa Gửi.**
3. Mở thư mục đã chọn ở nút *More* lúc ghi. Chuột phải file `GHI_TONG_QUAN.vbs` › **Open with** (Mở bằng) › **Notepad**. *Không bấm đúp* – bấm đúp là chạy kịch bản.
4. Bấm **Ctrl+H** (Thay thế). Ô *Find what*: gõ MST thật đã dùng lúc ghi. Ô *Replace with*: gõ `0100000000`. Bấm **Replace All**.
5. Bấm **Ctrl+A** (chọn hết) rồi **Ctrl+C** (copy). Đóng Notepad, bấm **Don't Save** (Không lưu).
6. Quay lại Gemini. Trong ô chat, dùng chuột **bôi đen dòng** `(dán toàn bộ nội dung file .vbs vừa ghi)`, bấm **Ctrl+V** – kịch bản thay vào đúng chỗ.
7. Kiểm tra không còn MST thật, rồi bấm **Gửi**. (Cần xuống dòng trong ô chat thì bấm **Shift+Enter**, vì Enter là gửi luôn.)
8. Lưu file `.vbs` Gemini trả về (mục *Lưu code thành file* bên dưới). Rồi copy khung **Prompt 2**, dán vào **cùng cuộc trò chuyện**, gửi, lưu tiếp 3 file.

> Code bị cắt giữa chừng? Gõ: "Gửi lại toàn bộ file trong 1 khối code, không lược."

---

## Prompt 1 – Bot `.vbs`

👉 **COPY KHUNG NÀY dán vào Gemini** (bấm nút copy góc phải khung):

```
[VAI TRÒ] Bạn là lập trình viên VBScript chuyên SAP GUI Scripting.
[BỐI CẢNH] Tôi là cán bộ thuế, dùng TMS (SAP GUI) trên máy cơ quan. Máy KHÔNG cài được Python, KHÔNG chạy macro VBA. Chỉ có: wscript (VBScript), PowerShell 5, Excel.
[KỊCH BẢN GHI] Đây là kịch bản tôi ghi bằng Script Recording khi tra tổng quan 1 MST trên màn FMCACOV. MST mẫu trong kịch bản là 0100000000:
"""
(dán toàn bộ nội dung file .vbs vừa ghi)
"""
[VIỆC CẦN LÀM] Viết file BOT_TONG_QUAN.vbs chạy cho NHIỀU MST:
1. Đọc MST_CHAY.txt cùng thư mục: mỗi dòng 1 MST, bỏ dòng trống, bỏ MST trùng. MST hợp lệ: 10 số, 10 số-3 số, hoặc 12 số.
2. Kết nối SAP GUI đang mở bằng GetObject("SAPGUI"), dùng phiên đầu tiên. Chưa mở SAP hoặc chưa bật scripting thì báo lỗi dễ hiểu rồi thoát.
3. Mỗi MST: làm lại đúng các bước trong kịch bản, thay MST mẫu bằng MST đang chạy. Chờ SAP xử lý xong (session.Busy) trước khi đọc, chờ tối đa 30 giây.
4. Đọc lưới kết quả (GuiGridView): lấy HẾT dòng (cuộn firstVisibleRow để tải đủ) và hết cột.
5. Ghi KET_QUA_TONG_QUAN.xlsx: sheet "Ket qua" (cột A = MST, sau đó các cột của lưới), sheet "Tien do" (STT, MST, tên NNT, trạng thái OK / KHONG CO DU LIEU / LOI + chữ ở sbar). Lưu file sau MỖI MST để dừng giữa chừng vẫn còn dữ liệu.
6. Gặp popup lạ (wnd[1], wnd[2]...): đóng bằng sendVKey 12, quay về FMCACOV, ghi LOI cho MST đó rồi chạy tiếp MST sau, không dừng cả bot.
7. Trước mỗi MST: thấy file STOP_TONG_QUAN.txt thì lưu Excel và dừng.
8. Ghi LOG_TONG_QUAN.txt (giờ, MST, bước, lỗi) và PROGRESS_TONG_QUAN.txt (dòng/tổng|MST|trạng thái) để app đọc tiến độ.
[RÀNG BUỘC] Toàn bộ code chỉ dùng ký tự ASCII: không dấu tiếng Việt, kể cả chú thích và MsgBox. Có Option Explicit. Mỗi hàm có 1 dòng chú thích. Không dùng thư viện ngoài.
[TRẢ VỀ] Một khối code HOÀN CHỈNH, không lược đoạn nào. Sau code, liệt kê các ID findById tôi cần kiểm tra lại trên máy mình.
```

## Prompt 2 – App giao diện + 2 file `.bat`

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
- Tạo thư mục `BOT_TONG_QUAN` trên Desktop: chuột phải ra màn hình › New › Folder › gõ tên › Enter.

**Lưu từng file (làm 4 lần)**

1. Bấm phím **Windows**, gõ `notepad`, bấm **Enter**.
2. Sang Gemini, ở góc trên bên phải khối code, bấm nút **Copy** (biểu tượng 2 ô vuông). Đừng bôi đen bằng chuột.
3. Quay lại Notepad, bấm **Ctrl+V**. Kéo xuống cuối xem code còn đủ.
4. Bấm **Ctrl+Shift+S** (File › Save as). Bên trái chọn **Desktop**, bấm đúp mở thư mục `BOT_TONG_QUAN`.
5. Ô **Save as type**: chọn **All files (\*.\*)**.
6. Ô **File name**: gõ đúng tên theo bảng dưới.
7. Ô **Encoding** (cạnh nút Save): chọn theo bảng, rồi bấm **Save**.
8. Bấm **Ctrl+N** mở trang trắng, làm lại từ bước 2 cho file tiếp theo.

| Gõ vào ô File name | Lấy code từ | Ô Encoding chọn |
|---|---|---|
| `BOT_TONG_QUAN.vbs` | Prompt 1 | ANSI |
| `APP_TONG_QUAN.ps1` | Prompt 2 – khối A | UTF-8 with BOM |
| `CHAY.bat` | Prompt 2 – khối B | ANSI |
| `DUNG.bat` | Prompt 2 – khối C | ANSI |

**Kiểm tra:** thư mục có đúng 4 file, đuôi `.vbs` `.ps1` `.bat`, không có `.txt`.
- Lỡ có `.txt`: chuột phải file › **Rename** › xóa `.txt` ở cuối › Enter › Yes.
- Lỡ sai encoding: mở lại file bằng Notepad › Ctrl+Shift+S › chọn lại Encoding › Save › Yes để ghi đè.

## Chạy thử với 3–5 MST

1. Chuột phải từng file › **Properties** › nếu dưới cùng có ô **Unblock** thì tick › **OK**.
2. Mở SAP, đăng nhập TMS, về màn hình chính.
3. Bấm đúp **CHAY.bat** (không bấm đúp file `.vbs`). Cửa sổ app hiện ra.
4. Copy 3–5 MST thử (mỗi dòng 1 MST, có 1 MST sai), bấm **Dan clipboard** hoặc Ctrl+V vào ô danh sách.
5. Bấm **LAY TONG QUAN** › **Yes**. Không đụng chuột vào SAP và Excel khi bot chạy.
6. Xong, mở `KET_QUA_TONG_QUAN.xlsx` trong thư mục, so 1 MST với tra tay trên FMCACOV.

## Prompt sửa lỗi

👉 **COPY KHUNG NÀY dán vào Gemini** khi bot báo lỗi (điền vào chỗ trống trước khi gửi):

```
Bot của tôi bị lỗi khi chạy. Thông tin:
- Thông báo lỗi (gõ lại nguyên văn): Dòng ___, Ký tự ___, Lỗi: ______________
- Màn hình SAP lúc lỗi đang ở: (tên màn hình / popup đang hiện)
- 20 dòng cuối LOG_TONG_QUAN.txt (đã thay MST thật bằng MST mẫu):
"""
(dán log)
"""
Hãy:
1) Giải thích nguyên nhân trong 2–3 câu.
2) Chỉ sửa hàm liên quan, trả về NGUYÊN hàm đã sửa, không đổi phần khác.
3) Nói rõ tôi thay đoạn đó vào chỗ nào trong file.
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
