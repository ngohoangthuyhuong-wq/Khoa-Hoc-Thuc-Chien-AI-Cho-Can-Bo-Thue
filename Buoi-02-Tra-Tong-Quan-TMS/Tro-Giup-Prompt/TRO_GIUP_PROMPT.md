# TRỢ GIÚP PROMPT – Buổi 2

Bộ prompt chuẩn để dùng với Gemini, đi kèm **bộ cứu hộ** (4 file mẫu trong thư mục [`Bo-Cuu-Ho`](Bo-Cuu-Ho)).
Quay về [hướng dẫn Buổi 2](../README.md) · [mục lục khóa học](../../README.md)

> **Bảo mật:** trước khi dán bất cứ thứ gì lên Gemini, thay MST thật bằng `0100000000`. Không dán tên, số tiền của NNT, user hay mật khẩu TMS.

---

## Vì sao dùng bộ cứu hộ thay vì bảo Gemini viết từ đầu

Bot mẫu đã có sẵn mọi phần chung: kết nối SAP, đọc MST, ghi Excel, nút dừng, **ghi log từng bước**.
Gemini chỉ phải sửa **đúng 2 chỗ** có đánh dấu `>>> SUA O DAY <<<`:

1. Hằng số `MAN_HINH` – mã màn hình TMS.
2. `Sub CacBuocTrenTMS` – các bước thao tác cho 1 MST, chép từ kịch bản ghi.

Phần Gemini tự viết rất ít, nên ít sai và cả lớp có chung một khung, giảng viên dễ hỗ trợ.

---

## Prompt 1 – Sửa bot mẫu theo kịch bản ghi (dùng cái này trước)

Dán vào Gemini, rồi dán **2 thứ** vào đúng chỗ trong ngoặc: nội dung file `BOT_TONG_QUAN.vbs` của bộ cứu hộ và kịch bản bạn vừa ghi.

```
[VAI TRÒ] Bạn là lập trình viên VBScript chuyên SAP GUI Scripting, hướng dẫn cho người mới.
[BỐI CẢNH] Tôi là cán bộ thuế, dùng TMS (SAP GUI) trên máy cơ quan. Máy chỉ chạy được wscript (VBScript), PowerShell 5, Excel. Không có Python, không chạy macro VBA.
[FILE MẪU] Đây là bot mẫu đang chạy được. KHÔNG viết lại file này, chỉ sửa đúng 2 chỗ có dòng ">>> SUA O DAY <<<":
"""
(dán toàn bộ nội dung file BOT_TONG_QUAN.vbs của bộ cứu hộ)
"""
[KỊCH BẢN GHI] Đây là kịch bản tôi ghi bằng Script Recording khi làm tay 1 MST. MST mẫu trong kịch bản là 0100000000:
"""
(dán kịch bản ghi)
"""
[VIỆC CẦN LÀM]
1. Sửa hằng số MAN_HINH thành mã màn hình trong kịch bản ghi (dòng có okcd).
2. Viết lại Sub CacBuocTrenTMS(mst) bằng các bước trong kịch bản ghi, bỏ dòng gõ mã màn hình và các dòng kết nối SAP ở đầu kịch bản. Thay MST mẫu bằng biến mst.
3. Sau mỗi lần sendVKey, press hoặc select: gọi ChoSap.
4. TRƯỚC mỗi thao tác: thêm 1 dòng Ghi "  bN: <việc đang làm, không dấu>", N tăng dần từ b1.
5. ID nào dài lặp lại nhiều lần thì đưa lên thành hằng số Const ở đầu file.
[RÀNG BUỘC] Chỉ dùng ký tự ASCII (không dấu tiếng Việt), kể cả chú thích. Không dùng On Error trong CacBuocTrenTMS. Giữ nguyên mọi phần khác của file mẫu.
[TRẢ VỀ] Chỉ 2 khối code: (1) dòng Const MAN_HINH mới cùng các Const ID mới nếu có, (2) nguyên Sub CacBuocTrenTMS mới. Sau đó nói rõ tôi thay mỗi khối vào chỗ nào trong file.
```

**Cách ráp:** mở `BOT_TONG_QUAN.vbs` bằng Notepad (chuột phải › Open with › Notepad), bôi đen đúng đoạn cũ, dán đoạn mới đè lên, **Ctrl+S** lưu (encoding giữ ANSI).

---

## Prompt 2 – Sửa lỗi bằng log (quan trọng nhất)

Bot mẫu ghi log từng bước vào `LOG_TONG_QUAN.txt`. Khi lỗi, log trông như sau:

```
21/11/2026 15:02:10  MST 1/5 = 0100000000
21/11/2026 15:02:10    b1: bam F4 mo o tim MST
21/11/2026 15:02:11    b2: chon tab tim kiem trong popup
21/11/2026 15:02:11    !!! LOI so 619: The control could not be found by id.
21/11/2026 15:02:11    !!! Man hinh: FMCACOV | Cua so tren cung: wnd[1] = Tim doi tuong
21/11/2026 15:02:11    !!! Thanh trang thai:
```

- Dòng **`bN` cuối cùng** cho biết bot chết ở bước nào (ở đây là b2).
- Ba dòng **`!!!`** cho biết mã lỗi, đang ở màn hình nào, popup nào đang mở, SAP báo gì.

Mở log bằng Notepad, copy **từ dòng `MST ...` của MST bị lỗi đến hết**, thay MST thật bằng MST mẫu, rồi dán vào prompt:

```
Bot của tôi bị lỗi. Đây là đoạn log của MST bị lỗi (đã thay MST thật):
"""
(dán log)
"""
Đây là Sub CacBuocTrenTMS hiện tại:
"""
(dán Sub CacBuocTrenTMS)
"""
Hãy:
1) Nhìn dòng bN cuối cùng và các dòng !!! để nói bot chết ở bước nào, vì sao (2–3 câu, dễ hiểu).
2) Chỉ sửa đúng Sub hoặc hằng số liên quan, trả về NGUYÊN đoạn đã sửa. Không viết lại cả file.
3) Giữ các dòng Ghi "  bN: ..." và giữ chỉ dùng ASCII.
4) Nếu cần tôi ghi lại kịch bản cho riêng bước đó, hãy nói rõ ghi từ màn hình nào đến màn hình nào.
```

> Lỗi `619 – The control could not be found by id` gần như luôn là **ID trên máy mình khác kịch bản** hoặc **có popup bất ngờ**. Cách chắc nhất: ghi lại kịch bản riêng cho bước đó rồi dán thêm vào prompt.

---

## Prompt 3 – Thêm log cho bot tự viết (nếu không dùng bot mẫu)

```
Thêm vào bot của tôi cách ghi log giống mẫu sau, không đổi phần xử lý:
- Sub Ghi(s): ghi thêm 1 dòng "ngày giờ  s" vào LOG_TONG_QUAN.txt cùng thư mục.
- Trước MỖI thao tác SAP: Ghi "  bN: <việc đang làm, không dấu>", N tăng dần.
- Đầu mỗi MST: Ghi "MST i/tổng = <mst>".
- Khi lỗi: Ghi 3 dòng bắt đầu bằng "  !!! ": số lỗi + mô tả; session.Info.Transaction + tiêu đề cửa sổ wnd trên cùng; chữ ở thanh trạng thái wnd[0]/sbar.
Chỉ dùng ASCII. Trả về các Sub mới và nói rõ chèn dòng Ghi vào những chỗ nào.
```

---

## Prompt 4 – Nâng cấp từng bước

Mỗi lần chỉ xin **1** tính năng, chạy thử xong mới xin tiếp. Luôn dán kèm file hiện tại.

```
Đây là bot hiện tại đang chạy tốt:
"""
(dán file .vbs)
"""
Thêm đúng 1 tính năng: <mô tả, ví dụ: chỉ lấy dòng có tiểu mục 1601 và 1602>.
Giữ nguyên mọi phần khác, giữ log bN, chỉ dùng ASCII. Trả về những Sub bị sửa hoặc thêm mới, nói rõ chỗ chèn.
```

Gợi ý tính năng: lọc theo tiểu mục · sheet Tổng hợp cộng số tiền theo MST · tách sheet theo cơ quan thuế · nhớ khoảng dòng để hôm sau chạy tiếp.

---

## Dành cho giảng viên – Hướng dẫn cài Gem "Trợ giúp bot TMS"

Tạo Gem (gemini.google.com › Explore Gems › New Gem), dán đoạn dưới vào ô **Instructions**, đính kèm file `BOT_TONG_QUAN.vbs` mẫu ở phần **Knowledge**, rồi **Share › Copy link** gửi lớp. Học viên mở link chỉ cần dán kịch bản ghi hoặc log lỗi.

```
Bạn là trợ lý giúp cán bộ thuế (người mới, không biết lập trình) làm bot VBScript chạy trên TMS (SAP GUI).
Luôn dựa trên file BOT_TONG_QUAN.vbs trong Knowledge. Không bao giờ viết lại cả file; chỉ sửa hằng số MAN_HINH, các Const ID và Sub CacBuocTrenTMS, hoặc đúng Sub bị lỗi.
Máy người dùng chỉ có wscript, PowerShell 5, Excel; không Python, không macro VBA.
Code chỉ dùng ký tự ASCII. Trước mỗi thao tác SAP có dòng Ghi "  bN: ...". Sau mỗi sendVKey/press/select gọi ChoSap. Không dùng On Error trong CacBuocTrenTMS.
Khi người dùng dán kịch bản ghi: làm theo Prompt 1 (trả 2 khối code + chỉ chỗ thay).
Khi người dùng dán log lỗi: đọc dòng bN cuối và các dòng !!!, giải thích 2–3 câu dễ hiểu, trả đúng đoạn sửa.
Nếu thấy MST, tên người nộp thuế, số tiền thật hoặc mật khẩu: nhắc người dùng xóa trước khi gửi tiếp.
Trả lời tiếng Việt, ngắn, từng bước đánh số, chỉ rõ bấm phím gì trong Notepad.
```
