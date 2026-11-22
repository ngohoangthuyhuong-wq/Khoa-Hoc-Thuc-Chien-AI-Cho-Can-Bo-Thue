# TRỢ GIÚP PROMPT – Buổi 2

Bộ prompt chuẩn để dùng với Gemini, đi kèm **bộ trợ giúp** (4 file mẫu trong thư mục [**Bo-Tro-Giup**](Bo-Tro-Giup)).
Quay về [hướng dẫn Buổi 2](../README.md) · [mục lục khóa học](../../README.md)

> ⚠️ **Tắt dịch tự động của trình duyệt** (biểu tượng 文A trên thanh địa chỉ › *Hiển thị bản gốc*) – bản dịch máy làm sai chữ.
>
> 📖 Chữ thường là **hướng dẫn để đọc**. Thứ dán vào Gemini chỉ là **khung xám to có nút copy ở góc**, nằm ngay dưới dòng 👉 *COPY KHUNG NÀY*.

> **Bảo mật:** trước khi dán bất cứ thứ gì lên Gemini, thay MST thật bằng **0100000000**. Không dán tên, số tiền của NNT, user hay mật khẩu TMS.

## Nguyên tắc chung – nhớ 1 câu

**Copy cái nào, dán ngay cái đó. Dán tài liệu trước, dán prompt sau cùng, rồi Gửi.**

- Máy chỉ nhớ **1 thứ vừa copy**. Copy cái mới là cái cũ mất. Nên copy xong phải dán vào Gemini ngay.
- Trong ô chat Gemini: **Enter = Gửi luôn**. Muốn xuống dòng bấm **Shift+Enter**.
- Không phải bôi đen hay thay chữ gì trong prompt.

---

## Vì sao dùng bộ trợ giúp thay vì bảo Gemini viết từ đầu

Bot mẫu đã có sẵn mọi phần chung: kết nối SAP, đọc MST, ghi Excel, nút dừng, **ghi log từng bước**.
Gemini chỉ phải sửa **đúng 2 chỗ** có đánh dấu **>>> SUA O DAY <<<**:

1. Hằng số **MAN_HINH** – mã màn hình TMS.
2. **Sub CacBuocTrenTMS** – các bước thao tác cho 1 MST, chép từ kịch bản ghi.

Phần Gemini tự viết rất ít, nên ít sai và cả lớp có chung một khung, giảng viên dễ hỗ trợ.

---

## Prompt 1 – Sửa bot mẫu theo kịch bản ghi (dùng cái này trước)

| Bước | Làm gì |
|---|---|
| 1 | Mở **gemini.google.com** › **New chat**. |
| 2 | Chuột phải **BOT_TONG_QUAN.vbs** (trong thư mục bộ trợ giúp) › **Open with › Notepad** › **Ctrl+A**, **Ctrl+C** › đóng Notepad › về Gemini bấm vào ô chat › **Ctrl+V**. |
| 3 | **Shift+Enter** 2 lần. |
| 4 | Chuột phải file **kịch bản bạn vừa ghi** › **Open with › Notepad** › **Ctrl+H** thay MST thật bằng **0100000000** › **Ctrl+A**, **Ctrl+C** › đóng, bấm **Don't Save** › về Gemini **Ctrl+V**. |
| 5 | **Shift+Enter** 2 lần › copy **khung dưới đây** › về Gemini **Ctrl+V** › **Gửi**. |

👉 **COPY KHUNG NÀY dán vào Gemini, ngay dưới 2 phần đã dán:**

```
[VAI TRÒ] Bạn là lập trình viên VBScript chuyên SAP GUI Scripting, hướng dẫn cho người mới.
[BỐI CẢNH] Tôi là cán bộ thuế, dùng TMS (SAP GUI) trên máy cơ quan. Máy chỉ chạy được wscript (VBScript), PowerShell 5, Excel. Không có Python, không chạy macro VBA.
[TÀI LIỆU] Phía trên tin nhắn này tôi đã dán 2 phần:
- Phần 1 (dán trước): bot mẫu BOT_TONG_QUAN.vbs đang chạy được, bắt đầu bằng dòng "' ====". KHÔNG viết lại file này, chỉ sửa đúng 2 chỗ có dòng ">>> SUA O DAY <<<".
- Phần 2 (dán sau): kịch bản tôi ghi bằng Script Recording khi làm tay 1 MST, thường bắt đầu bằng "If Not IsObject(application)". MST mẫu trong kịch bản là 0100000000.
[VIỆC CẦN LÀM]
1. Sửa hằng số MAN_HINH thành mã màn hình trong kịch bản ghi (dòng có okcd).
2. Viết lại Sub CacBuocTrenTMS(mst) bằng các bước trong kịch bản ghi, bỏ dòng gõ mã màn hình và các dòng kết nối SAP ở đầu kịch bản. Thay MST mẫu bằng biến mst.
3. Sau mỗi lần sendVKey, press hoặc select: gọi ChoSap.
4. TRƯỚC mỗi thao tác: thêm 1 dòng Ghi "  bN: <việc đang làm, không dấu>", N tăng dần từ b1.
5. ID nào dài lặp lại nhiều lần thì đưa lên thành hằng số Const ở đầu file.
[RÀNG BUỘC] Chỉ dùng ký tự ASCII (không dấu tiếng Việt), kể cả chú thích. Không dùng On Error trong CacBuocTrenTMS. Giữ nguyên mọi phần khác của file mẫu.
[TRẢ VỀ] Chỉ 2 khối code: (1) dòng Const MAN_HINH mới cùng các Const ID mới nếu có, (2) nguyên Sub CacBuocTrenTMS mới. Sau đó hướng dẫn người mới từng bước: mở file bằng Notepad, tìm đoạn cũ ở đâu, bôi đen từ dòng nào đến dòng nào, dán đè, Ctrl+S.
```

---

## Prompt 2 – Sửa lỗi bằng log (quan trọng nhất)

Bot mẫu ghi log từng bước vào **LOG_TONG_QUAN.txt**. Khi lỗi, log trông như sau (đây chỉ là **ví dụ để nhìn**, không copy):

> 21/11/2026 15:02:10 · MST 1/5 = 0100000000
> 21/11/2026 15:02:10 · b1: bam F4 mo o tim MST
> 21/11/2026 15:02:11 · **b2: chon tab tim kiem trong popup** ← bước cuối, bot chết ở đây
> 21/11/2026 15:02:11 · !!! LOI so 619: The control could not be found by id.
> 21/11/2026 15:02:11 · !!! Man hinh: FMCACOV | Cua so tren cung: wnd[1] = Tim doi tuong
> 21/11/2026 15:02:11 · !!! Thanh trang thai:

- Dòng bN** cuối cùng** cho biết bot chết ở bước nào (ở đây là b2).
- Ba dòng !!! cho biết mã lỗi, đang ở màn hình nào, popup nào đang mở, SAP báo gì.

| Bước | Làm gì |
|---|---|
| 1 | Mở **đúng cuộc trò chuyện Gemini đã tạo bot** (cột trái Gemini có lịch sử). |
| 2 | Mở **LOG_TONG_QUAN.txt** bằng Notepad › **Ctrl+End** xuống cuối › bôi đen từ dòng **MST ...** của MST bị lỗi đến hết › **Ctrl+C** › về Gemini **Ctrl+V**. Sửa MST thật thành **0100000000**. |
| 3 | **Shift+Enter** 2 lần › copy **khung dưới đây** › **Ctrl+V** › **Gửi**. |

👉 **COPY KHUNG NÀY dán vào Gemini, ngay dưới đoạn log đã dán:**

```
Bot của tôi bị lỗi. Phần tôi dán ở PHÍA TRÊN là đoạn log của MST bị lỗi (đã thay MST thật). Bot là file BOT_TONG_QUAN.vbs bạn đã giúp tôi sửa trong cuộc trò chuyện này.
Hãy:
1) Nhìn dòng bN cuối và các dòng !!! để nói bot chết ở bước nào, vì sao (2–3 câu, dễ hiểu).
2) Chỉ sửa đúng Sub hoặc hằng số liên quan, trả về NGUYÊN đoạn đã sửa. Không viết lại cả file.
3) Giữ các dòng Ghi "  bN: ..." và giữ chỉ dùng ASCII.
4) Nếu cần tôi ghi lại kịch bản cho riêng bước đó, hãy nói rõ ghi từ màn hình nào đến màn hình nào.
```

> Lỗi **619 – The control could not be found by id** gần như luôn là **ID trên máy mình khác kịch bản** hoặc **có popup bất ngờ**. Cách chắc nhất: ghi lại kịch bản riêng cho bước đó, dán vào Gemini và nói "đây là kịch bản ghi lại bước bị lỗi".

---

## Prompt 3 – Thêm log cho bot tự viết (nếu không dùng bot mẫu)

| Bước | Làm gì |
|---|---|
| 1 | Mở file **.vbs** của bạn bằng Notepad › **Ctrl+A**, **Ctrl+C** › về Gemini **Ctrl+V**. |
| 2 | **Shift+Enter** 2 lần › copy **khung dưới đây** › **Ctrl+V** › **Gửi**. |

👉 **COPY KHUNG NÀY dán vào Gemini, ngay dưới file đã dán:**

```
Phía trên là bot VBScript của tôi. Thêm vào cách ghi log giống mẫu sau, không đổi phần xử lý:
- Sub Ghi(s): ghi thêm 1 dòng "ngày giờ  s" vào LOG_TONG_QUAN.txt cùng thư mục.
- Trước MỖI thao tác SAP: Ghi "  bN: <việc đang làm, không dấu>", N tăng dần.
- Đầu mỗi MST: Ghi "MST i/tổng = <mst>".
- Khi lỗi: Ghi 3 dòng bắt đầu bằng "  !!! ": số lỗi + mô tả; session.Info.Transaction + tiêu đề cửa sổ wnd trên cùng; chữ ở thanh trạng thái wnd[0]/sbar.
Chỉ dùng ASCII. Trả về NGUYÊN file đã sửa trong 1 khối code, không lược đoạn nào.
```

---

## Prompt 4 – Nâng cấp từng bước

Mỗi lần chỉ xin **1** tính năng, chạy thử xong mới xin tiếp.

| Bước | Làm gì |
|---|---|
| 1 | Mở file **.vbs** đang chạy tốt bằng Notepad › **Ctrl+A**, **Ctrl+C** › về Gemini **Ctrl+V**. |
| 2 | **Shift+Enter** 2 lần › copy **khung dưới đây** › **Ctrl+V** › gõ tính năng muốn thêm vào cuối › **Gửi**. |

👉 **COPY KHUNG NÀY dán vào Gemini, ngay dưới file đã dán:**

```
Phía trên là bot của tôi đang chạy tốt. Thêm đúng 1 tính năng, giữ nguyên mọi phần khác, giữ log bN, chỉ dùng ASCII. Trả về những Sub bị sửa hoặc thêm mới, hướng dẫn người mới chèn vào chỗ nào.
Tính năng cần thêm:
```

Gợi ý tính năng (gõ vào sau chữ "Tính năng cần thêm:"): chỉ lấy dòng có tiểu mục 1601 và 1602 · thêm sheet Tổng hợp cộng số tiền theo MST · tách sheet theo cơ quan thuế · nhớ khoảng dòng để hôm sau chạy tiếp.
