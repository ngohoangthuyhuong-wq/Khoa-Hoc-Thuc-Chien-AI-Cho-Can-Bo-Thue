# Buổi 3 – Tập prompt HTML

Giảng viên: Ngô Hoàng Thuỳ Hương · Zalo 0328 878 297

Slide bài giảng: [Buoi3_Slide.pdf](Buoi3_Slide.pdf) · Quay về [mục lục khóa học](../README.md)

Làm trên trang **[gemini.google.com](https://gemini.google.com)**. Không cài thêm phần mềm.

> ⚠️ **Tắt dịch tự động của trình duyệt** (bấm biểu tượng 文A trên thanh địa chỉ › *Hiển thị bản gốc*). Trang này viết sẵn tiếng Việt.

Buổi 3 tập viết prompt để Gemini làm ra **tool dạng trang HTML**: một file, bấm đúp là mở bằng Chrome, gửi qua Zalo cho đồng nghiệp là dùng được.

---

## Công thức prompt HTML: 5 phần

| Phần | Nói gì |
|---|---|
| 1. File gì, chạy thế nào | Một file HTML tên …, bấm đúp mở bằng Chrome, giao diện tiếng Việt |
| 2. Tool làm việc gì | Một câu: đưa gì vào, lấy gì ra |
| 3. Từng bước xử lý | Nói rõ từng bước, càng rõ Gemini làm càng đúng |
| 4. Dùng công cụ gì | Thư viện nào, lấy qua link CDN |
| 5. Trả code ra sao | Một khối code đầy đủ; báo lỗi thì sửa và gửi lại cả file |

---

## Bài 1 – Tool chuyển PDF sang Excel (đọc được cả file scan)

Kéo vào PDF có chữ, PDF scan hoặc ảnh chụp bảng (JPG, PNG). Tool tự nhận ra trang nào là scan để đọc bằng OCR tiếng Việt, rồi xuất ra Excel: mỗi file một sheet.

### 1. Gửi Gemini – một prompt

Vào [gemini.google.com](https://gemini.google.com) › đăng nhập › **New chat**. Giữ nguyên cuộc trò chuyện này đến hết bài.

Copy khung prompt bên dưới › Ctrl+V vào Gemini › **Gửi**.

👉 **COPY KHUNG NÀY:**

```
Viết cho tôi một file HTML duy nhất tên PDF_SANG_EXCEL.html.
Bấm đúp file là mở bằng Chrome và dùng được, không cài gì thêm.
Giao diện tiếng Việt có dấu. Đầu file có <meta charset="utf-8">.

Việc của tool: chuyển file PDF và ảnh chụp (JPG, PNG) sang Excel.
- Có ô kéo thả và nút Chọn file. Cho chọn nhiều file một lúc.
- PDF có chữ thật thì lấy chữ trực tiếp.
- PDF scan và ảnh chụp thì đọc bằng OCR:
  + Tesseract.js, ngôn ngữ ["vie", "eng"]. Chỉ tạo bộ đọc 1 lần rồi dùng lại.
  + Trang PDF vẽ ra ảnh phóng 2 lần rồi mới đọc.
  + Ảnh nhỏ thì phóng to cho cạnh dài khoảng 1800 điểm ảnh, nền trắng.
  + Bỏ chữ có độ tin cậy dưới 35.
  + Trong dãy số, chữ o hoặc O đọc nhầm thì đổi thành số 0.
- Giữ dạng bảng: chữ cùng một dòng là một hàng. Chữ sát nhau ghép thành một ô.
  Các ô thẳng mép trái với nhau là cùng một cột.
- Số tiền kiểu 1.234.567 chuyển thành số để cộng được.
  Mã số thuế, số có số 0 ở đầu, số ngắn như STT, tiểu mục thì giữ nguyên dạng chữ.
- Hiện thanh tiến độ: đang ở file nào, trang mấy.
- Đọc xong hiện bảng xem trước trên màn hình.
- Nút Tải Excel: ra file .xlsx, mỗi file là một sheet, cột đầu là số trang.

Dùng thư viện qua link CDN: pdf.js 3.11.174, Tesseract.js 5, SheetJS.
File đọc ngay trên máy, không gửi lên đâu.

Trả về một khối code đầy đủ, không cắt bớt.
Nếu tôi gửi tiếp một câu báo lỗi, sửa rồi gửi lại đủ cả file.
```

Code bị cắt giữa chừng thì gõ: `Gửi lại đủ cả file, một khối code.`

### 2. Lưu file

Windows 11: mở thư mục › **View** › **Show** › tick **File name extensions** (làm một lần).

1. Bấm **Copy** ở góc khối code trên Gemini (đừng bôi đen bằng chuột).
2. Mở Notepad › Ctrl+V. Kéo xuống cuối, thấy dòng **&lt;/html&gt;** là code đủ.
3. Ctrl+Shift+S › chọn **Desktop**.
4. **Save as type:** All files (\*.\*).
5. **Encoding:** **UTF-8** *(khác buổi 2: buổi 2 lưu ANSI)*.
6. **File name:** `PDF_SANG_EXCEL.html` › Save.

Lỡ thành `PDF_SANG_EXCEL.html.txt`: chuột phải › **Rename** › xóa `.txt` › Enter.

### 3. Chạy thử

1. Máy phải có mạng (lần đầu tool tải thư viện và bộ đọc tiếng Việt).
2. Bấm đúp **PDF_SANG_EXCEL.html**.
3. Kéo vào 1 PDF có chữ, 1 PDF scan, 1 ảnh chụp bảng. Chưa có file thì dùng file mẫu trong thư mục [Bai-1-PDF-Sang-Excel](Bai-1-PDF-Sang-Excel).
4. Chờ thanh tiến độ chạy hết, xem bảng trên màn hình.
5. Bấm **Tải Excel**, mở file ra. Dò lại 1 mã số thuế và 1 số tiền với bản gốc.

PDF có chữ đọc gần như ngay. Trang scan mất vài giây mỗi trang. Scan càng thẳng, càng rõ thì đọc càng đúng.

> 💡 Cách biết PDF có chữ hay scan: mở PDF, thử bôi đen một chữ. Bôi được là PDF có chữ. Tool tự nhận ra, không cần chọn.

### Khi bị lỗi

Ở **đúng cuộc trò chuyện** vừa tạo tool, dán nguyên câu báo lỗi rồi Gửi. Lưu đè file (UTF-8, All files) rồi bấm đúp chạy lại.

| Thấy gì | Làm gì |
|---|---|
| Chữ trên tool bị lỗi font (Ã, Æ…) | Lưu lại file, chọn Encoding **UTF-8** |
| Trang trắng, bấm nút không chạy | Kiểm tra mạng. Vẫn lỗi: bấm **F12** › tab **Console**, copy dòng chữ đỏ gửi Gemini |
| Đứng mãi ở "Đang tải bộ đọc tiếng Việt" | Lần đầu chờ 1–2 phút. Mạng cơ quan chặn thì thử mạng khác |
| Excel lệch cột, dính cột | Nhắn: `Cột Số tiền bị dính vào cột Tên, tách ra. Gửi lại cả file.` |
| Số tiền thành chữ, không cộng được | Nhắn: `Số kiểu 1.234.567 phải thành số. Gửi lại cả file.` |

Gemini sửa mãi không được: mở file mẫu [PDF_SANG_EXCEL.html](Bai-1-PDF-Sang-Excel/PDF_SANG_EXCEL.html) (bấm nút tải ở góc phải) để có bản chạy được, rồi tập tiếp phần nâng cấp.

### Tự nâng cấp – mỗi lần một câu

Mỗi lần chỉ thêm một ý, cuối câu luôn có **Gửi lại cả file**:

- `Thêm nút Gộp 1 sheet: dồn tất cả file vào một sheet, cột đầu là tên file. Gửi lại cả file.`
- `Bỏ những dòng chỉ có tiêu đề, chỉ giữ dòng có mã số thuế. Gửi lại cả file.`
- `Thêm dòng Tổng cộng cuối mỗi sheet cho các cột số tiền. Gửi lại cả file.`
- `Ghi tên tôi và số Zalo ở chân trang tool. Gửi lại cả file.`

### Đạt bài 1 khi

- Bấm đúp tool mở được, chữ tiếng Việt không lỗi.
- Đọc được cả 3 loại: PDF có chữ, PDF scan, ảnh chụp.
- Excel giữ mã số thuế có số 0 ở đầu, số tiền cộng được.
- Tự thêm được 1 nâng cấp bằng một câu nhắn Gemini.

**Mình nói rõ. Gemini viết. Mình lưu và chạy.**
