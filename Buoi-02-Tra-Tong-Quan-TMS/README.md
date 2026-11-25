# Buổi 2 – Tra tổng quan TMS trên web Gemini (1 prompt)

Giảng viên: Ngô Hoàng Thuỳ Hương · Zalo 0328 878 297

Slide bài giảng: [Buoi2_Slide.pdf](Buoi2_Slide.pdf) · Quay về [mục lục khóa học](../README.md)

Làm trên trang **[gemini.google.com](https://gemini.google.com)**. Không cài thêm phần mềm.

> ⚠️ **Tắt dịch tự động của trình duyệt** (bấm biểu tượng 文A trên thanh địa chỉ › *Hiển thị bản gốc*). Trang này viết sẵn tiếng Việt.

Mục tiêu: ghi tay 1 mã số thuế trên FMCACOV, dán **một prompt** vào Gemini, có cửa sổ để dán nhiều mã rồi bấm Chạy. Không cần biết lập trình.

> **Bảo mật:** không bao giờ dán user hoặc mật khẩu TMS lên Gemini.

---

## Làm 4 việc

### 1. Bật Scripting (một lần)

SAP đang mở, đã đăng nhập TMS.

- Alt+F12 › **Options** › **Accessibility & Scripting** › **Scripting**.
- Tick **Enable scripting**. Bỏ tick hai ô **Notify when…**. OK.

### 2. Ghi 1 mã số thuế

- Alt+F12 › **Script Recording and Playback…** › **More** (lưu ra Desktop, tên `ghi.vbs`) › **Record**.
- Tra tổng quan đúng 1 MST trên FMCACOV, đến khi thấy bảng kết quả › **Stop**.
- Không bấm đúp `ghi.vbs`. Chuột phải file › **Open with** › **Notepad** › Ctrl+A › Ctrl+C.

### 3. Gửi Gemini — một prompt

Vào [gemini.google.com](https://gemini.google.com) › đăng nhập › **New chat**. Giữ nguyên cuộc trò chuyện này đến hết buổi.

1. Ctrl+V để dán kịch bản. **Chưa Gửi.**
2. **Shift+Enter** hai lần (Enter là gửi luôn).
3. Copy khung prompt bên dưới › Ctrl+V › **Gửi**.

👉 **COPY KHUNG NÀY, dán ngay dưới kịch bản:**

```
Tôi dán ở trên kịch bản SAP lúc tra 1 mã số thuế trên màn hình FMCACOV.
Mã số thuế lúc ghi nằm ở dòng .text = "..."

Viết đúng 3 file. Chữ trong code và trên cửa sổ không dấu.

1. BOT.vbs
Đọc file MST.txt cùng thư mục, mỗi dòng 1 mã số thuế, bỏ dòng trống.
Nối SAP đang mở bằng GetObject("SAPGUI"), dùng phiên đầu tiên. Chưa mở SAP thì báo "Hay mo SAP truoc" rồi dừng.
Với mỗi mã, làm lại đúng các bước trong kịch bản, chỉ đổi mã số thuế.
Chờ SAP xử lý xong rồi mới đọc bảng, chờ tối đa 30 giây. Đọc hết dòng, hết cột.
Ghi file KET_QUA.csv. Cột đầu là mã số thuế. Lưu sau mỗi mã.
Mã lỗi hoặc không có dữ liệu: ghi Loi, chạy mã tiếp theo, không dừng cả bot.
Có hộp thoại lạ thì đóng lại rồi chạy tiếp.
Thấy file DUNG.txt thì lưu csv và dừng.

2. MAN_HINH.ps1
Một cửa sổ:
- Ô để dán nhiều mã số thuế
- Nút Chay: lưu nội dung ô thành MST.txt, xóa DUNG.txt nếu có, rồi chạy BOT.vbs
- Nút Dung: tạo file DUNG.txt

3. MO.bat
Mở cửa sổ bằng lệnh:
powershell -NoProfile -ExecutionPolicy Bypass -STA -File MAN_HINH.ps1

Trả về 3 khối code riêng. Mỗi khối ghi tên file ở dòng đầu. Không cắt bớt.
Nếu tôi gửi tiếp một câu báo lỗi, sửa rồi gửi lại đủ 3 file, mỗi file một khối code.
```

Code bị cắt giữa chừng thì gõ: `Gửi lại đủ 3 file, mỗi file một khối code.`

### 4. Lưu 3 file và chạy

Tạo thư mục **BOT** trên Desktop (chuột phải màn hình › New › Folder).

Windows 11: mở thư mục › **View** › **Show** › tick **File name extensions**.

Mỗi khối code làm một lần:

1. Bấm **Copy** ở góc khối code trên Gemini (đừng bôi đen bằng chuột).
2. Mở Notepad › Ctrl+V. Kéo xuống cuối xem code còn đủ.
3. Ctrl+Shift+S › mở thư mục **BOT**.
4. **Save as type:** All files (\*.\*).
5. **Encoding:** ANSI.
6. **File name:** gõ đúng tên trong bảng › Save.

| Gõ vào ô File name | Khối code |
|---|---|
| `BOT.vbs` | file 1 |
| `MAN_HINH.ps1` | file 2 |
| `MO.bat` | file 3 |

Thư mục **BOT** có đúng 3 file, không có đuôi `.txt` thừa. Lỡ thành `BOT.vbs.txt`: chuột phải › **Rename** › xóa `.txt` › Enter.

Chạy thử:

1. SAP vẫn mở, đang ở màn hình chính.
2. Bấm đúp **MO.bat**.
3. Dán 3 mã số thuế vào ô (mỗi dòng 1 mã, thêm 1 mã cố ý gõ sai) › bấm **Chay**.
4. Không đụng chuột vào SAP lúc bot đang chạy.
5. Xong, mở **KET_QUA.csv** bằng Excel. So 1 mã với tra tay trên FMCACOV.

Muốn dừng giữa chừng: bấm **Dung**. File csv vẫn còn dữ liệu đã chạy.

## Khi bị lỗi

Ở **đúng cuộc trò chuyện** vừa tạo bot, dán nguyên câu báo lỗi trên màn hình rồi Gửi. Prompt lúc nãy đã dặn Gemini sửa và gửi lại đủ 3 file. Lưu đè 3 file (ANSI, All files) rồi bấm đúp **MO.bat** chạy lại.

| Thấy gì | Làm gì |
|---|---|
| Hay mo SAP truoc | Mở SAP, đăng nhập, chạy lại |
| The control could not be found | Ghi lại đúng bước đó, dán kịch bản mới vào cùng cuộc trò chuyện và nhắn: thay ID cũ bằng ID trong kịch bản mới, gửi lại đủ 3 file |
| Invalid character | Nhắn: viết lại đủ 3 file, chỉ dùng chữ không dấu |
| Không thấy KET_QUA.csv hoặc file đang mở | Đóng Excel, chạy lại |

Đạt buổi 2 khi có file csv đủ các mã thử, mã gõ sai vẫn không làm dừng các mã còn lại, và 1 mã khớp với tra tay.

**Mình ghi. Gemini viết. Mình lưu và chạy.**
