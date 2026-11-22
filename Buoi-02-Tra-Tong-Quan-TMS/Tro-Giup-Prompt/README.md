# Bộ trợ giúp + Trợ giúp prompt – Buổi 2

Dùng khi bot của bạn không chạy, hoặc muốn có khung chuẩn để Gemini sửa cho nhanh.
Quay về [hướng dẫn Buổi 2](../README.md) · [mục lục khóa học](../../README.md)

## Có gì trong này

| File | Dùng để |
|---|---|
| [Bo_Tro_Giup_Buoi2.zip](Bo_Tro_Giup_Buoi2.zip) | **Tải 1 lần là đủ** – chứa cả 5 file bên dưới |
| **Bo-Tro-Giup/BOT_TONG_QUAN.vbs** | Bot mẫu tra tổng quan FMCACOV, có ghi log từng bước |
| **Bo-Tro-Giup/APP_TONG_QUAN.ps1** | Cửa sổ app: dán MST, nút chạy, nút dừng |
| **Bo-Tro-Giup/CHAY.bat** | Bấm đúp để mở app |
| **Bo-Tro-Giup/DUNG.bat** | Bấm đúp để dừng bot |
| **Bo-Tro-Giup/MST_MAU.txt** | Ví dụ cách ghi danh sách MST |
| [TRO_GIUP_PROMPT.md](TRO_GIUP_PROMPT.md) | **Prompt chuẩn** cho Gemini: sửa bot mẫu, sửa lỗi bằng log, nâng cấp |

## Chạy bộ trợ giúp – từng bước

1. Bấm vào **Bo_Tro_Giup_Buoi2.zip** ở bảng trên › bấm nút tải (mũi tên ↓ góc phải).
2. Mở thư mục Downloads, chuột phải file zip › **Extract All** (Giải nén tất cả) › chọn **Desktop** › **Extract**.
3. Mở thư mục vừa giải nén. Chuột phải từng file › **Properties** › nếu dưới cùng có ô **Unblock** thì tick › **OK**.
4. Mở SAP, đăng nhập TMS, bật Scripting (Alt+F12 › Options › Accessibility & Scripting › Scripting › tick *Enable scripting*).
5. Bấm đúp **CHAY.bat**. Cửa sổ app hiện ra.
6. Dán 3–5 MST thật để thử (mỗi dòng 1 MST), bấm **LAY TONG QUAN** › **Yes**.
7. Không đụng chuột vào SAP và Excel khi bot chạy. Muốn dừng: bấm **DUNG BOT**.
8. Xong, mở **KET_QUA_TONG_QUAN.xlsx** trong thư mục: sheet *Ket qua* là dữ liệu, sheet *Tien do* là trạng thái từng MST.

## Khi bot lỗi

Mở **LOG_TONG_QUAN.txt** bằng Notepad:
- Dòng bN cuối cùng của MST lỗi = bot chết ở bước đó.
- Các dòng !!! = mã lỗi, màn hình đang đứng, popup đang mở, SAP báo gì.

Copy đoạn log đó (thay MST thật bằng **0100000000**) và dùng **Prompt 2** trong [TRO_GIUP_PROMPT.md](TRO_GIUP_PROMPT.md).

> Bot mẫu viết theo màn FMCACOV trên máy giảng viên. Nếu máy bạn ID màn hình khác, log sẽ báo ở bước **b2** hoặc **b3**: ghi lại kịch bản và dùng **Prompt 1** để Gemini sửa đúng chỗ.
