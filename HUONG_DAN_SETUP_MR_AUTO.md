# Hướng dẫn cài đặt MR Auto Send Supplier

Tài liệu này áp dụng cho **MR Auto Send Supplier v1.3.0** trên Windows.

## 1. Điều kiện hệ thống

Máy sử dụng cần có:

- Windows PowerShell 5.1 tại `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`.
- Microsoft Excel desktop.
- Microsoft Outlook classic desktop đã đăng nhập email.
- Quyền Excel/Outlook COM automation.

> New Outlook không hỗ trợ COM automation. Nếu installer báo thiếu Outlook, hãy cài hoặc chuyển sang Outlook classic.

## 2. Cài bằng installer

1. Tải `MR-Auto-Setup-v1.3.0.exe` từ [GitHub Releases](../../releases/latest).
2. Đóng các phiên MR Auto đang chạy.
3. Chạy installer và đọc phần **What's New**.
4. Chọn tạo desktop shortcut nếu cần.
5. Sau khi cài, ứng dụng nằm tại:

```text
%LOCALAPPDATA%\MR Auto Send Supplier
```

Installer kiểm tra Excel và Outlook trước khi cài. Khi nâng cấp, các file sau được cập nhật:

- `Scripts\MR-Launcher.ps1`
- `Scripts\MR-Outlook.ps1`
- `Input\Template.xlsx`

Các dữ liệu người dùng sau được giữ nguyên:

- `Config\suppliers.csv`
- `Config\reply_folder.txt`
- toàn bộ `Input` và `MR_Out`

## 3. Cài thủ công từ source

Clone hoặc tải source rồi giữ nguyên cấu trúc:

```text
MR Auto Send Supplier
|-- Config
|   |-- suppliers.example.csv
|   `-- reply_folder.txt
|-- Input
|   `-- Template.xlsx
|-- MR_Out
|-- Scripts
|   |-- MR-Launcher.ps1
|   `-- MR-Outlook.ps1
`-- Tests
```

Tạo cấu hình local:

```powershell
Copy-Item .\Config\suppliers.example.csv .\Config\suppliers.csv
New-Item -ItemType Directory -Force .\MR_Out | Out-Null
```

Không đổi tên `Config`, `Input`, `MR_Out` hoặc `Scripts` nếu chưa cập nhật các đường dẫn trong script.

## 4. Cấu hình supplier

Mở `Config\suppliers.csv` và thay dữ liệu mẫu bằng dữ liệu nội bộ.

| Cột | Nội dung |
|---|---|
| `Keyword` | Alias dùng để match tên supplier; nhiều alias phân tách bằng `;` |
| `VendorCode` | Mã supplier |
| `VendorName` | Tên supplier chuẩn |
| `Email To` | Người nhận chính; nhiều email phân tách bằng `;` |
| `Email CC` | Danh sách CC |
| `MC` | MC phụ trách |
| `Buyer` | Buyer phụ trách |

Ứng dụng bỏ qua supplier có địa chỉ email không hợp lệ. Luôn chạy test với **Display only, do not send** sau khi thay đổi cấu hình.

## 5. Cấu hình thư mục Outlook

Giá trị mặc định trong `Config\reply_folder.txt` là:

```text
MR_REQUEST
```

Tạo thư mục `Inbox\MR_REQUEST` trong Outlook nếu muốn tách riêng reply. Khi thư mục không tồn tại, ứng dụng cảnh báo và quét Inbox.

## 6. Chuẩn bị input

Trong `Input`, tạo thư mục theo định dạng `dd.MM.yyyy`, ví dụ:

```text
Input\13.08.2026
```

Đặt file `.xlsx` hoặc `.xlsb` vào thư mục ngày mới nhất. File đầu vào cần:

- sheet `MR`;
- header nằm trong 100 dòng đầu;
- các cột chính như `Vendor Name`, `Part No.`, `Request Qty` và `Material Need By Date`.

Tác vụ Prepare tạo:

```text
Input\dd.MM.yyyy\MR_Master_Input.txt
Input\dd.MM.yyyy\MR_Master_Input.xlsx
```

## 7. Quy trình vận hành an toàn

1. Mở Excel và Outlook, xử lý mọi popup đăng nhập.
2. Mở **MR Auto Send Supplier**.
3. Giữ chọn **Display only, do not send**.
4. Chọn **Prepare Input**.
5. Chọn một supplier có ít dòng để kiểm tra trước.
6. Chọn **Send MR**.
7. Kiểm tra trong Outlook:
   - To/CC đúng;
   - subject có dạng `MR_REQUEST|Supplier|yyyyMMdd_HHmmss`;
   - bảng MR và attachment đúng supplier;
   - ngày và số lượng hiển thị đúng.
8. Chỉ bỏ chọn Display khi đã xác nhận toàn bộ nội dung.

## 8. Quét reply và reminder

- **Scan Replies** đọc attachment từ thư mục Outlook đã cấu hình, lưu bản sao vào `MR_Out\Replies` và cập nhật master.
- **Remind Pending** tạo reply trong đúng email thread cho các yêu cầu quá 24 giờ chưa có phản hồi.
- Khi chạy ở chế độ Display, email/reminder không được ghi nhận là đã gửi.

## 9. Chạy bằng command line

```powershell
# Launcher
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Launcher.ps1

# Prepare
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode prepare

# Preview Send
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode send -Display

# Scan Replies
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode scan -ReplyFolder MR_REQUEST

# Preview Reminder
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode remind -Display
```

## 10. Lỗi thường gặp

### Installer báo thiếu Excel hoặc Outlook

Xác nhận ứng dụng desktop đã được cài đúng kiến trúc và mở được thủ công. New Outlook không đáp ứng yêu cầu COM automation.

### Không đọc được input

Kiểm tra tên thư mục `dd.MM.yyyy`, sheet `MR`, các cột bắt buộc và bảo đảm file không bị khóa trong Excel.

### Không thấy supplier trong launcher

Chạy **Prepare Input**, sau đó kiểm tra `Vendor Name` trong master và mapping trong `Config\suppliers.csv`.

### Supplier bị bỏ qua khi Send

Kiểm tra `Email To`, `Email CC` và `MC`. Mỗi địa chỉ phải hợp lệ và được phân tách bằng `;`.

### Scan không cập nhật master

Đóng file master đang mở, kiểm tra subject reply, attachment Excel/CSV, trace columns và thư mục Outlook trong `reply_folder.txt`.

### Shortcut cũ không chạy

Gỡ shortcut cũ và chạy lại installer. Shortcut thủ công cần trỏ đến `powershell.exe` với tham số:

```text
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "<thu_muc_cai_dat>\Scripts\MR-Launcher.ps1"
```

## 11. Bảo vệ dữ liệu

- Không commit `Config\suppliers.csv`, `Input\dd.MM.yyyy` hoặc `MR_Out`.
- Không đưa file MR thật vào issue, pull request hoặc release asset.
- Chỉ phát hành installer/ZIP đã được kiểm tra và không chứa dữ liệu vận hành.
- Repo đã từng chứa dữ liệu mẫu vận hành cũ; khi cần loại bỏ khỏi toàn bộ lịch sử, hãy thực hiện một quy trình history rewrite riêng có phê duyệt và phối hợp với mọi người dùng repo.
