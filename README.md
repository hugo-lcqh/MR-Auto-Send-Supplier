# MR Auto Send Supplier

[![Version](https://img.shields.io/badge/version-1.3.0-0E6F9F)](../../releases/tag/v1.3.0)
![Platform](https://img.shields.io/badge/platform-Windows-0078D4)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1-5391FE)
[![CI](https://github.com/moi0329/MR-Auto-Send-Supplier/actions/workflows/ci.yml/badge.svg)](../../actions/workflows/ci.yml)

Ứng dụng Windows tự động hóa quy trình **Material Request (MR)**: chuẩn hóa dữ liệu Excel, tạo file riêng cho từng supplier, soạn/gửi email Outlook, quét file phản hồi và nhắc các yêu cầu quá hạn.

> [!IMPORTANT]
> Đây là công cụ nội bộ. Không commit danh bạ supplier, file MR đầu vào, email log hoặc file output thật. Repo chỉ lưu cấu hình mẫu dùng địa chỉ `example.com`.

## Điểm nổi bật

- Launcher WinForms theo dạng **Material Request Control Center**, hỗ trợ DPI, thay đổi kích thước, bàn phím và accessibility.
- Bốn tác vụ theo đúng luồng vận hành: **Prepare Input → Send MR → Scan Replies → Remind Pending**.
- Tạo attachment riêng cho supplier và chỉ hiển thị 17 cột cần thiết.
- Chèn bảng MR vào email; có thể tắt bằng tùy chọn **Show MR table in email**.
- Gửi nhiều supplier trong một worker, hiển thị tiến độ và kết quả trên terminal.
- Mặc định an toàn **Display only, do not send** để người dùng duyệt email trước khi gửi thật.
- Bộ cài Inno Setup giữ lại `Config`, `Input` và `MR_Out` khi nâng cấp hoặc gỡ cài đặt.

## Yêu cầu hệ thống

- Windows có **Windows PowerShell 5.1**.
- Microsoft Excel desktop.
- Microsoft Outlook classic desktop đã đăng nhập; New Outlook không hỗ trợ COM automation được ứng dụng sử dụng.
- Quyền đọc/ghi tại thư mục cài đặt của người dùng.

## Cài đặt

### Cách khuyến nghị

1. Mở [Latest release](../../releases/latest).
2. Tải `MR-Auto-Setup-v1.3.0.exe`.
3. Chạy installer. Bộ cài sẽ kiểm tra Excel và Outlook desktop trước khi tiếp tục.
4. Cập nhật `Config\suppliers.csv` bằng dữ liệu nội bộ hợp lệ. Khi nâng cấp, installer giữ nguyên file hiện có.
5. Khởi chạy **MR Auto Send Supplier** từ Start Menu hoặc desktop shortcut.

Xem hướng dẫn chi tiết tại [HUONG_DAN_SETUP_MR_AUTO.md](HUONG_DAN_SETUP_MR_AUTO.md).

### Chạy trực tiếp từ source

```powershell
Copy-Item .\Config\suppliers.example.csv .\Config\suppliers.csv
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Launcher.ps1
```

Thay dữ liệu mẫu bằng danh bạ nội bộ trước khi vận hành thật. Không commit `Config\suppliers.csv`.

## Sử dụng nhanh

1. Đặt file MR vào `Input\dd.MM.yyyy\`.
2. Mở launcher và giữ chọn **Display only, do not send**.
3. Chọn **Prepare Input** để tạo master và tải Supplier Queue.
4. Chọn supplier cần xử lý, sau đó chọn **Send MR**.
5. Kiểm tra To/CC, subject, nội dung và attachment trong Outlook.
6. Chỉ tắt chế độ Display khi đã xác nhận dữ liệu chính xác.
7. Dùng **Scan Replies** để cập nhật phản hồi và **Remind Pending** cho yêu cầu quá 24 giờ.

## Cấu hình

| File | Mục đích | Quản lý trên Git |
|---|---|---|
| `Config/suppliers.example.csv` | Schema và dữ liệu giả lập an toàn | Có |
| `Config/suppliers.csv` | Danh bạ supplier thật | Không |
| `Config/reply_folder.txt` | Tên thư mục Outlook dùng để quét reply | Có |
| `Input/Template.xlsx` | Template tạo master/attachment | Có |
| `Input/dd.MM.yyyy/*` | Dữ liệu MR vận hành | Không |
| `MR_Out/*` | Output, reply attachment và email tracking | Không |

Các cột bắt buộc trong `suppliers.csv`:

```text
Keyword, VendorCode, VendorName, Email To, Email CC, MC, Buyer
```

Nhiều alias hoặc địa chỉ email trong một ô được phân tách bằng dấu `;`.

## Lệnh chính

```powershell
# Chuẩn hóa input và tạo master
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode prepare

# Mở email để duyệt, không gửi
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode send -Display

# Quét attachment phản hồi
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode scan -ReplyFolder MR_REQUEST

# Mở reminder để duyệt, không gửi
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\MR-Outlook.ps1 -Mode remind -Display
```

## Cấu trúc repo

```text
Config/       Cấu hình mẫu và Outlook reply folder
Input/        Template được version-control; dữ liệu thật bị ignore
Installer/    Inno Setup script và release notes
Scripts/      Launcher UI và backend automation
Tests/        UI contract, version contract và regression tests
.github/      GitHub Actions workflow
```

## Kiểm thử

Chạy toàn bộ quality gate bằng Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Launcher.UI.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Launcher.Version.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Outlook.Regression.Tests.ps1 -SupplierPath Config\suppliers.example.csv
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Attachment.Columns.Tests.ps1
```

GitHub Actions chạy lại cùng các kiểm tra trên `windows-latest` cho mọi push và pull request vào `main`.

## Phiên bản và thay đổi

Dự án dùng [Semantic Versioning](https://semver.org/). Xem lịch sử thay đổi tại [CHANGELOG.md](CHANGELOG.md) và ghi chú cài đặt tại [Installer/ReleaseNotes.txt](Installer/ReleaseNotes.txt).

## Đóng góp và bảo mật

- Quy trình branch, commit và pull request: [CONTRIBUTING.md](CONTRIBUTING.md).
- Cách báo cáo vấn đề bảo mật và quy tắc xử lý dữ liệu: [SECURITY.md](SECURITY.md).

Maintainer: [@moi0329](https://github.com/moi0329) · Hugo Le Chi Quoc Hung
