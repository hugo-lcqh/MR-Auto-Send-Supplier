# Security Policy

## Supported versions

| Version | Status |
|---|---|
| 1.3.x | Supported |
| < 1.3 | Unsupported |

## Báo cáo vấn đề bảo mật

Không đăng thông tin nhạy cảm trong issue hoặc pull request. Hãy dùng **GitHub Private Vulnerability Reporting** nếu tính năng này được bật, hoặc liên hệ trực tiếp repository owner qua kênh nội bộ.

Khi báo cáo, cung cấp:

- phiên bản ứng dụng;
- mô tả ảnh hưởng và cách tái hiện bằng dữ liệu giả;
- file/log đã loại bỏ email, supplier, mã nhân viên và nội dung MR thật;
- đề xuất giảm thiểu nếu có.

## Quy tắc dữ liệu

- `Config\suppliers.csv`, `Input\dd.MM.yyyy` và `MR_Out` là dữ liệu local, không được version-control.
- Release asset chỉ chứa chương trình, template và cấu hình mẫu dùng domain `example.com`.
- Không ghi password, token hoặc credential vào source, config, log hay GitHub Actions.
- Nếu credential từng được commit, phải revoke/rotate trước; xóa file hoặc rewrite history không làm credential cũ an toàn trở lại.
- Email thật hoặc file MR đã xuất hiện trong lịch sử Git cần được xử lý bằng một kế hoạch history rewrite riêng, có backup và thông báo cho mọi người dùng repo.
