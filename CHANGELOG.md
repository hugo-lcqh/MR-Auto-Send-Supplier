# Changelog

Mọi thay đổi đáng chú ý của dự án được ghi lại tại đây. Định dạng dựa trên [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) và dự án dùng [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [1.3.0] - 2026-08-13

### Added

- Launcher WinForms dạng Material Request Control Center với bố cục responsive, DPI scaling, keyboard navigation và accessibility metadata.
- Trạng thái tác vụ trực tiếp trên launcher và terminal progress cho Prepare, Send, Scan và Remind.
- Bảng MR trong nội dung email với tùy chọn ẩn bảng.
- Bộ cài Inno Setup kiểm tra Excel/Outlook desktop và giữ lại dữ liệu người dùng khi nâng cấp hoặc gỡ cài đặt.
- UI, version, attachment-column và regression test suites.

### Changed

- Chạy nhiều supplier trong một PowerShell worker thay vì mở một process cho từng supplier.
- Supplier attachment chỉ hiển thị đúng 17 cột vận hành; các cột còn lại được ẩn.
- Định dạng ngày ETA và Material Need By Date rõ ràng, nhất quán trong Excel và email.
- Reminder trả lời đúng email thread đã gửi thay vì tạo conversation mới.
- Master output được publish qua staging để tránh mất file hiện có khi ghi thất bại.

### Fixed

- Không ghi preview email/reminder vào tracking log như một lần gửi thật.
- Tách index vòng lặp attachment/header để tránh cập nhật sai khi scan reply.
- Cảnh báo rõ khi master bị khóa hoặc supplier xử lý thất bại.
- Đồng bộ version `1.3.0` giữa launcher, installer và release notes.

### Security

- Tắt macro khi Excel mở workbook qua automation.
- Validate danh sách To/CC/MC trước khi tạo email.
- Loại dữ liệu supplier, input, output, log và binary release khỏi version control; installer dùng cấu hình mẫu `example.com`.

[Unreleased]: https://github.com/hugo-lcqh/MR-Auto-Send-Supplier/compare/v1.3.0...HEAD
[1.3.0]: https://github.com/hugo-lcqh/MR-Auto-Send-Supplier/releases/tag/v1.3.0
