# Contributing

MR Auto Send Supplier là dự án nội bộ. Mọi thay đổi cần giữ an toàn dữ liệu và khả năng chạy trên Windows PowerShell 5.1.

## Quy trình

1. Tạo branch ngắn từ `main`: `feature/<ten>`, `fix/<ten>` hoặc `chore/<ten>`.
2. Chỉ dùng `Config\suppliers.example.csv` và dữ liệu giả lập trong commit/test.
3. Thay đổi một concern trong mỗi commit, dùng message dạng `feat:`, `fix:`, `test:`, `docs:` hoặc `chore:`.
4. Chạy quality gate bên dưới.
5. Mở pull request, mô tả ảnh hưởng người dùng, cách test và rủi ro dữ liệu/email.

## Quality gate

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Launcher.UI.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Launcher.Version.Tests.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Outlook.Regression.Tests.ps1 -SupplierPath Config\suppliers.example.csv
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\MR-Attachment.Columns.Tests.ps1
```

## Checklist pull request

- [ ] Không có supplier/email/file MR thật trong diff.
- [ ] Mặc định **Display only, do not send** vẫn được giữ nguyên.
- [ ] Contract `-Mode`, `-VendorSelection`, `-ReplyFolder`, `-Display` và `-HideEmailTable` không bị thay đổi ngoài ý muốn.
- [ ] Version launcher, installer, release notes và changelog được đồng bộ nếu phát hành release.
- [ ] Tất cả test liên quan đã pass trên Windows PowerShell 5.1.

Không force-push `main`, không commit file trong `Release/`, `MR_Out/` hoặc thư mục input theo ngày.
