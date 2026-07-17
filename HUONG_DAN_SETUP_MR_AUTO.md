# Huong dan setup chuong trinh MR Auto

Tai lieu nay dung de copy va thiet lap chuong trinh **MR Auto Send Supplier** tren mot may tinh khac.

## 1. Dieu kien tren may moi

May moi can co:

- Windows.
- Microsoft Excel desktop.
- Microsoft Outlook desktop da dang nhap email.
- PowerShell Windows co san tai:
  `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
- Quyen cho phep Excel/Outlook automation bang COM. Neu Outlook/Excel hien canh bao lan dau, hay mo thu cong Excel va Outlook truoc, sau do chay lai chuong trinh.

## 2. Copy dung cau truc thu muc

Copy nguyen thu muc `MR Auto Send Supplier` sang may moi. Cau truc toi thieu can giu nhu sau:

```text
MR Auto Send Supplier
|-- Config
|   |-- suppliers.csv
|   `-- reply_folder.txt
|-- Input
|   |-- Template.xlsx
|   `-- dd.MM.yyyy
|       |-- file MR dau vao .xlsx/.xlsb
|-- MR_Out
|-- Scripts
|   |-- MR-Launcher.ps1
|   `-- MR-Outlook.ps1
`-- Test MR-Outlook.lnk
```

Khong doi ten cac thu muc `Config`, `Input`, `MR_Out`, `Scripts` neu khong sua lai script.

## 3. Chuan bi file cau hinh supplier

Mo file:

```text
Config\suppliers.csv
```

Kiem tra cac cot chinh:

- `Keyword`: tu khoa/alias de match ten supplier trong file MR. Neu co nhieu alias, cach nhau bang dau `;`.
- `VendorName`: ten supplier.
- `Email To`: danh sach email supplier nhan mail.
- `Email CC`: danh sach email CC.
- `MC`: email MC phu trach.

Luu y:

- Danh sach email trong mot o co the cach nhau bang dau `;`.
- Ten supplier trong input can match voi `VendorName` hoac alias trong `Keyword`.
- Nen test bang che do **Display only, do not send** truoc khi bo tick de gui that.

## 4. Chuan bi Outlook folder de scan reply

File cau hinh:

```text
Config\reply_folder.txt
```

Mac dinh dang la:

```text
MR_REQUEST
```

Trong Outlook, tao folder con duoi Inbox voi dung ten nay neu muon chuong trinh scan reply trong folder rieng:

```text
Inbox\MR_REQUEST
```

Neu folder khong ton tai, script se canh bao va scan Inbox thay the.

## 5. Chuan bi input MR

Trong thu muc `Input`, tao mot folder theo ngay voi dinh dang:

```text
dd.MM.yyyy
```

Vi du:

```text
Input\08.07.2026
```

Bo cac file MR Excel dau vao vao folder ngay nay. Script se tu chon folder ngay moi nhat theo dinh dang tren.

Yeu cau file dau vao:

- File Excel co sheet ten `MR`.
- Header nam trong 100 dong dau.
- Can co cac cot quan trong nhu `Vendor Name`, `Part No.`, `Shortage Qty` hoac `Request Qty`, va `Material Need By Date`.

Sau khi chay Prepare, chuong trinh se tao:

```text
Input\dd.MM.yyyy\MR_Master_Input.txt
Input\dd.MM.yyyy\MR_Master_Input.xlsx
```

## 6. Chay bang giao dien Launcher

Cach khuyen dung:

1. Double click `Test MR-Outlook.lnk`.
2. Giu tick **Display only, do not send** khi test.
3. Bam **Prepare Input** de tao master input va load danh sach supplier.
4. Chon supplier can xu ly.
5. Bam mot trong cac nut:
   - **Send MR**: tao file MR theo supplier va tao/gui email.
   - **Scan Replies**: doc reply trong Outlook va cap nhat feedback vao master input.
   - **Remind**: gui reminder cho request qua 24h chua reply.

Neu shortcut khong chay sau khi copy sang may moi, tao lai shortcut theo muc 7.

## 7. Tao lai shortcut neu duong dan bi sai

Shortcut cu co the dang tro ve duong dan cua may goc. Tren may moi, tao shortcut moi voi:

**Target:**

```text
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
```

**Arguments:**

```text
-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "DUONG_DAN_MOI\MR Auto Send Supplier\Scripts\MR-Launcher.ps1"
```

**Start in:**

```text
DUONG_DAN_MOI\MR Auto Send Supplier
```

Thay `DUONG_DAN_MOI` bang noi da copy thu muc tren may moi.

## 8. Lenh chay truc tiep khi can debug

Mo PowerShell tai thu muc `MR Auto Send Supplier`, sau do chay:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\MR-Launcher.ps1"
```

Chay prepare:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\MR-Outlook.ps1" -Mode prepare
```

Mo mail de xem truoc, khong gui:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\MR-Outlook.ps1" -Mode send -Display
```

Scan reply:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\MR-Outlook.ps1" -Mode scan -ReplyFolder "MR_REQUEST"
```

Gui reminder dang display, khong gui that:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Scripts\MR-Outlook.ps1" -Mode remind -Display
```

## 9. Output sau khi chay

Thu muc output chinh:

```text
MR_Out
```

File va folder thuong gap:

- `MR_Out\MR_<Supplier>_<yyyyMMdd_HHmmss>.xlsx`: file gui cho supplier.
- `MR_Out\mr_sent.csv`: log cac email MR da tao/gui, dung de scan reply va remind.
- `MR_Out\Replies\dd.MM-dd.MM.yyyy`: file reply attachment da luu theo tuan.

## 10. Quy trinh test an toan tren may moi

1. Mo Excel va Outlook thu cong, dam bao khong bi popup dang nhap.
2. Mo launcher.
3. Giu tick **Display only, do not send**.
4. Bam **Prepare Input**.
5. Chon 1 supplier it dong.
6. Bam **Send MR**.
7. Kiem tra email hien len trong Outlook:
   - To/CC dung.
   - Subject co dang `MR_REQUEST|Supplier|yyyyMMdd_HHmmss`.
   - Attachment dung supplier.
8. Neu tat ca OK, moi bo tick **Display only, do not send** de gui that.

## 11. Loi thuong gap

### Shortcut khong mo

Nguyen nhan thuong la shortcut dang tro ve duong dan may cu. Tao lai shortcut theo muc 7.

### Khong doc duoc input

Kiem tra:

- Folder input co dung format `dd.MM.yyyy` khong.
- File Excel co sheet `MR` khong.
- Header co cac cot bat buoc khong.
- File Excel co dang mo va bi lock khong.

### Khong thay supplier trong launcher

Chay **Prepare Input** truoc. Neu van khong thay, kiem tra cot `Vendor Name` trong input va file `Config\suppliers.csv`.

### Email gui sai hoac khong co email

Kiem tra `Config\suppliers.csv`:

- `VendorName` hoac `Keyword` co match voi ten supplier trong input khong.
- `Email To` co bi trong khong.
- Cac email co cach nhau bang `;` khong.

### Scan reply khong cap nhat data

Kiem tra:

- Reply email co subject dang reply/forward cua `MR_REQUEST|...` khong.
- Reply co attachment Excel khong.
- Attachment con cac cot `_SourceKey`, `Reply Qty`, `ETA Vendor can Supply`, `IF CAN NOT` khong.
- Reply co nam trong tuan hien tai khong.
- Folder Outlook trong `Config\reply_folder.txt` co dung khong.

## 12. Luu y khi chuyen cho nguoi khac

- Nen xoa file output cu trong `MR_Out` neu khong muon mang theo lich su test.
- Neu muon giu lich su da gui/remind/scan thi giu `MR_Out\mr_sent.csv`.
- Khong xoa `Input\Template.xlsx` vi file nay dung de tao master va format output.
- Khi test tren may moi, luon bat dau bang **Display only, do not send**.
