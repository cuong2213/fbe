# Vehicle Defect Mapping – Hướng dẫn triển khai

Gồm 3 file code (đã thống nhất tên Sheet / Table / Shape / Cell):

| File | Dán vào đâu |
|---|---|
| `modVehicleMapping.bas` | VBA Editor → File → **Import File** (hoặc dán toàn bộ, bỏ dòng `Attribute VB_Name`) |
| `Sheet_MAP_VEHICLE.txt` | Sheet `MAP_VEHICLE` (Microsoft Excel Objects) |
| `ThisWorkbook.txt` | `ThisWorkbook` |

> Chuỗi thông báo trong code viết **tiếng Việt không dấu** để VBA Editor không bị lỗi mã hóa font trên máy có locale khác.

## 1. Cài đặt (làm 1 lần)

1. Tạo Workbook mới → **Save As** `Vehicle_Defect_Mapping.xlsm`.
2. `Alt+F11` → Import `modVehicleMapping.bas`.
3. Chạy `SetupWorkbook` (Alt+F8). Macro tự tạo:
   - 4 Sheet: `MAP_VEHICLE`, `MASTER`, `DATABASE`, `SUMMARY` (xóa Sheet trống thừa).
   - MASTER: `tblModels`, `tblDefectTypes`, `tblShifts`, `tblGridZones` + 2 cột trợ giúp `V`, `W` (danh sách chỉ lấy Active = Yes – **đừng xóa**).
   - DATABASE: `tblDefects` (17 cột). SUMMARY: `tblSummary`.
   - Giao diện MAP_VEHICLE, Data Validation, 3 nút SAVE / CLEAR SELECTION / RESET INPUT.
   - Hình **placeholder** `IMG_VF3`, `IMG_VF5`, `IMG_MNV` (hình chữ nhật xám, ẩn).
4. Dán code vào `MAP_VEHICLE` và `ThisWorkbook`.
5. Thay placeholder bằng hình Master thật (mục 2). Lưu file, đóng, mở lại.

## 2. Chuẩn bị hình Master

1. Trên sheet `MAP_VEHICLE`: xóa placeholder `IMG_VF3` (hoặc tạm thời để nguyên để thử).
2. Insert → Pictures, chọn hình Master gồm 5 góc nhìn LH/RH/FRONT/REAR/TOP.
3. Đặt tên shape trong **Name Box** (ô bên trái thanh công thức) đúng bằng cột `ImageShapeName`, ví dụ `IMG_VF3`.
4. Không cần tự canh vị trí/kích thước: khi chọn Model, code tự đặt tại B10, rộng 1000 pt (giữ tỷ lệ; đổi hằng `IMG_TARGET_WIDTH` nếu cần).

Model mới (VF8): thêm dòng cuối `tblModels` (`VF8 | 24 | 36 | IMG_VF8 | Yes`) + chèn hình đặt tên `IMG_VF8`. Không sửa code. Defect Type mới: thêm dòng cuối `tblDefectTypes`. Dropdown tự cập nhật (tối đa 60 mục; sửa `MAX_LIST_ROWS` nếu cần).

## 3. Cách dùng

1. Chọn Model (B2) → ảnh + Grid tự tải, B5/B6 cập nhật.
2. Click B3 → quét barcode (hoặc gõ tay). VIN giữ dạng Text, giữ số 0 đầu, CR/LF thừa bị loại. Quét xong **không tự Save**; con trỏ nhảy sang D4 (hoặc D5 nếu Inspector đã có).
3. Chọn Defect Type (D5), Qty/Cell (D6), Inspector (D4), Description (D7).
4. Click ô trên hình Master: chọn (xanh lá) / click lại: bỏ chọn. B4 = số ô đã chọn.
5. **SAVE**: kiểm tra 7 điều kiện → ghi 1 dòng/ô, chung `BatchID`, `DateTime`, `Shift`, `ShiftDate` → refresh SUMMARY → xóa chọn, reset VIN / Defect Type / Description, Qty = 1, giữ Model và Inspector → con trỏ về B3.
6. **CLEAR SELECTION**: chỉ xóa chọn Grid. **RESET INPUT**: xóa VIN, Defect Type, Description, Qty = 1 (giữ Model, Inspector, Grid, Database).

Shift/ShiftDate: tính bằng `GetShiftInfo`; `Time < 06:00` → S3 và ShiftDate = hôm qua. Cập nhật khi mở file, khi di chuyển ô, và lúc Save (không dùng Timer).

## 4. Cấu hình máy quét USB HID

- Máy quét ở chế độ **USB HID Keyboard** (mặc định).
- Bật hậu tố **Enter** (CR) hoặc Tab – code đã xử lý cả hai; hoặc để không hậu tố đều được.
- Bàn phím Windows để đúng layout (US) để không sai ký tự; tắt Caps Lock nếu máy quét không tự xử lý.
- Click B3 trước khi quét. Đừng để Excel ở chế độ Edit ô khác.
- Nếu VIN toàn số và ô B3 bị đổi định dạng, code tự đặt lại `@` (Text) và yêu cầu quét lại.

## 5. Bảo vệ sheet & môi trường

- `MAP_VEHICLE` được bảo vệ (không mật khẩu): chỉ B2:B3, D4:D7 sửa được; shape/nút vẫn bấm bình thường. Bỏ bảo vệ: Review → Unprotect Sheet.
- Cần bật Macro (Trusted Location khuyến nghị). Code không dùng API nên chạy được 32-bit và 64-bit. Cần Excel 2010+ (công thức `AGGREGATE`) và Windows (`Scripting.Dictionary`).
- `tblGridZones`: tự khai báo vùng theo hình thật (`Model, View, XFrom, XTo, YFrom, YTo`; X = cột, Y = hàng). Ô không nằm trong vùng nào → View = `MASTER`.

## 6. Test case

| # | Thao tác | Kết quả mong đợi |
|---|---|---|
| A | Thêm VF8 vào tblModels + chèn `IMG_VF8` | Có trong dropdown B2; chọn → ảnh + Grid 24×36 |
| B | Thêm Pin Hole vào tblDefectTypes | Có trong D5 |
| C | Quét `0012345678` + Enter | B3 giữ nguyên `0012345678`, không Save |
| D | Gọi `GetShiftInfo` với 05:59 / 06:00 / 13:59 / 14:00 / 21:59 / 22:00 / 23:59 / 00:00 | S3 hôm qua / S1 / S1 / S2 / S2 / S3 / S3 / S3 hôm qua |
| E | Chọn 3 ô, Scratch, Qty 2, Save | 3 dòng cùng BatchID/VIN/DefectType, Qty = 2, tổng 6 |
| F | Chọn nhiều ô → CLEAR SELECTION | Grid về mặc định, B4 = 0, các ô nhập giữ nguyên |
| G | Nhập VIN + Defect Type → RESET INPUT | Xóa theo quy tắc; Model và Database còn nguyên |
| H | Save sát 06:00 / 14:00 / 22:00 | Mọi dòng trong Batch cùng Shift và ShiftDate (lấy `Now` đúng 1 lần) |

Kiểm tra nhanh Shift (Immediate Window, Ctrl+G):

```vba
Dim s As String, d As Date
GetShiftInfo #11/13/2026 2:00:00 AM#, s, d
Debug.Print s, d   ' S3  12/11/2026
```

## 7. Xử lý lỗi thường gặp

| Hiện tượng | Cách xử lý |
|---|---|
| "Thieu hinh Master" | Chèn hình, đặt Name đúng `ImageShapeName` |
| Dropdown trống | Kiểm tra cột `Active` = `Yes` và công thức cột V/W trên MASTER chưa bị xóa |
| Click Grid không chạy | Macro bị tắt, hoặc Grid bị tạo lại thất bại: chọn lại Model ở B2 |
| "Grid khong thuoc Model" | Chọn lại Model để tạo lại Grid |
| Nút/Grid báo không tìm thấy macro | Module phải tên `modVehicleMapping`; chạy lại `SetupWorkbook` để gắn lại nút |
| Lỗi `Scripting.Dictionary` | Máy Mac/không có scripting runtime – hệ thống thiết kế cho Windows |
| Mất liên kết sau khi VBA bị reset | Hệ thống tự dựng lại danh sách ô chọn từ màu Grid khi cần |
