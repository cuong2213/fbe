# Production Quality Defect Dashboard – Stage 2

Dashboard được thêm vào chính file `Vehicle_Defect_Mapping.xlsm`. Dữ liệu nguồn là `DATABASE > tblDefects` (mỗi dòng = 1 ô lỗi, số lượng = `Qty`).

## Các quyết định đã chốt

| Mục | Quyết định |
|---|---|
| Location | Là **ô trên Grid (Cell ID)**, không dùng tên vùng. Không cần sheet `Defect_Map` |
| Cách đếm lỗi | **Tổng `Qty`** (đổi sang đếm số dòng: Config > `CountMode` = `Records`) |
| Vị trí | Thêm vào file cũ. Thêm 3 sheet: `Dashboard`, `Config`, `DASH_CALC` (ẩn, chứa dữ liệu biểu đồ) |
| Nguồn dữ liệu | Chỉ `DATABASE`. Không tạo bảng Database hợp nhất thứ hai |
| Defect Rate | Bỏ (không có Production Quantity) |
| Production Date / Shift | Dùng `ShiftDate` / `Shift` đã lưu lúc SAVE. Nếu trống thì tính lại từ `DateTime` bằng `GetShiftInfo` |

## Cài đặt

1. Thay `modVehicleMapping.bas` bằng bản mới (chỉ sửa `FindBlankSheet` để `SetupWorkbook` không xóa 3 sheet mới).
2. VBA Editor → File → **Import File** lần lượt: `modDashUtil.bas`, `modDashData.bas`, `modDashRender.bas`, `modDashboard.bas`.
3. Chạy `InitializeDashboard` (Alt+F8). Macro tạo 3 sheet, Config, bố cục, 5 biểu đồ, 3 nút, bộ lọc mặc định rồi tính lần đầu.
4. Dán nội dung `Sheet_Dashboard.txt` vào sheet **Dashboard** (Microsoft Excel Objects). Việc này bật tự cập nhật khi đổi bộ lọc.
5. Lưu file `.xlsm`.

## Bố cục sheet Dashboard (A:N)

| Vùng | Ô |
|---|---|
| Header, Last refresh, 3 nút | A1:N3 |
| Bộ lọc: From, To, Model, Shift, Defect Type, Location (Cell ID), Top N | A4:N6 (nhập ở hàng 5) |
| 4 KPI | A8:N11 |
| Trend theo ngày và ca / Trend theo loại lỗi (Top N) | A13:G27 / H13:N27 |
| Defect Summary / Defect by Shift | A29:G43 / H29:N43 |
| Defect by Model / Top 3 Locations (biểu đồ) | A45:G58 / H45:N58 |
| Top 3 Locations by Model / by Defect Type | A60:G74 / H60:N74 |
| Defect Type × Model (heatmap màu) | A76:N90 |
| Vehicle Defect Heatmap (Stage 3) | A92:N118 |
| Defect Detail (mới nhất trước, mặc định 500 dòng) | A120 trở xuống |

## Cách dùng

- Đổi bất kỳ bộ lọc nào ở hàng 5: Dashboard tự cập nhật. Dropdown Model, Shift, Defect Type, Location lấy từ dữ liệu thực tế (cộng thêm mục Active trong MASTER).
- **REFRESH DASHBOARD**: đọc lại dữ liệu mới và báo số dòng không hợp lệ. **RESET FILTERS**: To = hôm nay, From = hôm nay − 29 ngày, các lọc khác = All. **EXPORT TO PDF**: lưu PDF cùng thư mục file.
- Mở sheet Dashboard cũng tự cập nhật (tắt: Config > `AutoRefreshOnActivate` = `No`).
- Model mới, Defect Type mới, ô mới: tự xuất hiện sau khi có dữ liệu, không sửa code. Màu của loại lỗi mới được thêm vào Config > `tblDefectColors` (tô màu ô cột Color để đổi).

## Quy tắc tính

- Total Defects = tổng `Qty`. Affected Vehicles = số VIN khác nhau. Defects per Vehicle = Total / Affected (0 nếu không có xe).
- Previous Period dài bằng kỳ hiện tại, liền trước From date. Previous = 0 → Change % hiện `N/A`; có lỗi hiện tại thì Trend = Increasing.
- Trend Status: Change % vượt ±`StablePct` (mặc định 5) → Increasing / Decreasing, còn lại Stable.
- Top 3 vị trí: Cell ID có số lỗi cao nhất. Bằng nhau → ô hàng trên, cột trái xếp trước. Không hiện ô có 0 lỗi; ít hơn 3 ô thì chỉ hiện ô có.
- Bảng tối đa 12 dòng (Summary, Top 3, Matrix) và Matrix tối đa 12 Model. Vượt quá sẽ có dòng ghi chú "Hiển thị x/y".
- Dòng dữ liệu thiếu Model / VIN / DefectType / CellID, Qty không phải số dương, hoặc không xác định được ngày/ca sẽ bị **bỏ qua và báo**: dòng Last refresh có "Skipped", nút Refresh báo dòng đầu tiên bị lỗi.

## Config (sheet Config)

`TopN`, `StablePct`, `DetailMaxRows`, `DefaultDays`, `CountMode`, `AutoRefreshOnActivate`; bảng màu `tblDefectColors`. Ngưỡng heatmap trên hình xe sẽ thêm ở Stage 3.

## Test cần chạy (Stage 2)

Nhập dữ liệu bằng MAP_VEHICLE rồi kiểm tra:
1. Nhiều Model, nhiều VIN, một VIN có nhiều ô và nhiều loại lỗi → Total = tổng Qty, Affected = số VIN khác nhau.
2. Lưu lúc S1, S2, S3, trước và sau 00:00 và 06:00 → đúng Shift, ShiftDate (S3 02:00 thuộc ngày hôm trước).
3. Từng bộ lọc: Date, Model, Shift, Defect Type, Location; và kết hợp nhiều bộ lọc.
4. Trend theo ngày và ca; Trend theo loại lỗi (Top N, đổi N, chọn một loại lỗi).
5. Defect Summary: kỳ trước = 0 (N/A), tăng, giảm, ổn định; sắp theo Total giảm dần.
6. Top 3 theo Model và theo Defect Type, kể cả trường hợp bằng nhau và ít hơn 3 ô.
7. Ma trận Defect Type × Model; Defect by Shift; Defect by Model.
8. DATABASE rỗng, bộ lọc không có dữ liệu, thêm dữ liệu rồi Refresh, thêm Model / Defect Type mới.
9. Bấm Refresh nhiều lần: không có biểu đồ hoặc nút bị nhân đôi.

## Lưu ý

- Mỗi lần cập nhật đọc toàn bộ `tblDefects` vào mảng. Với vài chục nghìn dòng chỉ mất khoảng dưới một giây; hàng trăm nghìn dòng sẽ chậm dần (xem phần lưu trữ định kỳ đã bàn).
- Nếu Excel bị tắt Events/ScreenUpdating sau một lỗi VBA, chạy `ForceRestoreApp`.
- `Worksheet_Activate` / `Worksheet_Change` chỉ chạy sau khi đã dán `Sheet_Dashboard.txt`.

## Stage 3 – Vehicle Defect Heatmap (theo yêu cầu Location của bạn)

Hiểu như sau, bạn xác nhận hoặc chỉnh:
- Dùng hình xe + Grid của Model đang chọn (cùng kích thước hình trong `tblSettings`).
- **Mỗi loại lỗi một màu** (lấy từ `tblDefectColors`).
- **Mỗi loại lỗi chỉ hiện top 3 ô** nhiều lỗi nhất; số trên ô là số lỗi.
- **Mỗi Model hiện toàn bộ loại lỗi trên Grid** của Model đó; chọn Model = All thì hiện bảng tổng hợp thay vì gộp các Grid khác nhau.
- Một ô có nhiều loại lỗi: tô màu loại lỗi nhiều nhất.
