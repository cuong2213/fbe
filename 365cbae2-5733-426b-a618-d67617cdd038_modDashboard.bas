Attribute VB_Name = "modDashboard"
Option Explicit
'=====================================================================
' DASHBOARD - MODULE CHINH: khoi tao, bo cuc, lam moi, bo loc, xuat PDF
' Cach dung: chay InitializeDashboard mot lan; sau do dung nut tren sheet Dashboard.
' Khong dung UserForm.
'=====================================================================

'=====================================================================
' 1. CAC THAO TAC NGUOI DUNG (gan cho nut / su kien)
'=====================================================================
Public Sub RefreshDashboard()
    RunDashboard True
End Sub

Public Sub UpdateDashboard()
    'Goi khi doi bo loc (Worksheet_Change cua sheet Dashboard)
    RunDashboard False
End Sub

Public Sub DashboardOnActivate()
    If LCase$(Trim$(CStr(CfgValue("AutoRefreshOnActivate", "Yes")))) = "yes" Then RunDashboard False
End Sub

Public Sub ResetDashboardFilters()
    On Error GoTo EH
    FastOn
    SetFilterDefaults DashSheet(SH_DASH)
    FastOff
    RunDashboard False
    Exit Sub
EH:
    FastOff
    MsgBox "ResetDashboardFilters loi: " & Err.Description, vbCritical, "Dashboard"
End Sub

Private Function RunDashboard(ByVal showMsg As Boolean) As Boolean
    Dim msg As String
    On Error GoTo EH
    FastOn
    gDashStep = "kiem tra sheet"
    If Not (DashSheetExists(SH_DASH) And DashSheetExists(SH_CFG) And DashSheetExists(SH_CALC)) Then
        MsgBox "Dashboard chua duoc khoi tao. Hay chay InitializeDashboard.", vbExclamation, "Dashboard"
        GoTo CLEANUP
    End If

    gDashStep = "doc bo loc"
    If Not ReadFilters(msg) Then
        MsgBox msg, vbExclamation, "Dashboard"
        GoTo CLEANUP
    End If

    If Not LoadAndAggregate() Then GoTo CLEANUP

    gDashStep = "mau Defect Type (Config)"
    SyncDefectColors lDefects
    gDashStep = "dropdown bo loc"
    UpdateFilterDropdowns
    gDashStep = "KPI"
    UpdateKPI
    gDashStep = "Defect Summary"
    UpdateSummary
    gDashStep = "Top 3 vi tri"
    UpdateTop3Tables
    gDashStep = "ma tran Defect Type x Model"
    UpdateMatrix
    gDashStep = "bang chi tiet"
    UpdateDetail
    UpdateCharts
    gDashStep = "tieu de / trang thai"
    UpdateHeaderStatus
    RunDashboard = True

    If showMsg Then
        If rRows = 0 Then
            MsgBox "DATABASE (tblDefects) chua co du lieu.", vbInformation, "Dashboard"
        ElseIf rMatch = 0 Then
            MsgBox "Khong co du lieu phu hop bo loc hien tai.", vbInformation, "Dashboard"
        End If
        If rSkipped > 0 Then
            MsgBox rSkipped & " dong trong tblDefects khong hop le (thieu Model / VIN / DefectType / CellID, Qty sai hoac khong xac dinh duoc ngay-ca) nen bi bo qua." & _
                   vbCrLf & "Dong dau tien bi loi: dong " & rFirstBad & " cua sheet DATABASE.", vbExclamation, "Dashboard"
        End If
    End If

CLEANUP:
    FastOff
    Exit Function
EH:
    MsgBox "Dashboard loi: " & Err.Description & vbCrLf & "Buoc: " & gDashStep, vbCritical, "Dashboard"
    Resume CLEANUP
End Function

Private Sub UpdateHeaderStatus()
    Dim ws As Worksheet, s As String
    Set ws = DashSheet(SH_DASH)
    s = "Last refresh: " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & "     |     Period: " & _
        Format$(CDate(fFrom), "dd/mm/yyyy") & " - " & Format$(CDate(fTo), "dd/mm/yyyy") & _
        "     |     Rows: " & Format$(rRows, "#,##0")
    If rSkipped > 0 Then s = s & "     |     Skipped (invalid): " & rSkipped
    ws.Range("A3").Value = s
End Sub

Public Sub ExportDashboardPDF()
    Dim ws As Worksheet, folder As String, f As String, lastRow As Long, nDet As Long
    On Error GoTo EH
    Set ws = DashSheet(SH_DASH)
    folder = ThisWorkbook.Path
    If Len(folder) = 0 Then folder = CurDir$
    f = folder & Application.PathSeparator & "Dashboard_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    nDet = rDetailN
    If nDet > 40 Then nDet = 40                         'PDF chi in 40 dong chi tiet dau
    lastRow = ROW_DT_FIRST + nDet
    If lastRow < ROW_DT_FIRST + 1 Then lastRow = ROW_DT_FIRST + 1
    With ws.PageSetup
        .PrintArea = "$A$1:$N$" & lastRow
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "Da xuat PDF:" & vbCrLf & f, vbInformation, "Dashboard"
    Exit Sub
EH:
    MsgBox "Xuat PDF loi: " & Err.Description & vbCrLf & "(Can co may in mac dinh / Microsoft Print to PDF)", vbExclamation, "Dashboard"
End Sub

'=====================================================================
' 2. KHOI TAO
'=====================================================================
Public Sub InitializeDashboard()
    On Error GoTo EH
    If MsgBox("Khoi tao se tao / dung lai cac sheet Dashboard, Config, DASH_CALC (sheet an)." & vbCrLf & _
              "Bo cuc va bo loc cua Dashboard duoc dat lai. Du lieu DATABASE va cac sheet khac KHONG bi anh huong." & vbCrLf & _
              "Tiep tuc?", vbYesNo + vbQuestion, "Dashboard") <> vbYes Then Exit Sub

    FastOn
    gDashStep = "tao sheet"
    EnsureDashSheets
    gDashStep = "Config"
    BuildConfigSheet
    gDashStep = "DASH_CALC"
    BuildCalcSheet
    gDashStep = "bo cuc Dashboard"
    BuildDashboardLayout DashSheet(SH_DASH)
    gDashStep = "bieu do"
    CreateCharts DashSheet(SH_DASH)
    gDashStep = "nut bam"
    AddDashButtons DashSheet(SH_DASH)
    gDashStep = "bo loc mac dinh"
    SetFilterDefaults DashSheet(SH_DASH)

    gDashStep = "cua so"
    On Error Resume Next
    DashSheet(SH_DASH).Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .DisplayGridlines = False
        .Zoom = 90
    End With
    DashSheet(SH_DASH).Range("A7").Select
    ActiveWindow.FreezePanes = True
    DashSheet(SH_DASH).Range("A1").Select
    DashSheet(SH_CALC).Visible = xlSheetHidden
    On Error GoTo EH

    FastOff
    RunDashboard True
    Exit Sub
EH:
    FastOff
    MsgBox "InitializeDashboard loi: " & Err.Description & vbCrLf & "Buoc: " & gDashStep, vbCritical, "Dashboard"
End Sub

Private Sub EnsureDashSheets()
    Dim nm As Variant, ws As Worksheet
    For Each nm In Array(SH_DASH, SH_CFG, SH_CALC)
        If Not DashSheetExists(CStr(nm)) Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(nm)
        End If
    Next nm
    DashSheet(SH_CALC).Visible = xlSheetVisible
End Sub

Private Sub BuildConfigSheet()
    Dim ws As Worksheet, lo As ListObject
    Set ws = DashSheet(SH_CFG)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "DASHBOARD CONFIG"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Set lo = DashEnsureTable(ws, TBL_DSET, "A3", Array("Setting", "Value", "Note"), _
        Array(Array("TopN", 5, "So loai loi hien trong 'Defect Trend by Defect Type' (mac dinh o bo loc Top N, toi da 10)"), _
              Array("StablePct", 5, "Thay doi (%) so voi ky truoc nam trong +/- gia tri nay thi coi la Stable"), _
              Array("DetailMaxRows", 500, "So dong toi da cua bang Defect Detail (moi nhat truoc)"), _
              Array("DefaultDays", 30, "So ngay mac dinh cua bo loc ngay khi Reset (To = hom nay)"), _
              Array("CountMode", "Qty", "Qty = cong Qty cua tung dong (mac dinh); Records = moi dong tinh 1 loi"), _
              Array("AutoRefreshOnActivate", "Yes", "Yes = tu cap nhat khi mo sheet Dashboard")))

    DashEnsureTable ws, TBL_DCOL, "F3", Array("DefectType", "Color")
    ws.Range("F1").Value = "Mau theo Defect Type (to mau o cot Color de doi mau; loai loi moi duoc them tu dong)"
    ws.Range("F1").Font.Italic = True
    ws.Columns("A").ColumnWidth = 24
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 80
    ws.Columns("D:E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 24
    ws.Columns("G").ColumnWidth = 14
End Sub

Private Sub BuildCalcSheet()
    Dim ws As Worksheet
    Set ws = DashSheet(SH_CALC)
    ws.Range("A:D").NumberFormat = "@"        'danh sach dropdown luon la Text
End Sub

'=====================================================================
' 3. BO CUC DASHBOARD
'=====================================================================
Private Sub TitleBar(ByVal ws As Worksheet, ByVal addr As String, ByVal txt As String)
    With ws.Range(addr)
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range(addr).Cells(1, 1).Value = txt
End Sub

Private Sub HeaderRow(ByVal ws As Worksheet, ByVal firstAddr As String, ByVal heads As Variant)
    Dim i As Long, rg As Range
    Set rg = ws.Range(firstAddr).Resize(1, UBound(heads) - LBound(heads) + 1)
    For i = 0 To UBound(heads) - LBound(heads)
        rg.Cells(1, i + 1).Value = heads(LBound(heads) + i)
    Next i
    With rg
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Sub BoxTable(ByVal rg As Range)
    rg.Interior.Color = RGB(255, 255, 255)
    With rg.Borders
        .LineStyle = xlContinuous
        .Color = RGB(200, 200, 200)
        .Weight = xlThin
    End With
End Sub

Private Sub BuildDashboardLayout(ByVal ws As Worksheet)
    Dim f As Variant, c As Variant, fc As FormatCondition, cs As ColorScale, rg As Range, i As Long

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(242, 244, 247)
    ws.Columns("A:N").ColumnWidth = 12.5
    ws.Columns("O").ColumnWidth = 3
    ws.Rows("1:400").RowHeight = 16

    '--- Header ---
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 20
    ws.Rows(3).RowHeight = 20
    With ws.Range("A1:N3")
        .Interior.Color = RGB(16, 37, 66)
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A1").Value = "PRODUCTION QUALITY DEFECT DASHBOARD"
    ws.Range("A1").Font.Size = 20
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Vehicle Defect Analysis - Paint Shop"
    ws.Range("A2").Font.Size = 11
    ws.Range("A3").Font.Size = 9
    ws.Range("A3").Font.Color = RGB(190, 205, 225)

    '--- Bo loc ---
    ws.Rows(4).RowHeight = 18
    ws.Rows(5).RowHeight = 24
    ws.Rows(6).RowHeight = 16
    For Each f In Array(Array("A", "B", "From date"), Array("C", "D", "To date"), Array("E", "F", "Model"), _
                        Array("G", "H", "Shift"), Array("I", "J", "Defect Type"), Array("K", "L", "Location (Cell ID)"), _
                        Array("M", "N", "Top N defect types"))
        With ws.Range(f(0) & "4:" & f(1) & "4")
            .Merge
            .Interior.Color = RGB(89, 89, 89)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        ws.Range(f(0) & "4").Value = f(2)
        With ws.Range(f(0) & "5:" & f(1) & "5")
            .Merge
            .Interior.Color = RGB(255, 242, 204)
            .Font.Bold = True
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(150, 150, 150)
        End With
    Next f
    ws.Range(F_FROM).NumberFormat = "dd/mm/yyyy"
    ws.Range(F_TO).NumberFormat = "dd/mm/yyyy"
    ws.Range("A6").Value = "Vang = bo loc (doi la Dashboard tu cap nhat). Location = vi tri tren Grid (Cell ID). Dropdown lay tu du lieu thuc te."
    ws.Range("A6").Font.Italic = True
    ws.Range("A6").Font.Size = 9
    ws.Range("A6").Font.Color = RGB(90, 90, 90)

    '--- KPI ---
    ws.Rows(8).RowHeight = 20
    ws.Rows(9).RowHeight = 26
    ws.Rows(10).RowHeight = 26
    ws.Rows(11).RowHeight = 20
    For Each c In Array(Array("A", "C", "TOTAL DEFECTS", RGB(31, 78, 121)), Array("D", "G", "AFFECTED VEHICLES", RGB(31, 78, 121)), _
                        Array("H", "K", "DEFECTS PER VEHICLE", RGB(31, 78, 121)), Array("L", "N", "TOP DEFECT TYPE", RGB(192, 80, 0)))
        With ws.Range(c(0) & "8:" & c(1) & "8")
            .Merge
            .Interior.Color = c(3)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        ws.Range(c(0) & "8").Value = c(2)
        With ws.Range(c(0) & "9:" & c(1) & "10")
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Font.Size = 24
            .Font.Bold = True
            .Font.Color = RGB(16, 37, 66)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With ws.Range(c(0) & "11:" & c(1) & "11")
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Font.Size = 9
            .Font.Color = RGB(90, 90, 90)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With ws.Range(c(0) & "8:" & c(1) & "11")
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Color = RGB(242, 244, 247)
            .Borders(xlEdgeLeft).Weight = xlThick
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(242, 244, 247)
            .Borders(xlEdgeRight).Weight = xlThick
        End With
    Next c
    ws.Range(KPI1_VAL).NumberFormat = "#,##0"
    ws.Range(KPI2_VAL).NumberFormat = "#,##0"
    ws.Range(KPI3_VAL).NumberFormat = "0.00"
    ws.Range(KPI4_VAL).Font.Size = 18

    '--- Defect Summary ---
    TitleBar ws, "A29:G29", "DEFECT SUMMARY"
    HeaderRow ws, "A" & ROW_SUM_HDR, Array("Defect Type", "Total Defects", "Percentage", "Affected VIN", "Previous Period Defects", "Change %", "Trend Status")
    ws.Rows(ROW_SUM_HDR).RowHeight = 32
    BoxTable ws.Range("A" & ROW_SUM_FIRST & ":G" & (ROW_SUM_FIRST + SUM_MAXROWS - 1))
    ws.Range("B" & ROW_SUM_FIRST & ":B" & (ROW_SUM_FIRST + SUM_MAXROWS - 1)).NumberFormat = "#,##0"
    ws.Range("C" & ROW_SUM_FIRST & ":C" & (ROW_SUM_FIRST + SUM_MAXROWS - 1)).NumberFormat = "0.0%"
    ws.Range("D" & ROW_SUM_FIRST & ":E" & (ROW_SUM_FIRST + SUM_MAXROWS - 1)).NumberFormat = "#,##0"
    ws.Range("F" & ROW_SUM_FIRST & ":F" & (ROW_SUM_FIRST + SUM_MAXROWS - 1)).NumberFormat = "+0.0%;-0.0%;0.0%"
    ws.Range("B" & ROW_SUM_FIRST & ":G" & (ROW_SUM_FIRST + SUM_MAXROWS - 1)).HorizontalAlignment = xlCenter
    With ws.Range("A" & (ROW_SUM_FIRST + SUM_MAXROWS))
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(90, 90, 90)
    End With
    Set rg = ws.Range("G" & ROW_SUM_FIRST & ":G" & (ROW_SUM_FIRST + SUM_MAXROWS - 1))
    rg.FormatConditions.Delete
    Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Increasing""")
    fc.Font.Color = RGB(156, 0, 6): fc.Interior.Color = RGB(255, 199, 206)
    Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Decreasing""")
    fc.Font.Color = RGB(0, 97, 0): fc.Interior.Color = RGB(198, 239, 206)
    Set fc = rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Stable""")
    fc.Font.Color = RGB(64, 64, 64): fc.Interior.Color = RGB(217, 217, 217)

    '--- Top 3 vi tri theo Model / Defect Type ---
    TitleBar ws, "A60:G60", "TOP 3 DEFECT LOCATIONS BY MODEL (Cell ID)"
    TitleBar ws, "H60:N60", "TOP 3 DEFECT LOCATIONS BY DEFECT TYPE (Cell ID)"
    HeaderRow ws, "A" & ROW_T3_HDR, Array("Model", "Rank 1 Location", "Count", "Rank 2 Location", "Count", "Rank 3 Location", "Count")
    HeaderRow ws, "H" & ROW_T3_HDR, Array("Defect Type", "Rank 1 Location", "Count", "Rank 2 Location", "Count", "Rank 3 Location", "Count")
    ws.Rows(ROW_T3_HDR).RowHeight = 30
    For Each f In Array("A", "H")
        Set rg = ws.Range(f & ROW_T3_FIRST).Resize(T3_MAXROWS, 7)
        BoxTable rg
        rg.Offset(0, 1).Resize(T3_MAXROWS, 6).HorizontalAlignment = xlCenter
        With ws.Range(f & (ROW_T3_FIRST + T3_MAXROWS))
            .Font.Italic = True
            .Font.Size = 9
            .Font.Color = RGB(90, 90, 90)
        End With
    Next f

    '--- Ma tran Defect Type x Model ---
    TitleBar ws, "A76:N76", "DEFECT TYPE x MODEL MATRIX"
    HeaderRow ws, "A" & ROW_MX_HDR, Array("Defect Type", "", "", "", "", "", "", "", "", "", "", "", "", "Total")
    BoxTable ws.Range("A" & ROW_MX_FIRST & ":N" & (ROW_MX_FIRST + MX_MAXROWS))
    ws.Range("B" & ROW_MX_FIRST & ":N" & (ROW_MX_FIRST + MX_MAXROWS)).HorizontalAlignment = xlCenter
    ws.Range("B" & ROW_MX_FIRST & ":N" & (ROW_MX_FIRST + MX_MAXROWS)).NumberFormat = "#,##0"
    ws.Range("A" & (ROW_MX_FIRST + MX_MAXROWS) & ":N" & (ROW_MX_FIRST + MX_MAXROWS)).Font.Bold = True
    ws.Range("N" & ROW_MX_FIRST & ":N" & (ROW_MX_FIRST + MX_MAXROWS)).Font.Bold = True
    Set rg = ws.Range("B" & ROW_MX_FIRST & ":M" & (ROW_MX_FIRST + MX_MAXROWS - 1))
    rg.FormatConditions.Delete
    Set cs = rg.FormatConditions.AddColorScale(3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)
    cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 217, 102)
    cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(220, 53, 69)

    '--- Heatmap (Stage 3) ---
    TitleBar ws, "A92:N92", "VEHICLE DEFECT HEATMAP"
    ws.Range("A93").Value = "(Se bo sung o Stage 3: hinh xe + Grid, moi loai loi 1 mau, top 3 vi tri cho tung loai loi)"
    ws.Range("A93").Font.Italic = True
    ws.Range("A93").Font.Color = RGB(90, 90, 90)

    '--- Bang chi tiet ---
    TitleBar ws, "A" & (ROW_DT_HDR - 1) & ":N" & (ROW_DT_HDR - 1), "DEFECT DETAIL"
    HeaderRow ws, "A" & ROW_DT_HDR, Array("Production Date", "DateTime", "Shift", "Model", "VIN", "Defect Type", "Qty", _
                                           "Grid Row", "Grid Column", "Cell ID", "Inspector", "Description")
    ws.Rows(ROW_DT_HDR).RowHeight = 30
    With ws.Range("A" & ROW_DT_FIRST & ":N" & (ROW_DT_FIRST + 2000))
        .Interior.Color = RGB(255, 255, 255)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(225, 225, 225)
    End With
    ws.Range("A" & ROW_DT_FIRST & ":A" & (ROW_DT_FIRST + 5000)).NumberFormat = "dd/mm/yyyy"
    ws.Range("B" & ROW_DT_FIRST & ":B" & (ROW_DT_FIRST + 5000)).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("E" & ROW_DT_FIRST & ":E" & (ROW_DT_FIRST + 5000)).NumberFormat = "@"
    ws.Range("J" & ROW_DT_FIRST & ":J" & (ROW_DT_FIRST + 5000)).NumberFormat = "@"
    ws.Range("A" & ROW_DT_FIRST & ":D" & (ROW_DT_FIRST + 5000)).HorizontalAlignment = xlCenter
    ws.Range("G" & ROW_DT_FIRST & ":J" & (ROW_DT_FIRST + 5000)).HorizontalAlignment = xlCenter

    '--- Kiem tra du lieu nhap cho bo loc ---
    With ws.Range(F_FROM).MergeArea.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="=DATE(2000,1,1)"
        .ErrorTitle = "From date"
        .ErrorMessage = "Nhap ngay hop le (dd/mm/yyyy)."
    End With
    With ws.Range(F_TO).MergeArea.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="=DATE(2000,1,1)"
        .ErrorTitle = "To date"
        .ErrorMessage = "Nhap ngay hop le (dd/mm/yyyy)."
    End With
    With ws.Range(F_TOPN).MergeArea.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="10"
        .ErrorTitle = "Top N"
        .ErrorMessage = "Nhap so nguyen tu 1 den 10."
    End With
End Sub

Private Sub SetFilterDefaults(ByVal ws As Worksheet)
    Dim days As Long
    days = CLng(Val(CStr(CfgValue("DefaultDays", 30))))
    If days < 1 Then days = 30
    ws.Range(F_TO).Value = Date
    ws.Range(F_FROM).Value = Date - days + 1
    ws.Range(F_MODEL).Value = "All"
    ws.Range(F_SHIFT).Value = "All"
    ws.Range(F_DEF).Value = "All"
    ws.Range(F_CELL).Value = "All"
    ws.Range(F_TOPN).Value = CLng(Val(CStr(CfgValue("TopN", 5))))
End Sub

'=====================================================================
' 4. BIEU DO (tao 1 lan) VA NUT BAM
'=====================================================================
Private Sub EnsureChart(ByVal ws As Worksheet, ByVal nm As String, ByVal rng As Range, ByVal ctype As Long, ByVal ttl As String)
    Dim co As ChartObject
    On Error Resume Next
    Set co = ws.ChartObjects(nm)
    On Error GoTo 0
    If co Is Nothing Then
        Set co = ws.ChartObjects.Add(rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
        co.Name = nm
        With co.Chart
            .ChartType = ctype
            .HasTitle = True
            .ChartTitle.Text = ttl
            .ChartTitle.Font.Size = 11
            .ChartTitle.Font.Bold = True
            If ctype = xlLine Then
                .HasLegend = True
                .Legend.Position = xlLegendPositionBottom
            Else
                .HasLegend = False
            End If
        End With
        On Error Resume Next
        co.Chart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        co.Chart.ChartArea.Format.Line.ForeColor.RGB = RGB(200, 200, 200)
        co.Chart.ChartArea.Font.Size = 9
        On Error GoTo 0
    Else
        co.Left = rng.Left + 2: co.Top = rng.Top + 2
        co.Width = rng.Width - 4: co.Height = rng.Height - 4
    End If
    co.Placement = xlMoveAndSize
End Sub

Private Sub CreateCharts(ByVal ws As Worksheet)
    EnsureChart ws, "chtTrendShift", ws.Range("A13:G27"), xlLine, "DEFECT TREND BY DATE AND SHIFT"
    EnsureChart ws, "chtTrendDefect", ws.Range("H13:N27"), xlLine, "DEFECT TREND BY DEFECT TYPE"
    EnsureChart ws, "chtByShift", ws.Range("H29:N43"), xlColumnClustered, "DEFECT BY SHIFT"
    EnsureChart ws, "chtByModel", ws.Range("A45:G58"), xlBarClustered, "DEFECT BY MODEL"
    EnsureChart ws, "chtTop3", ws.Range("H45:N58"), xlBarClustered, "TOP 3 DEFECT LOCATIONS (CELL ID)"
End Sub

Private Sub DashButton(ByVal ws As Worksheet, ByVal shpName As String, ByVal caption As String, _
                       ByVal rng As Range, ByVal clr As Long, ByVal macroName As String)
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(shpName).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left + 3, rng.Top + 4, rng.Width - 6, rng.Height - 8)
    With s
        .Name = shpName
        .Fill.ForeColor.RGB = clr
        .Line.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlFreeFloating
        With .TextFrame2.TextRange
            .Text = caption
            .Font.Size = 10
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
End Sub

Private Sub AddDashButtons(ByVal ws As Worksheet)
    DashButton ws, "BTN_DASH_REFRESH", "REFRESH DASHBOARD", ws.Range("H1:I2"), RGB(0, 140, 70), "modDashboard.RefreshDashboard"
    DashButton ws, "BTN_DASH_RESET", "RESET FILTERS", ws.Range("J1:K2"), RGB(200, 120, 0), "modDashboard.ResetDashboardFilters"
    DashButton ws, "BTN_DASH_PDF", "EXPORT TO PDF", ws.Range("L1:M2"), RGB(31, 120, 180), "modDashboard.ExportDashboardPDF"
End Sub
