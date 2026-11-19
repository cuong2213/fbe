Attribute VB_Name = "modDashRender"
Option Explicit
'=====================================================================
' DASHBOARD - HIEN THI: dropdown, KPI, bang, ma tran, chi tiet, bieu do
' Bieu do tao 1 lan (modDashboard.CreateCharts); moi lan Refresh chi doi Series.
' Du lieu bieu do nam o sheet an DASH_CALC.
'=====================================================================

'=====================================================================
' Dropdown bo loc
'=====================================================================
Private Sub WriteList(ByVal calc As Worksheet, ByVal col As Long, ByVal title As String, _
                      ByRef ks() As String, ByVal n As Long, ByVal target As Range)
    Dim arr() As Variant, i As Long, lastRow As Long
    ReDim arr(1 To n + 1, 1 To 1)
    arr(1, 1) = "All"
    For i = 1 To n
        arr(i + 1, 1) = ks(i)
    Next i
    calc.Cells(1, col).Value = title
    calc.Cells(2, col).Resize(n + 1, 1).Value = arr
    lastRow = n + 2
    With target.MergeArea.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & SH_CALC & "!" & calc.Range(calc.Cells(2, col), calc.Cells(lastRow, col)).Address
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

Public Sub UpdateFilterDropdowns()
    Dim ws As Worksheet, calc As Worksheet, ks() As String, vs() As Double, n As Long
    Set ws = DashSheet(SH_DASH)
    Set calc = DashSheet(SH_CALC)
    calc.Range("A1:D3000").ClearContents

    n = SortDictKeys(lModels, ks, vs, True)         'bang nhau -> theo ten tang dan
    WriteList calc, 1, "Models", ks, n, ws.Range(F_MODEL)

    ReDim ks(1 To 1)
    n = ShiftNames(ks)
    WriteList calc, 2, "Shifts", ks, n, ws.Range(F_SHIFT)

    n = SortDictKeys(lDefects, ks, vs, True)
    WriteList calc, 3, "DefectTypes", ks, n, ws.Range(F_DEF)

    n = SortDictKeys(lCells, ks, vs, False)         'theo hang roi cot
    WriteList calc, 4, "Cells", ks, n, ws.Range(F_CELL)
End Sub

Private Function ShiftNames(ByRef ks() As String) As Long
    'Danh sach ca: tblShifts (Active = Yes); neu khong co thi lay tu du lieu
    Dim lo As ListObject, dat As Variant, i As Long, n As Long, cS As Long, cA As Long
    Dim vs() As Double
    Set lo = DashTryTable(SH_MASTER, TBL_SHIFTS)
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            cS = DashColIndex(lo, "Shift", False)
            cA = DashColIndex(lo, "Active", False)
            If cS > 0 Then
                dat = lo.DataBodyRange.Value2
                ReDim ks(1 To UBound(dat, 1))
                For i = 1 To UBound(dat, 1)
                    If Len(Trim$(CStr(dat(i, cS)))) > 0 Then
                        If cA = 0 Then
                            n = n + 1: ks(n) = UCase$(Trim$(CStr(dat(i, cS))))
                        ElseIf IsActiveValue(dat(i, cA)) Then
                            n = n + 1: ks(n) = UCase$(Trim$(CStr(dat(i, cS))))
                        End If
                    End If
                Next i
            End If
        End If
    End If
    If n = 0 Then n = SortDictKeys(lShifts, ks, vs, True)
    ShiftNames = n
End Function

'=====================================================================
' KPI
'=====================================================================
Public Sub UpdateKPI()
    Dim ws As Worksheet, ks() As String, vs() As Double, n As Long, v As Long, dpv As Double, pct As Double
    Set ws = DashSheet(SH_DASH)
    v = rVin.Count

    ws.Range(KPI1_VAL).Value = rTotal
    ws.Range(KPI1_SUB).Value = rMatch & " ban ghi"

    ws.Range(KPI2_VAL).Value = v
    ws.Range(KPI2_SUB).Value = "VIN co it nhat 1 loi"

    dpv = 0
    If v > 0 Then dpv = rTotal / v
    ws.Range(KPI3_VAL).Value = dpv
    ws.Range(KPI3_SUB).Value = "Total Defects / Affected Vehicles"

    n = SortDictKeys(rDefQty, ks, vs, True)
    If n > 0 Then
        pct = 0
        If rTotal > 0 Then pct = vs(1) / rTotal
        ws.Range(KPI4_VAL).Value = ks(1)
        ws.Range(KPI4_SUB).Value = Format$(vs(1), "#,##0") & " loi  (" & Format$(pct, "0.0%") & ")"
    Else
        ws.Range(KPI4_VAL).Value = "-"
        ws.Range(KPI4_SUB).Value = "Khong co du lieu"
    End If
End Sub

'=====================================================================
' DEFECT SUMMARY
'=====================================================================
Public Sub UpdateSummary()
    Dim ws As Worksheet, ks() As String, vs() As Double, n As Long, m As Long, i As Long
    Dim arr() As Variant, cur As Double, prv As Double, chg As Double, st As String, thr As Double
    Dim vinCnt As Long, noteRow As Long, note As String

    Set ws = DashSheet(SH_DASH)
    noteRow = ROW_SUM_FIRST + SUM_MAXROWS
    ws.Range("A" & ROW_SUM_FIRST & ":G" & noteRow).ClearContents
    thr = Val(CStr(CfgValue("StablePct", 5)))

    n = SortDictKeys(rDefQty, ks, vs, True)
    m = n
    If m > SUM_MAXROWS Then m = SUM_MAXROWS
    note = "Ky truoc: " & Format$(CDate(pFrom), "dd/mm/yyyy") & " - " & Format$(CDate(pTo), "dd/mm/yyyy")

    If n = 0 Then
        ws.Range("A" & ROW_SUM_FIRST).Value = "Khong co du lieu theo bo loc hien tai"
    Else
        ReDim arr(1 To m, 1 To 7)
        For i = 1 To m
            cur = vs(i)
            prv = 0
            If rDefPrev.Exists(ks(i)) Then prv = CDbl(rDefPrev(ks(i)))
            vinCnt = 0
            If rDefVin.Exists(ks(i)) Then vinCnt = rDefVin(ks(i)).Count
            arr(i, 1) = ks(i)
            arr(i, 2) = cur
            If rTotal > 0 Then arr(i, 3) = cur / rTotal Else arr(i, 3) = 0
            arr(i, 4) = vinCnt
            arr(i, 5) = prv
            If prv = 0 Then
                arr(i, 6) = "N/A"
                If cur > 0 Then st = "Increasing" Else st = "Stable"
            Else
                chg = (cur - prv) / prv
                arr(i, 6) = chg
                If chg * 100 > thr Then
                    st = "Increasing"
                ElseIf chg * 100 < -thr Then
                    st = "Decreasing"
                Else
                    st = "Stable"
                End If
            End If
            arr(i, 7) = st
        Next i
        ws.Range("A" & ROW_SUM_FIRST).Resize(m, 7).Value = arr
        If n > m Then note = note & "   |   Hien thi " & m & "/" & n & " loai loi"
    End If
    ws.Range("A" & noteRow).Value = note
End Sub

'=====================================================================
' TOP 3 VI TRI (CellID) THEO MODEL / THEO DEFECT TYPE
'=====================================================================
Private Sub FillTop3(ByVal ws As Worksheet, ByVal firstCol As String, ByVal groupQty As Object, _
                     ByVal groupCells As Object, ByVal noteText As String)
    Dim ks() As String, vs() As Double, n As Long, m As Long, i As Long, r As Long, c As Long
    Dim nm() As String, ct() As Double, cnt As Long, arr() As Variant, lastCol As String, noteRow As Long
    Dim rngTop As Long, colIdx As Long

    colIdx = ws.Range(firstCol & "1").Column
    lastCol = Split(ws.Cells(1, colIdx + 6).Address(True, False), "$")(0)
    noteRow = ROW_T3_FIRST + T3_MAXROWS
    ws.Range(firstCol & ROW_T3_FIRST & ":" & lastCol & noteRow).ClearContents

    n = SortDictKeys(groupQty, ks, vs, True)
    If n = 0 Then
        ws.Range(firstCol & ROW_T3_FIRST).Value = "Khong co du lieu theo bo loc hien tai"
        Exit Sub
    End If
    m = n
    If m > T3_MAXROWS Then m = T3_MAXROWS
    ReDim arr(1 To m, 1 To 7)
    For i = 1 To m
        arr(i, 1) = ks(i)
        If groupCells.Exists(ks(i)) Then
            cnt = Top3(groupCells(ks(i)), lCells, nm, ct)
            For r = 1 To cnt
                arr(i, 2 * r) = nm(r)
                arr(i, 2 * r + 1) = ct(r)
            Next r
        End If
    Next i
    ws.Range(firstCol & ROW_T3_FIRST).Resize(m, 7).Value = arr
    If n > m Then ws.Range(firstCol & noteRow).Value = "Hien thi " & m & "/" & n & " muc"
End Sub

Public Sub UpdateTop3Tables()
    Dim ws As Worksheet
    Set ws = DashSheet(SH_DASH)
    FillTop3 ws, "A", rModelQty, rModelCell, "Model"
    FillTop3 ws, "H", rDefQty, rDefCell, "Defect Type"
End Sub

'=====================================================================
' DEFECT TYPE x MODEL
'=====================================================================
Public Sub UpdateMatrix()
    Dim ws As Worksheet, dk() As String, dv() As Double, mk() As String, mv() As Double
    Dim nD As Long, nM As Long, i As Long, j As Long, key As String, arr() As Variant
    Dim totRow As Long, lastRow As Long, rowTot As Double

    Set ws = DashSheet(SH_DASH)
    totRow = ROW_MX_FIRST + MX_MAXROWS
    ws.Range("A" & ROW_MX_HDR & ":N" & totRow).ClearContents

    nD = SortDictKeys(rDefQty, dk, dv, True)
    nM = SortDictKeys(rModelQty, mk, mv, True)
    If nD > MX_MAXROWS Then nD = MX_MAXROWS
    If nM > 12 Then nM = 12

    ws.Range("A" & ROW_MX_HDR).Value = "Defect Type"
    ws.Range("N" & ROW_MX_HDR).Value = "Total"
    If rDefQty.Count = 0 Then
        ws.Range("A" & ROW_MX_FIRST).Value = "Khong co du lieu theo bo loc hien tai"
        Exit Sub
    End If

    For j = 1 To nM
        ws.Cells(ROW_MX_HDR, 1 + j).Value = mk(j)
    Next j

    ReDim arr(1 To nD, 1 To 14)
    For i = 1 To nD
        arr(i, 1) = dk(i)
        For j = 1 To nM
            key = dk(i) & "|" & mk(j)
            If rMatrix.Exists(key) Then arr(i, 1 + j) = rMatrix(key)
        Next j
        arr(i, 14) = rDefQty(dk(i))
    Next i
    ws.Range("A" & ROW_MX_FIRST).Resize(nD, 14).Value = arr

    ws.Range("A" & totRow).Value = "Total"
    For j = 1 To nM
        ws.Cells(totRow, 1 + j).Value = rModelQty(mk(j))
    Next j
    ws.Range("N" & totRow).Value = rTotal
End Sub

'=====================================================================
' BANG CHI TIET
'=====================================================================
Public Sub UpdateDetail()
    Dim ws As Worksheet, ttl As String
    Set ws = DashSheet(SH_DASH)
    ws.Range("A" & ROW_DT_FIRST & ":L" & (ROW_DT_FIRST + 5000)).ClearContents
    ttl = "DEFECT DETAIL"
    If rDetailN > 0 Then
        ws.Range("A" & ROW_DT_FIRST).Resize(rDetailN, 12).Value = rDetail
        ttl = ttl & "   (moi nhat " & rDetailN & " / " & rMatch & " ban ghi)"
    Else
        ws.Range("A" & ROW_DT_FIRST).Value = "Khong co du lieu theo bo loc hien tai"
    End If
    ws.Range("A" & (ROW_DT_HDR - 1)).Value = ttl
End Sub

'=====================================================================
' BIEU DO
'=====================================================================
Private Function GetChart(ByVal ws As Worksheet, ByVal nm As String) As Chart
    Dim co As ChartObject
    On Error Resume Next
    Set co = ws.ChartObjects(nm)
    On Error GoTo 0
    If co Is Nothing Then
        Err.Raise vbObjectError + 710, "GetChart", "Thieu bieu do '" & nm & "'. Hay chay InitializeDashboard."
    End If
    Set GetChart = co.Chart
End Function

Private Sub ClearSeries(ByVal cht As Chart)
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
End Sub

Private Sub AddSeries(ByVal cht As Chart, ByVal nameCell As Range, ByVal xr As Range, ByVal yr As Range, _
                      ByVal clr As Long, ByVal asLine As Boolean, ByVal withLabels As Boolean)
    Dim s As Series
    Set s = cht.SeriesCollection.NewSeries
    s.Name = CStr(nameCell.Value)
    s.XValues = xr
    s.Values = yr
    If asLine Then
        s.Format.Line.ForeColor.RGB = clr
        s.Format.Line.Weight = 2.25
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 5
        s.MarkerBackgroundColor = clr
        s.MarkerForegroundColor = clr
    Else
        s.Format.Fill.ForeColor.RGB = clr
    End If
    If withLabels Then s.HasDataLabels = True
End Sub

Private Function DateRows() As Long
    DateRows = fTo - fFrom + 1
    If DateRows > 2000 Then DateRows = 2000
    If DateRows < 1 Then DateRows = 1
End Function

Private Sub UpdateTrendShift(ByVal ws As Worksheet, ByVal calc As Worksheet)
    Dim cht As Chart, sh() As String, nSh As Long, nD As Long, i As Long, j As Long
    Dim arr() As Variant, dd As Long, key As String
    Const C0 As Long = 6
    Set cht = GetChart(ws, "chtTrendShift")
    ClearSeries cht
    ReDim sh(1 To 1)
    nSh = ShiftNames(sh)
    If nSh = 0 Then Exit Sub
    nD = DateRows()
    ReDim arr(1 To nD + 1, 1 To nSh + 1)
    arr(1, 1) = "Date"
    For j = 1 To nSh
        arr(1, j + 1) = sh(j)
    Next j
    For i = 1 To nD
        dd = fFrom + i - 1
        arr(i + 1, 1) = CDbl(dd)
        For j = 1 To nSh
            key = dd & "|" & sh(j)
            If rDateShift.Exists(key) Then arr(i + 1, j + 1) = rDateShift(key) Else arr(i + 1, j + 1) = 0
        Next j
    Next i
    calc.Cells(1, C0).Resize(nD + 1, nSh + 1).Value = arr
    calc.Range(calc.Cells(2, C0), calc.Cells(nD + 1, C0)).NumberFormat = "dd/mm"
    For j = 1 To nSh
        AddSeries cht, calc.Cells(1, C0 + j), calc.Range(calc.Cells(2, C0), calc.Cells(nD + 1, C0)), _
                  calc.Range(calc.Cells(2, C0 + j), calc.Cells(nD + 1, C0 + j)), PaletteColor(j - 1), True, False
    Next j
    On Error Resume Next
    cht.Axes(xlCategory).CategoryType = xlCategoryScale
    cht.Axes(xlCategory).TickLabels.NumberFormat = "dd/mm"
    cht.Axes(xlValue).HasTitle = True
    cht.Axes(xlValue).AxisTitle.Text = "Defect Count"
    On Error GoTo 0
End Sub

Private Sub UpdateTrendDefect(ByVal ws As Worksheet, ByVal calc As Worksheet)
    Dim cht As Chart, ks() As String, vs() As Double, n As Long, k As Long, nD As Long, i As Long, j As Long
    Dim arr() As Variant, dd As Long, key As String
    Const C0 As Long = 14
    Set cht = GetChart(ws, "chtTrendDefect")
    ClearSeries cht
    n = SortDictKeys(rDefQty, ks, vs, True)
    If n = 0 Then Exit Sub
    k = n
    If k > fTopN Then k = fTopN                        'Top N loai loi (khong hien loai co so loi = 0)
    nD = DateRows()
    ReDim arr(1 To nD + 1, 1 To k + 1)
    arr(1, 1) = "Date"
    For j = 1 To k
        arr(1, j + 1) = ks(j)
    Next j
    For i = 1 To nD
        dd = fFrom + i - 1
        arr(i + 1, 1) = CDbl(dd)
        For j = 1 To k
            key = dd & "|" & ks(j)
            If rDateDef.Exists(key) Then arr(i + 1, j + 1) = rDateDef(key) Else arr(i + 1, j + 1) = 0
        Next j
    Next i
    calc.Cells(1, C0).Resize(nD + 1, k + 1).Value = arr
    calc.Range(calc.Cells(2, C0), calc.Cells(nD + 1, C0)).NumberFormat = "dd/mm"
    For j = 1 To k
        AddSeries cht, calc.Cells(1, C0 + j), calc.Range(calc.Cells(2, C0), calc.Cells(nD + 1, C0)), _
                  calc.Range(calc.Cells(2, C0 + j), calc.Cells(nD + 1, C0 + j)), DefectColor(ks(j)), True, False
    Next j
    On Error Resume Next
    cht.Axes(xlCategory).CategoryType = xlCategoryScale
    cht.Axes(xlCategory).TickLabels.NumberFormat = "dd/mm"
    cht.Axes(xlValue).HasTitle = True
    cht.Axes(xlValue).AxisTitle.Text = "Defect Count"
    On Error GoTo 0
End Sub

Private Sub UpdateByShift(ByVal ws As Worksheet, ByVal calc As Worksheet)
    Dim cht As Chart, sh() As String, nSh As Long, i As Long, arr() As Variant, q As Double, v As Long, dpv As Double
    Const C0 As Long = 27
    Set cht = GetChart(ws, "chtByShift")
    ClearSeries cht
    ReDim sh(1 To 1)
    nSh = ShiftNames(sh)
    If nSh = 0 Then Exit Sub
    ReDim arr(1 To nSh + 1, 1 To 4)
    arr(1, 1) = "Shift": arr(1, 2) = "Defects": arr(1, 3) = "Affected VINs": arr(1, 4) = "Defects/Vehicle"
    For i = 1 To nSh
        q = 0: v = 0: dpv = 0
        If rShiftQty.Exists(sh(i)) Then q = CDbl(rShiftQty(sh(i)))
        If rShiftVin.Exists(sh(i)) Then v = rShiftVin(sh(i)).Count
        If v > 0 Then dpv = q / v
        arr(i + 1, 1) = sh(i): arr(i + 1, 2) = q: arr(i + 1, 3) = v: arr(i + 1, 4) = dpv
    Next i
    calc.Cells(1, C0).Resize(nSh + 1, 4).Value = arr
    AddSeries cht, calc.Cells(1, C0 + 1), calc.Range(calc.Cells(2, C0), calc.Cells(nSh + 1, C0)), _
              calc.Range(calc.Cells(2, C0 + 1), calc.Cells(nSh + 1, C0 + 1)), RGB(31, 78, 121), False, True
End Sub

Private Sub UpdateByModel(ByVal ws As Worksheet, ByVal calc As Worksheet)
    Dim cht As Chart, ks() As String, vs() As Double, n As Long, i As Long, arr() As Variant
    Const C0 As Long = 32
    Set cht = GetChart(ws, "chtByModel")
    ClearSeries cht
    n = SortDictKeys(rModelQty, ks, vs, True)
    If n = 0 Then Exit Sub
    If n > 20 Then n = 20
    ReDim arr(1 To n + 1, 1 To 2)
    arr(1, 1) = "Model": arr(1, 2) = "Defects"
    For i = 1 To n
        arr(i + 1, 1) = ks(i): arr(i + 1, 2) = vs(i)
    Next i
    calc.Cells(1, C0).Resize(n + 1, 2).Value = arr
    AddSeries cht, calc.Cells(1, C0 + 1), calc.Range(calc.Cells(2, C0), calc.Cells(n + 1, C0)), _
              calc.Range(calc.Cells(2, C0 + 1), calc.Cells(n + 1, C0 + 1)), RGB(192, 80, 40), False, True
    On Error Resume Next
    cht.Axes(xlCategory).ReversePlotOrder = True       'lon nhat o tren cung
    On Error GoTo 0
End Sub

Private Sub UpdateTop3Chart(ByVal ws As Worksheet, ByVal calc As Worksheet)
    Dim cht As Chart, nm() As String, ct() As Double, m As Long, i As Long, arr() As Variant
    Const C0 As Long = 35
    Set cht = GetChart(ws, "chtTop3")
    ClearSeries cht
    m = Top3(rCellQty, lCells, nm, ct)
    If m = 0 Then Exit Sub
    ReDim arr(1 To m + 1, 1 To 2)
    arr(1, 1) = "Cell": arr(1, 2) = "Defects"
    For i = 1 To m
        arr(i + 1, 1) = nm(i): arr(i + 1, 2) = ct(i)
    Next i
    calc.Cells(1, C0).Resize(m + 1, 2).Value = arr
    AddSeries cht, calc.Cells(1, C0 + 1), calc.Range(calc.Cells(2, C0), calc.Cells(m + 1, C0)), _
              calc.Range(calc.Cells(2, C0 + 1), calc.Cells(m + 1, C0 + 1)), RGB(192, 57, 43), False, True
    On Error Resume Next
    cht.Axes(xlCategory).ReversePlotOrder = True
    On Error GoTo 0
End Sub

Public Sub UpdateCharts()
    Dim ws As Worksheet, calc As Worksheet
    Set ws = DashSheet(SH_DASH)
    Set calc = DashSheet(SH_CALC)
    calc.Range("F1:AL3000").ClearContents
    gDashStep = "bieu do xu huong theo ca"
    UpdateTrendShift ws, calc
    gDashStep = "bieu do xu huong theo loai loi"
    UpdateTrendDefect ws, calc
    gDashStep = "bieu do theo ca"
    UpdateByShift ws, calc
    gDashStep = "bieu do theo Model"
    UpdateByModel ws, calc
    gDashStep = "bieu do Top 3 vi tri"
    UpdateTop3Chart ws, calc
End Sub
