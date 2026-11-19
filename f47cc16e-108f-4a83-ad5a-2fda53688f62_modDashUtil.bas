Attribute VB_Name = "modDashUtil"
Option Explicit
'=====================================================================
' DASHBOARD - TIEN ICH DUNG CHUNG (hang so bo cuc, Config, mau, sap xep)
' Khong dung UserForm. Chuoi hien thi tieng Viet KHONG DAU.
'=====================================================================

Public Const SH_DASH As String = "Dashboard"
Public Const SH_CFG As String = "Config"
Public Const SH_CALC As String = "DASH_CALC"

Public Const TBL_DSET As String = "tblDashSettings"
Public Const TBL_DCOL As String = "tblDefectColors"

'--- O bo loc tren Dashboard (moi bo loc la 1 cap cot gop o, hang 5; o chinh = o tren-trai) ---
Public Const F_FROM As String = "A5"
Public Const F_TO As String = "C5"
Public Const F_MODEL As String = "E5"
Public Const F_SHIFT As String = "G5"
Public Const F_DEF As String = "I5"
Public Const F_CELL As String = "K5"
Public Const F_TOPN As String = "M5"

'--- O KPI ---
Public Const KPI1_VAL As String = "A9"
Public Const KPI1_SUB As String = "A11"
Public Const KPI2_VAL As String = "D9"
Public Const KPI2_SUB As String = "D11"
Public Const KPI3_VAL As String = "H9"
Public Const KPI3_SUB As String = "H11"
Public Const KPI4_VAL As String = "L9"
Public Const KPI4_SUB As String = "L11"

'--- Hang cac bang tren Dashboard ---
Public Const ROW_SUM_HDR As Long = 30
Public Const ROW_SUM_FIRST As Long = 31
Public Const SUM_MAXROWS As Long = 12
Public Const ROW_T3_HDR As Long = 61
Public Const ROW_T3_FIRST As Long = 62
Public Const T3_MAXROWS As Long = 12
Public Const ROW_MX_HDR As Long = 77
Public Const ROW_MX_FIRST As Long = 78
Public Const MX_MAXROWS As Long = 12
Public Const ROW_DT_HDR As Long = 121
Public Const ROW_DT_FIRST As Long = 122

Public gDashStep As String          'buoc dang chay (hien trong thong bao loi)
Public gDefColors As Object         'DefectType -> RGB (doc tu Config > tblDefectColors)

Private mDepth As Long
Private mSU As Boolean
Private mEV As Boolean

'=====================================================================
' Sheet / Table
'=====================================================================
Public Function DashSheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then DashSheetExists = True: Exit Function
    Next ws
End Function

Public Function DashSheet(ByVal sheetName As String) As Worksheet
    If Not DashSheetExists(sheetName) Then
        Err.Raise vbObjectError + 701, "DashSheet", "Khong tim thay sheet '" & sheetName & "'. Hay chay InitializeDashboard."
    End If
    Set DashSheet = ThisWorkbook.Worksheets(sheetName)
End Function

Public Function DashTryTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set DashTryTable = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
End Function

Public Function DashTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Set DashTable = DashTryTable(sheetName, tableName)
    If DashTable Is Nothing Then
        Err.Raise vbObjectError + 702, "DashTable", "Khong tim thay bang '" & tableName & "' tren sheet '" & sheetName & "'."
    End If
End Function

Public Function DashColIndex(ByVal lo As ListObject, ByVal colName As String, Optional ByVal required As Boolean = True) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then DashColIndex = lc.Index: Exit Function
    Next lc
    If required Then
        Err.Raise vbObjectError + 703, "DashColIndex", "Bang '" & lo.Name & "' thieu cot '" & colName & "'."
    End If
End Function

Public Function DashEnsureTable(ByVal ws As Worksheet, ByVal tblName As String, ByVal topLeft As String, _
                                ByVal headers As Variant, Optional ByVal rowsData As Variant) As ListObject
    Dim lo As ListObject, i As Long, j As Long, nCols As Long, nRows As Long, rng As Range
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    If Not lo Is Nothing Then Set DashEnsureTable = lo: Exit Function

    nCols = UBound(headers) - LBound(headers) + 1
    nRows = 1
    If Not IsMissing(rowsData) Then nRows = UBound(rowsData) - LBound(rowsData) + 1
    Set rng = ws.Range(topLeft).Resize(nRows + 1, nCols)
    For j = 0 To nCols - 1
        rng.Cells(1, j + 1).Value = headers(LBound(headers) + j)
    Next j
    If Not IsMissing(rowsData) Then
        For i = 0 To nRows - 1
            For j = 0 To nCols - 1
                rng.Cells(i + 2, j + 1).Value = rowsData(LBound(rowsData) + i)(j)
            Next j
        Next i
    End If
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleMedium2"
    Set DashEnsureTable = lo
End Function

'=====================================================================
' Trang thai Excel (nhanh + khoi phuc)
'=====================================================================
Public Sub FastOn()
    If mDepth = 0 Then
        mSU = Application.ScreenUpdating
        mEV = Application.EnableEvents
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
    mDepth = mDepth + 1
End Sub

Public Sub FastOff()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth = 0 Then
        Application.EnableEvents = mEV
        Application.ScreenUpdating = mSU
    End If
End Sub

Public Sub ForceRestoreApp()
    'Dung khi VBA bi reset giua chung va Excel con bi tat Events / ScreenUpdating
    mDepth = 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

'=====================================================================
' Dictionary va tong hop
'=====================================================================
Public Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                    'vbTextCompare
    Set NewDict = d
End Function

Public Sub AddTo(ByVal d As Object, ByVal k As String, ByVal v As Double)
    If d.Exists(k) Then
        d(k) = CDbl(d(k)) + v
    Else
        d.Add k, v
    End If
End Sub

Public Sub AddNested(ByVal outer As Object, ByVal k1 As String, ByVal k2 As String, ByVal v As Double)
    Dim inner As Object
    If Not outer.Exists(k1) Then outer.Add k1, NewDict()
    Set inner = outer(k1)
    AddTo inner, k2, v
End Sub

Public Function SortDictKeys(ByVal d As Object, ByRef ks() As String, ByRef vs() As Double, ByVal descending As Boolean) As Long
    'Sap xep khoa theo gia tri (giam dan neu descending), bang nhau thi theo ten tang dan. Tra ve so phan tu.
    Dim n As Long, i As Long, j As Long, k As Variant, tk As String, tv As Double, moveIt As Boolean
    n = d.Count
    If n = 0 Then
        ReDim ks(1 To 1): ReDim vs(1 To 1)
        Exit Function
    End If
    ReDim ks(1 To n): ReDim vs(1 To n)
    i = 0
    For Each k In d.Keys
        i = i + 1
        ks(i) = CStr(k)
        vs(i) = CDbl(d(k))
    Next k
    For i = 2 To n
        tk = ks(i): tv = vs(i)
        j = i - 1
        Do While j >= 1
            If vs(j) = tv Then
                moveIt = (StrComp(ks(j), tk, vbTextCompare) > 0)
            ElseIf descending Then
                moveIt = (vs(j) < tv)
            Else
                moveIt = (vs(j) > tv)
            End If
            If Not moveIt Then Exit Do
            ks(j + 1) = ks(j): vs(j + 1) = vs(j)
            j = j - 1
        Loop
        ks(j + 1) = tk: vs(j + 1) = tv
    Next i
    SortDictKeys = n
End Function

Public Function Top3(ByVal d As Object, ByVal tie As Object, ByRef nm() As String, ByRef ct() As Double) As Long
    'Top 3 khoa co gia tri lon nhat. Bang nhau -> khoa co gia tri tie nho hon (hang truoc, cot truoc) xep truoc.
    Dim r As Long, k As Variant, best As String, bestV As Double, bestT As Double
    Dim found As Boolean, v As Double, t As Double, used As Object
    ReDim nm(1 To 3): ReDim ct(1 To 3)
    Set used = NewDict()
    For r = 1 To 3
        found = False
        For Each k In d.Keys
            If Not used.Exists(CStr(k)) Then
                v = CDbl(d(k))
                t = 0
                If tie.Exists(CStr(k)) Then t = CDbl(tie(CStr(k)))
                If Not found Then
                    best = CStr(k): bestV = v: bestT = t: found = True
                ElseIf v > bestV Or (v = bestV And t < bestT) Then
                    best = CStr(k): bestV = v: bestT = t
                End If
            End If
        Next k
        If Not found Then Exit For
        used.Add best, True
        nm(r) = best
        ct(r) = bestV
        Top3 = r
    Next r
End Function

'=====================================================================
' Config
'=====================================================================
Public Function CfgValue(ByVal settingName As String, ByVal defVal As Variant) As Variant
    Dim lo As ListObject, i As Long
    CfgValue = defVal
    Set lo = DashTryTable(SH_CFG, TBL_DSET)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.DataBodyRange.Rows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, 1).Value)), settingName, vbTextCompare) = 0 Then
            If Len(Trim$(CStr(lo.DataBodyRange.Cells(i, 2).Value))) > 0 Then CfgValue = lo.DataBodyRange.Cells(i, 2).Value
            Exit Function
        End If
    Next i
End Function

'=====================================================================
' Mau theo Defect Type (Config > tblDefectColors, tu them loai loi moi)
'=====================================================================
Public Function PaletteColor(ByVal i As Long) As Long
    Dim p As Variant
    p = Array(RGB(31, 119, 180), RGB(255, 127, 14), RGB(44, 160, 44), RGB(214, 39, 40), RGB(148, 103, 189), _
              RGB(140, 86, 75), RGB(227, 119, 194), RGB(127, 127, 127), RGB(188, 189, 34), RGB(23, 190, 207), _
              RGB(0, 0, 128), RGB(255, 192, 0))
    PaletteColor = CLng(p(i Mod (UBound(p) + 1)))
End Function

Public Sub SyncDefectColors(ByVal names As Object)
    Dim lo As ListObject, i As Long, nm As String, clr As Long, k As Variant, r As ListRow
    Dim body As Range, emptyRow As Long
    Set gDefColors = NewDict()
    Set lo = DashTryTable(SH_CFG, TBL_DCOL)
    If lo Is Nothing Then Exit Sub

    If Not lo.DataBodyRange Is Nothing Then
        Set body = lo.DataBodyRange
        For i = 1 To body.Rows.Count
            nm = Trim$(CStr(body.Cells(i, 1).Value))
            If Len(nm) > 0 Then
                If Not gDefColors.Exists(nm) Then
                    clr = body.Cells(i, 2).Interior.Color
                    If clr = 16777215 Then                   'chua gan mau -> gan mau mac dinh
                        clr = PaletteColor(gDefColors.Count)
                        body.Cells(i, 2).Interior.Color = clr
                    End If
                    gDefColors.Add nm, clr
                End If
            End If
        Next i
    End If

    For Each k In names.Keys
        nm = CStr(k)
        If Not gDefColors.Exists(nm) Then
            clr = PaletteColor(gDefColors.Count)
            emptyRow = 0
            If Not lo.DataBodyRange Is Nothing Then
                For i = 1 To lo.DataBodyRange.Rows.Count
                    If Len(Trim$(CStr(lo.DataBodyRange.Cells(i, 1).Value))) = 0 Then emptyRow = i: Exit For
                Next i
            End If
            If emptyRow > 0 Then
                lo.DataBodyRange.Cells(emptyRow, 1).Value = nm
                lo.DataBodyRange.Cells(emptyRow, 2).Interior.Color = clr
            Else
                Set r = lo.ListRows.Add
                r.Range.Cells(1, 1).Value = nm
                r.Range.Cells(1, 2).Interior.Color = clr
            End If
            gDefColors.Add nm, clr
        End If
    Next k
End Sub

Public Function DefectColor(ByVal defName As String) As Long
    If gDefColors Is Nothing Then Set gDefColors = NewDict()
    If gDefColors.Exists(defName) Then
        DefectColor = CLng(gDefColors(defName))
    Else
        DefectColor = RGB(127, 127, 127)
    End If
End Function
