Option Explicit
'==============================================================================
' TONG HOP DU LIEU BC / CC / PC  ->  TC_Base  va  TC_Primer&Clear  (ThisWorkbook)
'
' GIA DINH DUOC SU DUNG (neu cau truc that khac, sua o day):
'  1. TC_Base: moi bang co 1 dong TIEU DE dang  MA_Ten mau  (vd CE18_SPM-380 ...).
'     Ngay duoi tieu de (toi da 3 dong) la dong header co o "Robot" va cac cot ngay.
'     Cac dong robot R11..R24 va dong SUM nam duoi header, cung cot voi o "Robot".
'  2. TC_Primer&Clear: moi bang bat dau bang 1 o co gia tri la ma robot (R11, R12...).
'     Trong 8 dong ke tiep co dong header chua o "Chemical name" (va "Area").
'     Dong ngay (cac o ngay) nam trong 3 dong tu header tro xuong.
'     Cot Area co the chi ghi o dong dau cua nhom -> macro lay tiep gia tri Area phia tren.
'     Area chua "PRIMER" -> lay du lieu tu file PC ; Area chua "CLEAR" -> lay tu file CC.
'  3. Cac o ngay phai la ngay that (kieu Date hoac so co dinh dang ngay).
'  4. Cfg_Map: hang 1 = ten hoa chat (khop cot Chemical name), tu hang 2 tro xuong la cac Color.
'  5. Cfg_Base: A = tieu de day du cua bang TC_Base, B = Type (co the trong), C = Color.
'     Neu tieu de khong co trong Cfg_Base -> dung ma truoc dau "_" lam Color (khong loc Type);
'     neu ma do xuat hien o nhieu bang -> bo qua va ghi canh bao.
'  6. Hoa chat: MAX rieng tung Color (khong phan biet Type) roi lay MAX giua cac Color.
'==============================================================================

Private Const SH_BASE As String = "TC_Base"
Private Const SH_PC As String = "TC_Primer&Clear"
Private Const SH_MAP As String = "Cfg_Map"
Private Const SH_BASEMAP As String = "Cfg_Base"
Private Const SH_LOG As String = "Run_Log"

Private gLog As Collection
Private gInserted As Object
Private gBlocks As Long
Private gChem As Long
Private gCells As Long
Private gNoData As Long
Private gNoMap As Long
Private gSkipNoR As Long

'==============================================================================
Public Sub Main_UpdateData()
    Dim oCalc As Long
    Dim oScr As Boolean
    Dim oEvt As Boolean
    Dim pBC As String
    Dim pCC As String
    Dim pPC As String
    Dim dictBC As Object
    Dim dictCC As Object
    Dim dictPC As Object
    Dim dToday As Date
    Dim errText As String
    Dim okText As String

    oCalc = Application.Calculation
    oScr = Application.ScreenUpdating
    oEvt = Application.EnableEvents

    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set gLog = New Collection
    Set gInserted = CreateObject("Scripting.Dictionary")
    gBlocks = 0
    gChem = 0
    gCells = 0
    gNoData = 0
    gNoMap = 0
    gSkipNoR = 0

    If Not SelectSourceFiles(pBC, pCC, pPC) Then GoTo CleanUp

    EnsureConfig

    dToday = Date
    Set dictBC = LoadSource(pBC, dToday, "BC")
    Set dictCC = LoadSource(pCC, dToday, "CC")
    Set dictPC = LoadSource(pPC, dToday, "PC")

    If dictBC Is Nothing Then
        AddLog "BC: khong doc duoc -> bo qua TC_Base"
    Else
        UpdateTC_Base dictBC, dToday
    End If
    UpdateTC_PrimerClear dictCC, dictPC, dToday

    WriteRunLog pBC, pCC, pPC, dToday

    okText = "Hoan tat." & vbCrLf & _
             "Ngay xu ly: " & Format(dToday, "dd/mm/yyyy") & vbCrLf & _
             "Bang TC_Base da cap nhat: " & gBlocks & vbCrLf & _
             "Dong hoa chat da xu ly: " & gChem & vbCrLf & _
             "O da ghi gia tri: " & gCells & vbCrLf & _
             "Dong/o khong co du lieu: " & gNoData & vbCrLf & _
             "Thieu anh xa / cau hinh sai: " & gNoMap & vbCrLf & _
             "Bo qua (thieu robot/header): " & gSkipNoR & vbCrLf & _
             "Chi tiet xem sheet " & SH_LOG
    GoTo CleanUp

EH:
    errText = "Loi " & Err.Number & ": " & Err.Description
    On Error Resume Next
    AddLog errText
    WriteRunLog pBC, pCC, pPC, dToday

CleanUp:
    On Error Resume Next
    Application.Calculation = oCalc
    Application.EnableEvents = oEvt
    Application.ScreenUpdating = oScr
    On Error GoTo 0
    If errText <> "" Then
        MsgBox errText, vbCritical, "Update BC/CC/PC"
    ElseIf okText <> "" Then
        MsgBox okText, vbInformation, "Update BC/CC/PC"
    End If
End Sub

'==============================================================================
' TIEN ICH CHUNG
'==============================================================================
Private Function RobotList() As Variant
    RobotList = Array("R11", "R21", "R12", "R22", "R13", "R23", "R14", "R24")
End Function

Private Function IsRobot(ByVal s As String) As Boolean
    Dim i As Long
    Dim a As Variant
    a = RobotList()
    For i = 0 To UBound(a)
        If a(i) = s Then
            IsRobot = True
            Exit Function
        End If
    Next i
End Function

Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = UCase$(s)
End Function

Private Sub AddLog(ByVal s As String)
    If gLog Is Nothing Then Exit Sub
    gLog.Add s
End Sub

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name = nm Then
            SheetExists = True
            Exit Function
        End If
    Next s
End Function

Private Function ToDay(ByVal v As Variant, ByRef d As Long) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        d = Int(CDbl(v))
        ToDay = True
    ElseIf VarType(v) = vbString Then
        If IsDate(v) Then
            d = Int(CDbl(CDate(v)))
            ToDay = True
        End If
    ElseIf IsNumeric(v) Then
        d = Int(CDbl(v))
        ToDay = True
    End If
End Function

Private Function ToNum(ByVal v As Variant, ByRef x As Double) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(CStr(v)) = "" Then Exit Function
    End If
    If IsNumeric(v) Then
        x = CDbl(v)
        ToNum = True
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    LastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Function

Private Function LastUsedCol(ByVal ws As Worksheet) As Long
    LastUsedCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
End Function

'==============================================================================
' CAU HINH
'==============================================================================
Private Sub EnsureConfig()
    Dim ws As Worksheet
    If Not SheetExists(SH_MAP) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_MAP
        ws.Range("A1").Value = "OTO H880-3 WHITE PRIMER"
        ws.Range("B1").Value = "OTO H880-3 DARK GRAY PRIMER"
        ws.Range("C1").Value = "1K SPO 380 CLEAR (VF)"
        ws.Range("A1:C1").Font.Bold = True
        ws.Range("E1").Value = "Moi cot = 1 hoa chat (hang 1 khop cot Chemical name). Hang 2 tro xuong: cac Color cua hoa chat do."
        AddLog "Da tao sheet " & SH_MAP & " - hay dien Color vao tung cot roi chay lai."
    End If
    If Not SheetExists(SH_BASEMAP) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_BASEMAP
        ws.Range("A1").Value = "Block title (day du)"
        ws.Range("B1").Value = "Type"
        ws.Range("C1").Value = "Color"
        ws.Range("A1:C1").Font.Bold = True
        ws.Range("E1").Value = "A = tieu de day du cua bang trong TC_Base; B = Type (co the trong); C = Color."
    End If
End Sub

' Doc Cfg_Map -> Dictionary: ten hoa chat (chuan hoa) -> chuoi "COLOR1|COLOR2|..."
Private Function LoadCfgMap() As Object
    Dim d As Object
    Dim ws As Worksheet
    Dim nr As Long
    Dim nc As Long
    Dim arr As Variant
    Dim c As Long
    Dim r As Long
    Dim nm As String
    Dim col As String
    Dim cur As String

    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(SH_MAP)
    nr = LastUsedRow(ws)
    nc = LastUsedCol(ws)
    If nr >= 1 And nc >= 1 Then
        arr = ws.Range(ws.Cells(1, 1), ws.Cells(nr, nc)).Value
        If IsArray(arr) Then
            For c = 1 To UBound(arr, 2)
                nm = Norm(arr(1, c))
                If nm <> "" Then
                    cur = ""
                    If d.Exists(nm) Then cur = d(nm)
                    For r = 2 To UBound(arr, 1)
                        col = Norm(arr(r, c))
                        If col <> "" Then
                            If cur = "" Then
                                cur = col
                            Else
                                cur = cur & "|" & col
                            End If
                        End If
                    Next r
                    d(nm) = cur
                End If
            Next c
        End If
    End If
    Set LoadCfgMap = d
End Function

' Doc Cfg_Base -> Dictionary: tieu de (chuan hoa) -> "TYPE|COLOR"
Private Function LoadCfgBase() As Object
    Dim d As Object
    Dim ws As Worksheet
    Dim nr As Long
    Dim arr As Variant
    Dim r As Long

    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(SH_BASEMAP)
    nr = LastUsedRow(ws)
    If nr >= 2 Then
        arr = ws.Range(ws.Cells(1, 1), ws.Cells(nr, 3)).Value
        If IsArray(arr) Then
            For r = 2 To UBound(arr, 1)
                If Norm(arr(r, 1)) <> "" Then
                    d(Norm(arr(r, 1))) = Norm(arr(r, 2)) & "|" & Norm(arr(r, 3))
                End If
            Next r
        End If
    End If
    Set LoadCfgBase = d
End Function

'==============================================================================
' CHON FILE NGUON
'==============================================================================
Private Function SelectSourceFiles(ByRef pBC As String, ByRef pCC As String, ByRef pPC As String) As Boolean
    pBC = PickFile("Chon file BC (nguon TC_Base)")
    If pBC = "" Then Exit Function
    pCC = PickFile("Chon file CC (nguon Clear)")
    If pCC = "" Then Exit Function
    pPC = PickFile("Chon file PC (nguon Primer)")
    If pPC = "" Then Exit Function
    SelectSourceFiles = True
End Function

Private Function PickFile(ByVal ttl As String) As String
    Dim v As Variant
    v = Application.GetOpenFilename("Excel / CSV (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , ttl)
    If VarType(v) = vbBoolean Then Exit Function
    PickFile = CStr(v)
End Function

'==============================================================================
' DOC FILE NGUON
'==============================================================================
Private Function FindHeaderRow(ByVal arr As Variant) As Long
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim t As String
    Dim mx As Long
    mx = UBound(arr, 1)
    If mx > 60 Then mx = 60
    For r = 1 To mx
        n = 0
        For c = 1 To UBound(arr, 2)
            t = Norm(arr(r, c))
            If t = "ENTER TIME" Or t = "TYPE" Or t = "COLOR" Or t = "CONSUMPTION" Then n = n + 1
        Next c
        If n >= 4 Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Function BuildColumnMap(ByVal arr As Variant, ByVal hr As Long) As Object
    Dim d As Object
    Dim c As Long
    Dim t As String
    Set d = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(arr, 2)
        t = Norm(arr(hr, c))
        If t <> "" Then
            If Not d.Exists(t) Then d.Add t, c
        End If
    Next c
    Set BuildColumnMap = d
End Function

Private Sub UpdMax(ByVal d As Object, ByVal key As String, ByVal v As Double)
    If d.Exists(key) Then
        If v > d(key) Then d(key) = v
    Else
        d(key) = v
    End If
End Sub

' Key: "TYPE|COLOR|Rxx" va "*|COLOR|Rxx" -> MAX trong ngay xu ly
Private Function LoadSource(ByVal path As String, ByVal dToday As Date, ByVal tag As String) As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim arr As Variant
    Dim hr As Long
    Dim cm As Object
    Dim mx As Object
    Dim i As Long
    Dim k As Long
    Dim cT As Long
    Dim cTy As Long
    Dim cCo As Long
    Dim rc(0 To 7) As Long
    Dim rb As Variant
    Dim dd As Long
    Dim typ As String
    Dim col As String
    Dim v As Double
    Dim miss As String
    Dim todayNum As Long

    On Error GoTo EH
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)
    arr = ws.UsedRange.Value
    wb.Close SaveChanges:=False
    Set wb = Nothing

    If Not IsArray(arr) Then
        AddLog tag & ": file rong"
        Exit Function
    End If
    hr = FindHeaderRow(arr)
    If hr = 0 Then
        AddLog tag & ": khong tim thay dong header (Enter Time / Type / Color / Consumption)"
        Exit Function
    End If
    Set cm = BuildColumnMap(arr, hr)
    If Not cm.Exists("ENTER TIME") Then miss = miss & "Enter Time; "
    If Not cm.Exists("TYPE") Then miss = miss & "Type; "
    If Not cm.Exists("COLOR") Then miss = miss & "Color; "
    If Not cm.Exists("CONSUMPTION") Then miss = miss & "Consumption; "
    If miss <> "" Then
        AddLog tag & ": thieu header bat buoc: " & miss
        Exit Function
    End If
    cT = cm("ENTER TIME")
    cTy = cm("TYPE")
    cCo = cm("COLOR")

    rb = RobotList()
    For k = 0 To 7
        If cm.Exists(rb(k) & " CONSUME") Then
            rc(k) = cm(rb(k) & " CONSUME")
        Else
            rc(k) = 0
            AddLog tag & ": thieu header " & rb(k) & " Consume"
        End If
    Next k

    todayNum = CLng(Int(CDbl(dToday)))
    Set mx = CreateObject("Scripting.Dictionary")
    For i = hr + 1 To UBound(arr, 1)
        If ToDay(arr(i, cT), dd) Then
            If dd = todayNum Then
                typ = Norm(arr(i, cTy))
                col = Norm(arr(i, cCo))
                For k = 0 To 7
                    If rc(k) > 0 Then
                        If ToNum(arr(i, rc(k)), v) Then
                            UpdMax mx, typ & "|" & col & "|" & rb(k), v
                            UpdMax mx, "*|" & col & "|" & rb(k), v
                        End If
                    End If
                Next k
            End If
        End If
    Next i

    AddLog tag & ": da doc " & path & " (" & mx.Count & " nhom khoa trong ngay)"
    Set LoadSource = mx
    Exit Function
EH:
    AddLog tag & ": loi doc file - " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadSource = Nothing
End Function

Private Function GetMax(ByVal mx As Object, ByVal typ As String, ByVal col As String, _
                        ByVal robot As String, ByRef res As Double) As Boolean
    Dim key As String
    If mx Is Nothing Then Exit Function
    If typ = "" Then
        key = "*|" & col & "|" & robot
    Else
        key = typ & "|" & col & "|" & robot
    End If
    If mx.Exists(key) Then
        res = mx(key)
        GetMax = True
    End If
End Function

'==============================================================================
' COT NGAY
'==============================================================================
Private Function IsDateCell(ByVal c As Range) As Boolean
    Dim v As Variant
    Dim f As String
    v = c.Value
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        IsDateCell = True
        Exit Function
    End If
    If VarType(v) = vbDouble Then
        f = LCase$(c.NumberFormat)
        If v > 30000 Then
            If InStr(f, "d") > 0 Or InStr(f, "m") > 0 Or InStr(f, "y") > 0 Then IsDateCell = True
        End If
    End If
End Function

Private Function GetDateCols(ByVal ws As Worksheet, ByVal hr As Long, ByVal fromCol As Long) As Collection
    Dim cl As Collection
    Dim c As Long
    Dim lc As Long
    Set cl = New Collection
    lc = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    If fromCol < 1 Then fromCol = 1
    For c = fromCol To lc
        If IsDateCell(ws.Cells(hr, c)) Then cl.Add c
    Next c
    Set GetDateCols = cl
End Function

' Tra ve cot ngay (co san hoac chen moi ngay sau cot ngay cuoi). 0 neu khong xac dinh duoc.
Private Function EnsureDateColumn(ByVal ws As Worksheet, ByVal hr As Long, ByVal fromCol As Long, ByVal d As Date) As Long
    Dim cl As Collection
    Dim i As Long
    Dim lastC As Long
    Dim newC As Long
    Dim nf As String
    Dim key As String
    Dim target As Long

    Set cl = GetDateCols(ws, hr, fromCol)
    If cl.Count = 0 Then Exit Function
    target = CLng(Int(CDbl(d)))
    For i = 1 To cl.Count
        If CLng(Int(CDbl(ws.Cells(hr, cl(i)).Value))) = target Then
            EnsureDateColumn = cl(i)
            Exit Function
        End If
    Next i

    lastC = cl(cl.Count)
    nf = ws.Cells(hr, lastC).NumberFormat
    newC = lastC + 1
    key = ws.Name & "|" & newC
    If gInserted.Exists(key) And IsEmpty(ws.Cells(hr, newC).Value) Then
        ' cot nay da duoc chen khi xu ly bang khac -> dung lai
    Else
        ws.Columns(newC).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        gInserted(key) = True
    End If
    ws.Cells(hr, newC).Value = CDbl(target)
    ws.Cells(hr, newC).NumberFormat = nf
    EnsureDateColumn = newC
End Function

Private Function BuildAddr(ByVal ws As Worksheet, ByVal r As Long, ByVal cl As Collection) As String
    Dim i As Long
    Dim contiguous As Boolean
    Dim s As String
    If cl.Count = 0 Then Exit Function
    contiguous = True
    For i = 2 To cl.Count
        If cl(i) <> cl(i - 1) + 1 Then contiguous = False
    Next i
    If contiguous Then
        BuildAddr = ws.Range(ws.Cells(r, cl(1)), ws.Cells(r, cl(cl.Count))).Address(False, False)
    Else
        For i = 1 To cl.Count
            If s <> "" Then s = s & ","
            s = s & ws.Cells(r, cl(i)).Address(False, False)
        Next i
        BuildAddr = s
    End If
End Function

Private Function FindColInRows(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal names As Variant) As Long
    Dim r As Long
    Dim c As Long
    Dim lc As Long
    Dim i As Long
    Dim t As String
    lc = LastUsedCol(ws)
    For r = r1 To r2
        For c = 1 To lc
            t = Norm(ws.Cells(r, c).Value)
            If t <> "" Then
                For i = LBound(names) To UBound(names)
                    If t = Norm(names(i)) Then
                        FindColInRows = c
                        Exit Function
                    End If
                Next i
            End If
        Next c
    Next r
End Function

'==============================================================================
' TC_BASE
'==============================================================================
Private Function IsTitleText(ByVal s As String) As Boolean
    Dim p As Long
    Dim code As String
    Dim i As Long
    Dim ch As String
    Dim hasDigit As Boolean
    p = InStr(s, "_")
    If p < 3 Then Exit Function
    If p >= Len(s) Then Exit Function
    code = Left$(s, p - 1)
    If Len(code) > 8 Then Exit Function
    For i = 1 To Len(code)
        ch = Mid$(code, i, 1)
        If ch >= "0" And ch <= "9" Then
            hasDigit = True
        ElseIf Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")) Then
            Exit Function
        End If
    Next i
    IsTitleText = hasDigit
End Function

Private Sub UpdateTC_Base(ByVal dictBC As Object, ByVal dToday As Date)
    Dim ws As Worksheet
    Dim cfg As Object
    Dim cnt As Object
    Dim tRow As Collection
    Dim tText As Collection
    Dim lastRow As Long
    Dim scanCols As Long
    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim n As Long
    Dim s As String
    Dim code As String
    Dim v As Variant
    Dim endR As Long

    If Not SheetExists(SH_BASE) Then
        AddLog "Khong co sheet " & SH_BASE
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SH_BASE)
    Set cfg = LoadCfgBase()
    Set cnt = CreateObject("Scripting.Dictionary")
    Set tRow = New Collection
    Set tText = New Collection

    lastRow = LastUsedRow(ws)
    scanCols = LastUsedCol(ws)
    If scanCols > 15 Then scanCols = 15

    For r = 1 To lastRow
        For c = 1 To scanCols
            v = ws.Cells(r, c).Value
            If VarType(v) = vbString Then
                s = Trim$(CStr(v))
                If IsTitleText(s) Then
                    tRow.Add r
                    tText.Add s
                    code = Norm(Left$(s, InStr(s, "_") - 1))
                    If cnt.Exists(code) Then
                        cnt(code) = cnt(code) + 1
                    Else
                        cnt(code) = 1
                    End If
                    Exit For
                End If
            End If
        Next c
    Next r

    n = tRow.Count
    If n = 0 Then
        AddLog SH_BASE & ": khong tim thay tieu de bang (dang MA_Ten mau)"
        Exit Sub
    End If

    For i = 1 To n
        If i < n Then
            endR = CLng(tRow(i + 1)) - 1
        Else
            endR = lastRow
        End If
        ProcessBaseTable ws, dictBC, dToday, cfg, cnt, CLng(tRow(i)), CStr(tText(i)), endR
    Next i
End Sub

Private Sub ProcessBaseTable(ByVal ws As Worksheet, ByVal dictBC As Object, ByVal dToday As Date, _
                             ByVal cfg As Object, ByVal cnt As Object, ByVal titleRow As Long, _
                             ByVal title As String, ByVal endRow As Long)
    Dim hdrRow As Long
    Dim robotCol As Long
    Dim r As Long
    Dim c As Long
    Dim lc As Long
    Dim code As String
    Dim typ As String
    Dim col As String
    Dim key As String
    Dim dc As Long
    Dim lab As String
    Dim sumRow As Long
    Dim addr As String
    Dim v As Double
    Dim nRobot As Long
    Dim probe As Long
    Dim cl As Collection

    code = Norm(Left$(title, InStr(title, "_") - 1))
    typ = ""
    col = code

    If cfg.Exists(Norm(title)) Then
        key = cfg(Norm(title))
        typ = Left$(key, InStr(key, "|") - 1)
        If Mid$(key, InStr(key, "|") + 1) <> "" Then col = Mid$(key, InStr(key, "|") + 1)
    ElseIf cnt(code) > 1 Then
        gNoMap = gNoMap + 1
        AddLog SH_BASE & ": ma mau " & code & " xuat hien o nhieu bang - can khai bao trong " & SH_BASEMAP & " (bo qua bang '" & title & "')"
        Exit Sub
    End If

    ' tim header co o "Robot" trong 3 dong duoi tieu de
    lc = LastUsedCol(ws)
    If lc > 15 Then lc = 15
    For probe = titleRow + 1 To titleRow + 3
        If probe > endRow Then Exit For
        For c = 1 To lc
            If Norm(ws.Cells(probe, c).Value) = "ROBOT" Then
                hdrRow = probe
                robotCol = c
                Exit For
            End If
        Next c
        If hdrRow > 0 Then Exit For
    Next probe
    If hdrRow = 0 Then
        gSkipNoR = gSkipNoR + 1
        AddLog SH_BASE & ": khong tim thay header 'Robot' duoi tieu de '" & title & "'"
        Exit Sub
    End If

    dc = EnsureDateColumn(ws, hdrRow, robotCol + 1, dToday)
    If dc = 0 Then
        gSkipNoR = gSkipNoR + 1
        AddLog SH_BASE & ": khong tim thay cot ngay trong bang '" & title & "'"
        Exit Sub
    End If

    sumRow = 0
    addr = ""
    nRobot = 0
    For r = hdrRow + 1 To endRow
        lab = Norm(ws.Cells(r, robotCol).Value)
        If IsRobot(lab) Then
            nRobot = nRobot + 1
            ws.Cells(r, dc).ClearContents
            If GetMax(dictBC, typ, col, lab, v) Then
                ws.Cells(r, dc).Value = v
                gCells = gCells + 1
            Else
                gNoData = gNoData + 1
                AddLog SH_BASE & " [" & title & "] " & lab & ": khong co du lieu ngay " & Format(dToday, "dd/mm/yyyy")
            End If
            If addr <> "" Then addr = addr & ","
            addr = addr & ws.Cells(r, dc).Address(False, False)
        ElseIf lab = "SUM" Then
            sumRow = r
        End If
    Next r

    If nRobot = 0 Then
        gSkipNoR = gSkipNoR + 1
        AddLog SH_BASE & ": bang '" & title & "' khong co dong robot"
        Exit Sub
    End If
    If sumRow > 0 Then
        ws.Cells(sumRow, dc).Formula = "=IF(COUNT(" & addr & ")=0,"""",SUM(" & addr & "))"
    End If
    gBlocks = gBlocks + 1
End Sub

'==============================================================================
' TC_PRIMER&CLEAR
'==============================================================================
Private Sub UpdateTC_PrimerClear(ByVal dictCC As Object, ByVal dictPC As Object, ByVal dToday As Date)
    Dim ws As Worksheet
    Dim cm As Object
    Dim claimed As Object
    Dim tR As Collection
    Dim tH As Collection
    Dim tC As Collection
    Dim tCode As Collection
    Dim arr As Variant
    Dim ur As Range
    Dim baseRow As Long
    Dim baseCol As Long
    Dim lastRow As Long
    Dim lc As Long
    Dim r As Long
    Dim c As Long
    Dim p As Long
    Dim n As Long
    Dim i As Long
    Dim hr As Long
    Dim chemCol As Long
    Dim t As String
    Dim hk As String
    Dim endR As Long

    If Not SheetExists(SH_PC) Then
        AddLog "Khong co sheet " & SH_PC
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets(SH_PC)
    Set cm = LoadCfgMap()
    If cm.Count = 0 Then
        AddLog SH_MAP & ": chua co du lieu - khong cap nhat " & SH_PC
        Exit Sub
    End If

    Set claimed = CreateObject("Scripting.Dictionary")
    Set tR = New Collection
    Set tH = New Collection
    Set tC = New Collection
    Set tCode = New Collection

    Set ur = ws.UsedRange
    baseRow = ur.Row - 1
    baseCol = ur.Column - 1
    lastRow = LastUsedRow(ws)
    lc = LastUsedCol(ws)
    arr = ur.Value
    If Not IsArray(arr) Then
        AddLog SH_PC & ": sheet rong"
        Exit Sub
    End If

    ' tim cac o ma robot -> header Chemical name trong 8 dong ke tiep
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            t = Norm(arr(r, c))
            If IsRobot(t) Then
                hr = 0
                chemCol = 0
                For p = r To r + 8
                    If p > UBound(arr, 1) Then Exit For
                    For i = 1 To UBound(arr, 2)
                        If Norm(arr(p, i)) = "CHEMICAL NAME" Then
                            hr = p
                            chemCol = i
                            Exit For
                        End If
                    Next i
                    If hr > 0 Then Exit For
                Next p
                If hr = 0 Then
                    gSkipNoR = gSkipNoR + 1
                    AddLog SH_PC & ": o robot " & t & " (dong " & (r + baseRow) & ") khong co header 'Chemical name' ngay duoi"
                Else
                    hk = hr & "|" & chemCol
                    If Not claimed.Exists(hk) Then
                        claimed(hk) = True
                        tR.Add r + baseRow
                        tH.Add hr + baseRow
                        tC.Add chemCol + baseCol
                        tCode.Add t
                    End If
                End If
            End If
        Next c
    Next r

    n = tR.Count
    If n = 0 Then
        AddLog SH_PC & ": khong tim thay bang robot nao"
        Exit Sub
    End If

    For i = 1 To n
        If i < n Then
            endR = CLng(tR(i + 1)) - 1
        Else
            endR = lastRow
        End If
        ProcessPrimerTable ws, cm, dictCC, dictPC, dToday, CStr(tCode(i)), CLng(tH(i)), CLng(tC(i)), endR
    Next i
End Sub

Private Sub ProcessPrimerTable(ByVal ws As Worksheet, ByVal cm As Object, ByVal dictCC As Object, _
                               ByVal dictPC As Object, ByVal dToday As Date, ByVal robot As String, _
                               ByVal hr As Long, ByVal chemCol As Long, ByVal endRow As Long)
    Dim areaCol As Long
    Dim dRow As Long
    Dim dc As Long
    Dim avgCol As Long
    Dim p As Long
    Dim c As Long
    Dim r As Long
    Dim k As Long
    Dim cl As Collection
    Dim curArea As String
    Dim a As String
    Dim nm As String
    Dim src As Object
    Dim srcName As String
    Dim colors As Variant
    Dim v As Double
    Dim best As Double
    Dim found As Boolean
    Dim addr As String
    Dim lc As Long

    ' cot Area (cung dong header, ben trai hoac tai cot Chemical name)
    For c = 1 To chemCol
        If Norm(ws.Cells(hr, c).Value) = "AREA" Then
            areaCol = c
            Exit For
        End If
    Next c
    If areaCol = 0 Then
        gSkipNoR = gSkipNoR + 1
        AddLog SH_PC & " [" & robot & "]: khong tim thay cot 'Area' tren dong " & hr & " - bo qua bang"
        Exit Sub
    End If

    ' dong ngay: trong 3 dong tu header, dong co o ngay ben phai cot Chemical name
    For p = hr To hr + 3
        If p > endRow Then Exit For
        If GetDateCols(ws, p, chemCol + 1).Count > 0 Then
            dRow = p
            Exit For
        End If
    Next p
    If dRow = 0 Then
        gSkipNoR = gSkipNoR + 1
        AddLog SH_PC & " [" & robot & "]: khong tim thay dong ngay - bo qua bang"
        Exit Sub
    End If

    dc = EnsureDateColumn(ws, dRow, chemCol + 1, dToday)
    If dc = 0 Then
        gSkipNoR = gSkipNoR + 1
        AddLog SH_PC & " [" & robot & "]: khong xac dinh duoc cot ngay - bo qua bang"
        Exit Sub
    End If

    avgCol = FindColInRows(ws, hr, dRow, Array("Trung b" & ChrW(236) & "nh", "Trung binh", "Average"))
    Set cl = GetDateCols(ws, dRow, chemCol + 1)

    curArea = ""
    For r = dRow + 1 To endRow
        a = Norm(ws.Cells(r, areaCol).Value)
        If a <> "" Then curArea = a
        nm = Norm(ws.Cells(r, chemCol).Value)
        If nm <> "" Then
            Set src = Nothing
            srcName = ""
            If InStr(curArea, "PRIMER") > 0 Then
                Set src = dictPC
                srcName = "PC"
            ElseIf InStr(curArea, "CLEAR") > 0 Then
                Set src = dictCC
                srcName = "CC"
            End If

            If srcName = "" Then
                gNoMap = gNoMap + 1
                AddLog SH_PC & " [" & robot & "] dong " & r & ": Area '" & curArea & "' khong phai PRIMER/CLEAR - bo qua"
            ElseIf Not cm.Exists(nm) Then
                gNoMap = gNoMap + 1
                AddLog SH_PC & " [" & robot & "] dong " & r & ": '" & ws.Cells(r, chemCol).Value & "' khong co trong " & SH_MAP
            ElseIf cm(nm) = "" Then
                gNoMap = gNoMap + 1
                AddLog SH_PC & " [" & robot & "] dong " & r & ": '" & ws.Cells(r, chemCol).Value & "' chua khai bao Color trong " & SH_MAP
            ElseIf src Is Nothing Then
                gNoData = gNoData + 1
                AddLog SH_PC & " [" & robot & "] dong " & r & ": file " & srcName & " khong doc duoc"
            Else
                colors = Split(cm(nm), "|")
                found = False
                best = 0
                For k = LBound(colors) To UBound(colors)
                    If GetMax(src, "", CStr(colors(k)), robot, v) Then
                        If Not found Then
                            best = v
                            found = True
                        ElseIf v > best Then
                            best = v
                        End If
                    End If
                Next k

                gChem = gChem + 1
                ws.Cells(r, dc).ClearContents
                If found Then
                    ws.Cells(r, dc).Value = best
                    gCells = gCells + 1
                Else
                    gNoData = gNoData + 1
                    AddLog SH_PC & " [" & robot & "] dong " & r & ": khong co du lieu " & srcName & " cho '" & ws.Cells(r, chemCol).Value & "'"
                End If

                If avgCol > 0 Then
                    addr = BuildAddr(ws, r, cl)
                    If addr <> "" Then
                        ws.Cells(r, avgCol).Formula = "=IF(COUNT(" & addr & ")=0,"""",AVERAGE(" & addr & "))"
                    End If
                End If
            End If
        End If
    Next r
End Sub

'==============================================================================
' NHAT KY
'==============================================================================
Private Sub WriteRunLog(ByVal pBC As String, ByVal pCC As String, ByVal pPC As String, ByVal dToday As Date)
    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long

    If SheetExists(SH_LOG) Then
        Set ws = ThisWorkbook.Worksheets(SH_LOG)
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
    End If

    ws.Range("A1").Value = "Thoi diem chay"
    ws.Range("B1").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
    ws.Range("A2").Value = "Ngay xu ly"
    ws.Range("B2").Value = Format(dToday, "dd/mm/yyyy")
    ws.Range("A3").Value = "File BC"
    ws.Range("B3").Value = pBC
    ws.Range("A4").Value = "File CC"
    ws.Range("B4").Value = pCC
    ws.Range("A5").Value = "File PC"
    ws.Range("B5").Value = pPC
    ws.Range("A6").Value = "Bang TC_Base da cap nhat"
    ws.Range("B6").Value = gBlocks
    ws.Range("A7").Value = "Dong hoa chat da xu ly"
    ws.Range("B7").Value = gChem
    ws.Range("A8").Value = "O da ghi du lieu"
    ws.Range("B8").Value = gCells
    ws.Range("A9").Value = "Dong/o khong co du lieu phu hop"
    ws.Range("B9").Value = gNoData
    ws.Range("A10").Value = "Thieu anh xa / cau hinh khong hop le"
    ws.Range("B10").Value = gNoMap
    ws.Range("A11").Value = "Bo qua do thieu robot / header"
    ws.Range("B11").Value = gSkipNoR
    ws.Range("A13").Value = "Canh bao / chi tiet"
    ws.Range("A13").Font.Bold = True

    r = 14
    If Not gLog Is Nothing Then
        For i = 1 To gLog.Count
            ws.Cells(r, 1).Value = gLog(i)
            r = r + 1
        Next i
    End If
    ws.Columns("A:B").AutoFit
End Sub

'==============================================================================
' VI DU KIEM THU MAX DOC LAP (cung ngay hom nay, Color = CE18):
'   Consumption  R11  R21
'   850          150  160
'   888          157  155
'   870          155  169
' Ket qua: R11 = 157 (dong 2), R21 = 169 (dong 3) - MAX tung cot, khong tu 1 dong.
' Hoa chat nhieu Color: R11 cua 3 Color la 100, 120, 110 -> ghi 120 (khong cong).
'==============================================================================
