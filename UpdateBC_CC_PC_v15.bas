Option Explicit
'==============================================================================
' TONG HOP DU LIEU BC / CC / PC  ->  TC_Base  va  TC_Primer&Clear  (ThisWorkbook)
'
' QUY TAC:
'  - LOC THEO TYPE LA BAT BUOC: chi doc cac dong co Type nam trong TYPE_LIST (1 hoac 2 Type,
'    2 Type duoc xem la 1 -> MAX chung). Khong lay dong cua Type khac.
'  - Khong loc theo ngay trong file nguon. Gia tri ghi vao cot co ngay chay macro (Date).
'  - Moi robot: MAX rieng cua cot "Rxx Consume" trong cac dong thoa Type + Color.
'  - TC_Base: Color = ma truoc dau "_" trong tieu de bang (vd CE18_SPM-380...).
'  - TC_Primer&Clear: Area PRIMER -> file PC, Area CLEAR -> file CC. Color lay tu Cfg_Map.
'    Hoa chat nhieu Color: MAX tung Color roi lay MAX giua cac Color.
'    Bang co o tieu de "ROBOT" = bang tong: cong cac bang R11, R12... theo tung hoa chat.
'  - Cfg_Map: hang 1 = ten hoa chat (khop cot Chemical name); hang 2 tro xuong = cac Color.
'==============================================================================

Private Const SH_BASE As String = "TC_Base"
Private Const SH_PC As String = "TC_Primer&Clear"
Private Const SH_MAP As String = "Cfg_Map"
Private Const SH_LOG As String = "Run_Log"

' ===== TYPE CAN LOC (SUA O DAY) =====
' Vi du: "T1"  hoac  "T1,T2".  De trong "" -> macro hien hop thoai hoi Type moi lan chay.
Private Const TYPE_LIST As String = ""

Private gLog As Collection
Private gTypes As Object
Private gTypeText As String
Private gInserted As Object
Private gChemAddr As Object
Private gBlocks As Long
Private gChem As Long
Private gCells As Long
Private gNoData As Long
Private gNoMap As Long
Private gSkip As Long

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
    Set gChemAddr = CreateObject("Scripting.Dictionary")
    gBlocks = 0
    gChem = 0
    gCells = 0
    gNoData = 0
    gNoMap = 0
    gSkip = 0

    If Not ResolveTypes() Then GoTo CleanUp
    EnsureConfig
    If Not SelectSourceFiles(pBC, pCC, pPC) Then GoTo CleanUp

    dToday = Date
    Set dictBC = LoadSource(pBC, "BC")
    Set dictCC = LoadSource(pCC, "CC")
    Set dictPC = LoadSource(pPC, "PC")

    If dictBC Is Nothing Then
        AddLog "BC: khong doc duoc -> bo qua TC_Base"
    Else
        UpdateTC_Base dictBC, dToday
    End If
    UpdateTC_PrimerClear dictCC, dictPC, dToday

    WriteRunLog pBC, pCC, pPC, dToday

    okText = "Hoan tat." & vbCrLf & _
             "Type da loc: " & gTypeText & vbCrLf & _
             "Ngay ghi: " & Format(dToday, "dd/mm/yyyy") & vbCrLf & _
             "Bang TC_Base da cap nhat: " & gBlocks & vbCrLf & _
             "Dong hoa chat da xu ly: " & gChem & vbCrLf & _
             "O da ghi gia tri: " & gCells & vbCrLf & _
             "Dong/o khong co du lieu: " & gNoData & vbCrLf & _
             "Thieu anh xa / cau hinh: " & gNoMap & vbCrLf & _
             "Bo qua (thieu robot/header): " & gSkip & vbCrLf & _
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
' TIEN ICH
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

Private Sub AddAddr(ByVal d As Object, ByVal key As String, ByVal addr As String)
    If d.Exists(key) Then
        d(key) = d(key) & "," & addr
    Else
        d(key) = addr
    End If
End Sub

' Ghi cong thuc tong cac o trong danh sach dia chi (cung sheet); neu qua nhieu o thi ghi gia tri.
Private Sub WriteSumOf(ByVal ws As Worksheet, ByVal target As Range, ByVal addrList As String)
    Dim parts As Variant
    Dim i As Long
    Dim x As Double
    Dim tot As Double
    Dim n As Long
    parts = Split(addrList, ",")
    If UBound(parts) < 250 Then
        target.Formula = "=IF(COUNT(" & addrList & ")=0,"""",SUM(" & addrList & "))"
    Else
        For i = 0 To UBound(parts)
            If ToNum(ws.Range(parts(i)).Value, x) Then
                tot = tot + x
                n = n + 1
            End If
        Next i
        If n > 0 Then
            target.Value = tot
        Else
            target.ClearContents
        End If
    End If
End Sub

'==============================================================================
' TYPE VA CAU HINH
'==============================================================================
Private Function ResolveTypes() As Boolean
    Dim txt As String
    Dim parts As Variant
    Dim i As Long
    Dim t As String

    txt = TYPE_LIST
    If Trim$(txt) = "" Then
        txt = InputBox("Nhap Type can loc (1 Type, hoac 2 Type cach nhau dau phay)." & vbCrLf & _
                       "Vi du: T1   hoac   T1,T2", "Loc theo Type")
    End If
    Set gTypes = CreateObject("Scripting.Dictionary")
    parts = Split(txt, ",")
    For i = LBound(parts) To UBound(parts)
        t = Norm(parts(i))
        If t <> "" Then gTypes(t) = True
    Next i
    If gTypes.Count = 0 Then
        MsgBox "Chua nhap Type. Macro dung, khong ghi du lieu.", vbExclamation, "Loc theo Type"
        Exit Function
    End If
    gTypeText = Join(gTypes.Keys, ",")
    ResolveTypes = True
End Function

Private Sub EnsureConfig()
    Dim ws As Worksheet
    If SheetExists(SH_MAP) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SH_MAP
    ws.Range("A1").Value = "OTO H880-3 WHITE PRIMER"
    ws.Range("B1").Value = "OTO H880-3 DARK GRAY PRIMER"
    ws.Range("C1").Value = "1K SPO 380 CLEAR (VF)"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("L1").Value = "Moi cot = 1 hoa chat (hang 1 khop cot Chemical name). Hang 2 tro xuong: cac Color cua hoa chat do."
    AddLog "Da tao sheet " & SH_MAP & " - hay dien Color vao tung cot roi chay lai."
End Sub

' Cfg_Map -> Dictionary: ten hoa chat (chuan hoa) -> "COLOR1|COLOR2|..."
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
                If nm <> "" And nm <> "TYPE" Then
                    cur = ""
                    If d.Exists(nm) Then cur = d(nm)
                    For r = 2 To UBound(arr, 1)
                        col = Norm(arr(r, c))
                        If col <> "" Then
                            If cur <> "" Then cur = cur & "|"
                            cur = cur & col
                        End If
                    Next r
                    d(nm) = cur
                End If
            Next c
        End If
    End If
    Set LoadCfgMap = d
End Function

'==============================================================================
' CHON FILE NGUON (1 thu muc chua BC, CC, PC)
'==============================================================================
Private Function SelectSourceFiles(ByRef pBC As String, ByRef pCC As String, ByRef pPC As String) As Boolean
    Dim fso As Object
    Dim folder As String
    Dim prev As String
    Dim ans As Long
    Dim miss As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If SheetExists(SH_LOG) Then
        prev = Trim$(CStr(ThisWorkbook.Worksheets(SH_LOG).Range("B3").Value))
        If prev <> "" Then folder = fso.GetParentFolderName(prev)
    End If
    If folder <> "" Then
        If fso.FolderExists(folder) Then
            ans = MsgBox("Dung thu muc nguon lan truoc?" & vbCrLf & folder & vbCrLf & vbCrLf & _
                         "Yes = dung thu muc nay, No = chon thu muc khac", vbYesNoCancel + vbQuestion, "Thu muc BC / CC / PC")
            If ans = vbCancel Then Exit Function
            If ans = vbNo Then folder = ""
        Else
            folder = ""
        End If
    End If
    If folder = "" Then
        folder = PickFolder("Chon thu muc chua file BC, CC, PC")
        If folder = "" Then Exit Function
    End If

    pBC = FindSourceFile(folder, "BC")
    pCC = FindSourceFile(folder, "CC")
    pPC = FindSourceFile(folder, "PC")

    If pBC = "" Then miss = miss & "BC "
    If pCC = "" Then miss = miss & "CC "
    If pPC = "" Then miss = miss & "PC "
    If miss <> "" Then
        MsgBox "Khong tim thay trong thu muc: " & miss & vbCrLf & "Hay chon thu cong cac file con thieu.", vbExclamation
        If pBC = "" Then pBC = PickFile("Chon file BC (nguon TC_Base)")
        If pBC = "" Then Exit Function
        If pCC = "" Then pCC = PickFile("Chon file CC (nguon Clear)")
        If pCC = "" Then Exit Function
        If pPC = "" Then pPC = PickFile("Chon file PC (nguon Primer)")
        If pPC = "" Then Exit Function
    End If
    SelectSourceFiles = True
End Function

Private Function PickFolder(ByVal ttl As String) As String
    With Application.FileDialog(4)
        .Title = ttl
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

Private Function PickFile(ByVal ttl As String) As String
    Dim v As Variant
    v = Application.GetOpenFilename("Excel / CSV (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , ttl)
    If VarType(v) = vbBoolean Then Exit Function
    PickFile = CStr(v)
End Function

' File co ten (khong tinh duoi) = nm, khong phan biet hoa/thuong. Nhieu file -> lay file moi nhat.
Private Function FindSourceFile(ByVal folder As String, ByVal nm As String) As String
    Dim fso As Object
    Dim f As Object
    Dim ext As String
    Dim bestDate As Date
    Dim best As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Function
    For Each f In fso.GetFolder(folder).Files
        If UCase$(fso.GetBaseName(f.Name)) = UCase$(nm) Then
            ext = LCase$(fso.GetExtensionName(f.Name))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "csv" Then
                If best = "" Or f.DateLastModified > bestDate Then
                    best = f.Path
                    bestDate = f.DateLastModified
                End If
            End If
        End If
    Next f
    FindSourceFile = best
End Function

'==============================================================================
' DOC FILE NGUON  ->  Dictionary  "COLOR|Rxx" -> MAX  (chi cac dong thuoc Type da chon)
'==============================================================================
Private Function FindHeaderRow(arr As Variant) As Long
    Dim r As Long
    Dim c As Long
    Dim mx As Long
    Dim hasT As Boolean
    Dim hasC As Boolean
    Dim t As String
    mx = UBound(arr, 1)
    If mx > 60 Then mx = 60
    For r = 1 To mx
        hasT = False
        hasC = False
        For c = 1 To UBound(arr, 2)
            t = Norm(arr(r, c))
            If t = "TYPE" Then hasT = True
            If t = "COLOR" Then hasC = True
        Next c
        If hasT And hasC Then
            FindHeaderRow = r
            Exit Function
        End If
    Next r
End Function

Private Sub UpdMax(ByVal d As Object, ByVal key As String, ByVal v As Double)
    If d.Exists(key) Then
        If v > d(key) Then d(key) = v
    Else
        d(key) = v
    End If
End Sub

Private Function LoadSource(ByVal path As String, ByVal tag As String) As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim arr As Variant
    Dim hr As Long
    Dim cm As Object
    Dim mx As Object
    Dim seen As Object
    Dim i As Long
    Dim k As Long
    Dim c As Long
    Dim cTy As Long
    Dim cCo As Long
    Dim rc(0 To 7) As Long
    Dim rb As Variant
    Dim typ As String
    Dim col As String
    Dim t As String
    Dim v As Double
    Dim nRows As Long
    Dim missR As String

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
        AddLog tag & ": khong tim thay dong header co cot Type va Color"
        Exit Function
    End If

    Set cm = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(arr, 2)
        t = Norm(arr(hr, c))
        If t <> "" Then
            If Not cm.Exists(t) Then cm.Add t, c
        End If
    Next c
    cTy = cm("TYPE")
    cCo = cm("COLOR")

    rb = RobotList()
    For k = 0 To 7
        If cm.Exists(rb(k) & " CONSUME") Then
            rc(k) = cm(rb(k) & " CONSUME")
        Else
            rc(k) = 0
            missR = missR & rb(k) & " "
        End If
    Next k
    If missR <> "" Then AddLog tag & ": khong co cot Consume cua robot: " & missR

    Set mx = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    For i = hr + 1 To UBound(arr, 1)
        typ = Norm(arr(i, cTy))
        col = Norm(arr(i, cCo))
        If typ <> "" Then seen(typ) = True
        If col <> "" And gTypes.Exists(typ) Then
            nRows = nRows + 1
            For k = 0 To 7
                If rc(k) > 0 Then
                    If ToNum(arr(i, rc(k)), v) Then UpdMax mx, col & "|" & rb(k), v
                End If
            Next k
        End If
    Next i

    AddLog tag & ": " & path
    If seen.Count > 0 Then AddLog tag & ": Type co trong file: " & Join(seen.Keys, ", ")
    If nRows = 0 Then
        AddLog tag & ": CANH BAO - khong co dong nao thuoc Type " & gTypeText
    Else
        AddLog tag & ": " & nRows & " dong thuoc Type " & gTypeText & " -> " & mx.Count & " to hop Color-Robot"
    End If
    Set LoadSource = mx
    Exit Function
EH:
    AddLog tag & ": loi doc file - " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadSource = Nothing
End Function

Private Function GetMax(ByVal mx As Object, ByVal col As String, ByVal robot As String, ByRef res As Double) As Boolean
    Dim key As String
    If mx Is Nothing Then Exit Function
    key = col & "|" & robot
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

' Tra ve cot ngay chay macro (co san, hoac chen moi sau cot ngay cuoi). 0 neu khong xac dinh duoc.
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
' TC_BASE  (cac bang theo mau)
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
        ProcessBaseTable ws, dictBC, dToday, cnt, CLng(tRow(i)), CStr(tText(i)), endR
    Next i
End Sub

Private Sub ProcessBaseTable(ByVal ws As Worksheet, ByVal dictBC As Object, ByVal dToday As Date, _
                             ByVal cnt As Object, ByVal titleRow As Long, _
                             ByVal title As String, ByVal endRow As Long)
    Dim hdrRow As Long
    Dim robotCol As Long
    Dim r As Long
    Dim c As Long
    Dim lc As Long
    Dim col As String
    Dim dc As Long
    Dim lab As String
    Dim sumRow As Long
    Dim addr As String
    Dim v As Double
    Dim nRobot As Long
    Dim probe As Long

    col = Norm(Left$(title, InStr(title, "_") - 1))
    If cnt(col) > 1 Then AddLog SH_BASE & ": Color " & col & " co o " & cnt(col) & " bang - cac bang nay nhan cung gia tri"

    ' header co o "Robot" trong 3 dong duoi tieu de
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
        gSkip = gSkip + 1
        AddLog SH_BASE & ": khong tim thay header 'Robot' duoi tieu de '" & title & "'"
        Exit Sub
    End If

    dc = EnsureDateColumn(ws, hdrRow, robotCol + 1, dToday)
    If dc = 0 Then
        gSkip = gSkip + 1
        AddLog SH_BASE & ": khong tim thay cot ngay trong bang '" & title & "'"
        Exit Sub
    End If

    ' moi bang ket thuc tai dong SUM cua no
    For r = hdrRow + 1 To endRow
        lab = Norm(ws.Cells(r, robotCol).Value)
        If IsRobot(lab) Then
            nRobot = nRobot + 1
            ws.Cells(r, dc).ClearContents
            If GetMax(dictBC, col, lab, v) Then
                ws.Cells(r, dc).Value = v
                gCells = gCells + 1
            Else
                gNoData = gNoData + 1
                AddLog SH_BASE & " [" & title & "] " & lab & ": khong co du lieu (Type " & gTypeText & ")"
            End If
            If addr <> "" Then addr = addr & ","
            addr = addr & ws.Cells(r, dc).Address(False, False)
        ElseIf lab = "SUM" Then
            sumRow = r
            Exit For
        End If
    Next r

    If nRobot = 0 Then
        gSkip = gSkip + 1
        AddLog SH_BASE & ": bang '" & title & "' khong co dong robot"
        Exit Sub
    End If
    If sumRow > 0 Then WriteSumOf ws, ws.Cells(sumRow, dc), addr
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
    Dim r As Long
    Dim c As Long
    Dim p As Long
    Dim n As Long
    Dim i As Long
    Dim hr As Long
    Dim chemCol As Long
    Dim pass As Long
    Dim endR As Long
    Dim t As String
    Dim hk As String

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
    arr = ur.Value
    If Not IsArray(arr) Then
        AddLog SH_PC & ": sheet rong"
        Exit Sub
    End If

    ' o ma robot (R11, R12...) hoac "ROBOT" -> header Chemical name trong 8 dong ke tiep
    For r = 1 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            t = Norm(arr(r, c))
            If IsRobot(t) Or t = "ROBOT" Then
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
                    If IsRobot(t) Then
                        gSkip = gSkip + 1
                        AddLog SH_PC & ": o robot " & t & " (dong " & (r + baseRow) & ") khong co header 'Chemical name' ngay duoi"
                    End If
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

    ' luot 1: bang R11, R12...   luot 2: bang "ROBOT" (tong)
    For pass = 1 To 2
        For i = 1 To n
            If (pass = 1 And CStr(tCode(i)) <> "ROBOT") Or (pass = 2 And CStr(tCode(i)) = "ROBOT") Then
                If i < n Then
                    endR = CLng(tR(i + 1)) - 1
                Else
                    endR = lastRow
                End If
                ProcessPrimerTable ws, cm, dictCC, dictPC, dToday, CStr(tCode(i)), CLng(tH(i)), CLng(tC(i)), endR
            End If
        Next i
    Next pass
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
    Dim isTotal As Boolean

    isTotal = (robot = "ROBOT")

    For c = 1 To chemCol
        If Norm(ws.Cells(hr, c).Value) = "AREA" Then
            areaCol = c
            Exit For
        End If
    Next c
    If areaCol = 0 And Not isTotal Then
        gSkip = gSkip + 1
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
        gSkip = gSkip + 1
        AddLog SH_PC & " [" & robot & "]: khong tim thay dong ngay - bo qua bang"
        Exit Sub
    End If

    dc = EnsureDateColumn(ws, dRow, chemCol + 1, dToday)
    If dc = 0 Then
        gSkip = gSkip + 1
        AddLog SH_PC & " [" & robot & "]: khong xac dinh duoc cot ngay - bo qua bang"
        Exit Sub
    End If

    avgCol = FindColInRows(ws, hr, dRow, Array("Trung b" & ChrW(236) & "nh", "Trung binh", "Average"))
    Set cl = GetDateCols(ws, dRow, chemCol + 1)

    curArea = ""
    For r = dRow + 1 To endRow
        If areaCol > 0 Then
            a = Norm(ws.Cells(r, areaCol).Value)
            If a <> "" Then curArea = a
        End If
        nm = Norm(ws.Cells(r, chemCol).Value)
        If nm <> "" Then
            If isTotal Then
                ' bang tong: cong cac bang robot theo tung hoa chat
                ws.Cells(r, dc).ClearContents
                If gChemAddr.Exists(nm) Then
                    WriteSumOf ws, ws.Cells(r, dc), CStr(gChemAddr(nm))
                    gCells = gCells + 1
                    gChem = gChem + 1
                Else
                    gNoData = gNoData + 1
                    AddLog SH_PC & " [bang tong ROBOT] dong " & r & ": '" & ws.Cells(r, chemCol).Value & "' khong co du lieu tu cac bang robot"
                End If
                If avgCol > 0 Then
                    addr = BuildAddr(ws, r, cl)
                    If addr <> "" Then ws.Cells(r, avgCol).Formula = "=IF(COUNT(" & addr & ")=0,"""",AVERAGE(" & addr & "))"
                End If
            Else
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
                        If GetMax(src, CStr(colors(k)), robot, v) Then
                            If Not found Then
                                best = v
                                found = True
                            ElseIf v > best Then
                                best = v
                            End If
                        End If
                    Next k

                    gChem = gChem + 1
                    AddAddr gChemAddr, nm, ws.Cells(r, dc).Address(False, False)
                    ws.Cells(r, dc).ClearContents
                    If found Then
                        ws.Cells(r, dc).Value = best
                        gCells = gCells + 1
                    Else
                        gNoData = gNoData + 1
                        AddLog SH_PC & " [" & robot & "] dong " & r & ": khong co du lieu " & srcName & " (Type " & gTypeText & ") cho '" & ws.Cells(r, chemCol).Value & "'"
                    End If

                    If avgCol > 0 Then
                        addr = BuildAddr(ws, r, cl)
                        If addr <> "" Then ws.Cells(r, avgCol).Formula = "=IF(COUNT(" & addr & ")=0,"""",AVERAGE(" & addr & "))"
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
    ws.Range("A2").Value = "Ngay ghi"
    ws.Range("B2").Value = Format(dToday, "dd/mm/yyyy")
    ws.Range("A3").Value = "File BC"
    ws.Range("B3").Value = pBC
    ws.Range("A4").Value = "File CC"
    ws.Range("B4").Value = pCC
    ws.Range("A5").Value = "File PC"
    ws.Range("B5").Value = pPC
    ws.Range("A6").Value = "Type da loc"
    ws.Range("B6").Value = gTypeText
    ws.Range("A7").Value = "Bang TC_Base da cap nhat"
    ws.Range("B7").Value = gBlocks
    ws.Range("A8").Value = "Dong hoa chat da xu ly"
    ws.Range("B8").Value = gChem
    ws.Range("A9").Value = "O da ghi du lieu"
    ws.Range("B9").Value = gCells
    ws.Range("A10").Value = "Dong/o khong co du lieu phu hop"
    ws.Range("B10").Value = gNoData
    ws.Range("A11").Value = "Thieu anh xa / cau hinh"
    ws.Range("B11").Value = gNoMap
    ws.Range("A12").Value = "Bo qua do thieu robot / header"
    ws.Range("B12").Value = gSkip
    ws.Range("A14").Value = "Canh bao / chi tiet"
    ws.Range("A14").Font.Bold = True

    r = 15
    If Not gLog Is Nothing Then
        For i = 1 To gLog.Count
            ws.Cells(r, 1).Value = gLog(i)
            r = r + 1
        Next i
    End If
    ws.Columns("A:B").AutoFit
End Sub
