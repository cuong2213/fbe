Option Explicit
'==============================================================================
' Tong hop du lieu BC / CC / PC -> TC_Base va TC_Primer&Clear (ThisWorkbook)
' Quy tac chinh:
'  - Chi lay du lieu cua ngay chay macro (Date)
'  - MAX rieng tung cot Rxx Consume (khong chon theo dong, khong cong tong)
'  - Hoa chat: chi dien cot robot khai bao trong Cfg_Map (cot Robot). Dong khong co R -> BO QUA
'  - TC_Base: robot R11..R24 lay dung cot Rxx Consume tuong ung, loc theo Color (+Type neu co)
'==============================================================================

Private Const SH_BASE As String = "TC_Base"
Private Const SH_PC As String = "TC_Primer&Clear"
Private Const SH_MAP As String = "Cfg_Map"
Private Const SH_BASEMAP As String = "Cfg_Base"
Private Const SH_LOG As String = "Run_Log"

Private gLog As Collection
Private gCells As Long, gBlocks As Long, gChem As Long
Private gNoData As Long, gNoMap As Long, gSkipNoR As Long

'------------------------------------------------------------------------------
Public Sub Main_UpdateData()
    Dim oCalc As XlCalculation, oScr As Boolean, oEvt As Boolean
    Dim pBC As String, pCC As String, pPC As String
    Dim sBC As Object, sCC As Object, sPC As Object
    Dim dToday As Date

    oCalc = Application.Calculation
    oScr = Application.ScreenUpdating
    oEvt = Application.EnableEvents
    On Error GoTo EH

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set gLog = New Collection
    gCells = 0: gBlocks = 0: gChem = 0: gNoData = 0: gNoMap = 0: gSkipNoR = 0

    If Not EnsureConfig() Then GoTo CleanUp
    If Not SelectSourceFiles(pBC, pCC, pPC) Then GoTo CleanUp

    dToday = Date
    Set sBC = LoadSource(pBC, dToday, "BC")
    Set sCC = LoadSource(pCC, dToday, "CC")
    Set sPC = LoadSource(pPC, dToday, "PC")

    If Not sBC Is Nothing Then UpdateTC_Base ThisWorkbook, dToday, sBC
    UpdateTC_PrimerClear ThisWorkbook, dToday, sCC, sPC

    WriteRunLog pBC, pCC, pPC, dToday
    Application.Calculation = oCalc
    Application.Calculate
    MsgBox "Hoan tat." & vbCrLf & _
           "Ngay xu ly: " & Format(dToday, "dd/mm/yyyy") & vbCrLf & _
           "Khoi TC_Base da cap nhat: " & gBlocks & vbCrLf & _
           "Dong hoa chat da cap nhat: " & gChem & vbCrLf & _
           "O da ghi gia tri: " & gCells & vbCrLf & _
           "Dong khong co du lieu phu hop: " & gNoData & vbCrLf & _
           "Dong thieu anh xa / trung Color: " & gNoMap & vbCrLf & _
           "Dong bo qua (khong co R): " & gSkipNoR & vbCrLf & _
           "Chi tiet: sheet " & SH_LOG, vbInformation, "Update BC/CC/PC"
CleanUp:
    Application.Calculation = oCalc
    Application.EnableEvents = oEvt
    Application.ScreenUpdating = oScr
    Exit Sub
EH:
    Dim em As String
    em = "Loi " & Err.Number & ": " & Err.Description
    Application.Calculation = oCalc
    Application.EnableEvents = oEvt
    Application.ScreenUpdating = oScr
    MsgBox em, vbCritical, "Update BC/CC/PC"
End Sub

'------------------------------------------------------------------------------
Private Function RobotList() As Variant
    RobotList = Array("R11", "R21", "R12", "R22", "R13", "R23", "R14", "R24")
End Function

Private Function IsRobot(ByVal s As String) As Boolean
    Dim i As Long, a As Variant
    a = RobotList()
    For i = 0 To UBound(a)
        If a(i) = s Then IsRobot = True: Exit Function
    Next
End Function

Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
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
    gLog.Add s
End Sub

'------------------------------------------------------------------------------
' Tao sheet cau hinh neu chua co
Private Function EnsureConfig() As Boolean
    Dim ws As Worksheet, created As Boolean
    If Not SheetExists(SH_MAP) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_MAP
        ws.Range("A1:E1").Value = Array("Chemical name", "Source file", "Type", "Color", "Robot")
        ws.Range("A1:E1").Font.Bold = True
        ws.Range("G1").Value = "Huong dan: moi dong = 1 to hop (Hoa chat, File CC/PC, Type, Color, Robot R11..R24)."
        ws.Range("G2").Value = "Hoa chat nhieu Color -> khai bao nhieu dong. Type de trong = khong loc Type."
        ws.Range("G3").Value = "Robot de trong = BO QUA dong do. Chemical name phai trung ten trong TC_Primer&Clear."
        created = True
    End If
    If Not SheetExists(SH_BASEMAP) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_BASEMAP
        ws.Range("A1:C1").Value = Array("Block title (day du)", "Type", "Color")
        ws.Range("A1:C1").Font.Bold = True
        ws.Range("E1").Value = "Tuy chon. Chi can khi 1 Color xuat hien o nhieu khoi TC_Base. Title phai giong tieu de khoi."
    End If
    If created Then
        MsgBox "Da tao sheet " & SH_MAP & " (va " & SH_BASEMAP & ")." & vbCrLf & _
               "Hay dien bang anh xa roi chay lai macro.", vbInformation
        EnsureConfig = False
    Else
        EnsureConfig = True
    End If
End Function

Private Function SheetExists(ByVal nm As String) As Boolean
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name = nm Then SheetExists = True: Exit Function
    Next
End Function

'------------------------------------------------------------------------------
Private Function SelectSourceFiles(ByRef pBC As String, ByRef pCC As String, ByRef pPC As String) As Boolean
    pBC = PickFile("Chon file BC (nguon TC_Base)")
    If pBC = "" Then Exit Function
    pCC = PickFile("Chon file CC (nguon TC_Primer&Clear)")
    If pCC = "" Then Exit Function
    pPC = PickFile("Chon file PC (nguon TC_Primer&Clear)")
    If pPC = "" Then Exit Function
    SelectSourceFiles = True
End Function

Private Function PickFile(ByVal ttl As String) As String
    Dim v As Variant
    v = Application.GetOpenFilename("Excel / CSV (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", , ttl)
    If VarType(v) = vbBoolean Then Exit Function
    PickFile = CStr(v)
End Function

'------------------------------------------------------------------------------
Private Function FindHeaderRow(arr As Variant) As Long
    Dim r As Long, c As Long, n As Long, t As String, mx As Long
    mx = UBound(arr, 1): If mx > 60 Then mx = 60
    For r = 1 To mx
        n = 0
        For c = 1 To UBound(arr, 2)
            t = Norm(arr(r, c))
            If t = "ENTER TIME" Or t = "TYPE" Or t = "COLOR" Or t = "CONSUMPTION" Then n = n + 1
        Next
        If n >= 4 Then FindHeaderRow = r: Exit Function
    Next
End Function

Private Function BuildColumnMap(arr As Variant, ByVal hr As Long) As Object
    Dim d As Object, c As Long, t As String
    Set d = CreateObject("Scripting.Dictionary")
    For c = 1 To UBound(arr, 2)
        t = Norm(arr(hr, c))
        If t <> "" Then If Not d.Exists(t) Then d.Add t, c
    Next
    Set BuildColumnMap = d
End Function

Private Function ToDay(ByVal v As Variant, ByRef d As Long) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        d = Int(CDbl(v)): ToDay = True
    ElseIf VarType(v) = vbString Then
        If IsDate(v) Then d = Int(CDbl(CDate(v))): ToDay = True
    ElseIf IsNumeric(v) Then
        d = Int(CDbl(v)): ToDay = True
    End If
End Function

Private Function ToNum(ByVal v As Variant, ByRef x As Double) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(CStr(v)) = "" Then Exit Function
    End If
    If IsNumeric(v) Then x = CDbl(v): ToNum = True
End Function

' Doc 1 file nguon vao Dictionary: key "TYPE|COLOR|Rxx" va "*|COLOR|Rxx" -> MAX
Private Function LoadSource(ByVal path As String, ByVal dToday As Date, ByVal tag As String) As Object
    Dim wb As Workbook, ws As Worksheet, arr As Variant
    Dim hr As Long, cm As Object, i As Long, k As Long
    Dim cT As Long, cTy As Long, cCo As Long, rc(0 To 7) As Long
    Dim dd As Long, typ As String, col As String, v As Double
    Dim rb As Variant, mx As Object, key As String, miss As String

    On Error GoTo EH
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wb.Worksheets(1)
    arr = ws.UsedRange.Value
    wb.Close SaveChanges:=False
    Set wb = Nothing
    If Not IsArray(arr) Then AddLog tag & ": file rong": Exit Function

    hr = FindHeaderRow(arr)
    If hr = 0 Then AddLog tag & ": khong tim thay dong header (Enter Time/Type/Color/Consumption)": Exit Function
    Set cm = BuildColumnMap(arr, hr)
    If Not cm.Exists("ENTER TIME") Then miss = miss & "Enter Time; "
    If Not cm.Exists("TYPE") Then miss = miss & "Type; "
    If Not cm.Exists("COLOR") Then miss = miss & "Color; "
    If Not cm.Exists("CONSUMPTION") Then miss = miss & "Consumption; "
    If miss <> "" Then AddLog tag & ": thieu header bat buoc: " & miss: Exit Function
    cT = cm("ENTER TIME"): cTy = cm("TYPE"): cCo = cm("COLOR")

    rb = RobotList()
    For k = 0 To 7
        If cm.Exists(rb(k) & " CONSUME") Then
            rc(k) = cm(rb(k) & " CONSUME")
        Else
            rc(k) = 0
            AddLog tag & ": thieu header " & rb(k) & " Consume"
        End If
    Next

    Set mx = CreateObject("Scripting.Dictionary")
    For i = hr + 1 To UBound(arr, 1)
        If ToDay(arr(i, cT), dd) Then
            If dd = CLng(Int(CDbl(dToday))) Then
                typ = Norm(arr(i, cTy)): col = Norm(arr(i, cCo))
                For k = 0 To 7
                    If rc(k) > 0 Then
                        If ToNum(arr(i, rc(k)), v) Then
                            key = typ & "|" & col & "|" & rb(k)
                            If Not mx.Exists(key) Then mx(key) = v Else If v > mx(key) Then mx(key) = v
                            key = "*|" & col & "|" & rb(k)
                            If Not mx.Exists(key) Then mx(key) = v Else If v > mx(key) Then mx(key) = v
                        End If
                    End If
                Next
            End If
        End If
    Next
    AddLog tag & ": da doc " & path & " (" & mx.Count & " nhom khoa)"
    Set LoadSource = mx
    Exit Function
EH:
    AddLog tag & ": loi doc file - " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Private Function GetMax(mx As Object, ByVal typ As String, ByVal col As String, _
                        ByVal robot As String, ByRef res As Double) As Boolean
    Dim key As String
    If typ = "" Then key = "*|" & col & "|" & robot Else key = typ & "|" & col & "|" & robot
    If mx.Exists(key) Then res = mx(key): GetMax = True
End Function

'------------------------------------------------------------------------------
' Cot ngay
Private Function IsDateCell(c As Range) As Boolean
    Dim v As Variant, f As String
    v = c.Value
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then IsDateCell = True: Exit Function
    If VarType(v) = vbDouble Then
        f = LCase$(c.NumberFormat)
        If v > 30000 And (InStr(f, "d") > 0 Or InStr(f, "m") > 0 Or InStr(f, "y") > 0) Then IsDateCell = True
    End If
End Function

Private Function FindDateHeaderRow(ws As Worksheet) As Long
    Dim r As Long, c As Long, n As Long, lc As Long
    For r = 1 To 40
        lc = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        n = 0
        For c = 1 To lc
            If IsDateCell(ws.Cells(r, c)) Then n = n + 1
        Next
        If n >= 2 Then FindDateHeaderRow = r: Exit Function
    Next
End Function

Private Function GetDateCols(ws As Worksheet, ByVal hr As Long) As Collection
    Dim cl As New Collection, c As Long, lc As Long
    lc = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lc
        If IsDateCell(ws.Cells(hr, c)) Then cl.Add c
    Next
    Set GetDateCols = cl
End Function

' Tra ve so cot ngay (co san hoac them moi sau cot ngay cuoi, truoc Xa/Trung binh)
Private Function EnsureDateColumn(ws As Worksheet, ByVal hr As Long, ByVal d As Date) As Long
    Dim cl As Collection, i As Long, lastC As Long, newC As Long, nf As String
    Set cl = GetDateCols(ws, hr)
    If cl.Count = 0 Then Exit Function
    For i = 1 To cl.Count
        If CLng(Int(CDbl(ws.Cells(hr, cl(i)).Value))) = CLng(Int(CDbl(d))) Then
            EnsureDateColumn = cl(i): Exit Function
        End If
    Next
    lastC = cl(cl.Count)
    nf = ws.Cells(hr, lastC).NumberFormat
    newC = lastC + 1
    ws.Columns(newC).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(hr, newC).Value = CDbl(Int(CDbl(d)))
    ws.Cells(hr, newC).NumberFormat = nf
    EnsureDateColumn = newC
End Function

'------------------------------------------------------------------------------
Private Function FindTextRow(ws As Worksheet, ByVal txt As String, ByRef colOut As Long) As Long
    Dim r As Long, c As Long, lc As Long
    For r = 1 To 40
        lc = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        For c = 1 To lc
            If Norm(ws.Cells(r, c).Value) = txt Then FindTextRow = r: colOut = c: Exit Function
        Next
    Next
End Function

Private Function FindColByNames(ws As Worksheet, ByVal hr As Long, names As Variant) As Long
    Dim c As Long, lc As Long, i As Long, t As String
    lc = ws.Cells(hr, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lc
        t = Norm(ws.Cells(hr, c).Value)
        For i = LBound(names) To UBound(names)
            If t = Norm(names(i)) Then FindColByNames = c: Exit Function
        Next
    Next
End Function

'------------------------------------------------------------------------------
Private Sub UpdateTC_Base(wb As Workbook, ByVal dToday As Date, mx As Object)
    Dim ws As Worksheet, hrDate As Long, dc As Long, lastRow As Long
    Dim labelCol As Long, r As Long, i As Long, n As Long
    Dim tRows As New Collection, tText As New Collection
    Dim cnt As Object, s As String, code As String
    Dim ov As Object, wsB As Worksheet, ba As Variant
    Dim rStart As Long, rEnd As Long, typ As String, col As String, key As String
    Dim lab As String, sumRow As Long, addr As String, v As Double, any1 As Boolean

    If Not SheetExists(SH_BASE) Then AddLog "Khong co sheet " & SH_BASE: Exit Sub
    Set ws = wb.Worksheets(SH_BASE)
    hrDate = FindDateHeaderRow(ws)
    If hrDate = 0 Then AddLog SH_BASE & ": khong tim thay dong ngay": Exit Sub
    dc = EnsureDateColumn(ws, hrDate, dToday)
    If dc = 0 Then AddLog SH_BASE & ": khong tao duoc cot ngay": Exit Sub

    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    labelCol = FindRobotLabelCol(ws, hrDate, lastRow)
    If labelCol = 0 Then AddLog SH_BASE & ": khong tim thay cac dong robot R11..R24": Exit Sub

    ' override (tuy chon)
    Set ov = CreateObject("Scripting.Dictionary")
    Set wsB = wb.Worksheets(SH_BASEMAP)
    If wsB.UsedRange.Rows.Count > 1 Then
        ba = wsB.Range("A1").Resize(wsB.UsedRange.Rows.Count, 3).Value
        For i = 2 To UBound(ba, 1)
            If Norm(ba(i, 1)) <> "" Then ov(Norm(ba(i, 1))) = Norm(ba(i, 2)) & "|" & Norm(ba(i, 3))
        Next
    End If

    ' tim tieu de khoi
    Set cnt = CreateObject("Scripting.Dictionary")
    For r = hrDate + 1 To lastRow
        s = TitleAt(ws, r, labelCol)
        If s <> "" Then
            tRows.Add r: tText.Add s
            code = Norm(Left$(s, InStr(s, "_") - 1))
            cnt(code) = cnt(code) + 1
        End If
    Next
    n = tRows.Count
    If n = 0 Then AddLog SH_BASE & ": khong tim thay tieu de khoi": Exit Sub

    For i = 1 To n
        rStart = tRows(i) + 1
        If i < n Then rEnd = tRows(i + 1) - 1 Else rEnd = lastRow
        s = tText(i)
        code = Norm(Left$(s, InStr(s, "_") - 1))
        typ = "": col = code
        If ov.Exists(Norm(s)) Then
            key = ov(Norm(s))
            typ = Left$(key, InStr(key, "|") - 1)
            If Mid$(key, InStr(key, "|") + 1) <> "" Then col = Mid$(key, InStr(key, "|") + 1)
        ElseIf cnt(code) > 1 Then
            gNoMap = gNoMap + 1
            AddLog SH_BASE & ": Color " & code & " xuat hien o nhieu khoi - can khai bao trong " & SH_BASEMAP & " (bo qua khoi '" & s & "')"
            GoTo NextBlock
        End If

        sumRow = 0: addr = "": any1 = False
        For r = rStart To rEnd
            lab = Norm(ws.Cells(r, labelCol).Value)
            If IsRobot(lab) Then
                ws.Cells(r, dc).ClearContents
                If GetMax(mx, typ, col, lab, v) Then
                    ws.Cells(r, dc).Value = v
                    gCells = gCells + 1
                    any1 = True
                Else
                    gNoData = gNoData + 1
                End If
                If addr <> "" Then addr = addr & ","
                addr = addr & ws.Cells(r, dc).Address(False, False)
            ElseIf lab = "SUM" Then
                sumRow = r
            End If
        Next
        If sumRow > 0 And addr <> "" Then
            ws.Cells(sumRow, dc).Formula = "=IF(COUNT(" & addr & ")=0,"""",SUM(" & addr & "))"
        End If
        If addr <> "" Then gBlocks = gBlocks + 1
NextBlock:
    Next
End Sub

Private Function FindRobotLabelCol(ws As Worksheet, ByVal r0 As Long, ByVal r1 As Long) As Long
    Dim c As Long, r As Long, n As Long, best As Long
    For c = 1 To 10
        n = 0
        For r = r0 + 1 To r1
            If IsRobot(Norm(ws.Cells(r, c).Value)) Then n = n + 1
        Next
        If n > best Then best = n: FindRobotLabelCol = c
    Next
End Function

' Tra ve text tieu de neu dong r la dong tieu de khoi, nguoc lai ""
Private Function TitleAt(ws As Worksheet, ByVal r As Long, ByVal labelCol As Long) As String
    Dim c As Long, s As String, p As Long, code As String
    For c = 1 To labelCol
        If VarType(ws.Cells(r, c).Value) = vbString Then
            s = Trim$(CStr(ws.Cells(r, c).Value))
            p = InStr(s, "_")
            If p > 1 Then
                code = Left$(s, p - 1)
                If Len(code) >= 2 And Len(code) <= 8 And InStr(code, " ") = 0 And code Like "*#*" Then
                    TitleAt = s: Exit Function
                End If
            End If
        End If
    Next
End Function

'------------------------------------------------------------------------------
Private Sub UpdateTC_PrimerClear(wb As Workbook, ByVal dToday As Date, sCC As Object, sPC As Object)
    Dim ws As Worksheet, hrChem As Long, chemCol As Long, hrDate As Long, dc As Long
    Dim avgCol As Long, lastRow As Long, r As Long, m As Long
    Dim wm As Worksheet, ma As Variant
    Dim nm As String, hasMap As Boolean, hasRobot As Boolean, found As Boolean
    Dim best As Double, v As Double, rb As String, srcN As String, src As Object
    Dim cl As Collection, addr As String, i As Long

    If Not SheetExists(SH_PC) Then AddLog "Khong co sheet " & SH_PC: Exit Sub
    Set ws = wb.Worksheets(SH_PC)
    hrChem = FindTextRow(ws, "CHEMICAL NAME", chemCol)
    If hrChem = 0 Then AddLog SH_PC & ": khong tim thay header 'Chemical name'": Exit Sub
    hrDate = FindDateHeaderRow(ws)
    If hrDate = 0 Then AddLog SH_PC & ": khong tim thay dong ngay": Exit Sub
    dc = EnsureDateColumn(ws, hrDate, dToday)
    If dc = 0 Then AddLog SH_PC & ": khong tao duoc cot ngay": Exit Sub

    avgCol = FindColByNames(ws, hrDate, Array("Trung b" & ChrW(236) & "nh", "Trung binh", "Average"))
    If avgCol = 0 Then avgCol = FindColByNames(ws, hrChem, Array("Trung b" & ChrW(236) & "nh", "Trung binh", "Average"))

    Set wm = wb.Worksheets(SH_MAP)
    If wm.UsedRange.Rows.Count < 2 Then
        AddLog SH_MAP & ": chua co du lieu anh xa - khong dien hoa chat"
        Exit Sub
    End If
    ma = wm.Range("A1").Resize(wm.UsedRange.Rows.Count, 5).Value

    Set cl = GetDateCols(ws, hrDate)

    lastRow = ws.Cells(ws.Rows.Count, chemCol).End(xlUp).Row
    For r = hrChem + 1 To lastRow
        nm = Norm(ws.Cells(r, chemCol).Value)
        If nm <> "" Then
            hasMap = False: hasRobot = False: found = False: best = 0
            For m = 2 To UBound(ma, 1)
                If Norm(ma(m, 1)) = nm Then
                    hasMap = True
                    rb = Norm(ma(m, 5))
                    If IsRobot(rb) And Norm(ma(m, 4)) <> "" Then
                        srcN = Norm(ma(m, 2))
                        Set src = Nothing
                        If srcN = "CC" Then Set src = sCC
                        If srcN = "PC" Then Set src = sPC
                        If src Is Nothing Then
                            AddLog SH_PC & " dong " & r & ": file nguon '" & srcN & "' khong hop le / khong doc duoc"
                        Else
                            hasRobot = True
                            If GetMax(src, Norm(ma(m, 3)), Norm(ma(m, 4)), rb, v) Then
                                If Not found Or v > best Then best = v: found = True
                            End If
                        End If
                    End If
                End If
            Next
            If Not hasMap Then
                gNoMap = gNoMap + 1
                AddLog SH_PC & " dong " & r & ": '" & ws.Cells(r, chemCol).Value & "' khong co trong " & SH_MAP
            ElseIf Not hasRobot Then
                gSkipNoR = gSkipNoR + 1
            Else
                gChem = gChem + 1
                ws.Cells(r, dc).ClearContents
                If found Then
                    ws.Cells(r, dc).Value = best
                    gCells = gCells + 1
                Else
                    gNoData = gNoData + 1
                    AddLog SH_PC & " dong " & r & ": khong co du lieu phu hop trong ngay"
                End If
            End If
            ' tinh lai Trung binh cho dong hoa chat co trong bang anh xa
            If hasMap And avgCol > 0 Then
                addr = ""
                For i = 1 To cl.Count
                    If addr <> "" Then addr = addr & ","
                    addr = addr & ws.Cells(r, cl(i)).Address(False, False)
                Next
                ws.Cells(r, avgCol).Formula = "=IF(COUNT(" & addr & ")=0,"""",AVERAGE(" & addr & "))"
            End If
        End If
    Next
End Sub

'------------------------------------------------------------------------------
Private Sub WriteRunLog(ByVal pBC As String, ByVal pCC As String, ByVal pPC As String, ByVal dToday As Date)
    Dim ws As Worksheet, r As Long, i As Long
    If SheetExists(SH_LOG) Then
        Set ws = ThisWorkbook.Worksheets(SH_LOG): ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
    End If
    ws.Range("A1").Value = "Thoi diem chay": ws.Range("B1").Value = Format(Now, "dd/mm/yyyy hh:nn:ss")
    ws.Range("A2").Value = "Ngay xu ly": ws.Range("B2").Value = Format(dToday, "dd/mm/yyyy")
    ws.Range("A3").Value = "File BC": ws.Range("B3").Value = pBC
    ws.Range("A4").Value = "File CC": ws.Range("B4").Value = pCC
    ws.Range("A5").Value = "File PC": ws.Range("B5").Value = pPC
    ws.Range("A6").Value = "Khoi TC_Base cap nhat": ws.Range("B6").Value = gBlocks
    ws.Range("A7").Value = "Dong hoa chat cap nhat": ws.Range("B7").Value = gChem
    ws.Range("A8").Value = "O da ghi": ws.Range("B8").Value = gCells
    ws.Range("A9").Value = "Dong khong co du lieu": ws.Range("B9").Value = gNoData
    ws.Range("A10").Value = "Dong thieu anh xa / trung Color": ws.Range("B10").Value = gNoMap
    ws.Range("A11").Value = "Dong bo qua (khong co R)": ws.Range("B11").Value = gSkipNoR
    ws.Range("A13").Value = "Chi tiet / canh bao": ws.Range("A13").Font.Bold = True
    r = 14
    For i = 1 To gLog.Count
        ws.Cells(r, 1).Value = gLog(i): r = r + 1
    Next
    ws.Columns("A:B").AutoFit
End Sub

'==============================================================================
' VI DU KIEM THU MAX DOC LAP (cung ngay hom nay, Type=T1, Color=CE18):
'   Consumption  R11  R21
'   850          150  160
'   888          157  155
'   870          155  169
' Ket qua: R11 = 157 (dong 2), R21 = 169 (dong 3) -> MAX lay rieng tung cot.
' Neu them dong R11 = 0 va cac dong khac trong: 0 van la gia tri hop le.
'==============================================================================
