Attribute VB_Name = "modCheckBOM"
Option Explicit

' Kiem tra file MOI (Super BOM / Single BOM) da khop file GOC chua.
' Chay macro: CheckBOM
' Toi uu cho ~100000 dong: doc ca 2 file vao mang, tra cuu bang Dictionary + linked list,
' cache ket qua NormVar, ghi ket qua 1 lan, tat ScreenUpdating/Calculation/Events.
' Luu y: de ASCII (khong dau) vi VBE khong hien thi Unicode.

Private gRe As Object
Private gCache As Object
Private gWs As Worksheet
Private gBuf() As Variant
Private gCap As Long
Private gCnt As Long

Public Sub CheckBOM()
    Dim runDate As Date, pNew As String, pOrg As String
    Dim wbN As Workbook, wbO As Workbook, wsN As Worksheet
    Dim rSup As Long, rSin As Long, eN As Long, eD As String
    Dim calcMode As XlCalculation, t0 As Double

    runDate = ParseDMY(InputBox("Ngay chay (dd/mm/yyyy):", "Kiem tra BOM", Format(Date, "dd/mm/yyyy")))
    If runDate = 0 Then Exit Sub

    pNew = PickFile("Chon FILE MOI")
    If Len(pNew) = 0 Then Exit Sub

    t0 = Timer
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Set gCache = Nothing
    On Error GoTo Fin
    Set wbN = Workbooks.Open(pNew, UpdateLinks:=0, ReadOnly:=True)
    Set wsN = wbN.Worksheets(1)

    rSup = FindMarker(wsN, "super bom")
    rSin = FindMarker(wsN, "single bom")
    If rSup = 0 And rSin = 0 Then
        MsgBox "Khong thay 'Super BOM' hoac 'Single BOM' (cot A:C) trong file moi.", vbExclamation
        GoTo Fin
    End If

    InitResult

    If rSup > 0 Then
        pOrg = PickFile("Chon FILE GOC - SUPER BOM")
        If Len(pOrg) > 0 Then
            Set wbO = Workbooks.Open(pOrg, UpdateLinks:=0, ReadOnly:=True)
            CheckSuper wsN, rSup, wbO.Worksheets(1), runDate
            wbO.Close False: Set wbO = Nothing
        End If
    End If

    If rSin > 0 Then
        pOrg = PickFile("Chon FILE GOC - SINGLE BOM")
        If Len(pOrg) > 0 Then
            Set wbO = Workbooks.Open(pOrg, UpdateLinks:=0, ReadOnly:=True)
            CheckSingle wsN, rSin, wbO.Worksheets(1), runDate
            wbO.Close False: Set wbO = Nothing
        End If
    End If

    Application.StatusBar = False
    ShowResult Timer - t0

Fin:
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    If Not wbO Is Nothing Then wbO.Close False
    If Not wbN Is Nothing Then wbN.Close False
    Application.StatusBar = False
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If eN <> 0 Then MsgBox "Loi: " & eD, vbCritical
End Sub

' ===================== SUPER BOM =====================
Private Sub CheckSuper(wsN As Worksheet, mRow As Long, wsO As Worksheet, runDate As Date)
    Dim hN As Long, cP As Long, cU As Long, cQ As Long, cO As Long, cV As Long, lastColN As Long, lastRowN As Long
    Dim hO As Long, kM As Long, kF1 As Long, kT1 As Long, kF2 As Long, kT2 As Long
    Dim kQ As Long, kU As Long, kB As Long, kS As Long, lastO As Long, lastColO As Long
    Dim vO As Variant, vN As Variant, d As Object, nxt() As Long, idx As Long
    Dim i As Long, r As Long, key As String, part As String
    Dim det As String, cols As String, n As Long, best As String, bestCols As String, bestN As Long

    hN = FindHdr(wsN, mRow, "Part Number", 6)
    If hN = 0 Then AddRes "Super BOM", mRow, "", "Header", "Khong thay header 'Part Number' o file moi": Exit Sub
    cP = ColOf(wsN, hN, "Part Number", 1)
    cU = ColOf(wsN, hN, "UOM", 1)
    cQ = ColOf(wsN, hN, "Q'ty", 1)
    cO = ColOf(wsN, hN, "Operation ID", 1)
    cV = NearRight(wsN, hN, "Variant", cO)
    If cV = 0 Then cV = ColOf(wsN, hN, "Variant", 1)
    If cP * cU * cQ * cO * cV = 0 Then AddRes "Super BOM", hN, "", "Header", "Thieu cot (Part Number/UOM/Q'ty/Operation ID/Variant) o file moi": Exit Sub
    lastColN = wsN.Cells(hN, wsN.Columns.Count).End(xlToLeft).Column

    hO = FindHdr(wsO, 1, "BOP ID", 30)
    If hO = 0 Then AddRes "Super BOM", 0, "", "Header", "Khong thay header 'BOP ID' o file goc": Exit Sub
    kM = ColOf(wsO, hO, "Material", -1)          ' cot Material nam ben phai nhat = part number
    kT1 = ColOf(wsO, hO, "Valid-To Date", 1)
    kT2 = ColOf(wsO, hO, "Valid To", 1)
    kF1 = NearLeft(wsO, hO, "Valid From", kT1)   ' Valid From gan nhat ben trai moi cot Valid To
    kF2 = NearLeft(wsO, hO, "Valid From", kT2)
    kQ = ColOf(wsO, hO, "Quantity", 1)
    kU = ColOf(wsO, hO, "Unit of Measure", 1)
    kB = ColOf(wsO, hO, "BOP ID", 1)
    kS = ColOf(wsO, hO, "Selection Rule", 1)
    If kM * kQ * kU * kB * kS = 0 Then AddRes "Super BOM", 0, "", "Header", "Thieu cot (Material/Quantity/Unit of Measure/BOP ID/Selection Rule) o file goc": Exit Sub
    lastColO = wsO.Cells(hO, wsO.Columns.Count).End(xlToLeft).Column

    lastO = wsO.Cells(wsO.Rows.Count, kM).End(xlUp).Row + 1
    If lastO <= hO + 1 Then AddRes "Super BOM", 0, "", "File goc", "File goc khong co du lieu": Exit Sub
    Application.StatusBar = "Dang doc file goc Super BOM..."
    vO = wsO.Range(wsO.Cells(hO + 1, 1), wsO.Cells(lastO, lastColO)).Value

    ' index file goc: key -> dong cuoi, nxt(i) -> dong truoc cung key
    Set d = CreateObject("Scripting.Dictionary")
    ReDim nxt(1 To UBound(vO, 1))
    For i = 1 To UBound(vO, 1)
        If Len(Trim(CStr(vO(i, kM)))) > 0 Then
            If RowValid(vO, i, kF1, kT1, runDate) And RowValid(vO, i, kF2, kT2, runDate) Then
                key = UCase(Trim(CStr(vO(i, kM))))
                If d.Exists(key) Then nxt(i) = d(key) Else nxt(i) = 0
                d(key) = i
            End If
        End If
    Next i

    lastRowN = LastUsedRow(wsN) + 1
    If lastRowN < hN + 2 Then lastRowN = hN + 2
    vN = wsN.Range(wsN.Cells(hN + 1, 1), wsN.Cells(lastRowN, lastColN)).Value

    For i = 1 To UBound(vN, 1)
        If RowBlank(vN, i, lastColN) Then Exit For          ' dong trong dau tien -> dung
        r = hN + i
        If i Mod 5000 = 0 Then Application.StatusBar = "Super BOM: " & i & " dong..."
        part = Trim(CStr(vN(i, cP)))
        key = UCase(part)
        If Not d.Exists(key) Then
            AddRes "Super BOM", r, part, "Part Number", "Khong co trong file goc (hoac chua/het hieu luc)"
        Else
            bestN = 999
            idx = d(key)
            Do While idx > 0
                det = DiffSuper(vN(i, cU), vN(i, cQ), vN(i, cO), vN(i, cV), _
                                vO(idx, kU), vO(idx, kQ), vO(idx, kB), vO(idx, kS), cols, n)
                If n < bestN Then bestN = n: best = det: bestCols = cols
                If n = 0 Then Exit Do
                idx = nxt(idx)
            Loop
            If bestN > 0 Then AddRes "Super BOM", r, part, bestCols, best
        End If
    Next i
End Sub

Private Function DiffSuper(nU As Variant, nQ As Variant, nO As Variant, nV As Variant, _
                           oU As Variant, oQ As Variant, oB As Variant, oS As Variant, _
                           ByRef cols As String, ByRef n As Long) As String
    Dim s As String, vn As String
    cols = "": n = 0
    If UCase(Trim(CStr(nU))) <> UCase(Trim(CStr(oU))) Then AddDiff s, cols, n, "UOM", CStr(nU), CStr(oU)
    If Not SameId(nO, oB) Then AddDiff s, cols, n, "Operation ID", CStr(nO), CStr(oB)
    vn = NormVar(CStr(nV))
    If HasDup(vn) Then
        AddDiff s, cols, n, "Variant (trung gia tri)", vn, NormVar(CStr(oS))
    ElseIf vn <> NormVar(CStr(oS)) Then
        AddDiff s, cols, n, "Variant", vn, NormVar(CStr(oS))
    End If
    If Not SameNum(nQ, oQ) Then AddDiff s, cols, n, "Q'ty", CStr(nQ), CStr(oQ)
    DiffSuper = s
End Function

' ===================== SINGLE BOM =====================
Private Sub CheckSingle(wsN As Worksheet, mRow As Long, wsO As Worksheet, runDate As Date)
    Dim hN As Long, cP As Long, cC As Long, cU As Long, cQ As Long, lastColN As Long, lastRowN As Long
    Dim hO As Long, kM As Long, kC As Long, kU As Long, kQ As Long, kF As Long, kT As Long
    Dim lastColO As Long, lastO As Long, vO As Variant, vN As Variant, d As Object, nxt() As Long, idx As Long
    Dim i As Long, r As Long, key As String, part As String, child As String
    Dim det As String, cols As String, n As Long, best As String, bestCols As String, bestN As Long

    hN = FindHdr(wsN, mRow, "Old Parent Part", 6)
    If hN = 0 Then AddRes "Single BOM", mRow, "", "Header", "Khong thay header 'Old Parent Part' o file moi": Exit Sub
    cP = ColOf(wsN, hN, "Old Parent Part", 1)
    cC = ColOf(wsN, hN, "Old Child part", 1)
    cU = ColOf(wsN, hN, "Old UOM", 1)
    cQ = ColOf(wsN, hN, "Old Q'ty", 1)
    If cP * cC * cU * cQ = 0 Then AddRes "Single BOM", hN, "", "Header", "Thieu cot (Old Parent Part/Old Child part/Old UOM/Old Q'ty) o file moi": Exit Sub
    lastColN = wsN.Cells(hN, wsN.Columns.Count).End(xlToLeft).Column

    hO = FindHdr(wsO, 1, "Component", 30)
    If hO = 0 Then AddRes "Single BOM", 0, "", "Header", "Khong thay header 'Component' o file goc": Exit Sub
    kM = ColOf(wsO, hO, "Material", 1)
    kC = ColOf(wsO, hO, "Component", 1)
    kU = ColOf(wsO, hO, "UoM", 1)
    kQ = ColOf(wsO, hO, "Quantity", 1)
    kF = ColOf(wsO, hO, "Valid from", 1)
    kT = ColOf(wsO, hO, "Valid Date To", 1)
    If kM * kC * kU * kQ = 0 Then AddRes "Single BOM", 0, "", "Header", "Thieu cot (Material/Component/UoM/Quantity) o file goc": Exit Sub
    lastColO = wsO.Cells(hO, wsO.Columns.Count).End(xlToLeft).Column

    lastO = wsO.Cells(wsO.Rows.Count, kM).End(xlUp).Row + 1
    If lastO <= hO + 1 Then AddRes "Single BOM", 0, "", "File goc", "File goc khong co du lieu": Exit Sub
    Application.StatusBar = "Dang doc file goc Single BOM..."
    vO = wsO.Range(wsO.Cells(hO + 1, 1), wsO.Cells(lastO, lastColO)).Value

    Set d = CreateObject("Scripting.Dictionary")
    ReDim nxt(1 To UBound(vO, 1))
    For i = 1 To UBound(vO, 1)
        If Len(Trim(CStr(vO(i, kM)))) > 0 Then
            If RowValid(vO, i, kF, kT, runDate) Then
                key = UCase(Trim(CStr(vO(i, kM)))) & "|" & UCase(Trim(CStr(vO(i, kC))))
                If d.Exists(key) Then nxt(i) = d(key) Else nxt(i) = 0
                d(key) = i
            End If
        End If
    Next i

    lastRowN = LastUsedRow(wsN) + 1
    If lastRowN < hN + 2 Then lastRowN = hN + 2
    vN = wsN.Range(wsN.Cells(hN + 1, 1), wsN.Cells(lastRowN, lastColN)).Value

    For i = 1 To UBound(vN, 1)
        If RowBlank(vN, i, lastColN) Then Exit For
        r = hN + i
        If i Mod 5000 = 0 Then Application.StatusBar = "Single BOM: " & i & " dong..."
        part = Trim(CStr(vN(i, cP)))
        child = Trim(CStr(vN(i, cC)))
        If Len(part) > 0 And Len(child) > 0 Then   ' dong khong co Old part (them moi) -> bo qua
            key = UCase(part) & "|" & UCase(child)
            If Not d.Exists(key) Then
                AddRes "Single BOM", r, part & " / " & child, "Old Parent Part / Old Child part", "Khong co cap Material-Component nay trong file goc (hoac het hieu luc)"
            Else
                bestN = 999
                idx = d(key)
                Do While idx > 0
                    det = "": cols = "": n = 0
                    If UCase(Trim(CStr(vN(i, cU)))) <> UCase(Trim(CStr(vO(idx, kU)))) Then _
                        AddDiff det, cols, n, "Old UOM", CStr(vN(i, cU)), CStr(vO(idx, kU))
                    If Not SameNum(vN(i, cQ), vO(idx, kQ)) Then _
                        AddDiff det, cols, n, "Old Q'ty", CStr(vN(i, cQ)), CStr(vO(idx, kQ))
                    If n < bestN Then bestN = n: best = det: bestCols = cols
                    If n = 0 Then Exit Do
                    idx = nxt(idx)
                Loop
                If bestN > 0 Then AddRes "Single BOM", r, part & " / " & child, bestCols, best
            End If
        End If
    Next i
End Sub

' ===================== HELPERS =====================
Private Function LastUsedRow(ws As Worksheet) As Long
    With ws.UsedRange
        LastUsedRow = .Row + .Rows.Count - 1
    End With
End Function

Private Function RowBlank(v As Variant, i As Long, nCols As Long) As Boolean
    Dim c As Long
    For c = 1 To nCols
        If IsError(v(i, c)) Then Exit Function
        If Len(CStr(v(i, c))) > 0 Then Exit Function
    Next c
    RowBlank = True
End Function

Private Sub AddDiff(ByRef s As String, ByRef cols As String, ByRef n As Long, nm As String, vNew As String, vOld As String)
    If n > 0 Then s = s & "; ": cols = cols & ", "
    s = s & nm & ": moi=" & vNew & " | goc=" & vOld
    cols = cols & nm
    n = n + 1
End Sub

Private Function RowValid(v As Variant, i As Long, cFrom As Long, cTo As Long, runDate As Date) As Boolean
    Dim d As Variant
    RowValid = True
    If cFrom > 0 Then
        d = ToDate(v(i, cFrom))
        If Not IsEmpty(d) Then If CDate(d) > runDate Then RowValid = False: Exit Function
    End If
    If cTo > 0 Then
        d = ToDate(v(i, cTo))
        If Not IsEmpty(d) Then If CDate(d) <= runDate Then RowValid = False
    End If
End Function

Private Function ToDate(v As Variant) As Variant
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        ToDate = v
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then ToDate = CDate(CDbl(v))
    ElseIf IsDate(v) Then
        ToDate = CDate(v)
    End If
End Function

Private Function SameId(a As Variant, b As Variant) As Boolean
    Dim sa As String, sb As String
    sa = Trim(CStr(a)): sb = Trim(CStr(b))
    If sa = sb Then SameId = True: Exit Function
    If IsNumeric(sa) And IsNumeric(sb) And Len(sa) > 0 And Len(sb) > 0 Then
        SameId = (CDbl(sa) = CDbl(sb))
    Else
        SameId = (UCase(sa) = UCase(sb))
    End If
End Function

Private Function SameNum(a As Variant, b As Variant) As Boolean
    Dim sa As String, sb As String
    sa = Trim(CStr(a)): sb = Trim(CStr(b))
    If Len(sa) = 0 And Len(sb) = 0 Then SameNum = True: Exit Function
    If IsNumeric(sa) And IsNumeric(sb) And Len(sa) > 0 And Len(sb) > 0 Then
        SameNum = (Abs(CDbl(sa) - CDbl(sb)) < 0.000001)
    Else
        SameNum = (sa = sb)
    End If
End Function

' Lay cac gia tri sau dau "=", KHONG bo trung, sap xep. Neu khong co "=" thi tach theo , ; / space OR
' Co cache theo chuoi goc (Selection Rule lap lai rat nhieu) va khong dung Dictionary ben trong.
Private Function NormVar(ByVal raw As String) As String
    Dim s As String, ms As Object, i As Long, j As Long, n As Long, t As Variant, p() As String
    Dim a() As String, tmp As String

    If Len(raw) = 0 Then Exit Function
    If gCache Is Nothing Then Set gCache = CreateObject("Scripting.Dictionary")
    If gCache.Exists(raw) Then NormVar = gCache(raw): Exit Function

    s = UCase(Trim(raw))
    If Len(s) = 0 Then Exit Function
    If InStr(s, "=") > 0 Then
        If gRe Is Nothing Then
            Set gRe = CreateObject("VBScript.RegExp")
            gRe.Global = True
            gRe.Pattern = "=\s*'?([^'\s()]+)'?"
        End If
        Set ms = gRe.Execute(s)
        If ms.Count = 0 Then gCache(raw) = "": Exit Function
        ReDim a(0 To ms.Count - 1)
        For i = 0 To ms.Count - 1
            a(i) = CStr(ms(i).SubMatches(0))
        Next i
        n = ms.Count
    Else
        For Each t In Array(",", ";", "/", "|", "'", """", vbLf, vbCr, vbTab)
            s = Replace(s, CStr(t), " ")
        Next t
        p = Split(s, " ")
        ReDim a(0 To UBound(p))
        For i = 0 To UBound(p)
            If Len(p(i)) > 0 And p(i) <> "OR" Then a(n) = p(i): n = n + 1
        Next i
        If n = 0 Then gCache(raw) = "": Exit Function
        ReDim Preserve a(0 To n - 1)
    End If

    For i = 1 To n - 1                        ' insertion sort
        tmp = a(i): j = i - 1
        Do While j >= 0
            If a(j) > tmp Then
                a(j + 1) = a(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        a(j + 1) = tmp
    Next i
    NormVar = Join(a, ",")
    gCache(raw) = NormVar
End Function

Private Function Norm(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = CStr(v)
    s = Replace(s, ChrW(8217), "'")
    s = Replace(s, Chr(160), " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = LCase(Trim(s))
End Function

Private Function FindHdr(ws As Worksheet, r0 As Long, nm As String, nRows As Long) As Long
    Dim r As Long, c As Long, t As String
    t = Norm(nm)
    For r = r0 To r0 + nRows
        For c = 1 To 80
            If Norm(ws.Cells(r, c).Value) = t Then FindHdr = r: Exit Function
        Next c
    Next r
End Function

Private Function ColOf(ws As Worksheet, hRow As Long, nm As String, occ As Long) As Long
    Dim c As Long, k As Long, t As String
    t = Norm(nm)
    For c = 1 To 80
        If Norm(ws.Cells(hRow, c).Value) = t Then
            k = k + 1
            If k = occ Then ColOf = c: Exit Function
            If occ < 0 Then ColOf = c   ' occ = -1: lay cot cuoi cung trung ten
        End If
    Next c
End Function

Private Function NearLeft(ws As Worksheet, hRow As Long, nm As String, refCol As Long) As Long
    Dim c As Long, t As String
    If refCol = 0 Then Exit Function
    t = Norm(nm)
    For c = refCol - 1 To 1 Step -1
        If Norm(ws.Cells(hRow, c).Value) = t Then NearLeft = c: Exit Function
    Next c
End Function

Private Function NearRight(ws As Worksheet, hRow As Long, nm As String, refCol As Long) As Long
    Dim c As Long, t As String
    If refCol = 0 Then Exit Function
    t = Norm(nm)
    For c = refCol + 1 To 80
        If Norm(ws.Cells(hRow, c).Value) = t Then NearRight = c: Exit Function
    Next c
End Function

' s la chuoi da sap xep, ngan cach bang ","
Private Function HasDup(ByVal s As String) As Boolean
    Dim a() As String, i As Long
    If Len(s) = 0 Then Exit Function
    a = Split(s, ",")
    For i = 1 To UBound(a)
        If a(i) = a(i - 1) Then HasDup = True: Exit Function
    Next i
End Function

Private Function FindMarker(ws As Worksheet, txt As String) As Long
    Dim f As Range
    Set f = ws.Range("A:C").Find(What:=txt, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False)
    If Not f Is Nothing Then FindMarker = f.Row
End Function

Private Function PickFile(ttl As String) As String
    With Application.FileDialog(3)
        .Title = ttl
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xls;*.xlsb"
        If .Show = -1 Then PickFile = .SelectedItems(1)
    End With
End Function

Private Function ParseDMY(s As String) As Date
    Dim p As Variant
    If Len(s) = 0 Then Exit Function
    p = Split(s, "/")
    If UBound(p) <> 2 Then MsgBox "Ngay khong hop le (dd/mm/yyyy)": Exit Function
    On Error Resume Next
    ParseDMY = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
    If Err.Number <> 0 Then ParseDMY = 0: MsgBox "Ngay khong hop le (dd/mm/yyyy)"
End Function

' ===================== KET QUA (buffer trong mang, ghi 1 lan) =====================
Private Sub InitResult()
    Set gWs = Nothing
    On Error Resume Next
    Set gWs = ThisWorkbook.Worksheets("Check_Result")
    On Error GoTo 0
    If gWs Is Nothing Then
        Set gWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        gWs.Name = "Check_Result"
    End If
    gWs.Cells.Clear
    gWs.Columns(3).NumberFormat = "@"
    gWs.Range("A1:E1").Value = Array("Loai BOM", "Dong (file moi)", "Part", "Cot loi", "Chi tiet")
    gWs.Range("A1:E1").Font.Bold = True
    gCap = 1000: gCnt = 0
    ReDim gBuf(1 To gCap, 1 To 5)
End Sub

Private Sub AddRes(typ As String, rw As Long, part As String, cols As String, det As String)
    Dim tmp() As Variant, i As Long, j As Long
    If gCnt >= gCap Then
        ReDim tmp(1 To gCap * 2, 1 To 5)
        For i = 1 To gCnt
            For j = 1 To 5
                tmp(i, j) = gBuf(i, j)
            Next j
        Next i
        gBuf = tmp
        gCap = gCap * 2
    End If
    gCnt = gCnt + 1
    gBuf(gCnt, 1) = typ
    gBuf(gCnt, 2) = rw
    gBuf(gCnt, 3) = part
    gBuf(gCnt, 4) = cols
    gBuf(gCnt, 5) = det
End Sub

Private Sub ShowResult(secs As Double)
    Dim i As Long, j As Long, m As String, lim As Long, out() As Variant
    If gCnt = 0 Then
        MsgBox "OK - tat ca dong trong file moi da khop file goc." & vbLf & "(" & Format(secs, "0.0") & " giay)", vbInformation
        Exit Sub
    End If
    ReDim out(1 To gCnt, 1 To 5)
    For i = 1 To gCnt
        For j = 1 To 5
            out(i, j) = gBuf(i, j)
        Next j
    Next i
    gWs.Range("A2").Resize(gCnt, 5).Value = out
    gWs.Columns("A:D").AutoFit
    lim = gCnt: If lim > 15 Then lim = 15
    For i = 1 To lim
        m = m & gBuf(i, 1) & " - dong " & gBuf(i, 2) & " [" & gBuf(i, 3) & "]: " & gBuf(i, 4) & vbLf
    Next i
    If gCnt > lim Then m = m & "... con " & (gCnt - lim) & " loi nua (xem sheet Check_Result)"
    gWs.Activate
    MsgBox "Co " & gCnt & " loi (" & Format(secs, "0.0") & " giay):" & vbLf & vbLf & m, vbExclamation
End Sub
