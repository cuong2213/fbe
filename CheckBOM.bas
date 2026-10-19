Attribute VB_Name = "modCheckBOM"
Option Explicit

' Kiem tra file MOI (Super BOM / Single BOM) da khop file GOC chua.
' Chay macro: CheckBOM
' Luu y: de ASCII (khong dau) vi VBE khong hien thi Unicode.

Private gRe As Object
Private gWs As Worksheet
Private gRow As Long
Private gCnt As Long

Public Sub CheckBOM()
    Dim runDate As Date, pNew As String, pOrg As String
    Dim wbN As Workbook, wbO As Workbook, wsN As Worksheet
    Dim rSup As Long, rSin As Long, eN As Long, eD As String

    runDate = ParseDMY(InputBox("Ngay chay (dd/mm/yyyy):", "Kiem tra BOM", Format(Date, "dd/mm/yyyy")))
    If runDate = 0 Then Exit Sub

    pNew = PickFile("Chon FILE MOI")
    If Len(pNew) = 0 Then Exit Sub

    Application.ScreenUpdating = False
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

    ShowResult

Fin:
    eN = Err.Number: eD = Err.Description
    On Error Resume Next
    If Not wbO Is Nothing Then wbO.Close False
    If Not wbN Is Nothing Then wbN.Close False
    Application.ScreenUpdating = True
    If eN <> 0 Then MsgBox "Loi: " & eD, vbCritical
End Sub

' ===================== SUPER BOM =====================
Private Sub CheckSuper(wsN As Worksheet, mRow As Long, wsO As Worksheet, runDate As Date)
    Dim hN As Long, cP As Long, cU As Long, cQ As Long, cO As Long, cV As Long, lastColN As Long
    Dim hO As Long, kM As Long, kF1 As Long, kT1 As Long, kF2 As Long, kT2 As Long
    Dim kQ As Long, kU As Long, kB As Long, kS As Long, lastO As Long, lastColO As Long
    Dim vO As Variant, d As Object, col As Collection, idx As Variant
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
    vO = wsO.Range(wsO.Cells(hO + 1, 1), wsO.Cells(lastO, lastColO)).Value

    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(vO, 1)
        If Len(Trim(CStr(vO(i, kM)))) > 0 Then
            If RowValid(vO, i, kF1, kT1, runDate) And RowValid(vO, i, kF2, kT2, runDate) Then
                key = UCase(Trim(CStr(vO(i, kM))))
                If Not d.Exists(key) Then
                    Set col = New Collection
                    d.Add key, col
                End If
                d(key).Add i
            End If
        End If
    Next i

    r = hN + 1
    Do While Application.WorksheetFunction.CountA(wsN.Range(wsN.Cells(r, 1), wsN.Cells(r, lastColN))) > 0
        part = Trim(CStr(wsN.Cells(r, cP).Value))
        key = UCase(part)
        If Not d.Exists(key) Then
            AddRes "Super BOM", r, part, "Part Number", "Khong co trong file goc (hoac chua/het hieu luc)"
        Else
            bestN = 999
            For Each idx In d(key)
                det = DiffSuper(wsN.Cells(r, cU).Value, wsN.Cells(r, cQ).Value, wsN.Cells(r, cO).Value, wsN.Cells(r, cV).Value, _
                                vO(idx, kU), vO(idx, kQ), vO(idx, kB), vO(idx, kS), cols, n)
                If n < bestN Then bestN = n: best = det: bestCols = cols
                If n = 0 Then Exit For
            Next idx
            If bestN > 0 Then AddRes "Super BOM", r, part, bestCols, best
        End If
        r = r + 1
    Loop
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
    Dim hN As Long, cP As Long, cC As Long, cU As Long, cQ As Long, lastColN As Long
    Dim hO As Long, kM As Long, kC As Long, kU As Long, kQ As Long, kF As Long, kT As Long
    Dim lastColO As Long, lastO As Long, vO As Variant, d As Object, col As Collection, idx As Variant
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
    vO = wsO.Range(wsO.Cells(hO + 1, 1), wsO.Cells(lastO, lastColO)).Value

    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(vO, 1)
        If Len(Trim(CStr(vO(i, kM)))) > 0 Then
            If RowValid(vO, i, kF, kT, runDate) Then
                key = UCase(Trim(CStr(vO(i, kM)))) & "|" & UCase(Trim(CStr(vO(i, kC))))
                If Not d.Exists(key) Then
                    Set col = New Collection
                    d.Add key, col
                End If
                d(key).Add i
            End If
        End If
    Next i

    r = hN + 1
    Do While Application.WorksheetFunction.CountA(wsN.Range(wsN.Cells(r, 1), wsN.Cells(r, lastColN))) > 0
        part = Trim(CStr(wsN.Cells(r, cP).Value))
        child = Trim(CStr(wsN.Cells(r, cC).Value))
        If Len(part) > 0 And Len(child) > 0 Then   ' dong khong co Old part (them moi) -> bo qua
            key = UCase(part) & "|" & UCase(child)
            If Not d.Exists(key) Then
                AddRes "Single BOM", r, part & " / " & child, "Old Parent Part / Old Child part", "Khong co cap Material-Component nay trong file goc (hoac het hieu luc)"
            Else
                bestN = 999
                For Each idx In d(key)
                    det = "": cols = "": n = 0
                    If UCase(Trim(CStr(wsN.Cells(r, cU).Value))) <> UCase(Trim(CStr(vO(idx, kU)))) Then _
                        AddDiff det, cols, n, "Old UOM", CStr(wsN.Cells(r, cU).Value), CStr(vO(idx, kU))
                    If Not SameNum(wsN.Cells(r, cQ).Value, vO(idx, kQ)) Then _
                        AddDiff det, cols, n, "Old Q'ty", CStr(wsN.Cells(r, cQ).Value), CStr(vO(idx, kQ))
                    If n < bestN Then bestN = n: best = det: bestCols = cols
                    If n = 0 Then Exit For
                Next idx
                If bestN > 0 Then AddRes "Single BOM", r, part & " / " & child, bestCols, best
            End If
        End If
        r = r + 1
    Loop
End Sub

' ===================== HELPERS =====================
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
        If Not IsEmpty(d) Then If CDate(d) > runDate Then RowValid = False
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
Private Function NormVar(ByVal s As String) As String
    Dim d As Object, ms As Object, i As Long, j As Long, t As Variant, k As Variant
    Dim a() As String, tmp As String
    s = UCase(Trim(s))
    If Len(s) = 0 Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    If InStr(s, "=") > 0 Then
        If gRe Is Nothing Then
            Set gRe = CreateObject("VBScript.RegExp")
            gRe.Global = True
            gRe.Pattern = "=\s*'?([^'\s()]+)'?"
        End If
        Set ms = gRe.Execute(s)
        For i = 0 To ms.Count - 1
            d(CStr(ms(i).SubMatches(0)) & "#" & d.Count) = 1
        Next i
    Else
        For Each t In Array(",", ";", "/", "|", "'", """", vbLf, vbCr, vbTab)
            s = Replace(s, CStr(t), " ")
        Next t
        For Each t In Split(s, " ")
            If Len(t) > 0 And t <> "OR" Then d(CStr(t) & "#" & d.Count) = 1
        Next t
    End If
    If d.Count = 0 Then Exit Function
    ReDim a(0 To d.Count - 1)
    j = 0
    For Each k In d.Keys
        a(j) = Left(CStr(k), InStrRev(CStr(k), "#") - 1): j = j + 1
    Next k
    For i = 1 To UBound(a)
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

' ===================== KET QUA =====================
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
    gWs.Range("A1:E1").Value = Array("Loai BOM", "Dong (file moi)", "Part", "Cot loi", "Chi tiet")
    gWs.Range("A1:E1").Font.Bold = True
    gRow = 2: gCnt = 0
End Sub

Private Sub AddRes(typ As String, rw As Long, part As String, cols As String, det As String)
    gWs.Cells(gRow, 1).Value = typ
    gWs.Cells(gRow, 2).Value = rw
    gWs.Cells(gRow, 3).NumberFormat = "@"
    gWs.Cells(gRow, 3).Value = part
    gWs.Cells(gRow, 4).Value = cols
    gWs.Cells(gRow, 5).Value = det
    gRow = gRow + 1: gCnt = gCnt + 1
End Sub

Private Sub ShowResult()
    Dim i As Long, m As String, lim As Long
    gWs.Columns("A:E").AutoFit
    If gCnt = 0 Then
        MsgBox "OK - tat ca dong trong file moi da khop file goc.", vbInformation
        Exit Sub
    End If
    lim = gCnt: If lim > 15 Then lim = 15
    For i = 2 To lim + 1
        m = m & gWs.Cells(i, 1).Value & " - dong " & gWs.Cells(i, 2).Value & " [" & gWs.Cells(i, 3).Value & "]: " & gWs.Cells(i, 4).Value & vbLf
    Next i
    If gCnt > lim Then m = m & "... con " & (gCnt - lim) & " loi nua (xem sheet Check_Result)"
    gWs.Activate
    MsgBox "Co " & gCnt & " loi:" & vbLf & vbLf & m, vbExclamation
End Sub
