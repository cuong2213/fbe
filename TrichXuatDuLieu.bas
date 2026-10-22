Attribute VB_Name = "TrichXuatDuLieu"
Option Explicit
'==============================================================================
' TRICH XUAT DU LIEU TU FILE NGUON VAO SHEET DATA (file dich = file chua macro)
' Cau truc sheet MODEL (chi doc cot A:J):
'   A        : danh sach Exterior Color (A2 tro xuong, A1 la tieu de)
'   B..E     : moi cot la 1 Model (hang 1 = ten Model), tu hang 2 = Part name
'   F        : cot trong ngan cach
'   G..J     : bang mapping  G = "Model_Part name" | H = Point | I = Row | J = Col
' Body Type = ten Model (lay tu hang 1 cua MODEL). MNV va MINIVAN duoc coi la 1.
' Moi lan chay: XOA het data cu o DATA roi ghi lai tu dau, KHONG loc trung.
' Dinh dang cot (cu): A STT | B Date | C Shift (khong dong toi) | D Week | E Part name
'   F Body Type | G Exterior Color | H Point | I N1 | J N3 | K Row | L Col
' Sheet DATA_FULL : danh sach DAY DU moi dong doc duoc (A:L).
' Sheet DATA      : chi cac dong co N1 > N_LIMIT hoac N3 > N_LIMIT (A:L) + 2 cot cong thuc
'   O = Size (luon = SIZE_VALUE = 10) | M = K&"-"&L (Row-Col, vd 1-12) | N = COUNTIFS dem so dong cung Row&Col, Week, Body Type, Point
' File nay chi dung ky tu ASCII de import vao VBE khong bi loi font.
'==============================================================================

'--- CAU HINH ---------------------------------------------------------------
Private Const SH_MODEL As String = "MODEL"
Private Const SH_DATA As String = "DATA"
Private Const SH_LOG As String = "LOG"
Private Const SH_FULL As String = "DATA_FULL"
Private Const MDL_FIRST_COL As Long = 2      'B: cot Model dau tien
Private Const MDL_LAST_COL As Long = 5       'E: cot Model cuoi cung
Private Const MAP_FIRST_COL As Long = 7      'G: bat dau bang mapping
Private Const MODEL_LAST_COL As Long = 10    'J: gioi han doc cua sheet MODEL
Private Const DATA_FIRST_ROW As Long = 2
Private Const WEEK_FORMULA As String = "ISOWEEKNUM"   'doi thanh WEEKNUM neu can
Private Const COL_DATE_FMT As String = "dd/mm/yyyy"
Private Const SIZE_VALUE As Long = 10          'DATA cot O (Size) luon bang gia tri nay
Private Const N_LIMIT As Double = 7            'dat khi N1 > 7 hoac N3 > 7 (lon hon han, khong gom bang 7)

Private Type TModel
    Name As String
    Tok1 As String
    Tok2 As String
    nParts As Long
    PartsO() As String
    PartsN() As String
End Type

Private mModels() As TModel
Private mNModels As Long
Private mColorsO() As String
Private mColorsN() As String
Private mNColors As Long
Private mMap As Object        'key = "MODEL PART|POINT" -> Array(Row, Col) hoac "DUP"
Private mOut As Collection    'moi dong: Array(Date,Part,Model,Color,Point,N1,N3,Row,Col,Hit)
Private mWritten As Long, mWrittenFull As Long
Private mLog As Collection
Private mSrcName As String
Private mRx As Object
Private mTables As Long, mAdded As Long, mErr As Long

'==============================================================================
Public Sub TrichXuatDuLieu()
    Dim wbDest As Workbook, wbSrc As Workbook, wb As Workbook
    Dim wsM As Worksheet, wsD As Worksheet, ws As Worksheet
    Dim srcPath As Variant, openedHere As Boolean
    Dim calc As XlCalculation, scr As Boolean, evt As Boolean
    Dim errMsg As String

    Set wbDest = ThisWorkbook
    On Error Resume Next
    Set wsM = wbDest.Worksheets(SH_MODEL)
    Set wsD = wbDest.Worksheets(SH_DATA)
    On Error GoTo 0
    If wsM Is Nothing Or wsD Is Nothing Then
        MsgBox "Thieu sheet " & SH_MODEL & " hoac " & SH_DATA & " trong file dich.", vbExclamation
        Exit Sub
    End If

    srcPath = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*", , "Chon file nguon")
    If VarType(srcPath) = vbBoolean Then Exit Sub
    If StrComp(CStr(srcPath), wbDest.FullName, vbTextCompare) = 0 Then
        MsgBox "Khong duoc chon chinh file dich lam file nguon.", vbExclamation
        Exit Sub
    End If

    calc = Application.Calculation
    scr = Application.ScreenUpdating
    evt = Application.EnableEvents
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set mLog = New Collection
    Set mOut = New Collection
    mWritten = 0: mWrittenFull = 0
    Set mMap = CreateObject("Scripting.Dictionary")
    Set mRx = CreateObject("VBScript.RegExp")
    mRx.Pattern = "(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})"
    mRx.Global = False
    mTables = 0: mAdded = 0: mErr = 0

    'File nguon: dung lai neu da mo, neu chua thi tu mo (chi doc)
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(srcPath), vbTextCompare) = 0 Then Set wbSrc = wb: Exit For
    Next wb
    If wbSrc Is Nothing Then
        Set wbSrc = Workbooks.Open(CStr(srcPath), UpdateLinks:=0, ReadOnly:=True)
        openedHere = True
    End If
    mSrcName = wbSrc.Name

    If Not LoadModel(wsM) Then GoTo Cleanup

    For Each ws In wbSrc.Worksheets
        ProcessSheet ws
    Next ws

    WriteData wsD
Cleanup:
    On Error Resume Next
    WriteLog wbDest
    If openedHere Then wbSrc.Close SaveChanges:=False
    Application.Calculation = calc
    Application.EnableEvents = evt
    Application.ScreenUpdating = scr
    If Len(errMsg) > 0 Then
        MsgBox "Dung do loi, DATA chua duoc ghi. Xem sheet LOG." & vbLf & errMsg, vbCritical
    Else
        MsgBox "Xong." & vbLf & "Bang xu ly: " & mTables & vbLf & "Dong nguon doc duoc: " & mAdded & _
               vbLf & "Ghi vao " & SH_FULL & ": " & mWrittenFull & _
               vbLf & "Ghi vao DATA (N1 hoac N3 > " & N_LIMIT & "): " & mWritten & _
               vbLf & "Dong/bang loi: " & mErr & _
               vbLf & "Chi tiet xem sheet LOG.", vbInformation
    End If
    Exit Sub
EH:
    errMsg = "Loi " & Err.Number & ": " & Err.Description
    AddLog "", "", "", "", "", "", "LOI HE THONG", "", errMsg
    Resume Cleanup
End Sub

'==============================================================================
' DOC SHEET MODEL (A:J)
'==============================================================================
Private Function LoadModel(wsM As Worksheet) As Boolean
    Dim lastR As Long, r As Long, c As Long, v As Variant
    Dim nm As String, k As Long, dic As Object, s As String, g As String, key As String

    lastR = 2
    For c = 1 To MODEL_LAST_COL
        r = wsM.Cells(wsM.Rows.Count, c).End(xlUp).Row
        If r > lastR Then lastR = r
    Next c
    v = wsM.Range(wsM.Cells(1, 1), wsM.Cells(lastR, MODEL_LAST_COL)).Value2

    'Mau
    Set dic = CreateObject("Scripting.Dictionary")
    ReDim mColorsO(1 To lastR): ReDim mColorsN(1 To lastR)
    mNColors = 0
    For r = 2 To lastR
        s = NormV(v(r, 1))
        If Len(s) > 0 Then
            If Not dic.Exists(s) Then
                dic.Add s, True
                mNColors = mNColors + 1
                mColorsO(mNColors) = Trim$(CStr(v(r, 1)))
                mColorsN(mNColors) = s
            End If
        End If
    Next r

    'Model + Part name
    mNModels = 0
    For c = MDL_FIRST_COL To MDL_LAST_COL
        nm = NormV(v(1, c))
        If Len(nm) > 0 Then
            mNModels = mNModels + 1
            ReDim Preserve mModels(1 To mNModels)
            With mModels(mNModels)
                .Name = Trim$(CStr(v(1, c)))
                .Tok1 = nm
                .Tok2 = ""
                If nm = "MNV" Then .Tok2 = "MINIVAN"
                If nm = "MINIVAN" Then .Tok2 = "MNV"
                ReDim .PartsO(1 To lastR): ReDim .PartsN(1 To lastR)
                .nParts = 0
                Set dic = CreateObject("Scripting.Dictionary")
                For r = 2 To lastR
                    s = NormV(v(r, c))
                    If Len(s) > 0 Then
                        If Not dic.Exists(s) Then
                            dic.Add s, True
                            .nParts = .nParts + 1
                            .PartsO(.nParts) = Trim$(CStr(v(r, c)))
                            .PartsN(.nParts) = s
                        End If
                    End If
                Next r
            End With
        End If
    Next c

    'Mapping G:J
    For r = 2 To lastR
        g = NormV(v(r, MAP_FIRST_COL))
        If Len(g) > 0 Then
            key = g & "|" & PointKey(v(r, MAP_FIRST_COL + 1))
            If mMap.Exists(key) Then
                mMap(key) = "DUP"
            Else
                mMap.Add key, Array(v(r, MAP_FIRST_COL + 2), v(r, MAP_FIRST_COL + 3))
            End If
        End If
    Next r

    If mNModels = 0 Then
        AddLog "", "", "", "", "", "", "LOI", "MODEL", "Khong doc duoc Model nao o hang 1 (cot B:E) cua sheet MODEL"
        mErr = mErr + 1
        Exit Function
    End If
    If mNColors = 0 Then
        AddLog "", "", "", "", "", "", "LOI", "Exterior Color", "Cot A cua sheet MODEL khong co mau"
        mErr = mErr + 1
        Exit Function
    End If
    LoadModel = True
End Function

'==============================================================================
' QUET MOT SHEET NGUON
'==============================================================================
Private Sub ProcessSheet(ws As Worksheet)
    Dim ur As Range, v As Variant
    Dim r0 As Long, c0 As Long, nR As Long, nC As Long
    Dim r As Long, c As Long, x As Variant, mi As Long
    Dim titles As Collection, tabs As Collection, tt As Variant, t2 As Variant, cEnd As Long

    Set ur = ws.UsedRange
    If ur.Cells.CountLarge < 2 Then Exit Sub
    v = ur.Value2
    r0 = ur.Row: c0 = ur.Column
    nR = UBound(v, 1): nC = UBound(v, 2)

    Set titles = New Collection
    For r = 1 To nR
        For c = 1 To nC
            x = v(r, c)
            If VarType(x) = vbString Then
                If Len(x) >= 3 Then
                    mi = FindModel(Norm(CStr(x)))
                    If mi <> 0 Then
                        titles.Add Array(r, c, mi, CStr(x))
                    End If
                End If
            End If
        Next c
    Next r

    'Moi bang chiem cac cot tu o tieu de den truoc o tieu de ke ben (cung hang)
    Set tabs = New Collection
    For Each tt In titles
        cEnd = nC
        For Each t2 In titles
            If t2(0) = tt(0) And t2(1) > tt(1) And t2(1) - 1 < cEnd Then cEnd = t2(1) - 1
        Next t2
        tabs.Add Array(tt(0), tt(1), tt(2), tt(3), cEnd)
    Next tt

    For Each tt In tabs
        ProcessTable ws, v, r0, c0, nR, nC, tt, tabs
    Next tt
End Sub

Private Sub ProcessTable(ws As Worksheet, v As Variant, ByVal r0 As Long, ByVal c0 As Long, _
                         ByVal nR As Long, ByVal nC As Long, tt As Variant, tabs As Collection)
    Dim tr As Long, tc As Long, mi As Long, titleText As String, addr As String
    Dim tn As String, modelName As String, partO As String, colorO As String
    Dim pi As Long, ci As Long, dVal As Double, dState As Long, dOut As Variant
    Dim hr As Long, c As Long, h As String, missing As String, dupHdr As String
    Dim cP As Long, cN1 As Long, cN3 As Long, cS As Long, cE As Long
    Dim rr As Long, pv As Variant, n1 As Variant, n3 As Variant, n1v As Variant, n3v As Variant
    Dim prefix As String, mk As String, mp As Variant, rowAddr As String
    Dim nAdd As Long, nErr As Long, bad As Boolean
    Dim dKey As String, hit As Long

    tr = tt(0): tc = tt(1): mi = tt(2): titleText = tt(3)
    cS = tc: cE = tt(4)
    addr = ws.Cells(r0 + tr - 1, c0 + tc - 1).Address(False, False)
    mTables = mTables + 1

    If mi < 0 Then
        AddLog ws.Name, addr, titleText, "(nhieu Model)", "", "", "LOI", "Model", "Tieu de khop nhieu Model, khong chon tuy tien"
        mErr = mErr + 1: Exit Sub
    End If
    modelName = mModels(mi).Name
    tn = Norm(titleText)

    'Part name (khop dai nhat)
    pi = BestMatch(tn, mModels(mi).PartsN, mModels(mi).nParts)
    If pi = 0 Then
        AddLog ws.Name, addr, titleText, modelName, "", "", "LOI", "Part name", "Model hop le nhung khong co Part name nao khop"
        mErr = mErr + 1: Exit Sub
    ElseIf pi < 0 Then
        AddLog ws.Name, addr, titleText, modelName, "", "", "LOI", "Part name", "Nhieu Part name cung khop (do dai bang nhau)"
        mErr = mErr + 1: Exit Sub
    End If
    partO = mModels(mi).PartsO(pi)

    'Exterior Color (khop dai nhat)
    ci = BestMatch(tn, mColorsN, mNColors)
    If ci = 0 Then
        AddLog ws.Name, addr, titleText, modelName, partO, "", "LOI", "Exterior Color", "Khong tim thay mau nao trong danh muc (cot A MODEL) khop voi tieu de"
        mErr = mErr + 1: Exit Sub
    ElseIf ci < 0 Then
        AddLog ws.Name, addr, titleText, modelName, partO, "", "LOI", "Exterior Color", "Nhieu mau cung khop (do dai bang nhau)"
        mErr = mErr + 1: Exit Sub
    End If
    colorO = mColorsO(ci)

    'Ngay
    dState = ParseDate(titleText, dVal)
    If dState = 1 Then
        dOut = dVal: dKey = CStr(CLng(dVal))
    Else
        dOut = Empty: dKey = ""
        If dState = 2 Then AddLog ws.Name, addr, titleText, modelName, partO, colorO, "CANH BAO", "Ngay", "Ngay trong tieu de khong hop le, de trong"
    End If

    'Hang tieu de cot (ngay duoi tieu de)
    hr = tr + 1
    If hr > nR Then
        AddLog ws.Name, addr, titleText, modelName, partO, colorO, "LOI", "Tieu de cot", "Khong co hang tieu de cot duoi tieu de bang"
        mErr = mErr + 1: Exit Sub
    End If
    For c = cS To cE
        If Not IsError(v(hr, c)) Then
            h = Norm(CStr(v(hr, c)))
            Select Case h
                Case "POINT": If cP <> 0 Then dupHdr = dupHdr & "Point " Else cP = c
                Case "N1": If cN1 <> 0 Then dupHdr = dupHdr & "N1 " Else cN1 = c
                Case "N3": If cN3 <> 0 Then dupHdr = dupHdr & "N3 " Else cN3 = c
            End Select
        End If
    Next c
    If cP = 0 Then missing = missing & "Point "
    If cN1 = 0 Then missing = missing & "N1 "
    If cN3 = 0 Then missing = missing & "N3 "
    If Len(missing) > 0 Or Len(dupHdr) > 0 Then
        AddLog ws.Name, addr, titleText, modelName, partO, colorO, "LOI", "Tieu de cot", _
               IIf(Len(missing) > 0, "Thieu cot bat buoc: " & Trim$(missing) & ". ", "") & _
               IIf(Len(dupHdr) > 0, "Trung cot: " & Trim$(dupHdr), "")
        mErr = mErr + 1: Exit Sub
    End If

    prefix = Norm(modelName & " " & partO)

    'Doc du lieu
    For rr = hr + 1 To nR
        If TitleInBand(tabs, rr, cS, cE) Then Exit For
        If RowBlank(v, rr, cS, cE) Then Exit For
        pv = v(rr, cP): n1 = v(rr, cN1): n3 = v(rr, cN3)
        If IsBlank(pv) And IsBlank(n1) And IsBlank(n3) Then GoTo NextRow
        rowAddr = ws.Cells(r0 + rr - 1, c0 + cP - 1).Address(False, False)

        If IsBlank(pv) Or IsError(pv) Then
            AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "LOI DONG", "Point", "Dong " & (r0 + rr - 1) & ": Point trong/loi nen khong tra duoc mapping"
            nErr = nErr + 1: GoTo NextRow
        End If

        bad = False
        If Not ToNum(n1, n1v) Then
            AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "LOI DONG", "N1", "Dong " & (r0 + rr - 1) & ": N1 khong phai so"
            bad = True
        End If
        If Not ToNum(n3, n3v) Then
            AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "LOI DONG", "N3", "Dong " & (r0 + rr - 1) & ": N3 khong phai so"
            bad = True
        End If
        If bad Then nErr = nErr + 1: GoTo NextRow
        If IsEmpty(n1v) Then AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "CANH BAO", "N1", "Dong " & (r0 + rr - 1) & ": N1 trong, de trong (khong thay bang 0)"
        If IsEmpty(n3v) Then AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "CANH BAO", "N3", "Dong " & (r0 + rr - 1) & ": N3 trong, de trong (khong thay bang 0)"

        mk = prefix & "|" & PointKey(pv)
        If Not mMap.Exists(mk) Then
            AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "LOI DONG", "Row/Col", "Dong " & (r0 + rr - 1) & ": khong co mapping cho " & mk
            nErr = nErr + 1: GoTo NextRow
        End If
        If VarType(mMap(mk)) = vbString Then
            AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "LOI DONG", "Row/Col", "Dong " & (r0 + rr - 1) & ": nhieu mapping trung khoa " & mk
            nErr = nErr + 1: GoTo NextRow
        End If
        mp = mMap(mk)
        If IsBlank(mp(0)) Or IsBlank(mp(1)) Or IsError(mp(0)) Or IsError(mp(1)) Then
            AddLog ws.Name, rowAddr, titleText, modelName, partO, colorO, "LOI DONG", "Row/Col", "Dong " & (r0 + rr - 1) & ": mapping thieu Row hoac Col cho " & mk
            nErr = nErr + 1: GoTo NextRow
        End If

        'Dem dong dat dieu kien: N1 > N_LIMIT hoac N3 > N_LIMIT (o trong khong tinh)
        hit = 0
        If Not IsEmpty(n1v) Then
            If n1v > N_LIMIT Then hit = 1
        End If
        If hit = 0 And Not IsEmpty(n3v) Then
            If n3v > N_LIMIT Then hit = 1
        End If

        mOut.Add Array(dOut, partO, modelName, colorO, pv, n1v, n3v, mp(0), mp(1), hit)
        nAdd = nAdd + 1
NextRow:
    Next rr

    mAdded = mAdded + nAdd: mErr = mErr + nErr
    AddLog ws.Name, addr, titleText, modelName, partO, colorO, IIf(nErr > 0, "CO LOI DONG", "OK"), "", _
           "Doc " & nAdd & " dong; loi " & nErr
End Sub

'==============================================================================
' GHI DATA VA LOG
'==============================================================================
Private Sub WriteData(wsD As Worksheet)
    Dim wsF As Worksheet
    If mOut.Count = 0 Then Exit Sub          'khong co gi de ghi -> giu nguyen data cu

    On Error Resume Next
    Set wsF = wsD.Parent.Worksheets(SH_FULL)
    On Error GoTo 0
    If wsF Is Nothing Then
        Set wsF = wsD.Parent.Worksheets.Add(After:=wsD)
        wsF.Name = SH_FULL
    End If

    mWrittenFull = WriteSheet(wsF, False)
    mWritten = WriteSheet(wsD, True)
    If mWritten = 0 Then AddLog "", "", "", "", "", "", "CANH BAO", "", "Khong co dong nao co N1 hoac N3 > " & N_LIMIT & " de ghi vao DATA"
End Sub

'Xoa du lieu cu roi ghi lai. onlyHit = True: chi dong co N1 hoac N3 > N_LIMIT, them cot M, N (cong thuc).
'Tra ve so dong da ghi.
Private Function WriteSheet(ws As Worksheet, ByVal onlyHit As Boolean) As Long
    Dim n As Long, k As Long, j As Long, lastNew As Long, it As Variant
    Dim a1() As Variant, a2() As Variant, r1 As Long

    ClearOldData ws

    ws.Range("A1:L1").Value2 = Array("STT", "DATE/NGAY", "Shift/Ca", "Week", "Part name", "Body Type", _
        "Exterior Color", "Point", "N1", "N3", "Row", "Col")
    If onlyHit Then
        ws.Range("M1:O1").Value2 = Array("Row-Col", "SL (N1 hoac N3 > " & N_LIMIT & ")", "Size")
    End If

    n = mOut.Count
    ReDim a1(1 To n, 1 To 2)     'A:B
    ReDim a2(1 To n, 1 To 8)     'E:L
    For j = 1 To n
        it = mOut(j)
        If Not onlyHit Or it(9) = 1 Then
            k = k + 1
            a1(k, 1) = k             'STT tinh lai tu 1
            a1(k, 2) = it(0)         'Date
            a2(k, 1) = it(1)         'Part name
            a2(k, 2) = it(2)         'Body Type (= Model)
            a2(k, 3) = it(3)         'Exterior Color
            a2(k, 4) = it(4)         'Point
            a2(k, 5) = it(5)         'N1
            a2(k, 6) = it(6)         'N3
            a2(k, 7) = it(7)         'Row
            a2(k, 8) = it(8)         'Col
        End If
    Next j
    If k = 0 Then Exit Function

    lastNew = DATA_FIRST_ROW + k - 1
    ws.Range(ws.Cells(DATA_FIRST_ROW, 2), ws.Cells(lastNew, 2)).NumberFormat = COL_DATE_FMT
    ws.Range(ws.Cells(DATA_FIRST_ROW, 1), ws.Cells(lastNew, 2)).Value2 = ResizeRows(a1, k)
    ws.Range(ws.Cells(DATA_FIRST_ROW, 5), ws.Cells(lastNew, 12)).Value2 = ResizeRows(a2, k)

    'Cot C (Shift) bo qua. Cot D (Week) tinh tu cot B.
    r1 = DATA_FIRST_ROW
    ws.Range(ws.Cells(r1, 4), ws.Cells(lastNew, 4)).Formula = _
        "=IF(B" & r1 & "="""",""""," & WEEK_FORMULA & "(B" & r1 & "))"

    If onlyHit Then
        ws.Range(ws.Cells(r1, 13), ws.Cells(lastNew, 13)).Formula = "=K" & r1 & "&""-""&L" & r1
        ws.Range(ws.Cells(r1, 14), ws.Cells(lastNew, 14)).Formula = _
            "=COUNTIFS($M:$M,M" & r1 & ",$D:$D,D" & r1 & ",$F:$F,F" & r1 & ",$H:$H,H" & r1 & ")"
        ws.Range(ws.Cells(r1, 15), ws.Cells(lastNew, 15)).Value2 = SIZE_VALUE    'O: Size = 10
    End If
    WriteSheet = k
End Function

'Cat mang 2 chieu (chi so 1..) xuong k dong dau
Private Function ResizeRows(arr() As Variant, ByVal k As Long) As Variant
    Dim r() As Variant, i As Long, c As Long, nc As Long
    If k = UBound(arr, 1) Then ResizeRows = arr: Exit Function
    nc = UBound(arr, 2)
    ReDim r(1 To k, 1 To nc)
    For i = 1 To k
        For c = 1 To nc
            r(i, c) = arr(i, c)
        Next c
    Next i
    ResizeRows = r
End Function

'Xoa noi dung cu (giu dinh dang va tieu de hang 1)
Private Sub ClearOldData(ws As Worksheet)
    Dim lastR As Long, c As Long, r As Long
    lastR = DATA_FIRST_ROW
    For c = 1 To 15
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > lastR Then lastR = r
    Next c

    ws.Range(ws.Cells(DATA_FIRST_ROW, 1), ws.Cells(lastR, 2)).ClearContents    'A:B
    ws.Range(ws.Cells(DATA_FIRST_ROW, 4), ws.Cells(lastR, 15)).ClearContents   'D:O

    'Cot C (Shift): chi xoa gia tri go tay, giu lai cong thuc neu co
    On Error Resume Next
    ws.Range(ws.Cells(DATA_FIRST_ROW, 3), ws.Cells(lastR, 3)).SpecialCells(xlCellTypeConstants).ClearContents
    On Error GoTo 0
End Sub

Private Sub WriteLog(wbDest As Workbook)
    Dim wsL As Worksheet, n As Long, i As Long, k As Long, a() As Variant, it As Variant
    Dim t As String

    n = mLog.Count
    On Error Resume Next
    Set wsL = wbDest.Worksheets(SH_LOG)
    On Error GoTo LogErr
    If wsL Is Nothing Then
        If n = 0 Then Exit Sub
        Set wsL = wbDest.Worksheets.Add(After:=wbDest.Worksheets(wbDest.Worksheets.Count))
        wsL.Name = SH_LOG
    End If

    'Xoa sach sheet LOG (noi dung + dinh dang), moi lan chay ghi lai tu dau
    wsL.Cells.Clear

    wsL.Range("A1:K1").Value2 = Array("Thoi gian", "File nguon", "Sheet nguon", "O tieu de", "Noi dung tieu de", _
        "Model", "Part name", "Exterior Color", "Trang thai", "Truong thieu/khong khop", "Chi tiet")
    wsL.Range("A1:K1").Font.Bold = True
    If n = 0 Then Exit Sub

    ReDim a(1 To n, 1 To 11)
    For i = 1 To n
        it = mLog(i)
        a(i, 1) = CDbl(it(0))
        For k = 1 To 10
            t = CStr(it(k))
            If Left$(t, 1) = "=" Then t = "'" & t     'tranh bi hieu thanh cong thuc
            a(i, k + 1) = t
        Next k
    Next i
    wsL.Range(wsL.Cells(2, 1), wsL.Cells(n + 1, 1)).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    wsL.Range(wsL.Cells(2, 1), wsL.Cells(n + 1, 11)).Value2 = a
    Exit Sub
LogErr:
    MsgBox "Loi khi ghi sheet LOG: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub AddLog(ByVal sh As String, ByVal addr As String, ByVal title As String, ByVal mdl As String, _
                   ByVal part As String, ByVal clr As String, ByVal status As String, _
                   ByVal fld As String, ByVal detail As String)
    mLog.Add Array(Now, mSrcName, sh, addr, title, mdl, part, clr, status, fld, detail)
End Sub

'==============================================================================
' HAM TRO GIUP
'==============================================================================
Private Function LastDataRow(ws As Worksheet) As Long
    Dim c As Variant, r As Long, m As Long
    m = 1
    For Each c In Array(1, 2, 5, 8)
        r = ws.Cells(ws.Rows.Count, CLng(c)).End(xlUp).Row
        If r > m Then m = r
    Next c
    LastDataRow = m
End Function

Private Function Norm(ByVal s As String) As String
    s = UCase$(s)
    s = Replace(s, "_", " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = Trim$(s)
End Function

Private Function NormV(x As Variant) As String
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then Exit Function
    NormV = Norm(CStr(x))
End Function

Private Function PointKey(x As Variant) As String
    PointKey = NormV(x)
End Function

Private Function IsBlank(x As Variant) As Boolean
    If IsError(x) Then Exit Function
    If IsEmpty(x) Then IsBlank = True: Exit Function
    If VarType(x) = vbString Then IsBlank = (Len(Trim$(CStr(x))) = 0)
End Function

Private Function TitleInBand(tabs As Collection, ByVal r As Long, ByVal cS As Long, ByVal cE As Long) As Boolean
    Dim t As Variant
    For Each t In tabs
        If t(0) = r And t(1) >= cS And t(1) <= cE Then TitleInBand = True: Exit Function
    Next t
End Function

Private Function RowBlank(v As Variant, ByVal r As Long, ByVal cS As Long, ByVal cE As Long) As Boolean
    Dim c As Long, x As Variant
    For c = cS To cE
        x = v(r, c)
        If IsError(x) Then Exit Function
        If Not IsEmpty(x) Then
            If VarType(x) <> vbString Then Exit Function
            If Len(Trim$(CStr(x))) > 0 Then Exit Function
        End If
    Next c
    RowBlank = True
End Function

'Chuyen o thanh so. Trong -> Empty (hop le). Khong phai so/loi -> False.
Private Function ToNum(x As Variant, ByRef out As Variant) As Boolean
    out = Empty
    If IsError(x) Then Exit Function
    If IsBlank(x) Then ToNum = True: Exit Function
    If VarType(x) = vbString Then
        If Not IsNumeric(x) Then Exit Function
        out = CDbl(x): ToNum = True
    ElseIf IsNumeric(x) Then
        out = CDbl(x): ToNum = True
    End If
End Function

'0 = khong co Model, -1 = nhieu Model, >0 = chi so Model
Private Function FindModel(ByVal t As String) As Long
    Dim i As Long, cnt As Long, idx As Long, p As String
    p = " " & t & " "
    For i = 1 To mNModels
        If InStr(1, p, " " & mModels(i).Tok1 & " ", vbBinaryCompare) > 0 Then
            cnt = cnt + 1: idx = i
        ElseIf Len(mModels(i).Tok2) > 0 Then
            If InStr(1, p, " " & mModels(i).Tok2 & " ", vbBinaryCompare) > 0 Then cnt = cnt + 1: idx = i
        End If
    Next i
    If cnt = 0 Then
        FindModel = 0
    ElseIf cnt = 1 Then
        FindModel = idx
    Else
        FindModel = -1
    End If
End Function

'Khop dai nhat. 0 = khong khop, -1 = hoa nhau (cung do dai, khac ten), >0 = chi so
Private Function BestMatch(ByVal t As String, arr() As String, ByVal n As Long) As Long
    Dim i As Long, bestLen As Long, best As Long, tie As Boolean, p As String
    If n <= 0 Then Exit Function
    p = " " & t & " "
    For i = 1 To n
        If Len(arr(i)) > 0 Then
            If InStr(1, p, " " & arr(i) & " ", vbBinaryCompare) > 0 Then
                If Len(arr(i)) > bestLen Then
                    bestLen = Len(arr(i)): best = i: tie = False
                ElseIf Len(arr(i)) = bestLen Then
                    tie = True
                End If
            End If
        End If
    Next i
    If best = 0 Then
        BestMatch = 0
    ElseIf tie Then
        BestMatch = -1
    Else
        BestMatch = best
    End If
End Function

'0 = khong co ngay, 1 = hop le (d), 2 = co dinh dang ngay nhung khong hop le
Private Function ParseDate(ByVal s As String, ByRef d As Double) As Long
    Dim m As Object, dd As Long, mm As Long, yy As Long, dt As Date
    Set m = mRx.Execute(s)
    If m.Count = 0 Then Exit Function
    dd = CLng(m(0).SubMatches(0)): mm = CLng(m(0).SubMatches(1)): yy = CLng(m(0).SubMatches(2))
    ParseDate = 2
    If yy < 1900 Or yy > 2200 Or mm < 1 Or mm > 12 Or dd < 1 Or dd > 31 Then Exit Function
    dt = DateSerial(yy, mm, dd)
    If Day(dt) <> dd Or Month(dt) <> mm Then Exit Function
    d = CDbl(dt)
    ParseDate = 1
End Function
