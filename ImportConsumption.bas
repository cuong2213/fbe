Option Explicit

'==============================================================================
'  IMPORT TIEU HAO SON TU REPORT CONSUMPTION (PRIMER / BASE / CLEAR)
'  Chay macro: ImportConsumption   (dat trong chinh file Dinh_muc_son)
'
'  - Chon folder -> macro tu tim file co ten chua "primer", "base", "clear"
'  - PRIMER / CLEAR : Type = PART NAME, ghi vao block PRIMER / CLEAR
'  - BASE           : them Color -> ghi vao block mau tuong ung
'  - Cot ngay = NGAY CHAY MACRO (hang 13 = ngay, dinh dang dd-mmm).
'    Chua co cot ngay hom nay -> tu chen cot moi o ca 3 sheet; hang 13 cua
'    TOTAL_SKID / ROBOT_TIEU_HAO_1PART lien ket ve hang 13 cua ROBOT_TIEU_HAO_1SKID
'  - Gia tri ghi = consumption XUAT HIEN NHIEU NHAT (mode) theo tung part
'    (bang nhau -> lay gia tri xuat hien truoc trong report)
'  - Type/Color khong khop ten sheet -> khai bao 1 lan o sheet MAP_TYPE
'==============================================================================

Private Const SH_ROBOT As String = "ROBOT_TIEU_HAO_1SKID"
Private Const SH_TOTAL As String = "TOTAL_SKID"
Private Const SH_PART As String = "ROBOT_TIEU_HAO_1PART"
Private Const SH_MAP As String = "MAP_TYPE"

Private Const FIRST_DAY_COL As Long = 6        'cot F = ngay dau tien
Private Const ROW_WEEKDAY As Long = 2          'hang "Thu"
Private Const ROW_DAY As Long = 13             'hang ngay (gia tri la ngay that)
Private Const FIRST_DATA_ROW As Long = 14

Private Const ONLY_PRODUCT As Boolean = True   'True: chi lay dong Ghost Car = Product

Private mCnt As Object          'key "dong|gia tri" -> so lan xuat hien
Private mNum As Object          'key "dong|gia tri" -> gia tri (so)
Private mBlk As Object          'ten block (chuan hoa) -> 1
Private mPart As Object         'block|part (chuan hoa) -> so dong trong ROBOT
Private mTypeMap As Object      'MAP_TYPE: Type -> part name
Private mColorMap As Object     'MAP_TYPE: Color -> block
Private mNewType As Object      'Type chua co trong MAP_TYPE
Private mNewColor As Object     'Color chua co trong MAP_TYPE
Private mMiss As Object         'block|part khong co dong trong ROBOT
Private mRead As Long, mUsed As Long
Private mSkipGhost As Long, mSkipBad As Long, mSkipMap As Long
Private mWarn As String

'------------------------------------------------------------------------------
Public Sub ImportConsumption()
    Dim wsR As Worksheet, files As Collection, best As Object, bestN As Object
    Dim folder As String, it As Variant, p() As String, k As Variant
    Dim calcOld As Long, nFile As Long, col As Long, rw As Long, n As Long
    Dim nW As Long, nFormula As Long, added As Boolean, msg As String

    calcOld = Application.Calculation
    On Error GoTo EH

    Set wsR = ThisWorkbook.Worksheets(SH_ROBOT)
    folder = PickFolder()
    If Len(folder) = 0 Then Exit Sub

    Set files = ListReportFiles(folder)
    If files.Count = 0 Then
        MsgBox "Khong thay file nao co ten chua primer / base / clear trong folder nay.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    InitState
    BuildBlocks wsR
    LoadMaps

    For Each it In files
        p = Split(CStr(it), "|")
        ReadReport folder & "\" & p(0), p(1)
        nFile = nFile + 1
    Next it

    AppendNewMaps
    wsR.Activate

    If mCnt.Count = 0 Then
        MsgBox "Khong co dong du lieu hop le de ghi." & vbCr & "Da doc " & mRead & " dong report." & BuildNotes(), vbExclamation
        GoTo Done
    End If

    'gia tri xuat hien nhieu nhat cho tung dong part
    Set best = NewDict(): Set bestN = NewDict()
    For Each k In mCnt.Keys
        p = Split(CStr(k), "|")
        rw = CLng(p(0))
        n = mCnt(k)
        If Not bestN.Exists(CStr(rw)) Then
            best(CStr(rw)) = mNum(k): bestN(CStr(rw)) = n
        ElseIf n > bestN(CStr(rw)) Then
            best(CStr(rw)) = mNum(k): bestN(CStr(rw)) = n
        End If
    Next k

    'cot ngay hom nay (tu them neu chua co)
    col = GetTodayCol(wsR, added)
    Application.Calculate

    For Each k In best.Keys
        rw = CLng(k)
        If wsR.Cells(rw, col).HasFormula Then
            nFormula = nFormula + 1
        Else
            wsR.Cells(rw, col).Value = best(k)
            nW = nW + 1
        End If
    Next k

    Application.Calculate
    msg = "Hoan tat - cot ngay " & Format(Date, "dd-mmm") & IIf(added, " (them moi)", " (ghi de)") & vbCr & _
          "- File da doc: " & nFile & vbCr & _
          "- Dong report: " & mRead & " (dung duoc: " & mUsed & ")" & vbCr & _
          "- O da ghi: " & nW
    If nFormula > 0 Then msg = msg & vbCr & "- Bo qua (o dang la cong thuc): " & nFormula
    msg = msg & BuildNotes()
    MsgBox msg, vbInformation, "Import consumption"

Done:
    Application.Calculation = calcOld
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    Exit Sub
EH:
    MsgBox "Loi: " & Err.Description, vbCritical, "Import consumption"
    Resume Done
End Sub

'------------------------------------------------------------------------------
Private Function BuildNotes() As String
    Dim s As String, k As Variant
    If mSkipGhost > 0 Then s = s & vbCr & "- Bo qua Ghost Car <> Product: " & mSkipGhost
    If mSkipBad > 0 Then s = s & vbCr & "- Bo qua dong thieu color: " & mSkipBad
    If mSkipMap > 0 Then s = s & vbCr & "- Bo qua do chua map / SKIP: " & mSkipMap
    If mNewType.Count + mNewColor.Count > 0 Then
        s = s & vbCr & "- Co " & mNewType.Count & " Type, " & mNewColor.Count & _
            " Color moi -> da them vao sheet " & SH_MAP & " (o vang). Dien PART NAME / block roi chay lai."
    End If
    For Each k In mMiss.Keys
        s = s & vbCr & "- Sheet khong co dong: " & k
    Next k
    BuildNotes = s & mWarn
End Function

Private Sub InitState()
    Set mCnt = NewDict(): Set mNum = NewDict()
    Set mBlk = NewDict(): Set mPart = NewDict()
    Set mTypeMap = NewDict(): Set mColorMap = NewDict()
    Set mNewType = NewDict(): Set mNewColor = NewDict()
    Set mMiss = NewDict()
    mRead = 0: mUsed = 0
    mSkipGhost = 0: mSkipBad = 0: mSkipMap = 0
    mWarn = ""
End Sub

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

'chuan hoa: hoa, bo khoang trang thua, " " va "-" -> "_"
Private Function Norm(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Or IsNull(s) Or IsEmpty(s) Then Exit Function
    t = CStr(s)
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, vbTab, " ")
    t = UCase$(Trim$(t))
    t = Replace(t, " ", "_")
    t = Replace(t, "-", "_")
    Do While InStr(t, "__") > 0
        t = Replace(t, "__", "_")
    Loop
    Norm = t
End Function

'dong "XA" (xa) trong sheet: 2 ky tu, bat dau bang X
Private Function IsXa(ByVal pn As String) As Boolean
    IsXa = (Len(pn) = 2 And Left$(pn, 1) = "X")
End Function

'------------------------------------------------------------------------------
'  CHON FOLDER + LIST FILE
'------------------------------------------------------------------------------
Private Function PickFolder() As String
    Dim fd As Object, startPath As String, res As String
    startPath = GetSetting("ConsumptionImport", "Main", "Folder", "")
    Set fd = Application.FileDialog(4)   'msoFileDialogFolderPicker
    With fd
        .Title = "Chon folder chua 3 file report: primer, base, clear"
        If Len(startPath) > 0 Then .InitialFileName = startPath & "\"
        If .Show = -1 Then res = .SelectedItems(1)
    End With
    If Right$(res, 1) = "\" Then res = Left$(res, Len(res) - 1)
    If Len(res) > 0 Then SaveSetting "ConsumptionImport", "Main", "Folder", res
    PickFolder = res
End Function

Private Function ListReportFiles(ByVal folder As String) As Collection
    Dim col As New Collection
    Dim f As String, ln As String, ext As String, kind As String
    f = Dir(folder & "\*.*")
    Do While Len(f) > 0
        ln = LCase$(f)
        If Left$(ln, 2) <> "~$" And folder & "\" & f <> ThisWorkbook.FullName Then
            ext = Mid$(ln, InStrRev(ln, ".") + 1)
            If ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Or ext = "csv" Then
                kind = ""
                If InStr(ln, "primer") > 0 Then
                    kind = "PRIMER"
                ElseIf InStr(ln, "clear") > 0 Then
                    kind = "CLEAR"
                ElseIf InStr(ln, "base") > 0 Then
                    kind = "BASE"
                End If
                If Len(kind) > 0 Then col.Add f & "|" & kind
            End If
        End If
        f = Dir()
    Loop
    Set ListReportFiles = col
End Function

'------------------------------------------------------------------------------
'  DOC CAU TRUC SHEET ROBOT: block (cot B) va part (cot D)
'------------------------------------------------------------------------------
Private Sub BuildBlocks(ws As Worksheet)
    Dim lastRow As Long, r As Long, cur As String, pn As String
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, "B").Value))) > 0 Then
            cur = Norm(ws.Cells(r, "B").Value)
            mBlk(cur) = 1
        End If
        If Len(cur) > 0 And Len(Trim$(CStr(ws.Cells(r, "D").Value))) > 0 Then
            pn = Norm(ws.Cells(r, "D").Value)
            If Not IsXa(pn) Then mPart(cur & "|" & pn) = r
        End If
    Next r
End Sub

'------------------------------------------------------------------------------
'  SHEET MAP_TYPE
'------------------------------------------------------------------------------
Private Function GetMapSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_MAP)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_MAP
        ws.Range("A1").Value = "Type (trong report)"
        ws.Range("B1").Value = "PART NAME (cot D sheet ROBOT) hoac SKIP"
        ws.Range("D1").Value = "Color (trong report)"
        ws.Range("E1").Value = "Block mau (cot B sheet ROBOT) hoac SKIP"
        ws.Range("A1:E1").Font.Bold = True
        ws.Range("A2").Value = "VF2_R_BUMPER": ws.Range("B2").Value = "REAR_BUMPER"
        ws.Range("A3").Value = "VF2_F_BUMPER": ws.Range("B3").Value = "FRONT_BUMER"
        ws.Range("D2").Value = "SILVER": ws.Range("E2").Value = "SLIVER"
        ws.Columns("A:E").ColumnWidth = 38
        ws.Range("G1").Value = "Ghi chu: Type/Color trung ten sheet ROBOT thi khong can khai bao. Ghi SKIP de bo qua."
    End If
    Set GetMapSheet = ws
End Function

Private Function MapVal(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If s = "-" Then
        MapVal = "SKIP"
    Else
        MapVal = Norm(s)
    End If
End Function

Private Sub LoadMaps()
    Dim ws As Worksheet, r As Long, lr As Long, k As String
    Set ws = GetMapSheet()
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lr
        k = Norm(ws.Cells(r, "A").Value)
        If Len(k) > 0 Then mTypeMap(k) = MapVal(ws.Cells(r, "B").Value)
    Next r
    lr = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    For r = 2 To lr
        k = Norm(ws.Cells(r, "D").Value)
        If Len(k) > 0 Then mColorMap(k) = MapVal(ws.Cells(r, "E").Value)
    Next r
End Sub

Private Sub AppendNewMaps()
    Dim ws As Worksheet, k As Variant, r As Long
    If mNewType.Count + mNewColor.Count = 0 Then Exit Sub
    Set ws = GetMapSheet()
    For Each k In mNewType.Keys
        r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(r, "A").Value = mNewType(k)
        ws.Cells(r, "B").Interior.Color = RGB(255, 235, 156)
    Next k
    For Each k In mNewColor.Keys
        r = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row + 1
        ws.Cells(r, "D").Value = mNewColor(k)
        ws.Cells(r, "E").Interior.Color = RGB(255, 235, 156)
    Next k
End Sub

'------------------------------------------------------------------------------
'  DOC 1 FILE REPORT
'------------------------------------------------------------------------------
Private Sub ReadReport(ByVal fPath As String, ByVal kind As String)
    Dim wb As Workbook, ws As Worksheet, v As Variant
    Dim r As Long, c As Long, hdr As Long, i As Long
    Dim cType As Long, cColor As Long, cCons As Long, cGhost As Long, maxC As Long
    Dim lastRow As Long, cons As Double, colorTxt As String
    Dim skipRow As Boolean

    On Error Resume Next
    Set wb = Workbooks.Open(fPath, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo 0
    If wb Is Nothing Then
        mWarn = mWarn & vbCr & "- Khong mo duoc file: " & fPath
        Exit Sub
    End If

    On Error GoTo Fail
    Set ws = wb.Worksheets(1)

    'tim dong tieu de (chua "Consumption")
    For r = 1 To 15
        For c = 1 To 40
            If Norm(ws.Cells(r, c).Value) = "CONSUMPTION" Then hdr = r: Exit For
        Next c
        If hdr > 0 Then Exit For
    Next r
    If hdr = 0 Then
        mWarn = mWarn & vbCr & "- Khong thay tieu de 'Consumption' trong: " & wb.Name
        GoTo CloseIt
    End If

    For c = 1 To 60
        Select Case Norm(ws.Cells(hdr, c).Value)
            Case "TYPE": cType = c
            Case "COLOR": cColor = c
            Case "CONSUMPTION": cCons = c
            Case "GHOST_CAR": cGhost = c
        End Select
    Next c
    If cType = 0 Or cCons = 0 Or (kind = "BASE" And cColor = 0) Then
        mWarn = mWarn & vbCr & "- Thieu cot Type/Color/Consumption trong: " & wb.Name
        GoTo CloseIt
    End If

    maxC = cType
    If cType > maxC Then maxC = cType
    If cColor > maxC Then maxC = cColor
    If cCons > maxC Then maxC = cCons
    If cGhost > maxC Then maxC = cGhost

    lastRow = ws.Cells(ws.Rows.Count, cType).End(xlUp).Row
    If lastRow <= hdr Then GoTo CloseIt
    v = ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(lastRow, maxC)).Value2

    For i = 1 To UBound(v, 1)
        mRead = mRead + 1
        skipRow = False
        If ONLY_PRODUCT And cGhost > 0 Then
            If LCase$(Trim$(CStr(v(i, cGhost)))) <> "product" Then
                mSkipGhost = mSkipGhost + 1
                skipRow = True
            End If
        End If
        If Not skipRow Then
            cons = ToNum(v(i, cCons))
            If cons > 0 Then
                colorTxt = ""
                If cColor > 0 Then colorTxt = CStr(v(i, cColor))
                Accumulate kind, CStr(v(i, cType)), colorTxt, cons
            End If
        End If
    Next i

CloseIt:
    On Error Resume Next
    wb.Close SaveChanges:=False
    Exit Sub
Fail:
    mWarn = mWarn & vbCr & "- Loi khi doc " & wb.Name & ": " & Err.Description
    Resume CloseIt
End Sub

'xac dinh dong trong sheet ROBOT cho (kind, color, type) va cong don
Private Sub Accumulate(ByVal kind As String, ByVal typ As String, _
                       ByVal colorTxt As String, ByVal cons As Double)
    Dim blk As String, tn As String, cn As String, m As String
    Dim rw As Long, key As String

    '--- 1) block (PRIMER / CLEAR / mau)
    Select Case kind
        Case "PRIMER": blk = "PRIMER"
        Case "CLEAR": blk = "CLEAR"
        Case Else
            cn = Norm(colorTxt)
            If Len(cn) = 0 Then
                mSkipBad = mSkipBad + 1
                Exit Sub
            End If
            If mBlk.Exists(cn) Then
                blk = cn
            ElseIf mColorMap.Exists(cn) Then
                m = mColorMap(cn)
                If Len(m) = 0 Or m = "SKIP" Then
                    mSkipMap = mSkipMap + 1
                    Exit Sub
                End If
                If Not mBlk.Exists(m) Then
                    mMiss("block " & m) = 1
                    Exit Sub
                End If
                blk = m
            Else
                If Not mNewColor.Exists(cn) Then mNewColor.Add cn, Trim$(colorTxt)
                mSkipMap = mSkipMap + 1
                Exit Sub
            End If
    End Select
    If Not mBlk.Exists(blk) Then
        mMiss("block " & blk) = 1
        Exit Sub
    End If

    '--- 2) part
    tn = Norm(typ)
    If mPart.Exists(blk & "|" & tn) Then
        rw = mPart(blk & "|" & tn)
    ElseIf mTypeMap.Exists(tn) Then
        m = mTypeMap(tn)
        If Len(m) = 0 Or m = "SKIP" Then
            mSkipMap = mSkipMap + 1
            Exit Sub
        End If
        If mPart.Exists(blk & "|" & m) Then
            rw = mPart(blk & "|" & m)
        Else
            mMiss(blk & " / " & m) = 1
            Exit Sub
        End If
    Else
        If Not mNewType.Exists(tn) Then mNewType.Add tn, Trim$(typ)
        mSkipMap = mSkipMap + 1
        Exit Sub
    End If

    key = CStr(rw) & "|" & CStr(cons)
    If mCnt.Exists(key) Then
        mCnt(key) = mCnt(key) + 1
    Else
        mCnt.Add key, 1
        mNum.Add key, cons
    End If
    mUsed = mUsed + 1
End Sub

Private Function ToNum(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        ToNum = Val(Replace(Trim$(CStr(v)), ",", ""))
    ElseIf IsNumeric(v) Then
        ToNum = CDbl(v)
    End If
End Function

'------------------------------------------------------------------------------
'  COT NGAY: tim cot hom nay, them moi, lien ket 3 sheet
'------------------------------------------------------------------------------
'cac cot ngay = tu cot F, hang 13 co gia tri so/ngay lien tiep; dung o o text/trong
Private Function GetTodayCol(ws As Worksheet, ByRef added As Boolean) As Long
    Dim c As Long, lastCol As Long, v As Variant, td As Double, newCol As Long
    td = CDbl(Date)
    lastCol = FIRST_DAY_COL - 1
    For c = FIRST_DAY_COL To FIRST_DAY_COL + 400
        v = ws.Cells(ROW_DAY, c).Value2
        If IsError(v) Or IsEmpty(v) Then Exit For
        If VarType(v) = vbString Then Exit For
        lastCol = c
        If CDbl(v) = td Then
            GetTodayCol = c
            Exit Function
        End If
    Next c

    newCol = lastCol + 1
    If lastCol >= FIRST_DAY_COL Then AddDayColumn newCol   'chua co cot nao -> dung cot F san co
    SetDayHeader newCol
    added = True
    GetTodayCol = newCol
End Function

'chen 1 cot moi tai newCol o ca 3 sheet, copy cong thuc tu cot ngay truoc
Private Sub AddDayColumn(ByVal newCol As Long)
    Dim nm As Variant, ws As Worksheet, prevCol As Long, lastRow As Long
    prevCol = newCol - 1
    For Each nm In Array(SH_ROBOT, SH_TOTAL, SH_PART)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

        ws.Columns(newCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Columns(newCol).ColumnWidth = ws.Columns(prevCol).ColumnWidth
        ExtendMerges ws, prevCol, newCol

        ws.Range(ws.Cells(ROW_DAY, prevCol), ws.Cells(lastRow, prevCol)).Copy ws.Cells(ROW_DAY, newCol)

        'bo so lieu nhap tay bi copy sang (chi giu cong thuc)
        On Error Resume Next
        ws.Range(ws.Cells(FIRST_DATA_ROW, newCol), ws.Cells(lastRow, newCol)).SpecialCells(xlCellTypeConstants).ClearContents
        On Error GoTo 0
    Next nm
    Application.CutCopyMode = False
End Sub

'hang 13: ROBOT = ngay chay; 2 sheet kia = lien ket ve ROBOT. Hang 2 = thu
Private Sub SetDayHeader(ByVal col As Long)
    Dim nm As Variant, ws As Worksheet, wsR As Worksheet, a As String
    Set wsR = ThisWorkbook.Worksheets(SH_ROBOT)
    For Each nm In Array(SH_ROBOT, SH_TOTAL, SH_PART)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        If ws.Name = SH_ROBOT Then
            ws.Cells(ROW_DAY, col).Value = CLng(Date)
        Else
            ws.Cells(ROW_DAY, col).Formula = "=" & SH_ROBOT & "!" & wsR.Cells(ROW_DAY, col).Address(True, False)
        End If
        ws.Cells(ROW_DAY, col).NumberFormat = "dd-mmm"
        a = ws.Cells(ROW_DAY, col).Address(True, False)
        ws.Cells(ROW_WEEKDAY, col).Formula = "=IF(" & a & "="""","""",CHOOSE(WEEKDAY(" & a & ",2)," & _
            """T2"",""T3"",""T4"",""T5"",""T6"",""T7"",""CN""))"
    Next nm
End Sub

'o tieu de (hang 8-13) dang merge den cot ngay cuoi -> keo dai them 1 cot
Private Sub ExtendMerges(ws As Worksheet, ByVal prevCol As Long, ByVal newCol As Long)
    Dim r As Long, ma As Range, rg As Range
    For r = 8 To ROW_DAY
        If ws.Cells(r, prevCol).MergeCells Then
            Set ma = ws.Cells(r, prevCol).MergeArea
            If ma.Column + ma.Columns.Count - 1 = prevCol Then
                Set rg = ws.Range(ma.Cells(1, 1), ws.Cells(ma.Row + ma.Rows.Count - 1, newCol))
                Application.DisplayAlerts = False
                ma.UnMerge
                rg.Merge
                Application.DisplayAlerts = True
            End If
        End If
    Next r
End Sub
