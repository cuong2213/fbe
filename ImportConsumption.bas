Option Explicit

'==============================================================================
'  IMPORT TIEU HAO SON TU REPORT CONSUMPTION (PRIMER / BASE / CLEAR)
'  Chay macro: ImportConsumption   (dat trong chinh file Dinh_muc_son)
'
'  - Chon folder -> tu tim file co ten chua primer / base / clear (khong phan biet hoa thuong)
'  - ROBOT_TIEU_HAO_1SKID: tu tao cot C "COLOR" (neu chua co), merge giong cot B
'  - PRIMER / CLEAR : Type -> PART NAME, ghi vao dong co COLOR = PRIMER / CLEAR
'  - BASE           : them cot Color -> ghi vao dong co COLOR tuong ung + PART NAME
'  - Cot ngay = NGAY CHAY MACRO (hang 13 = ngay, dd-mmm). Dung cot trong san co,
'    het cot trong thi tu chen them cot o: ROBOT_TIEU_HAO_1SKID, TOTAL_SKID,
'    ROBOT_TIEU_HAO_1PART, TONG_HOP, TONG_HOP_Detail (ca khoi mml va khoi g),
'    copy cong thuc tu cot truoc -> tu tro dung cot cung ngay
'  - Gia tri ghi = consumption XUAT HIEN NHIEU NHAT (mode) theo tung dong part
'    (bang nhau -> lay gia tri xuat hien truoc trong report)
'  - MAP_TYPE: 1 Type co the co nhieu dong (nhieu PART NAME), 1 PART NAME co the
'    nam o nhieu Type. Type -> nhieu part: moi part nhan cung gia tri.
'    Nhieu Type -> 1 part: cac gia tri duoc gop chung roi lay mode.
'==============================================================================

Private Const SH_ROBOT As String = "ROBOT_TIEU_HAO_1SKID"
Private Const SH_TOTAL As String = "TOTAL_SKID"
Private Const SH_PART As String = "ROBOT_TIEU_HAO_1PART"
Private Const SH_TH As String = "TONG_HOP"
Private Const SH_DET As String = "TONG_HOP_Detail"
Private Const SH_MAP As String = "MAP_TYPE"

Private Const ROW_WEEKDAY As Long = 2          'hang "Thu" (3 sheet ROBOT/TOTAL/PART)
Private Const ROW_DAY As Long = 13             'hang ngay (gia tri la ngay that)
Private Const ROW_HDR As Long = 11             'hang tieu de cot (PART NAME, COLOR...)
Private Const ROW_LBL As Long = 4              'hang nhan ngay cua TONG_HOP / TONG_HOP_Detail
Private Const FIRST_DATA_ROW As Long = 14

Private Const ONLY_PRODUCT As Boolean = True   'True: chi lay dong Ghost Car = Product

Private mCnt As Object          'key "dong|gia tri" -> so lan xuat hien
Private mNum As Object          'key "dong|gia tri" -> gia tri (so)
Private mBlk As Object          'COLOR (chuan hoa) -> 1
Private mPart As Object         'COLOR|part (chuan hoa) -> so dong trong ROBOT
Private mTypeMap As Object      'MAP_TYPE: Type -> "PART1|PART2|..." (rong = chua dien)
Private mColorMap As Object     'MAP_TYPE: Color -> block
Private mNewType As Object      'Type chua co trong MAP_TYPE
Private mNewColor As Object     'Color chua co trong MAP_TYPE
Private mMiss As Object         'COLOR / part khong co dong trong ROBOT
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

    EnsureColorColumn wsR
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

    'cot ngay hom nay (dung cot trong, het thi tu chen them o tat ca cac sheet)
    col = PrepareTodaySlot(added)
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
    msg = "Hoan tat - cot ngay " & Format(Date, "dd-mmm") & IIf(added, " (moi)", " (ghi de)") & vbCr & _
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

Private Function CellText(ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Dim v As Variant
    v = ws.Cells(r, c).Value
    If IsError(v) Or IsEmpty(v) Then Exit Function
    CellText = CStr(v)
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    LastUsedRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Function

'------------------------------------------------------------------------------
'  CHON FOLDER + LIST FILE (khong phan biet hoa thuong)
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
    Dim f As String, ext As String, kind As String
    f = Dir(folder & "\*.*")
    Do While Len(f) > 0
        If Left$(f, 2) <> "~$" And StrComp(folder & "\" & f, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            ext = LCase$(Mid$(f, InStrRev(f, ".") + 1))
            If ext = "xlsx" Or ext = "xls" Or ext = "xlsm" Or ext = "csv" Then
                kind = ""
                If InStr(1, f, "primer", vbTextCompare) > 0 Then
                    kind = "PRIMER"
                ElseIf InStr(1, f, "clear", vbTextCompare) > 0 Then
                    kind = "CLEAR"
                ElseIf InStr(1, f, "base", vbTextCompare) > 0 Then
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
'  CAU TRUC SHEET ROBOT: cot COLOR (C), cot PART NAME, cot ngay dau tien
'------------------------------------------------------------------------------
Private Function HeaderCol(ws As Worksheet, ByVal key As String, ByVal r As Long, _
                           ByVal c1 As Long, ByVal c2 As Long) As Long
    Dim c As Long
    For c = c1 To c2
        If Norm(ws.Cells(r, c).Value) = key Then
            HeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Function PartCol(ws As Worksheet) As Long
    PartCol = HeaderCol(ws, "PART_NAME", ROW_HDR, 1, 12)
    If PartCol = 0 Then Err.Raise vbObjectError + 2, , "Khong thay tieu de PART NAME o hang " & ROW_HDR & " cua sheet " & ws.Name
End Function

'cot ngay dau tien = PART NAME + 2 (sau cot SO LUONG PART)
Private Function FirstDayCol(ws As Worksheet) As Long
    FirstDayCol = PartCol(ws) + 2
End Function

'cot dau tien cua nhom MANUAL (hang 12) = ngay sau cot ngay cuoi cung (ROBOT)
Private Function FindManualCol(ws As Worksheet) As Long
    Dim c As Long
    For c = FirstDayCol(ws) To FirstDayCol(ws) + 100
        If UCase$(Trim$(CellText(ws, 12, c))) = "MANUAL" Then
            FindManualCol = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 1, , "Khong tim thay chu MANUAL o hang 12 cua sheet " & ws.Name
End Function

'tao cot C "COLOR" neu chua co: cung cau truc (merge, dinh dang) voi cot B ben canh;
'ten color chi nam o o tren cung cua moi vung merge (giong cot B)
Private Sub EnsureColorColumn(ws As Worksheet)
    Dim cCol As Long, pCol As Long, r As Long, lastRow As Long

    PartCol ws   'kiem tra cau truc truoc
    cCol = HeaderCol(ws, "COLOR", ROW_HDR, 1, 12)
    If cCol = 0 Then
        ws.Columns(3).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        ws.Columns(3).ColumnWidth = ws.Columns(2).ColumnWidth
        lastRow = LastUsedRow(ws)
        'copy dinh dang + merge cua cot B (tu hang tieu de den het du lieu) sang cot C
        ws.Range(ws.Cells(ROW_HDR, 2), ws.Cells(lastRow, 2)).Copy
        ws.Cells(ROW_HDR, 3).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        ws.Cells(ROW_HDR, 3).Value = "COLOR"
        cCol = 3
    End If

    'dien ten color vao o tren cung cua tung vung (lay tu cot B), khong ghi de o da co
    lastRow = LastUsedRow(ws)
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(CellText(ws, r, 2))) > 0 Then
            If Len(Trim$(CellText(ws, r, cCol))) = 0 Then ws.Cells(r, cCol).Value = CellText(ws, r, 2)
        End If
    Next r
End Sub

Private Sub BuildBlocks(ws As Worksheet)
    Dim lastRow As Long, r As Long, cur As String, pn As String
    Dim cCol As Long, pCol As Long
    cCol = HeaderCol(ws, "COLOR", ROW_HDR, 1, 12)
    pCol = PartCol(ws)
    If cCol = 0 Then Err.Raise vbObjectError + 3, , "Khong thay cot COLOR trong sheet " & ws.Name
    lastRow = LastUsedRow(ws)
    For r = FIRST_DATA_ROW To lastRow
        If Len(Trim$(CellText(ws, r, cCol))) > 0 Then
            cur = Norm(ws.Cells(r, cCol).Value)
            mBlk(cur) = 1
        End If
        If Len(cur) > 0 And Len(Trim$(CellText(ws, r, pCol))) > 0 Then
            pn = Norm(ws.Cells(r, pCol).Value)
            If Not IsXa(pn) Then mPart(cur & "|" & pn) = r
        End If
    Next r
End Sub

'------------------------------------------------------------------------------
'  SHEET MAP_TYPE  (1 Type nhieu dong; 1 PART NAME nam o nhieu Type)
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
        ws.Range("B1").Value = "PART NAME (cot PART NAME sheet ROBOT) hoac SKIP"
        ws.Range("D1").Value = "Color (trong report)"
        ws.Range("E1").Value = "COLOR (cot C sheet ROBOT) hoac SKIP"
        ws.Range("A1:E1").Font.Bold = True
        ws.Range("A2").Value = "VF2_R_BUMPER": ws.Range("B2").Value = "REAR_BUMPER"
        ws.Range("A3").Value = "VF2_F_BUMPER": ws.Range("B3").Value = "FRONT_BUMER"
        ws.Range("D2").Value = "SILVER": ws.Range("E2").Value = "SLIVER"
        ws.Columns("A:E").ColumnWidth = 38
        ws.Range("G1").Value = "Ghi chu:"
        ws.Range("G2").Value = "- Type/Color trung ten sheet ROBOT thi khong can khai bao. Ghi SKIP de bo qua."
        ws.Range("G3").Value = "- 1 Type -> nhieu part: viet nhieu dong cung Type, moi dong 1 PART NAME (moi part nhan cung gia tri)."
        ws.Range("G4").Value = "- Nhieu Type -> 1 part: viet nhieu dong, cung PART NAME (gia tri cua cac Type duoc gop de lay mode)."
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
    Dim ws As Worksheet, r As Long, lr As Long, k As String, v As String
    Set ws = GetMapSheet()
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lr
        k = Norm(ws.Cells(r, "A").Value)
        If Len(k) > 0 Then
            If Not mTypeMap.Exists(k) Then mTypeMap(k) = ""
            v = MapVal(ws.Cells(r, "B").Value)
            If Len(v) > 0 Then
                If Len(mTypeMap(k)) > 0 Then mTypeMap(k) = mTypeMap(k) & "|"
                mTypeMap(k) = mTypeMap(k) & v
            End If
        End If
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

'xac dinh cac dong trong sheet ROBOT cho (kind, color, type) va cong don
Private Sub Accumulate(ByVal kind As String, ByVal typ As String, _
                       ByVal colorTxt As String, ByVal cons As Double)
    Dim blk As String, tn As String, cn As String, m As String
    Dim rw As Long, arr() As String, j As Long, n As Long, done As String

    '--- 1) COLOR (PRIMER / CLEAR / mau)
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
                    mMiss("COLOR " & m) = 1
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
        mMiss("COLOR " & blk) = 1
        Exit Sub
    End If

    '--- 2) part: trung ten truc tiep + cac dong MAP_TYPE (co the nhieu part)
    tn = Norm(typ)
    If mPart.Exists(blk & "|" & tn) Then
        rw = mPart(blk & "|" & tn)
        AddVal rw, cons
        done = "|" & rw & "|"
        n = n + 1
    End If
    If mTypeMap.Exists(tn) Then
        If Len(mTypeMap(tn)) > 0 Then
            arr = Split(mTypeMap(tn), "|")
            For j = 0 To UBound(arr)
                m = arr(j)
                If m <> "SKIP" Then
                    If mPart.Exists(blk & "|" & m) Then
                        rw = mPart(blk & "|" & m)
                        If InStr(done, "|" & rw & "|") = 0 Then
                            AddVal rw, cons
                            done = done & rw & "|"
                            n = n + 1
                        End If
                    Else
                        mMiss(blk & " / " & m) = 1
                    End If
                End If
            Next j
        End If
    ElseIf n = 0 Then
        If Not mNewType.Exists(tn) Then mNewType.Add tn, Trim$(typ)
    End If

    If n > 0 Then
        mUsed = mUsed + 1
    Else
        mSkipMap = mSkipMap + 1
    End If
End Sub

Private Sub AddVal(ByVal rw As Long, ByVal cons As Double)
    Dim key As String
    key = CStr(rw) & "|" & CStr(cons)
    If mCnt.Exists(key) Then
        mCnt(key) = mCnt(key) + 1
    Else
        mCnt.Add key, 1
        mNum.Add key, cons
    End If
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
'  COT NGAY
'  "Cot" = o co san trong sheet. Hang 13 la ngay that (>40000) -> da dung,
'  con lai (so 1..31, rong, "") -> con trong.
'------------------------------------------------------------------------------
Private Function IsDateSerial(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If IsNumeric(v) Then IsDateSerial = (CDbl(v) > 40000)
End Function

'tra ve cot ngay hom nay trong sheet ROBOT; them cot moi o moi sheet neu het cot trong
Private Function PrepareTodaySlot(ByRef added As Boolean) As Long
    Dim wsR As Worksheet, first As Long, nSlots As Long, i As Long
    Dim slot As Long, freeIdx As Long, found As Boolean, v As Variant

    Set wsR = ThisWorkbook.Worksheets(SH_ROBOT)
    first = FirstDayCol(wsR)
    nSlots = FindManualCol(wsR) - first
    freeIdx = -1

    For i = 0 To nSlots - 1
        v = wsR.Cells(ROW_DAY, first + i).Value2
        If IsDateSerial(v) Then
            If CDbl(v) = CDbl(Date) Then
                slot = i
                found = True
                Exit For
            End If
        ElseIf freeIdx < 0 Then
            freeIdx = i
        End If
    Next i

    If found Then
        added = False
        PrepareTodaySlot = first + slot
        Exit Function
    End If

    added = True
    If freeIdx >= 0 Then
        slot = freeIdx
    Else
        slot = nSlots
        InsertDaySlot nSlots
        nSlots = nSlots + 1
    End If

    SetSlotHeaders slot
    ClearFreeHeaders slot, nSlots
    PrepareTodaySlot = first + slot
End Function

'het cot trong: chen them 1 cot o ROBOT/TOTAL/PART va o khoi mml + khoi g cua TONG_HOP, TONG_HOP_Detail
Private Sub InsertDaySlot(ByVal nSlots As Long)
    Dim nm As Variant, ws As Worksheet, mmlS As Long, gS As Long, gE As Long

    For Each nm In Array(SH_ROBOT, SH_TOTAL, SH_PART)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        InsertColCopy ws, FirstDayCol(ws) + nSlots, ROW_DAY
    Next nm

    For Each nm In Array(SH_TH, SH_DET)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        GetBlocks ws, mmlS, gS, gE
        InsertColCopy ws, gS, ROW_LBL           'khoi mml: chen sau cot mml cuoi
        GetBlocks ws, mmlS, gS, gE
        InsertColCopy ws, gE + 1, ROW_LBL       'khoi g: chen sau cot g cuoi
    Next nm
End Sub

'khoi mml / g dua vao hang 3: "(mml)" ... "(g)" ... o co chu tiep theo (Ty le phun lai)
Private Sub GetBlocks(ws As Worksheet, ByRef mmlS As Long, ByRef gS As Long, ByRef gE As Long)
    Dim c As Long, t As String, lastC As Long
    mmlS = 0: gS = 0: gE = 0
    lastC = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    For c = 1 To lastC
        t = CellText(ws, 3, c)
        If mmlS = 0 Then
            If InStr(1, t, "(mml)", vbTextCompare) > 0 Then mmlS = c
        ElseIf gS = 0 Then
            If InStr(1, t, "(g)", vbTextCompare) > 0 Then gS = c
        ElseIf Len(Trim$(t)) > 0 Then
            gE = c - 1
            Exit For
        End If
    Next c
    If mmlS = 0 Or gS = 0 Or gE = 0 Then
        Err.Raise vbObjectError + 4, , "Khong xac dinh duoc khoi mml/g (hang 3) trong sheet " & ws.Name
    End If
End Sub

'chen 1 cot tai newCol, copy cong thuc/dinh dang tu cot truoc (tu hang hdrRow tro xuong)
Private Sub InsertColCopy(ws As Worksheet, ByVal newCol As Long, ByVal hdrRow As Long)
    Dim prevCol As Long, lastRow As Long
    prevCol = newCol - 1
    lastRow = LastUsedRow(ws)

    ws.Columns(newCol).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Columns(newCol).ColumnWidth = ws.Columns(prevCol).ColumnWidth
    ExtendMerges ws, prevCol, newCol, hdrRow

    ws.Range(ws.Cells(hdrRow, prevCol), ws.Cells(lastRow, prevCol)).Copy ws.Cells(hdrRow, newCol)

    'bo so lieu nhap tay bi copy sang (chi giu cong thuc)
    On Error Resume Next
    ws.Range(ws.Cells(hdrRow + 1, newCol), ws.Cells(lastRow, newCol)).SpecialCells(xlCellTypeConstants).ClearContents
    On Error GoTo 0
    Application.CutCopyMode = False
End Sub

'o tieu de (hang 1..hdrRow) dang merge den cot truoc -> keo dai them 1 cot
Private Sub ExtendMerges(ws As Worksheet, ByVal prevCol As Long, ByVal newCol As Long, ByVal hdrRow As Long)
    Dim r As Long, ma As Range, rg As Range
    For r = 1 To hdrRow
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

'dat nhan ngay cho cot thu "slot" (0-based) o tat ca cac sheet
Private Sub SetSlotHeaders(ByVal slot As Long)
    Dim nm As Variant, ws As Worksheet, wsR As Worksheet, c As Long, a As String
    Dim rAddr As String, mmlS As Long, gS As Long, gE As Long

    Set wsR = ThisWorkbook.Worksheets(SH_ROBOT)
    rAddr = SH_ROBOT & "!" & wsR.Cells(ROW_DAY, FirstDayCol(wsR) + slot).Address(True, False)

    For Each nm In Array(SH_ROBOT, SH_TOTAL, SH_PART)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        c = FirstDayCol(ws) + slot
        If ws.Name = SH_ROBOT Then
            ws.Cells(ROW_DAY, c).Value = CLng(Date)
        Else
            ws.Cells(ROW_DAY, c).Formula = "=" & rAddr
        End If
        ws.Cells(ROW_DAY, c).NumberFormat = "dd-mmm"
        a = ws.Cells(ROW_DAY, c).Address(True, False)
        ws.Cells(ROW_WEEKDAY, c).Formula = "=IF(" & a & "="""","""",CHOOSE(WEEKDAY(" & a & ",2)," & _
            """T2"",""T3"",""T4"",""T5"",""T6"",""T7"",""CN""))"
    Next nm

    For Each nm In Array(SH_TH, SH_DET)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        GetBlocks ws, mmlS, gS, gE
        ws.Cells(ROW_LBL, mmlS + slot).Formula = "=" & rAddr
        ws.Cells(ROW_LBL, mmlS + slot).NumberFormat = "dd-mmm"
        ws.Cells(ROW_LBL, gS + slot).Formula = "=" & rAddr
        ws.Cells(ROW_LBL, gS + slot).NumberFormat = "dd-mmm"
    Next nm
End Sub

'xoa nhan cua cac cot con trong phia sau (tranh con so 1..30 cu / chuoi cong thuc cu)
Private Sub ClearFreeHeaders(ByVal slot As Long, ByVal nSlots As Long)
    Dim nm As Variant, ws As Worksheet, i As Long, c As Long
    Dim mmlS As Long, gS As Long, gE As Long

    For Each nm In Array(SH_ROBOT, SH_TOTAL, SH_PART)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        For i = slot + 1 To nSlots - 1
            c = FirstDayCol(ws) + i
            ws.Cells(ROW_DAY, c).ClearContents
            ws.Cells(ROW_WEEKDAY, c).ClearContents
        Next i
    Next nm

    For Each nm In Array(SH_TH, SH_DET)
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        GetBlocks ws, mmlS, gS, gE
        For i = slot + 1 To gS - mmlS - 1
            ws.Cells(ROW_LBL, mmlS + i).ClearContents
        Next i
        For i = slot + 1 To gE - gS
            ws.Cells(ROW_LBL, gS + i).ClearContents
        Next i
    Next nm
End Sub
