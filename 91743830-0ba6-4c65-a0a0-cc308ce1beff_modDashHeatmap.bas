Attribute VB_Name = "modDashHeatmap"
Option Explicit
'=====================================================================
' VEHICLE DEFECT HEATMAP (Stage 3)  -  ban viet lai theo modVehicleMapping
' - Tinh Top 3 vi tri (o Grid) cho tung Defect Type, TUNG MODEL, truc tiep tu tblDefects
'   (Model, X = cot, Y = hang, DefectType, Qty) theo bo loc cua Dashboard.
' - Hinh xe lay tu hinh Master (tblModels > ImageShapeName) tren MAP_VEHICLE.
'   Vi tri o (r, c) = ((c - 0.5) / Cols , (r - 0.5) / Rows) cua hinh (Grid chia deu).
' - Ve marker 1/2/3 mau theo Defect Type (Config). Re chuot len marker de xem Cell ID, so luong.
' - Bang Top 3 cu tren Dashboard bi AN (hang 60 tro xuong), khong con can cho heatmap.
' Phu thuoc (Public o module khac):
'   modVehicleMapping: SH_MAP, SH_MASTER, SH_DB, TBL_MODELS, TBL_DEFECTS
'   Dashboard: DashSheet, SH_DASH, SH_CFG, TBL_DCOL, TBL_DSET, ROW_T3_FIRST, T3_MAXROWS, ROW_DT_HDR
' Bo loc doc truc tiep tu o: A5 From, C5 To, E5 Model, G5 Shift, I5 Defect Type, K5 Cell ID
'=====================================================================

Private Const HM_FIRST_ROW As Long = 94
Private Const HM_MAX_TYPES As Long = 10
Private Const HM_TITLE_H As Double = 16
Private Const HM_GAP As Double = 8

Public Sub UpdateHeatmap(ByVal showMsg As Boolean)
    Dim ws As Worksheet, mdl As Object, grp As Object, totDef As Object, totMod As Object
    Dim defs() As String, nD As Long, mods() As String, nM As Long, k As Variant
    Dim availW As Double, availH As Double, legendH As Double, hmAvail As Double
    Dim pics() As Shape, asp() As Double, perRow As Long, nTileRows As Long
    Dim i As Long, tr As Long, tc As Long, tileW As Double, sumAsp As Double, mx As Double
    Dim yRow() As Double, x As Double, y As Double, nSkip As Long, prevSU As Boolean
    Dim info As Variant, s As Shape, pic As Shape, dd As Long, rr As Long
    Dim t As Variant, key As String, occ As Object, n As Long, base As Double, sz As Double
    Dim mx0 As Double, my0 As Double, rowN As Long, colN As Long, v As Double, clr As Long, tip As String

    On Error GoTo EH
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = DashSheet(SH_DASH)

    ws.Rows("60:" & (ROW_T3_FIRST + T3_MAXROWS)).Hidden = True        'an bang Top 3 cu

    availW = ws.Range("O1").Left - ws.Range("A1").Left - 8
    availH = ws.Rows(ROW_DT_HDR - 1).Top - ws.Rows(HM_FIRST_ROW).Top - 6
    If availH < 120 Then
        If showMsg Then MsgBox "Vung Vehicle Heatmap qua thap (hang " & HM_FIRST_ROW & " den " & (ROW_DT_HDR - 2) & "). Hay tang ROW_DT_HDR.", vbExclamation, "Dashboard"
        GoTo CLEANUP
    End If

    ClearHeatShapes ws
    ws.Range("A93").Value = "So 1-2-3 = thu hang vi tri loi (o Grid) cua tung Defect Type, theo tung Model. Mau = Defect Type (Config). Re chuot len so de xem Cell ID."

    Set mdl = ReadModels()
    Aggregate ws, mdl, grp, totDef, totMod, nSkip
    If totMod.Count = 0 Then
        DeleteStalePics ws, totMod
        ws.Range("A93").Value = "Khong co du lieu phu hop bo loc (hoac Model khong co trong tblModels)."
        GoTo CLEANUP
    End If
    DeleteStalePics ws, totMod

    'Defect Type theo tong giam dan (toi da HM_MAX_TYPES)
    defs = SortedKeys(totDef, HM_MAX_TYPES)
    nD = UBound(defs)

    'Danh sach Model theo thu tu tblModels
    nM = 0
    ReDim mods(1 To totMod.Count)
    For Each k In mdl.Keys
        If totMod.Exists(CStr(k)) Then
            nM = nM + 1
            mods(nM) = CStr(k)
        End If
    Next k

    'Hinh nen cho tung Model
    ReDim pics(1 To nM)
    ReDim asp(1 To nM)
    For i = 1 To nM
        info = mdl(mods(i))
        Set pics(i) = ModelPicture(ws, mods(i), CStr(info(2)), showMsg)
        If pics(i) Is Nothing Then
            If showMsg Then MsgBox "Khong lay duoc hinh Master '" & CStr(info(2)) & "' cua Model " & mods(i) & ".", vbExclamation, "Dashboard"
            asp(i) = 0.5
        Else
            asp(i) = pics(i).Height / pics(i).Width
        End If
    Next i

    'Bo cuc: legend, roi cac o (tile) 1 hoac 2 cot
    legendH = DrawLegend(ws, defs, nD, availW)
    hmAvail = availH - legendH - 4
    If nM = 1 Then perRow = 1 Else perRow = 2
    nTileRows = (nM + perRow - 1) \ perRow
    tileW = (availW - HM_GAP * (perRow - 1)) / perRow
    ReDim yRow(0 To nTileRows - 1)
    For tr = 0 To nTileRows - 1
        mx = 0
        For tc = 0 To perRow - 1
            i = tr * perRow + tc + 1
            If i <= nM Then
                If asp(i) > mx Then mx = asp(i)
            End If
        Next tc
        sumAsp = sumAsp + mx
    Next tr
    If tileW * sumAsp + nTileRows * HM_TITLE_H + HM_GAP * (nTileRows - 1) > hmAvail Then
        tileW = (hmAvail - nTileRows * HM_TITLE_H - HM_GAP * (nTileRows - 1)) / sumAsp
    End If
    If tileW < 120 Then
        If showMsg Then MsgBox "Vung Vehicle Heatmap qua thap cho " & nM & " Model. Hay tang ROW_DT_HDR.", vbExclamation, "Dashboard"
        GoTo CLEANUP
    End If

    y = ws.Rows(HM_FIRST_ROW).Top + 2 + legendH + 4
    For tr = 0 To nTileRows - 1
        yRow(tr) = y
        mx = 0
        For tc = 0 To perRow - 1
            i = tr * perRow + tc + 1
            If i <= nM Then
                If asp(i) > mx Then mx = asp(i)
            End If
        Next tc
        y = y + HM_TITLE_H + tileW * mx + HM_GAP
    Next tr

    base = tileW / 45
    If base < 11 Then base = 11
    If base > 20 Then base = 20
    Set occ = CreateObject("Scripting.Dictionary")

    For i = 1 To nM
        tr = (i - 1) \ perRow
        tc = (i - 1) Mod perRow
        x = ws.Range("A1").Left + 2 + tc * (tileW + HM_GAP)
        info = mdl(mods(i))

        'tieu de o
        Set s = ws.Shapes.AddShape(msoShapeRectangle, x, yRow(tr), tileW, HM_TITLE_H)
        With s
            .Name = "HM_T_" & mods(i)
            .Fill.ForeColor.RGB = RGB(31, 56, 100)
            .Line.Visible = msoFalse
            .Placement = xlMove
            With .TextFrame2
                .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
                .VerticalAnchor = msoAnchorMiddle
                .WordWrap = msoFalse
                .TextRange.Text = mods(i) & "  (" & info(0) & " x " & info(1) & ")  -  " & Format$(totMod(mods(i)), "#,##0")
                .TextRange.Font.Size = 9
                .TextRange.Font.Bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            End With
        End With

        'hinh nen
        Set pic = pics(i)
        If Not pic Is Nothing Then
            With pic
                .LockAspectRatio = msoFalse
                .Left = x
                .Top = yRow(tr) + HM_TITLE_H
                .Width = tileW
                .Height = tileW * asp(i)
                .Line.Visible = msoTrue
                .Line.ForeColor.RGB = RGB(150, 150, 150)
                .Line.Weight = 0.75
            End With
            mx0 = pic.Left: my0 = pic.Top

            'marker Top 3 cho tung Defect Type
            For dd = 1 To nD
                key = mods(i) & "|" & defs(dd)
                If grp.Exists(key) Then
                    t = grp(key)
                    clr = HmColor(defs(dd), dd)
                    For rr = 1 To 3
                        v = t(rr, 1)
                        If v > 0 Then
                            rowN = CLng(t(rr, 2)): colN = CLng(t(rr, 3))
                            sz = base * (1 - 0.15 * (rr - 1))
                            n = 0
                            If occ.Exists(i & "|" & rowN & "," & colN) Then n = occ(i & "|" & rowN & "," & colN)
                            occ(i & "|" & rowN & "," & colN) = n + 1
                            Set s = ws.Shapes.AddShape(msoShapeOval, _
                                    mx0 + (colN - 0.5) / CLng(info(1)) * pic.Width + n * sz * 0.45 - sz / 2, _
                                    my0 + (rowN - 0.5) / CLng(info(0)) * pic.Height - sz / 2, sz, sz)
                            tip = mods(i) & " - " & defs(dd) & " - hang " & rr & " - " & CellLetter(colN) & rowN & " (" & CStr(CLng(v)) & ")"
                            StyleMarker s, "HM_M_" & i & "_" & dd & "_" & rr, CStr(rr), clr, tip, sz
                        End If
                    Next rr
                End If
            Next dd
        End If
    Next i

    If nSkip > 0 Then
        ws.Range("A93").Value = ws.Range("A93").Value & "  (Bo qua " & nSkip & " ban ghi ngoai Grid)"
    End If
CLEANUP:
    Application.CutCopyMode = False
    Application.ScreenUpdating = prevSU
    Exit Sub
EH:
    If showMsg Then MsgBox "Heatmap loi: " & Err.Description, vbExclamation, "Dashboard"
    Resume CLEANUP
End Sub

'=====================================================================
' DU LIEU
'=====================================================================
'Model -> Array(Rows, Cols, ImageShapeName), thu tu theo tblModels
Private Function ReadModels() As Object
    Dim d As Object, lo As ListObject, dat As Variant, i As Long, nm As String
    Dim cM As Long, cR As Long, cC As Long, cI As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set lo = ThisWorkbook.Worksheets(SH_MASTER).ListObjects(TBL_MODELS)
    If Not lo.DataBodyRange Is Nothing Then
        cM = lo.ListColumns("Model").Index
        cR = lo.ListColumns("Rows").Index
        cC = lo.ListColumns("Cols").Index
        cI = lo.ListColumns("ImageShapeName").Index
        dat = lo.DataBodyRange.Value2
        For i = 1 To UBound(dat, 1)
            nm = Trim$(CStr(dat(i, cM)))
            If Len(nm) > 0 And IsNumeric(dat(i, cR)) And IsNumeric(dat(i, cC)) Then
                If CLng(dat(i, cR)) >= 1 And CLng(dat(i, cC)) >= 1 And Not d.Exists(nm) Then
                    d.Add nm, Array(CLng(dat(i, cR)), CLng(dat(i, cC)), Trim$(CStr(dat(i, cI))))
                End If
            End If
        Next i
    End If
    Set ReadModels = d
End Function

Private Function CleanFilter(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = Trim$(CStr(v))
    Select Case LCase$(s)
        Case "", "all", "(all)", "tat ca", "*"
            CleanFilter = ""
        Case Else
            CleanFilter = s
    End Select
End Function

Private Function CountModeIsRecords() As Boolean
    Dim lo As ListObject, c As Range
    On Error GoTo Done
    Set lo = DashSheet(SH_CFG).ListObjects(TBL_DSET)
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each c In lo.ListColumns(1).DataBodyRange.Cells
        If StrComp(Trim$(CStr(c.Value)), "CountMode", vbTextCompare) = 0 Then
            CountModeIsRecords = (StrComp(Trim$(CStr(c.Offset(0, 1).Value)), "Records", vbTextCompare) = 0)
            Exit Function
        End If
    Next c
Done:
End Function

'Gop tblDefects theo bo loc Dashboard:
'  grp("Model|DefectType") = mang 3x3 Top 3 (gia tri, hang, cot); totDef / totMod = tong theo Defect Type / Model
Private Sub Aggregate(ByVal ws As Worksheet, ByVal mdl As Object, ByRef grp As Object, _
                      ByRef totDef As Object, ByRef totMod As Object, ByRef nSkip As Long)
    Dim lo As ListObject, dat As Variant, i As Long, agg As Object, k As Variant
    Dim cM As Long, cX As Long, cY As Long, cQ As Long, cD As Long, cSD As Long, cSh As Long, cCell As Long
    Dim fModel As String, fShift As String, fDef As String, fCell As String
    Dim hasFrom As Boolean, hasTo As Boolean, dFrom As Long, dTo As Long, sd As Long
    Dim mName As String, dName As String, r As Long, c As Long, v As Double, info As Variant
    Dim useRecords As Boolean, key As String, p() As String, rc() As String, gk As String
    Dim t As Variant, blank(1 To 3, 1 To 3) As Double

    Set grp = CreateObject("Scripting.Dictionary")
    Set totDef = CreateObject("Scripting.Dictionary")
    Set totMod = CreateObject("Scripting.Dictionary")
    Set agg = CreateObject("Scripting.Dictionary")
    grp.CompareMode = vbTextCompare
    totDef.CompareMode = vbTextCompare
    totMod.CompareMode = vbTextCompare
    nSkip = 0

    fModel = CleanFilter(ws.Range("E5").Value)
    fShift = UCase$(CleanFilter(ws.Range("G5").Value))
    fDef = CleanFilter(ws.Range("I5").Value)
    fCell = CleanFilter(ws.Range("K5").Value)
    If IsDate(ws.Range("A5").Value) Then hasFrom = True: dFrom = CLng(Int(CDbl(CDate(ws.Range("A5").Value))))
    If IsDate(ws.Range("C5").Value) Then hasTo = True: dTo = CLng(Int(CDbl(CDate(ws.Range("C5").Value))))
    useRecords = CountModeIsRecords()

    Set lo = ThisWorkbook.Worksheets(SH_DB).ListObjects(TBL_DEFECTS)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    cM = lo.ListColumns("Model").Index
    cX = lo.ListColumns("X").Index
    cY = lo.ListColumns("Y").Index
    cQ = lo.ListColumns("Qty").Index
    cD = lo.ListColumns("DefectType").Index
    cSD = lo.ListColumns("ShiftDate").Index
    cSh = lo.ListColumns("Shift").Index
    cCell = lo.ListColumns("CellID").Index
    dat = lo.DataBodyRange.Value2

    For i = 1 To UBound(dat, 1)
        If IsError(dat(i, cM)) Or IsError(dat(i, cX)) Or IsError(dat(i, cY)) Or IsError(dat(i, cD)) Then GoTo NextRow
        mName = Trim$(CStr(dat(i, cM)))
        dName = Trim$(CStr(dat(i, cD)))
        If Len(mName) = 0 Or Len(dName) = 0 Then GoTo NextRow
        If Not mdl.Exists(mName) Then GoTo NextRow
        If Len(fModel) > 0 Then
            If StrComp(mName, fModel, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If Len(fDef) > 0 Then
            If StrComp(dName, fDef, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If Len(fShift) > 0 Then
            If IsError(dat(i, cSh)) Then GoTo NextRow
            If UCase$(Trim$(CStr(dat(i, cSh)))) <> fShift Then GoTo NextRow
        End If
        If Len(fCell) > 0 Then
            If IsError(dat(i, cCell)) Then GoTo NextRow
            If StrComp(Trim$(CStr(dat(i, cCell))), fCell, vbTextCompare) <> 0 Then GoTo NextRow
        End If
        If hasFrom Or hasTo Then
            sd = 0
            If Not IsError(dat(i, cSD)) Then
                If IsNumeric(dat(i, cSD)) Then
                    sd = CLng(dat(i, cSD))
                ElseIf IsDate(dat(i, cSD)) Then
                    sd = CLng(Int(CDbl(CDate(dat(i, cSD)))))
                End If
            End If
            If hasFrom Then
                If sd < dFrom Then GoTo NextRow
            End If
            If hasTo Then
                If sd > dTo Then GoTo NextRow
            End If
        End If
        If Not (IsNumeric(dat(i, cX)) And IsNumeric(dat(i, cY))) Then GoTo NextRow
        c = CLng(dat(i, cX)): r = CLng(dat(i, cY))
        info = mdl(mName)
        If r < 1 Or r > CLng(info(0)) Or c < 1 Or c > CLng(info(1)) Then
            nSkip = nSkip + 1
            GoTo NextRow
        End If
        If useRecords Then
            v = 1
        ElseIf IsNumeric(dat(i, cQ)) And Not IsError(dat(i, cQ)) Then
            v = CDbl(dat(i, cQ))
        Else
            v = 0
        End If
        If v <= 0 Then GoTo NextRow

        key = mName & "|" & dName & "|" & r & "," & c
        If agg.Exists(key) Then agg(key) = agg(key) + v Else agg.Add key, v
        If totDef.Exists(dName) Then totDef(dName) = totDef(dName) + v Else totDef.Add dName, v
        If totMod.Exists(mName) Then totMod(mName) = totMod(mName) + v Else totMod.Add mName, v
NextRow:
    Next i

    For Each k In agg.Keys
        p = Split(CStr(k), "|")
        rc = Split(p(2), ",")
        gk = p(0) & "|" & p(1)
        If grp.Exists(gk) Then
            t = grp(gk)
        Else
            Erase blank
            t = blank
        End If
        InsertTop3 t, CDbl(agg(k)), CLng(rc(0)), CLng(rc(1))
        grp(gk) = t
    Next k
End Sub

Private Sub InsertTop3(ByRef t As Variant, ByVal v As Double, ByVal r As Long, ByVal c As Long)
    Dim p As Long, q As Long
    For p = 1 To 3
        If v > t(p, 1) Then
            For q = 3 To p + 1 Step -1
                t(q, 1) = t(q - 1, 1): t(q, 2) = t(q - 1, 2): t(q, 3) = t(q - 1, 3)
            Next q
            t(p, 1) = v: t(p, 2) = r: t(p, 3) = c
            Exit Sub
        End If
    Next p
End Sub

Private Function SortedKeys(ByVal d As Object, ByVal maxN As Long) As String()
    Dim k() As String, v() As Double, n As Long, i As Long, j As Long, key As Variant
    Dim ts As String, tv As Double, out() As String, m As Long
    n = d.Count
    ReDim k(1 To n)
    ReDim v(1 To n)
    i = 0
    For Each key In d.Keys
        i = i + 1
        k(i) = CStr(key)
        v(i) = CDbl(d(key))
    Next key
    For i = 1 To n - 1
        For j = i + 1 To n
            If v(j) > v(i) Then
                tv = v(i): v(i) = v(j): v(j) = tv
                ts = k(i): k(i) = k(j): k(j) = ts
            End If
        Next j
    Next i
    m = n
    If m > maxN Then m = maxN
    ReDim out(1 To m)
    For i = 1 To m
        out(i) = k(i)
    Next i
    SortedKeys = out
End Function

Private Function CellLetter(ByVal n As Long) As String
    'Cell ID dang Excel giong SaveDefectRecords: ColumnLetter(c) & r
    Dim s As String
    Do While n > 0
        s = Chr$(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    CellLetter = s
End Function

'=====================================================================
' HINH VA SHAPE
'=====================================================================
Private Function FindShp(ByVal ws As Worksheet, ByVal nm As String) As Shape
    On Error Resume Next
    Set FindShp = ws.Shapes(nm)
    On Error GoTo 0
End Function

'Uu tien hinh that; hinh tam (AutoShape) chi dung khi chua co hinh that
Private Function FindImg(ByVal wm As Worksheet, ByVal nm As String) As Shape
    Dim s As Shape, temp As Shape
    For Each s In wm.Shapes
        If StrComp(s.Name, nm, vbTextCompare) = 0 Then
            If s.Type = msoAutoShape Then
                If temp Is Nothing Then Set temp = s
            Else
                Set FindImg = s
                Exit Function
            End If
        End If
    Next s
    Set FindImg = temp
End Function

'Copy hinh Master cua Model vao Dashboard (tam bat hien thi neu dang an, roi tra lai nhu cu).
'Chi copy lai khi force = True (REFRESH) hoac chua co.
Private Function ModelPicture(ByVal ws As Worksheet, ByVal model As String, ByVal imgName As String, _
                              ByVal force As Boolean) As Shape
    Dim nm As String, s As Shape, src As Shape, wm As Worksheet, wasVis As Boolean
    Dim tries As Long, ok As Boolean
    nm = "HM_PIC_" & model
    Set s = FindShp(ws, nm)
    If Not s Is Nothing Then
        If Not force Then
            Set ModelPicture = s
            Exit Function
        End If
        s.Delete
        Set s = Nothing
    End If
    If Len(imgName) = 0 Then Exit Function
    Set wm = ThisWorkbook.Worksheets(SH_MAP)
    Set src = FindImg(wm, imgName)
    If src Is Nothing Then Exit Function

    wasVis = (src.Visible = msoTrue)
    If Not wasVis Then src.Visible = msoTrue
    For tries = 1 To 3
        On Error Resume Next
        Err.Clear
        src.Copy
        If Err.Number = 0 Then ws.Paste Destination:=ws.Range("A" & HM_FIRST_ROW)
        ok = (Err.Number = 0)
        On Error GoTo 0
        If ok Then Exit For
        DoEvents
    Next tries
    On Error Resume Next
    If Not wasVis Then src.Visible = msoFalse
    On Error GoTo 0
    If Not ok Then Exit Function

    Set s = ws.Shapes(ws.Shapes.Count)
    s.Name = nm
    s.Placement = xlMove
    Set ModelPicture = s
End Function

'Xoa marker / tieu de / legend (giu hinh nen da copy)
Private Sub ClearHeatShapes(ByVal ws As Worksheet)
    Dim i As Long, nm As String
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Left$(nm, 5) = "HM_M_" Or Left$(nm, 5) = "HM_T_" Or Left$(nm, 5) = "HM_L_" Then ws.Shapes(i).Delete
    Next i
End Sub

'Xoa hinh nen cua Model khong con hien thi
Private Sub DeleteStalePics(ByVal ws As Worksheet, ByVal keep As Object)
    Dim i As Long, nm As String
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Left$(nm, 7) = "HM_PIC_" Then
            If Not keep.Exists(Mid$(nm, 8)) Then ws.Shapes(i).Delete
        End If
    Next i
End Sub

Private Function DrawLegend(ByVal ws As Worksheet, ByRef defs() As String, ByVal nD As Long, _
                            ByVal availW As Double) As Double
    Dim i As Long, x As Double, y As Double, w As Double, x0 As Double, y0 As Double
    Dim clr As Long, s As Shape
    x0 = ws.Range("A1").Left + 2
    y0 = ws.Rows(HM_FIRST_ROW).Top + 2
    x = x0: y = y0
    For i = 1 To nD
        w = Len(defs(i)) * 5.6 + 18
        If w < 44 Then w = 44
        If x + w > x0 + availW And x > x0 Then
            x = x0
            y = y + 18
        End If
        clr = HmColor(defs(i), i)
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, 15)
        With s
            .Name = "HM_L_" & i
            .Fill.ForeColor.RGB = clr
            .Line.Visible = msoFalse
            .Placement = xlMove
            With .TextFrame2
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
                .VerticalAnchor = msoAnchorMiddle
                .WordWrap = msoFalse
                .TextRange.Text = defs(i)
                .TextRange.Font.Size = 8
                .TextRange.Font.Bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = TextColorFor(clr)
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        x = x + w + 6
    Next i
    DrawLegend = (y - y0) + 15
End Function

Private Sub StyleMarker(ByVal s As Shape, ByVal nm As String, ByVal txt As String, _
                        ByVal clr As Long, ByVal tip As String, ByVal sz As Double)
    With s
        .Name = nm
        .Fill.ForeColor.RGB = clr
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 1.25
        .Placement = xlMove
        .AlternativeText = tip
        With .TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            With .TextRange
                .Text = txt
                .Font.Size = IIf(sz >= 14, 9, 8)
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = TextColorFor(clr)
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    End With
End Sub

Private Function TextColorFor(ByVal clr As Long) As Long
    Dim lum As Double
    lum = 0.299 * (clr And 255) + 0.587 * ((clr \ 256) And 255) + 0.114 * ((clr \ 65536) And 255)
    If lum > 150 Then TextColorFor = RGB(0, 0, 0) Else TextColorFor = RGB(255, 255, 255)
End Function

'Mau theo Config (cot Color cua bang Defect Type); chua co thi dung bang mau mac dinh
Private Function HmColor(ByVal defectName As String, ByVal idx As Long) As Long
    Dim lo As ListObject, c As Range, pal As Variant
    pal = Array(RGB(31, 119, 180), RGB(255, 127, 14), RGB(44, 160, 44), RGB(214, 39, 40), RGB(148, 103, 189), _
                RGB(140, 86, 75), RGB(227, 119, 194), RGB(127, 127, 127), RGB(188, 189, 34), RGB(23, 190, 207))
    HmColor = pal((idx - 1) Mod 10)
    On Error Resume Next
    Set lo = DashSheet(SH_CFG).ListObjects(TBL_DCOL)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each c In lo.ListColumns(1).DataBodyRange.Cells
        If StrComp(Trim$(CStr(c.Value)), defectName, vbTextCompare) = 0 Then
            If c.Offset(0, 1).Interior.ColorIndex <> xlNone Then HmColor = c.Offset(0, 1).Interior.Color
            Exit Function
        End If
    Next c
End Function
