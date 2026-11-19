Attribute VB_Name = "modDashData"
Option Explicit
'=====================================================================
' DASHBOARD - DOC DU LIEU tblDefects, LOC VA TONG HOP (mang + Dictionary)
' Nguon: sheet DATABASE > tblDefects (moi dong = 1 o loi; so luong = Qty)
' Location = o tren Grid (CellID). Khong sua du lieu nguon.
'=====================================================================

'--- Bo loc hien tai (doc tu Dashboard) ---
Public fFrom As Long, fTo As Long          'ky hien tai (so serial ngay)
Public pFrom As Long, pTo As Long          'ky truoc (cung do dai)
Public fModel As String, fShift As String, fDef As String, fCell As String
Public fTopN As Long

'--- Ket qua (ky hien tai, sau loc) ---
Public rRows As Long                       'so dong doc tu tblDefects
Public rSkipped As Long                    'so dong du lieu khong hop le bi bo qua
Public rFirstBad As Long                   'so dong Excel dau tien bi loi
Public rMatch As Long                      'so dong khop bo loc (ky hien tai)
Public rTotal As Double                    'tong so loi (tong Qty hoac so dong, theo Config > CountMode)
Public rVin As Object                      'tap VIN duy nhat
Public rDefQty As Object                   'DefectType -> so loi
Public rDefVin As Object                   'DefectType -> tap VIN
Public rDefPrev As Object                  'DefectType -> so loi ky truoc
Public rDateShift As Object                'ngay|ca -> so loi
Public rDateDef As Object                  'ngay|loai loi -> so loi
Public rShiftQty As Object                 'ca -> so loi
Public rShiftVin As Object                 'ca -> tap VIN
Public rModelQty As Object                 'Model -> so loi
Public rModelCell As Object                'Model -> (CellID -> so loi)
Public rDefCell As Object                  'DefectType -> (CellID -> so loi)
Public rCellQty As Object                  'CellID -> so loi
Public rMatrix As Object                   'loai loi|Model -> so loi

'--- Danh sach cho dropdown (toan bo du lieu, khong phu thuoc bo loc) ---
Public lModels As Object
Public lDefects As Object
Public lCells As Object                    'CellID -> khoa sap xep (hang*10000 + cot)
Public lShifts As Object

'--- Bang chi tiet ---
Public rDetail As Variant                  'mang 2 chieu (1..n, 1..12)
Public rDetailN As Long

Private cDT As Long, cSD As Long, cSh As Long, cMod As Long, cVin As Long, cDef As Long
Private cCell As Long, cX As Long, cY As Long, cQty As Long, cIns As Long, cDesc As Long
Private mCountRecords As Boolean

'=====================================================================
' Bo loc
'=====================================================================
Private Function FilterText(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    s = Trim$(CStr(v))
    If LCase$(s) = "all" Then s = ""
    FilterText = s
End Function

Public Function ReadFilters(ByRef msg As String) As Boolean
    Dim ws As Worksheet, v As Variant, days As Long
    Set ws = DashSheet(SH_DASH)

    days = CLng(Val(CStr(CfgValue("DefaultDays", 30))))
    If days < 1 Then days = 30

    v = ws.Range(F_TO).Value
    If IsDate(v) Then fTo = CLng(Int(CDbl(CDate(v)))) Else fTo = CLng(Int(CDbl(Date)))
    v = ws.Range(F_FROM).Value
    If IsDate(v) Then fFrom = CLng(Int(CDbl(CDate(v)))) Else fFrom = fTo - days + 1
    If fFrom > fTo Then
        msg = "From date lon hon To date. Hay chon lai khoang ngay."
        Exit Function
    End If
    pTo = fFrom - 1
    pFrom = fFrom - (fTo - fFrom + 1)

    fModel = FilterText(ws.Range(F_MODEL).Value)
    fShift = UCase$(FilterText(ws.Range(F_SHIFT).Value))
    fDef = FilterText(ws.Range(F_DEF).Value)
    fCell = FilterText(ws.Range(F_CELL).Value)

    fTopN = CLng(Val(CStr(ws.Range(F_TOPN).Value)))
    If fTopN < 1 Then fTopN = CLng(Val(CStr(CfgValue("TopN", 5))))
    If fTopN < 1 Then fTopN = 5
    If fTopN > 10 Then fTopN = 10
    ReadFilters = True
End Function

Private Function PassBase(ByVal mdl As String, ByVal sh As String, ByVal df As String, ByVal cl As String) As Boolean
    If Len(fModel) > 0 Then
        If StrComp(mdl, fModel, vbTextCompare) <> 0 Then Exit Function
    End If
    If Len(fShift) > 0 Then
        If StrComp(sh, fShift, vbTextCompare) <> 0 Then Exit Function
    End If
    If Len(fDef) > 0 Then
        If StrComp(df, fDef, vbTextCompare) <> 0 Then Exit Function
    End If
    If Len(fCell) > 0 Then
        If StrComp(cl, fCell, vbTextCompare) <> 0 Then Exit Function
    End If
    PassBase = True
End Function

'=====================================================================
' Doc 1 dong du lieu
'=====================================================================
Private Function ParseRow(ByRef src As Variant, ByVal i As Long, ByRef sd As Long, ByRef sh As String, _
                          ByRef mdl As String, ByRef vn As String, ByRef df As String, ByRef cl As String, _
                          ByRef q As Double) As Boolean
    'Tra ve False neu dong khong hop le (thieu Model/VIN/Defect/Cell, Qty sai, khong xac dinh duoc ngay/ca)
    Dim v As Variant, dt As Date, hasDT As Boolean, pd As Date, sn As String
    On Error GoTo Bad

    mdl = Trim$(CStr(src(i, cMod)))
    vn = Trim$(CStr(src(i, cVin)))
    df = Trim$(CStr(src(i, cDef)))
    cl = Trim$(CStr(src(i, cCell)))
    If Len(mdl) = 0 Or Len(vn) = 0 Or Len(df) = 0 Or Len(cl) = 0 Then Exit Function

    If mCountRecords Then
        q = 1
    Else
        v = src(i, cQty)
        If IsError(v) Then Exit Function
        If IsEmpty(v) Then Exit Function
        If Not IsNumeric(v) Then Exit Function
        q = CDbl(v)
        If q <= 0 Then Exit Function
    End If

    v = src(i, cDT)
    If Not IsEmpty(v) And Not IsError(v) Then
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then dt = CDate(CDbl(v)): hasDT = True
        ElseIf IsDate(v) Then
            dt = CDate(v): hasDT = True
        End If
    End If

    sd = 0
    v = src(i, cSD)
    If Not IsEmpty(v) And Not IsError(v) Then
        If IsNumeric(v) Then
            If CDbl(v) > 0 Then sd = CLng(Int(CDbl(v)))
        ElseIf IsDate(v) Then
            sd = CLng(Int(CDbl(CDate(v))))
        End If
    End If

    sh = UCase$(Trim$(CStr(src(i, cSh))))
    If hasDT And (sd = 0 Or Len(sh) = 0) Then
        GetShiftInfo dt, sn, pd                    'ham co san trong modVehicleMapping
        If sd = 0 Then sd = CLng(Int(CDbl(pd)))
        If Len(sh) = 0 Then sh = sn
    End If
    If sd = 0 Or Len(sh) = 0 Then Exit Function
    ParseRow = True
    Exit Function
Bad:
    ParseRow = False
End Function

Private Function CellSortKey(ByVal x As Variant, ByVal y As Variant) As Double
    CellSortKey = 99999999
    If IsEmpty(x) Or IsEmpty(y) Or IsError(x) Or IsError(y) Then Exit Function
    If IsNumeric(x) And IsNumeric(y) Then CellSortKey = CDbl(y) * 10000# + CDbl(x)
End Function

Private Sub InitCols(ByVal lo As ListObject)
    cDT = DashColIndex(lo, "DateTime")
    cSD = DashColIndex(lo, "ShiftDate")
    cSh = DashColIndex(lo, "Shift")
    cMod = DashColIndex(lo, "Model")
    cVin = DashColIndex(lo, "VIN")
    cDef = DashColIndex(lo, "DefectType")
    cCell = DashColIndex(lo, "CellID")
    cX = DashColIndex(lo, "X")
    cY = DashColIndex(lo, "Y")
    cQty = DashColIndex(lo, "Qty")
    cIns = DashColIndex(lo, "Inspector", False)
    cDesc = DashColIndex(lo, "Description", False)
    mCountRecords = (LCase$(Trim$(CStr(CfgValue("CountMode", "Qty")))) = "records")
End Sub

Private Sub ResetResults()
    rRows = 0: rSkipped = 0: rFirstBad = 0: rMatch = 0: rTotal = 0
    Set rVin = NewDict(): Set rDefQty = NewDict(): Set rDefVin = NewDict(): Set rDefPrev = NewDict()
    Set rDateShift = NewDict(): Set rDateDef = NewDict(): Set rShiftQty = NewDict(): Set rShiftVin = NewDict()
    Set rModelQty = NewDict(): Set rModelCell = NewDict(): Set rDefCell = NewDict(): Set rCellQty = NewDict()
    Set rMatrix = NewDict()
    Set lModels = NewDict(): Set lDefects = NewDict(): Set lCells = NewDict(): Set lShifts = NewDict()
    rDetail = Empty: rDetailN = 0
End Sub

Private Sub MergeMasterList(ByVal tblName As String, ByVal nameCol As String, ByVal target As Object)
    'Them cac muc Active = Yes tu MASTER vao danh sach dropdown (de chon duoc ca muc chua co du lieu)
    Dim lo As ListObject, dat As Variant, i As Long, cN As Long, cA As Long, nm As String
    Set lo = DashTryTable(SH_MASTER, tblName)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    cN = DashColIndex(lo, nameCol, False)
    cA = DashColIndex(lo, "Active", False)
    If cN = 0 Then Exit Sub
    dat = lo.DataBodyRange.Value2
    For i = 1 To UBound(dat, 1)
        nm = Trim$(CStr(dat(i, cN)))
        If Len(nm) > 0 Then
            If cA = 0 Then
                target(nm) = True
            ElseIf IsActiveValue(dat(i, cA)) Then
                target(nm) = True
            End If
        End If
    Next i
End Sub

'=====================================================================
' Doc + loc + tong hop (1 luot qua du lieu)
'=====================================================================
Public Function LoadAndAggregate() As Boolean
    Dim lo As ListObject, src As Variant, n As Long, i As Long, j As Long
    Dim sd As Long, sh As String, mdl As String, vn As String, df As String, cl As String, q As Double
    Dim maxDetail As Long, d() As Variant, t() As Variant, a As Long, b As Long
    Dim firstRow As Long

    ResetResults
    gDashStep = "doc tblDefects"
    Set lo = DashTable(SH_DB, TBL_DEFECTS)
    InitCols lo
    If lo.DataBodyRange Is Nothing Then
        LoadAndAggregate = True
        GoTo Finish
    End If
    firstRow = lo.DataBodyRange.Row
    src = lo.DataBodyRange.Value2
    n = UBound(src, 1)

    gDashStep = "tong hop du lieu"
    For i = 1 To n
        If IsEmpty(src(i, cMod)) And IsEmpty(src(i, cVin)) And IsEmpty(src(i, cDef)) Then GoTo NextRow   'dong trong
        rRows = rRows + 1
        If Not ParseRow(src, i, sd, sh, mdl, vn, df, cl, q) Then
            rSkipped = rSkipped + 1
            If rFirstBad = 0 Then rFirstBad = firstRow + i - 1
            GoTo NextRow
        End If

        lModels(mdl) = True
        lDefects(df) = True
        lShifts(sh) = True
        If Not lCells.Exists(cl) Then lCells.Add cl, CellSortKey(src(i, cX), src(i, cY))

        If Not PassBase(mdl, sh, df, cl) Then GoTo NextRow

        If sd >= fFrom And sd <= fTo Then
            rMatch = rMatch + 1
            rTotal = rTotal + q
            rVin(vn) = True
            AddTo rDefQty, df, q
            AddNested rDefVin, df, vn, 1
            AddTo rDateShift, sd & "|" & sh, q
            AddTo rDateDef, sd & "|" & df, q
            AddTo rShiftQty, sh, q
            AddNested rShiftVin, sh, vn, 1
            AddTo rModelQty, mdl, q
            AddNested rModelCell, mdl, cl, q
            AddNested rDefCell, df, cl, q
            AddTo rCellQty, cl, q
            AddTo rMatrix, df & "|" & mdl, q
        ElseIf sd >= pFrom And sd <= pTo Then
            AddTo rDefPrev, df, q
        End If
NextRow:
    Next i

    'Bang chi tiet: lay cac dong moi nhat khop bo loc (duyet tu cuoi len)
    gDashStep = "bang chi tiet"
    maxDetail = CLng(Val(CStr(CfgValue("DetailMaxRows", 500))))
    If maxDetail < 1 Then maxDetail = 500
    If maxDetail > 5000 Then maxDetail = 5000
    If rMatch > 0 Then
        If maxDetail > rMatch Then maxDetail = rMatch
        ReDim d(1 To maxDetail, 1 To 12)
        j = 0
        For i = n To 1 Step -1
            If ParseRow(src, i, sd, sh, mdl, vn, df, cl, q) Then
                If PassBase(mdl, sh, df, cl) Then
                    If sd >= fFrom And sd <= fTo Then
                        j = j + 1
                        d(j, 1) = CDbl(sd)
                        d(j, 2) = src(i, cDT)
                        d(j, 3) = sh
                        d(j, 4) = mdl
                        d(j, 5) = vn
                        d(j, 6) = df
                        d(j, 7) = src(i, cQty)
                        d(j, 8) = src(i, cY)
                        d(j, 9) = src(i, cX)
                        d(j, 10) = cl
                        If cIns > 0 Then d(j, 11) = src(i, cIns)
                        If cDesc > 0 Then d(j, 12) = src(i, cDesc)
                        If j >= maxDetail Then Exit For
                    End If
                End If
            End If
        Next i
        If j = maxDetail Then
            rDetail = d
        ElseIf j > 0 Then
            ReDim t(1 To j, 1 To 12)
            For a = 1 To j
                For b = 1 To 12
                    t(a, b) = d(a, b)
                Next b
            Next a
            rDetail = t
        End If
        rDetailN = j
    End If
    LoadAndAggregate = True

Finish:
    MergeMasterList TBL_MODELS, "Model", lModels
    MergeMasterList TBL_DEFECTTYPES, "DefectType", lDefects
End Function
