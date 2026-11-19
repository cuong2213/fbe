Attribute VB_Name = "modVehicleMapping"
Option Explicit
'=====================================================================
' VEHICLE DEFECT MAPPING SYSTEM  -  Vehicle_Defect_Mapping.xlsm
' Module chinh. KHONG dung UserForm.
' Quy uoc: toan bo chuoi hien thi dung tieng Viet KHONG DAU de tranh
'          loi ma hoa trong VBA Editor.
' Cach dung nhanh: chay Sub SetupWorkbook mot lan de tao 4 sheet,
'          cac bang, giao dien, nut bam va hinh placeholder.
'=====================================================================

'---------------------------- HANG SO -------------------------------
Public Const SH_MAP As String = "MAP_VEHICLE"
Public Const SH_MASTER As String = "MASTER"
Public Const SH_DB As String = "DATABASE"
Public Const SH_SUM As String = "SUMMARY"

Public Const TBL_MODELS As String = "tblModels"
Public Const TBL_DEFECTTYPES As String = "tblDefectTypes"
Public Const TBL_SHIFTS As String = "tblShifts"
Public Const TBL_ZONES As String = "tblGridZones"
Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_DEFECTS As String = "tblDefects"
Public Const TBL_SUMMARY As String = "tblSummary"

Private Const ANCHOR_IMG As String = "F2"        'goc tren-trai cua hinh Master (ben PHAI panel nhap lieu)
Private Const IMG_MAX_WIDTH As Double = 1100       'chi THU NHO neu hinh rong hon (point); 0 = giu nguyen kich thuoc hinh
Private Const IMG_PLACEHOLDER_WIDTH As Double = 900
Private Const GRID_PREFIX As String = "CELL_"
Private Const IMG_PREFIX As String = "IMG_"
Private Const COLOR_SELECTED As Long = 5287936    'RGB(0,176,80) xanh la
Private Const MAX_LIST_ROWS As Long = 60          'so dong toi da cua danh sach dropdown

'---------------------------- BIEN TOAN CUC -------------------------
Public gModel As String
Public gRows As Long
Public gCols As Long
Public gSelectedCells As Object        'Scripting.Dictionary, key = "row,col"
Public gVinJustEntered As Boolean      'co bao: vua nhap VIN xong (dung cho SelectionChange)

Private mSeq As Long
Private mGridRows As Long             'kich thuoc Grid dang co tren sheet (de xoa nhanh bang 1 lenh)
Private mGridCols As Long
Private mHeat As Object               'Dictionary "row,col" -> gia tri dang hien thi tren Grid (heatmap)
Private mHeatMax As Double
Private mHeatActive As Boolean        'True sau khi bam SHOW DATA (de tu lam moi sau Save)
Private mStep As String               'buoc dang chay (hien trong thong bao loi de biet loi o dau)
Private mCalOpen As Boolean           'lich chon ngay (Shape) dang mo
Private mCalTarget As String          '"B21" hoac "B22"
Private mCalYear As Long
Private mCalMonth As Long
Private mCalCol As Collection         'ten cac shape cua lich

'=====================================================================
' 1. HAM TIEN ICH
'=====================================================================
Private Function GetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Err.Raise vbObjectError + 513, "GetSheet", "Khong tim thay sheet '" & sheetName & "'. Hay chay SetupWorkbook."
    End If
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then SheetExists = True: Exit Function
    Next ws
End Function

Public Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Set ws = GetSheet(sheetName)
    On Error Resume Next
    Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
    If GetTable Is Nothing Then
        Err.Raise vbObjectError + 514, "GetTable", "Khong tim thay bang '" & tableName & "' tren sheet '" & sheetName & "'."
    End If
End Function

Private Function ColIdx(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then ColIdx = lc.Index: Exit Function
    Next lc
    Err.Raise vbObjectError + 515, "ColIdx", "Bang '" & lo.Name & "' thieu cot '" & colName & "'."
End Function

Private Function TableHasData(ByVal lo As ListObject) As Boolean
    If lo.DataBodyRange Is Nothing Then Exit Function
    TableHasData = (Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0)
End Function

Private Function FindShape(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set FindShape = ws.Shapes(shapeName)
    On Error GoTo 0
End Function

Private Function FindImageShape(ByVal ws As Worksheet, ByVal imgName As String) As Shape
    'Uu tien hinh that (Picture...). Hinh tam (AutoShape) chi dung khi chua co hinh that; neu co hinh that thi xoa hinh tam trung ten.
    Dim s As Shape, real As Shape, temp As Shape, i As Long
    Set s = FindShape(ws, imgName)                  'duong nhanh: tra theo ten
    If Not s Is Nothing Then
        If s.Type <> msoAutoShape Then Set FindImageShape = s: Exit Function
    End If
    Set s = Nothing
    For Each s In ws.Shapes                         'duong cham: co hinh tam / trung ten
        If StrComp(s.Name, imgName, vbTextCompare) = 0 Then
            If s.Type = msoAutoShape Then
                If temp Is Nothing Then Set temp = s
            Else
                If real Is Nothing Then Set real = s
            End If
        End If
    Next s
    If Not real Is Nothing Then
        For i = ws.Shapes.Count To 1 Step -1
            If StrComp(ws.Shapes(i).Name, imgName, vbTextCompare) = 0 Then
                If ws.Shapes(i).Type = msoAutoShape Then ws.Shapes(i).Delete
            End If
        Next i
        Set FindImageShape = real
    Else
        Set FindImageShape = temp
    End If
End Function

Public Function ColumnLetter(ByVal colNum As Long) As String
    Dim n As Long, s As String
    n = colNum
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColumnLetter = s
End Function

Public Function NewID(ByVal prefix As String) As String
    mSeq = (mSeq + 1) Mod 100
    NewID = prefix & Format$(Now, "yyyymmddhhnnss") & _
            Format$(Int((Timer - Int(Timer)) * 1000), "000") & Format$(mSeq, "00")
End Function

Public Function IsActiveValue(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    IsActiveValue = (LCase$(Trim$(CStr(v))) = "yes")
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    IsNum = IsNumeric(v)
End Function

Public Function GetGridCoordinates(ByVal shapeName As String, ByRef r As Long, ByRef c As Long) As Boolean
    'Ten shape: CELL_R{row}_C{col}
    Dim p() As String
    If Left$(shapeName, 6) <> "CELL_R" Then Exit Function
    p = Split(Mid$(shapeName, 7), "_C")
    If UBound(p) <> 1 Then Exit Function
    If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then Exit Function
    r = CLng(p(0))
    c = CLng(p(1))
    GetGridCoordinates = True
End Function

Private Function GetModelInfo(ByVal modelName As String, ByRef nRows As Long, ByRef nCols As Long, _
                              ByRef imgName As String) As Boolean
    'Tra ve True neu Model ton tai trong tblModels va Active = Yes
    Dim lo As ListObject, dat As Variant, i As Long
    Dim cM As Long, cR As Long, cC As Long, cI As Long, cA As Long
    Set lo = GetTable(SH_MASTER, TBL_MODELS)
    If Not TableHasData(lo) Then Exit Function
    cM = ColIdx(lo, "Model"): cR = ColIdx(lo, "Rows"): cC = ColIdx(lo, "Cols")
    cI = ColIdx(lo, "ImageShapeName"): cA = ColIdx(lo, "Active")
    dat = lo.DataBodyRange.Value2
    For i = 1 To UBound(dat, 1)
        If StrComp(Trim$(CStr(dat(i, cM))), Trim$(modelName), vbTextCompare) = 0 Then
            If IsActiveValue(dat(i, cA)) And IsNum(dat(i, cR)) And IsNum(dat(i, cC)) Then
                nRows = CLng(dat(i, cR))
                nCols = CLng(dat(i, cC))
                imgName = Trim$(CStr(dat(i, cI)))
                If nRows >= 1 And nCols >= 1 Then GetModelInfo = True
            End If
            Exit Function
        End If
    Next i
End Function

Private Function GetGridModel(ByVal ws As Worksheet) As String
    'Model ma Grid hien tai duoc tao cho (luu trong AlternativeText cua shape)
    Dim s As Shape
    Set s = FindShape(ws, GRID_PREFIX & "R1_C1")
    If Not s Is Nothing Then GetGridModel = s.AlternativeText
End Function

'=====================================================================
' 2. SHIFT / SHIFTDATE
'=====================================================================
Public Sub GetShiftInfo(ByVal inputTime As Date, ByRef shiftName As String, ByRef productionDate As Date)
    Dim h As Long
    h = Hour(inputTime)
    If h >= 6 And h < 14 Then
        shiftName = "S1": productionDate = DateValue(inputTime)
    ElseIf h >= 14 And h < 22 Then
        shiftName = "S2": productionDate = DateValue(inputTime)
    ElseIf h >= 22 Then
        shiftName = "S3": productionDate = DateValue(inputTime)
    Else                                   '00:00 - 05:59
        shiftName = "S3": productionDate = DateValue(inputTime) - 1
    End If
End Sub

Public Sub UpdateShiftDisplay()
    Dim ws As Worksheet, t As Date, sn As String, pd As Date, prevEv As Boolean
    On Error GoTo EH
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Set ws = GetSheet(SH_MAP)
    t = Now
    GetShiftInfo t, sn, pd
    ws.Range("B7").Value = sn
    ws.Range("D2").Value = pd
    ws.Range("D3").Value = t
CLEANUP:
    Application.EnableEvents = prevEv
    Exit Sub
EH:
    Resume CLEANUP
End Sub

'=====================================================================
' 3. MODEL / ANH / GRID
'=====================================================================
Private Function GetImageSize(ByRef w As Double, ByRef h As Double) As Boolean
    'Doc kich thuoc hinh Master CHUNG cho moi Model tu MASTER > tblSettings (ImageWidth / ImageHeight, don vi point)
    'Tra ve False neu khong co bang hoac gia tri khong hop le -> dung cach cu (giu ti le, toi da IMG_MAX_WIDTH)
    Dim lo As ListObject, dat As Variant, i As Long, cS As Long, cV As Long, k As String
    On Error GoTo Fail
    w = 0: h = 0
    Set lo = GetTable(SH_MASTER, TBL_SETTINGS)
    If Not TableHasData(lo) Then Exit Function
    cS = ColIdx(lo, "Setting"): cV = ColIdx(lo, "Value")
    dat = lo.DataBodyRange.Value2
    For i = 1 To UBound(dat, 1)
        k = LCase$(Trim$(CStr(dat(i, cS))))
        If k = "imagewidth" And IsNum(dat(i, cV)) Then w = CDbl(dat(i, cV))
        If k = "imageheight" And IsNum(dat(i, cV)) Then h = CDbl(dat(i, cV))
    Next i
    GetImageSize = (w >= 50 And h >= 50)
    Exit Function
Fail:
    GetImageSize = False
End Function

Private Function CanReuseGrid(ByVal ws As Worksheet, ByVal nR As Long, ByVal nC As Long, ByVal img As Shape) As Boolean
    'Grid dang co tren sheet khop hoan toan (so hang/cot, vi tri, kich thuoc o) voi hinh -> khong can tao lai
    Dim a As Shape, z As Shape, tol As Double
    tol = 0.5
    Set a = FindShape(ws, GRID_PREFIX & "R1_C1")
    If a Is Nothing Then Exit Function
    Set z = FindShape(ws, GRID_PREFIX & "R" & nR & "_C" & nC)
    If z Is Nothing Then Exit Function
    If Not FindShape(ws, GRID_PREFIX & "R" & (nR + 1) & "_C1") Is Nothing Then Exit Function
    If Not FindShape(ws, GRID_PREFIX & "R1_C" & (nC + 1)) Is Nothing Then Exit Function
    If Abs(a.Left - img.Left) > tol Or Abs(a.Top - img.Top) > tol Then Exit Function
    If Abs(a.Width - img.Width / nC) > tol Or Abs(a.Height - img.Height / nR) > tol Then Exit Function
    If Abs((z.Left + z.Width) - (img.Left + img.Width)) > tol Then Exit Function
    If Abs((z.Top + z.Height) - (img.Top + img.Height)) > tol Then Exit Function
    CanReuseGrid = True
End Function

Private Sub ResetModelState(ByVal ws As Worksheet)
    'Xoa Grid, an anh, dua trang thai ve rong (dung khi khong co Model hop le)
    DeleteGridShapes
    HideAllImages
    Set gSelectedCells = CreateObject("Scripting.Dictionary")
    gModel = "": gRows = 0: gCols = 0
    Set mHeat = Nothing: mHeatMax = 0: mHeatActive = False
    ws.Range("A29:D29").ClearContents
    ws.Range("B4").Value = 0
    ws.Range("B5:B6").ClearContents
End Sub

Public Function LoadVehicleModel(ByVal modelName As String) As Boolean
    Dim ws As Worksheet, nR As Long, nC As Long, imgName As String, shp As Shape, g As Shape
    Dim prevEv As Boolean, wasProt As Boolean, t0 As Double, usingPlaceholder As Boolean
    Dim reused As Boolean, perfSet As Boolean, prevCalc As Long, prevPB As Boolean
    Dim hasSize As Boolean, imgW As Double, imgH As Double
    On Error GoTo EH
    mStep = "khoi dong"
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set ws = GetSheet(SH_MAP)
    t0 = Timer
    If ws.ProtectContents Then wasProt = True: ws.Unprotect   'bao ve sheet lam cham thao tac shape

    On Error Resume Next                    'tang toc: tat tinh toan va hien thi ngat trang trong luc xu ly
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    prevPB = ws.DisplayPageBreaks
    ws.DisplayPageBreaks = False
    perfSet = True
    Err.Clear
    On Error GoTo EH

    If Len(Trim$(modelName)) = 0 Then
        mStep = "xoa Grid (khong chon Model)"
        ResetModelState ws
        GoTo CLEANUP
    End If

    mStep = "doc tblModels"
    If Not GetModelInfo(modelName, nR, nC, imgName) Then
        ResetModelState ws
        MsgBox "Model '" & modelName & "' khong co trong tblModels, khong Active, hoac Rows/Cols khong hop le.", _
               vbExclamation, "Loi Model"
        GoTo CLEANUP
    End If

    mStep = "tim hinh " & imgName
    Set shp = FindImageShape(ws, imgName)
    If shp Is Nothing Then
        ResetModelState ws
        MsgBox "Khong tim thay hinh Master '" & imgName & "' cho Model '" & modelName & "'." & vbCrLf & _
               "Hay chen hinh vao sheet MAP_VEHICLE va dat Name = " & imgName & " (Name Box)." & vbCrLf & _
               "Grid KHONG duoc tao.", vbCritical, "Thieu hinh Master"
        GoTo CLEANUP
    End If

    mStep = "an anh cu"
    HideAllImages

    '4: hien thi anh moi (cac thuoc tinh phu bo qua loi neu loai hinh khong ho tro)
    mStep = "dat kich thuoc/vi tri hinh " & imgName
    hasSize = GetImageSize(imgW, imgH)       'kich thuoc chung cho moi Model (MASTER > tblSettings)
    On Error Resume Next
    If hasSize Then shp.LockAspectRatio = msoFalse Else shp.LockAspectRatio = msoTrue
    Err.Clear
    On Error GoTo EH
    If hasSize Then
        If Abs(shp.Width - imgW) > 0.01 Then shp.Width = imgW
        If Abs(shp.Height - imgH) > 0.01 Then shp.Height = imgH
    ElseIf IMG_MAX_WIDTH > 0 Then
        If shp.Width > IMG_MAX_WIDTH Then shp.Width = IMG_MAX_WIDTH   'Grid chia theo dung kich thuoc hinh
    End If
    shp.Left = ws.Range(ANCHOR_IMG).Left
    shp.Top = ws.Range(ANCHOR_IMG).Top
    On Error Resume Next
    shp.Placement = xlFreeFloating
    Err.Clear
    shp.Visible = msoTrue
    Err.Clear
    shp.ZOrder msoSendToBack
    Err.Clear
    usingPlaceholder = (shp.Type = msoAutoShape)
    Err.Clear
    On Error GoTo EH

    '5: Grid - dung lai neu khop hoan toan, nguoc lai xoa va tao moi
    mStep = "kiem tra Grid hien co"
    reused = CanReuseGrid(ws, nR, nC, shp)
    If reused Then
        mStep = "dung lai Grid (xoa chon / heatmap)"
        ClearSelection
        ClearHeatDisplay
        Application.ScreenUpdating = False
    Else
        mStep = "xoa Grid cu"
        DeleteGridShapes
    End If

    Set gSelectedCells = CreateObject("Scripting.Dictionary")
    Set mHeat = Nothing: mHeatMax = 0: mHeatActive = False
    gModel = "": gRows = 0: gCols = 0
    ws.Range("A29:D29").ClearContents
    ws.Range("B4").Value = 0
    Application.StatusBar = False

    If reused Then
        mStep = "ghi nhan Model cho Grid dung lai"
        Set g = FindShape(ws, GRID_PREFIX & "R1_C1")
        g.AlternativeText = modelName
        mGridRows = nR: mGridCols = nC
    Else
        mStep = "tao Grid"
        CreateVehicleGrid modelName, nR, nC, shp
    End If

    '6: cap nhat Rows/Cols
    mStep = "ghi Rows/Cols"
    gModel = modelName: gRows = nR: gCols = nC
    ws.Range("B5").Value = nR
    ws.Range("B6").Value = nC
    LoadVehicleModel = True
    Application.StatusBar = "Da tai Model " & modelName & " (" & nR & "x" & nC & ") trong " & Format$(Timer - t0, "0.0") & " giay" & _
                            IIf(reused, " - dung lai Grid", "")

CLEANUP:
    On Error Resume Next
    If perfSet Then
        ws.DisplayPageBreaks = prevPB
        Application.Calculation = prevCalc
    End If
    If wasProt Then ProtectMap ws
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    If usingPlaceholder Then
        MsgBox "Dang hien thi hinh TAM (o xam) cho Model '" & modelName & "'." & vbCrLf & _
               "Hay chen hinh Master that, dat Name = " & imgName & " (Name Box). Code se tu dung hinh that va bo hinh tam.", _
               vbExclamation, "Chua co hinh Master"
    End If
    Exit Function
EH:
    MsgBox "LoadVehicleModel loi: " & Err.Description & vbCrLf & "Buoc: " & mStep & vbCrLf & "Ma loi: " & Err.Number, vbCritical
    Resume CLEANUP
End Function

Public Sub HideAllImages()
    'An tat ca hinh Master theo cot ImageShapeName cua tblModels (tra theo ten, nhanh)
    Dim ws As Worksheet, lo As ListObject, dat As Variant, i As Long, cI As Long, nm As String, s As Shape
    Set ws = GetSheet(SH_MAP)
    Set lo = GetTable(SH_MASTER, TBL_MODELS)
    If Not TableHasData(lo) Then Exit Sub
    cI = ColIdx(lo, "ImageShapeName")
    dat = lo.DataBodyRange.Value2
    For i = 1 To UBound(dat, 1)
        nm = Trim$(CStr(dat(i, cI)))
        If Len(nm) > 0 Then
            Set s = FindShape(ws, nm)
            If Not s Is Nothing Then s.Visible = msoFalse
        End If
    Next i
End Sub

Public Sub DeleteGridShapes()
    'Duong nhanh: xoa ca ShapeRange bang 1 lenh. Duong cham (du phong): duyet tung shape.
    Dim ws As Worksheet, i As Long, r As Long, c As Long, k As Long
    Dim nR As Long, nC As Long, nm() As Variant
    Set ws = GetSheet(SH_MAP)

    nR = mGridRows: nC = mGridCols
    If nR < 1 Or nC < 1 Then                     'mat trang thai (vua mo file) -> suy ra tu B5/B6
        nR = CLng(Val(CStr(ws.Range("B5").Value)))
        nC = CLng(Val(CStr(ws.Range("B6").Value)))
    End If

    If nR >= 1 And nC >= 1 Then
        If Not FindShape(ws, GRID_PREFIX & "R" & nR & "_C" & nC) Is Nothing Then
            ReDim nm(0 To nR * nC - 1)
            For r = 1 To nR
                For c = 1 To nC
                    nm(k) = GRID_PREFIX & "R" & r & "_C" & c
                    k = k + 1
                Next c
            Next r
            On Error Resume Next
            ws.Shapes.Range(nm).Delete
            Err.Clear
            On Error GoTo 0
        End If
    End If
    mGridRows = 0: mGridCols = 0

    'Du phong: con sot shape Grid (khong khop kich thuoc) -> duyet tung shape
    If Not FindShape(ws, GRID_PREFIX & "R1_C1") Is Nothing Then
        For i = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(i).Name, Len(GRID_PREFIX)) = GRID_PREFIX Then ws.Shapes(i).Delete
        Next i
    End If
End Sub

Public Sub CreateVehicleGrid(ByVal modelName As String, ByVal nR As Long, ByVal nC As Long, ByVal img As Shape)
    'Tao o bang AddShape + Name, sau do dinh dang hang loat (nhanh). Neu dinh dang hang loat loi
    'hoac khong an -> tu dong chuyen sang dinh dang tung o (cham hon nhung chac chan).
    Dim ws As Worksheet, r As Long, c As Long, k As Long, s As Shape
    Dim cw As Double, ch As Double, nm() As Variant, okGroup As Boolean
    Set ws = GetSheet(SH_MAP)
    mStep = "xoa Grid cu (trong CreateVehicleGrid)"
    DeleteGridShapes                       'dam bao khong trung shape
    cw = img.Width / nC
    ch = img.Height / nR
    mStep = "tao o Grid"
    ReDim nm(0 To nR * nC - 1)
    For r = 1 To nR
        For c = 1 To nC
            Set s = ws.Shapes.AddShape(msoShapeRectangle, img.Left + (c - 1) * cw, img.Top + (r - 1) * ch, cw, ch)
            nm(k) = GRID_PREFIX & "R" & r & "_C" & c
            s.Name = CStr(nm(k))
            k = k + 1
        Next c
    Next r

    mStep = "dinh dang hang loat"
    okGroup = FormatGridGroup(ws, nm, modelName)
    If okGroup Then
        On Error Resume Next
        okGroup = (ws.Shapes(CStr(nm(0))).Fill.Transparency > 0.9)
        If Err.Number <> 0 Then okGroup = False: Err.Clear
        On Error GoTo 0
    End If

    If Not okGroup Then
        mStep = "dinh dang tung o"
        For k = 0 To UBound(nm)
            Set s = ws.Shapes(CStr(nm(k)))
            s.AlternativeText = modelName
            On Error Resume Next
            s.Placement = xlFreeFloating
            s.Shadow.Visible = msoFalse
            Err.Clear
            On Error GoTo 0
            s.OnAction = "modVehicleMapping.GridCell_Click"
            SetCellHighlight s, False
        Next k
    End If
    mGridRows = nR: mGridCols = nC
End Sub

Private Function FormatGridGroup(ByVal ws As Worksheet, ByRef nm() As Variant, ByVal modelName As String) As Boolean
    Dim sr As ShapeRange
    On Error GoTo Fail
    Set sr = ws.Shapes.Range(nm)
    sr.AlternativeText = modelName          'ghi nho Model cua grid
    sr.Placement = xlFreeFloating
    sr.Shadow.Visible = msoFalse
    sr.OnAction = "modVehicleMapping.GridCell_Click"
    sr.Fill.Visible = msoTrue
    sr.Fill.Solid
    sr.Fill.ForeColor.RGB = RGB(255, 255, 255)   'mac dinh = SetCellHighlight(shp, False)
    sr.Fill.Transparency = 0.97
    sr.Line.ForeColor.RGB = RGB(110, 110, 110)
    sr.Line.Weight = 0.5
    sr.Line.Transparency = 0.6
    FormatGridGroup = True
    Exit Function
Fail:
    FormatGridGroup = False
End Function

Public Sub SetCellHighlight(ByVal shp As Shape, ByVal isOn As Boolean)
    Dim r As Long, c As Long, hasHeat As Boolean
    If Not isOn Then
        If Not mHeat Is Nothing Then
            If GetGridCoordinates(shp.Name, r, c) Then hasHeat = mHeat.Exists(r & "," & c)
        End If
    End If
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        If isOn Then
            .Fill.ForeColor.RGB = COLOR_SELECTED
            .Fill.Transparency = 0.55
            .Line.ForeColor.RGB = RGB(0, 120, 50)
            .Line.Weight = 1.25
            .Line.Transparency = 0.2
        ElseIf hasHeat Then
            SetHeatFill shp, CDbl(mHeat(r & "," & c))
        Else
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Fill.Transparency = 0.97          'gan nhu trong suot nhung van nhan click
            .Line.ForeColor.RGB = RGB(110, 110, 110)
            .Line.Weight = 0.5
            .Line.Transparency = 0.6
        End If
    End With
End Sub

'=====================================================================
' 4. CHON O LOI
'=====================================================================
Public Sub EnsureSelection()
    'Khoi tao dictionary; neu mat (VBA reset) thi dung lai tu mau fill cua cac shape
    Dim ws As Worksheet, shp As Shape, r As Long, c As Long
    If Not gSelectedCells Is Nothing Then Exit Sub
    Set gSelectedCells = CreateObject("Scripting.Dictionary")
    Set ws = GetSheet(SH_MAP)
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(GRID_PREFIX)) = GRID_PREFIX Then
            If shp.Fill.ForeColor.RGB = COLOR_SELECTED Then
                If GetGridCoordinates(shp.Name, r, c) Then gSelectedCells(r & "," & c) = True
            End If
        End If
    Next shp
End Sub

Public Sub GridCell_Click()
    'Gan cho OnAction cua moi o Grid
    Dim nm As String, r As Long, c As Long
    On Error GoTo EH
    CloseCalendar
    If mHeatActive Then
        Application.StatusBar = "Dang o che do XEM DU LIEU - bam HIDE DATA de chon o loi."
        Exit Sub
    End If
    nm = CStr(Application.Caller)
    If Not GetGridCoordinates(nm, r, c) Then Exit Sub
    ToggleGridCell r, c
    Exit Sub
EH:
    MsgBox "Loi khi chon o: " & Err.Description, vbExclamation
End Sub

Public Sub ToggleGridCell(ByVal r As Long, ByVal c As Long)
    Dim ws As Worksheet, shp As Shape, key As String
    Set ws = GetSheet(SH_MAP)
    EnsureSelection
    Set shp = FindShape(ws, GRID_PREFIX & "R" & r & "_C" & c)
    If shp Is Nothing Then Exit Sub
    key = r & "," & c
    If gSelectedCells.Exists(key) Then
        gSelectedCells.Remove key
        SetCellHighlight shp, False
    Else
        gSelectedCells.Add key, True
        SetCellHighlight shp, True
    End If
    UpdateSelCount
End Sub

Private Sub UpdateSelCount()
    Dim ws As Worksheet, prevEv As Boolean
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Set ws = GetSheet(SH_MAP)
    ws.Range("B4").Value = gSelectedCells.Count
    Application.EnableEvents = prevEv
End Sub

Public Sub ClearSelection()
    Dim ws As Worksheet, k As Variant, p() As String, shp As Shape, prevEv As Boolean
    On Error GoTo EH
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set ws = GetSheet(SH_MAP)
    EnsureSelection
    For Each k In gSelectedCells.Keys
        p = Split(CStr(k), ",")
        Set shp = FindShape(ws, GRID_PREFIX & "R" & p(0) & "_C" & p(1))
        If Not shp Is Nothing Then SetCellHighlight shp, False
    Next k
    gSelectedCells.RemoveAll
    ws.Range("B4").Value = 0
CLEANUP:
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    Exit Sub
EH:
    MsgBox "ClearSelection loi: " & Err.Description, vbExclamation
    Resume CLEANUP
End Sub

'=====================================================================
' 5. VIN
'=====================================================================
Public Sub NormalizeVIN()
    'Goi tu Worksheet_Change khi B3 thay doi. KHONG tu dong Save.
    Dim ws As Worksheet, c As Range, v As Variant, s As String, cleaned As String, prevEv As Boolean
    On Error GoTo EH
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Set ws = GetSheet(SH_MAP)
    Set c = ws.Range("B3")
    v = c.Value
    If IsEmpty(v) Then
        gVinJustEntered = False
        GoTo CLEANUP
    End If
    If VarType(v) <> vbString Then
        'Gia tri bi doi thanh so -> da mat so 0 dau
        c.NumberFormat = "@"
        c.ClearContents
        gVinJustEntered = False
        MsgBox "VIN bi nhan dang la SO (mat so 0 dau). O B3 da duoc dat lai dinh dang Text. Hay quet/nhap lai VIN.", vbExclamation
        GoTo CLEANUP
    End If
    s = CStr(v)
    cleaned = Replace(Replace(s, vbCr, ""), vbLf, "")
    If cleaned <> s Then c.Value = cleaned
    gVinJustEntered = (Len(cleaned) > 0)
CLEANUP:
    Application.EnableEvents = prevEv
    Exit Sub
EH:
    gVinJustEntered = False
    Resume CLEANUP
End Sub

'=====================================================================
' 6. VALIDATE + SAVE
'=====================================================================
Public Function ValidateInput(ByRef msg As String) As Boolean
    Dim ws As Worksheet, model As String, nR As Long, nC As Long, img As String
    Dim loD As ListObject, dat As Variant, i As Long, cT As Long, cA As Long
    Dim defect As String, okDef As Boolean, v As Variant, bad As Boolean

    ValidateInput = False
    Set ws = GetSheet(SH_MAP)

    '1. Model
    model = Trim$(CStr(ws.Range("B2").Value))
    If Len(model) = 0 Then msg = "Chua chon Model (o B2).": Exit Function
    If Not GetModelInfo(model, nR, nC, img) Then
        msg = "Model '" & model & "' khong hop le hoac khong Active trong tblModels.": Exit Function
    End If

    '2. VIN
    If Len(Trim$(CStr(ws.Range("B3").Value))) = 0 Then
        msg = "VIN dang trong (o B3). Hay quet barcode hoac nhap VIN.": Exit Function
    End If

    '3. Cell da chon
    EnsureSelection
    If gSelectedCells.Count = 0 Then msg = "Chua chon o loi nao tren hinh Master.": Exit Function

    '4. Defect Type
    defect = Trim$(CStr(ws.Range("D5").Value))
    If Len(defect) = 0 Then msg = "Chua chon Defect Type (o D5).": Exit Function
    Set loD = GetTable(SH_MASTER, TBL_DEFECTTYPES)
    If TableHasData(loD) Then
        cT = ColIdx(loD, "DefectType"): cA = ColIdx(loD, "Active")
        dat = loD.DataBodyRange.Value2
        For i = 1 To UBound(dat, 1)
            If StrComp(Trim$(CStr(dat(i, cT))), defect, vbTextCompare) = 0 And IsActiveValue(dat(i, cA)) Then
                okDef = True: Exit For
            End If
        Next i
    End If
    If Not okDef Then msg = "Defect Type '" & defect & "' khong co trong tblDefectTypes hoac khong Active.": Exit Function

    '5. Qty
    v = ws.Range("D6").Value
    If IsError(v) Or IsEmpty(v) Then
        bad = True
    ElseIf Not IsNumeric(v) Then
        bad = True
    ElseIf CDbl(v) < 1 Or CDbl(v) <> Int(CDbl(v)) Then
        bad = True
    End If
    If bad Then msg = "Qty per Cell (o D6) phai la so nguyen duong.": Exit Function

    '6. Inspector
    If Len(Trim$(CStr(ws.Range("D4").Value))) = 0 Then msg = "Chua nhap Inspector (o D4).": Exit Function

    '7. Grid dung Model
    If StrComp(GetGridModel(ws), model, vbTextCompare) <> 0 Or _
       FindShape(ws, GRID_PREFIX & "R" & nR & "_C" & nC) Is Nothing Then
        msg = "Grid hien tai khong thuoc Model '" & model & "'. Hay chon lai Model o B2.": Exit Function
    End If

    gModel = model: gRows = nR: gCols = nC
    ValidateInput = True
End Function

Public Sub SaveDefectRecords()
    Dim ws As Worksheet, loDB As ListObject, msg As String
    Dim model As String, vin As String, defect As String, inspector As String, desc As String
    Dim qty As Long, saveTime As Date, saveShift As String, saveShiftDate As Date
    Dim batchID As String, n As Long, r As Long, c As Long, i As Long, k As Long
    Dim arr() As Variant, ix(1 To 17) As Long, hdr As Variant
    Dim nOld As Long, tgt As Range, totalQty As Long, sumErr As String, prevEv As Boolean
    Dim viewName As String

    On Error GoTo EH
    Set ws = GetSheet(SH_MAP)
    If Not ValidateInput(msg) Then MsgBox msg, vbExclamation, "Khong the luu": Exit Sub

    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Set loDB = GetTable(SH_DB, TBL_DEFECTS)
    hdr = Array("DefectID", "BatchID", "DateTime", "ShiftDate", "Model", "VIN", "View", "X", "Y", _
                "CellID", "GridID", "DefectType", "Qty", "Shift", "Inspector", "Description", "Status")
    For k = 0 To 16
        ix(k + 1) = ColIdx(loDB, CStr(hdr(k)))
    Next k

    model = Trim$(CStr(ws.Range("B2").Value))
    vin = CStr(ws.Range("B3").Value)
    defect = Trim$(CStr(ws.Range("D5").Value))
    qty = CLng(ws.Range("D6").Value)
    inspector = Trim$(CStr(ws.Range("D4").Value))
    desc = CStr(ws.Range("D7").Value)

    'Thoi diem Save: tinh MOT LAN, dung chung cho ca Batch
    saveTime = Now
    GetShiftInfo saveTime, saveShift, saveShiftDate
    batchID = NewID("B")

    n = gSelectedCells.Count
    ReDim arr(1 To n, 1 To loDB.ListColumns.Count)
    i = 0
    For r = 1 To gRows
        For c = 1 To gCols
            If gSelectedCells.Exists(r & "," & c) Then
                i = i + 1
                viewName = GetViewForCell(model, c, r)
                arr(i, ix(1)) = batchID & "-" & Format$(i, "000")
                arr(i, ix(2)) = batchID
                arr(i, ix(3)) = saveTime
                arr(i, ix(4)) = saveShiftDate
                arr(i, ix(5)) = model
                arr(i, ix(6)) = vin
                arr(i, ix(7)) = viewName
                arr(i, ix(8)) = c                       'X = cot
                arr(i, ix(9)) = r                       'Y = hang
                arr(i, ix(10)) = ColumnLetter(c) & r    'CellID dang Excel
                arr(i, ix(11)) = "R" & r & "C" & c      'GridID
                arr(i, ix(12)) = defect
                arr(i, ix(13)) = qty
                arr(i, ix(14)) = saveShift
                arr(i, ix(15)) = inspector
                arr(i, ix(16)) = desc
                arr(i, ix(17)) = "Open"
            End If
        Next c
    Next r
    If i <> n Then
        MsgBox "Danh sach o chon khong dong bo voi Grid. Hay bam CLEAR SELECTION va chon lai.", vbExclamation
        GoTo CLEANUP
    End If

    'Ghi vao tblDefects (khong xoa du lieu cu)
    nOld = 0
    If TableHasData(loDB) Then nOld = loDB.DataBodyRange.Rows.Count
    loDB.Resize loDB.Range.Resize(nOld + n + 1)
    Set tgt = loDB.DataBodyRange.Rows(nOld + 1).Resize(n)
    tgt.Columns(ix(3)).NumberFormat = "dd/mm/yyyy hh:mm:ss"
    tgt.Columns(ix(4)).NumberFormat = "dd/mm/yyyy"
    tgt.Columns(ix(6)).NumberFormat = "@"       'VIN luon la Text
    tgt.Columns(ix(10)).NumberFormat = "@"
    tgt.Columns(ix(11)).NumberFormat = "@"
    tgt.Value = arr

    totalQty = n * qty

    On Error Resume Next
    RefreshSummary
    If Err.Number <> 0 Then sumErr = Err.Description: Err.Clear
    On Error GoTo EH

    ClearSelection
    ResetInputFields False
    If mHeatActive Then ShowDatabaseOnGrid True
    UpdateShiftDisplay
    Application.ScreenUpdating = True
    On Error Resume Next
    ws.Range("B3").Select
    On Error GoTo EH

    msg = "LUU THANH CONG" & vbCrLf & vbCrLf & _
          "BatchID: " & batchID & vbCrLf & _
          "So ban ghi: " & n & vbCrLf & _
          "Tong Qty: " & totalQty & vbCrLf & _
          "Ca: " & saveShift & "   ShiftDate: " & Format$(saveShiftDate, "dd/mm/yyyy")
    If Len(sumErr) > 0 Then msg = msg & vbCrLf & vbCrLf & "Luu y: Refresh SUMMARY loi: " & sumErr
    MsgBox msg, vbInformation, "Vehicle Defect Mapping"

CLEANUP:
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    Exit Sub
EH:
    MsgBox "SaveDefectRecords loi: " & Err.Description, vbCritical
    Resume CLEANUP
End Sub

Public Function GetViewForCell(ByVal modelName As String, ByVal x As Long, ByVal y As Long) As String
    'Tra ve LH/RH/FRONT/REAR/TOP theo tblGridZones; khong co vung -> "MASTER"
    Dim lo As ListObject, dat As Variant, i As Long
    Dim cM As Long, cV As Long, cXF As Long, cXT As Long, cYF As Long, cYT As Long
    GetViewForCell = "MASTER"
    On Error GoTo EH
    Set lo = GetTable(SH_MASTER, TBL_ZONES)
    If Not TableHasData(lo) Then Exit Function
    cM = ColIdx(lo, "Model"): cV = ColIdx(lo, "View")
    cXF = ColIdx(lo, "XFrom"): cXT = ColIdx(lo, "XTo")
    cYF = ColIdx(lo, "YFrom"): cYT = ColIdx(lo, "YTo")
    dat = lo.DataBodyRange.Value2
    For i = 1 To UBound(dat, 1)
        If StrComp(Trim$(CStr(dat(i, cM))), modelName, vbTextCompare) = 0 Then
            If IsNum(dat(i, cXF)) And IsNum(dat(i, cXT)) And IsNum(dat(i, cYF)) And IsNum(dat(i, cYT)) Then
                If x >= CDbl(dat(i, cXF)) And x <= CDbl(dat(i, cXT)) And _
                   y >= CDbl(dat(i, cYF)) And y <= CDbl(dat(i, cYT)) Then
                    If Len(Trim$(CStr(dat(i, cV)))) > 0 Then GetViewForCell = UCase$(Trim$(CStr(dat(i, cV))))
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    GetViewForCell = "MASTER"
End Function

'=====================================================================
' 7. RESET INPUT
'=====================================================================
Public Sub ResetInputFields(Optional ByVal clearInspector As Boolean = False)
    Dim ws As Worksheet, prevEv As Boolean
    On Error GoTo EH
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Set ws = GetSheet(SH_MAP)
    ws.Range("B3").ClearContents
    ws.Range("D5").ClearContents
    ws.Range("D6").Value = 1
    ws.Range("D7").ClearContents
    If clearInspector Then ws.Range("D4").ClearContents
    gVinJustEntered = False
CLEANUP:
    Application.EnableEvents = prevEv
    Exit Sub
EH:
    MsgBox "ResetInputFields loi: " & Err.Description, vbExclamation
    Resume CLEANUP
End Sub

'Ham gan cho nut bam
Public Sub BtnSave()
    SaveDefectRecords
End Sub

Public Sub BtnClear()
    ClearSelection
End Sub

Public Sub BtnReset()
    ResetInputFields False
    On Error Resume Next
    GetSheet(SH_MAP).Range("B3").Select
End Sub

'=====================================================================
' 7b. HIEN THI DU LIEU DATABASE LEN GRID (HEATMAP)
'=====================================================================
Private Sub SetHeatFill(ByVal shp As Shape, ByVal v As Double)
    Dim tt As Double
    If mHeatMax > 0 Then tt = v / mHeatMax
    If tt > 1 Then tt = 1
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255 - CLng(35 * tt), 235 - CLng(205 * tt), 60 - CLng(30 * tt))   'vang -> do
        .Fill.Transparency = 0.3
        .Line.ForeColor.RGB = RGB(150, 40, 40)
        .Line.Weight = 0.75
        .Line.Transparency = 0.3
    End With
End Sub

Private Sub SetHeatText(ByVal shp As Shape, ByVal txt As String)
    With shp.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = txt
        If Len(txt) > 0 Then
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Size = 8
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        End If
    End With
End Sub

Public Sub ClearHeatDisplay()
    'Xoa heatmap: chi cham vao cac o co du lieu (nhanh)
    Dim ws As Worksheet, old As Object, k As Variant, p() As String, shp As Shape
    If Not mHeat Is Nothing Then
        Set ws = GetSheet(SH_MAP)
        EnsureSelection
        Set old = mHeat
        Set mHeat = Nothing                    'de SetCellHighlight ve lai mau mac dinh
        For Each k In old.Keys
            p = Split(CStr(k), ",")
            Set shp = FindShape(ws, GRID_PREFIX & "R" & p(0) & "_C" & p(1))
            If Not shp Is Nothing Then
                SetHeatText shp, ""
                SetCellHighlight shp, gSelectedCells.Exists(CStr(k))
            End If
        Next k
    End If
    mHeatMax = 0
    mHeatActive = False
End Sub

Public Sub ShowDatabaseOnGrid(Optional ByVal silent As Boolean = False)
    Dim ws As Worksheet, lo As ListObject, dat As Variant, i As Long, tmp As Variant
    Dim cMod As Long, cX As Long, cY As Long, cQty As Long, cSD As Long, cSh As Long, cVin As Long, cDef As Long
    Dim model As String, fVin As String, fShift As String, fDef As String, useRecords As Boolean
    Dim hasFrom As Boolean, hasTo As Boolean, dFrom As Long, dTo As Long, sd As Long
    Dim agg As Object, key As String, r As Long, c As Long, v As Double, nR As Long, nC As Long
    Dim k As Variant, p() As String, shp As Shape, nSkip As Long, totalV As Double, maxV As Double
    Dim prevSU As Boolean, st As String

    On Error GoTo EH
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = GetSheet(SH_MAP)

    CloseCalendar
    model = Trim$(CStr(ws.Range("B2").Value))
    If Len(model) = 0 Then
        If Not silent Then MsgBox "Chua chon Model (o B2).", vbExclamation, "Xem du lieu"
        GoTo CLEANUP
    End If
    If StrComp(GetGridModel(ws), model, vbTextCompare) <> 0 Then
        If Not silent Then MsgBox "Grid hien tai khong thuoc Model '" & model & "'. Hay chon lai Model o B2.", vbExclamation, "Xem du lieu"
        GoTo CLEANUP
    End If
    nR = gRows: nC = gCols
    If nR < 1 Then nR = CLng(Val(CStr(ws.Range("B5").Value)))
    If nC < 1 Then nC = CLng(Val(CStr(ws.Range("B6").Value)))

    'Doc bo loc
    tmp = ws.Range("B21").Value
    If Len(Trim$(CStr(tmp))) > 0 Then
        If Not IsDate(tmp) Then
            If Not silent Then MsgBox "From date (B21) khong phai ngay hop le.", vbExclamation, "Xem du lieu"
            GoTo CLEANUP
        End If
        hasFrom = True: dFrom = CLng(Int(CDbl(CDate(tmp))))
    End If
    tmp = ws.Range("B22").Value
    If Len(Trim$(CStr(tmp))) > 0 Then
        If Not IsDate(tmp) Then
            If Not silent Then MsgBox "To date (B22) khong phai ngay hop le.", vbExclamation, "Xem du lieu"
            GoTo CLEANUP
        End If
        hasTo = True: dTo = CLng(Int(CDbl(CDate(tmp))))
    End If
    If hasFrom And hasTo Then
        If dFrom > dTo Then
            If Not silent Then MsgBox "From date lon hon To date.", vbExclamation, "Xem du lieu"
            GoTo CLEANUP
        End If
    End If
    fVin = Trim$(CStr(ws.Range("B23").Value))
    fShift = UCase$(Trim$(CStr(ws.Range("B24").Value)))
    fDef = Trim$(CStr(ws.Range("B25").Value))
    useRecords = (LCase$(Trim$(CStr(ws.Range("B26").Value))) = "records")

    'Tong hop theo o (X = cot, Y = hang)
    Set agg = CreateObject("Scripting.Dictionary")
    Set lo = GetTable(SH_DB, TBL_DEFECTS)
    If TableHasData(lo) Then
        cMod = ColIdx(lo, "Model"): cX = ColIdx(lo, "X"): cY = ColIdx(lo, "Y"): cQty = ColIdx(lo, "Qty")
        cSD = ColIdx(lo, "ShiftDate"): cSh = ColIdx(lo, "Shift"): cVin = ColIdx(lo, "VIN"): cDef = ColIdx(lo, "DefectType")
        dat = lo.DataBodyRange.Value2
        For i = 1 To UBound(dat, 1)
            If StrComp(Trim$(CStr(dat(i, cMod))), model, vbTextCompare) = 0 Then
                If Len(fVin) > 0 Then
                    If StrComp(Trim$(CStr(dat(i, cVin))), fVin, vbTextCompare) <> 0 Then GoTo NextRow
                End If
                If Len(fShift) > 0 Then
                    If UCase$(Trim$(CStr(dat(i, cSh)))) <> fShift Then GoTo NextRow
                End If
                If Len(fDef) > 0 Then
                    If StrComp(Trim$(CStr(dat(i, cDef))), fDef, vbTextCompare) <> 0 Then GoTo NextRow
                End If
                If hasFrom Or hasTo Then
                    sd = 0
                    If IsNumeric(dat(i, cSD)) Then
                        sd = CLng(dat(i, cSD))
                    ElseIf IsDate(dat(i, cSD)) Then
                        sd = CLng(Int(CDbl(CDate(dat(i, cSD)))))
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
                If r < 1 Or r > nR Or c < 1 Or c > nC Then
                    nSkip = nSkip + 1
                    GoTo NextRow
                End If
                If useRecords Then
                    v = 1
                ElseIf IsNumeric(dat(i, cQty)) Then
                    v = CDbl(dat(i, cQty))
                Else
                    v = 0
                End If
                key = r & "," & c
                If agg.Exists(key) Then agg(key) = agg(key) + v Else agg.Add key, v
            End If
NextRow:
        Next i
    End If

    'Ve lai: xoa heatmap cu, roi to mau cac o co du lieu
    ClearHeatDisplay
    EnsureSelection
    For Each k In agg.Keys
        If agg(k) > maxV Then maxV = agg(k)
        totalV = totalV + agg(k)
    Next k
    Set mHeat = agg
    mHeatMax = maxV
    mHeatActive = True
    For Each k In agg.Keys
        p = Split(CStr(k), ",")
        Set shp = FindShape(ws, GRID_PREFIX & "R" & p(0) & "_C" & p(1))
        If Not shp Is Nothing Then
            SetHeatText shp, CStr(CLng(agg(k)))
            SetCellHighlight shp, gSelectedCells.Exists(CStr(k))
        End If
    Next k

    st = "Model " & model & ": " & agg.Count & " o co loi  |  Tong " & IIf(useRecords, "Records", "Qty") & " = " & CLng(totalV)
    If nSkip > 0 Then st = st & "  |  bo qua " & nSkip & " ban ghi ngoai Grid"
    If agg.Count = 0 Then st = "Khong co du lieu phu hop bo loc cho Model " & model
    ws.Range("A29").Value = st

CLEANUP:
    Application.ScreenUpdating = prevSU
    Exit Sub
EH:
    Application.ScreenUpdating = prevSU
    MsgBox "ShowDatabaseOnGrid loi: " & Err.Description, vbCritical
End Sub

Public Sub HideDatabaseOnGrid()
    On Error GoTo EH
    CloseCalendar
    Application.ScreenUpdating = False
    ClearHeatDisplay
    GetSheet(SH_MAP).Range("A29:D29").ClearContents
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "HideDatabaseOnGrid loi: " & Err.Description, vbExclamation
End Sub

'Ham gan cho nut bam
Public Sub BtnShowData()
    ShowDatabaseOnGrid False
End Sub

Public Sub BtnHideData()
    HideDatabaseOnGrid
End Sub

'=====================================================================
' 7c. LICH CHON NGAY - hien khi chon o B21 / B22 (dung Shape, KHONG dung UserForm)
'=====================================================================
Private Function CalShape(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, _
                          ByVal l As Double, ByVal tp As Double, ByVal w As Double, ByVal h As Double, _
                          ByVal fillClr As Long, ByVal fontClr As Long, ByVal isBold As Boolean, _
                          Optional ByVal action As String = "") As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRectangle, l, tp, w, h)
    s.Name = nm
    s.Fill.ForeColor.RGB = fillClr
    s.Line.ForeColor.RGB = RGB(190, 190, 190)
    s.Line.Weight = 0.5
    On Error Resume Next
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    Err.Clear
    On Error GoTo 0
    If Len(action) > 0 Then s.OnAction = action
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = txt
        If Len(txt) > 0 Then
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Size = 9
            .TextRange.Font.Bold = IIf(isBold, msoTrue, msoFalse)
            .TextRange.Font.Fill.ForeColor.RGB = fontClr
        End If
    End With
    mCalCol.Add nm
    Set CalShape = s
End Function

Private Sub DeleteCalendarShapes()
    Dim ws As Worksheet, nm() As Variant, i As Long
    Set ws = GetSheet(SH_MAP)
    If Not mCalCol Is Nothing Then
        If mCalCol.Count > 0 Then
            ReDim nm(0 To mCalCol.Count - 1)
            For i = 1 To mCalCol.Count
                nm(i - 1) = CStr(mCalCol(i))
            Next i
            On Error Resume Next
            ws.Shapes.Range(nm).Delete
            Err.Clear
            On Error GoTo 0
        End If
    End If
    Set mCalCol = Nothing
    If Not FindShape(ws, "CAL_BG") Is Nothing Then      'con sot (vd: mat trang thai) -> duyet tung shape
        For i = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(i).Name, 4) = "CAL_" Then ws.Shapes(i).Delete
        Next i
    End If
End Sub

Public Sub ForceCloseCalendar()
    On Error Resume Next
    DeleteCalendarShapes
    mCalOpen = False
End Sub

Public Sub CloseCalendar()
    If Not mCalOpen Then Exit Sub
    On Error Resume Next
    DeleteCalendarShapes
    mCalOpen = False
End Sub

Private Sub DrawCalendar()
    Dim ws As Worksheet, d As Long, nDays As Long, idx As Long, wd As Long, i As Long
    Dim L0 As Double, T0 As Double, cw As Double, rh As Double, pad As Double
    Dim w As Double, hTotal As Double, yTop As Double, yBot As Double, bw As Double
    Dim cur As Variant, curSerial As Long, todaySerial As Long, serial As Long, fillClr As Long
    Dim wdNames As Variant, s As Shape

    Set ws = GetSheet(SH_MAP)
    DeleteCalendarShapes
    Set mCalCol = New Collection

    cw = 34: rh = 24: pad = 6
    w = 7 * cw + 2 * pad
    hTotal = pad + 26 + 20 + 6 * rh + 26 + pad
    L0 = ws.Range("C21").Left
    T0 = ws.Range(mCalTarget).Top
    yTop = T0 + pad
    yBot = yTop + 46 + 6 * rh

    cur = ws.Range(mCalTarget).Value
    If IsDate(cur) Then curSerial = CLng(Int(CDbl(CDate(cur))))
    todaySerial = CLng(Int(CDbl(Date)))

    Set s = CalShape(ws, "CAL_BG", "", L0, T0, w, hTotal, RGB(245, 245, 245), 0, False)
    s.Line.ForeColor.RGB = RGB(31, 56, 100)
    s.Line.Weight = 1.5

    CalShape ws, "CAL_NAV_PREV", "<", L0 + pad, yTop, cw, 26, RGB(31, 78, 121), RGB(255, 255, 255), True, "modVehicleMapping.Cal_Nav"
    CalShape ws, "CAL_TITLE", "Thang " & mCalMonth & "/" & mCalYear, L0 + pad + cw, yTop, 5 * cw, 26, RGB(31, 56, 100), RGB(255, 255, 255), True
    CalShape ws, "CAL_NAV_NEXT", ">", L0 + pad + 6 * cw, yTop, cw, 26, RGB(31, 78, 121), RGB(255, 255, 255), True, "modVehicleMapping.Cal_Nav"

    wdNames = Array("T2", "T3", "T4", "T5", "T6", "T7", "CN")
    For i = 0 To 6
        If i = 6 Then fillClr = RGB(192, 0, 0) Else fillClr = RGB(0, 0, 0)
        CalShape ws, "CAL_WD" & (i + 1), CStr(wdNames(i)), L0 + pad + i * cw, yTop + 26, cw, 20, RGB(217, 217, 217), fillClr, True
    Next i

    nDays = Day(DateSerial(mCalYear, mCalMonth + 1, 0))
    wd = Weekday(DateSerial(mCalYear, mCalMonth, 1), vbMonday) - 1      '0 = thu hai
    For d = 1 To nDays
        idx = wd + d - 1
        serial = CLng(DateSerial(mCalYear, mCalMonth, d))
        fillClr = RGB(255, 255, 255)
        If serial = todaySerial Then fillClr = RGB(255, 230, 153)
        If serial = curSerial Then fillClr = RGB(146, 208, 80)
        CalShape ws, "CAL_D" & d, CStr(d), L0 + pad + (idx Mod 7) * cw, yTop + 46 + (idx \ 7) * rh, cw, rh, _
                 fillClr, RGB(0, 0, 0), (serial = curSerial), "modVehicleMapping.Cal_DayClick"
    Next d

    bw = 7 * cw / 3
    CalShape ws, "CAL_BTN_TODAY", "Hom nay", L0 + pad, yBot, bw, 26, RGB(0, 140, 70), RGB(255, 255, 255), True, "modVehicleMapping.Cal_Today"
    CalShape ws, "CAL_BTN_CLEAR", "Xoa ngay", L0 + pad + bw, yBot, bw, 26, RGB(120, 120, 120), RGB(255, 255, 255), True, "modVehicleMapping.Cal_Clear"
    CalShape ws, "CAL_BTN_CLOSE", "Dong", L0 + pad + 2 * bw, yBot, bw, 26, RGB(160, 30, 30), RGB(255, 255, 255), True, "modVehicleMapping.Cal_Close"
End Sub

Public Sub ShowCalendar(ByVal targetAddr As String)
    'Goi tu Worksheet_SelectionChange khi chon B21 hoac B22
    Dim ws As Worksheet, v As Variant, dt As Date, prevSU As Boolean
    On Error GoTo EH
    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set ws = GetSheet(SH_MAP)
    mCalTarget = targetAddr
    v = ws.Range(targetAddr).Value
    If IsDate(v) Then
        dt = CDate(v)
    ElseIf targetAddr = "B22" And IsDate(ws.Range("B21").Value) Then
        dt = CDate(ws.Range("B21").Value)       'To date trong: mo o thang cua From date
    Else
        dt = Date
    End If
    mCalYear = Year(dt): mCalMonth = Month(dt)
    DrawCalendar
    mCalOpen = True
    Application.ScreenUpdating = prevSU
    Exit Sub
EH:
    Application.ScreenUpdating = prevSU
    mCalOpen = True                             'de CloseCalendar don dep phan da tao
    CloseCalendar
    MsgBox "Lich chon ngay loi: " & Err.Description & vbCrLf & "Co the nhap ngay truc tiep vao o (dd/mm/yyyy).", vbExclamation
End Sub

Private Sub CalPick(ByVal dt As Date)
    Dim ws As Worksheet, tgt As String, prevEv As Boolean
    Set ws = GetSheet(SH_MAP)
    tgt = mCalTarget
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(tgt).NumberFormat = "dd/mm/yyyy"
    ws.Range(tgt).Value = dt
    Application.EnableEvents = prevEv
    CloseCalendar
    On Error Resume Next
    If tgt = "B21" Then ws.Range("B22").Select Else ws.Range("B23").Select   'From -> mo tiep lich To
End Sub

Public Sub Cal_DayClick()
    Dim nm As String
    On Error GoTo EH
    nm = CStr(Application.Caller)
    If Left$(nm, 5) <> "CAL_D" Then Exit Sub
    If Not IsNumeric(Mid$(nm, 6)) Then Exit Sub
    CalPick DateSerial(mCalYear, mCalMonth, CLng(Mid$(nm, 6)))
    Exit Sub
EH:
    MsgBox "Chon ngay loi: " & Err.Description, vbExclamation
End Sub

Public Sub Cal_Today()
    On Error GoTo EH
    CalPick Date
    Exit Sub
EH:
    MsgBox "Chon ngay loi: " & Err.Description, vbExclamation
End Sub

Public Sub Cal_Clear()
    Dim ws As Worksheet
    On Error GoTo EH
    Set ws = GetSheet(SH_MAP)
    ws.Range(mCalTarget).ClearContents
    CloseCalendar
    Exit Sub
EH:
    MsgBox "Xoa ngay loi: " & Err.Description, vbExclamation
End Sub

Public Sub Cal_Close()
    CloseCalendar
End Sub

Public Sub Cal_Nav()
    Dim nm As String
    On Error GoTo EH
    nm = CStr(Application.Caller)
    If nm = "CAL_NAV_PREV" Then mCalMonth = mCalMonth - 1 Else mCalMonth = mCalMonth + 1
    If mCalMonth < 1 Then mCalMonth = 12: mCalYear = mCalYear - 1
    If mCalMonth > 12 Then mCalMonth = 1: mCalYear = mCalYear + 1
    Application.ScreenUpdating = False
    DrawCalendar
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "Chuyen thang loi: " & Err.Description, vbExclamation
End Sub

'=====================================================================
' 8. SUMMARY
'=====================================================================
Public Sub RefreshSummary()
    Dim loD As ListObject, loS As ListObject, ws As Worksheet, dat As Variant, dict As Object
    Dim i As Long, j As Long, m As Long, key As String, v As Variant, arr() As Variant
    Dim cMod As Long, cSD As Long, cSh As Long, cCell As Long, cX As Long, cY As Long, cDef As Long, cQty As Long
    Dim sIx(1 To 9) As Long, sHdr As Variant, sd As Long, qty As Double
    Dim nOld As Long, n As Long, newRows As Long, prevEv As Boolean

    prevEv = Application.EnableEvents
    On Error GoTo EH
    Application.EnableEvents = False
    Set loD = GetTable(SH_DB, TBL_DEFECTS)
    Set loS = GetTable(SH_SUM, TBL_SUMMARY)
    Set ws = GetSheet(SH_SUM)
    Set dict = CreateObject("Scripting.Dictionary")

    If TableHasData(loD) Then
        cMod = ColIdx(loD, "Model"): cSD = ColIdx(loD, "ShiftDate"): cSh = ColIdx(loD, "Shift")
        cCell = ColIdx(loD, "CellID"): cX = ColIdx(loD, "X"): cY = ColIdx(loD, "Y")
        cDef = ColIdx(loD, "DefectType"): cQty = ColIdx(loD, "Qty")
        dat = loD.DataBodyRange.Value2
        For i = 1 To UBound(dat, 1)
            If Len(CStr(dat(i, cMod))) > 0 Then
                If IsNumeric(dat(i, cSD)) Then
                    sd = CLng(dat(i, cSD))
                Else
                    sd = CLng(CDate(dat(i, cSD)))
                End If
                qty = 0
                If IsNumeric(dat(i, cQty)) Then qty = CDbl(dat(i, cQty))
                key = CStr(dat(i, cMod)) & "|" & sd & "|" & CStr(dat(i, cSh)) & "|" & _
                      CStr(dat(i, cCell)) & "|" & CStr(dat(i, cDef))
                If dict.Exists(key) Then
                    v = dict(key)
                    v(8) = v(8) + qty
                    v(9) = v(9) + 1
                    dict(key) = v
                Else
                    ReDim v(1 To 9)
                    v(1) = CStr(dat(i, cMod))
                    v(2) = sd
                    v(3) = CStr(dat(i, cSh))
                    v(4) = CStr(dat(i, cCell))
                    v(5) = dat(i, cX)
                    v(6) = dat(i, cY)
                    v(7) = CStr(dat(i, cDef))
                    v(8) = qty
                    v(9) = 1
                    dict.Add key, v
                End If
            End If
        Next i
    End If

    sHdr = Array("Model", "ShiftDate", "Shift", "CellID", "X", "Y", "DefectType", "TotalQty", "RecordCount")
    For m = 0 To 8
        sIx(m + 1) = ColIdx(loS, CStr(sHdr(m)))
    Next m

    nOld = 0
    If Not loS.DataBodyRange Is Nothing Then
        nOld = loS.DataBodyRange.Rows.Count
        loS.DataBodyRange.ClearContents
    End If
    n = dict.Count
    newRows = n
    If newRows < 1 Then newRows = 1
    loS.Resize loS.Range.Resize(newRows + 1)
    If nOld > newRows Then
        ws.Rows((loS.Range.Row + newRows + 1) & ":" & (loS.Range.Row + nOld)).Delete
    End If

    If n > 0 Then
        ReDim arr(1 To n, 1 To loS.ListColumns.Count)
        j = 0
        Dim k As Variant
        For Each k In dict.Keys
            j = j + 1
            v = dict(k)
            For m = 1 To 9
                arr(j, sIx(m)) = v(m)
            Next m
        Next k
        loS.ListColumns(sIx(2)).DataBodyRange.NumberFormat = "dd/mm/yyyy"
        loS.DataBodyRange.Value = arr
    End If
CLEANUP:
    Application.EnableEvents = prevEv
    Exit Sub
EH:
    Dim eN As Long, eD As String
    eN = Err.Number: eD = Err.Description
    Application.EnableEvents = prevEv
    Err.Raise eN, "RefreshSummary", eD
End Sub

'=====================================================================
' 9. KHOI TAO / DROPDOWN / BAO VE
'=====================================================================
Public Sub EnsureNamedLists()
    'Ten dong tro vao cot tro giup (chi Active = Yes) tren MASTER
    Dim lastRow As Long
    lastRow = MAX_LIST_ROWS + 1
    ThisWorkbook.Names.Add Name:="ActiveModels", RefersTo:="=OFFSET(" & SH_MASTER & "!$V$2,0,0,MAX(1,COUNTIF(" & _
        SH_MASTER & "!$V$2:$V$" & lastRow & ",""?*"")),1)"
    ThisWorkbook.Names.Add Name:="ActiveDefects", RefersTo:="=OFFSET(" & SH_MASTER & "!$W$2,0,0,MAX(1,COUNTIF(" & _
        SH_MASTER & "!$W$2:$W$" & lastRow & ",""?*"")),1)"
End Sub

Private Sub ApplyValidation(ByVal ws As Worksheet)
    With ws.Range("B2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ActiveModels"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Model"
        .ErrorMessage = "Hay chon Model trong danh sach."
    End With
    With ws.Range("D5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ActiveDefects"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Defect Type"
        .ErrorMessage = "Hay chon Defect Type trong danh sach."
    End With
    With ws.Range("D6").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="1"
        .ErrorTitle = "Qty"
        .ErrorMessage = "Qty per Cell phai la so nguyen >= 1."
    End With
    With ws.Range("B24").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=INDIRECT(""tblShifts[Shift]"")"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With ws.Range("B25").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ActiveDefects"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    With ws.Range("B26").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="Qty,Records"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With
End Sub

Private Sub ApplyProtection(ByVal ws As Worksheet)
    ws.Cells.Locked = True
    ws.Range("B2:B3").Locked = False
    ws.Range("D4:D7").Locked = False
    ws.Range("B21:B26").Locked = False
    ProtectMap ws
End Sub

Private Sub ProtectMap(ByVal ws As Worksheet)
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False, UserInterfaceOnly:=True
End Sub

Public Sub InitWorkbook()
    'Goi tu Workbook_Open
    Dim ws As Worksheet, model As String
    On Error GoTo EH
    Application.ScreenUpdating = False
    Set ws = GetSheet(SH_MAP)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo EH

    EnsureNamedLists
    ForceCloseCalendar
    ApplyValidation ws
    ws.Range("B3").NumberFormat = "@"
    UpdateShiftDisplay

    model = Trim$(CStr(ws.Range("B2").Value))
    If Len(model) > 0 Then
        LoadVehicleModel model
    Else
        Set gSelectedCells = CreateObject("Scripting.Dictionary")
        ws.Range("B4").Value = 0
    End If

    ApplyProtection ws

    On Error Resume Next
    ws.Activate
    With ActiveWindow
        .DisplayGridlines = False
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    ws.Range("F2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("B3").Select
    On Error GoTo EH
CLEANUP:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    MsgBox "InitWorkbook loi: " & Err.Description, vbCritical
    Resume CLEANUP
End Sub

'=====================================================================
' 10. SETUP (chay mot lan) - tao sheet, bang, giao dien
'=====================================================================
Public Sub SetupWorkbook()
    Dim prevEv As Boolean
    On Error GoTo EH
    If MsgBox("Setup se tao 4 sheet, cac bang, giao dien, nut bam va hinh placeholder." & vbCrLf & _
              "Du lieu trong cac bang da ton tai se KHONG bi xoa. Tiep tuc?", vbYesNo + vbQuestion, "Setup") <> vbYes Then Exit Sub
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    EnsureSheets
    BuildMasterSheet
    BuildDatabaseSheet
    BuildSummarySheet
    BuildMapSheet
    EnsureNamedLists
    CreatePlaceholderImages

    Application.EnableEvents = prevEv
    InitWorkbook
    Application.ScreenUpdating = True
    MsgBox "Setup xong." & vbCrLf & _
           "Buoc tiep theo: dan code vao Sheet MAP_VEHICLE va ThisWorkbook, thay hinh placeholder IMG_xxx bang hinh Master that, roi luu file .xlsm.", _
           vbInformation, "Setup"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    Application.EnableEvents = prevEv
    MsgBox "SetupWorkbook loi: " & Err.Description, vbCritical
End Sub

Private Function FindBlankSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, SH_MAP, vbTextCompare) <> 0 And StrComp(ws.Name, SH_MASTER, vbTextCompare) <> 0 And _
           StrComp(ws.Name, SH_DB, vbTextCompare) <> 0 And StrComp(ws.Name, SH_SUM, vbTextCompare) <> 0 And _
           StrComp(ws.Name, "Dashboard", vbTextCompare) <> 0 And StrComp(ws.Name, "Config", vbTextCompare) <> 0 And _
           StrComp(ws.Name, "DASH_CALC", vbTextCompare) <> 0 Then
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then
                Set FindBlankSheet = ws: Exit Function
            End If
        End If
    Next ws
End Function

Private Sub EnsureSheets()
    Dim names As Variant, i As Long, ws As Worksheet
    names = Array(SH_MAP, SH_MASTER, SH_DB, SH_SUM)
    For i = 0 To 3
        If Not SheetExists(CStr(names(i))) Then
            Set ws = FindBlankSheet()
            If ws Is Nothing Then
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            End If
            ws.Name = CStr(names(i))
        End If
    Next i
    For i = 3 To 0 Step -1
        ThisWorkbook.Worksheets(CStr(names(i))).Move Before:=ThisWorkbook.Worksheets(1)
    Next i
    'Xoa sheet thua con trong (chi giu dung 4 sheet)
    Application.DisplayAlerts = False
    Do
        Set ws = FindBlankSheet()
        If ws Is Nothing Then Exit Do
        ws.Delete
    Loop
    Application.DisplayAlerts = True
End Sub

Private Function EnsureTable(ByVal ws As Worksheet, ByVal tblName As String, ByVal topLeft As String, _
                             ByVal headers As Variant, Optional ByVal rowsData As Variant) As ListObject
    Dim lo As ListObject, i As Long, j As Long, nCols As Long, nRows As Long, rng As Range
    On Error Resume Next
    Set lo = ws.ListObjects(tblName)
    On Error GoTo 0
    If Not lo Is Nothing Then Set EnsureTable = lo: Exit Function

    nCols = UBound(headers) - LBound(headers) + 1
    nRows = 1
    If Not IsMissing(rowsData) Then nRows = UBound(rowsData) - LBound(rowsData) + 1
    Set rng = ws.Range(topLeft).Resize(nRows + 1, nCols)
    For j = 0 To nCols - 1
        rng.Cells(1, j + 1).Value = headers(LBound(headers) + j)
    Next j
    If Not IsMissing(rowsData) Then
        For i = 0 To nRows - 1
            For j = 0 To nCols - 1
                rng.Cells(i + 2, j + 1).Value = rowsData(LBound(rowsData) + i)(j)
            Next j
        Next i
    End If
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = tblName
    lo.TableStyle = "TableStyleMedium2"
    Set EnsureTable = lo
End Function

Private Sub BuildMasterSheet()
    Dim ws As Worksheet, lo As ListObject, lastRow As Long
    Set ws = GetSheet(SH_MASTER)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    EnsureTable ws, TBL_MODELS, "A1", Array("Model", "Rows", "Cols", "ImageShapeName", "Active"), _
        Array(Array("VF3", 20, 30, "IMG_VF3", "Yes"), Array("VF5", 20, 30, "IMG_VF5", "Yes"), Array("MNV", 24, 36, "IMG_MNV", "Yes"))

    EnsureTable ws, TBL_DEFECTTYPES, "G1", Array("DefectType", "Active"), _
        Array(Array("Dust", "Yes"), Array("Scratch", "Yes"), Array("Paint sagging", "Yes"), Array("Orange peel", "Yes"), Array("Dirt", "Yes"))

    Set lo = EnsureTable(ws, TBL_SHIFTS, "J1", Array("Shift", "StartTime", "EndTime", "Active"), _
        Array(Array("S1", TimeSerial(6, 0, 0), TimeSerial(14, 0, 0), "Yes"), _
              Array("S2", TimeSerial(14, 0, 0), TimeSerial(22, 0, 0), "Yes"), _
              Array("S3", TimeSerial(22, 0, 0), TimeSerial(6, 0, 0), "Yes")))
    lo.ListColumns("StartTime").DataBodyRange.NumberFormat = "hh:mm"
    lo.ListColumns("EndTime").DataBodyRange.NumberFormat = "hh:mm"

    'Khong tu tao toa do vung - nguoi dung tu khai bao theo bo cuc hinh that
    EnsureTable ws, TBL_ZONES, "O1", Array("Model", "View", "XFrom", "XTo", "YFrom", "YTo")

    'Cot tro giup cho dropdown (chi lay Active = Yes). KHONG xoa/sua.
    lastRow = MAX_LIST_ROWS + 1
    ws.Range("V1").Value = "ActiveModels (auto)"
    ws.Range("W1").Value = "ActiveDefects (auto)"
    ws.Range("V2:V" & lastRow).Formula = "=IFERROR(INDEX(tblModels[Model],AGGREGATE(15,6," & _
        "(ROW(tblModels[Model])-ROW(tblModels[[#Headers],[Model]]))/((tblModels[Active]=""Yes"")*(tblModels[Model]<>""""))," & _
        "ROWS(V$2:V2))),"""")"
    ws.Range("W2:W" & lastRow).Formula = "=IFERROR(INDEX(tblDefectTypes[DefectType],AGGREGATE(15,6," & _
        "(ROW(tblDefectTypes[DefectType])-ROW(tblDefectTypes[[#Headers],[DefectType]]))/((tblDefectTypes[Active]=""Yes"")*(tblDefectTypes[DefectType]<>""""))," & _
        "ROWS(W$2:W2))),"""")"
    With ws.Range("V1:W" & lastRow)
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
        .Interior.Color = RGB(242, 242, 242)
    End With
    ws.Range("V1:W1").Font.Bold = True
    ws.Columns("A:W").AutoFit
    ws.Columns("F").ColumnWidth = 3: ws.Columns("I").ColumnWidth = 3
    ws.Columns("N").ColumnWidth = 3: ws.Columns("U").ColumnWidth = 3
    ws.Columns("V:W").ColumnWidth = 22

    'Kich thuoc hinh Master CHUNG cho moi Model (point). Sua o day de doi kich thuoc/Grid cua tat ca Model.
    EnsureTable ws, TBL_SETTINGS, "Y1", Array("Setting", "Value", "Note"), _
        Array(Array("ImageWidth", 1000, "Chieu rong hinh Master (point) - ap dung cho TAT CA Model"), _
              Array("ImageHeight", 500, "Chieu cao hinh Master (point) - nen cung ti le voi hinh goc de hinh khong bi meo"))
    ws.Columns("X").ColumnWidth = 3
    ws.Columns("Y:Z").ColumnWidth = 14
    ws.Columns("AA").ColumnWidth = 80
End Sub

Private Sub BuildDatabaseSheet()
    Dim ws As Worksheet, lo As ListObject
    Set ws = GetSheet(SH_DB)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    Set lo = EnsureTable(ws, TBL_DEFECTS, "A1", Array("DefectID", "BatchID", "DateTime", "ShiftDate", "Model", "VIN", "View", _
        "X", "Y", "CellID", "GridID", "DefectType", "Qty", "Shift", "Inspector", "Description", "Status"))
    lo.ListColumns("VIN").DataBodyRange.NumberFormat = "@"
    lo.ListColumns("CellID").DataBodyRange.NumberFormat = "@"
    lo.ListColumns("GridID").DataBodyRange.NumberFormat = "@"
    lo.ListColumns("DateTime").DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    lo.ListColumns("ShiftDate").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    ws.Columns("A:Q").ColumnWidth = 14
    ws.Columns("A:B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("P").ColumnWidth = 30
End Sub

Private Sub BuildSummarySheet()
    Dim ws As Worksheet, lo As ListObject
    Set ws = GetSheet(SH_SUM)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "SUMMARY - Tong hop loi theo Model / ShiftDate / Shift / CellID / DefectType"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    Set lo = EnsureTable(ws, TBL_SUMMARY, "A3", Array("Model", "ShiftDate", "Shift", "CellID", "X", "Y", "DefectType", "TotalQty", "RecordCount"))
    lo.ListColumns("ShiftDate").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    ws.Columns("A:I").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 20
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal shpName As String, ByVal caption As String, _
                      ByVal rng As Range, ByVal clr As Long, ByVal macroName As String)
    Dim s As Shape
    Set s = FindShape(ws, shpName)
    If Not s Is Nothing Then s.Delete
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    With s
        .Name = shpName
        .Fill.ForeColor.RGB = clr
        .Line.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlFreeFloating
        With .TextFrame2.TextRange
            .Text = caption
            .Font.Size = 12
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
End Sub

Private Sub BuildMapSheet()
    'Bo cuc 2 phan: TRAI = panel Model/nhap lieu/nut (cot A:D), PHAI = hinh Master + Grid (tu F2)
    Dim ws As Worksheet
    Set ws = GetSheet(SH_MAP)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ws.Range("A8:AD30").UnMerge
    ws.Range("A8:AD30").Clear                 'don bo cuc cu (neu chay lai Setup)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F:AD").ColumnWidth = 8.43
    ws.Rows(1).RowHeight = 34
    ws.Rows("2:7").RowHeight = 22
    ws.Rows(8).RowHeight = 8
    ws.Rows("9:13").RowHeight = 24

    'Tieu de
    With ws.Range("A1:AD1")
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A1").Value = "VEHICLE DEFECT MAPPING"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True
    ws.Range("F1").Value = "VEHICLE MASTER  -  click vao o luoi de chon / bo chon vi tri loi"
    ws.Range("F1").Font.Size = 11
    ws.Range("F1").Font.Bold = True

    'Nhan
    ws.Range("A2").Value = "Model"
    ws.Range("A3").Value = "VIN (scan)"
    ws.Range("A4").Value = "Selected Cells"
    ws.Range("A5").Value = "Grid Rows"
    ws.Range("A6").Value = "Grid Columns"
    ws.Range("A7").Value = "Shift"
    ws.Range("C2").Value = "Shift Date"
    ws.Range("C3").Value = "Date Time"
    ws.Range("C4").Value = "Inspector"
    ws.Range("C5").Value = "Defect Type"
    ws.Range("C6").Value = "Qty / Cell"
    ws.Range("C7").Value = "Description"

    'Mau theo khu vuc: Vehicle (xanh duong) / Production (cam) / Defect (do)
    ws.Range("A2:A6").Interior.Color = RGB(31, 78, 121)
    ws.Range("A7").Interior.Color = RGB(191, 113, 0)
    ws.Range("C2:C4").Interior.Color = RGB(191, 113, 0)
    ws.Range("C5:C7").Interior.Color = RGB(160, 30, 30)
    With ws.Range("A2:A7,C2:C7")
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With

    'O nhap lieu (vang) / o tu dong (xam)
    ws.Range("B2:B3,D4:D7").Interior.Color = RGB(255, 242, 204)
    ws.Range("B2:B3,D4:D7").Font.Bold = True
    ws.Range("B4:B7,D2:D3").Interior.Color = RGB(217, 217, 217)
    ws.Range("B2:B7,D2:D7").HorizontalAlignment = xlLeft
    ws.Range("B2:B7,D2:D7").VerticalAlignment = xlCenter
    With ws.Range("A2:D7").Borders
        .LineStyle = xlContinuous
        .Color = RGB(150, 150, 150)
        .Weight = xlThin
    End With

    ws.Range("B3").NumberFormat = "@"
    ws.Range("B4:B6").NumberFormat = "0"
    ws.Range("D2").NumberFormat = "dd/mm/yyyy"
    ws.Range("D3").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("D6").NumberFormat = "0"
    If IsEmpty(ws.Range("B4").Value) Then ws.Range("B4").Value = 0
    If IsEmpty(ws.Range("D6").Value) Then ws.Range("D6").Value = 1

    'Huong dan ngan duoi cac nut
    With ws.Range("A15:D18")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(90, 90, 90)
    End With
    ws.Range("A15").Value = "Vang = nhap tay / chon  |  Xam = tu dong" & vbLf & _
        "Quy trinh: Model > VIN > Defect Type > Qty > Inspector > Description > click o loi tren hinh ben phai > SAVE"

    'Nut bam (panel ben trai)
    AddButton ws, "BTN_SAVE", "SAVE", ws.Range("A9:D11"), RGB(0, 140, 70), "modVehicleMapping.BtnSave"
    AddButton ws, "BTN_CLEAR", "CLEAR SELECTION", ws.Range("A12:B13"), RGB(200, 120, 0), "modVehicleMapping.BtnClear"
    AddButton ws, "BTN_RESET", "RESET INPUT", ws.Range("C12:D13"), RGB(120, 120, 120), "modVehicleMapping.BtnReset"

    'Panel XEM DU LIEU DATABASE tren Grid (heatmap)
    ws.Rows("20:29").RowHeight = 22
    ws.Range("A20").Value = "XEM DU LIEU DATABASE TREN GRID"
    With ws.Range("A20:D20")
        .Interior.Color = RGB(64, 64, 64)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A21").Value = "From date"
    ws.Range("A22").Value = "To date"
    ws.Range("A23").Value = "VIN"
    ws.Range("A24").Value = "Shift"
    ws.Range("A25").Value = "Defect Type"
    ws.Range("A26").Value = "Show value"
    With ws.Range("A21:A26")
        .Interior.Color = RGB(89, 89, 89)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .VerticalAlignment = xlCenter
    End With
    With ws.Range("B21:B26")
        .Interior.Color = RGB(255, 242, 204)
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B21:B22").NumberFormat = "dd/mm/yyyy"
    ws.Range("B23").NumberFormat = "@"
    ws.Range("B26").Value = "Qty"
    With ws.Range("C21:D26")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Italic = True
        .Font.Size = 9
        .Font.Color = RGB(90, 90, 90)
    End With
    ws.Range("C21").Value = "De trong = khong loc. Luon loc theo Model dang chon (B2)." & vbLf & _
        "Ngay loc theo ShiftDate. Mau: vang = it, do = nhieu (so voi o cao nhat). So tren o = Qty hoac so ban ghi."
    With ws.Range("A21:D26").Borders
        .LineStyle = xlContinuous
        .Color = RGB(150, 150, 150)
        .Weight = xlThin
    End With
    AddButton ws, "BTN_SHOW", "SHOW DATA", ws.Range("A27:B28"), RGB(31, 78, 121), "modVehicleMapping.BtnShowData"
    AddButton ws, "BTN_HIDE", "HIDE DATA", ws.Range("C27:D28"), RGB(120, 120, 120), "modVehicleMapping.BtnHideData"
    With ws.Range("A29:D29")
        .Merge
        .Font.Bold = True
        .Font.Size = 9
        .VerticalAlignment = xlCenter
        .Font.Color = RGB(31, 78, 121)
    End With
End Sub

Private Sub CreatePlaceholderImages()
    'Tao hinh tam (hinh chu nhat xam) cho moi ImageShapeName chua co. Thay bang hinh Master that.
    Dim ws As Worksheet, lo As ListObject, dat As Variant, i As Long, cI As Long, nm As String, s As Shape
    Dim pw As Double, ph As Double
    Set ws = GetSheet(SH_MAP)
    Set lo = GetTable(SH_MASTER, TBL_MODELS)
    If Not TableHasData(lo) Then Exit Sub
    If Not GetImageSize(pw, ph) Then pw = IMG_PLACEHOLDER_WIDTH: ph = IMG_PLACEHOLDER_WIDTH / 2
    cI = ColIdx(lo, "ImageShapeName")
    dat = lo.DataBodyRange.Value2
    For i = 1 To UBound(dat, 1)
        nm = Trim$(CStr(dat(i, cI)))
        If Len(nm) > 0 Then
            If FindShape(ws, nm) Is Nothing Then
                Set s = ws.Shapes.AddShape(msoShapeRectangle, ws.Range(ANCHOR_IMG).Left, ws.Range(ANCHOR_IMG).Top, _
                                           pw, ph)
                s.Name = nm
                s.Fill.ForeColor.RGB = RGB(225, 225, 225)
                s.Line.ForeColor.RGB = RGB(150, 150, 150)
                s.TextFrame2.TextRange.Text = nm & " - PLACEHOLDER: thay bang hinh Master that (LH / RH / FRONT / REAR / TOP) va dat Name = " & nm
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
                s.TextFrame2.TextRange.Font.Size = 14
                s.Visible = msoFalse
            End If
        End If
    Next i
End Sub
