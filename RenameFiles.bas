Attribute VB_Name = "RenameFiles"
Option Explicit

'==============================================================================
' MODULE : RenameFiles
' PURPOSE: Standardize file names in a folder AND all of its sub-folders
'          Target format : FVN_PSHT_[FileName]_[DDMMYYYY].[Ext]
'
' FLOW   : PreviewRename()  -> scans folder, fills sheet "Preview", NO rename
'          ConfirmRename()  -> renames rows with Status = "Ready", writes log
'
' SHEETS : "Preview"    A:STT  B:File Name Current  C:File Name New
'                       D:Extension  E:Status  F:Reason
'                       G:Folder Path (full folder of each file, used by ConfirmRename)
'          "Rename_Log" A:Time  B:Old Name  C:New Name  D:Result
'
' SPEED  : all data is read/written with arrays (no cell-by-cell loops),
'          so 10,000+ files are fine.
'==============================================================================

Private Const SHEET_PREVIEW As String = "Preview"
Private Const SHEET_LOG As String = "Rename_Log"
Private Const NAME_PREFIX As String = "FVN_PSHT_"
Private Const MAX_PATH_LEN As Long = 259

'==============================================================================
' PHASE 1 - PREVIEW (does NOT rename anything)
'==============================================================================
Public Sub PreviewRename()

    Dim folderPath As String
    Dim fso As Object, fol As Object, f As Object
    Dim names() As String
    Dim dirs() As String        ' folder (with trailing "\") of each file
    Dim stack As Collection     ' folders waiting to be scanned
    Dim sf As Object            ' sub-folder
    Dim curDir As String
    Dim cap As Long
    Dim n As Long, i As Long
    Dim existing As Object      ' all current full paths (case-insensitive)
    Dim newNames As Object      ' new full paths already assigned in this preview
    Dim result() As Variant
    Dim ws As Worksheet

    Dim curName As String, lName As String
    Dim baseName As String, ext As String, pos As Long
    Dim newBase As String, newFull As String
    Dim reason As String
    Dim cntReady As Long, cntSkip As Long

    ' ---- 1. Pick folder ------------------------------------------------------
    folderPath = PickFolder()
    If Len(folderPath) = 0 Then Exit Sub            ' user cancelled
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    On Error GoTo ErrHandler
    SpeedOn

    ' ---- 2. Scan files of the selected folder only (no sub-folders) ----------
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fol = fso.GetFolder(folderPath)
    Set existing = CreateObject("Scripting.Dictionary")
    existing.CompareMode = 1                        ' 1 = vbTextCompare
    Set newNames = CreateObject("Scripting.Dictionary")
    newNames.CompareMode = 1

    ' Iterative scan (stack instead of recursion): root folder + all sub-folders
    cap = 1024
    ReDim names(1 To cap)
    ReDim dirs(1 To cap)
    n = 0
    Set stack = New Collection
    stack.Add fol.Path

    Do While stack.Count > 0
        curDir = stack(stack.Count)                 ' pop last folder
        stack.Remove stack.Count
        If Right$(curDir, 1) <> "\" Then curDir = curDir & "\"

        ' folders we cannot access (permission...) are silently skipped
        On Error Resume Next
        Set fol = Nothing
        Set fol = fso.GetFolder(curDir)
        If Not fol Is Nothing Then
            For Each sf In fol.SubFolders
                stack.Add sf.Path
            Next sf
        End If
        On Error GoTo ErrHandler

        If Not fol Is Nothing Then
            For Each f In fol.Files
                n = n + 1
                If n > cap Then                     ' grow arrays (doubling)
                    cap = cap * 2
                    ReDim Preserve names(1 To cap)
                    ReDim Preserve dirs(1 To cap)
                End If
                names(n) = f.Name
                dirs(n) = curDir
                existing(curDir & f.Name) = 1       ' remember every existing path
            Next f
        End If

        Application.StatusBar = "Scanning folders... " & n & " files found"
    Loop

    If n = 0 Then
        SpeedOff
        MsgBox "No file found in the selected folder.", vbInformation
        Exit Sub
    End If

    ' ---- 3. Build preview data in memory ------------------------------------
    ReDim result(1 To n, 1 To 7)

    For i = 1 To n
        curName = names(i)
        lName = LCase$(curName)
        reason = ""

        ' split base name / extension (keep extension unchanged)
        pos = InStrRev(curName, ".")
        If pos > 1 Then
            baseName = Left$(curName, pos - 1)
            ext = Mid$(curName, pos + 1)
        Else
            baseName = curName
            ext = ""
        End If

        result(i, 1) = i
        result(i, 2) = curName
        result(i, 3) = ""
        result(i, 4) = ext
        result(i, 7) = dirs(i)                      ' folder of this file

        If IsSystemFile(lName) Then
            ' ---- skip system / temp files
            result(i, 5) = "Skip"
            result(i, 6) = "System File"

        ElseIf IsStandardized(lName) Then
            ' ---- skip files already starting with FVN_ / FVN- / FVN.
            result(i, 5) = "Skip"
            result(i, 6) = "Already Standardized"

        Else
            ' ---- build the new name (returns "" + reason when not possible)
            newBase = BuildNewBase(baseName, reason)

            If Len(newBase) = 0 Then
                result(i, 5) = "Skip"
                result(i, 6) = reason               ' No Date Found / Empty Name

            Else
                newFull = newBase
                If Len(ext) > 0 Then newFull = newFull & "." & ext

                If Len(dirs(i)) + Len(newFull) > MAX_PATH_LEN Then
                    result(i, 3) = newFull
                    result(i, 5) = "Skip"
                    result(i, 6) = "Path Too Long"

                ElseIf existing.Exists(dirs(i) & newFull) Or newNames.Exists(dirs(i) & newFull) Then
                    ' target already exists in the SAME folder, or another file got it first
                    result(i, 3) = newFull
                    result(i, 5) = "Skip"
                    result(i, 6) = "Duplicate Name"

                Else
                    newNames(dirs(i) & newFull) = 1
                    result(i, 3) = newFull
                    result(i, 5) = "Ready"
                    result(i, 6) = "Will Rename"
                    cntReady = cntReady + 1
                End If
            End If
        End If

        If result(i, 5) = "Skip" Then cntSkip = cntSkip + 1

        If i Mod 500 = 0 Then Application.StatusBar = "Scanning... " & i & " / " & n
    Next i

    ' ---- 4. Write everything to sheet "Preview" in ONE shot ------------------
    Set ws = GetOrCreateSheet(SHEET_PREVIEW)
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete

    ws.Range("A1:G1").Value = Array("STT", "File Name Current", "File Name New", _
                                    "Extension", "Status", "Reason", "Folder Path")

    ws.Range("B2:D" & n + 1).NumberFormat = "@"     ' text, avoid auto-conversion
    ws.Range("G2:G" & n + 1).NumberFormat = "@"
    ws.Range("A2").Resize(n, 7).Value = result

    ' ---- 5. Formatting -------------------------------------------------------
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
    ws.Columns("A:G").AutoFit
    If ws.Columns("B").ColumnWidth > 70 Then ws.Columns("B").ColumnWidth = 70
    If ws.Columns("C").ColumnWidth > 70 Then ws.Columns("C").ColumnWidth = 70
    If ws.Columns("G").ColumnWidth > 80 Then ws.Columns("G").ColumnWidth = 80
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:G" & n + 1).AutoFilter

    ' conditional formatting (cheap, no per-row loop)
    With ws.Range("A2:G" & n + 1)
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$E2=""Ready""")
            .Interior.Color = RGB(226, 239, 218)
        End With
        With .FormatConditions.Add(Type:=xlExpression, Formula1:="=$E2=""Skip""")
            .Interior.Color = RGB(237, 237, 237)
            .Font.Color = RGB(128, 128, 128)
        End With
    End With

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    SpeedOff
    MsgBox "Preview completed." & vbCrLf & _
           "Please review sheet Preview." & vbCrLf & _
           "Run ConfirmRename() to execute renaming." & vbCrLf & vbCrLf & _
           "Ready: " & cntReady & "  |  Skip: " & cntSkip, vbInformation
    Exit Sub

ErrHandler:
    SpeedOff
    MsgBox "Preview error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'==============================================================================
' PHASE 2 - CONFIRM RENAME (only rows with Status = "Ready")
'==============================================================================
Public Sub ConfirmRename()

    Dim ws As Worksheet, wsLog As Worksheet
    Dim fso As Object
    Dim folderPath As String
    Dim lastRow As Long, i As Long, k As Long
    Dim data As Variant
    Dim stat() As Variant           ' new Status / Reason (cols E:F)
    Dim logArr() As Variant         ' log rows
    Dim readyCount As Long
    Dim oldPath As String, newPath As String
    Dim cntOK As Long, cntDup As Long, cntErr As Long
    Dim errTxt As String
    Dim logStart As Long

    ' ---- 1. Read sheet Preview ----------------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PREVIEW)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet Preview not found. Run PreviewRename() first.", vbExclamation
        Exit Sub
    End If

    ' each row carries its own folder in column G (sub-folders supported)
    Set fso = CreateObject("Scripting.FileSystemObject")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Sheet Preview is empty.", vbExclamation
        Exit Sub
    End If

    data = ws.Range("A2:G" & lastRow).Value

    ' ---- 2. Count Ready rows -------------------------------------------------
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 5)) = "Ready" Then readyCount = readyCount + 1
    Next i

    If readyCount = 0 Then
        MsgBox "No file with Status = Ready.", vbInformation
        Exit Sub
    End If

    ' ---- 3. Ask for confirmation --------------------------------------------
    If MsgBox("Are you sure to rename " & readyCount & " files ?", _
              vbYesNo + vbQuestion, "Confirm Rename") <> vbYes Then Exit Sub

    ' ---- 4. Rename -----------------------------------------------------------
    SpeedOn
    ReDim stat(1 To UBound(data, 1), 1 To 2)
    ReDim logArr(1 To readyCount, 1 To 4)
    k = 0

    For i = 1 To UBound(data, 1)
        stat(i, 1) = data(i, 5)
        stat(i, 2) = data(i, 6)

        If CStr(data(i, 5)) = "Ready" Then
            k = k + 1
            oldPath = CStr(data(i, 7)) & CStr(data(i, 2))
            newPath = CStr(data(i, 7)) & CStr(data(i, 3))

            ' log full paths (same file name may exist in several sub-folders)
            logArr(k, 1) = Now
            logArr(k, 2) = oldPath
            logArr(k, 3) = newPath

            If Not fso.FileExists(oldPath) Then
                ' source disappeared since preview
                stat(i, 1) = "Error": stat(i, 2) = "Source file not found"
                logArr(k, 4) = "Error: Source file not found"
                cntErr = cntErr + 1

            ElseIf fso.FileExists(newPath) Then
                ' target already exists -> never overwrite
                stat(i, 1) = "Duplicate": stat(i, 2) = "New name already exists"
                logArr(k, 4) = "Duplicate: New name already exists"
                cntDup = cntDup + 1

            Else
                On Error Resume Next
                Name oldPath As newPath
                If Err.Number <> 0 Then
                    errTxt = Err.Number & " - " & Err.Description
                    Err.Clear
                    On Error GoTo 0
                    stat(i, 1) = "Error": stat(i, 2) = errTxt
                    logArr(k, 4) = "Error: " & errTxt
                    cntErr = cntErr + 1
                Else
                    On Error GoTo 0
                    stat(i, 1) = "Success": stat(i, 2) = "Renamed"
                    logArr(k, 4) = "Success"
                    cntOK = cntOK + 1
                End If
            End If

            If k Mod 200 = 0 Then Application.StatusBar = "Renaming... " & k & " / " & readyCount
        End If
    Next i

    ' ---- 5. Write Status/Reason back to Preview (one shot) -------------------
    ws.Range("E2").Resize(UBound(stat, 1), 2).Value = stat
    ws.Columns("E:F").AutoFit

    ' ---- 6. Append log to sheet Rename_Log -----------------------------------
    Set wsLog = GetOrCreateSheet(SHEET_LOG)
    If Len(CStr(wsLog.Range("A1").Value)) = 0 Then
        wsLog.Range("A1:D1").Value = Array("Time", "Old Name", "New Name", "Result")
        With wsLog.Range("A1:D1")
            .Font.Bold = True
            .Interior.Color = RGB(31, 78, 121)
            .Font.Color = vbWhite
        End With
    End If
    logStart = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    wsLog.Range("B" & logStart).Resize(readyCount, 2).NumberFormat = "@"
    wsLog.Range("A" & logStart).Resize(readyCount, 1).NumberFormat = "dd/mm/yyyy hh:nn:ss"
    wsLog.Range("A" & logStart).Resize(readyCount, 4).Value = logArr
    wsLog.Columns("A:D").AutoFit
    If wsLog.Columns("B").ColumnWidth > 70 Then wsLog.Columns("B").ColumnWidth = 70
    If wsLog.Columns("C").ColumnWidth > 70 Then wsLog.Columns("C").ColumnWidth = 70

    SpeedOff
    MsgBox "Rename completed." & vbCrLf & _
           "Success  : " & cntOK & vbCrLf & _
           "Duplicate: " & cntDup & vbCrLf & _
           "Error    : " & cntErr & vbCrLf & _
           "See sheet Rename_Log for details.", vbInformation
End Sub

'==============================================================================
' NAME BUILDING
'==============================================================================

' Returns the new base name (without extension), or "" with reason when the
' file must be skipped ("No Date Found" / "Empty Name").
Private Function BuildNewBase(ByVal baseName As String, ByRef reason As String) As String

    Dim s As String
    Dim leadDate As String, trailDate As String, d As String
    Dim finalDate As String
    Dim pos As Long, token As String
    Dim changed As Boolean
    Dim prefixes As Variant, p As Variant

    prefixes = Array("VF_PSHT_", "FVN_PSHT_", "VF_", "FVN_")   ' longest first

    ' 1) remove Vietnamese accents, replace special chars with "_", clean "_"
    s = SanitizeName(baseName)

    ' 2) strip leading dates and known prefixes (repeat until nothing changes)
    Do
        changed = False

        ' leading date: YYMMDD_ or DDMMYYYY_
        pos = InStr(s, "_")
        If pos > 0 Then
            d = ParseDateToken(Left$(s, pos - 1))
            If Len(d) > 0 Then
                If Len(leadDate) = 0 Then leadDate = d      ' keep the first one
                s = Mid$(s, pos + 1)
                changed = True
            End If
        End If

        ' prefixes VF_PSHT_, FVN_PSHT_, VF_, FVN_
        For Each p In prefixes
            If UCase$(Left$(s, Len(p))) = p Then
                s = Mid$(s, Len(p) + 1)
                changed = True
                Exit For
            End If
        Next p
    Loop While changed

    ' 3) strip trailing dates (avoids 30072026_..._30072026)
    Do
        pos = InStrRev(s, "_")
        If pos = 0 Then Exit Do
        d = ParseDateToken(Mid$(s, pos + 1))
        If Len(d) = 0 Then Exit Do
        If Len(trailDate) = 0 Then trailDate = d            ' nearest to the end
        s = Left$(s, pos - 1)
    Loop

    ' 4) choose date: leading date wins, otherwise trailing date
    If Len(leadDate) > 0 Then
        finalDate = leadDate
    Else
        finalDate = trailDate
    End If

    If Len(finalDate) = 0 Then
        reason = "No Date Found"
        Exit Function
    End If

    s = TrimUnderscore(s)
    If Len(s) = 0 Then
        reason = "Empty Name"
        Exit Function
    End If

    BuildNewBase = NAME_PREFIX & s & "_" & finalDate
End Function

' Accepts "YYMMDD" or "DDMMYYYY" (digits only). Returns "DDMMYYYY" or "".
Private Function ParseDateToken(ByVal t As String) As String
    Dim dd As Long, mm As Long, yy As Long

    If t Like "######" Then                       ' YYMMDD
        yy = 2000 + CLng(Left$(t, 2))
        mm = CLng(Mid$(t, 3, 2))
        dd = CLng(Right$(t, 2))
    ElseIf t Like "########" Then                 ' DDMMYYYY
        dd = CLng(Left$(t, 2))
        mm = CLng(Mid$(t, 3, 2))
        yy = CLng(Right$(t, 4))
        If yy < 1990 Or yy > 2099 Then Exit Function
    Else
        Exit Function
    End If

    If Not IsValidDate(yy, mm, dd) Then Exit Function
    ParseDateToken = Format$(dd, "00") & Format$(mm, "00") & Format$(yy, "0000")
End Function

Private Function IsValidDate(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Boolean
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    Dim dt As Date
    dt = DateSerial(y, m, d)
    IsValidDate = (Year(dt) = y And Month(dt) = m And Day(dt) = d)
End Function

' Remove accents, replace special chars with "_", collapse "__", trim "_"
Private Function SanitizeName(ByVal s As String) As String
    s = RemoveVietnamese(s)

    Dim chars As Variant, c As Variant
    chars = Array(" ", ",", ".", "-", "(", ")", "[", "]", "&", "/", "\", ":", ";", "'", vbTab)
    For Each c In chars
        s = Replace(s, c, "_")
    Next c

    Do While InStr(s, "__") > 0
        s = Replace(s, "__", "_")
    Loop

    SanitizeName = TrimUnderscore(s)
End Function

Private Function TrimUnderscore(ByVal s As String) As String
    Do While Left$(s, 1) = "_"
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = "_"
        s = Left$(s, Len(s) - 1)
    Loop
    TrimUnderscore = s
End Function

' Convert Vietnamese text to unsigned ASCII (uses char codes, so the code
' does not depend on the editor's code page).
Private Function RemoveVietnamese(ByVal s As String) As String
    Dim i As Long, n As Long, code As Long
    Dim buf As String, outPos As Long, r As String

    n = Len(s)
    buf = Space$(n)
    outPos = 0

    For i = 1 To n
        code = AscW(Mid$(s, i, 1))
        If code < 0 Then code = code + 65536
        r = ""

        If code < 192 Then
            r = ChrW$(code)                       ' plain ASCII etc. -> keep
        Else
            Select Case code
                Case 768 To 879: r = ""           ' combining accents -> drop
                Case 192 To 195, 258:   r = "A"
                Case 224 To 227, 259:   r = "a"
                Case 200 To 202:        r = "E"
                Case 232 To 234:        r = "e"
                Case 204, 205, 296:     r = "I"
                Case 236, 237, 297:     r = "i"
                Case 210 To 213, 416:   r = "O"
                Case 242 To 245, 417:   r = "o"
                Case 217, 218, 360, 431: r = "U"
                Case 249, 250, 361, 432: r = "u"
                Case 221:               r = "Y"
                Case 253:               r = "y"
                Case 272:               r = "D"
                Case 273:               r = "d"
                Case 7840 To 7863       ' A with dots/hooks/circumflex/breve
                    r = IIf(code Mod 2 = 0, "A", "a")
                Case 7864 To 7879       ' E
                    r = IIf(code Mod 2 = 0, "E", "e")
                Case 7880 To 7883       ' I
                    r = IIf(code Mod 2 = 0, "I", "i")
                Case 7884 To 7907       ' O
                    r = IIf(code Mod 2 = 0, "O", "o")
                Case 7908 To 7921       ' U
                    r = IIf(code Mod 2 = 0, "U", "u")
                Case 7922 To 7929       ' Y
                    r = IIf(code Mod 2 = 0, "Y", "y")
                Case Else
                    r = ChrW$(code)               ' other chars -> keep
            End Select
        End If

        If Len(r) > 0 Then
            outPos = outPos + 1
            Mid$(buf, outPos, 1) = r
        End If
    Next i

    RemoveVietnamese = Left$(buf, outPos)
End Function

'==============================================================================
' FILE FILTERS
'==============================================================================

' Thumbs.db, desktop.ini, ~$*, *.tmp, *.log, *.bak, *.dwl, *.dwl2
Private Function IsSystemFile(ByVal lName As String) As Boolean
    If lName = "thumbs.db" Or lName = "desktop.ini" Then
        IsSystemFile = True
    ElseIf Left$(lName, 2) = "~$" Then
        IsSystemFile = True
    ElseIf lName Like "*.tmp" Or lName Like "*.log" Or lName Like "*.bak" _
        Or lName Like "*.dwl" Or lName Like "*.dwl2" Then
        IsSystemFile = True
    End If
End Function

' Name starts with FVN_ / FVN- / FVN.
Private Function IsStandardized(ByVal lName As String) As Boolean
    Select Case Left$(lName, 4)
        Case "fvn_", "fvn-", "fvn."
            IsStandardized = True
    End Select
End Function

'==============================================================================
' HELPERS
'==============================================================================

' FileDialog folder picker; returns "" when cancelled
Private Function PickFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to standardize file names"
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

Private Function GetOrCreateSheet(ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = shName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub SpeedOn()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub SpeedOff()
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
