Attribute VB_Name = "FileOwnerReport"
Option Explicit

'==============================================================================
' MODULE : FileOwnerReport
' PURPOSE: List the files you can SEE in Explorer (folder + sub-folders) with
'          the same values as Windows > Properties > Details:
'            - File name
'            - Authors        (Details tab: Authors)
'            - Last Saved By  (Details tab: Last saved by)
'            - Date modified, Folder path, clickable Link
'          Sorted by PERSON first: Authors -> Last Saved By -> File name
'
' NOT LISTED (not visible in Explorer):
'          - hidden / system files and hidden / system folders
'          - "xxx_files" resource folders of a saved web page (xxx.htm / xxx.html)
'
' RUN    : ListFileOwners()   (Alt+F8)
' OUTPUT : sheet "File_Owner"  (cleared and rebuilt on every run)
'
' NOTES  : - Values are read with Shell.Application (FolderItem.ExtendedProperty),
'            the same source as the Details tab. Files are NOT opened in Office.
'          - Reading properties from a network drive can be slow.
'==============================================================================

Private Const SHEET_NAME As String = "File_Owner"
Private Const INCLUDE_SUBFOLDERS As Boolean = True   ' False = selected folder only

Public Sub ListFileOwners()

    Dim folderPath As String
    Dim fso As Object, shl As Object
    Dim fol As Object, sf As Object, f As Object
    Dim ns As Object, itm As Object
    Dim stack As Collection, folders As Collection
    Dim curDir As String, p As Variant
    Dim total As Long, k As Long, n As Long
    Dim res() As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim stt() As Variant, i As Long

    ' ---- 1. Pick folder ------------------------------------------------------
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder to list file owners"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set shl = CreateObject("Shell.Application")

    ' ---- 2. Pass 1: collect visible folders and count files ------------------
    Set stack = New Collection
    Set folders = New Collection
    stack.Add folderPath

    Do While stack.Count > 0
        curDir = stack(stack.Count)
        stack.Remove stack.Count
        If Right$(curDir, 1) <> "\" Then curDir = curDir & "\"

        ' folders we cannot access (permission...) are silently skipped
        On Error Resume Next
        Set fol = Nothing
        Set fol = fso.GetFolder(curDir)
        If Not fol Is Nothing Then
            folders.Add curDir
            total = total + fol.Files.Count
            If INCLUDE_SUBFOLDERS Then
                For Each sf In fol.SubFolders
                    ' skip hidden/system folders and web page "xxx_files" folders
                    If (sf.Attributes And 6) = 0 Then
                        If Not IsWebResourceFolder(fso, curDir, sf.Name) Then stack.Add sf.Path
                    End If
                Next sf
            End If
        End If
        On Error GoTo ErrHandler

        Application.StatusBar = "Counting files... " & total
    Loop

    If total = 0 Then
        RestoreApp
        MsgBox "No file found in the selected folder.", vbInformation
        Exit Sub
    End If

    ' ---- 3. Pass 2: read properties of every visible file --------------------
    ReDim res(1 To total, 1 To 7)
    k = 0

    For Each p In folders
        curDir = CStr(p)

        On Error Resume Next
        Set fol = Nothing
        Set ns = Nothing
        Set fol = fso.GetFolder(curDir)
        Set ns = shl.NameSpace(curDir)
        On Error GoTo ErrHandler

        If Not fol Is Nothing Then
            For Each f In fol.Files
                If k >= total Then Exit For         ' new file appeared meanwhile

                ' hidden / system files are not visible in Explorer -> not listed
                If (f.Attributes And 6) <> 0 Then GoTo NextFile

                k = k + 1
                res(k, 2) = f.Name
                res(k, 3) = ""
                res(k, 4) = ""
                res(k, 5) = f.DateLastModified
                res(k, 6) = curDir

                ' properties are read with error protection (some files deny it)
                On Error Resume Next
                If Not ns Is Nothing Then
                    Set itm = Nothing
                    Set itm = ns.ParseName(f.Name)
                    If Not itm Is Nothing Then
                        res(k, 3) = PropText(itm.ExtendedProperty("System.Author"))
                        res(k, 4) = PropText(itm.ExtendedProperty("System.Document.LastAuthor"))
                    End If
                End If
                On Error GoTo ErrHandler

                ' clickable link = Folder (col F) & File name (col B)
                res(k, 7) = "=HYPERLINK(F" & (k + 1) & "&B" & (k + 1) & ",""Open"")"

                If k Mod 100 = 0 Then Application.StatusBar = "Reading properties... " & k & " / " & total
NextFile:
            Next f
        End If
    Next p

    n = k
    If n = 0 Then
        RestoreApp
        MsgBox "No visible file found.", vbInformation
        Exit Sub
    End If

    ' ---- 4. Write result to sheet (one shot) -----------------------------------
    Set ws = GetOrCreateSheet(SHEET_NAME)
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("STT", "File Name", "Authors", _
                                    "Last Saved By", "Date Modified", "Folder Path", "Link")

    ws.Range("B2:D" & total + 1).NumberFormat = "@"
    ws.Range("F2:F" & total + 1).NumberFormat = "@"
    ws.Range("E2:E" & total + 1).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("A2").Resize(total, 7).Value = res

    ' hidden files were counted in "total" but not listed -> clear unused rows
    If n < total Then ws.Range("A" & n + 2 & ":G" & total + 1).Clear
    lastRow = n + 1

    ' ---- 5. Sort: PERSON first (Authors -> Last Saved By -> File Name) --------
    Application.Calculation = xlCalculationAutomatic
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("C2:C" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("D2:D" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("B2:B" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:G" & lastRow)
        .Header = xlYes
        .Apply
    End With

    ' numbering after sorting
    ReDim stt(1 To n, 1 To 1)
    For i = 1 To n
        stt(i, 1) = i
    Next i
    ws.Range("A2").Resize(n, 1).Value = stt

    ' ---- 6. Formatting ---------------------------------------------------------
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
    ws.Columns("A:G").AutoFit
    If ws.Columns("B").ColumnWidth > 60 Then ws.Columns("B").ColumnWidth = 60
    If ws.Columns("F").ColumnWidth > 70 Then ws.Columns("F").ColumnWidth = 70
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:G" & lastRow).AutoFilter
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    RestoreApp
    MsgBox "Done. " & n & " files listed." & vbCrLf & _
           "Sorted by: Authors > Last Saved By > File Name.", vbInformation
    Exit Sub

ErrHandler:
    RestoreApp
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

'==============================================================================
' HELPERS
'==============================================================================

' Property value -> text. "Authors" can be a list (array): join with "; "
Private Function PropText(ByVal v As Variant) As String
    On Error Resume Next
    If IsArray(v) Then
        PropText = Join(v, "; ")
    ElseIf Not IsNull(v) And Not IsEmpty(v) Then
        PropText = CStr(v)
    End If
End Function

' True when folder is the resource folder of a saved web page, e.g. folder
' "Report_files" next to "Report.htm" / "Report.html" (images, css, js inside)
Private Function IsWebResourceFolder(ByVal fso As Object, ByVal parentDir As String, _
                                     ByVal folderName As String) As Boolean
    Dim baseName As String
    If LCase$(Right$(folderName, 6)) <> "_files" Then Exit Function
    baseName = Left$(folderName, Len(folderName) - 6)
    If Len(baseName) = 0 Then Exit Function
    IsWebResourceFolder = fso.FileExists(parentDir & baseName & ".htm") _
                       Or fso.FileExists(parentDir & baseName & ".html")
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

Private Sub RestoreApp()
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
