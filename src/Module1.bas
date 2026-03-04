Option Explicit
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------

'========================
' Status Engine - Globals
'========================
Private mStartTimer As Single
Private mTotal As Long
Private mMode As String
Private mLastLogFolder As String
Private mLastLogPath As String
'Private mLastLogFolder As String
Private Const SKIP_BLANK_VALUES As Boolean = True
Private Const OVERWRITE_EXISTING As Boolean = True

Private Function Csv10( _
    ByVal fName As String, _
    ByVal fPath As String, _
    ByVal csvPath As String, _
    ByVal rowsRead As String, _
    ByVal applied As String, _
    ByVal skippedBlank As String, _
    ByVal csvMissing As String, _
    ByVal openErr As String, _
    ByVal saveErr As String, _
    ByVal saveWarn As String, _
    ByVal notes As String) As String
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
    Csv10 = _
        CsvQuote(fName) & "," & _
        CsvQuote(fPath) & "," & _
        CsvQuote(csvPath) & "," & _
        rowsRead & "," & _
        applied & "," & _
        skippedBlank & "," & _
        csvMissing & "," & _
        openErr & "," & _
        saveErr & "," & _
        saveWarn & "," & _
        CsvQuote(notes)
End Function
 
Private Sub btnImport_Click()
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
    On Error GoTo EH
 
    Dim modelFolder As String, csvFolder As String
    modelFolder = EnsureTrailingSlash(Trim$(txtModelFolderImp.Text))
    csvFolder = EnsureTrailingSlash(Trim$(txtCsvFolderImp.Text))
 
    If Len(modelFolder) = 0 Then
        MsgBox "Select Model Folder first.", vbExclamation
        Exit Sub
    End If
    
    If Len(csvFolder) = 0 Then
    MsgBox "Select CSV Folder first.", vbExclamation
    Exit Sub
End If
 
    If Len(csvFolder) = 0 Then
        MsgBox "Select CSV Folder first.", vbExclamation
        Exit Sub
    End If
 
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
 
    If Not fso.FolderExists(modelFolder) Then
        MsgBox "Model folder not found:" & vbCrLf & modelFolder, vbCritical
        Exit Sub
    End If
 
    If Not fso.FolderExists(csvFolder) Then
        MsgBox "CSV folder not found:" & vbCrLf & csvFolder, vbCritical
        Exit Sub
    End If
 
    'Count selected rows (IMPORT listbox)
    Dim totalChecked As Long
    totalChecked = CountCheckedRows_Import()
 
    If totalChecked = 0 Then
        MsgBox "No files selected. Tick rows then click Import.", vbInformation
        Exit Sub
    End If
 
    Dim swApp As Object: Set swApp = Application.SldWorks
 
    'Create LOG file (in CSV folder)
    Dim logPath As String
    logPath = csvFolder & "Import_Log_" & SafeNowStamp() & ".csv"
 
    Dim logTs As Object
    Set logTs = fso.CreateTextFile(logPath, True, False) 'Unicode=False
 
    logTs.WriteLine "FileName,ModelPath,CSV,RowsRead,Applied,SkippedBlank,CSV_NotFound,OpenErr,SaveErrs,SaveWarns,Notes"
 
    '==== STATUS BAR (BEGIN) ====
    Status_Begin "Import", totalChecked, logPath
 
    Dim done As Long: done = 0
    Dim i As Long
 
    For i = 0 To lstExportPreview.ListCount - 1
 
        If lstExportPreview.Selected(i) Then
 
            Dim fileName As String, fullPath As String
            fileName = CStr(lstExportPreview.List(i, 0))
            fullPath = CStr(lstExportPreview.List(i, 2))
 
            'Build CSV path by BaseName
            Dim csvPath As String
            csvPath = csvFolder & CleanBaseName(fileName) & ".csv"
 
            SetRowStatus_Import i, "Importing..."
 
            done = done + 1
            Status_Update done, "Importing: " & fileName
 
            'CSV missing?
            If Dir(csvPath) = "" Then
                SetRowStatus_Import i, "CSV not found"
                logTs.WriteLine Csv10(fileName, fullPath, csvPath, "0", "0", "0", "1", "0", "0", "0", "CSV not found")
                DoEvents
                GoTo NextRow
            End If
 
            'Open doc
            Dim ext As String
            ext = LCase$(fso.GetExtensionName(fileName))
 
            Dim docType As Long
            If ext = "sldasm" Then
                docType = 2
            Else
                docType = 1
            End If
 
            Dim openErr As Long, openWarn As Long
            Dim doc As Object
            Set doc = swApp.OpenDoc6(fullPath, docType, 64, "", openErr, openWarn) 'silent
 
            If doc Is Nothing Then
                SetRowStatus_Import i, "Open failed"
                logTs.WriteLine Csv10(fileName, fullPath, csvPath, "0", "0", "0", "0", CStr(openErr), "0", "0", "Open failed")
                DoEvents
                GoTo NextRow
            End If
 
            'Run your EXISTING core import function
            Dim result As Variant
            result = Import2ColCsv_ToActiveDoc_Smart(doc, csvPath) 'Array(rowsRead, applied, skippedBlank, notes)
 
            doc.ForceRebuild3 False
 
            Dim errs As Long, warns As Long
            errs = 0: warns = 0
 
            On Error Resume Next
            doc.Save3 1, errs, warns
            On Error GoTo EH
 
            swApp.CloseDoc doc.GetTitle
 
            'Row status
            If errs = 0 Then
                SetRowStatus_Import i, "Imported (" & result(1) & ")"
            Else
                SetRowStatus_Import i, "Save error"
            End If
 
            logTs.WriteLine Csv10(fileName, fullPath, csvPath, _
                                  CStr(result(0)), CStr(result(1)), CStr(result(2)), _
                                  "0", CStr(openErr), CStr(errs), CStr(warns), CStr(result(3)))
 
            DoEvents
        End If
 
NextRow:
    Next i
 
    logTs.Close
 
    '==== STATUS BAR END ====
    Status_End "Import completed."
 
    MsgBox "Import done!" & vbCrLf & "Log: " & logPath, vbInformation
    Exit Sub
 
EH:
    On Error Resume Next
    Status_End "Import failed."
    lblStatusMsg.Caption = "Import error: " & Err.Description
    MsgBox "Import Error: " & Err.Description, vbCritical
 
End Sub

'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' CORE IMPORT: CSV -> Custom Properties (Doc-level)
' Required by: btnImport_Click
'===========================================================
Private Function Import2ColCsv_ToActiveDoc_Smart(ByVal doc As Object, ByVal csvPath As String) As Variant

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    'Read as ASCII (NOT Unicode)
    Dim ts As Object: Set ts = fso.OpenTextFile(csvPath, 1, False, 0)
 
    Dim rowsRead As Long: rowsRead = 0
    Dim applied As Long: applied = 0
    Dim skippedBlank As Long: skippedBlank = 0
    Dim notes As String: notes = ""
 
    If ts.AtEndOfStream Then
        Import2ColCsv_ToActiveDoc_Smart = Array(0, 0, 0, "Empty file")
        Exit Function
    End If
 
    Dim header As String
    header = ts.ReadLine 'header line
 
    Dim cpm As Object: Set cpm = doc.Extension.CustomPropertyManager("")
 
    Do While Not ts.AtEndOfStream
 
        Dim line As String: line = ts.ReadLine
        If Len(Trim$(line)) = 0 Then GoTo ContinueLoop
 
        Dim pName As String, pVal As String
        If Not Parse2Cols(line, pName, pVal) Then GoTo ContinueLoop
 
        rowsRead = rowsRead + 1
 
        If SKIP_BLANK_VALUES And Len(Trim$(pVal)) = 0 Then
            skippedBlank = skippedBlank + 1
            GoTo ContinueLoop
        End If
 
        Dim ok As Boolean
        ok = UpsertCustomProp(cpm, pName, pVal)
 
        If ok Then
            applied = applied + 1
        Else
            notes = notes & "Failed:" & pName & "; "
        End If
 
ContinueLoop:
    Loop
 
    ts.Close
 
    If rowsRead = 0 Then notes = "Parsed 0 rows (check delimiter/format)"
 
    Import2ColCsv_ToActiveDoc_Smart = Array(rowsRead, applied, skippedBlank, notes)
 
End Function

Private Function Parse2Cols(ByVal line As String, ByRef col1 As String, ByRef col2 As String) As Boolean
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
    Dim delim As String
    If InStr(line, vbTab) > 0 Then
        delim = vbTab
    ElseIf InStr(line, ";") > 0 Then
        delim = ";"
    ElseIf InStr(line, ",") > 0 Then
        delim = ","
    Else
        Parse2Cols = False
        Exit Function
    End If
 
    Dim arr As Variant
    If InStr(line, """") > 0 Then
        arr = SplitQuoted(line, delim)
    Else
        arr = Split(line, delim)
    End If
 
    If IsEmpty(arr) Then Parse2Cols = False: Exit Function
    If UBound(arr) < 1 Then Parse2Cols = False: Exit Function
 
    col1 = CleanQuotes(CStr(arr(0)))
    col2 = CleanQuotes(CStr(arr(1)))
 
    Parse2Cols = True
 
End Function
 
Private Function SplitQuoted(ByVal s As String, ByVal delim As String) As Variant
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
    Dim out() As String, i As Long, c As String, inQ As Boolean, cur As String
    ReDim out(0): cur = "": inQ = False
 
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
 
        If c = """" Then
            If inQ And i < Len(s) And Mid$(s, i + 1, 1) = """" Then
                cur = cur & """": i = i + 1
            Else
                inQ = Not inQ
            End If
 
        ElseIf c = delim And Not inQ Then
            out(UBound(out)) = cur
            ReDim Preserve out(UBound(out) + 1)
            cur = ""
 
        Else
            cur = cur & c
        End If
    Next i
 
    out(UBound(out)) = cur
    SplitQuoted = out
 
End Function
 
Private Function CleanQuotes(ByVal s As String) As String
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
    s = Trim$(s)
    If Len(s) >= 2 Then
        If Left$(s, 1) = """" And Right$(s, 1) = """" Then
            s = Mid$(s, 2, Len(s) - 2)
            s = Replace(s, """""", """")
        End If
    End If
    CleanQuotes = s
End Function

Private Function UpsertCustomProp(ByVal cpm As Object, ByVal pName As String, ByVal pVal As String) As Boolean
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
    On Error GoTo Fail
 
   
    On Error Resume Next
    cpm.Delete2 pName
    On Error GoTo 0
 
    'Add (Text=30, DeleteAndAdd=2)
    Dim ret As Long
    ret = cpm.Add3(pName, 30, pVal, 2)
 
    'Verify
    Dim vRaw As String, vRes As String, wasResolved As Boolean
    Dim ok As Boolean
    ok = cpm.Get5(pName, False, vRaw, vRes, wasResolved)
 
    UpsertCustomProp = ok
    Exit Function
 
Fail:
    UpsertCustomProp = False
End Function
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
Private Function CountCheckedRows_Import() As Long
    Dim i As Long, c As Long
    For i = 0 To lstExportPreview.ListCount - 1
        If lstExportPreview.Selected(i) Then c = c + 1
    Next i
    CountCheckedRows_Import = c
End Function

'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
Private Sub SetRowStatus_Import(ByVal rowIndex As Long, ByVal statusText As String)
    lstExportPreview.List(rowIndex, 3) = statusText
End Sub

Private Sub lstImportPreview_Click()
 Footer_RefreshCounts
End Sub

 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------

'===========================================================
' FORM INITIALIZATION
'===========================================================
Private Sub UserForm_Initialize()
    On Error Resume Next
    lstImportPreview.Font.Name = "Segoe UI"
    On Error GoTo 0
 
    'Enable checkbox-style selection
    lstImportPreview.MultiSelect = fmMultiSelectMulti
 
    FitImportPreviewColumns
    
    '---- IMPORT tab listbox setup ----
lstExportPreview.Font.Name = "Segoe UI"
lstExportPreview.MultiSelect = fmMultiSelectMulti
FitExportPreviewColumns
 
    Status_Reset   'reset progress bar on open
    
End Sub

Private Sub FitExportPreviewColumns()
    On Error Resume Next
    'Same style as Export listbox. Adjust widths if needed.
    lstExportPreview.ColumnCount = 4
    lstExportPreview.ColumnWidths = "120 pt;30 pt;480 pt;90 pt"
    On Error GoTo 0
End Sub
 
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' RESIZE EVENT ? Keep columns fitted
'===========================================================
Private Sub UserForm_Resize()
    On Error Resume Next
    FitImportPreviewColumns
End Sub
 
 
'===========================================================
' RESET BUTTON
'===========================================================
Private Sub btnReset_Click()
    txtModelFolderExp.Value = ""
    txtCsvFolderExp.Value = ""
    lstImportPreview.Clear
    mLastLogPath = ""
    mLastLogFolder = ""
 
    'reset footer completely (removes log hyperlink)
    Status_Reset
End Sub
 
 
'===========================================================
' COLUMN FIT (Avoid Horizontal Scrollbar)
'===========================================================
Private Sub FitImportPreviewColumns()
 
    Dim w As Single
    w = lstImportPreview.Width - 35   'padding
 
    Dim wFile As Single, wType As Single, wStatus As Single, wPath As Single
    wFile = w * 0.28
    wType = w * 0.1
    wStatus = w * 0.17
    wPath = w - (wFile + wType + wStatus)
 
    With lstImportPreview
        .ColumnCount = 4
        .IntegralHeight = False
        .ColumnWidths = _
            CStr(Int(wFile)) & " pt;" & _
            CStr(Int(wType)) & " pt;" & _
            CStr(Int(wPath)) & " pt;" & _
            CStr(Int(wStatus)) & " pt"
    End With
 
End Sub
 
 
'===========================================================
' BROWSE MODEL FOLDER ? auto-list files
'===========================================================
Private Sub btnBrowseModel_Exp_Click()
 
    Dim p As String
    p = BrowseForFolder("Select MODEL folder")
    If Len(p) = 0 Then Exit Sub
 
    txtModelFolderExp.Text = EnsureTrailingSlash(p)
    ListSolidWorksFiles txtModelFolderExp.Text
 
End Sub
 
 
'===========================================================
' BROWSE CSV OUTPUT FOLDER
'===========================================================
Private Sub btnBrowseCsv_Exp_Click()
 
    Dim p As String
    p = BrowseForFolder("Select CSV OUTPUT folder")
    If Len(p) = 0 Then Exit Sub
 
    txtCsvFolderExp.Text = EnsureTrailingSlash(p)
 
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
 
'===========================================================
' LIST ONLY SW FILES (.sldprt / .sldasm)
'===========================================================
Private Sub ListSolidWorksFiles(ByVal folderPath As String)
 
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
 
    folderPath = EnsureTrailingSlash(Trim$(folderPath))
 
    'Clear previous list
    lstImportPreview.Clear
 
    'Footer reset (no progress bar needed here)
    Status_Reset
    lblLog.Caption = "Log: -"
    lblStatusMsg.Caption = "Scanning: " & folderPath
    lblProgress.Caption = "Progress: Scanning..."
    Progress_SetPercent 0
    DoEvents
 
    'Validate
    If Len(folderPath) = 0 Then
        lblStatusMsg.Caption = "No folder selected."
        lblProgress.Caption = "Progress: 0/0 files"
        Exit Sub
    End If
 
    If Not fso.FolderExists(folderPath) Then
        lblStatusMsg.Caption = "Folder not found."
        lblProgress.Caption = "Progress: 0/0 files"
        MsgBox "Model folder not found:" & vbCrLf & folderPath, vbCritical
        Exit Sub
    End If
 
    'Scan files
    Dim fol As Object
    Set fol = fso.GetFolder(folderPath)
 
    Dim fil As Object
    Dim swCount As Long: swCount = 0
    Dim otherCount As Long: otherCount = 0
 
    For Each fil In fol.Files
     '=== SKIP TEMP / LOCK / BACKUP FILES ===
        If Left$(fil.Name, 2) = "~$" Or Left$(fil.Name, 1) = "~" Then
    GoTo Nextfil
        End If
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(fil.Name))
 
        If ext = "sldprt" Or ext = "sldasm" Then
 
            Dim typ As String
            If ext = "sldprt" Then
                typ = "PRT"
            Else
                typ = "ASM"
            End If
 
            AddRowImport fil.Name, typ, fil.Path, "Ready"
            swCount = swCount + 1
 
        Else
            otherCount = otherCount + 1
        End If
Nextfil:
    Next fil
 
    'Final footer update (Checked/Total style)
    Progress_SetPercent 0
 
    If swCount = 0 Then
        lblStatusMsg.Caption = "No SolidWorks files found in this folder."
        lblProgress.Caption = "Progress: 0/0 files"
        Beep
    Else
        'No auto-check here (user can tick manually)
        lblStatusMsg.Caption = "Found " & swCount & " SolidWorks file(s)."
        lblProgress.Caption = "Progress: 0/" & swCount & " files"
    End If

End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
 '===========================================================
' ADD ONE ROW TO lstImportPreview
'===========================================================
Private Sub AddRowImport(ByVal fileName As String, _
                         ByVal fileType As String, _
                         ByVal fullPath As String, _
                         ByVal statusText As String)
 
    With lstImportPreview
        Dim newRow As Long
        newRow = .ListCount
 
        .AddItem fileName
        .List(newRow, 1) = fileType
        .List(newRow, 2) = fullPath
        .List(newRow, 3) = statusText
    End With
 
End Sub
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' ADD ONE ROW
'===========================================================
Private Sub AddRow(ByVal fileName As String, ByVal fileType As String, ByVal fullPath As String, ByVal statusText As String)
 
    Dim r As Long
    lstImportPreview.AddItem fileName
    r = lstImportPreview.ListCount - 1
 
    lstImportPreview.List(r, 1) = fileType
    lstImportPreview.List(r, 2) = fullPath
    lstImportPreview.List(r, 3) = statusText
 
    lstImportPreview.Selected(r) = False   'unchecked by default
 
End Sub

 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
 
'===========================================================
' SELECT ALL
'===========================================================
Private Sub btnSelectAllExport_Click()
    Dim i As Long
    For i = 0 To lstImportPreview.ListCount - 1
        lstImportPreview.Selected(i) = True
    Next i
    Footer_RefreshCounts
End Sub
 
 
'===========================================================
' CLEAR ALL
'===========================================================
Private Sub btnClearExport_Click()
    Dim i As Long
    For i = 0 To lstImportPreview.ListCount - 1
        lstImportPreview.Selected(i) = False
    Next i
    Footer_RefreshCounts
End Sub
 
 
'===========================================================
' REMOVE ONLY SELECTED ROWS
'===========================================================
Private Sub btnRemove_Click()
 
    Dim i As Long, removed As Long
    removed = 0
 
    'Remove selected rows (your checkbox selection)
    For i = lstImportPreview.ListCount - 1 To 0 Step -1
        If lstImportPreview.Selected(i) Then
            lstImportPreview.RemoveItem i
            removed = removed + 1
        End If
    Next i
 
    If removed = 0 Then
        Footer_RefreshCounts "No files selected to remove."
        MsgBox "No files selected to remove.", vbInformation
    Else
        Footer_RefreshCounts "Removed " & removed & " file(s)."
    End If
 
End Sub

'===========================================================
' EXPORT BUTTON (Exports properties to CSV for selected rows)
'===========================================================
Private Sub btnExport_Click()
 
    On Error GoTo EH
 
    Dim modelFolder As String, csvFolder As String
    modelFolder = EnsureTrailingSlash(Trim$(txtModelFolderExp.Text))
    csvFolder = EnsureTrailingSlash(Trim$(txtCsvFolderExp.Text))
 
    If Len(modelFolder) = 0 Then
        MsgBox "Select Model Folder first.", vbExclamation
        Exit Sub
    End If
 
    If Len(csvFolder) = 0 Then
        MsgBox "Select CSV Output Folder first.", vbExclamation
        Exit Sub
    End If
 
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
 
    If Not fso.FolderExists(modelFolder) Then
        MsgBox "Model folder not found:" & vbCrLf & modelFolder, vbCritical
        Exit Sub
    End If
 
    'Create CSV folder if missing
    If Not fso.FolderExists(csvFolder) Then
        On Error Resume Next
        fso.CreateFolder csvFolder
        On Error GoTo EH
        If Not fso.FolderExists(csvFolder) Then
            MsgBox "CSV folder cannot be created:" & vbCrLf & csvFolder, vbCritical
            Exit Sub
        End If
    End If
 
    'Count how many rows are selected
    Dim totalChecked As Long
    totalChecked = CountCheckedRows()
    If totalChecked = 0 Then
        MsgBox "No files selected. Tick rows then click Export.", vbInformation
        Exit Sub
    End If
 
    Dim swApp As Object: Set swApp = Application.SldWorks
 
    'Create LOG file
    Dim logPath As String
    logPath = csvFolder & "Export_Log_" & SafeNowStamp() & ".csv"
 
    Dim logTs As Object
    Set logTs = fso.CreateTextFile(logPath, True, False)
    logTs.WriteLine "FileName,FullPath,CSV,PropCount,OpenErr,Notes"
 
    '==== STATUS BAR (BEGIN) ====
    Status_Begin "Export", totalChecked, logPath
 
    Dim done As Long: done = 0
    Dim i As Long
 
    '=== MAIN LOOP: Process selected rows only ===
    For i = 0 To lstImportPreview.ListCount - 1
 
        If lstImportPreview.Selected(i) Then
 
            Dim fileName As String, fullPath As String
            fileName = CStr(lstImportPreview.List(i, 0))
            fullPath = CStr(lstImportPreview.List(i, 2))
 
            SetRowStatus i, "Exporting..."
 
            done = done + 1
            Status_Update done, "Exporting: " & fileName
 
            Dim outCsv As String
            outCsv = csvFolder & CleanBaseName(fileName) & ".csv"
 
            Dim propCount As Long, openErr As Long, notes As String
            propCount = 0: openErr = 0: notes = ""
 
            Dim ok As Boolean
            ok = ExportOneFile(swApp, fullPath, outCsv, propCount, openErr, notes)
 
            If ok Then
                SetRowStatus i, "Exported (" & propCount & ")"
            Else
                SetRowStatus i, "Failed"
            End If
 
            logTs.WriteLine Csv6(fileName, fullPath, outCsv, CStr(propCount), CStr(openErr), notes)
 
            DoEvents
        End If
    Next i
 
    logTs.Close
 
    '==== STATUS BAR END ====
    Status_End "Export completed."
 
    MsgBox "Export done!" & vbCrLf & "Log: " & logPath, vbInformation
    Exit Sub
 
EH:
    On Error Resume Next
    Status_End "Export failed."
    lblStatusMsg.Caption = "Export error: " & Err.Description
    MsgBox "Export Error: " & Err.Description, vbCritical
 
End Sub
 
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' SET STATUS IN LIST ROW (Column 3)
'===========================================================
Private Sub SetRowStatus(ByVal rowIndex As Long, ByVal statusText As String)
    If rowIndex < 0 Or rowIndex >= lstImportPreview.ListCount Then Exit Sub
    lstImportPreview.List(rowIndex, 3) = statusText
End Sub
 
 
'===========================================================
' COUNT SELECTED ROWS
'===========================================================
Private Function CountCheckedRows() As Long
    Dim i As Long, cnt As Long
    cnt = 0
    For i = 0 To lstImportPreview.ListCount - 1
        If lstImportPreview.Selected(i) Then cnt = cnt + 1
    Next i
    CountCheckedRows = cnt
End Function

'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------

'===========================================================
' EXPORT ONE FILE (Open -> Export Props -> Close)
'===========================================================
Private Function ExportOneFile( _
    ByVal swApp As Object, _
    ByVal filePath As String, _
    ByVal outCsv As String, _
    ByRef propCount As Long, _
    ByRef openErr As Long, _
    ByRef notes As String) As Boolean
 
    On Error GoTo Fail
 
    propCount = 0
    openErr = 0
    notes = ""
 
    Dim ext As String
    ext = LCase$(GetExtension(filePath))
 
    Dim docType As Long
    If ext = "sldasm" Then
        docType = 2
    ElseIf ext = "sldprt" Then
        docType = 1
    Else
        notes = "Not a SW model"
        ExportOneFile = False
        Exit Function
    End If
 
    Dim openWarn As Long
    openWarn = 0
 
    Dim doc As Object
    Set doc = swApp.OpenDoc6(filePath, docType, 64, "", openErr, openWarn) 'silent open
 
    If doc Is Nothing Then
        notes = "Open failed"
        ExportOneFile = False
        Exit Function
    End If
 
    propCount = ExportFileLevelProps_To2ColCsv(doc, outCsv)
 
    swApp.CloseDoc doc.GetTitle
 
    If propCount = 0 Then
        notes = "0 props (no custom properties?)"
    End If
 
    ExportOneFile = True
    Exit Function
 
Fail:
    notes = "Exception: " & Err.Description
    On Error Resume Next
    If Not doc Is Nothing Then swApp.CloseDoc doc.GetTitle
    On Error GoTo 0
    ExportOneFile = False
 
End Function
 
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------

'===========================================================
' EXPORT PROPERTIES TO 2-COLUMN CSV
'===========================================================
Private Function ExportFileLevelProps_To2ColCsv(ByVal doc As Object, ByVal outCsv As String) As Long
 
    On Error GoTo EH
 
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object: Set ts = fso.CreateTextFile(outCsv, True, False) 'Unicode=False
 
    ts.WriteLine "PropName,PropValue"
 
    Dim cpm As Object: Set cpm = doc.Extension.CustomPropertyManager("")
    Dim names As Variant: names = cpm.GetNames
 
    Dim cnt As Long: cnt = 0
 
    If Not IsEmpty(names) Then
 
        Dim i As Long
        For i = LBound(names) To UBound(names)
 
            Dim pName As String: pName = CStr(names(i))
            Dim valOut As String, resolved As String
            Dim wasResolved As Boolean, link As Boolean
 
            cpm.Get6 pName, False, valOut, resolved, wasResolved, link
 
            Dim finalVal As String
            finalVal = resolved
            If Len(finalVal) = 0 Then finalVal = valOut
 
            ts.WriteLine Csv2(pName, finalVal)
            cnt = cnt + 1
 
        Next i
    End If
 
    ts.Close
    ExportFileLevelProps_To2ColCsv = cnt
    Exit Function
 
EH:
    ExportFileLevelProps_To2ColCsv = 0
 
End Function
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------

'===========================================================
' STATUS BAR ENGINE (COMMON FOR EXPORT + IMPORT)
'
' Required controls on form:
'   lblStatusMsg
'   lblProgress
'   lblElapsed
'   lblLog
'   fraBar       (Frame)
'   lblBarFill   (Label inside frame)
'===========================================================
 
Private Sub Status_Reset()
    lblStatusMsg.Caption = ""
    lblProgress.Caption = "Progress: -"
    lblElapsed.Caption = "Elapsed: 00:00"
 
    '? reset log + remove hyperlink style
    lblLog.Caption = "Log: -"
    lblLog.Font.Underline = False
    lblLog.ForeColor = vbBlack
 
    Progress_SetPercent 0
End Sub
 
Private Sub Status_Begin(ByVal modeText As String, ByVal totalItems As Long, ByVal logPath As String)
 
    mMode = modeText
    mTotal = totalItems
    mStartTimer = Timer
 
    mLastLogPath = logPath
    mLastLogFolder = GetParentFolder(logPath)
 
    'normal (not hyperlink look)
    lblLog.Caption = "Log: " & ShortFileName(logPath)
    lblLog.Font.Underline = False
    lblLog.ForeColor = vbBlack
 
    lblProgress.Caption = "Progress: 0/" & CStr(mTotal) & " files"
    lblElapsed.Caption = "Elapsed: 00:00"
    lblStatusMsg.Caption = mMode & " started..."
 
    Progress_SetPercent 0
    DoEvents
End Sub
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
Private Function GetParentFolder(ByVal fullPath As String) As String
    Dim p As Long
    p = InStrRev(fullPath, "\")
    If p > 0 Then
        GetParentFolder = Left$(fullPath, p - 1)
    Else
        GetParentFolder = ""
    End If
End Function

Private Sub lblLog_Click()
    If Len(mLastLogFolder) = 0 Then Exit Sub
    Shell "explorer.exe """ & mLastLogFolder & """", vbNormalFocus
End Sub
 
Private Sub Status_Update(ByVal done As Long, ByVal msg As String)
 
    Dim pct As Single
    If mTotal <= 0 Then
        pct = 0
    Else
        pct = (done / mTotal) * 100
    End If
 
    lblProgress.Caption = "Progress: " & done & "/" & mTotal & " files"
    lblElapsed.Caption = "Time: " & FormatElapsed(TimerSafeDiff(mStartTimer))
    lblStatusMsg.Caption = msg
 
    Progress_SetPercent pct
    DoEvents
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
Private Sub Status_End(ByVal finalMsg As String)
 
    lblElapsed.Caption = "Done in: " & FormatElapsed(TimerSafeDiff(mStartTimer))
    lblStatusMsg.Caption = finalMsg
    Progress_SetPercent 100
 
    '? now make it look like hyperlink
    lblLog.Caption = "Log: " & ShortFileName(mLastLogPath)
    lblLog.Font.Underline = True
    lblLog.ForeColor = vbBlue
 
    DoEvents
End Sub
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' PROGRESS BAR FILL (lblBarFill inside fraBar)
'===========================================================
Private Sub Progress_SetPercent(ByVal pct As Single)
 
    If pct < 0 Then pct = 0
    If pct > 100 Then pct = 100
 
    lblBarFill.Left = 0
    lblBarFill.Top = 0
    lblBarFill.Height = fraBar.InsideHeight
    lblBarFill.Width = (fraBar.InsideWidth * pct) / 100
End Sub
 
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' TIMER / ELAPSED TIME HELPERS
'===========================================================
Private Function TimerSafeDiff(ByVal startT As Single) As Single
    Dim t As Single
    t = Timer - startT
 
    'Handles Timer rollover @ midnight (Timer resets)
    If t < 0 Then t = t + 86400
 
    TimerSafeDiff = t
End Function
 
Private Function FormatElapsed(ByVal sec As Single) As String
    Dim s As Long: s = CLng(sec)
    FormatElapsed = Format$(s \ 60, "00") & ":" & Format$(s Mod 60, "00")
End Function
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
 
'===========================================================
' SIMPLE FILE NAME (LAST PART ONLY)
'===========================================================
Private Function ShortFileName(ByVal p As String) As String
    Dim i As Long: i = InStrRev(p, "\")
    If i > 0 Then ShortFileName = Mid$(p, i + 1) Else ShortFileName = p
End Function
 
 
'===========================================================
' CSV HELPERS
'===========================================================
Private Function Csv2(ByVal a As String, ByVal b As String) As String
    Csv2 = CsvQuote(a) & "," & CsvQuote(b)
End Function
 
Private Function Csv6(ByVal a As String, ByVal b As String, ByVal c As String, _
                      ByVal d As String, ByVal e As String, ByVal f As String) As String
 
    Csv6 = CsvQuote(a) & "," & CsvQuote(b) & "," & CsvQuote(c) & "," & _
           CsvQuote(d) & "," & CsvQuote(e) & "," & CsvQuote(f)
End Function
 
Private Function CsvQuote(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, """", """""")
    CsvQuote = """" & s & """"
End Function
 
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' PATH / NAME HELPERS
'===========================================================
Private Function GetExtension(ByVal fullPathOrName As String) As String
    Dim p As Long: p = InStrRev(fullPathOrName, ".")
    If p > 0 Then GetExtension = Mid$(fullPathOrName, p + 1) _
    Else GetExtension = ""
End Function
 
Private Function CleanBaseName(ByVal s As String) As String
    s = Trim$(Replace(s, "*", ""))
 
    Dim dotPos As Long: dotPos = InStrRev(s, ".")
    If dotPos > 0 Then s = Left$(s, dotPos - 1)
 
    CleanBaseName = s
End Function
 
Private Function SafeNowStamp() As String
    Dim t As String
    t = CStr(Now)
 
    t = Replace(t, ":", "")
    t = Replace(t, "/", "")
    t = Replace(t, " ", "_")
 
    SafeNowStamp = t
End Function
 
 
'===========================================================
' FOLDER BROWSE HELPERS
'===========================================================
Private Function BrowseForFolder(ByVal title As String) As String
    On Error GoTo EH
 
    Dim sh As Object, f As Object
    Set sh = CreateObject("Shell.Application")
    Set f = sh.BrowseForFolder(0, title, 0, 0)
 
    If Not f Is Nothing Then
        BrowseForFolder = f.Self.Path
    Else
        BrowseForFolder = ""
    End If
    Exit Function
 
EH:
    BrowseForFolder = ""
End Function
 
Private Function EnsureTrailingSlash(ByVal p As String) As String
    p = Trim$(p)
 
    If Len(p) = 0 Then
        EnsureTrailingSlash = ""
    ElseIf Right$(p, 1) = "\" Then
        EnsureTrailingSlash = p
    Else
        EnsureTrailingSlash = p & "\"
    End If
End Function
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' FOOTER REFRESH (call after add/remove/select/clear)
' - Always updates Progress
' - Always updates lblStatusMsg (default message)
' - If caller passes statusMsg, it overrides default
'===========================================================
'===========================================================
' FOOTER REFRESH (call after add/remove/select/clear/click)
'===========================================================
Private Sub Footer_RefreshCounts(Optional ByVal statusMsg As String = "")
 
    Dim totalRows As Long
    Dim checkedRows As Long
 
    totalRows = lstImportPreview.ListCount
    checkedRows = CountCheckedRows()
 
    'Progress line (this one is already working for you)
    lblProgress.Caption = "Progress: " & checkedRows & "/" & totalRows & " files"
 
    'TOP BOLD LINE (lblStatusMsg) - ALWAYS refresh based on current list
    If totalRows = 0 Then
        lblStatusMsg.Caption = "No SolidWorks files found in this folder."
        Progress_SetPercent 0
    Else
        lblStatusMsg.Caption = "Found " & totalRows & " SolidWorks file(s)."
        Progress_SetPercent 0
    End If
 
    'If caller gave a custom message, override the default
    If Len(Trim$(statusMsg)) > 0 Then
        lblStatusMsg.Caption = statusMsg
    End If
 
    DoEvents
End Sub
'------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' IMPORT TAB — LISTBOX SETUP (Same 4 columns)
' Col0: File Name
' Col1: Type (PRT/ASM)
' Col2: Full Path
' Col3: Status (Ready / CSV Missing)
'
' Controls:
'   txtModelFolderImp, btnBrowseModel_Imp
'   txtCsvFolderImp,   btnBrowseCsv_Imp
'   lstExportPreview
'   btnImport, btnRemoveImport, btnSelectAllImport, btnClearImport, btnResetImport
' Footer (shared):
'   lblStatusMsg, lblProgress, lblElapsed, lblLog, fraBar, lblBarFill
'===========================================================
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'-------------------------------
' Fit columns - avoid horizontal bar
'-------------------------------
Private Sub Imp_FitColumns()
    Dim w As Single
    w = lstExportPreview.Width - 30 'padding
 
    Dim wFile As Single, wType As Single, wStatus As Single, wPath As Single
    wFile = w * 0.28
    wType = w * 0.1
    wStatus = w * 0.18
    wPath = w - (wFile + wType + wStatus)
 
    With lstExportPreview
        .ColumnCount = 4
        .IntegralHeight = False
        .ColumnWidths = _
            CStr(Int(wFile)) & " pt;" & _
            CStr(Int(wType)) & " pt;" & _
            CStr(Int(wPath)) & " pt;" & _
            CStr(Int(wStatus)) & " pt"
    End With
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' BROWSE MODEL FOLDER (Import) -> list files
'===========================================================
Private Sub btnBrowseModel_Imp_Click()
    Dim p As String
    p = BrowseForFolder("Select MODEL folder (Import)")
    If Len(p) = 0 Then Exit Sub
 
    txtModelFolderImp.Text = EnsureTrailingSlash(p)
 
    'Auto list after selecting model folder
    Imp_ListFiles
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' BROWSE CSV FOLDER (Import) -> update statuses
'===========================================================
Private Sub btnBrowseCsv_Imp_Click()
    Dim p As String
    p = BrowseForFolder("Select CSV folder (Import)")
    If Len(p) = 0 Then Exit Sub
 
    txtCsvFolderImp.Text = EnsureTrailingSlash(p)
 
    'If list already exists, refresh Ready/CSV Missing
    Imp_RefreshCsvStatus
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' LIST FILES (only SW) + status check using CSV folder
'===========================================================
Private Sub Imp_ListFiles()
 
    Dim modelFolder As String, csvFolder As String
    modelFolder = EnsureTrailingSlash(Trim$(txtModelFolderImp.Text))
    csvFolder = EnsureTrailingSlash(Trim$(txtCsvFolderImp.Text))
 
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
 
    lstExportPreview.Clear
    lblStatusMsg.Caption = "Scanning: " & modelFolder
    lblProgress.Caption = "Progress: Scanning..."
    Progress_SetPercent 0
    DoEvents
 
    If Len(modelFolder) = 0 Then
        lblStatusMsg.Caption = "No model folder selected."
        lblProgress.Caption = "Progress: 0/0 files"
        Exit Sub
    End If
 
    If Not fso.FolderExists(modelFolder) Then
        lblStatusMsg.Caption = "Model folder not found."
        lblProgress.Caption = "Progress: 0/0 files"
        Exit Sub
    End If
 
    Dim fol As Object: Set fol = fso.GetFolder(modelFolder)
    Dim fil As Object
    Dim swCount As Long: swCount = 0
 
    For Each fil In fol.Files
    '=== SKIP TEMP / LOCK / BACKUP FILES ===
    If Left$(fil.Name, 2) = "~$" Or Left$(fil.Name, 1) = "~" Then
    GoTo Nextfil
    End If
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(fil.Name))
 
        If ext = "sldprt" Or ext = "sldasm" Then
            Dim typ As String
            If ext = "sldprt" Then typ = "PRT" Else typ = "ASM"
 
            Dim st As String
            st = "Ready"
 
            'If CSV folder already chosen, validate csv exists
            If Len(csvFolder) > 0 Then
                Dim baseName As String
                baseName = CleanBaseName(fil.Name)
                If Dir$(csvFolder & baseName & ".csv") = "" Then
                    st = "CSV Missing"
                End If
            End If
 
            Imp_AddRow fil.Name, typ, fil.Path, st
            swCount = swCount + 1
        End If
        
Nextfil:
    Next fil
 
    'Footer summary
    If swCount = 0 Then
        lblStatusMsg.Caption = "No SolidWorks files found in this folder."
        lblProgress.Caption = "Progress: 0/0 files selected"
        Beep
    Else
        lblStatusMsg.Caption = "Found " & swCount & " SolidWorks file(s)."
        lblProgress.Caption = "Progress: 0/" & swCount & " files selected"
    End If
 
    Progress_SetPercent 0
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' Refresh status column based on CSV folder after it is chosen
'===========================================================
Private Sub Imp_RefreshCsvStatus()
    Dim csvFolder As String
    csvFolder = EnsureTrailingSlash(Trim$(txtCsvFolderImp.Text))
    If Len(csvFolder) = 0 Then Exit Sub
 
    Dim i As Long, baseName As String, fileName As String
 
    For i = 0 To lstExportPreview.ListCount - 1
        fileName = CStr(lstExportPreview.List(i, 0))
        baseName = CleanBaseName(fileName)
 
        If Dir$(csvFolder & baseName & ".csv") = "" Then
            lstExportPreview.List(i, 3) = "CSV Missing"
        Else
            lstExportPreview.List(i, 3) = "Ready"
        End If
    Next i
 
    lblStatusMsg.Caption = "CSV status refreshed."
    DoEvents
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' CLICK ROW -> toggle selection (checkbox behavior)
'===========================================================
Private Sub lstExportPreview_Click()
    Dim r As Long
    r = lstExportPreview.ListIndex
    If r < 0 Then Exit Sub
 
    'Toggle selection
    lstExportPreview.Selected(r) = Not lstExportPreview.Selected(r)
 
    'Update footer count
    Imp_RefreshCounts
End Sub
 
'===========================================================
' SELECT ALL / CLEAR ALL
'===========================================================
Private Sub btnSelectAllImport_Click()
    On Error GoTo EH
 
    Dim i As Long
    For i = 0 To lstExportPreview.ListCount - 1
        lstExportPreview.Selected(i) = True
    Next i
 
    Footer_RefreshCounts "Selected all files."
    Exit Sub
EH:
    MsgBox "Select All (Import) Error: " & Err.Description, vbCritical
End Sub
 
Private Sub btnClearImport_Click()
    On Error GoTo EH
 
    Dim i As Long
    For i = 0 To lstExportPreview.ListCount - 1
        lstExportPreview.Selected(i) = False
    Next i
 
    Footer_RefreshCounts "Cleared selection."
    Exit Sub
EH:
    MsgBox "Clear All (Import) Error: " & Err.Description, vbCritical
End Sub
 '------------------------------------
'SolidWorks Custom Property Manager
'Developed by Ramu Gopal
'The Tech Thinker
'https://thetechthinker.com
'-------------------------------------
'===========================================================
' REMOVE (only selected rows)
'===========================================================
Private Sub btnRemoveImport_Click()
    On Error GoTo EH
 
    If lstExportPreview.ListCount = 0 Then Exit Sub
 
    Dim removed As Long
    Dim i As Long
 
    'Remove from bottom to top (important!)
    For i = lstExportPreview.ListCount - 1 To 0 Step -1
        If lstExportPreview.Selected(i) Then
            lstExportPreview.RemoveItem i
            removed = removed + 1
        End If
    Next i
    
 If removed = 0 Then
        Footer_RefreshCounts "No files selected to remove."
        MsgBox "No files selected to remove.", vbInformation
    Else
        Footer_RefreshCounts "Removed " & removed & " file(s)."
    End If
    
    'Footer_RefreshCounts "Removed " & removed & " item(s)."
    Exit Sub
EH:
    MsgBox "Remove (Import) Error: " & Err.Description, vbCritical
End Sub
'===========================================================
' RESET IMPORT TAB
'===========================================================
Private Sub btnResetImport_Click()
    txtModelFolderImp.Text = ""
    txtCsvFolderImp.Text = ""
    lstExportPreview.Clear
 
    lblStatusMsg.Caption = "Import reset."
    lblProgress.Caption = "Progress: -"
    lblElapsed.Caption = "Elapsed: 00:00"
    lblLog.Caption = "Log: -"
    Progress_SetPercent 0
End Sub
 
'===========================================================
' Footer counts for Import list
'===========================================================
Private Sub Imp_RefreshCounts()
    Dim totalRows As Long, checkedRows As Long
    totalRows = lstExportPreview.ListCount
    checkedRows = Imp_CountSelected()
 
    lblProgress.Caption = "Progress: " & checkedRows & "/" & totalRows & " files selected"
End Sub
 
Private Function Imp_CountSelected() As Long
    Dim i As Long, cnt As Long
    cnt = 0
    For i = 0 To lstExportPreview.ListCount - 1
        If lstExportPreview.Selected(i) Then cnt = cnt + 1
    Next i
    Imp_CountSelected = cnt
End Function
 
'===========================================================
' Add one row (4 columns)
'===========================================================
Private Sub Imp_AddRow(ByVal fileName As String, ByVal fileType As String, ByVal fullPath As String, ByVal statusText As String)
    Dim r As Long
    lstExportPreview.AddItem fileName
    r = lstExportPreview.ListCount - 1
    lstExportPreview.List(r, 1) = fileType
    lstExportPreview.List(r, 2) = fullPath
    lstExportPreview.List(r, 3) = statusText
End Sub

Private Sub lblTabExport_Click()
    MultiPage1.Value = 0
    HighlightTabs 0
End Sub

