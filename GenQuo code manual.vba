Option Explicit    ' Require explicit declaration of all variables

#If VBA7 Then
    ' Declare ShellExecute for 64-bit Office
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
            ByVal hwnd As LongPtr, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long _
        ) As LongPtr
#Else
    ' Declare ShellExecute for 32-bit Office
    Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" ( _
            ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long _
        ) As Long
#End If

'==============================================
' Main routine: GenerateQuotation
'==============================================
Sub GenerateQuotation()
    ' File names for template and output
    Dim masterFileName As String, genFileName As String
    ' Path to inputs workbook
    Dim inputsPath As String
    ' Workbook objects for inputs and master
    Dim inputsWB As Workbook, masterWB As Workbook
    ' Worksheet objects for general inputs and section inputs
    Dim genSheet As Worksheet, secSheet As Worksheet
    ' Dictionary to hold placeholder key/value pairs
    Dim placeholders As Object
    ' Dictionaries to hold section data arrays
    Dim sectionsGroup1 As Object, sectionsGroup2 As Object
    ' Loop counters and row indices
    Dim lastRowGeneral As Long, lastRowGroup1 As Long, lastRowGroup2 As Long
    Dim i As Long, rowIndex As Long, k As Long
    ' Temporary variables for keys and headers
    Dim key As String, currentHeader As String, cellVal As String
    ' Temporary storage for one row of section data
    Dim rowData As Variant
    ' Collection to gather multiple rows per section
    Dim dataRows As Collection
    ' Temporary array to hold a section's data rows
    Dim sectionData() As Variant
    ' Worksheet for the master template
    Dim masterWS As Worksheet
    ' Paths and names for final documents
    Dim masterPath As String, newNameDoc As String, newNamePdf As String
    Dim finalPdfPath As String
    ' Quotation number counter
    Dim currentQuot As Long
    ' Workbook to update the inputs file
    Dim updatedWB As Workbook, genInputsSheet As Worksheet
    ' Variables for photo insertion
    Dim photoName As String, photoPath As String
    ' Variables for displaying final file names
    Dim fileNameDoc As String, fileNamePdf As String
    
    ' 0) Ensure template and previous output are not open
    masterFileName = "master_quotation_format.xlsx"
    genFileName    = "Generated Quotation.xlsx"
    If IsWorkbookOpen(masterFileName) Then
        MsgBox masterFileName & " is open. Close it and retry.", vbExclamation
        Exit Sub
    End If
    If IsWorkbookOpen(genFileName) Then
        MsgBox genFileName & " is open. Close it and retry.", vbExclamation
        Exit Sub
    End If
    
    ' 1) Create dictionary for placeholders
    Set placeholders = CreateObject("Scripting.Dictionary")
    
    ' 2) Open the inputs workbook
    inputsPath = ThisWorkbook.Path & "\quotation_inputs.xlsx"
    Set inputsWB = Workbooks.Open(inputsPath)
    
    ' 2a) Read general inputs from row 3 down (keys in B, values in C)
    Set genSheet = inputsWB.Sheets("General Inputs")
    lastRowGeneral = genSheet.Cells(genSheet.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRowGeneral
        key = Trim(genSheet.Cells(i, "B").Value)          ' Read key text
        key = Replace(key, ":", "")                       ' Remove any colon
        If key <> "" Then
            If Left(key, 1) = """" And Right(key, 1) = """" Then
                ' If wrapped in quotes, strip them and flag direct replace
                key = Mid(key, 2, Len(key) - 2)
                placeholders(key) = Array(genSheet.Cells(i, "C").Value, True)
            Else
                ' Otherwise store value and flag prefix replace
                placeholders(key) = Array(genSheet.Cells(i, "C").Value, False)
            End If
        End If
    Next i
    
    ' 2b) Read section inputs (two column groups) into dictionaries
    Set sectionsGroup1 = CreateObject("Scripting.Dictionary")
    Set sectionsGroup2 = CreateObject("Scripting.Dictionary")
    Set secSheet      = inputsWB.Sheets("Section Inputs")
    
    ' Group 1: header in column B (row 2), titles in C:G row 3, data from row 4+
    lastRowGroup1 = secSheet.Cells(secSheet.Rows.Count, "B").End(xlUp).Row
    rowIndex = 1
    Do While rowIndex <= lastRowGroup1
        cellVal = Trim(secSheet.Cells(rowIndex, "B").Value)
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            rowIndex = rowIndex + 2                     ' Skip header + title rows
            Set dataRows = New Collection
            ' Collect data rows until next header
            Do While rowIndex <= lastRowGroup1 _
                  And Trim(secSheet.Cells(rowIndex, "B").Value) = ""
                rowData = Array( _
                    secSheet.Cells(rowIndex, "C").Value, _
                    secSheet.Cells(rowIndex, "D").Value, _
                    secSheet.Cells(rowIndex, "E").Value, _
                    secSheet.Cells(rowIndex, "F").Value, _
                    secSheet.Cells(rowIndex, "G").Value)
                If Not IsAllEmpty(rowData) Then dataRows.Add rowData
                rowIndex = rowIndex + 1
            Loop
            ' Store into dictionary if any rows found
            If dataRows.Count > 0 Then
                ReDim sectionData(0 To dataRows.Count - 1)
                For k = 1 To dataRows.Count
                    sectionData(k - 1) = dataRows(k)
                Next k
                sectionsGroup1(currentHeader) = sectionData
            End If
        Else
            rowIndex = rowIndex + 1
        End If
    Loop
    
    ' Group 2: header in column J (row 2), titles in K:O row 3, data from row 4+
    lastRowGroup2 = secSheet.Cells(secSheet.Rows.Count, "J").End(xlUp).Row
    rowIndex = 1
    Do While rowIndex <= lastRowGroup2
        cellVal = Trim(secSheet.Cells(rowIndex, "J").Value)
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            rowIndex = rowIndex + 2
            Set dataRows = New Collection
            Do While rowIndex <= lastRowGroup2 _
                  And Trim(secSheet.Cells(rowIndex, "J").Value) = ""
                rowData = Array( _
                    secSheet.Cells(rowIndex, "K").Value, _
                    secSheet.Cells(rowIndex, "L").Value, _
                    secSheet.Cells(rowIndex, "M").Value, _
                    secSheet.Cells(rowIndex, "N").Value, _
                    secSheet.Cells(rowIndex, "O").Value)
                If Not IsAllEmpty(rowData) Then dataRows.Add rowData
                rowIndex = rowIndex + 1
            Loop
            If dataRows.Count > 0 Then
                ReDim sectionData(0 To dataRows.Count - 1)
                For k = 1 To dataRows.Count
                    sectionData(k - 1) = dataRows(k)
                Next k
                sectionsGroup2(currentHeader) = sectionData
            End If
        Else
            rowIndex = rowIndex + 1
        End If
    Loop
    
    ' Close inputs workbook without saving
    inputsWB.Close False
    
    ' 3) Open the master template workbook and its first sheet
    masterPath = ThisWorkbook.Path & "\dev(do not edit)\master_quotation_format.xlsx"
    Set masterWB = Workbooks.Open(masterPath)
    Set masterWS = masterWB.Sheets(1)
    
    ' 4a) Replace header placeholders in the template
    UpdateHeader masterWS, placeholders
    
    ' 4b) Insert photo if provided
    If placeholders.Exists("<<Photo>>") Then
        photoName = placeholders("<<Photo>>")(0)            ' File name of photo
        photoPath = ThisWorkbook.Path & "\photos\" & photoName
        If Dir(photoPath) <> "" Then
            InsertPhoto masterWS, "<<Photo>>", photoPath
        End If
    End If
    
    ' 4c) Populate sections from Group 1
    For Each sKey In sectionsGroup1.Keys
        sectionData = sectionsGroup1(sKey)
        If IsArray(sectionData) Then
            UpdateSection masterWS, CStr(sKey), sectionData
        End If
    Next sKey
    
    ' 4d) Populate sections from Group 2
    For Each sKey In sectionsGroup2.Keys
        sectionData = sectionsGroup2(sKey)
        If IsArray(sectionData) Then
            UpdateSection masterWS, CStr(sKey), sectionData
        End If
    Next sKey
    
    ' 4e) Optional: overwrite any "Sub Total Cost (USD):" text
    Dim c As Range
    For Each c In masterWS.UsedRange
        If VarType(c.Value) = vbString Then
            If InStr(c.Value, "Sub Total Cost (USD):") > 0 Then
                c.Value = "Sub Total Cost (USD): $24,390"
            End If
        End If
    Next c
    
    ' 5) Determine and apply the quotation number to output filenames
    If placeholders.Exists("Quotation Number") Then
        currentQuot = CLng(placeholders("Quotation Number")(0))
    Else
        currentQuot = 1
    End If
    newNameDoc = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".docx"
    newNamePdf = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".pdf"
    masterWB.SaveAs newNameDoc
    masterWB.ExportAsFixedFormat Type:=xlTypePDF, Filename:=newNamePdf
    
    ' 6) Increment the stored quotation number back in inputs file
    Set updatedWB = Workbooks.Open(ThisWorkbook.Path & "\quotation_inputs.xlsx")
    Set genInputsSheet = updatedWB.Sheets("General Inputs")
    For k = 3 To genInputsSheet.Cells(genInputsSheet.Rows.Count, "B").End(xlUp).Row
        If Trim(genInputsSheet.Cells(k, "B").Value) = "Quotation Number" Then
            genInputsSheet.Cells(k, "C").Value = currentQuot + 1
            Exit For
        End If
    Next k
    updatedWB.Save
    updatedWB.Close False
    
    ' 7) Open the resulting PDF automatically
    finalPdfPath = newNamePdf
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        Dim script As String
        script = "tell application ""Finder"" to open POSIX file """ & finalPdfPath & """"
        MacScript script
    Else
        ShellExecute 0, "open", finalPdfPath, vbNullString, vbNullString, 1
    End If
    
    ' Close the master workbook without saving further changes
    masterWB.Close False
    
    ' Show a confirmation message listing the created files
    fileNameDoc = Mid(newNameDoc, InStrRev(newNameDoc, "\") + 1)
    fileNamePdf = Mid(newNamePdf, InStrRev(newNamePdf, "\") + 1)
    MsgBox "Quotation saved as:" & vbCrLf & _
           "Word: " & fileNameDoc & vbCrLf & "PDF: " & fileNamePdf, _
           vbInformation, "Quotation Generation"
    
CleanExit:
    ' Ensure inputsWB is closed if an error occurred early
    On Error Resume Next
    If Not inputsWB Is Nothing Then inputsWB.Close False
    Exit Sub

ErrorHandler:
    ' Display any unexpected error
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'==============================================
' Helper: Check if all elements in a 1D array are empty
'==============================================
Private Function IsAllEmpty(arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Trim(CStr(arr(i))) <> "" Then
            IsAllEmpty = False
            Exit Function
        End If
    Next i
    IsAllEmpty = True
End Function

'==============================================
' UpdateHeader: Replace placeholders in the template sheet
'==============================================
Public Sub UpdateHeader(ws As Worksheet, placeholders As Object)
    Dim cell As Range, key As Variant
    Dim tmp As Variant, newVal As String, directReplace As Boolean
    For Each cell In ws.UsedRange
        If VarType(cell.Value) = vbString Then
            For Each key In placeholders.Keys
                tmp = placeholders(key)
                If IsArray(tmp) Then
                    newVal = CStr(tmp(0))        ' Extract stored value
                    directReplace = CBool(tmp(1)) ' Extract replace flag
                Else
                    newVal = CStr(tmp)
                    directReplace = False
                End If
                If directReplace Then
                    ' Direct replace anywhere the key appears
                    If InStr(cell.Value, key) > 0 Then cell.Value = newVal
                Else
                    ' Only replace if cell exactly matches "Key:*"
                    If Trim(cell.Value) Like key & ":*" Then
                        cell.Value = key & ": " & newVal
                    End If
                End If
            Next key
        End If
    Next cell
End Sub

'==============================================
' UpdateSection: Insert a block of rows under a header
'==============================================
Public Sub UpdateSection(ws As Worksheet, sectionHeader As String, dataList As Variant)
    Dim pos As Variant, startRow As Long
    Dim rawCount As Long, numRows As Long, numCols As Long
    Dim i As Long, j As Long, colIndex As Long, outRow As Long
    Dim outputArr() As Variant, rng As Range, r As Long
    
    ' Speed up by disabling screen updates and events
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Find the header cell
    pos = FindCell(ws, sectionHeader)
    startRow = pos(0)
    If startRow = 0 Then GoTo CleanupSection
    
    ' Calculate how many rows to insert (even number)
    rawCount = UBound(dataList) - LBound(dataList) + 1
    numRows = IIf(rawCount Mod 2 = 0, rawCount, rawCount + 1)
    
    ' Count how many columns (skip the 6th element = remarks)
    numCols = 0
    For j = LBound(dataList(0)) To UBound(dataList(0))
        If j <> 5 Then numCols = numCols + 1
    Next j
    
    ' Insert blank rows below the header
    ws.Rows(startRow + 1).Resize(numRows).Insert Shift:=xlDown
    
    ' Prepare output array and fill with dataList
    ReDim outputArr(1 To numRows, 1 To numCols)
    For i = LBound(dataList) To UBound(dataList)
        outRow = i - LBound(dataList) + 1
        colIndex = 1
        For j = LBound(dataList(i)) To UBound(dataList(i))
            If j <> 5 Then
                outputArr(outRow, colIndex) = dataList(i)(j)
                colIndex = colIndex + 1
            End If
        Next j
    Next i
    
    ' Write all at once
    ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + numRows, numCols)).Value = outputArr
    
    ' Ensure plain font
    Set rng = ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + numRows, numCols))
    With rng.Font
        .Bold = False
        .Underline = xlUnderlineStyleNone
        .Italic = False
    End With
    
    ' Apply alternating shading
    For r = 1 To numRows
        With ws.Range(ws.Cells(startRow + r, 1), ws.Cells(startRow + r, numCols)).Interior
            .Color = IIf(r Mod 2 = 1, RGB(242, 242, 242), RGB(255, 255, 255))
        End With
    Next r

CleanupSection:
    ' Restore application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'==============================================
' FindCell: Locate first cell containing searchText
'==============================================
Public Function FindCell(ws As Worksheet, searchText As String) As Variant
    Dim c As Range
    For Each c In ws.UsedRange
        If VarType(c.Value) = vbString Then
            If InStr(c.Value, searchText) > 0 Then
                FindCell = Array(c.Row, c.Column)
                Exit Function
            End If
        End If
    Next c
    FindCell = Array(0, 0)
End Function

'==============================================
' InsertPhoto: Place and size a picture in a merged cell
'==============================================
Public Sub InsertPhoto(ws As Worksheet, placeholderKey As String, photoPath As String)
    Dim foundRange As Range, targetRange As Range, picShape As Shape
    Dim fullW As Double, fullH As Double
    
    ' Find the placeholder cell
    Set foundRange = ws.Cells.Find(What:=placeholderKey, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If foundRange Is Nothing Then Exit Sub
    
    ' Get the merged area around it, clear contents
    Set targetRange = foundRange.MergeArea
    targetRange.Value = ""
    fullW = targetRange.Width
    fullH = targetRange.Height
    
    ' Insert and size picture
    On Error Resume Next
    Set picShape = ws.Shapes.AddPicture( _
        Filename:=photoPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=targetRange.Left, Top:=targetRange.Top, Width:=fullW, Height:=fullH)
    On Error GoTo 0
    If picShape Is Nothing Then Exit Sub
    
    With picShape
        .LockAspectRatio = msoFalse
        .Placement = xlMoveAndSize
    End With
End Sub

'==============================================
' IsWorkbookOpen: True if a workbook with that name is open
'==============================================
Public Function IsWorkbookOpen(wbName As String) As Boolean
    On Error Resume Next
    IsWorkbookOpen = Not (Workbooks(wbName) Is Nothing)
End Function
