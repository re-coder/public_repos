Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'==============================================
' Main routine: GenerateQuotation
'==============================================
Sub GenerateQuotation()
    Dim masterFileName As String, genFileName As String
    Dim inputsPath As String
    Dim inputsWB As Workbook, masterWB As Workbook
    Dim genSheet As Worksheet, secSheet As Worksheet
    Dim placeholders As Object
    Dim sectionsGroup1 As Object, sectionsGroup2 As Object
    Dim lastRowGeneral As Long, lastRow As Long, i As Long
    Dim key As String, rowData As Variant
    Dim rowIndex As Long, currentHeader As String, cellVal As String
    Dim dataRows As Collection
    Dim lastRowGroup1 As Long, lastRowGroup2 As Long
    Dim masterPath As String, newNameDoc As String, newNamePdf As String
    Dim currentQuot As Long, updatedWB As Workbook, genInputsSheet As Worksheet
    Dim k As Long, sKey As Variant, sectionData As Variant
    Dim masterWS As Worksheet
    Dim photoName As String, photoPath As String
    Dim finalPdfPath As String
    Dim fileNameDoc As String, fileNamePdf As String
    
    ' 0) Fileâ€open checks
    masterFileName = "master_quotation_format.xlsx"
    genFileName = "Generated Quotation.xlsx"
    If IsWorkbookOpen(masterFileName) Then
        MsgBox masterFileName & " is open. Close it and retry.", vbExclamation
        Exit Sub
    End If
    If IsWorkbookOpen(genFileName) Then
        MsgBox genFileName & " is open. Close it and retry.", vbExclamation
        Exit Sub
    End If
    
    ' 1) Build placeholder dictionary
    Set placeholders = CreateObject("Scripting.Dictionary")
    
    ' 2) Open inputs workbook
    inputsPath = ThisWorkbook.Path & "\quotation_inputs.xlsx"
    Set inputsWB = Workbooks.Open(inputsPath)
    
    ' 2a) General Inputs (keys in B, values in C from rowÂ 3 down)
    Set genSheet = inputsWB.Sheets("General Inputs")
    lastRowGeneral = genSheet.Cells(genSheet.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRowGeneral
        key = Trim(genSheet.Cells(i, "B").Value)
        key = Replace(key, ":", "")
        If key <> "" Then
            If Left(key, 1) = """" And Right(key, 1) = """" Then
                key = Mid(key, 2, Len(key) - 2)
                placeholders(key) = Array(genSheet.Cells(i, "C").Value, True)
            Else
                placeholders(key) = Array(genSheet.Cells(i, "C").Value, False)
            End If
        End If
    Next i
    
    ' 2b) Section Inputs â€” two groups
    Set sectionsGroup1 = CreateObject("Scripting.Dictionary")
    Set sectionsGroup2 = CreateObject("Scripting.Dictionary")
    Set secSheet = inputsWB.Sheets("Section Inputs")
    
    ' GroupÂ 1: header in B2, titles in C:G rowÂ 3, data rowÂ 4+
    lastRowGroup1 = secSheet.Cells(secSheet.Rows.Count, "B").End(xlUp).Row
    rowIndex = 1
    Do While rowIndex <= lastRowGroup1
        cellVal = Trim(secSheet.Cells(rowIndex, "B").Value)
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            rowIndex = rowIndex + 2
            Set dataRows = New Collection
            Do While rowIndex <= lastRowGroup1 And Trim(secSheet.Cells(rowIndex, "B").Value) = ""
                rowData = Array( _
                  secSheet.Cells(rowIndex, "C").Value, _
                  secSheet.Cells(rowIndex, "D").Value, _
                  secSheet.Cells(rowIndex, "E").Value, _
                  secSheet.Cells(rowIndex, "F").Value, _
                  secSheet.Cells(rowIndex, "G").Value)
                If Not IsAllEmpty(rowData) Then dataRows.Add rowData
                rowIndex = rowIndex + 1
            Loop
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
    
    ' GroupÂ 2: header in J2, titles in K:O rowÂ 3, data rowÂ 4+
    lastRowGroup2 = secSheet.Cells(secSheet.Rows.Count, "J").End(xlUp).Row
    rowIndex = 1
    Do While rowIndex <= lastRowGroup2
        cellVal = Trim(secSheet.Cells(rowIndex, "J").Value)
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            rowIndex = rowIndex + 2
            Set dataRows = New Collection
            Do While rowIndex <= lastRowGroup2 And Trim(secSheet.Cells(rowIndex, "J").Value) = ""
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
    
    inputsWB.Close False
    
    ' 3) Open master template
    masterPath = ThisWorkbook.Path & "\dev(do not edit)\master_quotation_format.xlsx"
    Set masterWB = Workbooks.Open(masterPath)
    Set masterWS = masterWB.Sheets(1)
    
    ' 4a) Header placeholders
    UpdateHeader masterWS, placeholders
    
    ' 4b) Photo
    If placeholders.Exists("<<Photo>>") Then
        photoName = placeholders("<<Photo>>")(0)
        photoPath = ThisWorkbook.Path & "\photos\" & photoName
        If Dir(photoPath) <> "" Then InsertPhoto masterWS, "<<Photo>>", photoPath
    End If
    
    ' 4c) Sections GroupÂ 1
    For Each sKey In sectionsGroup1.Keys
        sectionData = sectionsGroup1(sKey)
        If IsArray(sectionData) Then UpdateSection masterWS, CStr(sKey), sectionData
    Next sKey
    
    ' 4d) Sections GroupÂ 2
    For Each sKey In sectionsGroup2.Keys
        sectionData = sectionsGroup2(sKey)
        If IsArray(sectionData) Then UpdateSection masterWS, CStr(sKey), sectionData
    Next sKey
    
    ' 4e) Optional sub total
    Dim c As Range
    For Each c In masterWS.UsedRange
        If VarType(c.Value) = vbString Then
            If InStr(c.Value, "Sub Total Cost (USD):") > 0 Then
                c.Value = "Sub Total Cost (USD): $24,390"
            End If
        End If
    Next c
    
    ' 5) Rename outputs by Quotation Number
    If placeholders.Exists("Quotation Number") Then
        currentQuot = CLng(placeholders("Quotation Number")(0))
    Else
        currentQuot = 1
    End If
    newNameDoc = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".docx"
    newNamePdf = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".pdf"
    masterWB.SaveAs newNameDoc
    masterWB.ExportAsFixedFormat Type:=xlTypePDF, Filename:=newNamePdf
    
    ' 6) Increment Quotation Number in inputs
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
    
    ' 7) Open the PDF
    finalPdfPath = newNamePdf
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        Dim script As String
        script = "tell application ""Finder"" to open POSIX file """ & finalPdfPath & """"
        MacScript script
    Else
        ' Option A: use ShellExecute (already declared at top of module)
        ShellExecute 0, "open", finalPdfPath, vbNullString, vbNullString, 1
        
        ' — OR —
        ' Option B: call FollowHyperlink on the workbook
        ' ThisWorkbook.FollowHyperlink Address:=finalPdfPath
    End If

    
    masterWB.Close False
    
    fileNameDoc = Mid(newNameDoc, InStrRev(newNameDoc, "\") + 1)
    fileNamePdf = Mid(newNamePdf, InStrRev(newNamePdf, "\") + 1)
    MsgBox "Quotation saved as:" & vbCrLf & _
           "Word: " & fileNameDoc & vbCrLf & "PDF: " & fileNamePdf, vbInformation
    
CleanExit:
    On Error Resume Next
    If Not inputsWB Is Nothing Then inputsWB.Close False
    Set placeholders = Nothing
    Set sectionsGroup1 = Nothing
    Set sectionsGroup2 = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'==============================================
' Helper: Check if all elements in an array are empty.
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
' AppendToArray: Appends a new element to an existing array.
'==============================================
Public Function AppendToArray(oldArray As Variant, newValue As Variant) As Variant
    Dim newArray() As Variant, n As Long, i As Long
    On Error Resume Next: n = UBound(oldArray): On Error GoTo 0
    If n < 0 Then
        ReDim newArray(0 To 0)
        newArray(0) = newValue
    Else
        ReDim newArray(0 To n + 1)
        For i = 0 To n: newArray(i) = oldArray(i): Next i
        newArray(n + 1) = newValue
    End If
    AppendToArray = newArray
End Function

'==============================================
' UpdateHeader: Replaces placeholders in the master sheet.
'==============================================
Public Sub UpdateHeader(ws As Worksheet, placeholders As Object)
    Dim cell        As Range
    Dim key         As Variant
    Dim tmp         As Variant
    Dim newVal      As String
    Dim directReplace As Boolean

    For Each cell In ws.UsedRange
        If VarType(cell.Value) = vbString Then
            For Each key In placeholders.Keys
                tmp = placeholders(key)
                If IsArray(tmp) Then
                    ' array(0)=value, array(1)=True/False
                    newVal = CStr(tmp(0))
                    directReplace = CBool(tmp(1))
                Else
                    ' fallback: somebody stored a bare value
                    newVal = CStr(tmp)
                    directReplace = False
                End If

                If directReplace Then
                    If InStr(cell.Value, key) > 0 Then
                        cell.Value = newVal
                    End If
                Else
                    If Trim(cell.Value) Like key & ":*" Then
                        cell.Value = key & ": " & newVal
                    End If
                End If
            Next key
        End If
    Next cell
End Sub


'==============================================
' UpdateSection: Inserts section data below a header.
'==============================================
Public Sub UpdateSection(ws As Worksheet, sectionHeader As String, dataList As Variant)
    Dim pos As Variant, startRow As Long
    Dim rawCount As Long, numRows As Long, numCols As Long
    Dim i As Long, j As Long, colIndex As Long, outRow As Long
    Dim outputArr() As Variant, rng As Range, r As Long
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    pos = FindCell(ws, sectionHeader)
    startRow = pos(0)
    If startRow = 0 Then GoTo CleanupSection
    
    rawCount = UBound(dataList) - LBound(dataList) + 1
    numRows = IIf(rawCount Mod 2 = 0, rawCount, rawCount + 1)
    
    numCols = 0
    For j = LBound(dataList(0)) To UBound(dataList(0))
        If j <> 5 Then numCols = numCols + 1
    Next j
    
    ws.Rows(startRow + 1).Resize(numRows).Insert Shift:=xlDown
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
    
    ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + numRows, numCols)).Value = outputArr
    
    Set rng = ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + numRows, numCols))
    With rng.Font: .Bold = False: .Underline = xlUnderlineStyleNone: .Italic = False: End With
    
    For r = 1 To numRows
        With ws.Range(ws.Cells(startRow + r, 1), ws.Cells(startRow + r, numCols)).Interior
            .Color = IIf(r Mod 2 = 1, RGB(242, 242, 242), RGB(255, 255, 255))
        End With
    Next r

CleanupSection:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'==============================================
' FindCell: Returns {row, col} of first match or {0,0}.
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
' InsertPhoto: Inserts and sizes an image into a merged cell.
'==============================================
Public Sub InsertPhoto(ws As Worksheet, placeholderKey As String, photoPath As String)
    Dim foundRange As Range, targetRange As Range, picShape As Shape
    Dim fullW As Double, fullH As Double
    Set foundRange = ws.Cells.Find(What:=placeholderKey, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If foundRange Is Nothing Then Exit Sub
    Set targetRange = foundRange.MergeArea
    targetRange.Value = ""
    fullW = targetRange.Width: fullH = targetRange.Height
    On Error Resume Next
    Set picShape = ws.Shapes.AddPicture(Filename:=photoPath, _
        LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=targetRange.Left, Top:=targetRange.Top, _
        Width:=fullW, Height:=fullH)
    On Error GoTo 0
    If picShape Is Nothing Then Exit Sub
    With picShape: .LockAspectRatio = msoFalse: .Placement = xlMoveAndSize: End With
End Sub

'==============================================
' IsWorkbookOpen: True if a workbook by that name is in the collection.
'==============================================
Public Function IsWorkbookOpen(wbName As String) As Boolean
    On Error Resume Next
    IsWorkbookOpen = Not (Workbooks(wbName) Is Nothing)
End Function

