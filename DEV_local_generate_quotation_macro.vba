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
    Dim lastRowGeneral As Long, i As Long
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
    Dim skipRows As Long
    
    ' 0) Check that template and prior output are closed
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
    
    ' 1) Build placeholder dictionary
    Set placeholders = CreateObject("Scripting.Dictionary")
    
    ' 2) Open inputs workbook
    inputsPath = ThisWorkbook.Path & "\quotation_inputs.xlsx"
    Set inputsWB = Workbooks.Open(inputsPath)
    
    ' 2a) Read General Inputs (keys in B, values in C from row 3 down)
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
    
    ' 2b) Read Section Inputs — two groups
    Set sectionsGroup1 = CreateObject("Scripting.Dictionary")
    Set sectionsGroup2 = CreateObject("Scripting.Dictionary")
    Set secSheet = inputsWB.Sheets("Section Inputs")
    
    ' ---- Group 1: headers in Col B, titles in C:G on row 3, data from row 4+
    lastRowGroup1 = secSheet.Cells(secSheet.Rows.Count, "B").End(xlUp).Row
    rowIndex = 1
    Do While rowIndex <= lastRowGroup1
        cellVal = Trim(secSheet.Cells(rowIndex, "B").Value)
        ' if we’ve found a real section header (not the literal "Section Item")
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            ' decide how many rows to skip based on header format:
            '  - letter+dot  => skip 2 (header + titles)
            '  - letter+digit+dot => skip 1 (no titles)
            If Len(currentHeader) >= 2 And _
               IsNumeric(Mid(currentHeader, 2, 1)) Then
                skipRows = 1
            Else
                skipRows = 2
            End If
            rowIndex = rowIndex + skipRows
            
            ' collect all blank-keyed rows as data rows
            Set dataRows = New Collection
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
            
            ' store into dictionary
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
    
    ' ---- Group 2: headers in Col K, titles in L:P row 3, data from row 4+
    lastRowGroup2 = secSheet.Cells(secSheet.Rows.Count, "K").End(xlUp).Row
    rowIndex = 1
    Do While rowIndex <= lastRowGroup2
        cellVal = Trim(secSheet.Cells(rowIndex, "K").Value)
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            If Len(currentHeader) >= 2 And _
               IsNumeric(Mid(currentHeader, 2, 1)) Then
                skipRows = 1
            Else
                skipRows = 2
            End If
            rowIndex = rowIndex + skipRows
            
            Set dataRows = New Collection
            Do While rowIndex <= lastRowGroup2 _
                     And Trim(secSheet.Cells(rowIndex, "K").Value) = ""
                rowData = Array( _
                  secSheet.Cells(rowIndex, "L").Value, _
                  secSheet.Cells(rowIndex, "M").Value, _
                  secSheet.Cells(rowIndex, "N").Value, _
                  secSheet.Cells(rowIndex, "O").Value, _
                  secSheet.Cells(rowIndex, "P").Value)
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
    
    ' 3) Open the master template
    masterPath = ThisWorkbook.Path & "\dev(do not edit)\master_quotation_format.xlsx"
    Set masterWB = Workbooks.Open(masterPath)
    Set masterWS = masterWB.Sheets(1)
    
    ' 4a) Replace header placeholders
    UpdateHeader masterWS, placeholders
    
    ' 4b) Insert photo if requested
    If placeholders.Exists("<<Photo>>") Then
        photoName = placeholders("<<Photo>>")(0)
        photoPath = ThisWorkbook.Path & "\photos\" & photoName
        If Dir(photoPath) <> "" Then _
           InsertPhoto masterWS, "<<Photo>>", photoPath
    End If
    
    ' 4c) Write out Group 1 sections
    For Each sKey In sectionsGroup1.Keys
        sectionData = sectionsGroup1(sKey)
        If IsArray(sectionData) Then
            Dim displayHeader As String
            Select Case True
                Case sKey = "A."
                    displayHeader = "A. Construction Fixtures & Materials"
                Case sKey Like "A1.*"
                    displayHeader = "A1. Flooring"
                Case sKey Like "A2.*"
                    displayHeader = "A2. Hanging Banner/Structure"
                Case sKey Like "A3.*"
                    displayHeader = "A3. System and Basic Structure"
                Case sKey = "F."
                    displayHeader = "F. Project Management (includes: I & D / VISA / Airfare / Accommodation / Transport / Design Fee / Miscellaneous)"
                Case sKey Like "F1.*"
                    displayHeader = "F1. Manpower"
                Case sKey Like "F2.*"
                    displayHeader = "F2. Accommodation"
                Case sKey Like "F3.*"
                    displayHeader = "F3. Air Tickets"
                Case sKey Like "F4.*"
                    displayHeader = "F4. Transportation"
                Case sKey Like "F5.*"
                    displayHeader = "F5. Miscellaneous, tools, hardware, accessories"
                Case sKey Like "F6.*"
                    displayHeader = "F6. Admin"
                Case sKey Like "F7.*"
                    displayHeader = "F7. Photography"
                Case sKey Like "F8.*"
                    displayHeader = "F8. Professional fee (PE endorsement)"
                Case sKey Like "F9.*"
                    displayHeader = "F9. Courier and storage charges"
                Case sKey Like "F10.*"
                    displayHeader = "F10. Preshow maintenance, packing"
                Case sKey Like "F11.*"
                    displayHeader = "F11. Others"
                Case Else
                    displayHeader = CStr(sKey)
            End Select
            UpdateSection masterWS, displayHeader, sectionData
        End If
    Next sKey
    
    ' 4d) Write out Group 2 sections
    For Each sKey In sectionsGroup2.Keys
        sectionData = sectionsGroup2(sKey)
        If IsArray(sectionData) Then
            Select Case True
                Case sKey = "X."
                    displayHeader = "X. Payment to Organisers (paid by exhibitor)"
                Case sKey Like "X1.*"
                    displayHeader = "X1. Floral arrangements"
                Case sKey Like "X2.*"
                    displayHeader = "X2. Contractor Badges"
                Case sKey Like "X3.*"
                    displayHeader = "X3. Parking Passes"
                Case sKey Like "X4.*"
                    displayHeader = "X4. Stand Approval"
                Case sKey Like "X5.*"
                    displayHeader = "X5. Main Electrical connection"
                Case sKey Like "X6.*"
                    displayHeader = "X6. Build-up electrical connection"
                Case sKey Like "X7.*"
                    displayHeader = "X7. Internet connection"
                Case sKey Like "X8.*"
                    displayHeader = "X8. Rigging Services"
                Case sKey Like "X9.*"
                    displayHeader = "X9. Badges"
                Case sKey Like "X10.*"
                    displayHeader = "X10. Late charges"
                Case sKey Like "X11.*"
                    displayHeader = "X11. Others"
                Case Else
                    displayHeader = CStr(sKey)
            End Select
            UpdateSection masterWS, displayHeader, sectionData
        End If
    Next sKey
    
    ' 4e) Optional: overwrite any "Sub Total Cost" text
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
    
    ' 6) Increment that Quotation Number back in the inputs file
    Set updatedWB = Workbooks.Open(inputsPath)
    Set genInputsSheet = updatedWB.Sheets("General Inputs")
    For k = 3 To genInputsSheet.Cells(genInputsSheet.Rows.Count, "B").End(xlUp).Row
        If Trim(genInputsSheet.Cells(k, "B").Value) = "Quotation Number" Then
            genInputsSheet.Cells(k, "C").Value = currentQuot + 1
            Exit For
        End If
    Next k
    updatedWB.Save
    updatedWB.Close False
    
    ' 7) Finally open the PDF
    finalPdfPath = newNamePdf
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        Dim script As String
        script = "tell application ""Finder"" to open POSIX file """ & finalPdfPath & """"
        MacScript script
    Else
        ShellExecute 0, "open", finalPdfPath, vbNullString, vbNullString, 1
    End If
    
    masterWB.Close False
    
    fileNameDoc = Mid(newNameDoc, InStrRev(newNameDoc, "\") + 1)
    fileNamePdf = Mid(newNamePdf, InStrRev(newNamePdf, "\") + 1)
    MsgBox "Quotation saved as:" & vbCrLf & _
           "Word: " & fileNameDoc & vbCrLf & "PDF: " & fileNamePdf, vbInformation

CleanExit:
    On Error Resume Next
    If Not inputsWB Is Nothing Then inputsWB.Close False
    Set placeholders      = Nothing
    Set sectionsGroup1    = Nothing
    Set sectionsGroup2    = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'----------------------------------------------
' Helper: Returns True if every element of arr is blank
'----------------------------------------------
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
