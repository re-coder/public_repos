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
    Dim currencyCode As String
    Dim resp As VbMsgBoxResult

    Debug.Print "=== GenerateQuotation START ==="
    
    ' 0) Check that template and prior output are closed
    masterFileName = "master_quotation_format.xlsx"
    genFileName    = "Generated Quotation.xlsx"
    Debug.Print "Checking open books..."
    If IsWorkbookOpen(masterFileName) Then
        Debug.Print "ERROR: Template is open."
        MsgBox masterFileName & " is open. Close it and retry.", vbExclamation
        Exit Sub
    End If
    If IsWorkbookOpen(genFileName) Then
        Debug.Print "ERROR: Previous output is open."
        MsgBox genFileName & " is open. Close it and retry.", vbExclamation
        Exit Sub
    End If
    Debug.Print "OK: template & prior outputs closed."
    
    ' 1) Build placeholder dictionary
    Debug.Print "Building placeholder dictionary..."
    Set placeholders = CreateObject("Scripting.Dictionary")
    
    ' 2) Open inputs workbook
    inputsPath = ThisWorkbook.Path & "\quotation_inputs.xlsx"
    Debug.Print "Opening inputs at " & inputsPath
    Set inputsWB = Workbooks.Open(inputsPath)
    Debug.Print "Opened: " & inputsWB.Name
    
    ' 2a) Read General Inputs
    Debug.Print "Reading General Inputs..."
    Set genSheet = inputsWB.Sheets("General Inputs")
    lastRowGeneral = genSheet.Cells(genSheet.Rows.Count, "B").End(xlUp).Row
    Debug.Print " General Inputs last row: " & lastRowGeneral
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

    ' — confirm currency —
    If Not placeholders.Exists("Currency") Then
        MsgBox "Currency not found in General Inputs. Cannot continue.", vbCritical
        GoTo CleanExit
    End If
    currencyCode = CStr(placeholders("Currency")(0))
    resp = MsgBox("Currency is set to " & currencyCode & ".  OK to proceed?", vbQuestion + vbYesNo, "Confirm Currency")
    If resp <> vbYes Then
        Debug.Print "Generation cancelled by user at currency prompt."
        GoTo CleanExit
    End If
    Debug.Print "Currency confirmed: " & currencyCode

    ' 2b) Read Section Inputs
    Debug.Print "Reading Section Inputs..."
    Set sectionsGroup1 = CreateObject("Scripting.Dictionary")
    Set sectionsGroup2 = CreateObject("Scripting.Dictionary")
    Set secSheet = inputsWB.Sheets("Section Inputs")
    
    ' --- Group 1 ---
    Dim hdrLast1 As Long, dataLast1 As Long
    hdrLast1  = secSheet.Cells(secSheet.Rows.Count, "B").End(xlUp).Row
    dataLast1 = secSheet.Cells(secSheet.Rows.Count, "C").End(xlUp).Row
    lastRowGroup1 = Application.Max(hdrLast1, dataLast1)
    Debug.Print " Group1 last row: " & lastRowGroup1
    rowIndex = 1
    Do While rowIndex <= lastRowGroup1
        cellVal = Trim(secSheet.Cells(rowIndex, "B").Value)
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            Debug.Print "  G1 header @" & rowIndex & ": " & currentHeader
            If Len(currentHeader) >= 2 And IsNumeric(Mid(currentHeader, 2, 1)) Then
                skipRows = 1
            Else
                skipRows = 2
            End If
            Debug.Print "   skipRows=" & skipRows
            rowIndex = rowIndex + skipRows
            
            Set dataRows = New Collection
            Debug.Print "   collecting data from row " & rowIndex
            Do While rowIndex <= lastRowGroup1 And Trim(secSheet.Cells(rowIndex, "B").Value) = ""
                rowData = Array( _
                  secSheet.Cells(rowIndex, "C").Value, _
                  secSheet.Cells(rowIndex, "D").Value, _
                  secSheet.Cells(rowIndex, "E").Value, _
                  secSheet.Cells(rowIndex, "F").Value, _
                  secSheet.Cells(rowIndex, "G").Value)
                If Not IsAllEmpty(rowData) Then
                    dataRows.Add rowData
                    Debug.Print "    Added row@" & rowIndex & ": [" & Join(rowData, ", ") & "]"
                Else
                    Debug.Print "    Skipped blank row@" & rowIndex
                End If
                rowIndex = rowIndex + 1
            Loop
            
            If dataRows.Count > 0 Then
                ReDim sectionData(0 To dataRows.Count - 1)
                For k = 1 To dataRows.Count
                    sectionData(k - 1) = dataRows(k)
                Next k
                sectionsGroup1(currentHeader) = sectionData
                Debug.Print "   Stored G1 '" & currentHeader & "' (" & dataRows.Count & " items)"
            Else
                Debug.Print "   No data under '" & currentHeader & "'"
            End If
        Else
            rowIndex = rowIndex + 1
        End If
    Loop
    
    ' --- Group 2 ---
    Dim hdrLast2 As Long, dataLast2 As Long
    hdrLast2  = secSheet.Cells(secSheet.Rows.Count, "K").End(xlUp).Row
    dataLast2 = secSheet.Cells(secSheet.Rows.Count, "L").End(xlUp).Row
    lastRowGroup2 = Application.Max(hdrLast2, dataLast2)
    Debug.Print " Group2 last row: " & lastRowGroup2
    rowIndex = 1
    Do While rowIndex <= lastRowGroup2
        cellVal = Trim(secSheet.Cells(rowIndex, "K").Value)
        If cellVal <> "" And LCase(cellVal) <> "section item" Then
            currentHeader = cellVal
            Debug.Print "  G2 header @" & rowIndex & ": " & currentHeader
            If Len(currentHeader) >= 2 And IsNumeric(Mid(currentHeader, 2, 1)) Then
                skipRows = 1
            Else
                skipRows = 2
            End If
            Debug.Print "   skipRows=" & skipRows
            rowIndex = rowIndex + skipRows
            
            Set dataRows = New Collection
            Debug.Print "   collecting data from row " & rowIndex
            Do While rowIndex <= lastRowGroup2 And Trim(secSheet.Cells(rowIndex, "K").Value) = ""
                rowData = Array( _
                  secSheet.Cells(rowIndex, "L").Value, _
                  secSheet.Cells(rowIndex, "M").Value, _
                  secSheet.Cells(rowIndex, "N").Value, _
                  secSheet.Cells(rowIndex, "O").Value, _
                  secSheet.Cells(rowIndex, "P").Value)
                If Not IsAllEmpty(rowData) Then
                    dataRows.Add rowData
                    Debug.Print "    Added row@" & rowIndex & ": [" & Join(rowData, ", ") & "]"
                Else
                    Debug.Print "    Skipped blank row@" & rowIndex
                End If
                rowIndex = rowIndex + 1
            Loop
            
            If dataRows.Count > 0 Then
                ReDim sectionData(0 To dataRows.Count - 1)
                For k = 1 To dataRows.Count
                    sectionData(k - 1) = dataRows(k)
                Next k
                sectionsGroup2(currentHeader) = sectionData
                Debug.Print "   Stored G2 '" & currentHeader & "' (" & dataRows.Count & " items)"
            Else
                Debug.Print "   No data under '" & currentHeader & "'"
            End If
        Else
            rowIndex = rowIndex + 1
        End If
    Loop
    
    ' close inputs
    Debug.Print "Closing inputs workbook..."
    inputsWB.Close False
    
    ' 3) Open master template
    masterPath = ThisWorkbook.Path & "\dev(do not edit)\master_quotation_format.xlsx"
    Debug.Print "Opening master template: " & masterPath
    Set masterWB = Workbooks.Open(masterPath)
    Set masterWS = masterWB.Sheets(1)
    
    ' 4a) Replace header placeholders
    Debug.Print "Updating header placeholders..."
    UpdateHeader masterWS, placeholders
    
    ' 4b) Photo
    If placeholders.Exists("<<Photo>>") Then
        photoName = placeholders("<<Photo>>")(0)
        photoPath = ThisWorkbook.Path & "\photos\" & photoName
        Debug.Print "Inserting photo '" & photoName & "'..."
        If Dir(photoPath) <> "" Then
            InsertPhoto masterWS, "<<Photo>>", photoPath
            Debug.Print " Photo inserted."
        End If
    End If
    
    ' 4c) Write all sections
    Debug.Print "Writing Group1 sections..."
    For Each sKey In sectionsGroup1.Keys
        sectionData = sectionsGroup1(sKey)
        Debug.Print " Processing G1 key: " & sKey
        If IsArray(sectionData) Then
            Dim displayHeader As String
            Select Case True
                Case sKey = "A.":       displayHeader = "A. Construction Fixtures & Materials"
                Case sKey Like "A1.*":  displayHeader = "A1. Flooring"
                Case sKey Like "A2.*":  displayHeader = "A2. Hanging Banner/Structure"
                Case sKey Like "A3.*":  displayHeader = "A3. System and Basic Structure"
                Case sKey = "F.":       displayHeader = "F. Project Management (includes: I & D / VISA / Airfare / Accommodation / Transport / Design Fee / Miscellaneous)"
                Case sKey Like "F1.*":  displayHeader = "F1. Manpower"
                Case sKey Like "F2.*":  displayHeader = "F2. Accommodation"
                Case sKey Like "F3.*":  displayHeader = "F3. Air Tickets"
                Case sKey Like "F4.*":  displayHeader = "F4. Transportation"
                Case sKey Like "F5.*":  displayHeader = "F5. Miscellaneous, tools, hardware, accessories"
                Case sKey Like "F6.*":  displayHeader = "F6. Admin"
                Case sKey Like "F7.*":  displayHeader = "F7. Photography"
                Case sKey Like "F8.*":  displayHeader = "F8. Professional fee (PE endorsement)"
                Case sKey Like "F9.*":  displayHeader = "F9. Courier and storage charges"
                Case sKey Like "F10.*": displayHeader = "F10. Preshow maintenance, packing"
                Case sKey Like "F11.*": displayHeader = "F11. Others"
                Case Else:               displayHeader = sKey
            End Select
            Debug.Print "  Updating section '" & displayHeader & "'"
            UpdateSection masterWS, displayHeader, sectionData
        End If
    Next sKey
    
    Debug.Print "Writing Group2 sections..."
    For Each sKey In sectionsGroup2.Keys
        sectionData = sectionsGroup2(sKey)
        Debug.Print " Processing G2 key: " & sKey
        If IsArray(sectionData) Then
            Select Case True
                Case sKey = "X.":       displayHeader = "X. Payment to Organisers (paid by exhibitor)"
                Case sKey Like "X1.*":  displayHeader = "X1. Floral arrangements"
                Case sKey Like "X2.*":  displayHeader = "X2. Contractor Badges"
                Case sKey Like "X3.*":  displayHeader = "X3. Parking Passes"
                Case sKey Like "X4.*":  displayHeader = "X4. Stand Approval"
                Case sKey Like "X5.*":  displayHeader = "X5. Main Electrical connection"
                Case sKey Like "X6.*":  displayHeader = "X6. Build-up electrical connection"
                Case sKey Like "X7.*":  displayHeader = "X7. Internet connection"
                Case sKey Like "X8.*":  displayHeader = "X8. Rigging Services"
                Case sKey Like "X9.*":  displayHeader = "X9. Badges"
                Case sKey Like "X10.*": displayHeader = "X10. Late charges"
                Case sKey Like "X11.*": displayHeader = "X11. Others"
                Case Else:               displayHeader = sKey
            End Select
            Debug.Print "  Updating section '" & displayHeader & "'"
            UpdateSection masterWS, displayHeader, sectionData
        End If
    Next sKey
    
    ' 4e) Sub Total override
    Debug.Print "Overwriting any Sub Total if present..."
    Dim c As Range
    For Each c In masterWS.UsedRange
        If VarType(c.Value) = vbString And InStr(c.Value, "Sub Total Cost") > 0 Then
            c.Value = "Sub Total Cost (USD): $24,390"
            Debug.Print " Overwrote Sub Total at " & c.Address
        End If
    Next c
    
    ' —— Apply currency formatting ——
    Debug.Print "Applying currency format: " & currencyCode
    ApplyCurrencyFormat masterWS, currencyCode
    
    ' 5) Save and export
    Debug.Print "Determining Quotation#..."
    If placeholders.Exists("Quotation Number") Then
        currentQuot = CLng(placeholders("Quotation Number")(0))
    Else
        currentQuot = 1
    End If
    newNameDoc = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".docx"
    newNamePdf = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".pdf"
    Debug.Print " Saving DOCX: " & newNameDoc
    masterWB.SaveAs newNameDoc
    Debug.Print " Exporting PDF: " & newNamePdf
    masterWB.ExportAsFixedFormat Type:=xlTypePDF, Filename:=newNamePdf
    
    ' 6) Increment Quotation#
    Debug.Print "Incrementing Quotation Number in inputs..."
    Set updatedWB = Workbooks.Open(inputsPath)
    Set genInputsSheet = updatedWB.Sheets("General Inputs")
    For k = 3 To genInputsSheet.Cells(genInputsSheet.Rows.Count, "B").End(xlUp).Row
        If Trim(genInputsSheet.Cells(k, "B").Value) = "Quotation Number" Then
            genInputsSheet.Cells(k, "C").Value = currentQuot + 1
            Debug.Print "  Updated Quotation# at row " & k & " to " & currentQuot + 1
            Exit For
        End If
    Next k
    updatedWB.Save
    updatedWB.Close False
    
    ' 7) Open the PDF
    Debug.Print "Opening PDF..."
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

    Debug.Print "=== GenerateQuotation COMPLETE ==="
    
CleanExit:
    On Error Resume Next
    If Not inputsWB Is Nothing Then inputsWB.Close False
    Set placeholders      = Nothing
    Set sectionsGroup1    = Nothing
    Set sectionsGroup2    = Nothing
    Exit Sub

ErrorHandler:
    Debug.Print "Error " & Err.Number & ": " & Err.Description
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit
End Sub
