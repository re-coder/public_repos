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

Option Explicit

Sub GenerateQuotation()
    Dim masterFileName As String, genFileName As String
    Dim inputsPath As String
    Dim inputsWB As Workbook, masterWB As Workbook
    Dim genSheet As Worksheet, secSheet As Worksheet
    Dim placeholders As Object
    Dim fSections As Object        ' For keys starting with F (grouped by first two characters)
    Dim aSections As Object        ' For keys starting with A (grouped by first two characters)
    Dim otherSections As Object    ' For keys starting with B, C, D, E, G, H, I, J, and for X keys (grouped as needed)
    Dim lastRowGeneral As Long, lastRow As Long, i As Long
    Dim key As String, rowData As Variant
    Dim groupID As String
    Dim masterWS As Worksheet

    ' Define file names
    masterFileName = "master_quotation_format.xlsx"
    genFileName = "Generated Quotation.xlsx"
    
    ' Check if either of the workbooks are open.
    If IsWorkbookOpen(masterFileName) Then
         MsgBox masterFileName & " is currently open and unsaved. Please save and close it before running the macro.", vbExclamation, "Workbook Open"
         Exit Sub
    End If
    
    If IsWorkbookOpen(genFileName) Then
         MsgBox genFileName & " is currently open and unsaved. Please save and close it before running the macro.", vbExclamation, "Workbook Open"
         Exit Sub
    End If
    
    '-------------------------------
    ' 1. Initialize dictionaries
    '-------------------------------
    Set placeholders = CreateObject("Scripting.Dictionary")
    Set fSections = CreateObject("Scripting.Dictionary")
    Set aSections = CreateObject("Scripting.Dictionary")
    Set otherSections = CreateObject("Scripting.Dictionary")
    
    '-----------------------------------------------
    ' 2. Open Quotation_Inputs.xlsx and extract data
    '-----------------------------------------------
    inputsPath = ThisWorkbook.Path & "\quotation_inputs.xlsx"
    Set inputsWB = Workbooks.Open(inputsPath)
    
    ' --- Read General Inputs from sheet "General Inputs"
    Set genSheet = inputsWB.Sheets("General Inputs")
    lastRowGeneral = genSheet.Cells(genSheet.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRowGeneral
         key = Trim(genSheet.Cells(i, "B").Value)
         key = Replace(key, ":", "")  ' Remove the colon, if present
         If key <> "" Then
             ' If the identifier is enclosed in double quotes, remove them.
             If Left(key, 1) = """" And Right(key, 1) = """" Then
                 key = Mid(key, 2, Len(key) - 2)
                 placeholders(key) = Array(genSheet.Cells(i, "C").Value, True)
             Else
                 placeholders(key) = Array(genSheet.Cells(i, "C").Value, False)
             End If
         End If
    Next i
    
    ' --- Read Section Inputs from sheet "Section Inputs"
    Set secSheet = inputsWB.Sheets("Section Inputs")
    lastRow = secSheet.Cells(secSheet.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
         key = Trim(secSheet.Cells(i, "B").Value)
         If key <> "" Then
             ' Load the row's values into an array.
             rowData = Array( _
                 key, _
                 secSheet.Cells(i, "C").Value, _
                 secSheet.Cells(i, "D").Value, _
                 secSheet.Cells(i, "E").Value, _
                 secSheet.Cells(i, "F").Value, _
                 secSheet.Cells(i, "G").Value _
             )
             ' Grouping based on the first character:
             Select Case UCase(Left(key, 1))
                 Case "F", "A", "X"   ' These groups have sub-identifier (e.g., F1, A1, X1)
                     groupID = Left(key, 2)
                     Select Case UCase(Left(key, 1))
                         Case "F"
                             If Not fSections.Exists(groupID) Then fSections.Add groupID, Array()
                             fSections(groupID) = AppendToArray(fSections(groupID), rowData)
                         Case "A"
                             If Not aSections.Exists(groupID) Then aSections.Add groupID, Array()
                             aSections(groupID) = AppendToArray(aSections(groupID), rowData)
                         Case "X"
                             ' Payment keys grouped as needed.
                             If Not otherSections.Exists(groupID) Then otherSections.Add groupID, Array()
                             otherSections(groupID) = AppendToArray(otherSections(groupID), rowData)
                     End Select
                 Case Else
                     ' For keys B, C, D, E, G, H, I, J:
                     groupID = UCase(Left(key, 1))
                     If Not otherSections.Exists(groupID) Then otherSections.Add groupID, Array()
                     otherSections(groupID) = AppendToArray(otherSections(groupID), rowData)
             End Select
         End If
    Next i
    
    Dim k As Variant
    For Each k In placeholders.Keys
        Debug.Print "Placeholder key: " & k & "  Value: " & placeholders(k)(0)
    Next k

    inputsWB.Close False  ' Close the inputs file
    
    '-----------------------------------------------
    ' 3. Open the master quotation template.
    '-----------------------------------------------
    Dim masterPath As String
    masterPath = ThisWorkbook.Path & "\dev(do not edit)\master_quotation_format.xlsx"
    Set masterWB = Workbooks.Open(masterPath)
    Set masterWS = masterWB.Sheets(1)  ' Adjust if needed
    
    '-----------------------------------------------
    ' 4. Update the master template with extracted inputs.
    '-----------------------------------------------
    
    ' 4a. Update header placeholders.
    UpdateHeader masterWS, placeholders
    
    ' 4a.1. Insert photo if specified in General Inputs.
    If placeholders.Exists("<<Photo>>") Then
         Dim photoName As String, photoPath As String
         photoName = placeholders("<<Photo>>")(0)
         photoPath = ThisWorkbook.Path & "\photos\" & photoName
         If Dir(photoPath) <> "" Then
              InsertPhoto masterWS, "<<Photo>>", photoPath
         Else
              MsgBox "Photo file not found: " & photoPath
         End If
    End If
    
    ' 4b. Update F sections.
    Dim fKey As Variant, sectionHeader As String
    For Each fKey In fSections.Keys
        Select Case UCase(fKey)
            Case "F1": sectionHeader = "F1. Manpower"
            Case "F2": sectionHeader = "F2. Accommodation"
            Case "F3": sectionHeader = "F3. Air Tickets"
            Case "F4": sectionHeader = "F4. Transportation"
            Case "F5": sectionHeader = "F5. Miscellaneous, tools, hardware, accessories"
            Case "F6": sectionHeader = "F6. Admin"
            Case "F7": sectionHeader = "F7. Photography"
            Case "F8": sectionHeader = "F8. Professional fee (PE endorsement)"
            Case "F9": sectionHeader = "F9. Courier and storage charges"
            Case "F10": sectionHeader = "F10. Preshow maintenance, packing"
            Case "F11": sectionHeader = "F11. Others"
            Case Else: sectionHeader = fKey
        End Select
        UpdateSection masterWS, sectionHeader, fSections(fKey)
    Next fKey

    ' 4c. Update A sections.
    Dim aKey As Variant
    For Each aKey In aSections.Keys
        Select Case UCase(aKey)
            Case "A1": sectionHeader = "A1. Flooring"
            Case "A2": sectionHeader = "A2. Hanging Banner/Structure"
            Case "A3": sectionHeader = "A3. System and Basic Structure"
            Case Else: sectionHeader = aKey
        End Select
        UpdateSection masterWS, sectionHeader, aSections(aKey)
    Next aKey

    ' 4d. Update other sections (B, C, D, E, G, H, I, J and X keys)
    Dim oKey As Variant
    For Each oKey In otherSections.Keys
         Select Case oKey
             Case "B": sectionHeader = "B. Graphics Materials & Printing"
             Case "C": sectionHeader = "C. Electrical Fittings and Lightings"
             Case "D": sectionHeader = "D. AV and LED"
             Case "E": sectionHeader = "E. Furniture Supply and Rental"
             Case "G": sectionHeader = "G. Air Condition Equipment"
             Case "H": sectionHeader = "H. Plants, Flowers and Arrangement"
             Case "I": sectionHeader = "I. Miscellaneous, Stationery and Printing"
             Case "J": sectionHeader = "J. Barista and Refreshments"
             Case Else
                  If Left(oKey, 1) = "X" Then
                        Select Case oKey
                           Case "X1": sectionHeader = "X1. Floral arrangements"
                           Case "X2": sectionHeader = "X2. Contractor Badges"
                           Case "X3": sectionHeader = "X3. Parking Passes"
                           Case "X4": sectionHeader = "X4. Stand Approval"
                           Case "X5": sectionHeader = "X5. Main Electrical connection"
                           Case "X6": sectionHeader = "X6. Build-up electrical connection"
                           Case "X7": sectionHeader = "X7. Internet connection"
                           Case "X8": sectionHeader = "X8. Rigging Services"
                           Case "X9": sectionHeader = "X9. Badges"
                           Case "X10": sectionHeader = "X10. Late charges"
                           Case "X11": sectionHeader = "X11. Others"
                           Case Else: sectionHeader = oKey
                        End Select
                  Else
                        sectionHeader = oKey
                  End If
         End Select
         UpdateSection masterWS, sectionHeader, otherSections(oKey)
    Next oKey
    
    ' 4e. (Optional) Update overall sub total cost if present.
    Dim cell As Range
    For Each cell In masterWS.UsedRange
         If Not IsError(cell.Value) Then
              If VarType(cell.Value) = vbString Then
                  If InStr(cell.Value, "Sub Total Cost (USD):") > 0 Then
                      cell.Value = "Sub Total Cost (USD): $24,390"
                  End If
              End If
         End If
    Next cell

    ' ----- Rename output files based on the Quotation Number -----
    Dim currentQuot As Long
    If placeholders.Exists("Quotation Number") Then
        currentQuot = CLng(placeholders("Quotation Number")(0))
    Else
        currentQuot = 1
    End If
    
    Dim newNameDoc As String, newNamePdf As String
    newNameDoc = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".docx"
    newNamePdf = ThisWorkbook.Path & "\Quotation" & Format(currentQuot, "000") & ".pdf"
    
    masterWB.SaveAs newNameDoc
    masterWB.ExportAsFixedFormat Type:=xlTypePDF, Filename:=newNamePdf
    ' ----- End Rename -----

    ' ----- Update the Quotation Number in the inputs file -----
    Dim updatedWB As Workbook
    Dim genInputsSheet As Worksheet
    Set updatedWB = Workbooks.Open(ThisWorkbook.Path & "\quotation_inputs.xlsx")
    Set genInputsSheet = updatedWB.Sheets("General Inputs")
    Dim j As Long
    For j = 3 To genInputsSheet.Cells(genInputsSheet.Rows.Count, "B").End(xlUp).Row
         If Trim(genInputsSheet.Cells(j, "B").Value) = "Quotation Number" Then
              genInputsSheet.Cells(j, "C").Value = currentQuot + 1
              Exit For
         End If
    Next j
    updatedWB.Save
    updatedWB.Close False
    ' ----- End update -----

    ' Automatically open the generated PDF based on the operating system:
    Dim pdfPath As String
    pdfPath = newNamePdf
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        Dim script As String
        script = "tell application ""Finder"" to open POSIX file """ & pdfPath & """" 
        MacScript script
    Else
        ShellExecute 0, "open", pdfPath, vbNullString, vbNullString, 1
    End If

    masterWB.Close False

    Dim fileNameDoc As String, fileNamePdf As String
    fileNameDoc = Mid(newNameDoc, InStrRev(newNameDoc, "\") + 1)
    fileNamePdf = Mid(newNamePdf, InStrRev(newNamePdf, "\") + 1)

    MsgBox "Quotation generated and saved as:" & vbCrLf & _
           "Word: " & fileNameDoc & vbCrLf & "PDF: " & fileNamePdf, vbInformation, "Quotation Generation"
End Sub
