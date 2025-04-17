Sub AutoNumberSections()
    Const INPUT_WB_NAME As String = "quotation_inputs.xlsx"
    Const INPUT_SHEET  As String = "Section Inputs"
    
    Dim wb As Workbook, ws As Worksheet
    Dim folder As String
    
    ' 1) Ensure we have the inputs workbook open (or open it)
    On Error Resume Next
    Set wb = Workbooks(INPUT_WB_NAME)
    On Error GoTo 0
    
    If wb Is Nothing Then
        ' assume it's in the same folder as this macro workbook
        folder = ThisWorkbook.Path
        If folder = "" Then
            MsgBox "Can't determine macro workbook folder.", vbCritical
            Exit Sub
        End If
        On Error Resume Next
        Set wb = Workbooks.Open(folder & "\" & INPUT_WB_NAME)
        On Error GoTo 0
        If wb Is Nothing Then
            MsgBox "Could not open " & INPUT_WB_NAME, vbCritical
            Exit Sub
        End If
    End If
    
    ' 2) Get the Section Inputs sheet
    On Error Resume Next
    Set ws = wb.Sheets(INPUT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet '" & INPUT_SHEET & "' not found in " & INPUT_WB_NAME, vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim lastRow1 As Long, lastRow2 As Long
    Dim r As Long, rr As Long, startRow As Long
    Dim hdr As String, prefix As String
    Dim dataCount As Long
    
    ' -- Group 1: headers in Col B, data in C:D:E:F:G --
    lastRow1 = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    r = 1
    Do While r <= lastRow1
        hdr = Trim(ws.Cells(r, "B").Value)
        If hdr <> "" And LCase(hdr) <> "section item" Then
            ' decide prefix & first-data-row
            If hdr Like "[A-Za-z].*" Then
                ' letter·dot header (e.g. "C. ...")
                prefix = Left(hdr, 1)                ' "C"
                startRow = r + 2                     ' skip header+title
            ElseIf hdr Like "[A-Za-z][0-9].*" Then
                ' letter·digit·dot header (e.g. "A1. ...")
                prefix = Left(hdr, InStr(hdr, "."))
                startRow = r + 1                     ' skip header only
            Else
                r = r + 1: GoTo NextGroup1
            End If
            
            ' fill down until Col D is blank
            rr = startRow
            Do While rr <= lastRow1 And ws.Cells(rr, "D").Value <> ""
                dataCount = Application.CountIf( _
                    ws.Range(ws.Cells(startRow, "D"), ws.Cells(rr, "D")), "<>")
                ' build formula in Col C
                ws.Cells(rr, "C").Formula = "=IF(" & _
                    "D" & rr & "="""","""",""" & prefix & """ & " & dataCount & ")"
                rr = rr + 1
            Loop
            r = rr
        Else
NextGroup1:
            r = r + 1
        End If
    Loop
    
    ' -- Group 2: headers in Col K, data in L:M:N:O:P --
    lastRow2 = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    r = 1
    Do While r <= lastRow2
        hdr = Trim(ws.Cells(r, "K").Value)
        If hdr <> "" And LCase(hdr) <> "section item" Then
            If hdr Like "[A-Za-z].*" Then
                prefix = Left(hdr, 1)
                startRow = r + 2
            ElseIf hdr Like "[A-Za-z][0-9].*" Then
                prefix = Left(hdr, InStr(hdr, "."))
                startRow = r + 1
            Else
                r = r + 1: GoTo NextGroup2
            End If
            
            rr = startRow
            Do While rr <= lastRow2 And ws.Cells(rr, "L").Value <> ""
                dataCount = Application.CountIf( _
                    ws.Range(ws.Cells(startRow, "L"), ws.Cells(rr, "L")), "<>")
                ws.Cells(rr, "C").Formula = "=IF(" & _
                    "L" & rr & "="""","""",""" & prefix & """ & " & dataCount & ")"
                rr = rr + 1
            Loop
            r = rr
        Else
NextGroup2:
            r = r + 1
        End If
    Loop
    
    ' 3) Save & finish
    wb.Save
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Numbering complete!", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub




