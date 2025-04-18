Sub AutoNumberSections()
    Const INPUT_BASE_NAME As String = "quotation_inputs"
    Const INPUT_SHEET     As String = "Section Inputs"
    
    Dim wb        As Workbook
    Dim ws        As Worksheet
    Dim lastRow1  As Long, lastHdrRow2 As Long, lastDataRow2 As Long
    Dim hdrRows1  As Collection, hdrRows2 As Collection
    Dim r         As Long, i As Long, startRow As Long, endRow As Long, rr As Long
    Dim hdr       As String, prefix As String
    Dim dataCount As Long

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Debug.Print "=== Starting AutoNumberSections ==="
    
    ' — find your open "quotation_inputs" workbook —
    Debug.Print "Looking for an open workbook whose name begins with """ & INPUT_BASE_NAME & """..."
    For Each wb In Application.Workbooks
        If LCase$(Left$(wb.Name, Len(INPUT_BASE_NAME))) = INPUT_BASE_NAME Then
            Debug.Print "Found workbook: " & wb.Name
            Exit For
        End If
    Next wb
    If wb Is Nothing Then Err.Raise vbObjectError + 1, , "Could not find 'quotation_inputs' open"
    
    ' — get the Section Inputs sheet —
    Debug.Print "Looking for sheet """ & INPUT_SHEET & """ in " & wb.Name
    Set ws = wb.Sheets(INPUT_SHEET)
    If ws Is Nothing Then Err.Raise vbObjectError + 2, , "Sheet '" & INPUT_SHEET & "' not found"
    Debug.Print "Using sheet: " & ws.Name
    
    ' — build header row lists for Group 1 —
    Set hdrRows1 = New Collection
    lastRow1 = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For r = 1 To lastRow1
        hdr = Trim$(ws.Cells(r, "B").Value)
        If hdr <> "" And LCase$(hdr) <> "section item" Then
            hdrRows1.Add r
        End If
    Next r
    
    Debug.Print "Starting Group 1 numbering..."
    Debug.Print "  Last used row in B: " & lastRow1
    For i = 1 To hdrRows1.Count
        r = hdrRows1(i)
        hdr = Trim$(ws.Cells(r, "B").Value)
        Debug.Print "  Header found at row " & r & ": '" & hdr & "'"
        
        If i < hdrRows1.Count Then
            endRow = hdrRows1(i + 1) - 1
        Else
            endRow = lastRow1
        End If
        
        If hdr Like "[A-Za-z][0-9].*" Then
            prefix = Left$(hdr, InStr(hdr, "."))
            startRow = r + 1
            Debug.Print "    Detected letter+digit header ? prefix='" & prefix & "', startRow=" & startRow
        ElseIf hdr Like "[A-Za-z].*" Then
            prefix = Left$(hdr, 1)
            startRow = r + 2
            Debug.Print "    Detected letter-dot header ? prefix='" & prefix & "', startRow=" & startRow
        Else
            Debug.Print "    Unrecognized header format—skipping"
            GoTo NextHdr1
        End If
        
        Debug.Print "    Will number rows " & startRow & " to " & endRow
        dataCount = 0
        For rr = startRow To endRow
            If ws.Cells(rr, "D").Value <> "" Then
                dataCount = dataCount + 1
                ws.Cells(rr, "C").Value = prefix & dataCount
                Debug.Print "      Wrote C" & rr & " = " & ws.Cells(rr, "C").Value
            Else
                Debug.Print "      Skipped blank D" & rr
            End If
        Next rr
        Debug.Print "  Finished numbering section starting at row " & r
NextHdr1:
    Next i
    Debug.Print "Completed Group 1."
    
    ' — build header row lists for Group 2 —
    Set hdrRows2 = New Collection
    lastHdrRow2 = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    lastDataRow2 = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    For r = 1 To lastHdrRow2
        hdr = Trim$(ws.Cells(r, "K").Value)
        If hdr <> "" And LCase$(hdr) <> "section item" Then
            hdrRows2.Add r
        End If
    Next r
    
    Debug.Print "Starting Group 2 numbering..."
    Debug.Print "  Last header row in K: " & lastHdrRow2
    Debug.Print "  Last data row in M:   " & lastDataRow2
    For i = 1 To hdrRows2.Count
        r = hdrRows2(i)
        hdr = Trim$(ws.Cells(r, "K").Value)
        Debug.Print "  Header found at row " & r & ": '" & hdr & "'"
        
        If i < hdrRows2.Count Then
            endRow = hdrRows2(i + 1) - 1
        Else
            endRow = lastDataRow2
        End If
        
        If hdr Like "[A-Za-z][0-9].*" Then
            prefix = Left$(hdr, InStr(hdr, "."))
            startRow = r + 1
            Debug.Print "    Detected letter+digit header ? prefix='" & prefix & "', startRow=" & startRow
        ElseIf hdr Like "[A-Za-z].*" Then
            prefix = Left$(hdr, 1)
            startRow = r + 2
            Debug.Print "    Detected letter-dot header ? prefix='" & prefix & "', startRow=" & startRow
        Else
            Debug.Print "    Unrecognized header format—skipping"
            GoTo NextHdr2
        End If
        
        Debug.Print "    Will number rows " & startRow & " to " & endRow
        dataCount = 0
        For rr = startRow To endRow
            If ws.Cells(rr, "M").Value <> "" Then
                dataCount = dataCount + 1
                ws.Cells(rr, "L").Value = prefix & dataCount
                Debug.Print "      Wrote L" & rr & " = " & ws.Cells(rr, "L").Value
            Else
                Debug.Print "      Skipped blank M" & rr
            End If
        Next rr
        Debug.Print "  Finished numbering section starting at row " & r
NextHdr2:
    Next i
    Debug.Print "Completed Group 2."
    
    ' — save and finish —
    Debug.Print "Saving workbook " & wb.Name
    wb.Save
    Debug.Print "=== AutoNumberSections complete ==="
    MsgBox "Numbering complete!", vbInformation

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Debug.Print "? Error " & Err.Number & ": " & Err.Description
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "AutoNumberSections"
    Resume Cleanup
End Sub

