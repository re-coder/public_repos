Option Explicit

Public Sub InsertInvoiceAddDeductSections( _
    ByVal newDoc       As Object, _
    ByVal dataFilePath As String, _
    ByVal sheetName    As String _
)
    Const wdBorderBottom             As Long = 4
    Const wdPreferredWidthPercent    As Long = 2
    Const wdFindStop                 As Long = 0
    Const wdUnderlineSingle          As Long = 1
    
    Dim wb         As Workbook
    Dim ws         As Worksheet
    Dim lastRow    As Long
    Dim i          As Long, cnt As Long
    Dim addStart   As Long, addEnd   As Long
    Dim dedStart   As Long, dedEnd   As Long
    Dim addFlag    As String, dedFlag    As String
    Dim addItems() As String, addPrices() As String
    Dim dedItems() As String, dedPrices() As String
    
    On Error GoTo CleanExit
    
    ' 1) Open data workbook & sheet
    Set wb = Workbooks.Open(Filename:=dataFilePath, ReadOnly:=True)
    Set ws = wb.Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' 2) Find the start/end rows + YES/NO flags
    For i = 1 To lastRow
        Select Case UCase(Trim(ws.Cells(i, "B").Value))
          Case "ADDITIONAL", "ADDITIONAL ITEMS"
            addStart = i: addFlag = UCase(ws.Cells(i, "C").Text)
          Case "ADDITION SUBTOTAL:"
            If addStart > 0 Then addEnd = i
          Case "DEDUCT", "DEDUCTION ITEMS"
            dedStart = i: dedFlag = UCase(ws.Cells(i, "C").Text)
          Case "DEDUCTION SUBTOTAL:"
            If dedStart > 0 Then dedEnd = i
        End Select
    Next i
    
    If addStart = 0 Or addEnd < addStart _
      Or dedStart = 0 Or dedEnd < dedStart Then
        Err.Raise vbObjectError + 1, , "Section markers not found."
    End If
    
    ' 3) Load non-blank rows into arrays
    If addFlag = "YES" Then
      cnt = 0
      For i = addStart To addEnd
        If Len(Trim(ws.Cells(i, "B").Value)) > 0 Then cnt = cnt + 1
      Next i
      ReDim addItems(1 To cnt), addPrices(1 To cnt)
      cnt = 0
      For i = addStart To addEnd
        If Len(Trim(ws.Cells(i, "B").Value)) > 0 Then
          cnt = cnt + 1
          addItems(cnt)  = ws.Cells(i, "B").Text
          addPrices(cnt) = ws.Cells(i, "C").Text
        End If
      Next i
    End If
    
    cnt = 0
    If dedFlag = "YES" Then
      For i = dedStart To dedEnd
        If Len(Trim(ws.Cells(i, "B").Value)) > 0 Then cnt = cnt + 1
      Next i
      ReDim dedItems(1 To cnt), dedPrices(1 To cnt)
      cnt = 0
      For i = dedStart To dedEnd
        If Len(Trim(ws.Cells(i, "B").Value)) > 0 Then
          cnt = cnt + 1
          dedItems(cnt)  = ws.Cells(i, "B").Text
          dedPrices(cnt) = ws.Cells(i, "C").Text
        End If
      Next i
    End If
    
    ' 4) Prepare a clean Word.Find
    Dim fnd As Object
    Set fnd = newDoc.Content.Find
    With fnd
      .ClearFormatting
      .Replacement.ClearFormatting
      .Wrap = wdFindStop
      .Forward = True
    End With
    
    ' 5) Insert each table at its placeholder
    Dim placeholder As Variant
    For Each placeholder In Array( _
        "[[INSERT_ADDITION_TABLE_HERE]]", _
        "[[INSERT_DEDUCTION_TABLE_HERE]]" _
      )
      
      With fnd
        .Text = placeholder
        If .Execute Then
          Dim rng As Object
          Set rng = newDoc.Range(.Parent.Start, .Parent.End)
          rng.Text = ""
          
          Dim itemsArr As Variant, pricesArr As Variant
          If InStr(placeholder, "ADDITION") > 0 Then
            If addFlag <> "YES" Then GoTo SkipTbl
            itemsArr  = addItems:   pricesArr = addPrices
          Else
            If dedFlag <> "YES" Then GoTo SkipTbl
            itemsArr  = dedItems:   pricesArr = dedPrices
          End If
          
          Dim rowCount As Long
          rowCount = UBound(itemsArr)
          
          ' Build a 2-column, 60%-wide, borderless table
          Dim tbl As Object, b As Object, cel As Object
          Set tbl = newDoc.Tables.Add(rng, rowCount, 2)
          tbl.PreferredWidthType = wdPreferredWidthPercent
          tbl.PreferredWidth     = 60
          For Each b In tbl.Borders: b.LineStyle = 0: Next b
          
          ' Underline header bottom
          tbl.Cell(1, 1).Borders(wdBorderBottom).LineStyle = 1
          
          ' Fill and style, but blank out the header-row column 2
          Dim r As Long
          For r = 1 To rowCount
            tbl.Cell(r, 1).Range.Text = itemsArr(r)
            If r = 1 Then
              tbl.Cell(r, 2).Range.Text = ""           ' ← no more YES
            Else
              tbl.Cell(r, 2).Range.Text = pricesArr(r)
            End If
          Next r
          
          ' Header word bold + single-underline
          With tbl.Cell(1, 1).Range.Font
            .Bold = True
            .Underline = wdUnderlineSingle
          End With
          ' Subtotal row bold
          With tbl.Rows(rowCount).Range.Font
            .Bold = True
          End With
          
          ' Align: left column left, right column right, remove any col‑2 borders
          For Each cel In tbl.Columns(1).Cells
            cel.Range.ParagraphFormat.Alignment = 0
          Next cel
          For Each cel In tbl.Columns(2).Cells
            cel.Range.ParagraphFormat.Alignment = 2
            For Each b In cel.Borders: b.LineStyle = 0: Next b
          Next cel
        End If
      End With
SkipTbl:
    Next placeholder

CleanExit:
    On Error Resume Next
    wb.Close False
End Sub
