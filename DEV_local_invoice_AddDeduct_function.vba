Option Explicit

Sub PopulateInvoiceTables_NewDoc()
    Dim ws            As Worksheet
    Dim lastRow       As Long
    Dim addStart As Long, addEnd As Long
    Dim dedStart As Long, dedEnd As Long
    Dim i             As Long
    Dim addCount      As Long, dedCount As Long
    
    '— 1) Read “inputs” sheet
    Set ws = ThisWorkbook.Sheets("inputs")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 1 To lastRow
        Select Case UCase(Trim(ws.Cells(i, "B").Value))
          Case "ADDITIONAL ITEMS", "ADDITIONAL"
            addStart = i
          Case "ADDITION SUBTOTAL"
            If addStart > 0 Then addEnd = i        ' include Subtotal row
          Case "DEDUCTION ITEMS", "DEDUCT"
            dedStart = i
          Case "DEDUCTION SUBTOTAL"
            If dedStart > 0 Then dedEnd = i        ' include Subtotal row
        End Select
    Next i
    
    If addStart = 0 Or addEnd < addStart _
      Or dedStart = 0 Or dedEnd < dedStart Then
        MsgBox "Could not locate all sections in 'inputs'!", vbExclamation
        Exit Sub
    End If
    
    addCount = addEnd - addStart + 1
    dedCount = dedEnd - dedStart + 1
    
    '— 2) Launch Word late‑bound and create a NEW document from the template
    Dim wdApp    As Object
    Dim wdDoc    As Object
    Dim fnd      As Object
    Dim rng      As Object
    Dim tbl      As Object
    Dim b        As Object
    Dim cel      As Object
    
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    
    ' New doc based on template
    Set wdDoc = wdApp.Documents.Add( _
       Template:=ThisWorkbook.Path & "\dev(do not edit)\master_invoice.docx", _
       NewTemplate:=False _
    )
    Set fnd = wdDoc.Content.Find
    
    '— 3) Insert ADDITION table
    With fnd
      .Text = "[[INSERT_ADDITION_TABLE_HERE]]"
      .MatchCase = True
      If .Execute Then
        Set rng = wdDoc.Range(.Parent.Start, .Parent.End)
        rng.Text = ""
        
        Set tbl = wdDoc.Tables.Add( _
          Range:=rng, _
          NumRows:=addCount, _
          NumColumns:=2 _
        )
        ' hide all borders
        For Each b In tbl.Borders: b.LineStyle = 0: Next b
        ' 60% width
        tbl.PreferredWidthType = 2
        tbl.PreferredWidth     = 60
        ' underline header row only
        tbl.Rows(1).Borders(4).LineStyle = 1
        
        ' fill table from inputs
        For i = 0 To addCount - 1
          tbl.Cell(i + 1, 1).Range.Text = ws.Cells(addStart + i, "B").Text
          tbl.Cell(i + 1, 2).Range.Text = ws.Cells(addStart + i, "C").Text
        Next i
        
        ' clear the “Addition SubTotal” label but keep its amount
        tbl.Cell(addCount, 1).Range.Text = ""
        
        ' bold & underline the header text in first cell
        With tbl.Cell(1, 1).Range.Font
          .Bold = True
          .Underline = 1
        End With
        
        ' align columns per cell
        For Each cel In tbl.Columns(1).Cells
          cel.Range.ParagraphFormat.Alignment = 0  ' left
        Next cel
        For Each cel In tbl.Columns(2).Cells
          cel.Range.ParagraphFormat.Alignment = 2  ' right
        Next cel
        
      Else
        MsgBox "Addition placeholder not found!", vbExclamation
      End If
    End With
    
    '— 4) Insert DEDUCTION table
    With fnd
      .Text = "[[INSERT_DEDUCTION_TABLE_HERE]]"
      .MatchCase = True
      If .Execute Then
        Set rng = wdDoc.Range(.Parent.Start, .Parent.End)
        rng.Text = ""
        
        Set tbl = wdDoc.Tables.Add( _
          Range:=rng, _
          NumRows:=dedCount, _
          NumColumns:=2 _
        )
        For Each b In tbl.Borders: b.LineStyle = 0: Next b
        tbl.PreferredWidthType = 2
        tbl.PreferredWidth     = 60
        tbl.Rows(1).Borders(4).LineStyle = 1
        
        For i = 0 To dedCount - 1
          tbl.Cell(i + 1, 1).Range.Text = ws.Cells(dedStart + i, "B").Text
          tbl.Cell(i + 1, 2).Range.Text = ws.Cells(dedStart + i, "C").Text
        Next i
        
        ' clear the “Deduction SubTotal” label but keep its amount
        tbl.Cell(dedCount, 1).Range.Text = ""
        
        ' bold & underline the header text in first cell
        With tbl.Cell(1, 1).Range.Font
          .Bold = True
          .Underline = 1
        End With
        
        For Each cel In tbl.Columns(1).Cells
          cel.Range.ParagraphFormat.Alignment = 0
        Next cel
        For Each cel In tbl.Columns(2).Cells
          cel.Range.ParagraphFormat.Alignment = 2
        Next cel
        
      Else
        MsgBox "Deduction placeholder not found!", vbExclamation
      End If
    End With
    
    '— 5) Save as new file
    Dim outPath As String
    outPath = ThisWorkbook.Path & "\generated_invoice.docx"
    wdDoc.SaveAs2 FileName:=outPath
    wdDoc.Close False
    wdApp.Quit
    
    Set wdDoc = Nothing
    Set wdApp  = Nothing
    
    MsgBox "New invoice created: " & vbCrLf & outPath, vbInformation
End Sub
