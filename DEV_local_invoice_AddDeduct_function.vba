Option Explicit

Sub PopulateInvoiceTables_NewDoc()
    Dim ws              As Worksheet
    Dim lastRow         As Long
    Dim addStart As Long, addEnd As Long
    Dim dedStart As Long, dedEnd As Long
    Dim i               As Long, cnt As Long
    Dim addFlag As String, dedFlag As String
    
    '— 1) Locate blocks on “inputs” sheet
    Set ws = ThisWorkbook.Sheets("inputs")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 1 To lastRow
        Select Case UCase(Trim(ws.Cells(i, "B").Value))
          Case "ADDITIONAL", "ADDITIONAL ITEMS"
            addStart = i
            addFlag = UCase(Trim(ws.Cells(i, "C").Text))   ' YES/NO flag
          Case "ADDITION SUBTOTAL:"
            If addStart > 0 Then addEnd = i
          Case "DEDUCT", "DEDUCTION ITEMS"
            dedStart = i
            dedFlag = UCase(Trim(ws.Cells(i, "C").Text))
          Case "DEDUCTION SUBTOTAL:"
            If dedStart > 0 Then dedEnd = i
        End Select
    Next i
    
    If addStart = 0 Or addEnd < addStart _
      Or dedStart = 0 Or dedEnd < dedStart Then
        MsgBox "Could not locate all sections in 'inputs'!", vbExclamation
        Exit Sub
    End If
    
    '— 2) Pull non-blank rows into arrays
    Dim addItems()  As String, addPrices()  As String
    Dim dedItems()  As String, dedPrices()  As String
    
    If addFlag = "YES" Then
      cnt = 0
      For i = addStart To addEnd
        If Trim(ws.Cells(i, "B").Value) <> "" Then cnt = cnt + 1
      Next i
      ReDim addItems(1 To cnt), addPrices(1 To cnt)
      cnt = 0
      For i = addStart To addEnd
        If Trim(ws.Cells(i, "B").Value) <> "" Then
          cnt = cnt + 1
          addItems(cnt) = ws.Cells(i, "B").Text
          addPrices(cnt) = ws.Cells(i, "C").Text
        End If
      Next i
    End If
    
    If dedFlag = "YES" Then
      cnt = 0
      For i = dedStart To dedEnd
        If Trim(ws.Cells(i, "B").Value) <> "" Then cnt = cnt + 1
      Next i
      ReDim dedItems(1 To cnt), dedPrices(1 To cnt)
      cnt = 0
      For i = dedStart To dedEnd
        If Trim(ws.Cells(i, "B").Value) <> "" Then
          cnt = cnt + 1
          dedItems(cnt) = ws.Cells(i, "B").Text
          dedPrices(cnt) = ws.Cells(i, "C").Text
        End If
      Next i
    End If
    
    '— 3) Launch Word & new doc from template
    Dim wdApp    As Object, wdDoc As Object, fnd As Object
    Dim rng      As Object, tbl As Object, b As Object, cel As Object
    
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add( _
      Template:=ThisWorkbook.Path & "\dev(do not edit)\master_invoice.docx", _
      NewTemplate:=False _
    )
    Set fnd = wdDoc.Content.Find
    
    '— 4) Insert tables
    Dim placeholder As Variant
    Dim itemsArr   As Variant, pricesArr As Variant
    Dim rowCount   As Long
    
    For Each placeholder In Array( _
        "[[INSERT_ADDITION_TABLE_HERE]]", _
        "[[INSERT_DEDUCTION_TABLE_HERE]]" _
      )
      
      With fnd
        .Text = placeholder
        .MatchCase = True
        If .Execute Then
          Set rng = wdDoc.Range(.Parent.Start, .Parent.End)
          rng.Text = ""
          
          ' choose block
          If placeholder Like "*ADDITION*" Then
            If addFlag <> "YES" Then GoTo SkipTable
            itemsArr = addItems:   pricesArr = addPrices
          Else
            If dedFlag <> "YES" Then GoTo SkipTable
            itemsArr = dedItems:   pricesArr = dedPrices
          End If
          rowCount = UBound(itemsArr)
          
          ' add table
          Set tbl = wdDoc.Tables.Add( _
            Range:=rng, NumRows:=rowCount, NumColumns:=2 _
          )
          ' 60% width & no borders
          tbl.PreferredWidthType = 2: tbl.PreferredWidth = 60
          For Each b In tbl.Borders: b.LineStyle = 0: Next b
          
          ' underline header cell only
          tbl.Cell(1, 1).Borders(4).LineStyle = 1
          
          ' fill rows & clear YES/NO
          Dim r As Long
          For r = 1 To rowCount
            tbl.Cell(r, 1).Range.Text = itemsArr(r)
            tbl.Cell(r, 2).Range.Text = IIf(r = 1, "", pricesArr(r))
          Next r
          
          ' bold & underline header word
          With tbl.Cell(1, 1).Range.Font: .Bold = True: .Underline = 1: End With
          ' bold subtotal row
          With tbl.Rows(rowCount).Range.Font: .Bold = True: End With
          
          ' align & ensure no cell borders in col 2
          For Each cel In tbl.Columns(1).Cells
            cel.Range.ParagraphFormat.Alignment = 0
          Next cel
          For Each cel In tbl.Columns(2).Cells
            cel.Range.ParagraphFormat.Alignment = 2
            For Each b In cel.Borders: b.LineStyle = 0: Next b
          Next cel
        End If
      End With
SkipTable:
    Next placeholder
    
    '— 5) Save new file
    Dim outPath As String
    outPath = ThisWorkbook.Path & "\generated_invoice.docx"
    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close False
    wdApp.Quit
    
    Set wdDoc = Nothing: Set wdApp = Nothing
    
    MsgBox "New invoice created:" & vbCrLf & outPath, vbInformation
End Sub


