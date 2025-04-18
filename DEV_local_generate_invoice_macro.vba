Option Explicit

Sub GenerateInvoice()
    Const wdFormatDefault As Long = 16
    Const wdReplaceAll      As Long = 2
    Const wdExportPDF       As Long = 17

    On Error GoTo ErrHandler

    Dim wdApp      As Object
    Dim wdDoc      As Object
    Dim newDoc     As Object
    Dim xlBook     As Workbook
    Dim xlSheet    As Worksheet
    Dim lastRow    As Long
    Dim i          As Long
    Dim tplPath    As String
    Dim savePath   As String
    Dim pdfPath    As String
    Dim dataPath   As String
    Dim searchText As String
    Dim replaceText As String

    Const SHEET_NAME As String = "Inputs"
    Const PL_COL     As String = "D"
    Const VAL_COL    As String = "C"

    ' 1) Start Word and make a copy of the template
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    tplPath = ThisWorkbook.Path & "\dev(do not edit)\master_invoice.docx"
    If Dir(tplPath) = "" Then Err.Raise 1001, , "Template missing: " & tplPath

    Set wdDoc = wdApp.Documents.Open(tplPath)
    savePath = ThisWorkbook.Path & "\Generated Invoice.docx"
    wdDoc.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDefault
    wdDoc.Close False

    Set newDoc = wdApp.Documents.Open(savePath)

    ' 2) Open the Inputs workbook
    dataPath = ThisWorkbook.Path & "\invoice_inputs.xlsx"
    If Dir(dataPath) = "" Then Err.Raise 1002, , "Data file missing: " & dataPath
    Set xlBook = Workbooks.Open(dataPath, ReadOnly:=True)
    Set xlSheet = xlBook.Worksheets(SHEET_NAME)

    ' 3) Replace simple placeholders (col D → col C)
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, PL_COL).End(xlUp).Row
    With newDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = 0
        For i = 3 To lastRow
            searchText  = Trim(xlSheet.Cells(i, PL_COL).Value)
            replaceText = xlSheet.Cells(i, VAL_COL).Value
            replaceText = Replace(replaceText, vbLf, vbCrLf)
            replaceText = Replace(replaceText, Chr(9), "")
            replaceText = Replace(replaceText, ChrW(160), " ")
            replaceText = Replace(replaceText, vbCrLf & " ", vbCrLf)
            If Len(searchText) > 0 Then
                .Text = searchText
                .Replacement.Text = replaceText
                .Execute Replace:=wdReplaceAll
            End If
        Next i
    End With

    ' 4) Insert the Additional & Deduction tables
    InsertInvoiceAddDeductSections newDoc, dataPath, SHEET_NAME

    ' 5) Save document and export as PDF
    newDoc.Save
    pdfPath = ThisWorkbook.Path & "\Generated Invoice.pdf"
    newDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportPDF

    MsgBox "Invoice generated successfully:" & vbCrLf & _
           savePath & vbCrLf & pdfPath, vbInformation

CleanUp:
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    Set xlSheet = Nothing
    Set xlBook  = Nothing
    Set newDoc  = Nothing
    wdApp.Quit
    Set wdApp   = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub
