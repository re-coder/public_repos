Sub GenerateInvoice()
    ' Define Word constants for use with late-binding.
    Const wdFormatDocumentDefault As Long = 16
    Const wdFindContinue As Long = 1
    Const wdReplaceAll As Long = 2
    Const wdExportFormatPDF As Long = 17  ' New constant for PDF export

    On Error GoTo ErrorHandler

    Dim wdApp As Object         ' Word.Application
    Dim wdDoc As Object         ' Word.Document (template)
    Dim newDoc As Object        ' Word.Document (copy to edit)
    Dim xlBook As Workbook
    Dim xlSheet As Worksheet
    Dim filePath As String      ' Path to invoice_inputs.xlsx
    Dim savePath As String      ' Path to save the new Word document
    Dim pdfPath As String       ' Path to save the PDF version
    Dim templatePath As String  ' Path to the Word template master_invoice.docx
    Dim lastRow As Long
    Dim i As Long
    Dim searchText As String
    Dim replaceText As String

    Const PLACEHOLDER_COL As String = "D"
    Const VALUE_COL As String = "C"
    Const SHEET_NAME As String = "Inputs"

    Debug.Print "?? Starting placeholder replacement process..."

    ' Initialize Word application
    Set wdApp = CreateObject("Word.Application")

    wdApp.Visible = True
    Debug.Print "?? Word application started."

    ' Build the path to the template document (in subfolder "dev" of the Excel file's folder)
    templatePath = ThisWorkbook.Path & "\dev(do not edit)\master_invoice.docx"

    ' Check that the master template exists
    If Dir(templatePath) = "" Then
        MsgBox "Template file not found at: " & templatePath, vbExclamation
        Debug.Print "?? Template file missing: " & templatePath
        Exit Sub
    End If
    Debug.Print "?? Template file found: " & templatePath

    ' Open the master invoice template
    Set wdDoc = wdApp.Documents.Open(templatePath)
    Debug.Print "?? Master invoice template opened."

    ' Define the save path for the new invoice document (in the same folder as the Excel file)
    savePath = ThisWorkbook.Path & "\Generated Invoice.docx"
    
    ' Define the PDF save path
    pdfPath = ThisWorkbook.Path & "\Generated Invoice.pdf"

    ' Create a copy of the template to work with.
    wdDoc.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDocumentDefault
    Debug.Print "?? Document saved as: " & savePath

    ' Close the original template document (we now work with the new copy)
    wdDoc.Close SaveChanges:=False
    Debug.Print "?? Template document closed."

    ' Open the new copy to perform placeholder replacements
    Set newDoc = wdApp.Documents.Open(Filename:=savePath)
    Debug.Print "?? New document opened for editing."

    ' Build the path to the Excel file with replacement inputs (assumed in the same folder as the Excel workbook)
    filePath = ThisWorkbook.Path & "\invoice_inputs.xlsx"

    ' Check that the Excel inputs file exists
    If Dir(filePath) = "" Then
        MsgBox "Excel inputs file not found at: " & filePath, vbExclamation
        Debug.Print "?? Excel inputs file missing: " & filePath
        GoTo CleanExit
    End If
    Debug.Print "?? Excel inputs file found: " & filePath

    ' Open the invoice inputs workbook as read-only
    Set xlBook = Workbooks.Open(filePath, ReadOnly:=True)
    Debug.Print "?? Invoice inputs workbook opened: " & xlBook.Name

    ' Get the relevant worksheet
    On Error Resume Next
    Set xlSheet = xlBook.Worksheets(SHEET_NAME)
    On Error GoTo ErrorHandler
    If xlSheet Is Nothing Then
        MsgBox "Worksheet '" & SHEET_NAME & "' not found.", vbCritical
        Debug.Print "?? Worksheet missing: " & SHEET_NAME
        GoTo CleanExit
    End If
    Debug.Print "?? Worksheet found: " & SHEET_NAME

    ' Determine the last row based on the placeholder column (column D)
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, PLACEHOLDER_COL).End(xlUp).Row
    Debug.Print "?? Last row with data: " & lastRow

    ' Loop through each row and replace placeholders in the Word document
    For i = 3 To lastRow
        searchText = Trim(xlSheet.Cells(i, PLACEHOLDER_COL).Value)
        replaceText = xlSheet.Cells(i, VALUE_COL).Value

        ' Convert Excel's line feed (vbLf) to Word's carriage return-line feed (vbCrLf)
        replaceText = Replace(replaceText, vbLf, vbCrLf)
        ' Remove tab characters and replace non-breaking spaces with normal spaces
        replaceText = Replace(replaceText, Chr(9), "")
        replaceText = Replace(replaceText, ChrW(160), " ")
        ' Remove extra leading spaces after a new line
        replaceText = Replace(replaceText, vbCrLf & " ", vbCrLf)

        Debug.Print "?? Row " & i & ": Searching for '" & searchText & "'"

        If Len(searchText) > 0 Then
            With newDoc.Content.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = searchText
                .Replacement.Text = replaceText
                .Wrap = wdFindContinue
                .Execute Replace:=wdReplaceAll
            End With
            Debug.Print "?? Replaced with: '" & replaceText & "'"
        Else
            Debug.Print "?? Skipped empty placeholder in row " & i
        End If
    Next i

    newDoc.Save
    Debug.Print "?? Generated document saved."

    ' --- New functionality: Export the generated invoice as a PDF
    newDoc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdExportFormatPDF
    Debug.Print "?? PDF generated: " & pdfPath

    MsgBox "Invoice generated and saved as:" & vbCrLf & _
           "Word: 'Generated Invoice.docx'" & vbCrLf & "PDF: 'Generated Invoice.pdf'", vbInformation

CleanExit:
    ' Close the invoice inputs workbook without saving changes
    If Not xlBook Is Nothing Then xlBook.Close False
    ' Note: Do not quit the Word application so the generated invoice remains open
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set wdDoc = Nothing
    Set newDoc = Nothing
    Set wdApp = Nothing
    Debug.Print "?? Cleanup completed."
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Debug.Print "?? ERROR " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub



