Option Explicit

Sub GenerateInvoice()
    Dim wbOpen   As Workbook
    Dim ff       As Integer
    Dim tplPath  As String
    Dim savePath As String
    

    '————————————————————————————————————
    ' 0a) Refuse if invoice_inputs.xlsx is open
    '————————————————————————————————————
    For Each wbOpen In Application.Workbooks
        If LCase(wbOpen.name) = "invoice_inputs.xlsx" Then
            MsgBox "Please close invoice_inputs.xlsx before running this macro.", vbExclamation
            Exit Sub
        End If
    Next wbOpen

    '————————————————————————————————————
    ' 0b) Refuse if master_invoice.docx is open
    '————————————————————————————————————
    tplPath = ThisWorkbook.path & "\dev(do not edit)\master_invoice.docx"
    If Not CanLockFile(tplPath) Then
        MsgBox "Please close master_invoice.docx in Word before running.", vbExclamation
        Exit Sub
    End If

    '————————————————————————————————————
    ' 0c) Refuse if a previous Generated Invoice.docx is open
    '————————————————————————————————————
    savePath = ThisWorkbook.path & "\Generated Invoice.docx"
    If Not CanLockFile(savePath) Then
        MsgBox "Please close Generated Invoice.docx in Word before running.", vbExclamation
        Exit Sub
    End If

    '— Now safe to proceed…
    Application.StatusBar = "Software is running – do not move your cursor while the document is being printed."

    Const wdFormatDefault As Long = 16
    Const wdReplaceAll      As Long = 2
    Const wdExportPDF       As Long = 17
    On Error GoTo ErrHandler

    Dim wdApp      As Object, wdDoc As Object, newDoc As Object
    Dim xlBook     As Workbook, xlSheet As Worksheet
    Dim lastRow    As Long, i As Long
    Dim dataPath   As String
    Dim searchText As String, replaceText As String

    Const SHEET_NAME As String = "Inputs"
    Const PL_COL     As String = "D"
    Const VAL_COL    As String = "C"

    ' 1) Start Word, silence alerts, open template ReadOnly
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    wdApp.DisplayAlerts = 0

    Set wdDoc = wdApp.Documents.Open( _
        Filename:=tplPath, _
        ReadOnly:=True, _
        AddToRecentFiles:=False _
    )
    wdDoc.SaveAs2 Filename:=savePath, FileFormat:=wdFormatDefault
    wdDoc.Close False
    Set newDoc = wdApp.Documents.Open( _
        Filename:=savePath, _
        ReadOnly:=False, _
        AddToRecentFiles:=False _
    )

    ' 2) Open the Inputs workbook
    dataPath = ThisWorkbook.path & "\invoice_inputs.xlsx"
    If Dir(dataPath) = "" Then Err.Raise 1002, , "Data file missing: " & dataPath
    Set xlBook = Workbooks.Open(Filename:=dataPath)
    Set xlSheet = xlBook.Worksheets(SHEET_NAME)

    ' 2a) Currency formatting & confirmation…
    Dim currencyCell As Range, currencyCode As String
    Set currencyCell = xlSheet.UsedRange.Find( _
        What:="Currency", LookAt:=xlWhole, LookIn:=xlValues)
    If Not currencyCell Is Nothing Then
        currencyCode = Trim(currencyCell.Offset(0, 1).Value)
        ApplyCurrencyFormat xlSheet, currencyCode
        If MsgBox("Applied currency format using code '" & currencyCode & "'" & vbCrLf & _
                  "Continue generating invoice?", vbQuestion + vbYesNo, "Confirm Currency") <> vbYes Then
            MsgBox "Cancelled. Inputs file remains open.", vbInformation
            Exit Sub
        End If
    Else
        MsgBox "Could not find 'Currency' cell—skipping formatting.", vbExclamation
    End If

    ' 3) Replace placeholders (col D ? col C)
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, PL_COL).End(xlUp).Row
    With newDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = 0
        For i = 3 To lastRow
            searchText = Trim(xlSheet.Cells(i, PL_COL).Value)
            replaceText = xlSheet.Cells(i, VAL_COL).Text
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

    ' 4) Insert tables
    InsertInvoiceAddDeductSections newDoc, dataPath, SHEET_NAME

    ' 5) Save & export PDF
    newDoc.Save
    newDoc.ExportAsFixedFormat _
      OutputFileName:=ThisWorkbook.path & "\Generated Invoice.pdf", _
      ExportFormat:=wdExportPDF

    MsgBox "Invoice generated successfully!", vbInformation

CleanUp:
    On Error Resume Next
    Application.StatusBar = False

    If Not xlBook Is Nothing Then xlBook.Close SaveChanges:=True
    If Not newDoc Is Nothing Then newDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set xlSheet = Nothing: Set xlBook = Nothing
    Set newDoc = Nothing: Set wdApp = Nothing
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

'————————————————————————————————————
' Helper: try to lock a file for read/write
' Returns True if successful (i.e. file is not open elsewhere)
'————————————————————————————————————
Private Function CanLockFile(path As String) As Boolean
    Dim ff As Integer
    On Error Resume Next
    ff = FreeFile
    ' Attempt exclusive lock for read/write
    Open path For Binary Access Read Write Lock Read Write As #ff
    If Err.Number = 0 Then
        CanLockFile = True
        Close #ff
    Else
        CanLockFile = False
        Err.Clear
    End If
    On Error GoTo 0
End Function


