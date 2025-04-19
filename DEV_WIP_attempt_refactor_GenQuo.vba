Option Explicit

'==============================================
' Module: MainGenerate
' Description: Entry point and finalization routines
'==============================================
' Run this main sub to generate a quotation
Public Sub GenerateQuotation()
    On Error GoTo ErrorHandler
    Debug.Print "=== GenerateQuotation START ==="

    ' 0) Verify required workbooks are closed
    EnsureWorkbooksClosed "master_quotation_format.xlsx", "Generated Quotation.xlsx"

    ' 1) Load inputs
    Dim placeholders As Object, sectionsG1 As Object, sectionsG2 As Object
    Dim inputsWB As Workbook, currencyCode As String
    Set placeholders = CreateObject("Scripting.Dictionary")
    Set sectionsG1 = CreateObject("Scripting.Dictionary")
    Set sectionsG2 = CreateObject("Scripting.Dictionary")
    Set inputsWB = Workbooks.Open(ThisWorkbook.path & "\quotation_inputs.xlsx")
    ReadGeneralInputs inputsWB, placeholders
    currencyCode = ConfirmCurrency(placeholders)
    ReadSectionInputs inputsWB, sectionsG1, sectionsG2
    inputsWB.Close False

    ' 2) Open template and populate
    Dim masterWB As Workbook
    Set masterWB = Workbooks.Open(ThisWorkbook.path & "\dev(do not edit)\master_quotation_format.xlsx")
    UpdateHeader masterWB.Sheets(1), placeholders
    InsertPhoto masterWB.Sheets(1), "<<Photo>>", ThisWorkbook.path & "\photos\" & placeholders("<<Photo>>")(0)
    WriteSections masterWB.Sheets(1), sectionsG1, True
    WriteSections masterWB.Sheets(1), sectionsG2, False
    OverrideSubtotals masterWB.Sheets(1)
    ApplyCurrencyFormat masterWB.Sheets(1), currencyCode

    ' 3) Save/export & cleanup
    SaveAndExport masterWB, placeholders
    IncrementQuotationNumber ThisWorkbook.path & "\quotation_inputs.xlsx"
    OpenQuotationFile masterWB
    masterWB.Close False

    Debug.Print "=== GenerateQuotation COMPLETE ==="
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

'----------------------------------------------
' Writes all sections from dictionary under their headers
'----------------------------------------------
Private Sub WriteSections(ws As Worksheet, dict As Object, Optional isGroup1 As Boolean = True)
    Dim key As Variant
    For Each key In dict.Keys
        UpdateSection ws, key, dict(key)
    Next key
End Sub

'----------------------------------------------
' Overrides any subtotal text found
'----------------------------------------------
Private Sub OverrideSubtotals(ws As Worksheet)
    Dim c As Range
    For Each c In ws.UsedRange
        If VarType(c.Value) = vbString And InStr(c.Value, "Sub Total Cost") > 0 Then
            c.Value = "Sub Total Cost (USD): $24,390"
        End If
    Next c
End Sub

'----------------------------------------------
' Saves as .xlsx and exports to PDF
'----------------------------------------------
Private Sub SaveAndExport(wb As Workbook, dict As Object)
    Dim num As Long, xPath As String, pPath As String
    num = IIf(dict.Exists("Quotation Number"), CLng(dict("Quotation Number")(0)), 1)
    xPath = ThisWorkbook.path & "\Quotation" & Format(num, "000") & ".xlsx"
    pPath = ThisWorkbook.path & "\Quotation" & Format(num, "000") & ".pdf"
    wb.SaveAs xPath
    wb.ExportAsFixedFormat xlTypePDF, pPath
End Sub

'----------------------------------------------
' Increments the quotation number in inputs workbook
'----------------------------------------------
Private Sub IncrementQuotationNumber(path As String)
    Dim wb As Workbook, ws As Worksheet, r As Long
    Set wb = Workbooks.Open(path)
    Set ws = wb.Sheets("General Inputs")
    For r = 3 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        If Trim(ws.Cells(r, "B").Value) = "Quotation Number" Then
            ws.Cells(r, "C").Value = ws.Cells(r, "C").Value + 1
            Exit For
        End If
    Next r
    wb.Save
    wb.Close False
End Sub

'----------------------------------------------
' Opens the generated PDF file
'----------------------------------------------
Private Sub OpenQuotationFile(wb As Workbook)
    Dim pdfPath As String
    pdfPath = Replace(wb.FullName, ".xlsx", ".pdf")
    ShellExecute 0, "open", pdfPath, vbNullString, vbNullString, 1
End Sub






'----------------------------------------------'----------------------------------------------'----------------------------------------------'----------------------------------------------

Option Explicit

'----------------------------------------------
' Module: InputAndHelpers
' Description: Read inputs, confirm currency, update placeholders & sections
'----------------------------------------------
'--- Read General Inputs into dictionary with directReplace flag ---
Public Sub ReadGeneralInputs(wb As Workbook, dict As Object)
    Dim ws As Worksheet: Set ws = wb.Sheets("General Inputs")
    Dim lastRow As Long, i As Long, key As String, raw As String
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow
        raw = Trim(ws.Cells(i, "B").Value)
        If raw <> "" Then
            key = Replace(raw, ":", "")
            If Left(key, 1) = """" And Right(key, 1) = """" Then
                key = Mid(key, 2, Len(key) - 2)
                dict(key) = Array(ws.Cells(i, "C").Value, True)
            Else
                dict(key) = Array(ws.Cells(i, "C").Value, False)
            End If
        End If
    Next i
End Sub

'--- Confirm currency before proceeding ---
Public Function ConfirmCurrency(dict As Object) As String
    Dim code As String, resp As VbMsgBoxResult
    If Not dict.Exists("Currency") Then Err.Raise vbObjectError + 2, , "Currency missing"
    code = CStr(dict("Currency")(0))
    resp = MsgBox("Currency is " & code & ". Proceed?", vbYesNo + vbQuestion)
    If resp <> vbYes Then Err.Raise vbObjectError + 3, , "User cancelled"
    ConfirmCurrency = code
End Function

'--- Read Section Inputs into two dictionaries ---
Public Sub ReadSectionInputs(wb As Workbook, g1 As Object, g2 As Object)
    Dim ws As Worksheet: Set ws = wb.Sheets("Section Inputs")
    ReadGroup ws, "B", Array("C", "D", "E", "F", "G", "H"), g1
    ReadGroup ws, "L", Array("M", "N", "O", "P", "Q", "R"), g2
End Sub

Private Sub ReadGroup(ws As Worksheet, hdrCol As String, dataCols As Variant, dict As Object)
    Dim lastHdr As Long, lastData As Long, maxRow As Long
    Dim r As Long, skipRows As Long, hdr As String
    Dim arrData() As Variant, coll As Collection, idx As Long
    lastHdr = ws.Cells(ws.Rows.Count, hdrCol).End(xlUp).Row
    lastData = ws.Cells(ws.Rows.Count, dataCols(0)).End(xlUp).Row
    maxRow = Application.Max(lastHdr, lastData)
    r = 1
    Do While r <= maxRow
        hdr = Trim(ws.Cells(r, hdrCol).Value)
        If hdr <> "" And LCase(hdr) <> "section item" Then
            skipRows = IIf(Len(hdr) >= 2 And IsNumeric(Mid(hdr, 2, 1)), 1, 2)
            r = r + skipRows
            Set coll = New Collection
            Do While r <= maxRow And Trim(ws.Cells(r, hdrCol).Value) = ""
                ReDim arrData(0 To UBound(dataCols))
                For idx = LBound(dataCols) To UBound(dataCols)
                    arrData(idx) = ws.Cells(r, dataCols(idx)).Value
                Next idx
                If Not IsAllEmpty(arrData) Then coll.Add arrData
                r = r + 1
            Loop
            If coll.Count > 0 Then dict(hdr) = coll
        Else
            r = r + 1
        End If
    Loop
End Sub

'--- Update header placeholders in sheet ---
Public Sub UpdateHeader(ws As Worksheet, dict As Object)
    Dim cell As Range, key As Variant
    Dim tmp As Variant, newVal As String, direct As Boolean
    For Each cell In ws.UsedRange
        If VarType(cell.Value) = vbString Then
            For Each key In dict.Keys
                tmp = dict(key)
                newVal = CStr(tmp(0)): direct = CBool(tmp(1))
                If direct Then
                    If InStr(cell.Value, key) > 0 Then cell.Value = newVal
                ElseIf Trim(cell.Value) Like key & ":*" Then
                    cell.Value = key & ": " & newVal
                End If
            Next key
        End If
    Next cell
End Sub

'--- Update a section with an array of data rows ---
Public Sub UpdateSection(ws As Worksheet, sectionHeader As String, dataList As Variant)
    Dim pos As Variant, startRow As Long
    Dim rawCount As Long, numRows As Long, numCols As Long
    Dim i As Long, j As Long, colIndex As Long, outRow As Long
    Dim outputArr() As Variant, rng As Range
    pos = FindCell(ws, sectionHeader)
    startRow = pos(0)
    If startRow = 0 Then Exit Sub
    rawCount = UBound(dataList) - LBound(dataList) + 1
    numRows = IIf(rawCount Mod 2 = 0, rawCount, rawCount + 1)
    numCols = UBound(dataList(LBound(dataList))) - LBound(dataList(LBound(dataList))) + 1
    ws.Rows(startRow + 1).Resize(numRows).Insert Shift:=xlDown
    ReDim outputArr(1 To numRows, 1 To numCols)
    For i = LBound(dataList) To UBound(dataList)
        outRow = i - LBound(dataList) + 1: colIndex = 1
        For j = LBound(dataList(i)) To UBound(dataList(i))
            outputArr(outRow, colIndex) = dataList(i)(j)
            colIndex = colIndex + 1
        Next j
    Next i
    Set rng = ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + numRows, numCols))
    rng.Value = outputArr
    ' reset formatting
    With rng.Font: .Bold = False: .Underline = xlUnderlineStyleNone: .Italic = False: End With
    ' apply shading
    For i = 1 To numRows
        With ws.Rows(startRow + i).Interior
            .Color = IIf(i Mod 2 = 1, RGB(242, 242, 242), RGB(255, 255, 255))
        End With
    Next i
End Sub

'--- Insert a photo where placeholderKey appears ---
Public Sub InsertPhoto(ws As Worksheet, placeholderKey As String, photoPath As String)
    Dim fCell As Range, target As Range, pic As Shape
    Set fCell = ws.Cells.Find(What:=placeholderKey, LookIn:=xlValues, LookAt:=xlPart)
    If fCell Is Nothing Then Exit Sub
    Set target = fCell.MergeArea: target.Value = ""
    On Error Resume Next
    Set pic = ws.Shapes.AddPicture(photoPath, msoFalse, msoTrue, target.Left, target.Top, target.Width, target.Height)
    On Error GoTo 0
    If pic Is Nothing Then Exit Sub
    With pic: .LockAspectRatio = msoFalse: .Placement = xlMoveAndSize: .ZOrder msoBringToFront: End With
End Sub





'----------------------------------------------'----------------------------------------------'----------------------------------------------'----------------------------------------------
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

'----------------------------------------------
' Utility: Check if a workbook is open
'----------------------------------------------
Public Function IsWorkbookOpen(wbName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0
    IsWorkbookOpen = Not wb Is Nothing
End Function

'----------------------------------------------
' Utility: Ensure required workbooks are closed
'----------------------------------------------
Public Sub EnsureWorkbooksClosed(ParamArray names() As Variant)
    Dim nm As Variant
    For Each nm In names
        If IsWorkbookOpen(CStr(nm)) Then
            MsgBox nm & " is open. Close it before proceeding.", vbExclamation
            Err.Raise vbObjectError + 1, , "Workbook open"
        End If
    Next nm
End Sub

'----------------------------------------------
' Utility: Returns True if every element of array is empty
'----------------------------------------------
Public Function IsAllEmpty(arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Trim(CStr(arr(i))) <> "" Then Exit Function
    Next i
    IsAllEmpty = True
End Function

'----------------------------------------------
' Utility: Append a value to an array
'----------------------------------------------
Public Function AppendToArray(oldArray As Variant, newValue As Variant) As Variant
    Dim newArray() As Variant, i As Long, n As Long
    On Error Resume Next: n = UBound(oldArray): On Error GoTo 0
    If n < 0 Then
        ReDim newArray(0 To 0)
        newArray(0) = newValue
    Else
        ReDim newArray(0 To n + 1)
        For i = 0 To n
            newArray(i) = oldArray(i)
        Next i
        newArray(n + 1) = newValue
    End If
    AppendToArray = newArray
End Function

'----------------------------------------------
' Utility: Find a cell containing text; returns (row, col)
'----------------------------------------------
Public Function FindCell(ws As Worksheet, searchText As String) As Variant
    Dim cell As Range
    For Each cell In ws.UsedRange
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = vbString And InStr(cell.Value, searchText) > 0 Then
                FindCell = Array(cell.Row, cell.Column)
                Exit Function
            End If
        End If
    Next cell
    FindCell = Array(0, 0)
End Function


