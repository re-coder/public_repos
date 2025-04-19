Option Explicit

Sub UpdateChangesQuo()
    Dim wb           As Workbook
    Dim ws           As Worksheet
    Dim currencyCell As Range
    Dim currCode     As String

    Set wb = ActiveWorkbook

    ' ——— 1) Read currency code from the active sheet ———
    With ActiveSheet
        Set currencyCell = .UsedRange.Find( _
        What:=Chr(34) & "Currency" & Chr(34), _
            LookAt:=xlWhole, LookIn:=xlValues)
        If currencyCell Is Nothing Then
            MsgBox "Could not find the cell containing 'Currency' on the active sheet.", vbExclamation
            Exit Sub
        End If
        
        currCode = Trim$(currencyCell.Offset(0, 1).Value)
        If currCode = "" Then
            MsgBox "No currency code found to the right of 'Currency' on the active sheet.", vbExclamation
            Exit Sub
        End If
    End With

    ' ——— 2) Loop all sheets with “inputs” in their name ———
    For Each ws In wb.Worksheets
        If InStr(1, ws.name, "inputs", vbTextCompare) > 0 Then
            ApplyCurrencyFormat ws, currCode
        End If
    Next ws

    MsgBox "Currency format (" & currCode & ") applied successfully to every worksheet with '*inputs*' in its name.", vbInformation
End Sub



