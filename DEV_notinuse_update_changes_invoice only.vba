Option Explicit

Sub UpdateChanges()
    Dim wbInputs     As Workbook
    Dim shtInputs    As Worksheet
    Dim secSheet     As Worksheet
    Dim currencyCell As Range
    Dim currCode     As String

    Debug.Print "=== UpdateChanges START ==="
    
    ' 1) Get or open quotation_inputs.xlsx
    Debug.Print "Looking for already-open workbook 'quotation_inputs.xlsx'..."
    On Error Resume Next
    Set wbInputs = Workbooks("quotation_inputs.xlsx")
    On Error GoTo 0
    If wbInputs Is Nothing Then
        Debug.Print "'quotation_inputs.xlsx' not open — opening from path..."
        Set wbInputs = Workbooks.Open(ThisWorkbook.Path & "\quotation_inputs.xlsx")
    Else
        Debug.Print "'quotation_inputs.xlsx' already open."
    End If
    Debug.Print "Using inputs workbook: " & wbInputs.Name

    ' 2) Point to the General Inputs sheet
    Set shtInputs = wbInputs.Sheets("General Inputs")
    Debug.Print "Pointed to sheet: " & shtInputs.Name

    ' 3) Find the cell that literally contains Currency
    Debug.Print "Searching for 'Currency' in " & shtInputs.Name & "..."
    Set currencyCell = shtInputs.UsedRange.Find( _
        What:=Chr(34) & "Currency" & Chr(34), _
        LookAt:=xlWhole, _
        LookIn:=xlValues)

    If currencyCell Is Nothing Then
        Debug.Print "ERROR: 'Currency' not found."
        MsgBox "Could not find 'Currency' on sheet " & shtInputs.Name, vbExclamation
        Exit Sub
    Else
        Debug.Print "Found 'Currency' at " & currencyCell.Address
    End If

    ' 4) Read the code from the cell to its right
    currCode = Trim$(currencyCell.Offset(0, 1).Value)
    Debug.Print "Read currency code: '" & currCode & "'"
    If currCode = "" Then
        Debug.Print "ERROR: No currency code to the right of Currency."
        MsgBox "No currency code found to the right of Currency.", vbExclamation
        Exit Sub
    End If

    ' 5) Apply to both sheets
    Debug.Print "Applying currency format (" & currCode & ") to " & shtInputs.Name & "..."
    ApplyCurrencyFormat shtInputs, currCode
    Debug.Print "Applied currency format to " & shtInputs.Name

    Set secSheet = wbInputs.Sheets("Section Inputs")
    Debug.Print "Applying currency format (" & currCode & ") to " & secSheet.Name & "..."
    ApplyCurrencyFormat secSheet, currCode
    Debug.Print "Applied currency format to " & secSheet.Name

    Debug.Print "=== UpdateChanges COMPLETE ==="
End Sub


