Option Explicit

'----------------------------------------------
' AppendToArray: Appends a new element to an existing array.
'----------------------------------------------
Public Function AppendToArray(oldArray As Variant, newValue As Variant) As Variant
    Dim newArray() As Variant
    Dim i As Long, n As Long
    
    On Error Resume Next
    n = UBound(oldArray)
    On Error GoTo 0
    
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
' UpdateHeader: Scans the worksheet and updates any cell that starts with a key.
'----------------------------------------------
'----------------------------------------------
' UpdateHeader: Scans the worksheet and updates any cell that starts with a key.
'----------------------------------------------
Public Sub UpdateHeader(ws As Worksheet, placeholders As Object)
    Dim cell As Range
    Dim key  As Variant
    Dim tmp  As Variant
    Dim newVal       As String
    Dim directReplace As Boolean

    For Each cell In ws.UsedRange
        If Not IsError(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                For Each key In placeholders.Keys
                    tmp = placeholders(key)
                    ' only treat as array if it really is one:
                    If IsArray(tmp) Then
                        newVal = CStr(tmp(0))
                        directReplace = CBool(tmp(1))
                    Else
                        newVal = CStr(tmp)
                        directReplace = False
                    End If

                    If directReplace Then
                        If InStr(cell.Value, key) > 0 Then
                            cell.Value = newVal
                        End If
                    Else
                        If Trim(cell.Value) Like key & ":*" Then
                            cell.Value = key & ": " & newVal
                        End If
                    End If
                Next key
            End If
        End If
    Next cell
End Sub


'----------------------------------------------
' UpdateSection: Finds the section header and writes the data (array of arrays) below that header.
'----------------------------------------------
Public Sub UpdateSection(ws As Worksheet, sectionHeader As String, dataList As Variant)
    Dim pos        As Variant
    Dim startRow   As Long
    Dim i          As Long, j As Long
    Dim rawCount   As Long, numRows As Long, numCols As Long, colIndex As Long
    Dim outputArr() As Variant
    Dim rng        As Range
    Dim r          As Long

    On Error GoTo CleanUp
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' find the header
    pos = FindCell(ws, sectionHeader)
    startRow = pos(0)
    If startRow = 0 Then
        Debug.Print "Section header '" & sectionHeader & "' not found."
        GoTo CleanUp
    End If

    ' determine how many rows to insert (even number)
    rawCount = UBound(dataList) - LBound(dataList) + 1
    If rawCount Mod 2 = 0 Then
        numRows = rawCount
    Else
        numRows = rawCount + 1
    End If

    ' count every element in the first data row
    numCols = 0
    For j = LBound(dataList(LBound(dataList))) To UBound(dataList(LBound(dataList)))
        numCols = numCols + 1
    Next j

    ' make space
    ws.Rows(startRow + 1).Resize(numRows).Insert Shift:=xlDown

    ' prepare the array
    ReDim outputArr(1 To numRows, 1 To numCols)

    ' fill the array with all elements of each rowData
    For i = LBound(dataList) To UBound(dataList)
        Dim outRow As Long
        outRow = i - LBound(dataList) + 1
        colIndex = 1
        For j = LBound(dataList(i)) To UBound(dataList(i))
            outputArr(outRow, colIndex) = dataList(i)(j)
            colIndex = colIndex + 1
        Next j
    Next i

    ' dump into sheet
    ws.Range(ws.Cells(startRow + 1, 1), _
             ws.Cells(startRow + numRows, numCols)).Value = outputArr

    ' reset formatting
    Set rng = ws.Range(ws.Cells(startRow + 1, 1), _
                       ws.Cells(startRow + numRows, numCols))
    With rng.Font
        .Bold = False
        .Underline = xlUnderlineStyleNone
        .Italic = False
    End With

    ' apply alternating row shading
    For r = 1 To numRows
        With ws.Range(ws.Cells(startRow + r, 1), ws.Cells(startRow + r, numCols)).Interior
            If r Mod 2 = 1 Then
                .Color = RGB(242, 242, 242)
            Else
                .Color = RGB(255, 255, 255)
            End If
        End With
    Next r

CleanUp:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'----------------------------------------------
' FindCell: Searches the used range for a cell containing searchText.
' Returns an array {row, column}; if not found, returns {0,0}.
'----------------------------------------------
Public Function FindCell(ws As Worksheet, searchText As String) As Variant
    Dim cell As Range
    For Each cell In ws.UsedRange
         If Not IsError(cell.Value) Then
             If VarType(cell.Value) = vbString Then
                 If InStr(cell.Value, searchText) > 0 Then
                     FindCell = Array(cell.Row, cell.Column)
                     Exit Function
                 End If
             End If
         End If
    Next cell
    FindCell = Array(0, 0)
End Function

'----------------------------------------------
' InsertPhoto: Inserts a photo from a file path and stretches it to fit the target cell.
'----------------------------------------------
Public Sub InsertPhoto(ws As Worksheet, placeholderKey As String, photoPath As String)
    Dim foundRange As Range
    Dim targetRange As Range
    Dim picShape As Shape
    Dim fullWidth As Double, fullHeight As Double
    
    Set foundRange = ws.Cells.Find(What:=placeholderKey, LookIn:=xlValues, _
                                   LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If foundRange Is Nothing Then
        MsgBox "Cell for placeholder '" & placeholderKey & "' not found in master sheet."
        Exit Sub
    End If
    
    Set targetRange = foundRange.MergeArea
    targetRange.Value = ""
    fullWidth = targetRange.Width
    fullHeight = targetRange.Height
    
    On Error Resume Next
    Set picShape = ws.Shapes.AddPicture( _
                    Filename:=photoPath, _
                    LinkToFile:=msoFalse, _
                    SaveWithDocument:=msoTrue, _
                    Left:=targetRange.Left, _
                    Top:=targetRange.Top, _
                    Width:=fullWidth, _
                    Height:=fullHeight)
    On Error GoTo 0
    
    If picShape Is Nothing Then
        MsgBox "Unable to insert the photo. Please check the file: " & photoPath
        Exit Sub
    End If
    
    With picShape
        .LockAspectRatio = msoFalse
        .Placement = xlMoveAndSize
        .ZOrder msoBringToFront
    End With
End Sub

'----------------------------------------------
' IsWorkbookOpen: Checks if a workbook with the specified name is open.
'----------------------------------------------
Public Function IsWorkbookOpen(wbName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0
    IsWorkbookOpen = Not wb Is Nothing
End Function


'----------------------------------------------
' Helper: Returns True if every element of arr is blank
'----------------------------------------------
Public Function IsAllEmpty(arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Trim(CStr(arr(i))) <> "" Then
            IsAllEmpty = False
            Exit Function
        End If
    Next i
    IsAllEmpty = True
End Function






