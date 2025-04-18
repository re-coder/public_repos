Public Sub ApplyCurrencyFormat(ws As Worksheet, currencyCode As String)
    Dim fmt         As String
    Dim c           As Range
    Dim symArr      As Variant
    Dim sym         As Variant
    Dim nf          As String
    
    Debug.Print "=== ApplyCurrencyFormat on " & ws.Name & " with " & currencyCode
    
    ' 1) Pick the right NumberFormat
    Select Case UCase(currencyCode)
      Case "USD": fmt = "$#,##0.00"
      Case "EUR": fmt = "€#,##0.00"
      Case "GBP": fmt = "£#,##0.00"
      Case "JPY": fmt = "¥#,##0"
      Case "CAD": fmt = "C$#,##0.00"
      Case "HKD": fmt = """HK""$#,##0.00"
      Case "RMB", "CNY": fmt = "¥#,##0.00"
      Case "SGD": fmt = """S""$#,##0.00"
      Case "MYR": fmt = """RM""#,##0.00"
      Case Else:  fmt = "$#,##0.00"
    End Select
    Debug.Print "  ? fmt = " & fmt
    
    symArr = Array("$", "€", "£", "¥", "RM", "C$", "HK$", "S$")
    
    ' 2) Scan every cell in UsedRange
    For Each c In ws.UsedRange
        nf = c.NumberFormat
        On Error Resume Next
        
        ' If it has the built-in Currency style, just overwrite
        If c.Style = "Currency" Then
            If c.MergeCells Then
                c.MergeArea.NumberFormat = fmt
                Debug.Print "    [" & c.Address(False, False) & "] merged style=Currency ? set fmt"
            Else
                c.NumberFormat = fmt
                Debug.Print "    [" & c.Address(False, False) & "] style=Currency ? set fmt"
            End If
        
        ' Otherwise look for any of our symbols in the existing format
        Else
            For Each sym In symArr
                If InStr(nf, sym) > 0 Then
                    If c.MergeCells Then
                        c.MergeArea.NumberFormat = fmt
                    Else
                        c.NumberFormat = fmt
                    End If
                    Debug.Print "    [" & c.Address(False, False) & "] detected '" & sym & "' ? set fmt"
                    Exit For
                End If
            Next sym
        End If
        
        On Error GoTo 0
    Next c
    
    Debug.Print "=== ApplyCurrencyFormat done ==="
End Sub

