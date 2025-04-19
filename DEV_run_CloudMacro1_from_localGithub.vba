Option Explicit

Sub RunCloudMacro1FromTextFile()
    ' Declaration of variables
    Dim fso            As Object
    Dim ts             As Object
    Dim vbaCode        As String
    Dim filePath       As String
    Dim vbProj         As Object
    Dim vbComp         As Object
    Dim moduleName     As String
    Dim subroutineName As String
    Dim fileOpened     As Boolean
    Dim lines()        As String

    ' === CONFIGURATION ===
    filePath       = "C:\Users\stile\OneDrive\Desktop\GitHub\public_repos\DEV_local_generate_quotation_macro.vba"
    moduleName     = "CldMacro1BENQuo"
    subroutineName = "GenerateQuotation"  ' Update if needed

    fileOpened = False
    On Error GoTo ErrorHandler

    ' === Load file content ===
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "File not found: " & filePath, vbExclamation
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(filePath, 1) ' 1 = ForReading
    fileOpened = True
    vbaCode = ts.ReadAll
    ts.Close
    fileOpened = False

    ' === Remove trailing "()” from the code string ===
    vbaCode = RTrim$(vbaCode)
    If Right$(vbaCode, 2) = "()" Then
        vbaCode = Left$(vbaCode, Len(vbaCode) - 2)
    End If

    lines = Split(vbaCode, vbCrLf)
    Do While UBound(lines) >= 0
        If Trim$(lines(UBound(lines))) = "()" Then
            ReDim Preserve lines(UBound(lines) - 1)
        Else
            Exit Do
        End If
    Loop
    vbaCode = Join(lines, vbCrLf)

    ' === Ensure access to the VBProject ===
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo ErrorHandler
    If vbProj Is Nothing Then
        MsgBox "Enable 'Trust access to the VBA project object model' in Macro Settings.", vbExclamation
        Exit Sub
    End If

    ' === Delete existing module if present ===
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(moduleName)
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
    On Error GoTo ErrorHandler

    ' === Add a new standard module and inject the code ===
    Set vbComp = vbProj.VBComponents.Add(1)    ' 1 = standard module
    vbComp.Name = moduleName
    vbComp.CodeModule.AddFromString vbaCode

    ' === Clean up any trailing "()" line in the injected module ===
    With vbComp.CodeModule
        Dim totalLines As Long
        totalLines = .CountOfLines
        If totalLines > 0 Then
            If Trim$(.Lines(totalLines, 1)) = "()" Then
                .DeleteLines totalLines, 1
            End If
        End If
    End With

    ' === Run the injected macro ===
    Application.Run moduleName & "." & subroutineName

    Exit Sub

ErrorHandler:
    If fileOpened And Not ts Is Nothing Then ts.Close
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
