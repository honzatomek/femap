' =============================================================================
' Description: Reads an F06 file and helps identify/report nodes referenced by solver error messages.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

' Show_Error_Nodes.BAS
' Unified FEMAP / Nastran for Windows 2004 compatible version.
'
' Selects one *.f06 file, then offers all *.f06 files from the same
' directory in a list. The selected file is printed to the BASIC debug
' output exactly as the original script intended.
'
' Compatibility rules used here:
'   - late-bound FEMAP object + feGetObject()
'   - no VBA Array(), Split(), Join(), or InStrRev()
'   - no external type-library references
'   - only one-dimensional dynamic primitive arrays

Dim fls$()
Dim selectedFileIndex As Long
Const filetype = "*.f06"

Sub Main()
    Dim filename As String

    On Error GoTo ErrorHandler

    filename = GetFilename()
    If filename = "" Then Exit Sub

    Call PrintFileText(filename)
    Exit Sub

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Sub PrintFileText(ByVal filename As String)
    Dim App As Object
    Dim f As Object
    Dim rc As Long

    Set App = feGetObject()
    Set f = App.feRead

    rc = f.Open(filename, 100)
    If rc <> -1 Then
        MsgBox "Could not open file: " & filename
        Set f = Nothing
        Set App = Nothing
        Exit Sub
    End If

    Do While Not f.AtEOF()
        rc = f.Read()
        If rc <> -1 Then Exit Do
        Debug.Print f.Line
    Loop

    rc = f.Close()
    Set f = Nothing
    Set App = Nothing
End Sub

Function GetFilename() As String
    Dim App As Object
    Dim rc As Long
    Dim fName As String
    Dim fldr As String
    Dim a As Long

    Set App = feGetObject()
    GetFilename = ""
    selectedFileIndex = -1

    rc = App.feFileGetName("Select a File - will look at all in the same directory", "Select Directory", filetype, True, fName)
    If rc <> -1 Then
        Set App = Nothing
        Exit Function
    End If

    fldr = FolderFromPath(fName)
    If fldr = "" Then
        Set App = Nothing
        Exit Function
    End If

    ChDir fldr
    fName = Dir$(filetype)
    a = 0

    While fName <> ""
        ReDim Preserve fls(a)
        fls(a) = fName
        fName = Dir$()
        a = a + 1
    Wend

    If a = 0 Then
        MsgBox "No *.f06 files found in the selected directory."
        Set App = Nothing
        Exit Function
    End If

    Begin Dialog UserDialog 100,100,550,378,"Show erroneous nodes from *.f06 file:",.DlgFunc ' %GRID:10,7,1,1
        OKButton 170,350,80,21
        CancelButton 280,350,90,21
        ListBox 10,7,530,336,fls(),.List
    End Dialog
    Dim dlg As UserDialog

    If Dialog(dlg) = 0 Then
        Set App = Nothing
        Exit Function
    End If

    If selectedFileIndex < 0 Then
        selectedFileIndex = dlg.List
    End If

    If selectedFileIndex < 0 Or selectedFileIndex >= a Then
        Set App = Nothing
        Exit Function
    End If

    GetFilename = fldr & fls(selectedFileIndex)
    Set App = Nothing
End Function

Function FolderFromPath(ByVal fullPath As String) As String
    Dim i As Long
    Dim lastSlash As Long

    lastSlash = 0
    For i = 1 To Len(fullPath)
        If Mid$(fullPath, i, 1) = "\" Then lastSlash = i
    Next

    If lastSlash = 0 Then
        FolderFromPath = ""
    Else
        FolderFromPath = Left$(fullPath, lastSlash)
    End If
End Function

' See DialogFunc help topic for more information.
Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
    DlgFunc = True
    Select Case Action%
    Case 1 ' Dialog box initialization
    Case 2 ' Value changing or button pressed
        Select Case DlgItem$
        Case "OK"
            selectedFileIndex = DlgValue("List")
            DlgFunc = False
        Case "Cancel", "Esc"
            selectedFileIndex = -1
            DlgFunc = False
        End Select
    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function
