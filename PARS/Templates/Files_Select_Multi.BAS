' =============================================================================
' Description: Template for selecting multiple files with the FEMAP/WinWrap dialog interface.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit On

Dim m As Integer
Dim msg() As String
Dim cmbfiles$()
Dim selected As Variant
Dim selectedCount As Long

Sub Main
    Dim App As Object
    Dim rc As Long
    Dim fName As String
    Dim fldr As String
    Dim a As Long
    Dim i As Long
    Dim p As Long

    Set App = feGetObject()

    rc = App.feFileGetName("Select a File - will look at all in the same directory", "Select Directory", "*.cmb", True, fName)
    If rc = 0 Then GoTo Cleanup

    p = LastBackslash(fName)
    If p <= 0 Then
        MsgBox "Could not determine the selected file directory."
        GoTo Cleanup
    End If
    fldr = Left$(fName, p)

    ChDir fldr
    fName = Dir$("*.cmb")
    a = 0

    While fName <> ""
        ReDim Preserve cmbfiles(a)
        cmbfiles(a) = fName
        fName = Dir$()
        a = a + 1
    Wend

    If a = 0 Then
        MsgBox "No .cmb files were found in the selected directory."
        GoTo Cleanup
    End If

    Begin Dialog UserDialog 100,100,550,378,"Import .cmb Results:",.DlgFunc ' %GRID:10,7,1,1
        MultiListBox 10,49,520,301,cmbfiles(),.List
        OKButton 170,350,80,21
        CancelButton 280,350,90,21
        CheckBox 10,21,90,14,"Select All",.CheckBox1
    End Dialog
    Dim dlg As UserDialog
    If Dialog(dlg) = 0 Then GoTo Cleanup

    If dlg.CheckBox1 = 1 Then
        ReDim msg(a - 1)
        For i = 0 To a - 1
            msg(i) = cmbfiles(i)
        Next
    Else
        If selectedCount <= 0 Then
            MsgBox "No files selected."
            GoTo Cleanup
        End If
    End If

    For i = 0 To UBound(msg)
        Debug.Print msg(i)
    Next

Cleanup:
    Set App = Nothing
End Sub

Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
    DlgFunc = True
    Select Case Action%
    Case 2
        Select Case DlgItem$
        Case "OK"
            selectedCount = 0
            selected = DlgValue("List")
            If IsArray(selected) Then
                ReDim msg(UBound(selected))
                For m = 0 To UBound(selected)
                    msg(m) = cmbfiles(CLng(selected(m)))
                    selectedCount = selectedCount + 1
                Next
            End If
            DlgFunc = False
        Case "Cancel", "Esc"
            selectedCount = 0
            DlgFunc = False
        End Select
    End Select
End Function

Private Function LastBackslash(ByVal text As String) As Long
    Dim startPos As Long
    Dim foundPos As Long
    Dim lastPos As Long

    startPos = 1
    lastPos = 0
    Do
        foundPos = InStr(startPos, text, "\")
        If foundPos = 0 Then Exit Do
        lastPos = foundPos
        startPos = foundPos + 1
    Loop
    LastBackslash = lastPos
End Function
