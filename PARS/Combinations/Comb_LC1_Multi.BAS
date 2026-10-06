' =============================================================================
' Description: Imports multiple result files and creates output combinations/envelopes across them.
' Compatibility: Modern FEMAP only; the script shows an early compatibility message on N4W 2004.
' =============================================================================

Option Explicit On

Const FE_OK As Long = -1
Const FCM_NORMAL As Long = 0
Const FCM_WARNING As Long = 2
Const FT_OUT_CASE As Long = 28

Dim gFiles$()
Dim gSelected() As String
Dim gSelection As Variant
Dim gSelectedCount As Long
Dim gIndex As Integer

Sub Main()
    Dim App As Object
    Dim rc As Long
    Dim probeSet As Object
    Dim fName As String
    Dim folder As String
    Dim p As Long
    Dim fileCount As Long
    Dim i As Long
    Dim j As Long
    Dim dlg As UserDialog
    Dim doEnvelope(0 To 2) As Boolean
    Dim created As Object

    Set App = feGetObject()

    If Not SupportsEnvelopeFromSets(App) Then
        MsgBox "This script uses feOutputProcessEnvelopeFromSets and is not portable to Nastran for Windows 2004. Use Comb_LC1.BAS to import combinations without automatic envelopes."
        GoTo Cleanup
    End If

    rc = App.feFileGetName("Select a file in the directory containing combinations", "Linear Comb. File", "*.cmb", True, fName)
    If rc = 0 Then GoTo Cleanup

    p = LastBackslash(fName)
    If p <= 0 Then
        MsgBox "Could not determine the selected file directory."
        GoTo Cleanup
    End If
    folder = Left$(fName, p)
    ChDir folder

    fName = Dir$("*.cmb")
    fileCount = 0
    While fName <> ""
        ReDim Preserve gFiles(fileCount)
        gFiles(fileCount) = fName
        fileCount = fileCount + 1
        fName = Dir$()
    Wend
    If fileCount = 0 Then
        MsgBox "No .cmb files were found in the selected directory."
        GoTo Cleanup
    End If

    Begin Dialog UserDialog 100,100,550,378,"Import .cmb Results:",.DlgFunc ' %GRID:10,7,1,1
        MultiListBox 10,49,520,301,gFiles(),.List
        OKButton 170,350,80,21
        CancelButton 280,350,90,21
        CheckBox 10,21,90,14,"Select All",.CheckBox1
        GroupBox 120,7,410,35,"Process Combinations",.GroupBox1
        CheckBox 140,21,80,14,"Max",.CheckBoxMax
        CheckBox 250,21,80,14,"Min",.CheckBoxMin
        CheckBox 350,21,80,14,"MaxAbs",.CheckBoxMaxAbs
    End Dialog

    If Dialog(dlg) = 0 Then GoTo Cleanup

    If dlg.CheckBox1 = 1 Then
        ReDim gSelected(fileCount - 1)
        For i = 0 To fileCount - 1
            gSelected(i) = gFiles(i)
        Next
        gSelectedCount = fileCount
    ElseIf gSelectedCount <= 0 Then
        MsgBox "No files selected."
        GoTo Cleanup
    End If

    doEnvelope(0) = dlg.CheckBoxMax
    doEnvelope(1) = dlg.CheckBoxMin
    doEnvelope(2) = dlg.CheckBoxMaxAbs

    Set created = App.feSet

    For i = 0 To gSelectedCount - 1
        rc = created.Clear()
        Call App.feAppMessage(FCM_WARNING, "Opening file: " & gSelected(i))
        rc = ImportOutputCombinations(App, gSelected(i), created)
        If rc <> FE_OK Then GoTo Cleanup

        For j = 0 To 2
            If doEnvelope(j) And created.Count > 0 Then
                Call App.feAppMessage(FCM_NORMAL, "Creating envelope type " & CStr(j))
                rc = App.feOutputProcessEnvelopeFromSets(True, j, 2, False, 2, True, created.ID, 0, 0)
                If rc <> FE_OK Then
                    MsgBox "Could not create envelope for " & gSelected(i) & "."
                    GoTo Cleanup
                End If
            End If
        Next

        rc = created.Clear()
    Next

Cleanup:
    Set created = Nothing
    Set App = Nothing
End Sub

Function ImportOutputCombinations(ByRef App As Object, ByVal fName As String, ByRef created As Object) As Long
    Dim f As Object
    Dim os As Object
    Dim rc As Long
    Dim Snum As Long
    Dim Title As String
    Dim Notes As String
    Dim n As Long
    Dim fromSet(100) As Variant
    Dim factor(100) As Variant
    Dim dummy(100) As Variant
    Dim toSet(100) As Variant

    ImportOutputCombinations = 0
    Set f = App.feRead
    rc = f.Open(fName, 100)
    If rc <> FE_OK Then
        MsgBox "Could not open " & fName & "."
        GoTo Done
    End If
    rc = f.SetFreeFormat()

    Do While Not f.AtEOF()
        rc = f.Read()
        If rc <> FE_OK Then Exit Do
        If Trim$(f.Line) = "" Then GoTo NextBlock
        Snum = f.IntField(2, 0)

        rc = f.Read()
        If rc <> FE_OK Then GoTo ReadError
        Title = f.Line
        rc = f.Read()
        If rc <> FE_OK Then GoTo ReadError
        Notes = f.Line

        n = 0
        Do
            rc = f.Read()
            If rc <> FE_OK Then GoTo ReadError
            fromSet(n) = f.IntField(1, -999)
            factor(n) = f.RealField(2, 0#)
            If fromSet(n) = -999 Then Exit Do
            n = n + 1
            If n > 100 Then
                MsgBox "Combination contains more than 100 source output sets."
                GoTo Done
            End If
        Loop

        If n > 0 Then
            Set os = App.feOutputSet
            rc = os.Delete(Snum)
            os.Title = Title
            os.Notes = Notes
            os.Program = 4
            os.Analysis = 1
            os.Value = 0#
            rc = os.Put(Snum)
            If rc <> FE_OK Then GoTo ReadError

            toSet(0) = Snum
            rc = App.feOutputProcess(n, True, 2, dummy, fromSet, dummy, toSet, dummy, dummy, factor, dummy, dummy, dummy)
            If rc <> FE_OK Then GoTo ReadError
            rc = created.Add(Snum)
        End If
NextBlock:
    Loop

    ImportOutputCombinations = FE_OK
    GoTo Done
ReadError:
    MsgBox "Error while reading or processing " & fName & "."
Done:
    On Error Resume Next
    rc = f.Close()
    Set os = Nothing
    Set f = Nothing
End Function

Function SupportsEnvelopeFromSets(ByRef App As Object) As Boolean
    Dim s As Object
    Dim rc As Long

    SupportsEnvelopeFromSets = False
    If App.Info_Version < 10# Then Exit Function

    On Error Resume Next
    Set s = App.feSet
    rc = App.feOutputProcessEnvelopeFromSets(True, 0, 2, False, 2, False, s.ID, 0, 0)
    If Err.Number <> 0 Then
        Err.Clear
        SupportsEnvelopeFromSets = False
    Else
        SupportsEnvelopeFromSets = True
    End If
    On Error GoTo 0
    Set s = Nothing
End Function

Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
    Case 2
        Select Case DlgItem$
        Case "OK"
            gSelectedCount = 0
            gSelection = DlgValue("List")
            If IsArray(gSelection) Then
                ReDim gSelected(UBound(gSelection))
                For gIndex = 0 To UBound(gSelection)
                    gSelected(gIndex) = gFiles(Val(CStr(gSelection(gIndex))))
                    gSelectedCount = gSelectedCount + 1
                Next
            End If
        Case "Cancel", "Esc"
            gSelectedCount = 0
            DlgFunc = False
        End Select
    End Select
End Function

Function LastBackslash(ByVal text As String) As Long
    Dim startPos As Long
    Dim foundPos As Long
    Dim lastPos As Long
    startPos = 1
    lastPos = 0
    Do
        foundPos = InStr(startPos, text, "\\")
        If foundPos = 0 Then Exit Do
        lastPos = foundPos
        startPos = foundPos + 1
    Loop
    LastBackslash = lastPos
End Function
