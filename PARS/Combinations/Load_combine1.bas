' =============================================================================
' Description: Reads a text definition and creates combined FEMAP load sets with feLoadCombine.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit On

Const FE_OK As Long = -1

Sub Main()
    Dim App As Object
    Dim rc As Long
    Dim fName As String
    Dim f As Object
    Dim ls As Object
    Dim Snum As Long
    Dim Title As String
    Dim Notes As String
    Dim n As Long
    Dim i As Long
    Dim fromSet(100) As Long
    Dim factor(100) As Double

    Set App = feGetObject()

    rc = App.feFileGetName("Open LC combination file", "Linear Comb. File", "*.cmb", True, fName)
    If rc = 0 Then GoTo Cleanup

    Set f = App.feRead
    rc = f.Open(fName, 100)
    If rc <> FE_OK Then
        MsgBox "Could not open combination file."
        GoTo Cleanup
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
                MsgBox "Combination contains more than 100 source load sets."
                GoTo Cleanup
            End If
        Loop

        Set ls = App.feLoadSet
        rc = ls.Delete(Snum)
        ls.Title = Title
        ls.Notes = Notes
        rc = ls.Put(Snum)
        If rc <> FE_OK Then
            MsgBox "Could not create load set " & CStr(Snum) & "."
            GoTo Cleanup
        End If

        For i = 0 To n - 1
            rc = App.feLoadCombine(fromSet(i), Snum, factor(i))
            If rc <> FE_OK Then
                MsgBox "Could not combine load set " & CStr(fromSet(i)) & " into " & CStr(Snum) & "."
                GoTo Cleanup
            End If
        Next
NextBlock:
    Loop
    GoTo Cleanup

ReadError:
    MsgBox "Error while reading the combination file."

Cleanup:
    On Error Resume Next
    rc = f.Close()
    Set ls = Nothing
    Set f = Nothing
    Set App = Nothing
End Sub
