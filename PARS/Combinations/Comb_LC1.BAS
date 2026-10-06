' =============================================================================
' Description: Reads a text definition and creates output-set combinations from existing FEMAP results.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit On

Const FE_OK As Long = -1
Sub Main()
    Dim App As Object
    Dim rc As Long
    Dim fName As String
    Dim f As Object
    Dim os As Object
    Dim Snum As Long
    Dim Title As String
    Dim Notes As String
    Dim n As Long
    Dim fromSet(100) As Variant
    Dim factor(100) As Variant
    Dim dummy(100) As Variant
    Dim toSet(100) As Variant

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
                MsgBox "Combination contains more than 100 source output sets."
                GoTo Cleanup
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
            If rc <> FE_OK Then
                MsgBox "Could not create output set " & CStr(Snum) & "."
                GoTo Cleanup
            End If

            toSet(0) = Snum
            rc = App.feOutputProcess(n, True, 2, dummy, fromSet, dummy, toSet, dummy, dummy, factor, dummy, dummy, dummy)
            If rc <> FE_OK Then
                MsgBox "Could not create output combination " & CStr(Snum) & "."
                GoTo Cleanup
            End If
        End If

NextBlock:
        ' A blank separator line, if present, is simply consumed by the next loop.
    Loop
    GoTo Cleanup

ReadError:
    MsgBox "Error while reading the combination file."

Cleanup:
    On Error Resume Next
    rc = f.Close()
    Set os = Nothing
    Set f = Nothing
    Set App = Nothing
End Sub
