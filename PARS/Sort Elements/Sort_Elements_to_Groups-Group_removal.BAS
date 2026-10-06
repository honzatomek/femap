' =============================================================================
' Description: Removes groups produced by the element-sorting workflow.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FT_GROUP As Long = 24

Sub Main()
    Dim App As Object
    Dim rc As Long
    Dim i As Long
    Dim s As Object
    Dim grp As Object
    Dim txt As Object
    Dim gid As Long

    Set App = feGetObject()
    Set s = App.feSet()
    Set grp = App.feGroup()
    Set txt = App.feText()

    rc = s.Select(C_FT_GROUP, True, "Select BEAM group(s) to delete:")
    If rc <> -1 Then GoTo ExitHere

    rc = s.Reset()
    For i = 1 To s.Count()
        gid = s.Next()
        If grp.Get(gid) = -1 Then
            If Left$(grp.title, 4) = "BEAM" Then
                rc = txt.Get(gid)
                If rc = -1 Then rc = txt.Delete(gid)
                rc = grp.Delete(gid)
            End If
        End If
    Next

ExitHere:
    On Error Resume Next
    Set txt = Nothing
    Set grp = Nothing
    Set s = Nothing
    Set App = Nothing
End Sub
