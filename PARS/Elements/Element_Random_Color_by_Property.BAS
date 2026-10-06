' =============================================================================
' Description: Assigns a distinct random color to elements according to their property.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' N4W/FEMAP portable constants (do not rely on type-library enum names)
Const C_FT_PROP As Long = 11
Const C_MSG_NORMAL As Long = 0

Const DBUG As Boolean = True
Private er As Long

Sub Main()
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, i As Long

    Dim pset As Object
    Dim p As Object

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Set pset = App.feSet()
    Set p = App.feProp()
    rc = pset.AddAll(C_FT_PROP)
    If (rc <> -1) Or (pset.Count() < 2) Then GoTo Cleanup
    rc = pset.Reset()
    For i = 0 To pset.Count() - 1 Step 1
        rc = p.Get(pset.Next())
        If rc <> -1 Then GoTo Cleanup
        p.color = GetRandom(149, 0)
        rc = p.Put(p.ID)
    Next

Cleanup:
    Call App.feAppMessage(C_MSG_NORMAL, "The script exited with code: " & er)
    On Error Resume Next
    Set p = Nothing
    Set pset = Nothing
    Set App = Nothing
End Sub

Private Function GetRandom(ByVal upperbound As Long, ByVal lowerbound As Long) As Long
    GetRandom = CLng((upperbound - lowerbound + 1) * Rnd() + lowerbound)
End Function
