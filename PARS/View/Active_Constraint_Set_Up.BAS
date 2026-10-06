' =============================================================================
' Description: Activates the next FEMAP constraint set.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const DBUG As Boolean = False
Const C_FT_BC_DIR As Long = 17
Const C_FE_FAIL As Long = 0
Const C_FCM_NORMAL As Long = 0
Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long
    Dim c As Long
    c = C_FT_BC_DIR

    er = 0
    If Not DBUG Then On Error GoTo Cleanup

    Dim cset As Object
    Set cset = App.feSet()

    rc = cset.AddAll(c)
    If cset.Count() <= 0 Then GoTo Cleanup

    rc = cset.NextAfter(App.Info_ActiveID(c))
    If rc = C_FE_FAIL Then rc = cset.First()
    If rc <> C_FE_FAIL Then App.Info_ActiveID(c) = rc
    Call App.feViewRedraw(0)

Cleanup:
    If DBUG Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    On Error Resume Next
    Set cset = Nothing
    Set App = Nothing
End Sub
