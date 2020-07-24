Option Explicit On

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long
    Dim c As Long
    c = FT_BC_DIR

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim cset As Object
    Set cset = App.feSet()

    rc = cset.AddAll(c)

    rc = cset.PrevBefore(App.Info_ActiveID(c))
    If rc = FE_FAIL Then
    	rc = cset.Last()
    End If

    App.Info_ActiveID(c) = rc
    Call App.feViewRedraw(0)

Cleanup:
	If DBUG Then Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set cset = Nothing
	Set App = Nothing
End Sub
