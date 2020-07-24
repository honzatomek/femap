Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup


Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub
