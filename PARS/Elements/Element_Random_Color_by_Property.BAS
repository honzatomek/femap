Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    Dim pset As Object
    Dim p As Object

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

	Set pset = App.feSet()
	Set p = App.feProp()
	rc = pset.AddAll(FT_PROP)
	If (rc <> -1) Or (pset.Count() < 2) Then GoTo Cleanup
	rc = pset.Reset()
	For i = 0 To pset.Count() - 1 Step 1
	 	rc = p.Get(pset.Next())
	 	If rc <> -1 Then GoTo Cleanup
	 	p.color = GetRandom(149, 0)
		rc = p.Put(p.ID)
	Next

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub

Private Function GetRandom(Optional ByVal upperbound As Long = 256 * 256 * 256 - 1, Optional ByVal lowerbound As Long = 0) As Long
	'GetRandom = CLng((upperbound - lowerbound + 1) * Rnd(-CLng(Split(CStr(Now()), ":")(2))) + lowerbound)
	GetRandom = CLng((upperbound - lowerbound + 1) * Rnd() + lowerbound)
End Function
