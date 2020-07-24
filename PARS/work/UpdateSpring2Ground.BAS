Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim pset As Object
    Set pset = App.feSet()
    Dim pID As Long
    Dim p As Object
    Set p = App.feProp()
    Dim tmp As Double

    rc = pset.Select(FT_PROP, True, "Select Spring propeties to update:")
	If (rc <> -1) Or (pset.Count() < 1) Then
		GoTo Cleanup
	End If

	rc = pset.Reset()
	pID = pset.Next()
	Do While pID > 0
		rc = p.Get(pID)
		If p.cbush = 1 Then
			tmp = p.pval(36)
			p.pval(36) = p.pval(37) / 6
			p.pval(37) = tmp / 6
			p.pval(38) = p.pval(38) / 6
			p.Put(pID)
		End If
		' Debug.Print p.title & " " & p.cbush
		pID = pset.Next()
	Loop

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set pset = Nothing
	Set App = Nothing
End Sub
