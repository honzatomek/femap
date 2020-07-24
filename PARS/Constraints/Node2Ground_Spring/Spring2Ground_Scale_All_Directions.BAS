Option Explicit On

Const DBUG As Boolean = False
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
    Dim scale As Double

    scale = 1.0

    rc = pset.Select(FT_PROP, True, "Select Spring propeties to scale:")
    Do While rc = -1
		If (rc <> -1) Or (pset.Count() < 1) Then
			GoTo Cleanup
		End If

		rc = App.feGetReal("Input Stiffness Scale Value", -1.0e-16, 1.0e16, scale)
		If rc = FE_CANCEL Then GoTo Cleanup

		rc = pset.Reset()
		pID = pset.Next()
		Do While pID <> FE_FAIL
			rc = p.Get(pID)
			If (p.type = 41) And (p.cbush = 1) Then
				For i = 36 To 41 Step 1
					p.pval(i) = p.pval(i) * scale
				Next
				p.Put(pID)
			End If
			pID = pset.Next()
		Loop
		rc = pset.Select(FT_PROP, True, "Select Spring propeties to scale:")
	Loop

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set p = Nothing
	Set pset = Nothing
	Set App = Nothing
End Sub
