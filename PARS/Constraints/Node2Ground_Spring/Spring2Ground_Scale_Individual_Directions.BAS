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
    Dim stiff_coeff As Variant

    rc = pset.Select(FT_PROP, True, "Select Spring propeties to scale:")
    Do While rc = -1
		If (rc <> -1) Or (pset.Count() < 1) Then
			GoTo Cleanup
		End If

		rc = get_stiffness_coefficients(App, stiff_coeff)
    	If rc = FE_CANCEL Then GoTo Cleanup

		rc = pset.Reset()
		pID = pset.Next()
		Do While pID <> FE_FAIL
			rc = p.Get(pID)
			If (p.type = 41) And (p.cbush = 1) Then
				For i = 0 To 5 Step 1
					p.pval(i+36) = p.pval(i + 36) * stiff_coeff(i)
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

Function get_stiffness_coefficients(ByRef App As femap.model, ByRef stiffness_coefficients As Variant) As Long
	Dim rc As Long
	Dim vecBase1 As Variant, vecBase2 As Variant

	On Error GoTo Cleanup

	App.Info_DefaultX = 1.0
	App.Info_DefaultY = 1.0
	App.Info_DefaultZ = 1.0
	rc = App.feCoordPickByMethod(FCD_COORDINATES, "Input Translational Stiffness Scale:", vecBase1)
	If rc = FE_CANCEL Then GoTo Cleanup

	App.Info_DefaultX = 1.0
	App.Info_DefaultY = 1.0
	App.Info_DefaultZ = 1.0
	rc = App.feCoordPickByMethod(FCD_COORDINATES, "Input Rotational Stiffness Scale:", vecBase2)
	If rc = FE_CANCEL Then GoTo Cleanup

	stiffness_coefficients = Array(vecBase1(0), vecBase1(1), vecBase1(2), vecBase2(0), vecBase2(1), vecBase2(2))

Cleanup:
	get_stiffness_coefficients = rc
End Function
