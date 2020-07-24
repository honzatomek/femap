Option Explicit On

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim elset As Object
    Set elset = App.feSet()

    Dim el As Object
    Set el = App.feElem()
    Dim length As Double
    Dim totlen As Double

    Dim load As Double
    Dim ld As Object
    Set ld = App.feLoadMesh()

    Dim DOF(2) As Long
    DOF(0) =6
    DOF(1) = 0
    DOF(2) = 0

	Dim vals(5) As Double
	For i = 0 To 5 Step 1
		vals(i) = 0
	Next
	Dim func(4) As Long
	For i = 0 To 4 Step 1
		func(i) = 0
	Next

    rc = elset.Select(8, True, "Select elements to load:")
	If (rc <> -1) Or (elset.Count() < 1) Then
		er = 1
		GoTo Cleanup
	End If

	Dim eq As String, eq2 As Double
	Do
		eq = InputBox("Input the summed load:", "Load by summed force")
		If eq = "" Then GoTo Cleanup
		If InStr(1, ".", ",") <> 0 Then eq = Replace(eq, ".", ",")
		If IsNumeric(eq) Then Exit Do
	Loop
	eq2 = CDbl(eq)

	totlen = 0
	rc = elset.Reset()
	Do While elset.Next() <> 0
		rc = el.Get(elset.CurrentID())
		rc = el.Length(length)
		totlen = totlen + length
	Loop

	eq2 = eq2/totlen

	rc = elset.Reset()
	Do While elset.Next() <> 0
		rc = el.Get(elset.CurrentID())
		rc = el.Length(length)
		'load = eq2 * length
		load = eq2
		vals(0) = load
		vals(1) = load
		rc = ld.Add(-el.ID, 41, 0, DOF, vals, func)
	Loop

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set ld = Nothing
	Set el = Nothing
	Set elset = Nothing
	Set App = Nothing
End Sub
