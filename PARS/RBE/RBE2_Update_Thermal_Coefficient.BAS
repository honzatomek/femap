Attribute VB_Name = "Module1"
Option Explicit

Sub Main
	Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long

    On Error GoTo ExitHere

    Dim elset As Object
    Set elset = App.feSet()
    Dim el As Object
    Set el = App.feElem()
    Dim elID As Long
    Dim tmp As Variant
    Dim thexp As Double

	rc = elset.Select(8, True, "Select Rigid elements (RBE2)")
	If rc <> -1 Then GoTo ExitHere
	rc = elset.Reset()

	tmp = InputBox("Input new Thermal Expansion Coefficient", "RBE2 Thermal Update", "1.2e-05")
	If tmp = "" Then GoTo ExitHere

	If InStr(1, tmp, ".") <> 0 Then tmp = Replace(tmp, ".", ",")
	thexp = CDbl(tmp)

	elID = elset.Next()

	Do While elID <> 0
		rc = el.Get(elID)
		If (el.type = 29) And (el.topology = 13) Then 'And (Not el.RigidInterpolate)
			'Debug.Print el.ID
			el.RigidThermalExpansion = thexp
			rc = el.Put(el.ID)
		End If

		elID = elset.Next()
	Loop

ExitHere:
	On Error Resume Next
	Set el = Nothing
	Set elset = Nothing
	Set App = Nothing
End Sub
