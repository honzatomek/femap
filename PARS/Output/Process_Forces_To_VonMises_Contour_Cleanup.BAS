' =============================================================================
' Description: Removes temporary/custom output vectors created by the force-to-von-Mises processing workflow.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_FAIL As Long = 0

Sub Main
	Dim App As Object
	Set App = feGetObject()
	Dim rc As Long, i As Long
	Dim o As Object
	Set o = App.feOutput()

	On Error GoTo ExitHere

	Dim s As String
	Dim tmp As Long

	o.setID = 24
	rc = o.Reset()
	s = "End A Max VonMises Stress"
	Do While o.Next() <> C_FE_FAIL
		If InStr(1, o.title, s) <> 0 Then
			tmp = o.ID()
			rc = o.Delete(tmp)
			rc = o.Delete(tmp + 1)
			rc = o.Reset()
		End If
	Loop

ExitHere:
	On Error Resume Next
	Set o = Nothing
	Set App = Nothing
	On Error GoTo 0
End Sub
