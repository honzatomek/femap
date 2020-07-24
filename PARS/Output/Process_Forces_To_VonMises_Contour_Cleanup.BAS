Attribute VB_Name = "Module2"
Option Explicit

Sub Main
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long
	Dim o As Object
	Set o = App.feOutput()

	On Error GoTo ExitHere

	Dim s As String
	Dim tmp As Long

	o.setID = 24
	rc = o.Reset()
	s = "End A Max VonMises Stress"
	Do While o.Next() <> FE_FAIL
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
