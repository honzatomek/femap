Option Explicit

'Shows active property

Sub Main
    Dim App As femap.model
    Set App = feFemap()
	Dim rc As Long, i As Long

	On Error GoTo ExitHere

	Dim pID As Long
	Dim elset As Object
	Set elset = App.feSet()

	pID = App.Info_ActiveID(FT_PROP)

	rc = elset.AddRule(pID, FGD_ELEM_BYPROP)

	If elset.Count() > 0 Then
		Call App.feViewShow(FT_ELEM, elset.ID)
	End If


ExitHere:
	On Error Resume Next
		Set elset = Nothing
		Set App = Nothing
End Sub
