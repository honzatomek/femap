Option Explicit

'Shows active material

Sub Main
    Dim App As femap.model
    Set App = feFemap()
	Dim rc As Long, i As Long

	On Error GoTo ExitHere

	Dim maID As Long
	Dim elset As Object
	Set elset = App.feSet()

	maID = App.Info_ActiveID(FT_MATL)

	rc = elset.AddRule(maID, FGD_ELEM_BYMATL)

	If elset.Count() > 0 Then
		Call App.feViewShow(FT_ELEM, elset.ID)
	End If


ExitHere:
	On Error Resume Next
		Set elset = Nothing
		Set App = Nothing
End Sub
