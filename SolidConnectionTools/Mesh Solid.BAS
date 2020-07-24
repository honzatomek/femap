Rem File: MeshSolid.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feSet As femap.Set
	Set feSet = App.feSet()

	Dim Solid As femap.Solid
	Set Solid = App.feSolid
	Dim solidID As Long
	solidID = 0
	Dim isSolid As Boolean

	Dim msg As String

	SelectSolid:
	rc = feSet.SelectID (FT_SOLID, "Select Solid Volume to Mesh", solidID )
	If rc = 2 Or rc = 4 Then
		End
	End If
	Solid.Get ( solidID )
	isSolid = Solid.IsSolid
	Solid.Put ( solidID )

	If isSolid Then
		rc = App.feMeshTetSolid( FT_SOLID, -solidID, 0,0,True,1.0, True, True, False, 0.0, 0)
		App.feViewRegenerate ( 0 )
	Else
		msg = "Selected Solid is Not a Solid Volume" + _
			"Please Select a Solid Volume" + _
			"Press Cancel to End Macro"
		rc = App.feAppMessageBox ( 1, msg )
		If rc = 2 Then
			End
		End If
	End If

	GoTo SelectSolid

End Sub
