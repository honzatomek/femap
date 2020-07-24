Rem File: GroupMatchingThickness.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feProp As femap.Prop
	Set feProp = App.feProp

	Dim propSet As femap.Set
	Set propSet = App.feSet

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim propID As Long
	Dim thickness As Double

	Dim feGroup As femap.Group
	Set feGroup = App.feGroup

	SelectProperty:
	rc = propSet.SelectID ( FT_PROP, "Select Base Element Property", propID )

	If propID > 0 Then

		rc = feProp.Get( propID )
		If feProp.type <17 Or feProp.type > 20 Then
			rc = App.feAppMessageBox ( 1, "Selected Property Must Be of Type Plate or Plane Strain" )
			If rc = 2 Then
				GoTo EndMacro
			Else
				GoTo SelectProperty
			End If
		End If
		thickness = feProp.pval(0)

		feProp.Reset
		While feProp.Next
			If feProp.type >=17 Or feProp.type <= 20 Then
				If Abs( feProp.pval(0) - thickness ) < 1E-16 Then
					rc = propSet.Add( feProp.ID )
				End If
			End If
		Wend

		rc = elSet.AddSetRule( propSet.ID, FGD_ELEM_BYPROP )

		rc = feGroup.SetAdd( FT_PROP, propSet.ID )
		rc = feGroup.SetAdd( FT_ELEM, elSet.ID )

		feGroup.title = "Planar Elements with Thickness = " + Str$(Round(thickness,4))
		rc = feGroup.Put( feGroup.NextEmptyID )

	End If

	EndMacro:
	App.feViewRegenerate(0)
	
End Sub
