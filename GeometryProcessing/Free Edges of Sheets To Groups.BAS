Rem File: FreeEdgesofSheetsToGroups.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feCurve As femap.Curve
	Set feCurve = App.feCurve
	Dim cuSet As femap.Set
	Set cuSet = App.feSet
	Dim  FreeEdgeSet As femap.Set
	Set FreeEdgeSet = App.feSet
	Dim feGroup As femap.Group
	Set feGroup = App.feGroup

	rc = feCurve.Reset

	While feCurve.Next()
		rc = cuSet.Clear()
		rc = cuSet.AddRule( feCurve.ID, FGD_SURFACE_BYCURVE )

		If cuSet.count = 1 Then
			FreeEdgeSet.Add( feCurve.ID )
		End If
	Wend

	If FreeEdgeSet.count > 0 Then
		rc = feGroup.SetAdd( FT_CURVE, FreeEdgeSet.ID )
		feGroup.title = "Free Curves"
		rc = feGroup.Put( feGroup.NextEmptyID )
		rc = App.feViewShow( FT_CURVE, FreeEdgeSet.ID )
	End If

End Sub
