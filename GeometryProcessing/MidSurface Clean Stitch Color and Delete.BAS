Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim soSet As femap.Set
	Set soSet = App.feSet

	Dim solidID As Long

	Dim surfSet1 As femap.Set
	Dim surfSet2 As femap.Set
	Set surfSet1 = App.feSet
	Set surfSet2 = App.feSet
	Dim surfID As Long
	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface
	Dim surfColor As Long

	rc = soSet.SelectID( FT_SOLID, "Select Solid to Mid-Surface", solidID )

	If solidID > 0 Then
		rc = surfSet1.AddAll( FT_SURFACE )
		rc = surfSet2.AddRule( solidID, FGD_SURFACE_ONSOLID )

		surfID = surfSet2.First
		rc = feSurf.Get( surfID )
		surfColor = feSurf.color

		rc = App.feSurfaceMidAuto( surfSet2.ID, .25)

		rc  = surfSet2.Clear()
		rc = surfSet2.AddAll( FT_SURFACE )
		rc = surfSet2.RemoveSet( surfSet1.ID )

		Dim newSolidID As Long
		newSolidID = App.Info_NextID( FT_SOLID)
		rc = App.feSolidStitch( surfSet2.ID, 1E-6)

		rc = App.feSolidCleanup( -newSolidID, True, False, True )

		rc = surfSet1.Clear
		rc = surfSet1.AddRule( newSolidID, FGD_SURFACE_ONSOLID )
		If surfSet1.count > 0 Then
			rc = App.feModifyColor( FT_SURFACE, surfSet1.ID, surfColor, False )
		End If
		rc = App.feDelete( FT_SOLID, -solidID )

	End If

	rc = App.feViewRegenerate( 0 )

End Sub
