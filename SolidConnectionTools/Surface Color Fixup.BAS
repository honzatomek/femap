Rem File: SurfaceColorFixup.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feSolid As femap.Solid
	Dim surfSet As femap.Set
	Dim solidSet As femap.Set
	Dim solidColor As Long
	Dim feSurf As femap.Surface
	Dim feSurf2 As femap.Surface

	Set feSolid = App.feSolid()
	Set surfSet = App.feSet()
	Set feSurf = App.feSurface()
	Set feSurf2 = App.feSurface()
	Set solidSet = App.feSet()

	rc = feSolid.Reset

	surfColor = App.Info_Color(FT_SURFACE)

	While feSolid.Next
		'Get all the surfaces on this solid and reset their color
		solidColor = 0
		rc = surfSet.Clear
		rc = surfSet.AddRule( feSolid.ID, FGD_SURFACE_ONSOLID )
		rc = surfSet.Reset
		ID = surfSet.Next
		While ID <> 0
			rc = feSurf.Get(ID)
			If feSurf.color <> App.Info_Color( FT_SURFACE ) Then
				If feSurf.color <> FCL_BLUE Then
					solidColor = feSurf.color
				End If
			End If
			ID = surfSet.Next
		Wend
		If solidColor = 0 Then
			solidColor = App.Info_Color(FT_SURFACE)
		End If
		ID = surfSet.Reset
		ID = surfSet.Next
		While ID <> 0
			rc = feSurf.Get(ID)
			feSurf.color = solidColor
			rc = feSurf.Put(ID)
			ID = surfSet.Next
		Wend
	Wend

	rc = feSurf.Reset

	While feSurf.Next
		If feSurf.approach = 6 Then
			feSurf.color =FCL_BLUE
			rc = feSurf.Put(feSurf.ID)
			rc = feSurf2.Get(feSurf.approachData(0) )
			feSurf2.color = FCL_BLUE
			rc = feSurf2.Put(feSurf2.ID)
		End If
	Wend

rc = App.feViewRegenerate(0)


	
End Sub
