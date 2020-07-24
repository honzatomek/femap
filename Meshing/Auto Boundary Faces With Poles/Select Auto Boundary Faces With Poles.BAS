Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim surfSet As femap.Set
	Set surfSet = App.feSet

	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface

	Dim surfID As Long

	rc = surfSet.Select( FT_SURFACE, True, "Select Surfaces" )

	If surfSet.count > 0 Then
		surfID = surfSet.First()
		While surfID > 0
			rc = feSurf.Get( surfID )
				If feSurf.HasPole() = FE_OK Then
					rc = App.feBoundaryFromSurfaces( -feSurf.ID )
				End If
			surfID = surfSet.Next()
		Wend
	End If

End Sub

