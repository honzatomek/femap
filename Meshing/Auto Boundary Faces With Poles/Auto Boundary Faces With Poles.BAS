Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface

	Dim rc As femap.zReturnCode

	While feSurf.Next()
		If feSurf.HasPole() = FE_OK Then
			rc = App.feBoundaryFromSurfaces( -feSurf.ID )
		End If
	Wend

End Sub

