Rem File: ColorElementsToMatchSurfaces.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface

	Dim elSet As femap.Set
	Set elSet = App.feSet

	While feSurf.Next()
		rc = elSet.Clear
		rc = elSet.AddRule( feSurf.ID, FGD_ELEM_ATSURFACE )
		If elSet.count() > 0 Then
			rc = App.feModifyColor( FT_ELEM, elSet.ID, feSurf.color, 0)
			End If
	Wend

	rc = App.feViewRegenerate( 0 )


	
End Sub
