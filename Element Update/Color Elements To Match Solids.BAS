Rem File: ColorElementsToMatchSolids.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim surfSet As femap.Set
	Set surfSet = App.feSet

	Dim feSolid As femap.Solid
	Set feSolid = App.feSolid

	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface

	While feSolid.Next
		rc = elSet.Clear()
		rc = surfSet.Clear()

		rc = surfSet.AddRule( feSolid.ID, FGD_SURFACE_ONSOLID )

		rc = surfSet.Reset
		surfID = surfSet.First()

		rc = feSurf.Get( surfID )

		rc = elSet.AddRule( feSolid.ID, FGD_ELEM_ATSOLID )

		If elSet.count() > 0 Then
			rc = App.feModifyColor( FT_ELEM, elSet.ID, feSurf.color, 0)

		End If

	Wend

	rc = App.feViewRegenerate( 0 )

	
End Sub
