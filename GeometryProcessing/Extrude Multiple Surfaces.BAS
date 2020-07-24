Rem File: ExtrudeMultipleSurfaces.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feSurface As femap.Surface
	Set feSurface = App.feSurface
	Dim surfSet As femap.Set
	Set surfSet = App.feSet
	Dim surfID As Long

	Dim vecBase As Variant
	Dim vecDir As Variant
	Dim length As Double
	Dim vecDepth(3) As Double
	Dim vvecDepth As Variant

	rc = surfSet.Select( FT_SURFACE, True, "Select Surfaces to Extrude")

	If surfSet.count > 0 Then

		rc = App.feVectorPick( "Select Extrude Vector", False, length, vecBase, vecDir )

		vecDepth(0) = length
		vvecDepth = vecDepth

		rc = surfSet.Reset
		surfID = surfSet.First
		While surfID > 0

			rc = App.feSolidExtrude( 0, surfID, 3, 0, vecDir, vvecDepth )
			surfID = surfSet.Next
		Wend
	End If

rc = App.feViewRegenerate( 0)

End Sub
