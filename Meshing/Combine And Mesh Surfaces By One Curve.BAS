Rem File: CombineAndMeshSurfacesByOneCurve.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim curveSet As femap.Set
	Set curveSet = App.feSet()

	Dim surfSet As femap.Set
	Set surfSet = App.feSet

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim surfID As Long

	rc = curveSet.Select( FT_CURVE, True, "Select Curve to Remove")
	'rc = surfSet.Select( FT_SURFACE, True, "Select Surfaces to Combine")
	rc = surfSet.AddSetRule( curveSet.ID, FGD_SURFACE_BYCURVE )

	rc = elSet.AddSetRule( surfSet.ID, FGD_ELEM_ATSURFACE )
	rc = ndSet.AddSetRule( elSet.ID, FGD_NODE_ONELEM )

	surfID = App.Info_NextID( femap.FT_SURFACE )

	rc = App.feBoundaryFromSurfaces( surfSet.ID )

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	Dim propID As Long
	Dim elemID As Long

	elemID = elSet.First()

	rc = feElem.Get( elemID )

	propID = feElem.propID

	rc = App.feDelete( FT_ELEM, elSet.ID )
	rc = App.feDelete( FT_NODE, ndSet.ID )

	rc = App.feMeshSurface( -surfID, propID, 1, 2.0, 200, 1, 60.0, False, 0.0, True, 10, .001 )


	rc = App.feViewRegenerate( 0 )








	
End Sub
