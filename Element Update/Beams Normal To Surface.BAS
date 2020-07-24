Rem File: BeamsNormalToSurface.BAS

Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim nd1 As femap.Node
	Set nd1 = App.feNode

	Dim nd2 As femap.Node
	Set nd2 = App.feNode

	Dim midx As Double
	Dim midy As Double
	Dim midz As Double

	Dim surfSet As femap.Set
	Set surfSet = App.feSet

	Dim surfID As Long

	Dim oVec As Variant

	Dim elemID As Long

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	Dim midPoint(3) As Double
	Dim vmidPoint As Variant
	Dim surfNormal As Variant

	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface

	Dim u As Double
	Dim v As Double

	rc = elSet.Select( FT_ELEM, True, "Select Beams to Re-Orient" )
	rc = surfSet.SelectID( FT_SURFACE, "Select Surface", surfID )

	If surfID = 0 Then
		End
	End If

	rc = feSurf.Get( surfID )
	If elSet.count > 0 Then
		elemID = elSet.First()
		While elemID > 0
			rc = feElem.Get( elemID )
			' If the element is a beam or a bar
			If feElem.type = FET_L_BEAM Or FET_L_BAR Then
				'Get its two nodes
				rc = nd1.Get( feElem.Nodes(0))
				rc = nd2.Get( feElem.Nodes(1))
				'Find their average
				midx = (nd1.x + nd2.x) / 2.0
				midy = (nd1.y + nd2.y) / 2.0
				midz = (nd1.z + nd2.z) / 2.0
				'load the midpoint into a vector
				midPoint(0) = midx
				midPoint(1) = midy
				midPoint(2) = midz
				'hook it to a variant for passing to FEMAP
				vmidPoint = midPoint
				'project it onto the surface
				rc = feSurf.XYZToParam( vmidPoint, u, v )
				'find the normal at the projection
				rc = feSurf.normal( u, v, surfNormal )
				'update the element orientation vector
				rc = App.feModifyOrient( - elemID, 1, 0, surfNormal)
			End If
			elemID = elSet.Next
		Wend
	End If
	rc = App.feViewRegenerate( 0 )
End Sub
