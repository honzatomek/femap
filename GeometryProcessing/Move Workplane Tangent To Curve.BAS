Rem File: MoveWorkplaneTangetToCurve.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim cuSet As femap.Set
	Set cuSet = App.feSet

	Dim cuID As Long

	Dim feCurve As femap.Curve
	Set feCurve = App.feCurve

	Dim ptID1 As Long
	Dim ptID2 As Long
	Dim pt1Loc(3) As Double
	Dim vpt1 As Variant

	Dim fePoint As femap.Point
	Set fePoint = App.fePoint

	'Line Axis
	Dim v1 As Variant
	'First Perp
	Dim v2 As Variant
	'Sencond Perp
	Dim v3 As Variant

	rc = cuSet.SelectID( FT_CURVE, "Select Curve to Align CSys", cuID )

	If cuID > 0 Then

		rc = feCurve.Get( cuID )
		rc = feCurve.EndPoints( ptID1, ptID2 )
		rc = fePoint.Get( ptID1 )
		pt1Loc(0) = fePoint.x
		pt1Loc(1) = fePoint.y
		pt1Loc(2) = fePoint.z

		vpt1 = pt1Loc

		rc = feCurve.tangent( False, vpt1, v1 )
		rc = App.feVectorPerpendicular( v1, v2 )
		rc = App.feVectorCrossProduct( v2, v1, v3 )

		App.Info_WorkplaneOrigin(0) = fePoint.x
		App.Info_WorkplaneOrigin(1) = fePoint.y
		App.Info_WorkplaneOrigin(2) = fePoint.z

		App.Info_WorkplaneX(0) = -v2(0)
		App.Info_WorkplaneX(1) = -v2(1)
		App.Info_WorkplaneX(2) = -v2(2)

		App.Info_WorkplaneY(0) = v3(0)
		App.Info_WorkplaneY(1) = v3(1)
		App.Info_WorkplaneY(2) = v3(2)

		App.Info_WorkplaneZ(0) = v1(0)
		App.Info_WorkplaneZ(1) = v1(1)
		App.Info_WorkplaneZ(2) = v1(2)

		'App.Info_WorkplaneX(0) = v2(0)
		'App.Info_WorkplaneX(1) = v3(0)
		'App.Info_WorkplaneX(2) = v1(0)

		'App.Info_WorkplaneY(0) = v2(1)
		'App.Info_WorkplaneY(1) = v3(1)
		'App.Info_WorkplaneY(2) = v1(1)

		'App.Info_WorkplaneZ(0) = v2(2)
		'App.Info_WorkplaneZ(1) = v3(2)
		'App.Info_WorkplaneZ(2) = v1(2)

		rc = App.feViewRegenerate( 0 )




	End If



	
End Sub
