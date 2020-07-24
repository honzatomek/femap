Rem File: BreakCurvesAtMiddle.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim cuSet As femap.Set
	Set cuSet = App.feSet

	Dim cuID As Long

	Dim pt1 As femap.Point
	Dim pt2 As femap.Point
	Set pt1 = App.fePoint
	Set pt2 = App.fePoint

	Dim pt3(3) As Double

	Dim midPT As Variant

	Dim feCurve As femap.Curve
	Set feCurve = App.feCurve

	rc = cuSet.Select( FT_CURVE, True, "Select Curves to Split" )

	cuID = cuSet.First

	While cuID > 0
		rc = feCurve.Get( cuID )
		If feCurve.type = FCU_LINE Then
			rc = pt1.Get( feCurve.StdPoint( 0 ) )
			rc = pt2.Get( feCurve.StdPoint( 1 ) )
			pt3(0) = ( pt1.x + pt2.x ) / 2.0
			pt3(1) = ( pt1.y + pt2.y ) / 2.0
			pt3(2) = ( pt1.z + pt2.z ) / 2.0
			midPT = pt3
			rc = App.feCurveBreak( cuID, midPT )
			GoTo Next_Curve
		End If

		If feCurve.type = FCU_SOLID Then
			If feCurve.StdPoint(0) > 0 Then
				If feCurve.StdPoint(1) > 0 Then
					rc = pt1.Get( feCurve.StdPoint( 0 ) )
					rc = pt2.Get( feCurve.StdPoint( 1 ) )
					pt3(0) = ( pt1.x + pt2.x ) / 2.0
					pt3(1) = ( pt1.y + pt2.y ) / 2.0
					pt3(2) = ( pt1.z + pt2.z ) / 2.0
					midPT = pt3
					rc = App.feCurveBreak( cuID, midPT )
					GoTo Next_Curve
				End If
			End If
		End If

		rc = feCurve.ParamToXYZ( 0.5, midPT )
		rc = App.feCurveBreak( cuID, midPT )


Next_Curve:

		cuID = cuSet.Next
	Wend

	rc = App.feViewRegenerate( 0 )

End Sub
