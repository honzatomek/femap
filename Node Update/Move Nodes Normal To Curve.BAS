Rem File: MoveNodesNormaltoCurve.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim lnID As Long

	Dim feCurve As femap.Curve
	Set feCurve = App.feCurve
	Dim feSet As femap.Set
	Set feSet = App.feSet
	Dim ndLoc(3) As Double
	Dim vndLoc As Variant
	Dim projLoc As Variant
	Dim feNode As femap.Node
	Set feNode = App.feNode

	Dim ndID As Long

	Dim moveVec(3) As Double
	Dim vmoveVec As Variant

	Dim nVec As Variant

	Dim moveDist As Double

	lineID = 0

	rc = feSet.SelectID( FT_CURVE, "Select Line", lnID )

	If lnID Then
		rc = feSet.Select(FT_NODE, True, "Select Nodes to Move")
		rc = feCurve.Get( lnID )

		rc = App.feGetReal( "Offset Distance", 0.0, 1000.0, moveDist )

		rc = feSet.Reset

		ndID = feSet.Next

		While ndID
				rc = feNode.Get( ndID )
				ndLoc(0) = feNode.x
				ndLoc(1) = feNode.y
				ndLoc(2) = feNode.z
				vndLoc = ndLoc
				rc = feCurve.project( ndLoc, projLoc )
				moveVec(0) =  projLoc(0) - feNode.x
				moveVec(1) =  projLoc(1) - feNode.y
				moveVec(2) =  projLoc(2) - feNode.z

				vmoveVec = moveVec

				rc = App.feVectorUnit(vmoveVec, nVec )

				feNode.x = feNode.x + nVec(0) * moveDist
				feNode.y = feNode.y + nVec(1) * moveDist
				feNode.z = feNode.z + nVec(2) * moveDist

				rc = feNode.Put(ndID)
				ndID = feSet.Next
		Wend

		rc = App.feViewRegenerate( 0 )
	End If



	
End Sub
