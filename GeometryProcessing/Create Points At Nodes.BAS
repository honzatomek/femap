Rem File: CreatePointsAtNodes.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ndSet As femap.Set
	Set ndSet = App.feSet
	Dim ndID As Long
	Dim feNode As femap.Node
	Set feNode = App.feNode

	Dim fePoint As femap.Point
	Set fePoint = App.fePoint()


	rc = ndSet.Select( FT_NODE, True, "Select Nodes to Make Points" )

	rc = ndSet.Reset

	ndID = ndSet.First()

	While ndID > 0
		rc = feNode.Get( ndID )

		fePoint.x = feNode.x
		fePoint.y = feNode.y
		fePoint.z = feNode.z

		rc = fePoint.Put( fePoint.NextEmptyID)

		ndID = ndSet.Next()
	Wend
	rc = App.feViewRegenerate(0)




	
End Sub
