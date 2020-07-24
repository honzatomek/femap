Rem File: CreateNodesProjectNodesOntoCurve.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feNode1 As femap.Node
	Set feNode1 = App.feNode

	Dim feNode2 As femap.Node
	Set feNode2 = App.feNode
	Dim ndSet As femap.Set
	Set ndSet = App.feSet
	Dim cuSet As femap.Set
	Set cuSet = App.feSet

	Dim cuID As Long
	Dim ndID As Long

	Dim ndLoc1(3) As Double
	Dim vndLoc1 As Variant
	Dim vndLoc2 As Variant

	rc = ndSet.Select( FT_NODE, True, "Select Nodes to Project")
	rc = cuSet.SelectID( FT_CURVE, "Select Curve to Project On To", cuID )

	ndID = ndSet.First
	While ndID > 0
		rc = feNode1.Get(ndID)
		ndLoc1(0) = feNode1.x
		ndLoc1(1) = feNode1.y
		ndLoc1(2) = feNode1.z

		vndLoc1 = ndLoc1

		rc = App.feCoordOntoCurve( cuID, vndLoc1, vndLoc2 )

		feNode2.x = vndLoc2(0)
		feNode2.y = vndLoc2(1)
		feNode2.z = vndLoc2(2)

		rc = feNode2.Put( feNode2.NextEmptyID )

		ndID = ndSet.Next
	Wend

rc = App.feViewRegenerate( 0 )

End Sub
