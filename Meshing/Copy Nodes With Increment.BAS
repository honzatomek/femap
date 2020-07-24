Rem File: CopyNodesWithIncrement.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim vecLength As Double
	Dim vecBase As Variant
	Dim vecDir As Variant
	Dim ndID As Long
	Dim nInc As Long

	'If you always want to use the same increment, remove the apostrophe from in front of this line:

	'nInc = 200

	'...and enter an increment value.  Then "comment out this line:

	rc = App.feGetInt ("Enter Increment", 1, 9999999, nInc)

	'with an apostrophe.

	Dim feNode As femap.Node
	Set feNode = App.feNode
	Dim feNewNode As femap.Node
	Set feNewNode = App.feNode

	rc = ndSet.Select( FT_NODE, True, "Select Nodes to Copy")
	rc = App.feVectorPick( "Select Copy Vector", False, vecLength, vecBase, vecDir )

	rc = ndSet.Reset
	ndID = ndSet.Next()
	While ndID > 0
		rc = feNode.Get( ndID )
		If feNode.Exist( ndID + nInc ) Then
			myString = "Node " + Str$( ndID +nInc ) + " already exists, cannot copy."
			rc = App.feAppMessage( FCL_BLACK, myString )
		Else
			feNewNode.x = feNode.x + vecLength * vecDir(0)
			feNewNode.y = feNode.y + vecLength * vecDir(1)
			feNewNode.z = feNode.z + vecLength * vecDir(2)
			rc = feNewNode.Put( ndID + nInc )
		End If
		ndID = ndSet.Next()
	Wend

	rc = App.feViewRegenerate( 0 )

End Sub
