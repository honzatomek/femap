Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ndSet As femap.Set
	Set ndSet = App.feSet
	Dim ndID As Long

	Dim feNode As femap.Node
	Set feNode = App.feNode

	Dim feText As femap.text
	Set feText = App.feText

	feText.ModelPosition = True
	feText.AllViews = True
	feText.drawBorder = True
	feText.backColor = FCL_YELLOW

	ndID = 0

	rc = ndSet.SelectID( FT_NODE, "Select Node to Place Text", ndID )

	If ndID > 0 Then
		rc = feNode.Get( ndID )
		feText.TextPosition(0) = feNode.x
		feText.TextPosition(1) = feNode.y
		feText.TextPosition(2) = feNode.z
		feText.text = "Node " + Str$(ndID)
	End If

	rc = feText.Put( feText.NextEmptyID )

	rc = App.feViewRegenerate( 0 )

End Sub
