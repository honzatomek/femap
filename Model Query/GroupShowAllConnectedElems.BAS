Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim feElem As femap.Elem
    Set feElem = App.feElem

    Dim ndSet As femap.Set
    Set ndSet = App.feSet

    Dim elSet As femap.Set
    Set elSet = App.feSet

    Dim eCount As Long
    Dim AddCount As Long

    rc = feElem.SelectID( "Pick Seed Element" )
    If rc = -1 Then
    	rc = elSet.Add( feElem.ID )
    	eCount = elSet.Count()
    	AddCount = 1

    	While AddCount > 0
    		rc = ndSet.AddSetRule( elSet.ID, FGD_NODE_ONELEM )
			rc = elSet.AddSetRule( ndSet.ID, FGD_ELEM_BYNODE )
			AddCount = elSet.Count - eCount
			eCount = elSet.Count
    	Wend
		If elSet.Count > 0 Then
			rc = App.feViewShow( FT_ELEM, elSet.ID )
			Dim feGroup As femap.Group
			Set feGroup = App.feGroup
			feGroup.title = "All Mesh Connected to Elem " + Str$(feElem.ID )
			rc = feGroup.SetAdd( FT_ELEM, elSet.ID )
			rc = ndSet.Clear
			rc = ndSet.AddSetRule( elSet.ID, FGD_NODE_ONELEM )
			rc = feGroup.SetAdd( FT_NODE, ndSet.ID )
			rc = feGroup.Put( feGroup.NextEmptyID )
		End If
    End If

End Sub
