Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim feGroup As femap.Group
    Set feGroup = App.feGroup

    Dim surfID As Long

    Dim surfSet As femap.Set
    Set surfSet = App.feSet

    Dim ndSet As femap.Set
    Set ndSet = App.feSet

    rc = surfSet.SelectID( FT_SURFACE, "Select Surface to Create Node Group", surfID )

    rc = ndSet.AddRule( surfID, FGD_NODE_ATSURFACE )

    rc = feGroup.SetAdd( FT_NODE, ndSet.ID )

    feGroup.title = "Nodes On Surface " + Str$(surfID )

    rc = feGroup.Put( feGroup.NextEmptyID )

	
End Sub
