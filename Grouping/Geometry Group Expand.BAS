Rem File: GeometryGroupExpand.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim cuSet As femap.Set
	Set cuSet = App.feSet

	Dim suSet As femap.Set
	Set suSet = App.feSet

	Dim feGroup As femap.Group
	Set feGroup = App.feGroup

	grID = App.Info_ActiveID( FT_GROUP )

	rc = feGroup.Get( grID )

	rc = cuSet.AddGroup( FT_CURVE, grID )
	rc = suSet.AddGroup(FT_SURFACE, grID )

	rc = cuSet.AddSetRule( suSet.ID, FGD_CURVE_ONSURFACE )
	rc = suSet.AddSetRule( cuSet.ID, FGD_SURFACE_BYCURVE )
	rc = cuSet.AddSetRule( suSet.ID, FGD_CURVE_ONSURFACE )

	rc = feGroup.SetAdd( FT_CURVE, cuSet.ID )
	rc = feGroup.SetAdd( FT_SURFACE, suSet.ID )

	rc = feGroup.Put( grID )

	rc = App.feViewRegenerate( 0 )

End Sub
