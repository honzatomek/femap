Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim grID As Long

	grID = App.Info_NextID( FT_GROUP )

	rc = elSet.AddAll(FT_ELEM)

	rc = App.feCheckElemDistortion( elSet.ID, False, True, False, False, False, False, False, False, False, True, False, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.7, 0.0)

	Dim elSet2 As femap.Set
	Set elSet2 = App.feSet

	rc = elSet2.AddGroup( FT_ELEM, grID )

	rc= App.feWindowShow( FT_ELEM, elSet2.ID )

	
End Sub
