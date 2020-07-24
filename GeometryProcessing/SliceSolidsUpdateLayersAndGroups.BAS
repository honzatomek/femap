Sub Main
    Dim App As femap.model
    Set App = feFemap()

    Dim soSet As femap.Set
    Set soSet = App.feSet

    Dim plBase As Variant
    Dim plNorm As Variant
    Dim plAxis As Variant

    Dim grSet As femap.Set
    Set grSet = App.feSet

    Dim allSolids As femap.Set
    Set allSolids = App.feSet

    Dim newSolids As femap.Set
    Set newSolids = App.feSet

    Dim feSolid As femap.Solid
    Set feSolid = App.feSolid

    Dim suSet As femap.Set
    Set suSet = App.feSet
    Dim cuSet As femap.Set
    Set cuSet = App.feSet
    Dim ptSet As femap.Set
    Set ptSet = App.feSet

    Dim soID As Long
    Dim soID2 As Long
    Dim grID As Long
    Dim layerID As Long

    Dim feGroup As femap.Group
    Set feGroup = App.feGroup

    Dim grSolids As femap.Set
    Set grSolids = App.feSet

    rc = soSet.Select( FT_SOLID, True, "Select Solids to Split")

    If soSet.Count > 0 Then
		rc = App.fePlanePick( "Select Split Plane", plBase, plNorm, plAxis )
		If rc = -1 Then
			soID = soSet.First
			While soID > 0
				rc = allSolids.Clear
				rc = newSolids.Clear
				rc = grSet.Clear
				rc = suSet.Clear
				rc = cuSet.Clear
				rc = ptSet.Clear
				'get all the solids in the current model
				rc = allSolids.AddAll( FT_SOLID )
				'get the current solid
				rc = feSolid.Get( soID )
				layerID = feSolid.layer
				' now see what groups it is in
				rc = feGroup.Reset
				While feGroup.Next = -1
					rc = grSolids.Clear
					rc = grSolids.AddGroup( FT_SOLID, feGroup.ID )
					If grSolids.IsAdded( soID ) Then
						' Debug Messages during coding
						'Msg = "Solid " +Str$( soID ) + " is in Group " + Str$( feGroup.ID )
						'rc = App.feAppMessage( FCM_NORMAL, Msg )
						rc = suSet.AddRule( soID, FGD_SURFACE_ONSOLID )
						rc = cuSet.AddSetRule( suSet.ID, FGD_CURVE_ONSURFACE )
						rc = ptSet.AddSetRule( cuSet.ID, FGD_POINT_ONCURVE )
						rc = feGroup.SetAddOpt( FT_SURFACE, suSet.ID, 0 )
						rc = feGroup.SetAddOpt( FT_CURVE, cuSet.ID, 0 )
						rc = feGroup.SetAddOpt( FT_POINT, ptSet.ID, 0 )
						grSet.Add( feGroup.ID )
					End If
				Wend
				' Now Split the Solid
				rc = App.feSolidSlice( -soID, plBase, plNorm, False )
				rc = newSolids.AddAll( FT_SOLID )
				rc = newSolids.RemoveSet( allSolids.ID )
				If newSolids.Count > 0 Then
					' Update the groups to include everything from the old solid
					grID = grSet.First
					While grID > 0
						rc = suSet.Clear
						rc = cuSet.Clear
						rc = ptSet.Clear
						rc = suSet.AddRule( soID, FGD_SURFACE_ONSOLID )
						rc = cuSet.AddSetRule( suSet.ID, FGD_CURVE_ONSURFACE )
						rc = ptSet.AddSetRule( cuSet.ID, FGD_POINT_ONCURVE )
						rc = feGroup.Get( grID )
						rc = feGroup.SetAdd( FT_SURFACE, suSet.ID )
						rc = feGroup.SetAdd( FT_CURVE, cuSet.ID )
						rc = feGroup.SetAdd( FT_POINT, ptSet.ID )
						rc = feGroup.Put( grID )
						rc = App.feModifyLayer( FT_SURFACE, suSet.ID, feSolid.layer )
						rc = App.feModifyLayer( FT_CURVE, cuSet.ID, feSolid.layer )
						rc = App.feModifyLayer( FT_POINT, ptSet.ID, feSolid.layer )
						grID = grSet.Next
					Wend
					soID2 = newSolids.First
					While soID2 > 0
						rc = feSolid.Get( soID2 )
						feSolid.layer = layerID
						rc = feSolid.Put( feSolid.ID )
						' now update groups
						grID = grSet.First
						While grID > 0
							rc = feGroup.Get( grID )
							rc = feGroup.Add( FT_SOLID, soID2 )
							rc = suSet.Clear
							rc = cuSet.Clear
							rc = ptSet.Clear
							rc = suSet.AddRule( soID2, FGD_SURFACE_ONSOLID )
							rc = cuSet.AddSetRule( suSet.ID, FGD_CURVE_ONSURFACE )
							rc = ptSet.AddSetRule( cuSet.ID, FGD_POINT_ONCURVE )
							rc = feGroup.SetAdd( FT_SURFACE, suSet.ID )
							rc = feGroup.SetAdd( FT_CURVE, cuSet.ID )
							rc = feGroup.SetAdd( FT_POINT, ptSet.ID )
							rc = feGroup.Put( grID )
							rc = App.feModifyLayer( FT_SURFACE, suSet.ID, feSolid.layer )
							rc = App.feModifyLayer( FT_CURVE, cuSet.ID, feSolid.layer )
							rc = App.feModifyLayer( FT_POINT, ptSet.ID, feSolid.layer )
							grID = grSet.Next
						Wend
						soID2 = newSolids.Next
					Wend
				End If
				soID = soSet.Next
			Wend

		End If
    End If

	rc = App.feViewRegenerate( -1 )
End Sub
