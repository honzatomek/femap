Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim suSet As femap.Set
	Set suSet = App.feSet

	Dim suID1 As Long
	Dim suID2 As Long

	Dim ndSet1 As femap.Set
	Dim ndSet2 As femap.Set

	Set ndSet1 = App.feSet
	Set ndSet2 = App.feSet

	Dim ndCount1 As Long
	Dim ndCount2 As Long

	Dim dof(6) As Long
	Dim vdof As Variant

	dof(0) = 1
	dof(1) = 1
	dof(2) = 1
	dof(3) = 0
	dof(4) = 0
	dof(5) = 0
	vdof = dof

	Dim oVec(3) As Double
	oVec(0) = 0.0
	oVec(1) = 0.0
	oVec(2) = 1.0

	Dim vVec As Variant
	vVec = oVec

	rc = suSet.Select( FT_SURFACE, True, "Select Surfaces on One Side" )

	If suSet.count > 0 Then
		rc = ndSet1.AddSetRule( suSet.ID, FGD_NODE_ATSURFACE )
		rc = suSet.Select( FT_SURFACE, True, "Select Surfaces on Other Side" )

		If suSet.count > 0 Then
			rc = ndSet2.AddSetRule( suSet.ID, FGD_NODE_ATSURFACE )
			ndCount1 = ndSet1.count()
			ndCount2 = ndSet2.count()

			If ndCount1 > ndCount2 Then
				rc = App.feMeshClosestLink( ndSet2.ID, ndSet1.ID, 2, oVec, vdof, False )
			Else
				rc = App.feMeshClosestLink( ndSet1.ID, ndSet2.ID, 2, oVec, vdof, False )
			End If
		End If
	End If

	rc= App.feViewRegenerate( 0 )


End Sub
