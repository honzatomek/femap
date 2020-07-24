Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feOutput As femap.output
	Set feOutput = App.feOutput

	Dim newOutput As femap.output
	Set newOutput = App.feOutput

	Dim ovID As Long
	Dim j As Long

	Dim nElem As Long
	Dim maxCorner As Long
	Dim eID As Variant
	Dim cenVal As Variant
	Dim c1 As Variant
	Dim c2 As Variant
	Dim c3 As Variant
	Dim c4 As Variant
	Dim c5 As Variant
	Dim c6 As Variant
	Dim c7 As Variant
	Dim c8 As Variant
	Dim maxPct As Double
	Dim p1 As Double

	'Elemental to Nodal Conversion Variables
	Dim approach(1) As Long
	Dim from_SetID(1) As Long
	Dim from_vectorID(1) As Long
	Dim to_setID(1) As Long
	Dim to_vectorID(1) As Long
	Dim groupID(1) As Long
	Dim factor(1) As Double
	Dim overwrite(1) As Boolean
	Dim ply_summary(1) As Boolean
	Dim set_info(1) As Boolean

	Dim v1 As Variant
	Dim v2 As Variant
	Dim v3 As Variant
	Dim v4 As Variant
	Dim v5 As Variant
	Dim v6 As Variant
	Dim v7 As Variant
	Dim v8 As Variant
	Dim v9 As Variant
	Dim v10 As Variant
	Dim v11 As Variant

    Dim cuHLSet As femap.Set
    Set cuHLSet = App.feSet

    Dim dMaxval As double
    
    If App.Info_ActiveID(FT_OUT_CASE) Then
		If feOutput.Get(60031) Then
           If	feOutput.GetElemWithCorner( nElem, maxCorner, eID, cenVal, c1, c2, c3, c4, c5, c6, c7, c8) Then
                dMaxval = feOutput.maxval
                newOutput.ID = 8999999
			    ovID = newOutput.NextEmptyID()
                newOutput.InitScalarAtElem( App.Info_ActiveID(FT_OUT_CASE), ovID, "Vec", 7, FALSE )

    			For j = 0 To nElem-1
    				maxPct = 0.0
                    If cenVal(j) > dMaxval * .8 Then
						p1 = (Abs( c1(j) / cenVal(j) ) - 1 )*100
						If p1 > maxPct Then
							maxPct = p1
						End If
						p1 = (Abs( c2(j) / cenVal(j) ) - 1 )*100
						If p1 > maxPct Then
							maxPct = p1
						End If
						p1 = (Abs( c3(j) / cenVal(j) ) - 1 )*100
						If p1 > maxPct Then
							maxPct = p1
						End If
						p1 = (Abs( c5(j) / cenVal(j) ) - 1 )*100
						If p1 > maxPct Then
							maxPct = p1
						End If
						newOutput.value(eID(j)) = maxPct
					Else
						newOutput.value(eID(j)) = 0.0
					End If
				Next j
				rc = newOutput.Put(-1)
			End If
		End If
	End If

	Dim nodalOutVec As output
	Set nodalOutVec = App.feOutput

	nodalOutVec.ID = newOutput.ID
	nodalOutVec.ID = nodalOutVec.NextEmptyID

	approach(0) = 	1 'Max
	from_SetID(0) = App.Info_ActiveID(FT_OUT_CASE )
	from_vectorID(0) = newOutput.ID
	to_setID(0) = from_SetID(0)
	groupID(0) = 0

	to_vectorID(0) = nodalOutVec.ID
	factor(0) = 1.0
	overwrite(0) = False
	ply_summary(0) = False
	set_info(0) = False

	v1 = approach
	v2 = from_SetID
	v3 = from_vectorID
	v4 = to_setID
	v5 = groupID
	v6 = newVectorID
	v7 = to_vectorID
	v8 = factor
	v9 = overwrite
	v10 = ply_summary
	v11 = set_info

	rc = App.feOutputProcess( 1, True, 4, v1, v2, v3, v4, v7, v5, v8, v9, v10, v11)

	nodalOutVec.setID = to_setID(0)
	nodalOutVec.Get( to_vectorID(0) )

	Dim solidID As Long
	Dim solidSet As femap.Set
	Set solidSet = App.feSet

	rc = solidSet.SelectID( FT_SOLID, "Select Solid To Update Mesh Size", solidID )

	Dim cuID As Long
	Dim cuSet As femap.Set
	Set cuSet = App.feSet

	' make a set with all of the curves of the selected solid
	rc = cuSet.AddRule( solidID, FGD_CURVE_ONSOLID )
	rc = cuSet.Reset

	cuID = cuSet.First()

	Dim nodeSET As femap.Set
	Set nodeSET = App.feSet

	Dim UpdateFlag As Boolean
	Dim ndID As Long

	Dim feCurve As femap.Curve
	Set feCurve = App.feCurve

	While cuID > 0
		rc = nodeSET.Clear()
		rc = nodeSET.AddRule( cuID, FGD_NODE_ATCURVE )
		rc = nodeSET.Reset()
		UpdateFlag = False
		ndID = nodeSET.First()
		While ndID > 0
			If nodalOutVec.value( ndID ) > 10.0 Then
				UpdateFlag = True
			End If
			ndID = nodeSET.Next()
		Wend
		If UpdateFlag Then
			rc = feCurve.Get( cuID )
			Msg = "Updating Curve " + Str$(cuID)
			rc = App.feAppMessage( FCM_NORMAL, Msg )
			feCurve.meshElem(0) = Int( feCurve.meshElem(0) * 1.4 )
            cuHLSet.Add( cuID )
			rc = feCurve.Put( cuID )
		End If
		cuID = cuSet.Next()
	Wend

    If cuHLSet.count > 0 Then
    	Dim elemSET As femap.Set
        Set elemSET = App.feSet

        nodeSET.Clear()

        rc = nodeSET.AddRule( solidID, FGD_NODE_ATSOLID )
        rc = elemSET.AddRule( solidID, FGD_ELEM_ATSOLID )
        rc = App.feDelete( FT_ELEM, elemSET.ID )
         rc = App.feDelete( FT_NODE, nodeSET.ID )

		rc = App.feDeleteAll( False, False, True, True )

        'rc = App.feMeshSizeSolid( - solidID, False, 1.0, 0, 0, 0, 0, 1.2, False, 1.0, 0.0, False, False, False, False, False )

        rc = App.feMeshTetSolid( FT_SOLID, -solidID,  0, False, True, 1.1, True, True, False, 0.0, 0 )
    	 rc = App.feViewShow( FT_CURVE, cuHLSet.ID )
         rc = MsgBox( "Tightened Mesh Size Here", vbOkOnly, "Mesh Refinement")

         Dim analysisSetID As Long
         analysisSetID = App.Info_ActiveID( FT_AMGR_DIR )

         Dim feAnalysisSet As femap.AnalysisMgr
         Set feAnalysisSet = App.feAnalysisMgr

         rc = feAnalysisSet.Get( analysisSetID )

         rc = feAnalysisSet.Analyze( analysisSetID )

         Dim feView As femap.View
         Set feView = App.feView

         Dim viewID As Long

         rc = App.feAppGetActiveView( viewID )
         rc = feView.Get( viewID )
         feView.Contour = FCF_CONTOUR
         rc = App.feViewRegenerate( 0 )

    End If
End Sub
