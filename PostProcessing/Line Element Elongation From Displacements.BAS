Sub Main
    Dim App As femap.model
    Set App = feFemap()

    App.feAppMessage( FCM_HIGHLIGHT, "Compute Line Element Elongation Results" )

    Dim cRes As femap.Results
    Set cRes = App.feResults

    Dim nNumComp As Long
    Dim vColXYZ As Variant
    Dim sSets As femap.Set
    Dim nSetID As Long
    Dim nVecID As Long

    Dim cElemSet As femap.Set
    Set cElemSet = App.feSet
    cElemSet.Select( FT_ELEM, True, "Select Line Elements" )

    ' Remove any selected elements that are not line elements......
    Dim cLineElemSet As femap.Set
    Set cLineElemSet = App.feSet
	 cLineElemSet.AddRule( FTO_LINE2, FGD_ELEM_BYSHAPE )
	cElemSet.RemoveNotCommon( cLineElemSet.ID )

	If cElemSet.IsEmpty() Then
		MsgBox( "No Line Elements Selected. Unable to Process.", vbOkOnly, "Error" )
		Exit All
	End If

	 Dim cNodeSet As femap.Set
	 Set cNodeSet = App.feSet
	 cNodeSet.AddSetRule( cElemSet.ID, FGD_NODE_ONELEM )


	 Dim cOutput As femap.Output
	 Set cOutput = App.feOutput

	 Dim cElem As femap.Elem
	 Set cElem = App.feElem

	 Dim cNode1 As femap.Node
	 Dim cNode2 As femap.Node
	 Set cNode1 = App.feNode
	 Set cNode2 = App.feNode

	 Dim vDef1 As Variant
	 Dim vDef2 As Variant


	 Dim nElemID As Variant
 	 Dim nNumElem As Long
 	 cElemSet.GetArray( nNumElem, nElemID )

	 Dim dElong() As Double
	 ReDim dElong( nNumElem ) As Double
	 Dim dPct() As Double
	 ReDim dPct( nNumElem ) As Double
	 Dim dX() As Double
	 Dim dY() As Double
	 Dim dZ() As Double
	 ReDim dX( nNumElem ) As Double
	 ReDim dY( nNumElem ) As Double
	 ReDim dZ( nNumElem ) As Double

	 Dim vUndef(2) As Double
	 Dim vDef(2) As Double

	 Dim i As Long
	 Dim e As Long
     Dim dDefLen As Double
     Dim dUndefLen As Double

    If App.feSelectOutputSets( "Choose Output Set(s)", sSets ) = FE_OK Then
		sSets.Reset()
	    nSetID = sSets.Next()
		While nSetID <> FE_FAIL
			cRes.Clear()
			cRes.DataNeeded( FT_NODE, cNodeSet.ID )
  			If  cRes.AddColumn( nSetID, 1, True, nNumComp, vColXYZ ) = FE_OK Then
				cRes.Populate()

				For e = 0 To nNumElem-1
					If cElem.Get( nElemID(e) ) = FE_OK Then
						If cNode1.Get( cElem.Node(0) )= FE_OK Then
							If cNode2.Get( cElem.Node(1) ) = FE_OK Then
								If cRes.GetRowByID( cNode1.ID, vDef1 )  = FE_OK Then
									If cRes.GetRowByID( cNode2.ID, vDef2 ) = FE_OK Then
										For i=0 To 2
											vUndef(i) = cNode2.xyz(i)-cNode1.xyz(i)
											vDef(i) = ( vDef2(vColXYZ(i+1))+cNode2.xyz(i) ) - ( vDef1(vColXYZ(i+1))+cNode1.xyz(i) )
										Next i
                                        App.feVectorLength( vUndef, dUndefLen )
                                        App.feVectorLength( vDef, dDefLen )
                                        dElong(e) = dDefLen - dUndefLen

                                        If  dUndefLen = 0.0 Then
                                        	dPct(e) = 0.0
                                        Else
                                        	dPct(e) = ( dElong(e) / dUndefLen ) * 100.0

                                        dX(e) = vDef2(vColXYZ(1)) -  vDef1(vColXYZ(1))
                                        dY(e) = vDef2(vColXYZ(2)) -  vDef1(vColXYZ(2))
										dZ(e) = vDef2(vColXYZ(3)) -  vDef1(vColXYZ(3))
                                        End If
									End If
								End If
							End If
						End If
					End If
				Next e

				cOutput.setID = nSetID
				cOutput.ID = 9000000
				cOutput.InitScalarAtElem( nSetID, cOutput.NextEmptyID(), "Elongation", FOT_ANY, True )
				cOutput.PutScalarAtElem( nNumElem, nElemID, dElong )
				cOutput.Put( cOutput.ID )

				cOutput.InitScalarAtElem( nSetID, cOutput.NextEmptyID(), "Elongation Percent", FOT_ANY, True )
				cOutput.PutScalarAtElem( nNumElem, nElemID, dPct )
				cOutput.Put( cOutput.ID )

				cOutput.InitScalarAtElem( nSetID, cOutput.NextEmptyID(), "X Relative Deformation", FOT_ANY, True )
				cOutput.PutScalarAtElem( nNumElem, nElemID, dX )
				cOutput.Put( cOutput.ID )

				cOutput.InitScalarAtElem( nSetID, cOutput.NextEmptyID(), "Y Relative Deformation", FOT_ANY, True )
				cOutput.PutScalarAtElem( nNumElem, nElemID, dY )
				cOutput.Put( cOutput.ID )

				cOutput.InitScalarAtElem( nSetID, cOutput.NextEmptyID(), "Z Relative Deformation", FOT_ANY, True )
				cOutput.PutScalarAtElem( nNumElem, nElemID, dZ )
				cOutput.Put( cOutput.ID )
			Else
				App.feAppMessage( FCM_ERROR, "No Displacements Found for Output Set "+Str$(nSetID) )
		    End If
			nSetID = sSets.Next()
		Wend
    End If
End Sub
