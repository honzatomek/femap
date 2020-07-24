Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim feCSys As femap.CSys
	Set feCSys = App.feCSys
	Dim feCSysID As Long

	Dim surfID As Long
	Dim surfSet As femap.Set
	Set surfSet = App.feSet

	Dim feSurf As femap.Surface
	Set feSurf = App.feSurface

	Dim feNode As femap.Node
	Set feNode = App.feNode

	Dim ndID As Long

	Dim feProp As femap.Prop
	Set feProp = App.feProp
	Dim feProp2 As femap.Prop
	Set feProp2 = App.feProp
	Dim propID As Long

	Dim xyz(3) As Double
	Dim vxyz As Variant
	Dim nPt(3) As Double
	Dim vnPt As Variant
	Dim pPt(3) As Double
	Dim vpPt As Variant
	Dim u As Double
	Dim v As Double
	Dim surfNormal As Variant
	Dim surfPerp As Variant

	Dim x1Local(3) As Double
	Dim x0Local(3) As Double
	Dim x1Global As Variant
	Dim x0Global As Variant
	Dim vecA(3) As Double
	Dim vecB(3) As Double
	Dim dot As Double

	Dim feNode2 As femap.Node
	Set feNode2 = App.feNode
	Dim feNode2ID As Long

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	surfID = 0
	rc = surfSet.SelectID( FT_SURFACE, "Select Surface for One-Way Constraint", surfID )

	If surfID > 0 Then
		rc = ndSet.AddRule( surfID, FGD_NODE_ATSURFACE )
		If ndSet.count > 0 Then

			'load the surface
			rc = feSurf.Get( surfID )

			ndID = ndSet.First()
			While ndID > 0
				rc = feNode.Get( ndID )
				' find the normal here and create a CSys
				xyz(0) = feNode.x
				xyz(1) = feNode.y
				xyz(2) = feNode.z
				vxyz = xyz

				feSurf.XYZToParam( vxyz, u, v )
				feSurf.normal( u, v, surfNormal )

				rc = feProp2.Reset
				While feProp2.Next
					' see if there is matching gap property already
					If feProp2.title = "One-Way Gap Property" Then
						rc = feCSys.Get( feProp2.refCS)
						' If this x-axis is aligned, use this prop
						x1Local(0) = 1.0
						x1Local(1) = 0.0
						x1Local(2) = 0.0
						x0Local(0) = 0.0
						x0Local(1) = 0.0
						x0Local(2) = 0.0
						rc = App.feCoordTransform( feProp.refCS, x1Local, 0, x1Global )
						rc = App.feCoordTransform( feProp.refCS, x0Local, 0, x0Global )
						vecA(0) = x1Global(0) - x0Global(0)
						vecA(1) = x1Global(1) - x0Global(1)
						vecA(2) = x1Global(2) - x0Global(2)
						vecB(0) = surfNormal(0)
						vecB(1) = surfNormal(1)
						vecB(2) = surfNormal(2)
						rc = App.feVectorDotProduct( vecA, vecB, dot )
						If Abs( dot - 1.0) < 0.001 Then
							propID = feProp2.ID
							GoTo GOT_A_PROP
						End If
					End If
				Wend

				App.feVectorPerpendicular( surfNormal, surfPerp )

				nPt(0) = xyz(0) + surfNormal(0)
				nPt(1) = xyz(1) + surfNormal(1)
				nPt(2) = xyz(2) + surfNormal(2)
				vnPt = nPt

				pPt(0) = xyz(0) + surfPerp(0)
				pPt(1) = xyz(1) + surfPerp(1)
				pPt(2) = xyz(2) + surfPerp(2)
				vpPt = pPt

				feCSys.type = FCS_RECTANGULAR
				feCSys.title = "CSys for Node " + Str$( ndID )
				feCSys.XYPoints( xyz, vnPt, vpPt )
				feCSysID = App.Info_NextID(FT_CSYS )
				feCSys.Put( feCSysID )

				' Let's create our GAP Property
				feProp.type = FET_L_GAP
				feProp.pval(1) = 1.0E6 'Compression stiffness only
				feProp.title = "One-Way Gap Property for Node" + Str$(ndID)
				feProp.refCS = feCSysID
				propID = feProp.NextEmptyID
				feProp.flag(0) = True
				feProp.Put( propID )

GOT_A_PROP:

				feNode2.x = feNode.x
				feNode2.y = feNode.y
				feNode2.z = feNode.z

				feNode2ID = feNode.NextEmptyID

				rc =feNode2.Put( feNode2ID )

				feElem.type = FET_L_GAP
				feElem.topology = FTO_LINE2
				feElem.propID = propID
				feElem.Node(0) = ndID
				feElem.Node(1) = feNode2ID
				feElem.orient(0) = 1.0
				feElem.orient(1) = 0.0
				feElem.orient(2) = 0.0
				'feElem.Nodes(0) = ndID
				'feElem.Nodes(1) = feNode2ID
				feElem.Put( feElem.NextEmptyID )

				' Also need to constriain the node
				rc = App.feModifyPermBC( -feNode2ID, True, True, True, True, True, True )
				ndID = ndSet.Next()
			Wend
		End If
	End If
End Sub
