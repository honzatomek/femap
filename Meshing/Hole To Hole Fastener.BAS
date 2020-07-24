Sub Main

    Dim App As femap.model
    Set App = feFemap()

    Dim feNode As femap.Node

    Dim nodeSet As femap.Set

    Dim feElem As femap.Elem

    Dim nodeCount As Long
    Dim nodeX As Double
    Dim nodeY As Double
    Dim nodeZ As Double
    Dim nodeID As Long
    Dim elemID As Long
    Dim vNodeArray As Variant
    Dim RAD As Double

    Dim cuSet As femap.Set
    Set cuSet = App.feSet

	Set nodeSet = App.feSet()
    Set cuSet = App.feSet

    Set nodeSet = App.feSet()

   	Dim cuID As Long

	Dim feCurve As femap.Curve
	Set feCurve = App.feCurve

	Dim cuCenter As Variant
	Dim cuNormal As Variant
	Dim cuStartPt As Variant
	Dim cuEndPt As Variant
	Dim cuAngle As Double
	Dim cuRadius As Double
	Dim pt1(3) As Double
	Dim pt2(3) As Double
	Dim ptID1 As Long
	Dim ptID2 As Long

	Dim fePoint As femap.Point
	Set fePoint = App.fePoint

	Dim passCount As Long

	Dim nodeSet1 As femap.Set
    Dim nodeCount1 As Long
    Dim nodeX1 As Double
    Dim nodeY1 As Double
    Dim nodeZ1 As Double
    Dim nodeID1 As Long
    Dim elemID1 As Long
    Dim vNodeArray1 As Variant
    Dim passCount1 As Long

	Dim feProp As femap.Prop
	Set feProp = App.feProp

	Dim BeamArea As Double
    Dim Vec(3) As Variant
    Dim X As Double
    Dim Y As Double
    Dim Z As Double
    Dim X1 As Double
    Dim Y1 As Double
    Dim Z1 As Double
    Dim VecP As Variant

    Set nodeSet1 = App.feSet()

    Dim feMatl As femap.Matl
    Set feMatl = App.feMatl
    Dim matID As Long

    matID = App.Info_ActiveID( FT_MATL )

    If matID = 0 Then
    	Msg = "No active material for fastener, please activate a material and re-run."
        rc = MsgBox( Msg, vbOkOnly, "No Material" )
        GoTo Jumping_Out
    End If

    rc = feMatl.Get( matID )

DO_IT_AGAIN:
    RAD= 0.0
    nodeX = 0#
    nodeY = 0#
    nodeZ = 0#
    nodeX1 = 0#
    nodeY1 = 0#
    nodeZ1 = 0#


     ' Get a set of nodes for this rigid element

	rc = cuSet.Select( FT_CURVE, True, "Select Curves at Fastener Bottom")
	If cuSet.count = 0 Then
		GoTo Jumping_Out
	End If

	cuID = cuSet.First
	rc = feCurve.Get(cuID)
	rc = feCurve.ArcCircleInfo( cuCenter, cuNormal, cuStartPt, cuEndPt, cuAngle, cuRadius )
	If rc <> FE_OK Then
    	Msg = "Curves must be circular, exiting."
        rc = MsgBox( Msg, vbOkOnly, "No Material" )
        GoTo Jumping_Out
	End If

    RAD = cuRadius

	rc = nodeSet.Clear()
	rc = nodeSet.AddSetRule( cuSet.ID, FGD_NODE_ATCURVE )

    If rc = -1 Then 'return code FE_OK
	    'Lets see how many nodes were selected
	    nodeCount = nodeSet.count()
	    If nodeCount > 0 Then
	       ' Walk the nodes and find the average
	        Set feNode = App.feNode()
	        rc = nodeSet.Reset()
	        nodeID = nodeSet.Next()
	        ReDim nodeArray(nodeCount) As Long
	        passCount = 0
	        Do While nodeID <> 0
	            nodeArray(passCount) = nodeID
	            passCount = passCount + 1
	            rc = feNode.Get(nodeID)
	            nodeX = nodeX + feNode.x
	            nodeY = nodeY + feNode.y
	            nodeZ = nodeZ + feNode.z
	            nodeID = nodeSet.Next()
	        Loop
			vNodeArray = nodeArray
	        nodeID = feNode.NextEmptyID
	        feNode.ID = nodeID
	        feNode.x = nodeX / nodeCount
	        feNode.y = nodeY / nodeCount
	        feNode.z = nodeZ / nodeCount
	        rc = feNode.Put(nodeID)
	        If rc = -1 Then 'return code FE_OK
	        	' create the element
				Set feElem = App.feElem()
				elemID = feElem.NextEmptyID
	            feElem.type = FET_L_RIGID
                feElem.topology = FTO_RIGIDLIST
	            feElem.Node(0) = nodeID 'Independent Node
	            feElem.release(0, 0) = 1
	            feElem.release(0, 1) = 1
	            feElem.release(0, 2) = 1
	            feElem.ID = elemID
                rc = feElem.PutNodeList(0, nodeCount, vNodeArray, Null, Null, Null)
	            rc = feElem.Put(elemID)
	        End If
	    End If
	End If

    ' Get a set of nodes for this rigid element
    rc = cuSet.Select( FT_CURVE, True, "Select Curves at Top at Hole")
    rc = nodeSet1.Clear()
    rc = nodeSet1.AddSetRule( cuSet.ID, FGD_NODE_ATCURVE )
    If rc = -1 Then 'return code FE_OK
	    'Lets see how many nodes were selected
	    nodeCount1 = nodeSet1.count()
	    If nodeCount1 > 0 Then
	        ' Walk the nodes and find the average
	        Set feNode = App.feNode()
	        rc = nodeSet1.Reset()
	        nodeID1 = nodeSet1.Next()
	        ReDim nodeArray1(nodeCount1) As Long

	        passCount1 = 0
	        Do While nodeID1 <> 0
	        	nodeArray1(passCount1) = nodeID1
        		passCount1 = passCount1 + 1
	            rc = feNode.Get(nodeID1)
	            nodeX1 = nodeX1 + feNode.x
	            nodeY1 = nodeY1 + feNode.y
	            nodeZ1 = nodeZ1 + feNode.z
	            nodeID1 = nodeSet1.Next()
	        Loop
        	vNodeArray1 = nodeArray1
	        nodeID1 = feNode.NextEmptyID
	        feNode.ID = nodeID1
	        feNode.x = nodeX1/ nodeCount1
	        feNode.y = nodeY1/ nodeCount1
	        feNode.z = nodeZ1/ nodeCount1
	        rc = feNode.Put(nodeID1)
	        If rc = -1 Then 'return code FE_OK
	            ' create the element
	            Set feElem = App.feElem()
	            elemID = feElem.NextEmptyID
                feElem.type = FET_L_RIGID
                feElem.topology = FTO_RIGIDLIST
	            feElem.Node(0) = nodeID1 'Independent Node
	            feElem.release(0, 0) = 1
	            feElem.release(0, 1) = 1
	            feElem.release(0, 2) = 1
	            feElem.ID = elemID
	            rc = feElem.PutNodeList(0, nodeCount1, vNodeArray1, Null, Null, Null)
	            rc = feElem.Put(elemID)
	        End If
	    End If
	End If

    Dim El As femap.Elem
    Set El = App.feElem
    Dim BeamElID As Long
    Dim PropID As Long
    Dim P As femap.Prop
    Set P = App.feProp
    Dim BeamID As Long

	'Section Values A
	If RAD = 0.0 Then
		 App.feGetReal ("Enter Bolt Radius", 0.000000001, 100000, RAD)
	End If

	PropID = 0


	BeamArea = 3.141592653 * RAD * RAD

	While feProp.Next
		If feProp.title = "API Fastener" Then
			If Abs( feProp.pval(0) - BeamArea ) < 0.001 Then
				PropID = feProp.ID
				GoTo GOT_A_PROP
			End If
		End If
	Wend

	PropID = P.NextEmptyID
	P.title = "API Fastener"
    P.type = FET_L_BEAM
    P.flagI(1) = 5
    P.pval (40) = RAD
    P.ComputeShape (False, False, True)
    P.Put(PropID)

GOT_A_PROP:

    Msg = "Fastener Diameter " + Format$(2*cuRadius,"###.####") + ", Property " + Str$(PropID) + ", Material ID " + Str$(matID) + ", Titled " + feMatl.title
    rc = App.feAppMessage( FCM_NORMAL, Msg )

    BeamElID = El.NextEmptyID
    El.type = FET_L_BEAM
    El.Node (0) = nodeID
    El.Node (1) = nodeID1
    rc = feNode.Get (nodeID)
    X = feNode.x
    Y = feNode.y
    Z = feNode.z
    rc = feNode.Get (nodeID1)
    X1 = feNode.x
    Y1 = feNode.y
    Z1 = feNode.z
    Vec(0) = X - X1
    Vec(1) = Y - Y1
    Vec(2) = Z - Z1
    rc = App.feVectorPerpendicular (Vec, VecP)
    El.orient(0) = VecP(0)
    El.orient(1) = VecP(1)
    El.orient(2) = VecP(2)
    El.propID = PropID
    El.Put (BeamElID)

    rc = App.feViewRegenerate(0)

    GoTo DO_IT_AGAIN

Jumping_Out:

End Sub

