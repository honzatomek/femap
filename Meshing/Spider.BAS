Rem File: Spider.BAS
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

	nodeX = 0#
	nodeY = 0#
	nodeZ = 0#

	' Get a set of nodes for this rigid element

	Set nodeSet = App.feSet()

	rc = nodeSet.Select(FT_NODE, True, "Select Nodes for Rigid Spider")

	If rc = -1 Then 'return code FE_OK

    'Lets see how many nodes were selected
    nodeCount = nodeSet.count()

    If nodeCount > 0 Then

        ' Walk the nodes and find the average
        Set feNode = App.feNode()
        rc = nodeSet.Reset()
        nodeID = nodeSet.Next()

        ReDim nodeArray(nodeCount) As Long

        Dim passCount As Long

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

Call App.feViewRegenerate(0)

End Sub
