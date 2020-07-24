Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim data As femap.DataTable
	Set data = App.feDataTable

	rc = data.Clear

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim C1 As Variant 'Node IDs
	Dim C2 As Variant 'Node X
	Dim C3 As Variant 'Node Y
	Dim C4 As Variant 'Node Z
	Dim C5 As Variant 'X-Disp
	Dim C6 As Variant 'Y-Disp
	Dim C7 As Variant 'Z-Disp

	Dim isThereOutput As Boolean

	If App.Info_ActiveID( FT_OUT_CASE ) > 0 Then
		isThereOutput = True
	Else
		isThereOutput = False
		Msg = "No Current Output Set, Exiting..."
		rc = MsgBox( Msg, vbOkOnly, "Displacements to Data Table" )
		GoTo Jumping_Out
	End If

	Dim xDisp As femap.Output
	Set xDisp = App.feOutput
	Dim yDisp As femap.Output
	Set yDisp = App.feOutput
	Dim zDisp As femap.Output
	Set zDisp = App.feOutput

	Dim feNode As femap.Node
	Set feNode = App.feNode
	Dim ndID As Long
	Dim count As Long
	Dim ndCount As Long

	count = 0

	rc = ndSet.Select( FT_NODE, True, "Select Nodes to Add to Data Table" )
	ndCount = ndSet.Count
	If ndCount > 0 Then
		App.feAppManagePanes ( "Data Table", 1)
		data.Lock ( False )
		ReDim C1(ndCount) As Long
		ReDim C2(ndCount) As Double
		ReDim C3(ndCount) As Double
		ReDim C4(ndCount) As Double
		ReDim C5(ndCount) As Double
		ReDim C6(ndCount) As Double
		ReDim C7(ndCount) As Double

		If isThereOutput Then
			xDisp.setID = App.Info_ActiveID( FT_OUT_CASE )
			yDisp.setID = App.Info_ActiveID( FT_OUT_CASE )
			zDisp.setID = App.Info_ActiveID( FT_OUT_CASE )
			rc = xDisp.Get( 2 )
			rc = yDisp.Get( 3 )
			rc = zDisp.Get( 4 )
		End If

		ndID = ndSet.First
		While ndID > 0
			rc = feNode.Get( ndID )
			C1(count) = ndID
			C2(count) = feNode.x
			C3(count) = feNode.y
			C4(count) = feNode.z
			If isThereOutput Then
				C5(count) = xDisp.Value( ndID )
				C6(count) = yDisp.Value( ndID )
				C7(count) = zDisp.Value( ndID )
			End If
			count = count + 1
			ndID = ndSet.Next
		Wend
		Dim newColID As Long
		'rc = data.AddColumn( False, False, FT_NODE, 0, "Node ID", FCT_INT, ndCount, C1, C1, newColID )
		rc = data.AddColumn( False, False, FT_NODE, 0, "Node X", FCT_DOUBLE, ndCount, C1, C2, newColID )
		rc = data.AddColumn( False, False, FT_NODE, 0, "Node Y", FCT_DOUBLE, ndCount, C1, C3, newColID )
		rc = data.AddColumn( False, False, FT_NODE, 0, "Node Z", FCT_DOUBLE, ndCount, C1, C4, newColID )
		If isThereOutput Then
			rc = data.AddColumn( False, False, FT_NODE, 0, "X Disp", FCT_DOUBLE, ndCount, C1, C5, newColID )
			rc = data.AddColumn( False, False, FT_NODE, 0, "Y Disp", FCT_DOUBLE, ndCount, C1, C6, newColID )
			rc = data.AddColumn( False, False, FT_NODE, 0, "Z Disp", FCT_DOUBLE, ndCount, C1, C7, newColID )
		End If
	End If

	Jumping_Out:
End Sub
