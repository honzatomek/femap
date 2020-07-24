Option Explicit

'Creates RBE3 Element from Beam elements end-nodes
Private er As Long
Const DBUG As Boolean = True

Sub Main()
	Dim App As femap.model
	Set App = feFemap()
	Dim rc As Long, i As Long, j As Long

	On Error GoTo Cleanup
	er = 0

	Dim elset As Object
	Set elset = App.feSet()

	Dim el As Object
	Set el = App.feElem()

	Dim ndset As Object
	Set ndset = App.feSet()

	Dim nd As Object
	Set nd = App.feNode()

	Dim numnode As Long
	Dim entID As Variant
	Dim xyz As Variant
	Dim lay As Variant
	Dim col As Variant
	Dim tp As Variant
	Dim defCSys As Variant
	Dim outCSys As Variant
	Dim permBC As Variant

	Dim weight() As Double
	Dim DOF() As Long
	Dim midnd(2) As Double


	rc = elset.Select(8, True, "Select elements for RBE3 Spider:")
	If (rc <> -1) Or (elset.Count < 1) Then
		er = 1
		GoTo cleanup
	End If

	rc = elset.Reset

	Dim els() As Long

	ReDim els(elset.Count() * 2 - 1)
	i = -2

	On Error GoTo 0

	For j = 0 To elset.Count() - 1 Step 1
	 	rc = el.Get(elset.Next())
		If el.type = 5 Then
			i = i + 2
			els(i) = el.Node(0)
			els(i + 1) = el.Node(1)
			'rc = ndset.Add(el.Node(0))
			'rc =ndset.Add(el.Node(1))
		End If
	Next

	rc = ndset.AddArray(i + 2, els)

	If ndset.Count < 2 Then
		er = 2
		GoTo Cleanup
	End If

	rc = ndset.Reset()

	rc = nd.GetAllArray(ndset.ID, numnode, entID, xyz, lay, col, tp, defCSys, outCSys, permBC)

	ReDim DOF(numnode * 6 - 1)
	ReDim weight(numnode - 1)

	midnd(0) = 0
	midnd(1) = 0
	midnd(2) = 0

	For i = 0 To ndset.Count() - 1 Step 1
		weight(i) = 1

		DOF(i*6) = 1
		DOF(i*6 + 1) = 1
		DOF(i*6 + 2) = 1
		DOF(i*6 + 3) = 0
		DOF(i*6 + 4) = 0
		DOF(i*6 + 5) = 0

		midnd(0) = midnd(0) + xyz(i * 3)
		midnd(1) = midnd(1) + xyz(i * 3 + 1)
		midnd(2) = midnd(2) + xyz(i * 3 + 2)
	Next

	midnd(0) = midnd(0) / numnode
	midnd(1) = midnd(1) / numnode
	midnd(2) = midnd(2) / numnode

	nd.ID = nd.NextEmptyID()
	nd.x = midnd(0)
	nd.y = midnd(1)
	nd.z = midnd(2)
	rc = nd.Put(nd.ID)

	el.ID = el.NextEmptyID()
	el.type = 29
	el.topology = 13
	el.RigidInterpolate = True
	el.Node(0) = nd.ID()
	el.release(0, 0) = 1
	el.release(0, 1) = 1
	el.release(0, 2) = 1
	rc = el.PutNodeList(0, numnode, entID, Null, weight, DOF)
	rc =el.Put(el.ID)
	If DBUG Then App.feAppMessage(FCM_NORMAL, "Created element: " & el.ID)

Cleanup:
	On Error Resume Next
		Set elset = Nothing
		If DBUG Then App.feAppMessage(FCM_NORMAL, "The script exited with code " & er)
		Set App = Nothing
End Sub
