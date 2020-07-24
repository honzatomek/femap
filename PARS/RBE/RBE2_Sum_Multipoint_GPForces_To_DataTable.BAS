Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long, j As Long, k As Long, m As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim elset As Object
    Set elset = App.feSet()

    Dim el As Object
    Set el = App.feElem()

    Dim nd As Object
    Set nd = App.feNode()

    Dim osets As Object
    Set osets = App.feSet()

    Dim oset As Object
    Set oset = App.feOutputSet()

    Dim ovec As Object
    Set ovec = App.feOutput()

    Dim s As Object
    Set s = App.feSet()

    Dim DT As Object
    Set DT = App.feDataTable()

    Dim vecIDs() As Long
    Dim vecTit(5) As String
    vecTit(0) = "T1 MultiPoint GPForce"
    vecTit(1) = "T2 MultiPoint GPForce"
    vecTit(2) = "T3 MultiPoint GPForce"
    vecTit(3) = "R1 MultiPoint GPMoment"
    vecTit(4) = "R2 MultiPoint GPMoment"
    vecTit(5) = "R3 MultiPoint GPMoment"

    j = -1

    rc = osets.SelectMultiID(FT_OUT_CASE, 1, "Select Output Sets to Process:")
    If rc <> -1 Then
    	er = 1
    	GoTo Cleanup
    Else
    	ReDim vecIDs((osets.Count() * 6) - 1)
    	rc = osets.Reset()
    	Do While osets.Next() <> 0
    		j = j + 1
    		rc = oset.Get(osets.CurrentID())
    		rc = s.Clear()
    		rc = oset.OutputVectors(s.ID)
    		rc = s.Reset()
			For i = 0 To 5 Step 1
    			vecIDs(j*6 + i) = 0
			Next
    		i = 0
    		Do While s.Next() <> 0
				Set ovec = oset.Vector(s.CurrentID)
				If InStr(1, ovec.title, vecTit(i)) <> 0 Then
					vecIDs(j * 6 + i) = ovec.ID
					i = i + 1
					If i = 6 Then Exit Do
				End If
    		Loop
    	Loop
    End If

    Dim forces() As Double
    Dim ndIDs() As Long

    Dim ndcount As Long
    Dim nds
    Dim face
    Dim weight
    Dim DOF

    rc = elset.Select(8, True, "Select RB Spider to Sum Forces to Data Table:")
    If rc <> -1 Then
    	er = 2
    	GoTo Cleanup
    End If

	Dim colIDs() As Long
	ReDim colIDs(osets.Count() * elset.Count() - 1)
	Dim colX() As Double
	Dim colY() As Double
	Dim colZ() As Double
	Dim colSetIDs() As Long
	Dim colSetTitles() As String
	Dim colT1() As Double
	Dim colT2() As Double
	Dim colT3() As Double
	Dim colR1() As Double
	Dim colR2() As Double
	Dim colR3() As Double
	Dim rowID() As Long

	ReDim colX(UBound(colIDs))
	ReDim colY(UBound(colIDs))
	ReDim colZ(UBound(colIDs))
	ReDim colSetIDs(UBound(colIDs))
	ReDim colSetTitles(UBound(colIDs))
	ReDim colT1(UBound(colIDs))
	ReDim colT2(UBound(colIDs))
	ReDim colT3(UBound(colIDs))
	ReDim colR1(UBound(colIDs))
	ReDim colR2(UBound(colIDs))
	ReDim colR3(UBound(colIDs))
	ReDim rowID(UBound(colIDs))

	For i = 1 To UBound(colIDs) + 1
		rowID(i - 1) = i
	Next

	Dim row(11)
	Dim x As Double
	Dim y As Double
	Dim z As Double

	rc = osets.Reset()
	For i = 0 To osets.Count() - 1 Step 1
		rc = oset.Get(osets.Next())
		rc = elset.Reset()
		For j = 0 To elset.Count() - 1 Step 1
			rc = el.Get(elset.Next())
			row(0) = el.Node(0)
			rc = nd.Get(row(0))
			row(1) = nd.x
			row(2) = nd.y
			row(3) = nd.z
			row(4) = oset.ID
			row(5) = oset.title
			row(6) = 0
			row(7) = 0
			row(8) = 0
			row(9) = 0
			row(10) = 0
			row(11) = 0
			If (el.type() = 29) And (el.topology = 13) And (Not el.RigidInterpolate) Then
				rc = el.GetNodeList(0, ndcount, nds, face, weight, DOF)
				Set ovec = oset.Vector(vecIDs(i * 6))
				For k = 0 To ndcount - 1 Step 1
					row(6) = row(6) + ovec.Value(nds(k))
					rc = nd.Get(nds(k))
					x = nd.x - row(1)
					y = nd.y - row(2)
					z = nd.z - row(3)
					row(10) = row(10) + ovec.Value(nds(k)) * z
					Debug.Print nd.ID & vbTab & ovec.Value(nds(k)) & "*" & z & "=" & ovec.Value(nds(k)) * z
					row(11) = row(11) + ovec.Value(nds(k)) * (-y)
				Next
				Set ovec = oset.Vector(vecIDs(i * 6 + 1))
				For k = 0 To ndcount - 1 Step 1
					row(7) = row(7) + ovec.Value(nds(k))
					rc = nd.Get(nds(k))
					x = nd.x - row(1)
					y = nd.y - row(2)
					z = nd.z - row(3)
					row(9) = row(9) + ovec.Value(nds(k)) * (-z)
					row(11) = row(11) + ovec.Value(nds(k)) * x
				Next
				Set ovec = oset.Vector(vecIDs(i * 6 + 2))
				For k = 0 To ndcount - 1 Step 1
					row(8) = row(8) + ovec.Value(nds(k))
					rc = nd.Get(nds(k))
					x = nd.x - row(1)
					y = nd.y - row(2)
					z = nd.z - row(3)
					row(9) = row(9) + ovec.Value(nds(k)) * y
					row(10) = row(10) + ovec.Value(nds(k)) * (-x)
					Debug.Print nd.ID & vbTab & ovec.Value(nds(k)) & "*" & -x & "=" & ovec.Value(nds(k)) * (-x)
				Next
				Set ovec = oset.Vector(vecIDs(i * 6 + 3))
				For k = 0 To ndcount - 1 Step 1
					row(9) = row(9) + ovec.Value(nds(k))
				Next
				Set ovec = oset.Vector(vecIDs(i * 6 + 4))
				For k = 0 To ndcount - 1 Step 1
					row(10) = row(10) + ovec.Value(nds(k))
				Next
				Set ovec = oset.Vector(vecIDs(i * 6 + 5))
				For k = 0 To ndcount - 1 Step 1
					row(11) = row(11) + ovec.Value(nds(k))
				Next
			End If
			colIDs(i * elset.Count() + j) = row(0)
			colX(i * elset.Count() + j) = row(1)
			colY(i * elset.Count() + j) = row(2)
			colZ(i * elset.Count() + j) = row(3)
			colSetIDs(i * elset.Count() + j) = row(4)
			colSetTitles(i * elset.Count() + j) = row(5)
			colT1(i * elset.Count() + j) = row(6)
			colT2(i * elset.Count() + j) = row(7)
			colT3(i * elset.Count() + j) = row(8)
			colR1(i * elset.Count() + j) = row(9)
			colR2(i * elset.Count() + j) = row(10)
			colR3(i * elset.Count() + j) = row(11)
		Next
	Next

	Dim colID(11) As Long

	DT.Locked = False
	DT.Visible = True

	rc = DT.Clear()

	DT.type = 7
	rc = DT.AddColumn(False, False, FT_NODE, 0, "Node ID", 1, UBound(colIDs) + 1, rowID, colIDs, colID(0))
	rc = DT.AddColumn(False, False, FT_NODE, 0, "X", 2, UBound(colIDs) + 1, rowID, colX, colID(1))
	rc = DT.AddColumn(False, False, FT_NODE, 0, "Y", 2, UBound(colIDs) + 1, rowID, colY, colID(2))
	rc = DT.AddColumn(False, False, FT_NODE, 0, "Z", 2, UBound(colIDs) + 1, rowID, colZ, colID(3))
	rc = DT.AddColumn(False, False, FT_NODE, 0, "Set Value", 1, UBound(colIDs) + 1, rowID, colSetIDs, colID(4))
	rc = DT.AddColumn(False, False, FT_NODE, 0, "Set Title", 3, UBound(colIDs) + 1, rowID, colSetTitles, colID(5))
	rc = DT.AddColumn(False, False, FT_NODE, 0, vecTit(0), 2, UBound(colIDs) + 1, rowID, colT1, colID(6))
	rc = DT.AddColumn(False, False, FT_NODE, 0, vecTit(1), 2, UBound(colIDs) + 1, rowID, colT2, colID(7))
	rc = DT.AddColumn(False, False, FT_NODE, 0, vecTit(2), 2, UBound(colIDs) + 1, rowID, colT3, colID(8))
	rc = DT.AddColumn(False, False, FT_NODE, 0, vecTit(3), 2, UBound(colIDs) + 1, rowID, colR1, colID(9))
	rc = DT.AddColumn(False, False, FT_NODE, 0, vecTit(4), 2, UBound(colIDs) + 1, rowID, colR2, colID(10))
	rc = DT.AddColumn(False, False, FT_NODE, 0, vecTit(5), 2, UBound(colIDs) + 1, rowID, colR3, colID(11))

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set DT = Nothing
	Set ovec = Nothing
	Set oset = Nothing
	Set nd = Nothing
	Set el = Nothing
	Set elset = Nothing
	Set osets = Nothing
	Set App = Nothing
End Sub
