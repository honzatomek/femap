Attribute VB_Name = "Module1"
Option Explicit

'Will create a vector for  Von Mises Stress along beams

Sub Main
	Dim App As femap.model
	Set App = feFemap()
	Dim rc As Long

	On Error GoTo ExitHere

	Dim elset As Object
	Set elset = App.feSet()
	Dim el As Object
	Set el = App.feElem()
	Dim elID As Long

	Dim osset As Object
	Set osset = App.feSet()
	Dim os As Object
	Set os = App.feOutputSet()
	Dim o As Object
	Set o = App.feOutput()
	Dim oID1 As Long
	Dim oID2 As Long

	Dim bc As Object
	Set bc = App.feBeamCalculator()

	Dim ovIDA(5) As Long
	Dim ovIDB(5) As Long
	Dim i As Long, j As Long

	Dim dForcesA(5) As Double
	Dim dForcesB(5) As Double
	Dim enComponent As Variant
	Dim pnNumNodes As Long
	Dim pvCoordinates As Variant
	Dim pvStresses As Variant
	Dim pvMaxStressA As Variant
	Dim pvMaxStressB As Variant
	Dim pvMaxStressInfo As Variant
	Dim pvMinStress As Variant
	Dim pvMinStressInfo As Variant

	Dim vmStressA() As Double
	Dim vmStressB() As Double
	Dim elIDs() As Long

	Dim tmp As Variant


	rc = osset.Select(28, True, "Select Output sets to process")
	If (rc <> -1) Or (osset.Count() < 1) Then GoTo ExitHere

	rc = elset.Select(8, True, "Select Elements:")
	If (rc <> -1) Or (elset.Count() < 1) Then GoTo ExitHere

	rc = osset.Reset()
	rc = os.Get(osset.Next())

	ovIDA(0) = 9000030
	ovIDA(1) = 9000022
	ovIDA(2) = 9000024
	ovIDA(3) = 9000014
	ovIDA(4) = 9000016
	ovIDA(5) = 9000034

	ovIDB(0) = 9000031
	ovIDB(1) = 9000023
	ovIDB(2) = 9000025
	ovIDB(3) = 9000015
	ovIDB(4) = 9000017
	ovIDB(5) = 9000035

	ReDim vmStressA(elset.Count() - 1)
	ReDim vmStressB(elset.Count() - 1)
	ReDim elIDs(elset.Count() - 1)
	j = -1

	Dim outputs As Long
	Dim oput As Long
	Dim elements As Long
	Dim elmnt As Long

	oput = 0


	outputs = osset.Count()
	elements = elset.Count()

	Call App.feAppStatusShow(True, outputs * elements)

	rc = osset.Reset()
	Do While osset.Next() <> FE_FAIL
		oput = oput + 1

		rc = os.Get(osset.CurrentID())
		Call App.feAppMessage(FCM_NORMAL, "Processing VonMises Stress on Output ID " & os.ID & " - " & os.title)
		j = -1
		rc = elset.Reset()

		Call ReturnVecIDs(os.ID, ovIDA, "A")
		Call ReturnVecIDs(os.ID, ovIDB, "B")
		elmnt = 0

		Do While elset.Next() <> FE_FAIL
			elmnt = elmnt + 1
			j = j + 1
			elID = elset.CurrentID()
			elIDs(j) = elID
			rc = el.Get(elID)

			If el.type = 5 Then
				For i = 0 To 5 Step 1
					o.setID = os.ID()
					If ovIDA(i) <> 0 Then
						Set o = os.Vector(ovIDA(i))
						dForcesA(i) = o.Value(elID)
					Else
						dForcesA(i) = 0
					End If
					If ovIDB(i) <> 0 Then
						Set o = os.Vector(ovIDB(i))
						dForcesB(i) = o.Value(elID)
					Else
						dForcesB(i) = 0
					End If
				Next

				bc.Element = elID
				bc.Position = 0.0
				bc.MeshFactor = 1
				rc = bc.CalculateStressFromForces(dForcesA, dForcesB, 0, pnNumNodes, pvCoordinates, pvStresses, pvMaxStressA, pvMaxStressInfo, pvMinStress, pvMinStressInfo)
				vmStressA(j) = pvMaxStressA(0)

				bc.Element = elID
				bc.Position = 1.0
				bc.MeshFactor = 1
				rc = bc.CalculateStressFromForces(dForcesA, dForcesB, 0, pnNumNodes, pvCoordinates, pvStresses, pvMaxStressB, pvMaxStressInfo, pvMinStress, pvMinStressInfo)
				vmStressB(j) = -pvMaxStressB(0)
			End If

			Call App.feAppStatusUpdate( (oput - 1) * elements + elmnt )
		Loop

		o.setID = os.ID
		oID1 = VonMisesOutputVector(os.ID)
		If oID1 <> -1 Then
			oID2 = oID1 + 1
			rc = o.Delete(oID1)
			rc = o.Delete(oID2)
		Else
			rc = o.Last()
			oID1 = App.feMaxInt(o.ID + 1, 9000001)
			oID2 = oID1 + 1
		End If
		rc = o.InitScalarAtBeam(os.ID, oID1, oID2, "Max VonMises Stress", 2, 4, False, True)
		rc = o.PutScalarAtBeam(UBound(elIDs) + 1, elIDs, vmStressA,vmStressB)
		rc = o.Put(-1)

		Call App.feAppStatusUpdate( oput * elements )
	Loop



ExitHere:
	On Error Resume Next
	Call App.feAppStatusShow(False, 0)
	Set o = Nothing
	Set bc = Nothing
	Set os = Nothing
	Set el = Nothing
	Set elset = Nothing
	Set osset = Nothing
	Set App = Nothing
	On Error GoTo 0
End Sub

Private Sub ReturnVecIDs(ByVal OutSet As Long, ByRef ovIDs() As Long, ByVal e As String)
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long
	Dim o As Object
	Set o = App.feOutput()

	Dim s As String
	Dim arrs(5) As String

	arrs(0) = "Axial Force"
	arrs(1) = "Pl1 Shear Force"
	arrs(2) = "Pl2 Shear Force"
	arrs(3) = "Plane1 Moment"
	arrs(4) = "Plane2 Moment"
	arrs(5) = "Torque"

	For i = 0 To 5 Step 1
		o.setID = OutSet
		rc = o.Reset()
		s = "Beam End" & e & " " & arrs(i)
		ovIDs(i) = 0
		Do While o.Next() <> FE_FAIL
			If InStr(1, o.title, s) <> 0 Then
				ovIDs(i) = o.ID
				Exit Do
			End If
		Loop
	Next

	On Error Resume Next
	Set o = Nothing
	Set App = Nothing
	On Error GoTo 0
End Sub

Private Function VonMisesOutputVector(ByVal OutSet As Long) As Long
	Dim App As femap.model
	Set App = feFemap
	Dim rc As Long, i As Long
	Dim o As Object
	Set o = App.feOutput()

	Dim s As String
	Dim tmp As Long

	o.setID = OutSet
	rc = o.Reset()
	s = "End A Max VonMises Stress"
	Do While o.Next() <> FE_FAIL
		If InStr(1, o.title, s) <> 0 Then
			tmp = o.ID()
			Exit Do
		Else
			tmp = -1
		End If
	Loop

	VonMisesOutputVector = tmp

	On Error Resume Next
	Set o = Nothing
	Set App = Nothing
	On Error GoTo 0
End Function
