Attribute VB_Name = "Module3"
Option Explicit

Sub Main
	Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long
    Dim i As Long

    On Error GoTo ExitHere

	Dim ndset As Object
	Set ndset = App.feSet()

	rc = ndset.AddAll(7)

	Dim ndID As Long
	Dim nd As Object
	Set nd = App.feNode()

	Dim oset As Object
	Set oset = App.feSet()

	Dim os As Object
	Set os = App.feOutputSet()

	Dim o As Object
	Set o = App.feOutput()

	rc = oset.Select(28, True, "Select Output Set to change to Seismic Load")
	If rc <> -1 Then GoTo ExitHere
	rc = oset.Reset()

	Dim tmp As Variant
	tmp = InputBox("Input Seismic Shear Coefficient:", "Simplified Seismicity", "0.065")
	If tmp = "" Then GoTo ExitHere
	Dim scof As Double
	If InStr(1, tmp, ".") <> 0 Then tmp = Replace(tmp, ".", ",")
	scof = CDbl(tmp)

	rc = o.GetFromSet(oset.Next(), 174)
	rc = ndset.Reset()
	ndID = ndset.Next()

	Dim sumM As Double
	Dim sumMZ As Double
	Dim zmin As Double

	zmin = 0

	Do While ndID <> 0
		rc = nd.Get(ndID)
		If nd.z() < zmin Then zmin = nd.z()
		ndID = ndset.Next()
	Loop

	rc = ndset.Reset()
	ndID = ndset.Next()

	Do While ndID <> 0
		rc = nd.Get(ndID)
		sumM = sumM + o.Value(ndID)
		sumMZ = sumMZ + o.Value(ndID) * (nd.z() - zmin)
		ndID = ndset.Next()
	Loop

	sumM = Abs(sumM)
	sumMZ = Abs(sumMZ)

	Dim LdSet As Object
	Set LdSet = App.feLoadSet()
	Dim LdSetID As Long
	Dim Ld As Object
	Set Ld = App.feLoadMesh()
	Dim vect(3, 1) As Long
	Dim nm(3) As String
	vect(0,0) = 1
	vect(0,1) = 0
	vect(1,0) = -1
	vect(1,1) = 0
	vect(2,0) = 0
	vect(2,1) = 1
	vect(3,0) = 0
	vect(3,1) = -1
	nm(0) ="[SEI1]..+X"
	nm(1) = "[SEI2]..-X"
	nm(2) = "[SEI3]..+Y"
	nm(3) = "[SEI4]..-Y"


	For i = 0 To 3 Step 1
		LdSetID = LdSet.NextEmptyID()
		LdSet.title = nm(i)
		Call LdSet.Put(LdSetID)

		rc = ndset.Reset()
		ndID = ndset.Next()

		Do While ndID <> 0
			rc = nd.Get(ndID)
			If (o.Value(ndID) <> 0) Then
				Ld.setID = LdSetID
				Ld.ID = Ld.NextEmptyID()
				Ld.type = 1
				Ld.meshID = ndID
				Ld.load(0) = (sumM * Abs(o.Value(ndID) * (nd.z() - zmin)) / sumMZ) * vect(i, 0) * scof
				Ld.load(1) = (sumM * Abs(o.Value(ndID) * (nd.z() - zmin)) / sumMZ) * vect(i, 1) * scof
				Ld.load(2) = 0
				Ld.XOn = True
				Ld.YOn = True
				Ld.ZOn = True
				rc = Ld.Put(Ld.ID)
			End If
			ndID = ndset.Next()
		Loop
	Next

ExitHere:
	On Error Resume Next
	Set Ld = Nothing
	Set LdSet = Nothing
	Set o = Nothing
	Set oset = Nothing
	Set nd = Nothing
	Set ndset = Nothing
	Set App = Nothing
End Sub
