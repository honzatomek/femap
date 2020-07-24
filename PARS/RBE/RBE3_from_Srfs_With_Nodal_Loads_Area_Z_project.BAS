Sub Main
    Dim App As femap.model
    Set App = feFemap()

	Dim srfset As Object
	Set srfset = App.feSet()
	Dim srfcount As Long
	Dim vDOF As Variant
	Dim srfID As Long
	Dim srf As Object
	Set srf = App.feSurface
	'Dim srf As App.feSurface
	'Dim crvs As App.feSet
	'Dim crv As App.feCurve
	'Dim pts As App.feSet
	'Dim SPt As App.fePoint
	'Dim nodes() As App.feNode
	'Dim midnode As App.feNode

	Dim subdiv As Long
	Dim rc As Long
	Dim i As Long
	Dim numpoints As Long
	Dim points As Variant
	Dim midXYZ As Variant
	Dim nd As Object
	Set	nd = App.feNode()
	Dim ndset As Object
	Set ndset = App.feSet()
	Dim ndID As Long
	Dim pt As Object
	Set pt = App.fePoint()
	Dim nds() As Long
	Dim rb3 As Object
	Set rb3 = App.feElem()
	Dim DOF() As Long
	Dim weight() As Double

	Dim LdSet As Object
	Set LdSet = App.feLoadSet()
	Dim LdSetID As Long

	Dim Ld As Object
	Set Ld = App.feLoadMesh
	Dim area As Double
	Dim normal As Variant

	Dim alpha As Double
	Dim z(3) As Double
	z(0) = 0.0
	z(1) = 0.0
	z(2) = -1.0

	LdSetID = LdSet.NextEmptyID()
	LdSet.title = "RBE3 - Area"
	rc = LdSet.Put(LdSetID)

	Dim eID As Long

	rc = srfset.Select(5, True, "Select surfaces to convert to RBE3 Elements:")

	If rc = -1 Then
		srfcount = srfset.Count()
		If srfcount > 0 Then
			rc = srfset.Reset()
			srfID = srfset.Next()
			Do While srfID <> 0
				rc = srf.Get(srfID)
				If rc <> -1 Then Exit Sub
				rc = srf.Points(FCC_BOTH, numpoints, points)
				If rc <> -1 Then Exit Sub
				rc = srf.cg(midXYZ)
				If rc <> -1 Then Exit Sub
				ReDim nds(numpoints)
				ReDim DOF(numpoints * 6)
				ReDim weight(numpoints)
				rc = srf.Area(area)

				For i = 0 To numpoints - 1
					ndID = nd.NextEmptyID()
					nds(i) = ndID
					rc = pt.Get(points(i))
					nd.x = pt.x
					nd.y = pt.y
					nd.z = pt.z
					DOF(i * 6) = 1
					DOF(i * 6 + 1) = 1
					DOF(i * 6 + 2) = 1
					DOF(i * 6 + 3) = 0
					DOF(i * 6 + 4) = 0
					DOF(i * 6 + 5) = 0
					weight(i) = 1
					rc = nd.Put(ndID)
				Next
				ndID = nd.NextEmptyID()
				nd.x = midXYZ(0)
				nd.y = midXYZ(1)
				nd.z = midXYZ(2)
				rc = nd.Put(ndID)

				rc = srf.NormalAtXYZ(1, midXYZ, normal)

				alpha = ArcSin(Abs(normal(0)*z(0) + normal(1)*z(1) + normal(2)*z(2)) / (Sqr(normal(0)^2+normal(1)^2+normal(2)^2) * Sqr(z(0)^2 + z(1)^2 + z(2)^2)))
				alpha = min(Abs(Pi()/2 - alpha), Abs(alpha))

				normal(0) = normal(0) * area
				normal(1) = normal(1) * area
				normal(2) = normal(2) * area

				eID = rb3.NextEmptyID()
				rb3.type = 29 'Rigid
				rb3.topology = 13 'Rigid
				rb3.Node(0) = ndID 'Independent Node
				rb3.RigidInterpolate = True 'Interpolation element
				rb3.release(0, 0) = 1
				rb3.release(0, 1) = 1
				rb3.release(0, 2) = 1
				rc = rb3.PutNodeList(0, numpoints, nds, Null, weight, DOF )
				rc = rb3.Put(eID)

				Ld.setID = LdSetID
				Ld.ID = Ld.NextEmptyID()
				Ld.type = 1
				Ld.meshID = ndID
				Ld.load(0) = 0
				Ld.load(1) = 0
				Ld.load(2) = - area * Cos(alpha)
				Ld.XOn = True
				Ld.YOn = True
				Ld.ZOn = True
				rc = Ld.Put(Ld.ID)

				srfID = srfset.Next()
			Loop
		End If
	End If

	'pts = srf.PointsAsSet(FCC_BOTH)
	'crvs = srf.CurvesAsSet(FCC_OFF)
	'midnode = srf.cg

    
End Sub

Function min(ByVal arg1 As Double, ByVal arg2 As Double) As Double
	If arg1 <= arg2 Then
		min = arg1
	Else
		min = arg2
	End If
End Function

Function Pi() As Double
	Pi = 3.1415926535897
End Function

Function Deg(ByVal arg As Double) As Double
	Deg = 180 * arg / Pi()
End Function

Function Rad(ByVal arg As Double) As Double
	Rad = Pi() * arg / 180
End Function

Function ArcSin(ByVal x As Double) As Double
	ArcSin =Atn(x/(Sqr(1-x^2)))
End Function
