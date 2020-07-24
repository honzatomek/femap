Option Explicit
'RBE3_from_Nodes_With_Nodal_Loads_Area

'Creates a rigid interpolation spider RBE3 from selected nodes,
'computes the area encompassed by all Dependent Nodes of the RBE3 Element
'and creates a Load in selected Load Case in a selected direction in the Independent node

Sub Main
	Begin Dialog UserDialog 300,105,"Force Direction",.Settings ' %GRID:10,7,1,1
		GroupBox 10,7,280,63,"Select direction:",.GroupBox1
		OKButton 110,77,80,21
		CancelButton 200,77,90,21
		CheckBox 20,28,50,14,"+X",.px
		CheckBox 20,49,50,14,"-X",.mx
		CheckBox 80,28,50,14,"+Y",.py
		CheckBox 80,49,50,14,"-Y",.my
		CheckBox 140,28,50,14,"+Z",.pz
		CheckBox 140,49,50,14,"-Z",.mz
		CheckBox 200,28,80,14,"Normal",.normal
	End Dialog
	Dim dlg As UserDialog
	dlg.px = 0
	dlg.mx = 0
	dlg.py = 0
	dlg.mx = 0
	dlg.pz = 0
	dlg.mz = 1
	dlg.normal = 0
	If Dialog(dlg) = 0 Then Exit Sub

	'Debug.Print dlg.px - dlg.mx & ", " & dlg.py - dlg.my & ", " & dlg.pz - dlg.mz

	Dim App As femap.model
    Set App = feFemap()
    Dim v As Object
    Set v = App.feView()
    Dim rc As Long
    Dim settings As Long

    On Error GoTo ExitHere
    rc = v.Get(0)

	Dim LdSet As Object
	'Set LdSet = App.feLoadSet()
	Set LdSet = App.feSet()
	Dim LdSetID As Long

	Dim Ld As Object
	Set Ld = App.feLoadMesh()
	Dim force As Variant
	Dim A As Double

	Dim ndset As Object
	Set ndset = App.feSet()
	Dim midnd As Object
	Set midnd = App.feNode()
	Dim midndID As Long
	Dim MidXYZ As Variant
	Dim nds() As Long
	Dim nd As Object
	Set nd = App.feNode()

	Dim el As Object
	Set el = App.feElem()
	Dim DOF() As Long
	Dim weight() As Double
	Dim ndcount As Long

	rc = LdSet.SelectID(12, "Select ID of the Load Set for nodal forces (Cancel = New LS):", LdSetID)
	If rc <> - 1 Then
		Set LdSet = App.feLoadSet
		LdSetID = LdSet.NextEmptyID()
		LdSet.title = "LC" & LdSetID & " - RBE3 - area "
		rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")
		If rc <> -1 Then GoTo ExitHere
		Call LdSet.Put(LdSetID)
	Else
		rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")
	End If
	App.Info_ActiveID(FT_LOAD_DIR) = LdSetID

	Do While rc = -1
		rc = ndset.Reset()
		ndcount = ndset.Count()
		elID = el.NextEmptyID()
		ReDim DOF(ndcount * 6 - 1)
		ReDim weight(ndcount - 1)
		ReDim nds(ndcount - 1)
		For i = 0 To ndcount - 1
			nds(i) = ndset.Next()
			DOF(i * 6) = 1
			DOF(i * 6 + 1) = 1
			DOF(i * 6 + 2) = 1
			DOF(i * 6 + 3) = 0
			DOF(i * 6 + 4) = 0
			DOF(i * 6 + 5) = 0
			weight(i) = 1.0
		Next
		midndID = midnd.NextEmptyID()
		MidXYZ = Mid(ndset)
		midnd.x = MidXYZ(0)
		midnd.y = MidXYZ(1)
		midnd.z = MidXYZ(2)
		rc = midnd.Put(midndID)

		el.ID = elID
		el.type = 29 'Rigid
		el.topology = 13 'Rigid
		el.Node(0) = midndID 'Independent Node
		el.RigidInterpolate = True 'Interpolation element
		el.release(0, 0) = 1
		el.release(0, 1) = 1
		el.release(0, 2) = 1
		rc = el.PutNodeList(0, ndcount, nds, Null, weight, DOF )
		rc = el.Put(eID)

		If dlg.normal = 1 Then
			force = GetNormal(ndset)
		Else
			force = unit(Array(dlg.px - dlg.mx, dlg.py - dlg.my, dlg.pz - dlg.mz))
		End If
		A = Abs(Area(ndset))

		Ld.setID = LdSetID
		Ld.ID = Ld.NextEmptyID()
		Ld.type = 1
		Ld.meshID = midndID
		Ld.load(0) = force(0) * A
		Ld.load(1) = force(1) * A
		Ld.load(2) = force(2) * A
		Ld.XOn = True
		Ld.YOn = True
		Ld.ZOn = True
		rc = Ld.Put(Ld.ID)

		App.feViewRegenerate(0)
		'Call v.Redraw()
		'Call v.Regenerate()

		rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")
	Loop

ExitHere:
	Set el = Nothing
	Set v = Nothing
	Set Ld = Nothing
	Set nd = Nothing
	Set ndset = Nothing
	Set LdSet = Nothing
	Set App = Nothing
End Sub

Rem See DialogFunc help topic for more information.
Private Function settings(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		'Debug.Print DlgItem$ & " " & SuppValue?
		Select Case DlgItem$
		Case "px"
			Call DlgValue(9, 0)
			Call DlgValue(4, 0)
		Case "mx"
			Call DlgValue(9, 0)
			Call DlgValue(3, 0)
		Case "py"
			Call DlgValue(9, 0)
			Call DlgValue(6, 0)
		Case "my"
			Call DlgValue(9, 0)
			Call DlgValue(5, 0)
		Case "pz"
			Call DlgValue(9, 0)
			Call DlgValue(8, 0)
		Case "mz"
			Call DlgValue(9, 0)
			Call DlgValue(7, 0)
		Case "normal"
			If SuppValue? = 1 Then
				Call DlgValue(3, 0)
				Call DlgValue(4, 0)
				Call DlgValue(5, 0)
				Call DlgValue(6, 0)
				Call DlgValue(7, 0)
				Call DlgValue(8, 0)
			End If
		End Select
		Rem Settings = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : Settings = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function

Function Distance(ByVal pl() As Double, ByVal pt() As Double) As Double
	Distance =(pl(0) * pt(0) + pl(1) * pt(1) + pl(2) * pt(2) + pl(3))/Sqr(pl(0)^2 + pl(1)^2 + pl(2)^2)
End Function

Function Mid(ByRef ndset As Object) As Variant
	Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long

	Dim nd As Object
	Set nd = App.feNode()
	Dim ndID As Long
	Dim ndcount As Long
	Dim i As Long
	Dim midnd As Variant
	Dim points() As Double

	ndcount = ndset.Count()
	ReDim points(ndcount - 1, 2)
	rc = ndset.Reset()

	For i = 0 To ndcount - 1 Step 1
		ndID = ndset.Next()
		rc = nd.Get(ndID)
		points(i, 0) = nd.x()
		points(i, 1) = nd.y()
		points(i, 2) = nd.z()
	Next
	Call SortVertices(points)
	midnd = AreaCentroid(points)

	Set nd = Nothing
	Mid = midnd
End Function

Function Area2D(ByRef pts() As Double) As Double
	Dim i As Long, j As Long
	Dim A As Double

	A = 0
	For i = 0 To UBound(pts) Step 1
		If i = UBound(pts) Then
			j = 0
		Else
			j = i + 1
		End If
		A = A + (pts(i, 0) * pts(j, 1) - pts(j, 0) * pts(i, 1))
		'Debug.Print "(" & pts(i, 0) & " * " & pts(j, 1) & " - " & pts(j, 0) & " * " & pts(i, 1) & ") = " & (pts(i, 0) * pts(j, 1) - pts(j, 0) * pts(i, 1))
		'Debug.Print A
	Next
	A = A/2
	Area2D = A
End Function

Function Centroid2D(ByVal pts() As Double) As Variant
	Dim Area As Double
	Dim i As Long, j As Long
	Dim c(1) As Double
	Dim x As Double, y As Double

	x = 0
	y = 0

	For i = 0 To UBound(pts) Step 1
		If pts(i, 0) < x Then x = pts(i, 0)
		If pts(i, 1) < y Then y = pts(i, 1)
	Next

	For i = 0 To UBound(pts) Step 1
		pts(i, 0) = pts(i, 0) - x
		pts(i, 1) = pts(i, 1) - y
		'Debug.Print "i[x , y] = [" & pts(i,0) & " , " & pts(i,1) & "]"
	Next

	Area = Area2D(pts)
	'Debug.Print "Area = " & Area

	c(0) = 0
	c(1) = 0
	For i = 0 To UBound(pts) Step 1
		If i = UBound(pts) Then
			j = 0
		Else
			j = i + 1
		End If
		c(0) = c(0) + (pts(i, 0) + pts(j, 0))*(pts(i, 0)*pts(j, 1) - pts(j, 0) * pts(i, 1))
		c(1) = c(1) + (pts(i, 1) + pts(j, 1))*(pts(i, 0)*pts(j, 1) - pts(j, 0) * pts(i, 1))
		'Debug.Print "c[x , y] = [" & c(0) & " , " & c(1) & "]"
	Next
	c(0) = c(0)/(6 * Area) + x
	c(1) = c(1)/(6 * Area) + y
	'Debug.Print "c[x , y] = [" & c(0) & " , " & c(1) & "]"
	Centroid2D = c
End Function

Function AreaCentroid(ByRef pts() As Double) As Variant
Dim v1(2) As Double, v2(2) As Double
Dim num As Long
Dim pts2D() As Double
Dim cxy As Double, cxz As Double, cyz As Double
Dim i As Long, j As Long, k As Long
Dim tmp As Double
Dim C2D As Variant
Dim ACen(2) As Double

'For i = 0 To UBound(pts)
'	Debug.Print "[x, y, z] = [" & pts(i, 0) & " , " & pts(i,1) & " , " & pts(i,2) & "]"
'Next

num = UBound(pts)
ReDim pts2D(num, 1)

cxy = 0
cxz = 0
cyz = 0

For i = 0 To num Step 1
	If i = num - 1 Then
		j = i + 1
		k = 0
	ElseIf i = num Then
		j = 0
		k = 1
	Else
		j = i + 1
		k = i + 2
	End If
	v1(0) = pts(i, 0) - pts(j, 0)
	v1(1) = pts(i, 1) - pts(j, 1)
	v1(2) = pts(i, 2) - pts(j, 2)

	v2(0) = pts(k, 0) - pts(j, 0)
	v2(1) = pts(k, 1) - pts(j, 1)
	v2(2) = pts(k, 2) - pts(j, 2)

	tmp =Round(Sqr(Dot(cross(Array(v1(0), v1(1), 0), Array(v2(0), v2(1), 0)),cross(Array(v1(0), v1(1), 0), Array(v2(0), v2(1), 0)))), 5)
	cxy = MaxAbs(cxy, tmp)
	tmp = Round(Sqr(Dot(cross(Array(v1(0), 0, v1(2)), Array(v2(0), 0, v2(2))),cross(Array(v1(0), 0, v1(2)), Array(v2(0), 0, v2(2))))), 5)
	cxz = MaxAbs(cxz, tmp)
	tmp = Round(Sqr(Dot(cross(Array(0, v1(1), v1(2)), Array(0, v2(1), v2(2))),cross(Array(0, v1(1), v1(2)), Array(0, v2(1), v2(2))))), 5)
	cyz = MaxAbs(cyz, tmp)
	If (cxy * cxz <> 0) Or (cxy * cyz <> 0) Or (cxz * cyz <> 0) Then Exit For
Next

ACen(0) = 0
ACen(1) = 0
ACen(2) = 0

Select Case cxy
Case 0 'polygon perpendicular to xy
	Select Case cxz
	Case 0 'polygon perpendicular to xy and xz
		Select Case cyz
		Case 0 'polygon perpendicular to xy, xz and yz
			AreaCentroid = ACen
			Exit Function
		Case Else 'polygon perpendicular to xy, xz not to yz
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 1)
				pts2D(i, 1) = pts(i, 2)
			Next
			C2D = Centroid2D(pts2D)
			ACen(0) = pts(0, 0)
			ACen(1) = C2D(0)
			ACen(2) = C2D(1)
			AreaCentroid = ACen
			Exit Function
		End Select
	Case Else 'polygon perpendicular to xy, not to xz
		Select Case cyz
		Case 0  'polygon perpendicular to xy, yz, not to xz
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 0)
				pts2D(i, 1) = pts(i, 2)
			Next
			C2D = Centroid2D(pts2D)
			ACen(0) = C2D(0)
			ACen(1) =pts(0, 1)
			ACen(2) = C2D(1)
			AreaCentroid = ACen
			Exit Function
		Case Else 'polygon perpendicular to xy, not to xz and yz
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 0)
				pts2D(i, 1) = pts(i, 2)
			Next
			C2D = Centroid2D(pts2D)
			ACen(0) = C2D(0)
			ACen(2) = C2D(1)
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 1)
				pts2D(i, 1) = pts(i, 2)
			Next
			C2D = Centroid2D(pts2D)
			ACen(1) = C2D(0)
			AreaCentroid = ACen
			Exit Function
		End Select
	End Select
Case Else 'polygon not perpendicular to xy
	Select Case cxz
	Case 0 'polygon perpendicular to xz, not to xy
		Select Case cyz
		Case 0 'polygon perpendicular to xz, yz, not to xy
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 0)
				pts2D(i, 1) = pts(i, 1)
			Next
			C2D = Centroid2D(pts2D)
			ACen(0) = C2D(0)
			ACen(1) = C2D(1)
			ACen(2) = pts(0, 2)
			AreaCentroid = ACen
			Exit Function
		Case Else 'polygon perpendicular to xz, not to xy and yz
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 0)
				pts2D(i, 1) = pts(i, 1)
			Next
			C2D = Centroid2D(pts2D)
			ACen(0) = C2D(0)
			ACen(1) = C2D(1)
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 1)
				pts2D(i, 1) = pts(i, 2)
			Next
			C2D = Centroid2D(pts2D)
			ACen(2) = C2D(1)
			AreaCentroid = ACen
			Exit Function
		End Select
	Case Else 'polygon not perpendicular to xy and xz
		For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 0)
				pts2D(i, 1) = pts(i, 1)
			Next
			C2D = Centroid2D(pts2D)
			ACen(0) = C2D(0)
			ACen(1) = C2D(1)
			For i = 0 To num Step 1
				pts2D(i, 0) = pts(i, 0)
				pts2D(i, 1) = pts(i, 2)
			Next
			C2D = Centroid2D(pts2D)
			ACen(2) = C2D(1)
			AreaCentroid = ACen
			Exit Function
	End Select
End Select
End Function

Function MaxAbs(ByVal A As Double, ByVal B As Double) As Double
	If Abs(A) > Abs(B) Then
		MaxAbs = A
	Else
		MaxAbs = B
	End If
End Function


Function Det(ByVal A() As Double) As Double
	Det = A(0,0)*A(1,1)*A(2,2) + A(0,1)*A(1,2)*A(2,0) + A(0,2)*A(1,0)*A(2,1) - A(0,2)*A(1,1)*A(2,0) - A(0,1)*A(1,0)*A(2,2) - A(0,0)*A(1,2)*A(2,1)
End Function

'unit normal vector of plane defined by points a, b, And c
Function unit_normal(ByVal A() As Double, ByVal B() As Double, ByVal c() As Double) As Variant
	Dim det1(2, 2) As Double
	Dim det2(2, 2) As Double
	Dim det3(2, 2) As Double
	Dim x As Double, y As Double, z As Double

	det1(0, 0) = 1
	det1(0, 1) = A(1)
	det1(0, 2) = A(2)
	det1(1, 0) = 1
	det1(1, 1) = B(1)
	det1(1, 2) = B(2)
	det1(2, 0) = 1
	det1(2, 1) = c(1)
	det1(2, 2) = c(2)

	det2(0, 0) = A(0)
	det2(0, 1) = 1
	det2(0, 2) = A(2)
	det2(1, 0) = B(0)
	det2(1, 1) = 1
	det2(1, 2) = b(2)
	det2(2, 0) = c(0)
	det2(2, 1) = 1
	det2(2, 2) = c(2)

	det3(0, 0) = A(0)
	det3(0, 1) = A(1)
	det3(0, 2) = 1
	det3(1, 0) = B(0)
	det3(1, 1) = B(1)
	det3(1, 2) = 1
	det3(2, 0) = c(0)
	det3(2, 1) = c(1)
	det3(2, 2) = 1

    x = Det(det1)
    y = Det(det2)
    z = Det(det3)
    magnitude = (x^2 + y^2 + z^2)^(0.5)

    unit_normal = unit(Array(x, y, z))
End Function

'unitizing a vector
Function unit(ByVal vect() As Variant) As Variant
	Dim l As Double
	Dim tmp(2) As Double

	l = Sqr(vect(0)^2 + vect(1)^2 + vect(2)^2)

	tmp(0) = vect(0) / l
	tmp(1) = vect(1) / l
	tmp(2) = vect(2) / l

	unit = Array(0,0,0)

	unit = tmp
End Function


'dot product of vectors a And b
Function Dot(ByVal A As Variant, ByVal B As Variant) As Double
    Dot = A(0)*B(0) + A(1)*B(1) + A(2)*B(2)
End Function

'cross product of vectors a And b
Function cross(ByVal A As Variant, ByVal B As Variant) As Variant
	Dim x As Double, y As Double, z As Double

    x = A(1) * B(2) - A(2) * B(1)
    y = A(2) * B(0) - A(0) * B(2)
    z = A(0) * B(1) - A(1) * B(0)
    cross = Array(x, y, z)
End Function

'Area of polygon poly
Function Area(ByRef ndset As Object) As Double
	Dim poly As Variant
	Dim prod As Variant
	Dim total(2) As Double
	Dim vi1(2) As Double
	Dim vi2(2) As Double
	Dim result As Variant

	Dim App As femap.model
    Set App = feFemap()
	Dim nd As Object
	Set nd = App.feNode()
	Dim ndID As Long
	Dim rc As Long
	rc = ndset.Reset()
	Dim ndcount As Long
	ndcount = ndset.Count()
	ReDim poly(ndcount - 1, 2)
	Dim A(2) As Double, B(2) As Double, c(2) As Double

	 If ndcount < 3 Then 'Not a plane - no Area
        Area = 0
        Exit Function
    End If

	For i = 0 To ndcount - 1 Step 1
		ndID = ndset.Next()
		rc = nd.Get(ndID)
		poly(i, 0) = nd.x()
		poly(i, 1) = nd.y()
		poly(i, 2) = nd.z()
	Next

	Call SortVertices(poly)

    total(0) = 0
    total(1) = 0
    total(2) = 0

    For i = 0 To ndcount - 1 Step 1
        vi1(0) = poly(i, 0)
        vi1(1) = poly(i, 1)
        vi1(2) = poly(i, 2)
        If i = ndcount - 1 Then
            vi2(0) = poly(0, 0)
        	vi2(1) = poly(0, 1)
       		vi2(2) = poly(0, 2)
        Else
         	vi2(0) = poly(i + 1, 0)
        	vi2(1) = poly(i + 1, 1)
       		vi2(2) = poly(i + 1, 2)
        End If
        prod = cross(vi1, vi2)
        total(0) = total(0) + prod(0)
        total(1) = total(1) + prod(1)
        total(2) = total(2) + prod(2)
    Next
    result = Dot(total, GetNormal(ndset))
    Set nd = Nothing
    Set App = Nothing
    Area = Abs(result/2)
End Function

Function length(ByVal v() As Variant) As Double
	dist = Sqr(v(0)^2 + v(1)^2 + v(2)^2)
End Function


Function GetNormal(ByRef ndset As Object) As Variant
	Dim App As femap.model
	Set App = feFemap()
	Dim rc As Long
	Dim nd As Object
	Set nd = App.feNode()
	Dim ndcount As Long
	ndcount = ndset.Count()
	Dim ndIDs() As Long
	Dim nd1(2) As Double
	Dim nd2(2) As Double
	Dim nd3(2) As Double
	Dim vect1(2) As Double
	Dim vect2(2) As Double
	Dim cosa As Double
	Dim normallength As Double
	Dim i As Long, j As Long, k As Long

	If ndset.Count() < 3 Then
		GetNormal = Array(0, 0, 0)
		Set nd = Nothing
		Set App = Nothing
		Exit Function
	End If

	rc = ndset.Reset()
	ReDim ndIDs(ndcount - 1)
	For i = 0 To ndcount - 1 Step 1
		ndIDs(i) = ndset.Next()
	Next

	For i = 0 To ndcount - 3 Step 1
		rc = nd.Get(ndIDs(i))
		nd1(0) = nd.x()
		nd1(1) = nd.y()
		nd1(2) = nd.z()
		For j = i + 1 To ndcount - 2 Step 1
			rc = nd.Get(ndIDs(j))
			nd2(0) = nd.x()
			nd2(1) = nd.y()
			nd2(2) = nd.z()

			vect1(0) = nd1(0) - nd2(0)
			vect1(1) = nd1(1) - nd2(1)
			vect1(2) = nd1(2) - nd2(2)
			For k = j + 1 To ndcount - 1 Step 1
				rc = nd.Get(ndIDs(k))
				nd3(0) = nd.x()
				nd3(1) = nd.y()
				nd3(2) = nd.z()

				vect2(0) = nd1(0) - nd3(0)
				vect2(1) = nd1(1) - nd3(1)
				vect2(2) = nd1(2) - nd3(2)
				cosa = (vect1(0) * vect2(0) + vect1(1) * vect2(1) + vect1(2) * vect2(2))/(Sqr(vect1(0)^2 + vect1(1)^2 + vect1(2)^2) * Sqr(vect2(0)^2 + vect2(1)^2 + vect2(2)^2))
				If Abs(Round(cosa, 3)) <> 1 Then
					Set nd = Nothing
					Set App = Nothing
					GetNormal = unit_normal(nd1, nd2, nd3)
					Exit Function
				End If
			Next
		Next
	Next
	Set nd = Nothing
	Set App = Nothing
	GetNormal = Array(0, 0, 0)
End Function

Function PI() As Double
	PI = 3.14159265359
End Function

Function Atan2(ByVal y As Double, ByVal x As Double) As Variant
	If (x > 0) And (y >= 0) Then
		Atan2 = Atn(y/x)
	ElseIf x < 0 Then
		Atan2 = Atn(y/x) + PI()
	ElseIf (x > 0) And (y < 0) Then
		Atan2 = Atn(y/x) + 2* PI()
	ElseIf (x = 0) And (y > 0) Then
		Atan2 = PI()/2
	Else
		Atan2 = 3 * PI()/2
	End If
End Function


'Z RHINA:
Sub SortVertices(ByRef ptsXYZ As Variant)
	Dim midpt As Variant
	Dim tmp As Variant
	Dim i As Long, j As Long, k As Long
	Dim num As Long
	Dim intersect As Boolean
	Dim min As Variant
	Dim mina As Double
	Dim alpha() As Variant
	Dim v1(2) As Double, v2(2) As Double

	num = UBound(ptsXYZ)
	midpt = Array(0, 0, 0)
	min = Array(0,0,0)

	For i = 0 To num Step 1
		If i = 0 Then
			min(0) = ptsXYZ(i, 0)
			min(1) = ptsXYZ(i, 1)
			min(2) = ptsXYZ(i, 2)
		End If

		midpt(0) = midpt(0) + ptsXYZ(i, 0)
		midpt(1) = midpt(1) + ptsXYZ(i, 1)
		midpt(2) = midpt(2) + ptsXYZ(i, 2)

		If ptsXYZ(i, 0) < min(0) Then min(0) = ptsXYZ(i, 0)
		If ptsXYZ(i, 1) < min(0) Then min(1) = ptsXYZ(i, 1)
		If ptsXYZ(i, 2) < min(0) Then min(2) = ptsXYZ(i, 2)
	Next
	midpt(0) = midpt(0) / (num + 1) - min(0)
	midpt(1) = midpt(1) / (num + 1) - min(1)
	midpt(2) = midpt(2) / (num + 1) - min(2)

	For i = 0 To num Step 1
		ptsXYZ(i, 0) = ptsXYZ(i, 0) - min(0)
		ptsXYZ(i, 1) = ptsXYZ(i, 1) - min(1)
		ptsXYZ(i, 2) = ptsXYZ(i, 2) - min(2)
	Next

	ReDim alpha(num, 1)

	alpha(0, 0) = 0
	alpha(0, 1) = 0

	v1(0) = ptsXYZ(0, 0) - midpt(0)
	v1(1) = ptsXYZ(0, 1) - midpt(1)
	v1(2) = ptsXYZ(0, 2) - midpt(2)
	'Debug.Print "[v1x, v1y, v1z] = [" & v1(0) & ", " & v1(1) & ", " & v1(2) & "]"
	Dim n As Variant
	n = Array(0, 0, 0)
	Dim x As Double, y As Double
	Dim c As Variant

	For i = 1 To num Step 1
		If i = 1 Then
			j = 1
			Do While Round(Dot(cross(v1, v2), cross(v1,v2)),3) = 0
				j = j + 1
				v2(0) = ptsXYZ(j, 0) - midpt(0)
				v2(1) = ptsXYZ(j, 1) - midpt(1)
				v2(2) = ptsXYZ(j, 2) - midpt(2)
			Loop
			n = unit(cross(v1, v2))
		End If
		v2(0) = ptsXYZ(i, 0) - midpt(0)
		v2(1) = ptsXYZ(i, 1) - midpt(1)
		v2(2) = ptsXYZ(i, 2) - midpt(2)

		alpha(i, 0) = i
		c =cross(v1, v2)
		y = Dot(n, c)
		x =Dot(v1, v2)
		alpha(i, 1) = Atan2(y, x)
		'Debug.Print alpha(i, 1) & "sin = " & x & ", cos = " & y
	Next

	For i = 0 To num - 1
		mina = alpha(i, 1)
		For j = i + 1 To num
			If alpha(j, 1) < mina Then
				mina = alpha(j, 1)
				tmp = alpha(i, 0)
				alpha(i, 0) = alpha(j, 0)
				alpha(j, 0) = tmp

				tmp = alpha(i, 1)
				alpha(i, 1) = alpha(j, 1)
				alpha(j, 1) = tmp
				For k = 0 To 2 Step 1
					tmp = ptsXYZ(i, k)
					ptsXYZ(i, k) = ptsXYZ(j, k)
					ptsXYZ(j, k) = tmp
				Next
			End If
		Next
	Next

	For i = 0 To num Step 1
		ptsXYZ(i, 0) = ptsXYZ(i, 0) + min(0)
		ptsXYZ(i, 1) = ptsXYZ(i, 1) + min(1)
		ptsXYZ(i, 2) = ptsXYZ(i, 2) + min(2)
	Next
End Sub
