Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim plBase As Variant
    Dim plNormal As Variant
    Dim plAxis As Variant

    Dim element1() As Variant
    Dim element2() As Variant
    Dim coord(2) As Double
    Dim A As Double

    Dim DOF(2) As Variant
    Dim Values(4) As Variant
    Dim func(4) As Variant

    For i = 0 To 4 Step 1
    	Values(i) = 0
    	func(i) = 0
    Next

    Dim tmp As Variant
    Dim tmp2 As Long

    Dim elset As Object
    Set elset = App.feSet()

    Dim el As Object
    Set el = App.feElem()

    Dim nd As Object
    Set nd = App.feNode()

    Dim ld As Object
    Set ld = App.feLoadMesh()

    'rc = ld.GetSearch(42, 23, 1)

    rc = elset.Select(FT_ELEM, True, "Select elements to load by pressure")
    If (elset.Count < 1) Or (rc = FE_FAIL) Or (rc = FE_CANCEL) Then
    	er = 1
    	GoTo Cleanup
    End If

    Do
		tmp = InputBox("Input face number to be loaded:" & vbLf & "values from 1 to 6", "Pressure by projected area:", "1")
		If IsNumeric(tmp) Then
			tmp2 = CLng(tmp)
			If (tmp2 < 7) And (tmp2 > 0) Then Exit Do
		ElseIf tmp = "" Then
			er = 2
			GoTo Cleanup
		End If
    Loop

	rc = App.fePlanePick("Pick plane to project onto:", plBase, plNormal, plAxis)
	If rc = FE_CANCEL Then
		er = 3
		GoTo Cleanup
	End If

	ld.setID = App.Info_ActiveID(FT_LOAD_DIR)
	ld.flag(1) = True
	ld.fluxdir(0) = plNormal(0)
	ld.fluxdir(1) = plNormal(1)
	ld.fluxdir(2) = plNormal(2)
	DOF(0) = tmp2
	DOF(1) = 0
	DOF(2) = 0

	rc = elset.Reset()
	Do While elset.Next() <> 0
		rc = el.Get(elset.CurrentID())
		If ((el.type = 11) Or (el.type = 17)) And ((el.topology = 2) Or (el.topology = 4)) Then
			For i = 0 To 5 Step 1
				If el.Node(i) = 0 Then Exit For
				rc = nd.Get(el.Node(i))
				ReDim Preserve element1(i)
				coord(0) = nd.x
				coord(1) = nd.y
				coord(2) = nd.z
				element1(i) = coord
			Next
			element2 = element1
			For i = 0 To UBound(element1) Step 1
				rc = App.feCoordOntoPlane(element1(i), plNormal, plBase, plNormal, element2(i))
			Next
			A = Area(element2)
			Values(0) = A
			rc = ld.Add(-el.ID, 42, 0, DOF, Values, func)
		End If

	Loop


Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub


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

	Area2D = 0
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
	Dim magnitude As Double

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
	det2(1, 2) = B(2)
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
Function Area(ByRef element() As Variant) As Double
	Dim poly As Variant
	Dim prod As Variant
	Dim total(2) As Double
	Dim vi1(2) As Double
	Dim vi2(2) As Double
	Dim result As Variant
	Dim i As Long

	Dim ndID As Long
	Dim rc As Long
	Dim ndcount As Long
	ndcount = UBound(element) + 1
	ReDim poly(ndcount - 1, 2)
	Dim A(2) As Double, B(2) As Double, c(2) As Double

	 If ndcount < 3 Then 'Not a plane - no Area
        Area = 0
        Exit Function
    End If

	For i = 0 To ndcount - 1 Step 1
		poly(i, 0) = element(i)(0)
		poly(i, 1) = element(i)(1)
		poly(i, 2) = element(i)(2)
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
    result = Dot(total, GetNormal(element))
    Area = Abs(result/2)
End Function

Function length(ByVal v() As Variant) As Double
	length = Sqr(v(0)^2 + v(1)^2 + v(2)^2)
End Function

Function GetNormal(ByRef element() As Variant) As Variant
	Dim rc As Long
	Dim ndcount As Long
	ndcount = UBound(element) + 1
	Dim ndIDs() As Long
	Dim nd1(2) As Double
	Dim nd2(2) As Double
	Dim nd3(2) As Double
	Dim vect1(2) As Double
	Dim vect2(2) As Double
	Dim cosa As Double
	Dim normallength As Double
	Dim i As Long, j As Long, k As Long

	If ndcount < 3 Then
		GetNormal = Array(0, 0, 0)
		Exit Function
	End If

	For i = 0 To ndcount - 3 Step 1
		nd1(0) = element(i)(0)
		nd1(1) = element(i)(1)
		nd1(2) = element(i)(2)
		For j = i + 1 To ndcount - 2 Step 1
			nd2(0) = element(j)(0)
			nd2(1) = element(j)(1)
			nd2(2) = element(j)(2)

			vect1(0) = nd1(0) - nd2(0)
			vect1(1) = nd1(1) - nd2(1)
			vect1(2) = nd1(2) - nd2(2)
			For k = j + 1 To ndcount - 1 Step 1
				nd3(0) = element(k)(0)
				nd3(1) = element(k)(1)
				nd3(2) = element(k)(2)

				vect2(0) = nd1(0) - nd3(0)
				vect2(1) = nd1(1) - nd3(1)
				vect2(2) = nd1(2) - nd3(2)
				cosa = (vect1(0) * vect2(0) + vect1(1) * vect2(1) + vect1(2) * vect2(2))/(Sqr(vect1(0)^2 + vect1(1)^2 + vect1(2)^2) * Sqr(vect2(0)^2 + vect2(1)^2 + vect2(2)^2))
				If Abs(Round(cosa, 3)) <> 1 Then
					GetNormal = unit_normal(nd1, nd2, nd3)
					Exit Function
				End If
			Next
		Next
	Next
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
