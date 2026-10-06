' =============================================================================
' Description: Computes RBE3 polygon area on a best-fit plane and applies an area-magnitude nodal force at the reference node.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Creates nodal forces at the reference node of selected RBE3 elements.
' Force magnitude equals the polygon area enclosed by the listed RBE3 nodes.
' The polygon is projected to its best-fit plane; area is evaluated with the
' signed 2D cross-product sum. The best-fit normal comes from the analytical
' smallest eigenvalue/eigenvector of the 3x3 covariance matrix.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FT_LOAD_DIR As Long = 12
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13
Const C_LOAD_FORCE As Long = 1

Sub Main
    Begin Dialog UserDialog 300,105,"Force Direction",.Settings
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
    dlg.my = 0
    dlg.pz = 0
    dlg.mz = 1
    dlg.normal = 0
    If Dialog(dlg) = 0 Then Exit Sub

    Dim App As Object
    Set App = feGetObject()

    Dim elset As Object
    Dim el As Object
    Dim nd As Object
    Dim ldsel As Object
    Dim ldset As Object
    Dim ld As Object
    Dim rc As Long
    Dim loadSetID As Long
    Dim newLoadSet As Boolean
    Dim elID As Long
    Dim ndcount As Long
    Dim nds As Variant
    Dim face As Variant
    Dim weight As Variant
    Dim dof As Variant
    Dim xyzX() As Double
    Dim xyzY() As Double
    Dim xyzZ() As Double
    Dim origin(2) As Double
    Dim normal(2) As Double
    Dim centroid(2) As Double
    Dim area As Double
    Dim i As Long
    Dim loadID As Long
    Dim fx As Double
    Dim fy As Double
    Dim fz As Double

    Set elset = App.feSet()
    Set el = App.feElem()
    Set nd = App.feNode()
    Set ldsel = App.feSet()
    Set ld = App.feLoadMesh()

    On Error GoTo ErrorHandler

    rc = ldsel.SelectID(C_FT_LOAD_DIR, "Select Load Set for nodal forces (Cancel = New Load Set):", loadSetID)
    If rc <> C_FE_OK Then
        Set ldset = App.feLoadSet()
        loadSetID = ldset.NextEmptyID()
        ldset.title = "LC" & CStr(loadSetID) & " - RBE3 - area"
        newLoadSet = True
    Else
        newLoadSet = False
    End If

SelectAgain:
    rc = elset.Select(C_FT_ELEM, True, "Select RBE3 elements for nodal area loads:")
    If rc <> C_FE_OK Then GoTo ExitHere
    If elset.Count() = 0 Then GoTo ExitHere

    If newLoadSet Then
        rc = ldset.Put(loadSetID)
        If rc <> C_FE_OK Then GoTo ExitHere
        newLoadSet = False
    End If
    App.Info_ActiveID(C_FT_LOAD_DIR) = loadSetID

    rc = elset.Reset()
    elID = elset.Next()
    Do While elID <> 0
        rc = el.Get(elID)
        If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) And el.RigidInterpolate Then
            rc = el.GetNodeList(0, ndcount, nds, face, weight, dof)
            If (rc = C_FE_OK) And (ndcount >= 3) Then
                ReDim xyzX(ndcount - 1)
                ReDim xyzY(ndcount - 1)
                ReDim xyzZ(ndcount - 1)

                For i = 0 To ndcount - 1
                    rc = nd.Get(nds(i))
                    If rc <> C_FE_OK Then GoTo NextElement
                    xyzX(i) = nd.x
                    xyzY(i) = nd.y
                    xyzZ(i) = nd.z
                Next

                If BestFitPlane(xyzX, xyzY, xyzZ, ndcount, origin, normal) Then
                    If PolygonAreaCentroid(xyzX, xyzY, xyzZ, ndcount, origin, normal, area, centroid) Then
                        If dlg.normal = 1 Then
                            fx = normal(0)
                            fy = normal(1)
                            fz = normal(2)
                        Else
                            fx = dlg.px - dlg.mx
                            fy = dlg.py - dlg.my
                            fz = dlg.pz - dlg.mz
                        End If

                        loadID = ld.NextEmptyID()
                        ld.setID = loadSetID
                        ld.ID = loadID
                        ld.type = C_LOAD_FORCE
                        ld.meshID = el.Node(0)
                        ld.load(0) = fx * Abs(area)
                        ld.load(1) = fy * Abs(area)
                        ld.load(2) = fz * Abs(area)
                        ld.XOn = True
                        ld.YOn = True
                        ld.ZOn = True
                        rc = ld.Put(loadID)
                    End If
                End If
            End If
        End If
NextElement:
        elID = elset.Next()
    Loop

    Call App.feViewRegenerate(0)
    GoTo SelectAgain

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set ld = Nothing
    Set ldset = Nothing
    Set ldsel = Nothing
    Set nd = Nothing
    Set el = Nothing
    Set elset = Nothing
    Set App = Nothing
End Sub

Private Function Settings(DlgItem, Action, SuppValue) As Boolean
    Settings = True
    Select Case Action
    Case 2
        Select Case DlgItem
        Case "OK", "Cancel", "Esc"
            Settings = False
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
            If SuppValue = 1 Then
                Call DlgValue(3, 0)
                Call DlgValue(4, 0)
                Call DlgValue(5, 0)
                Call DlgValue(6, 0)
                Call DlgValue(7, 0)
                Call DlgValue(8, 0)
            End If
        End Select
    End Select
End Function
Function BestFitPlane(ByRef xyzX() As Double, ByRef xyzY() As Double, ByRef xyzZ() As Double, ByVal npts As Long, ByRef origin() As Double, ByRef n() As Double) As Boolean
    Dim i As Long
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim c00 As Double
    Dim c01 As Double
    Dim c02 As Double
    Dim c11 As Double
    Dim c12 As Double
    Dim c22 As Double
    Dim cov(5) As Double
    Dim eigVal(2) As Double

    BestFitPlane = False

    origin(0) = 0#
    origin(1) = 0#
    origin(2) = 0#

    For i = 0 To npts - 1
        origin(0) = origin(0) + xyzX(i)
        origin(1) = origin(1) + xyzY(i)
        origin(2) = origin(2) + xyzZ(i)
    Next

    origin(0) = origin(0) / npts
    origin(1) = origin(1) / npts
    origin(2) = origin(2) / npts

    c00 = 0#
    c01 = 0#
    c02 = 0#
    c11 = 0#
    c12 = 0#
    c22 = 0#

    For i = 0 To npts - 1
        dx = xyzX(i) - origin(0)
        dy = xyzY(i) - origin(1)
        dz = xyzZ(i) - origin(2)

        c00 = c00 + dx * dx
        c01 = c01 + dx * dy
        c02 = c02 + dx * dz
        c11 = c11 + dy * dy
        c12 = c12 + dy * dz
        c22 = c22 + dz * dz
    Next

    cov(0) = c00
    cov(1) = c01
    cov(2) = c02
    cov(3) = c11
    cov(4) = c12
    cov(5) = c22

    If Not SmallestEigenvectorSym3(cov, eigVal, n) Then Exit Function

    ' eigVal(0)=smallest, eigVal(1)=middle, eigVal(2)=largest.
    ' A unique plane requires two independent in-plane directions.
    ' If the middle covariance eigenvalue vanishes relative to the largest,
    ' the points are effectively collinear.
    If eigVal(2) <= 0# Then Exit Function
    If eigVal(1) <= eigVal(2) * 0.0000000001 Then Exit Function

    Call OrientNormal(n)
    BestFitPlane = True
End Function


' Closed-form eigenvalues of a real symmetric 3x3 matrix and eigenvector
' belonging to the smallest eigenvalue.
'
' Matrix:
'   [ a00 a01 a02 ]
'   [ a01 a11 a12 ]
'   [ a02 a12 a22 ]
Function SmallestEigenvectorSym3(ByRef a() As Double, ByRef eigVal() As Double, ByRef v() As Double) As Boolean
    Dim a00 As Double
    Dim a01 As Double
    Dim a02 As Double
    Dim a11 As Double
    Dim a12 As Double
    Dim a22 As Double
    Dim q As Double
    Dim p2 As Double
    Dim p As Double
    Dim b00 As Double
    Dim b01 As Double
    Dim b02 As Double
    Dim b11 As Double
    Dim b12 As Double
    Dim b22 As Double
    Dim detB As Double
    Dim r As Double
    Dim phi As Double
    Dim row1(2) As Double
    Dim row2(2) As Double
    Dim row3(2) As Double
    Dim v12(2) As Double
    Dim v13(2) As Double
    Dim v23(2) As Double
    Dim n12 As Double
    Dim n13 As Double
    Dim n23 As Double

    SmallestEigenvectorSym3 = False

    a00 = a(0)
    a01 = a(1)
    a02 = a(2)
    a11 = a(3)
    a12 = a(4)
    a22 = a(5)

    q = (a00 + a11 + a22) / 3#

    p2 = (a00 - q) * (a00 - q)
    p2 = p2 + (a11 - q) * (a11 - q)
    p2 = p2 + (a22 - q) * (a22 - q)
    p2 = p2 + 2# * (a01 * a01 + a02 * a02 + a12 * a12)
    p2 = p2 / 6#

    If p2 <= 0# Then Exit Function
    p = Sqr(p2)

    b00 = (a00 - q) / p
    b01 = a01 / p
    b02 = a02 / p
    b11 = (a11 - q) / p
    b12 = a12 / p
    b22 = (a22 - q) / p

    detB = b00 * (b11 * b22 - b12 * b12)
    detB = detB - b01 * (b01 * b22 - b12 * b02)
    detB = detB + b02 * (b01 * b12 - b11 * b02)

    r = detB / 2#
    If r < -1# Then r = -1#
    If r > 1# Then r = 1#

    phi = AcosCompat(r) / 3#

    ' For phi in [0, pi/3], these are ordered max, mid, min.
    eigVal(2) = q + 2# * p * Cos(phi)
    eigVal(0) = q + 2# * p * Cos(phi + 2# * PIValue() / 3#)
    eigVal(1) = 3# * q - eigVal(2) - eigVal(0)

    ' Clamp tiny negative covariance eigenvalues caused by roundoff.
    If eigVal(2) < 0# And Abs(eigVal(2)) < p * 0.000000000001 Then eigVal(2) = 0#
    If eigVal(1) < 0# And Abs(eigVal(1)) < p * 0.000000000001 Then eigVal(1) = 0#
    If eigVal(0) < 0# And Abs(eigVal(0)) < p * 0.000000000001 Then eigVal(0) = 0#

    row1(0) = a00 - eigVal(0)
    row1(1) = a01
    row1(2) = a02

    row2(0) = a01
    row2(1) = a11 - eigVal(0)
    row2(2) = a12

    row3(0) = a02
    row3(1) = a12
    row3(2) = a22 - eigVal(0)

    Call Cross3(row1, row2, v12)
    Call Cross3(row1, row3, v13)
    Call Cross3(row2, row3, v23)

    n12 = Dot3(v12, v12)
    n13 = Dot3(v13, v13)
    n23 = Dot3(v23, v23)

    If n12 >= n13 And n12 >= n23 Then
        v(0) = v12(0)
        v(1) = v12(1)
        v(2) = v12(2)
    ElseIf n13 >= n23 Then
        v(0) = v13(0)
        v(1) = v13(1)
        v(2) = v13(2)
    Else
        v(0) = v23(0)
        v(1) = v23(1)
        v(2) = v23(2)
    End If

    If Not Normalize3(v) Then Exit Function

    SmallestEigenvectorSym3 = True
End Function


' Eigenvectors have arbitrary sign.  Make the result deterministic by making
' the largest-magnitude global component positive.
Sub OrientNormal(ByRef n() As Double)
    Dim ax As Double
    Dim ay As Double
    Dim az As Double
    Dim flip As Boolean

    ax = Abs(n(0))
    ay = Abs(n(1))
    az = Abs(n(2))
    flip = False

    If ax >= ay And ax >= az Then
        If n(0) < 0# Then flip = True
    ElseIf ay >= az Then
        If n(1) < 0# Then flip = True
    Else
        If n(2) < 0# Then flip = True
    End If

    If flip Then
        n(0) = -n(0)
        n(1) = -n(1)
        n(2) = -n(2)
    End If
End Sub


' Project points onto the best-fit plane, sort around the projected arithmetic
' centroid by polar angle, then compute signed polygon area and polygon centroid
' from one 2-D cross-product (shoelace) loop.
Function PolygonAreaCentroid(ByRef xyzX() As Double, ByRef xyzY() As Double, ByRef xyzZ() As Double, ByVal npts As Long, ByRef origin() As Double, ByRef n() As Double, ByRef area As Double, ByRef cen() As Double) As Boolean
    Dim e1(2) As Double
    Dim e2(2) As Double
    Dim r3(2) As Double
    Dim rp(2) As Double
    Dim u() As Double
    Dim v() As Double
    Dim angle() As Double
    Dim i As Long
    Dim j As Long
    Dim d As Double
    Dim d2 As Double
    Dim maxD2 As Double
    Dim sumCross As Double
    Dim cr As Double
    Dim cu As Double
    Dim cv As Double

    PolygonAreaCentroid = False
    area = 0#
    cen(0) = 0#
    cen(1) = 0#
    cen(2) = 0#

    ' Choose the farthest projected node vector as the local X axis.
    maxD2 = 0#
    For i = 0 To npts - 1
        r3(0) = xyzX(i) - origin(0)
        r3(1) = xyzY(i) - origin(1)
        r3(2) = xyzZ(i) - origin(2)

        d = Dot3(r3, n)
        rp(0) = r3(0) - d * n(0)
        rp(1) = r3(1) - d * n(1)
        rp(2) = r3(2) - d * n(2)
        d2 = Dot3(rp, rp)

        If d2 > maxD2 Then
            maxD2 = d2
            e1(0) = rp(0)
            e1(1) = rp(1)
            e1(2) = rp(2)
        End If
    Next

    If maxD2 <= 0# Then Exit Function
    If Not Normalize3(e1) Then Exit Function

    Call Cross3(n, e1, e2)
    If Not Normalize3(e2) Then Exit Function

    ReDim u(npts - 1)
    ReDim v(npts - 1)
    ReDim angle(npts - 1)

    ' origin is the arithmetic mean of the master nodes, so the projected
    ' arithmetic centroid is at local (0,0), apart from roundoff.
    For i = 0 To npts - 1
        r3(0) = xyzX(i) - origin(0)
        r3(1) = xyzY(i) - origin(1)
        r3(2) = xyzZ(i) - origin(2)

        u(i) = Dot3(r3, e1)
        v(i) = Dot3(r3, e2)
        angle(i) = Atan2Positive(v(i), u(i))
    Next

    Call Sort2DByAngle(u, v, angle, npts)

    sumCross = 0#
    cu = 0#
    cv = 0#

    For i = 0 To npts - 1
        If i = npts - 1 Then
            j = 0
        Else
            j = i + 1
        End If

        cr = u(i) * v(j) - u(j) * v(i)
        sumCross = sumCross + cr
        cu = cu + (u(i) + u(j)) * cr
        cv = cv + (v(i) + v(j)) * cr
    Next

    If sumCross = 0# Then Exit Function

    area = Abs(sumCross) / 2#
    cu = cu / (3# * sumCross)
    cv = cv / (3# * sumCross)

    cen(0) = origin(0) + cu * e1(0) + cv * e2(0)
    cen(1) = origin(1) + cu * e1(1) + cv * e2(1)
    cen(2) = origin(2) + cu * e1(2) + cv * e2(2)

    PolygonAreaCentroid = True
End Function


Sub Sort2DByAngle(ByRef u() As Double, ByRef v() As Double, ByRef angle() As Double, ByVal npts As Long)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim tmpAngle As Double
    Dim tmpU As Double
    Dim tmpV As Double

    ' Simple selection sort: intentionally old-BASIC friendly.
    For i = 0 To npts - 2
        k = i
        For j = i + 1 To npts - 1
            If angle(j) < angle(k) Then k = j
        Next

        If k <> i Then
            tmpAngle = angle(i)
            angle(i) = angle(k)
            angle(k) = tmpAngle

            tmpU = u(i)
            tmpV = v(i)
            u(i) = u(k)
            v(i) = v(k)
            u(k) = tmpU
            v(k) = tmpV
        End If
    Next
End Sub


Function Normalize3(ByRef v() As Double) As Boolean
    Dim mag As Double

    mag = Sqr(v(0) * v(0) + v(1) * v(1) + v(2) * v(2))
    If mag = 0# Then
        Normalize3 = False
        Exit Function
    End If

    v(0) = v(0) / mag
    v(1) = v(1) / mag
    v(2) = v(2) / mag
    Normalize3 = True
End Function


Function Dot3(ByRef a() As Double, ByRef b() As Double) As Double
    Dot3 = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function


Sub Cross3(ByRef a() As Double, ByRef b() As Double, ByRef c() As Double)
    c(0) = a(1) * b(2) - a(2) * b(1)
    c(1) = a(2) * b(0) - a(0) * b(2)
    c(2) = a(0) * b(1) - a(1) * b(0)
End Sub


Function PIValue() As Double
    PIValue = 3.14159265358979
End Function


' Old WinWrap/N4W BASIC may not provide ACos directly.
Function AcosCompat(ByVal x As Double) As Double
    Dim t As Double

    If x < -1# Then x = -1#
    If x > 1# Then x = 1#

    t = 1# - x * x
    If t < 0# Then t = 0#
    AcosCompat = Atan2Positive(Sqr(t), x)
End Function


Function Atan2Positive(ByVal y As Double, ByVal x As Double) As Double
    If x > 0# Then
        Atan2Positive = Atn(y / x)
        If Atan2Positive < 0# Then Atan2Positive = Atan2Positive + 2# * PIValue()
    ElseIf x < 0# Then
        Atan2Positive = Atn(y / x) + PIValue()
    ElseIf y > 0# Then
        Atan2Positive = PIValue() / 2#
    ElseIf y < 0# Then
        Atan2Positive = 3# * PIValue() / 2#
    Else
        Atan2Positive = 0#
    End If
End Function
