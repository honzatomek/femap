' =============================================================================
' Description: Creates an RBE3 from selected nodes, places the reference node at the projected area centroid, and applies an area-magnitude nodal force.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_RBE3_COLOR As Long = 122

' RBE3_from_Nodes_With_Nodal_Loads_Area_compat
'
' Compatibility-oriented version intended for both old FEMAP/Nastran-for-Windows
' BASIC environments and modern FEMAP WinWrap Basic.
'
' Creates an interpolation rigid element from selected master/independent nodes,
' creates the dependent/reference node at the projected polygon area centroid,
' computes the projected area enclosed by the selected nodes, and applies
' a force of magnitude = area in the selected direction at the reference node.
'
' Important compatibility choices:
'   - late-bound FEMAP objects (As Object)
'   - feGetObject() for N4W 2004 / modern FEMAP compatibility
'   - geometry uses only one-dimensional primitive Double arrays
'   - PutNodeList uses direct primitive arrays; no Variant-array assignment
'   - RBE3 classification uses RigidInterpolate = True
'   - default RBE3 color is FEMAP palette 122 = RGB(128,255,255)
'   - only old/basic language constructs are used

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
    dlg.my = 0
    dlg.pz = 0
    dlg.mz = 1
    dlg.normal = 0

    If Dialog(dlg) = 0 Then Exit Sub

    Dim App As Object
    Dim LdSet As Object
    Dim Ld As Object
    Dim ndset As Object
    Dim nd As Object
    Dim midnd As Object
    Dim el As Object

    Dim rc As Long
    Dim LdSetID As Long
    Dim midndID As Long
    Dim eID As Long
    Dim loadID As Long
    Dim ndcount As Long
    Dim i As Long
    Dim newLoadSet As Boolean

    Dim nds() As Long
    Dim face() As Long
    Dim DOF() As Long
    Dim weight() As Double
    Dim stage As String

    Dim xyzX() As Double
    Dim xyzY() As Double
    Dim xyzZ() As Double
    Dim origin(2) As Double
    Dim normal(2) As Double
    Dim centroid(2) As Double
    Dim force(2) As Double
    Dim A As Double

    On Error GoTo ErrorHandler
    stage = "attach to FEMAP"

    Set App = feGetObject()
    stage = "create FEMAP API objects"
    Set ndset = App.feSet()
    Set nd = App.feNode()
    Set midnd = App.feNode()
    Set el = App.feElem()
    Set Ld = App.feLoadMesh()

    ' Use a Set object only for the load-set selection dialog.
    Set LdSet = App.feSet()
    stage = "select load set"
    rc = LdSet.SelectID(12, "Select ID of the Load Set for nodal forces (Cancel = New LS):", LdSetID)

    newLoadSet = False
    If rc <> -1 Then
        Set LdSet = App.feLoadSet()
        LdSetID = LdSet.NextEmptyID()
        LdSet.title = "LC" & LdSetID & " - RBE3 - area "
        newLoadSet = True
    Else
        App.Info_ActiveID(12) = LdSetID       ' FT_LOAD_DIR = 12
    End If

    stage = "select master nodes"
    rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")   ' FT_NODE = 7

    If rc = -1 And newLoadSet Then
        rc = LdSet.Put(LdSetID)
        If rc <> -1 Then
            MsgBox "Could not create Load Set " & CStr(LdSetID) & ". Return code: " & CStr(rc)
            GoTo ExitHere
        End If
        App.Info_ActiveID(12) = LdSetID
        rc = -1
    End If

    Do While rc = -1
        ndcount = ndset.Count()

        If ndcount < 3 Then
            MsgBox "Select at least three non-collinear nodes."
            rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")
            GoTo ContinueLoop
        End If

        stage = "read node coordinates"
        If Not ReadNodeCoordinates(App, ndset, xyzX, xyzY, xyzZ, ndcount) Then
            MsgBox "Could not read the selected node coordinates."
            GoTo ExitHere
        End If

        stage = "fit best-fit plane"
        If Not BestFitPlane(xyzX, xyzY, xyzZ, ndcount, origin, normal) Then
            MsgBox "The selected nodes are collinear, coincident, or do not define a unique best-fit plane."
            rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")
            GoTo ContinueLoop
        End If

        stage = "project polygon and compute area"
        If Not PolygonAreaCentroid(xyzX, xyzY, xyzZ, ndcount, origin, normal, A, centroid) Then
            MsgBox "Could not determine a non-zero enclosed area from the selected nodes."
            rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")
            GoTo ContinueLoop
        End If

        ' Load direction.
        If dlg.normal = 1 Then
            force(0) = normal(0)
            force(1) = normal(1)
            force(2) = normal(2)
        Else
            force(0) = CDbl(dlg.px - dlg.mx)
            force(1) = CDbl(dlg.py - dlg.my)
            force(2) = CDbl(dlg.pz - dlg.mz)
            If Not Normalize3(force) Then
                MsgBox "Select one force direction."
                GoTo ExitHere
            End If
        End If

        ' RBE3 dependent/reference node at the area centroid.
        stage = "create independent node"
        midndID = midnd.NextEmptyID()
        midnd.x = centroid(0)
        midnd.y = centroid(1)
        midnd.z = centroid(2)
        rc = midnd.Put(midndID)
        If rc <> -1 Then
            MsgBox "Could not create RBE3 reference node " & CStr(midndID) & ". Return code: " & CStr(rc)
            GoTo ExitHere
        End If

        ' Collect independent interpolation nodes and data.
        ReDim nds(ndcount - 1)
        ReDim face(ndcount - 1)
        ReDim weight(ndcount - 1)
        ReDim DOF(ndcount * 6 - 1)

        rc = ndset.Reset()
        For i = 0 To ndcount - 1
            nds(i) = ndset.Next()
            face(i) = 0
            weight(i) = 1#

            DOF(i * 6) = 1
            DOF(i * 6 + 1) = 1
            DOF(i * 6 + 2) = 1
            DOF(i * 6 + 3) = 0
            DOF(i * 6 + 4) = 0
            DOF(i * 6 + 5) = 0
        Next

        ' Rigid interpolation element / RBE3-style spider.
        ' Use direct typed arrays.  Old N4W BASIC cannot do Variant = typedArray.
        ' Do not pass Null for the face array either; supply a real zero-filled array.
        stage = "get next rigid element ID"
        eID = el.NextEmptyID()

        stage = "set rigid element ID"
        el.ID = eID

        stage = "set rigid element type"
        el.type = 29                    ' FET_L_RIGID

        stage = "set rigid topology"
        el.topology = 13               ' FTO_RIGIDLIST

        stage = "mark rigid interpolation / RBE3"
        el.RigidInterpolate = True
        el.color = C_RBE3_COLOR

        stage = "set RBE3 dependent/reference node"
        el.Node(0) = midndID

        stage = "set RBE3 dependent TX"
        el.release(0, 0) = 1
        stage = "set RBE3 dependent TY"
        el.release(0, 1) = 1
        stage = "set RBE3 dependent TZ"
        el.release(0, 2) = 1

        stage = "PutNodeList direct arrays"
        rc = el.PutNodeList(0, ndcount, nds, face, weight, DOF)
        If rc <> -1 Then
            MsgBox "PutNodeList failed for rigid element " & CStr(eID) & ". Return code: " & CStr(rc)
            GoTo ExitHere
        End If

        stage = "store rigid element"
        rc = el.Put(eID)
        If rc <> -1 Then
            MsgBox "Could not create rigid element " & CStr(eID) & ". Return code: " & CStr(rc)
            GoTo ExitHere
        End If

        ' Force magnitude equals projected area.
        stage = "create nodal load"
        Ld.setID = LdSetID
        loadID = Ld.NextEmptyID()
        Ld.ID = loadID
        Ld.type = 1
        Ld.meshID = midndID
        Ld.load(0) = force(0) * Abs(A)
        Ld.load(1) = force(1) * Abs(A)
        Ld.load(2) = force(2) * Abs(A)
        Ld.XOn = True
        Ld.YOn = True
        Ld.ZOn = True

        rc = Ld.Put(loadID)
        If rc <> -1 Then
            MsgBox "Could not create nodal force " & CStr(loadID) & ". Return code: " & CStr(rc)
            GoTo ExitHere
        End If

        App.feViewRegenerate(0)

        rc = ndset.Select(7, True, "Select Nodes for Rigid Spider")

ContinueLoop:
    Loop

    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & " during '" & stage & "': " & Err.Description

ExitHere:
    Set el = Nothing
    Set midnd = Nothing
    Set nd = Nothing
    Set Ld = Nothing
    Set ndset = Nothing
    Set LdSet = Nothing
    Set App = Nothing
End Sub


Rem See DialogFunc help topic for more information.
Private Function Settings(DlgItem, Action, SuppValue) As Boolean
    Settings = True
    Select Case Action
    Case 1 ' Dialog box initialization

    Case 2 ' Value changing or button pressed
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

    Case 3 ' TextBox or ComboBox text changed
    Case 4 ' Focus changed
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function


' Read selected FEMAP nodes into three one-dimensional Double arrays.
' Old N4W BASIC accepts dynamic primitive arrays but not dynamic arrays of UDTs.
Function ReadNodeCoordinates(ByRef App As Object, ByRef ndset As Object, ByRef xyzX() As Double, ByRef xyzY() As Double, ByRef xyzZ() As Double, ByVal ndcount As Long) As Boolean
    Dim nd As Object
    Dim rc As Long
    Dim i As Long
    Dim ndID As Long

    ReadNodeCoordinates = False
    ReDim xyzX(ndcount - 1)
    ReDim xyzY(ndcount - 1)
    ReDim xyzZ(ndcount - 1)

    Set nd = App.feNode()
    rc = ndset.Reset()

    For i = 0 To ndcount - 1
        ndID = ndset.Next()
        If ndID = 0 Then
            Set nd = Nothing
            Exit Function
        End If

        rc = nd.Get(ndID)
        If rc <> -1 Then
            Set nd = Nothing
            Exit Function
        End If

        xyzX(i) = nd.x()
        xyzY(i) = nd.y()
        xyzZ(i) = nd.z()
    Next

    Set nd = Nothing
    ReadNodeCoordinates = True
End Function


' Least-squares best-fit plane through the selected points.
'
' The covariance matrix is symmetric 3x3.  Its smallest eigenvalue is found
' analytically from the cubic characteristic equation in its stable symmetric
' form.  The corresponding eigenvector is the plane normal.
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
