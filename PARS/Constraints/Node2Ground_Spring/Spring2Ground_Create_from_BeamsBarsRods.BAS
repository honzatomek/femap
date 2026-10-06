' =============================================================================
' Description: Creates spring-to-ground elements and properties at nodes of selected beam, bar, and rod elements.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' N4W/FEMAP portable constants (do not rely on type-library enum names)
Const C_FE_FAIL As Long = 0
Const C_FE_CANCEL As Long = 2
Const C_FT_ELEM As Long = 8
Const C_FGD_ELEM_BYNODE As Long = 25
Const C_FGD_ELEM_BYTYPE As Long = 24
Const C_FGD_NODE_ONELEM As Long = 20
Const C_MSG_NORMAL As Long = 0
Const C_MSG_WARNING As Long = 2

Const DBUG As Boolean = False
Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim nset As Object
    Dim eset As Object
    Dim eset_s2g As Object
    Dim n As Object
    Dim els_by_node As Object
    Dim rc As Long
    Dim nID As Long
    Dim pID As Long
    Dim csID As Long
    Dim length As Double
    Dim pname As String
    Dim stiff_coeff(5) As Double

    er = 0
    If Not DBUG Then On Error GoTo Cleanup

    Set eset = App.feSet()
    Set eset_s2g = App.feSet()
    Set nset = App.feSet()
    Set n = App.feNode()
    Set els_by_node = App.feSet()

    rc = eset.Select(C_FT_ELEM, True, "Select elements to be supported by Spring2Ground:")
    If (rc <> -1) Or (eset.Count() < 1) Then
        Call App.feAppMessage(C_MSG_NORMAL, "No Element selected.")
        GoTo Cleanup
    End If

    rc = get_stiffness_coefficients(App, stiff_coeff)
    If rc = C_FE_CANCEL Then GoTo Cleanup

    rc = eset_s2g.AddRule(1, C_FGD_ELEM_BYTYPE)
    rc = eset_s2g.AddRule(2, C_FGD_ELEM_BYTYPE)
    rc = eset_s2g.AddRule(5, C_FGD_ELEM_BYTYPE)
    rc = eset.RemoveNotCommon(eset_s2g.ID)

    nset.Clear()
    rc = nset.AddSetRule(eset.ID, C_FGD_NODE_ONELEM)
    nset.Reset()
    nID = nset.Next()

    Do While nID > 0
        rc = n.Get(nID)
        If rc = C_FE_FAIL Then
            Call App.feAppMessage(C_MSG_WARNING, "Node " & nID & " not found.")
            GoTo Cleanup
        End If

        ' 1 - Create a set of elements on current node, get their length and property name.
        rc = get_length_name(App, els_by_node, pname, length, nID, eset.ID)
        If rc = -1 Then
            er = 1
            GoTo Cleanup
        End If

        ' 2 - Create coordinate system from line elements.
        rc = coordination_system_from_line_elements(App, els_by_node.ID, n.ID, pname)
        If rc = -1 Then
            er = 2
            GoTo Cleanup
        Else
            csID = rc
        End If

        n.outCSys = csID
        rc = n.Put(n.ID)

        ' 3 - Create Spring2Ground property.
        rc = spring_to_ground_property(App, pname, length, stiff_coeff, csID)
        If rc = -1 Then
            er = 3
            GoTo Cleanup
        Else
            pID = rc
        End If

        ' 4 - Create Spring2Ground element.
        rc = spring_to_ground_element(App, n.ID, pID)
        If rc = -1 Then
            er = 4
            GoTo Cleanup
        End If

        nID = nset.Next()
    Loop

Cleanup:
    On Error Resume Next
    Call App.feAppMessage(C_MSG_NORMAL, "The script exited with code: " & er)
    Set els_by_node = Nothing
    Set n = Nothing
    Set eset_s2g = Nothing
    Set eset = Nothing
    Set nset = Nothing
    Set App = Nothing
End Sub

Function get_stiffness_coefficients(ByRef App As Object, ByRef stiffness_coefficients() As Double) As Long
    Dim rc As Long
    Dim vecBase1(2) As Double
    Dim vecBase2(2) As Double

    On Error GoTo Cleanup

    App.Info_DefaultX = 1#
    App.Info_DefaultY = 1#
    App.Info_DefaultZ = 1#
    rc = App.feCoordPick("Input Translational Stiffness Coefficients:", vecBase1)
    If rc = C_FE_CANCEL Then GoTo Cleanup

    App.Info_DefaultX = 1#
    App.Info_DefaultY = 1#
    App.Info_DefaultZ = 1#
    rc = App.feCoordPick("Input Rotational Stiffness Coefficients:", vecBase2)
    If rc = C_FE_CANCEL Then GoTo Cleanup

    stiffness_coefficients(0) = vecBase1(0)
    stiffness_coefficients(1) = vecBase1(1)
    stiffness_coefficients(2) = vecBase1(2)
    stiffness_coefficients(3) = vecBase2(0)
    stiffness_coefficients(4) = vecBase2(1)
    stiffness_coefficients(5) = vecBase2(2)

Cleanup:
    get_stiffness_coefficients = rc
End Function

Function get_length_name(ByRef App As Object, ByRef eset As Object, ByRef ele_name As String, ByRef length As Double, ByVal nodeID As Long, ByVal base_element_set_ID As Long) As Long
    Dim rc As Long
    Dim e As Object
    Dim p As Object
    Dim eID As Long
    Dim tmp As Double

    On Error GoTo Cleanup

    rc = -1
    Set eset = App.feSet()
    Set e = App.feElem()
    Set p = App.feProp()
    ele_name = "S2G"
    length = 0#

    eset.Clear()
    rc = eset.AddRule(nodeID, C_FGD_ELEM_BYNODE)
    If (rc = C_FE_FAIL) Or (eset.Count() < 1) Then
        rc = -1
        Call App.feAppMessage(C_MSG_NORMAL, "No Elements on Node " & nodeID & ".")
        GoTo Cleanup
    End If

    rc = eset.RemoveNotCommon(base_element_set_ID)
    If rc = C_FE_FAIL Then
        rc = -1
        GoTo Cleanup
    End If

    eset.Reset()
    eID = eset.Next()
    Do While eID > 0
        rc = e.Get(eID)
        If rc = C_FE_FAIL Then
            rc = -1
            Call App.feAppMessage(C_MSG_NORMAL, "Element " & eID & " not found.")
            GoTo Cleanup
        End If

        If e.type = 1 Or e.type = 2 Or e.type = 5 Then
            rc = e.Length(tmp)
            If rc = C_FE_FAIL Then
                rc = -1
                GoTo Cleanup
            End If
            length = length + tmp

            rc = p.Get(e.propID)
            If rc = C_FE_FAIL Then
                rc = -1
                GoTo Cleanup
            End If
            If p.title <> ele_name Then ele_name = ele_name & "_" & p.title
        End If
        eID = eset.Next()
    Loop

    length = length / 2#
    rc = eset.ID

Cleanup:
    On Error Resume Next
    If rc = -1 Then Call App.feAppMessage(C_MSG_WARNING, "Problem With Line Elements on Node " & nodeID & ".")
    Set e = Nothing
    Set p = Nothing
    get_length_name = rc
End Function

Function coordination_system_from_line_elements(ByRef App As Object, ByVal element_set_ID As Long, ByVal nodeID As Long, ByVal csys_name As String) As Long
    ' Creates a CSys from line-element end points. Returns new CSys ID, or -1 on failure.
    Dim rc As Long
    Dim nset As Object
    Dim n As Object
    Dim cs As Object
    Dim csID As Long
    Dim nID As Long

    Dim x1 As Double, y1 As Double, z1 As Double
    Dim vx As Double, vy As Double, vz As Double
    Dim yx As Double, yy As Double, yz As Double
    Dim rx As Double, ry As Double, rz As Double
    Dim vlen As Double
    Dim p0(2) As Double
    Dim p1(2) As Double
    Dim p2(2) As Double

    On Error GoTo Cleanup
    rc = -1

    Set nset = App.feSet()
    Set n = App.feNode()
    Set cs = App.feCSys()

    csID = cs.NextEmptyID()

    nset.Clear()
    rc = nset.AddSetRule(element_set_ID, C_FGD_NODE_ONELEM)
    Select Case nset.Count()
    Case 0, 1
        Call App.feAppMessage(C_MSG_NORMAL, csys_name & " CSys creation: element set has " & nset.Count() & " nodes.")
        rc = -1
        GoTo Cleanup
    Case 2
    Case 3
        rc = nset.Remove(nodeID)
    Case Else
        Call App.feAppMessage(C_MSG_NORMAL, csys_name & " CSys creation: element set has " & nset.Count() & " nodes, only first two will be used.")
        rc = nset.Remove(nodeID)
    End Select

    nset.Reset()
    nID = nset.Next()
    rc = n.Get(nID)
    If rc = C_FE_FAIL Then
        rc = -1
        GoTo Cleanup
    End If
    x1 = n.x
    y1 = n.y
    z1 = n.z

    nID = nset.Next()
    rc = n.Get(nID)
    If rc = C_FE_FAIL Then
        rc = -1
        GoTo Cleanup
    End If

    vx = n.x - x1
    vy = n.y - y1
    vz = n.z - z1
    vlen = Sqr(vx * vx + vy * vy + vz * vz)
    If vlen <= 1.0E-20 Then
        rc = -1
        GoTo Cleanup
    End If
    Call unit3(vx, vy, vz)

    ' Match the original global-Z construction when possible; use global Y near Z-parallel lines.
    If Abs(vz) < 0.999 Then
        rx = 0#
        ry = 0#
        rz = 1#
    Else
        rx = 0#
        ry = 1#
        rz = 0#
    End If
    Call cross3(rx, ry, rz, vx, vy, vz, yx, yy, yz)
    Call unit3(yx, yy, yz)

    rc = n.Get(nodeID)
    If rc = C_FE_FAIL Then
        rc = -1
        GoTo Cleanup
    End If

    rc = cs.Get(0)
    p0(0) = n.x
    p0(1) = n.y
    p0(2) = n.z
    p1(0) = n.x + vx
    p1(1) = n.y + vy
    p1(2) = n.z + vz
    p2(0) = n.x + yx
    p2(1) = n.y + yy
    p2(2) = n.z + yz
    rc = cs.XYPoints(p0, p1, p2)
    cs.title = csys_name

    rc = cs.Put(csID)
    If rc = C_FE_FAIL Then
        rc = -1
    Else
        rc = csID
    End If

Cleanup:
    On Error Resume Next
    If rc = -1 Then Call App.feAppMessage(C_MSG_WARNING, "CSys from Line Elements " & csys_name & " failed.")
    Set nset = Nothing
    Set n = Nothing
    Set cs = Nothing
    coordination_system_from_line_elements = rc
End Function

Function spring_to_ground_property(ByRef App As Object, ByVal property_name As String, ByVal base_stiffness As Double, ByRef stiffness_coefficients() As Double, ByVal csysID As Long) As Long
    ' Creates a Spring To Ground property. Returns new property ID, or -1 on failure.
    Dim rc As Long, i As Long
    Dim p As Object
    Dim pID As Long

    On Error GoTo Cleanup
    rc = -1
    Set p = App.feProp()
    pID = p.NextEmptyID()

    p.title = property_name
    p.type = 41
    p.cbush = 1
    p.flag(0) = csysID
    p.refCS = csysID
    For i = 0 To 5 Step 1
        p.pval(36 + i) = base_stiffness * stiffness_coefficients(i)
    Next

    rc = p.Put(pID)
    If rc = C_FE_FAIL Then
        rc = -1
    Else
        rc = pID
    End If

Cleanup:
    On Error Resume Next
    If rc = -1 Then Call App.feAppMessage(C_MSG_WARNING, "Spring2Ground Property " & property_name & " failed.")
    Set p = Nothing
    spring_to_ground_property = rc
End Function

Function spring_to_ground_element(ByRef App As Object, ByVal nodeID As Long, ByVal propID As Long) As Long
    ' Creates a grounded spring element. Orientation is taken from the property for old-version compatibility.
    Dim rc As Long
    Dim e As Object
    Dim eID As Long

    On Error GoTo Cleanup
    rc = -1
    Set e = App.feElem()

    eID = e.NextEmptyID()
    e.type = 41
    e.propID = propID
    e.topology = 9
    e.Node(0) = nodeID

    ' Do not use SpringUseCID/SpringCID here: those Element API attributes were added in FEMAP 11.3.
    rc = e.Put(eID)
    If rc = C_FE_FAIL Then
        rc = -1
    Else
        rc = eID
    End If

Cleanup:
    On Error Resume Next
    If rc = -1 Then Call App.feAppMessage(C_MSG_WARNING, "Spring2Ground Element at Node " & nodeID & " failed.")
    Set e = Nothing
    spring_to_ground_element = rc
End Function

Sub cross3(ByVal ax As Double, ByVal ay As Double, ByVal az As Double, ByVal bx As Double, ByVal by As Double, ByVal bz As Double, ByRef cx As Double, ByRef cy As Double, ByRef cz As Double)
    cx = ay * bz - az * by
    cy = az * bx - ax * bz
    cz = ax * by - ay * bx
End Sub

Sub unit3(ByRef x As Double, ByRef y As Double, ByRef z As Double)
    Dim l As Double
    l = Sqr(x * x + y * y + z * z)
    If l > 1.0E-20 Then
        x = x / l
        y = y / l
        z = z / l
    End If
End Sub
