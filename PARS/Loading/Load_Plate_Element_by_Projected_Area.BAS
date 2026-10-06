' =============================================================================
' Description: Applies plate-element loading proportional to area projected onto a user-selected plane.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const DBUG As Boolean = True
Const C_FE_OK As Long = -1
Const C_FE_FAIL As Long = 0
Const C_FE_CANCEL As Long = 2
Const C_FT_ELEM As Long = 8
Const C_FT_LOAD_DIR As Long = 12
Const C_FCM_NORMAL As Long = 0
Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim rc As Long
    Dim i As Long
    Dim tmp As Variant
    Dim faceNo As Long

    er = 0
    If Not DBUG Then On Error GoTo Cleanup

    Dim plBase As Variant
    Dim plNormal As Variant
    Dim plAxis As Variant

    Dim nx As Double, ny As Double, nz As Double
    Dim exx As Double, exy As Double, exz As Double
    Dim eyx As Double, eyy As Double, eyz As Double

    Dim elset As Object
    Dim el As Object
    Dim nd As Object
    Dim ld As Object
    Set elset = App.feSet()
    Set el = App.feElem()
    Set nd = App.feNode()
    Set ld = App.feLoadMesh()

    Dim DOF(2) As Long
    Dim Values(4) As Double
    Dim func(4) As Long
    For i = 0 To 4
        Values(i) = 0#
        func(i) = 0
    Next

    rc = elset.Select(C_FT_ELEM, True, "Select elements to load by pressure")
    If (elset.Count() < 1) Or (rc = C_FE_FAIL) Or (rc = C_FE_CANCEL) Then
        er = 1
        GoTo Cleanup
    End If

    Do
        tmp = InputBox("Input face number to be loaded:" & vbLf & "values from 1 to 6", "Pressure by projected area:", "1")
        If tmp = "" Then
            er = 2
            GoTo Cleanup
        End If
        If IsNumeric(tmp) Then
            faceNo = CLng(tmp)
            If (faceNo >= 1) And (faceNo <= 6) Then Exit Do
        End If
    Loop

    rc = App.fePlanePick("Pick plane to project onto:", plBase, plNormal, plAxis)
    If rc = C_FE_CANCEL Then
        er = 3
        GoTo Cleanup
    End If

    nx = CDbl(plNormal(0))
    ny = CDbl(plNormal(1))
    nz = CDbl(plNormal(2))
    If Not BuildPlaneBasis(nx, ny, nz, exx, exy, exz, eyx, eyy, eyz) Then
        er = 4
        Call App.feAppMessage(C_FCM_NORMAL, "The selected plane has an invalid normal.")
        GoTo Cleanup
    End If

    ld.setID = App.Info_ActiveID(C_FT_LOAD_DIR)
    ld.flag(1) = True
    ld.fluxdir(0) = nx
    ld.fluxdir(1) = ny
    ld.fluxdir(2) = nz

    DOF(0) = faceNo
    DOF(1) = 0
    DOF(2) = 0

    rc = elset.Reset()
    Do While elset.Next() <> 0
        rc = el.Get(elset.CurrentID())

        'Preserve the original scope: linear shear-panel/plate triangles and quads.
        If ((el.type = 11) Or (el.type = 17)) And ((el.topology = 2) Or (el.topology = 4)) Then
            Values(0) = ProjectedArea(el, nd, plBase, exx, exy, exz, eyx, eyy, eyz)
            If Values(0) > 0# Then
                rc = ld.Add(-el.ID, 42, 0, DOF, Values, func)
            End If
        End If
    Loop

Cleanup:
    Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    On Error Resume Next
    Set ld = Nothing
    Set nd = Nothing
    Set el = Nothing
    Set elset = Nothing
    Set App = Nothing
End Sub

'Build two orthonormal in-plane directions from the selected plane normal.
Private Function BuildPlaneBasis(ByRef nx As Double, ByRef ny As Double, ByRef nz As Double, _
                                 ByRef exx As Double, ByRef exy As Double, ByRef exz As Double, _
                                 ByRef eyx As Double, ByRef eyy As Double, ByRef eyz As Double) As Boolean
    Dim nlen As Double
    Dim ax As Double, ay As Double, az As Double
    Dim elen As Double

    nlen = Sqr(nx * nx + ny * ny + nz * nz)
    If nlen <= 0# Then
        BuildPlaneBasis = False
        Exit Function
    End If

    nx = nx / nlen
    ny = ny / nlen
    nz = nz / nlen

    'Choose the global axis least parallel to the normal.
    If (Abs(nx) <= Abs(ny)) And (Abs(nx) <= Abs(nz)) Then
        ax = 1#: ay = 0#: az = 0#
    ElseIf Abs(ny) <= Abs(nz) Then
        ax = 0#: ay = 1#: az = 0#
    Else
        ax = 0#: ay = 0#: az = 1#
    End If

    'ex = n x a
    exx = ny * az - nz * ay
    exy = nz * ax - nx * az
    exz = nx * ay - ny * ax
    elen = Sqr(exx * exx + exy * exy + exz * exz)
    If elen <= 0# Then
        BuildPlaneBasis = False
        Exit Function
    End If
    exx = exx / elen
    exy = exy / elen
    exz = exz / elen

    'ey = n x ex
    eyx = ny * exz - nz * exy
    eyy = nz * exx - nx * exz
    eyz = nx * exy - ny * exx

    BuildPlaneBasis = True
End Function

'Area of the element after orthogonal projection onto the selected plane.
'Coordinates are expressed directly in the selected plane basis, so no
'3-D projected point array is needed.
Private Function ProjectedArea(ByRef el As Object, ByRef nd As Object, ByVal plBase As Variant, _
                               ByVal exx As Double, ByVal exy As Double, ByVal exz As Double, _
                               ByVal eyx As Double, ByVal eyy As Double, ByVal eyz As Double) As Double
    Dim ncorner As Long
    Dim i As Long, j As Long
    Dim rc As Long
    Dim qx As Double, qy As Double, qz As Double
    Dim u(3) As Double
    Dim v(3) As Double
    Dim area2 As Double

    If el.topology = 2 Then
        ncorner = 3
    ElseIf el.topology = 4 Then
        ncorner = 4
    Else
        ProjectedArea = 0#
        Exit Function
    End If

    For i = 0 To ncorner - 1
        rc = nd.Get(el.Node(i))
        If rc <> C_FE_OK Then
            ProjectedArea = 0#
            Exit Function
        End If

        qx = nd.x - CDbl(plBase(0))
        qy = nd.y - CDbl(plBase(1))
        qz = nd.z - CDbl(plBase(2))

        u(i) = qx * exx + qy * exy + qz * exz
        v(i) = qx * eyx + qy * eyy + qz * eyz
    Next

    area2 = 0#
    For i = 0 To ncorner - 1
        If i = ncorner - 1 Then
            j = 0
        Else
            j = i + 1
        End If
        area2 = area2 + u(i) * v(j) - u(j) * v(i)
    Next

    ProjectedArea = Abs(area2) / 2#
End Function
