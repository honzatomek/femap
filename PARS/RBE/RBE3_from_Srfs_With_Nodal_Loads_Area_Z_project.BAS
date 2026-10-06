' =============================================================================
' Description: Creates RBE3 spiders from selected surfaces and applies global-Z nodal forces based on XY-projected area.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Creates one RBE3 per selected geometry surface and applies a nodal force at
' the new dependent/reference node. Written against the common N4W 2004 /
' modern FEMAP BASIC/API subset.

Const C_FE_OK As Long = -1
Const C_FT_SURFACE As Long = 5
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13
Const C_RBE3_COLOR As Long = 122
Const C_FCC_BOTH As Long = 2
Const C_LOAD_FORCE As Long = 1

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim srfset As Object
    Dim srf As Object
    Dim pt As Object
    Dim nd As Object
    Dim rb3 As Object
    Dim ldset As Object
    Dim ld As Object
    Dim rc As Long
    Dim srfID As Long
    Dim numpoints As Long
    Dim points As Variant
    Dim midXYZ As Variant
    Dim normal As Variant
    Dim area As Double
    Dim projectedArea As Double
    Dim nlen As Double
    Dim fx As Double
    Dim fy As Double
    Dim fz As Double
    Dim i As Long
    Dim ndID As Long
    Dim refNodeID As Long
    Dim elemID As Long
    Dim loadID As Long
    Dim loadSetID As Long
    Dim nds() As Long
    Dim face() As Long
    Dim weight() As Double
    Dim dof() As Long

    Set srfset = App.feSet()
    Set srf = App.feSurface()
    Set pt = App.fePoint()
    Set nd = App.feNode()
    Set rb3 = App.feElem()
    Set ldset = App.feLoadSet()
    Set ld = App.feLoadMesh()

    On Error GoTo ErrorHandler

    rc = srfset.Select(C_FT_SURFACE, True, "Select surfaces to convert to RBE3 elements:")
    If (rc <> C_FE_OK) Or (srfset.Count() = 0) Then GoTo ExitHere

    loadSetID = ldset.NextEmptyID()
    ldset.title = "RBE3 - Projected Area - Z"
    rc = ldset.Put(loadSetID)
    If rc <> C_FE_OK Then GoTo ExitHere

    rc = srfset.Reset()
    srfID = srfset.Next()
    Do While srfID <> 0
        rc = srf.Get(srfID)
        If rc = C_FE_OK Then
            rc = srf.Points(C_FCC_BOTH, numpoints, points)
            If (rc = C_FE_OK) And (numpoints >= 3) Then
                rc = srf.cg(midXYZ)
                If rc <> C_FE_OK Then GoTo NextSurface
                rc = srf.Area(area)
                If (rc <> C_FE_OK) Or (area <= 0#) Then GoTo NextSurface

                ReDim nds(numpoints - 1)
                ReDim face(numpoints - 1)
                ReDim weight(numpoints - 1)
                ReDim dof(numpoints * 6 - 1)

                For i = 0 To numpoints - 1
                    rc = pt.Get(points(i))
                    If rc <> C_FE_OK Then GoTo NextSurface
                    ndID = nd.NextEmptyID()
                    nd.x = pt.x
                    nd.y = pt.y
                    nd.z = pt.z
                    rc = nd.Put(ndID)
                    If rc <> C_FE_OK Then GoTo NextSurface

                    nds(i) = ndID
                    face(i) = 0
                    weight(i) = 1#
                    dof(i * 6) = 1
                    dof(i * 6 + 1) = 1
                    dof(i * 6 + 2) = 1
                    dof(i * 6 + 3) = 0
                    dof(i * 6 + 4) = 0
                    dof(i * 6 + 5) = 0
                Next

                refNodeID = nd.NextEmptyID()
                nd.x = midXYZ(0)
                nd.y = midXYZ(1)
                nd.z = midXYZ(2)
                rc = nd.Put(refNodeID)
                If rc <> C_FE_OK Then GoTo NextSurface

                rc = srf.NormalAtXYZ(1, midXYZ, normal)
                If rc <> C_FE_OK Then GoTo NextSurface
                nlen = Sqr(normal(0) * normal(0) + normal(1) * normal(1) + normal(2) * normal(2))
                If nlen <= 0# Then GoTo NextSurface
                projectedArea = area * Abs(normal(2)) / nlen
                fx = 0#
                fy = 0#
                fz = -projectedArea

                elemID = rb3.NextEmptyID()
                rb3.ID = elemID
                rb3.type = C_FET_RIGID
                rb3.topology = C_FTO_RIGID
                rb3.RigidInterpolate = True
                rb3.color = C_RBE3_COLOR
                rb3.Node(0) = refNodeID
                rb3.release(0, 0) = 1
                rb3.release(0, 1) = 1
                rb3.release(0, 2) = 1
                rc = rb3.PutNodeList(0, numpoints, nds, face, weight, dof)
                If rc = C_FE_OK Then rc = rb3.Put(elemID)
                If rc <> C_FE_OK Then GoTo NextSurface

                ld.setID = loadSetID
                loadID = ld.NextEmptyID()
                ld.ID = loadID
                ld.type = C_LOAD_FORCE
                ld.meshID = refNodeID
                ld.load(0) = fx
                ld.load(1) = fy
                ld.load(2) = fz
                ld.XOn = True
                ld.YOn = True
                ld.ZOn = True
                rc = ld.Put(loadID)
            End If
        End If
NextSurface:
        srfID = srfset.Next()
    Loop

    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set ld = Nothing
    Set ldset = Nothing
    Set rb3 = Nothing
    Set nd = Nothing
    Set pt = Nothing
    Set srf = Nothing
    Set srfset = Nothing
    Set App = Nothing
End Sub
