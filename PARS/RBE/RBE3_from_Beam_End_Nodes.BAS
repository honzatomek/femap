' =============================================================================
' Description: Creates an RBE3 from unique end nodes of selected beam elements.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Creates one RBE3 from the unique end nodes of selected BEAM elements.
' The dependent/reference node is the arithmetic mean of the selected beam end nodes.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FET_BEAM As Long = 5
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13
Const C_RBE3_COLOR As Long = 122
Const C_FCM_NORMAL As Long = 0

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim elset As Object
    Dim ndset As Object
    Dim el As Object
    Dim nd As Object
    Dim rc As Long
    Dim elID As Long
    Dim ndID As Long
    Dim i As Long
    Dim numnode As Long
    Dim sx As Double
    Dim sy As Double
    Dim sz As Double
    Dim nds() As Long
    Dim face() As Long
    Dim weight() As Double
    Dim dof() As Long
    Dim newNodeID As Long
    Dim newElemID As Long

    Set elset = App.feSet()
    Set ndset = App.feSet()
    Set el = App.feElem()
    Set nd = App.feNode()

    On Error GoTo ErrorHandler

    rc = elset.Select(C_FT_ELEM, True, "Select BEAM elements for RBE3 spider:")
    If (rc <> C_FE_OK) Or (elset.Count() < 1) Then GoTo ExitHere

    rc = elset.Reset()
    elID = elset.Next()
    Do While elID <> 0
        rc = el.Get(elID)
        If (rc = C_FE_OK) And (el.type = C_FET_BEAM) Then
            rc = ndset.Add(el.Node(0))
            rc = ndset.Add(el.Node(1))
        End If
        elID = elset.Next()
    Loop

    numnode = ndset.Count()
    If numnode < 2 Then
        MsgBox "At least two unique BEAM end nodes are required."
        GoTo ExitHere
    End If

    ReDim nds(numnode - 1)
    ReDim face(numnode - 1)
    ReDim weight(numnode - 1)
    ReDim dof(numnode * 6 - 1)

    rc = ndset.Reset()
    For i = 0 To numnode - 1
        ndID = ndset.Next()
        nds(i) = ndID
        face(i) = 0
        weight(i) = 1#
        dof(i * 6) = 1
        dof(i * 6 + 1) = 1
        dof(i * 6 + 2) = 1
        dof(i * 6 + 3) = 0
        dof(i * 6 + 4) = 0
        dof(i * 6 + 5) = 0

        rc = nd.Get(ndID)
        If rc = C_FE_OK Then
            sx = sx + nd.x
            sy = sy + nd.y
            sz = sz + nd.z
        End If
    Next

    newNodeID = nd.NextEmptyID()
    nd.x = sx / numnode
    nd.y = sy / numnode
    nd.z = sz / numnode
    rc = nd.Put(newNodeID)
    If rc <> C_FE_OK Then GoTo ExitHere

    newElemID = el.NextEmptyID()
    el.ID = newElemID
    el.type = C_FET_RIGID
    el.topology = C_FTO_RIGID
    el.RigidInterpolate = True
    el.color = C_RBE3_COLOR
    el.Node(0) = newNodeID
    el.release(0, 0) = 1
    el.release(0, 1) = 1
    el.release(0, 2) = 1
    rc = el.PutNodeList(0, numnode, nds, face, weight, dof)
    If rc = C_FE_OK Then rc = el.Put(newElemID)
    If rc = C_FE_OK Then Call App.feAppMessage(C_FCM_NORMAL, "Created RBE3 element " & CStr(newElemID) & ".")
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set nd = Nothing
    Set el = Nothing
    Set ndset = Nothing
    Set elset = Nothing
    Set App = Nothing
End Sub
