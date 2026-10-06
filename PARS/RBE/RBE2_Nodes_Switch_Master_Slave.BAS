' =============================================================================
' Description: Switches the RBE2 reference node with its first listed dependent node.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Switches the reference node of a rigid element with its first listed node.
' Written against the common N4W 2004 / modern FEMAP BASIC/API subset.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13
Const C_FCM_NORMAL As Long = 0

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim el As Object
    Dim pick As Object
    Dim elID As Long
    Dim rc As Long
    Dim count As Long
    Dim oldRef As Long
    Dim newRef As Long
    Dim ndcount As Long
    Dim nds As Variant
    Dim face As Variant
    Dim weight As Variant
    Dim dof As Variant

    Set el = App.feElem()
    Set pick = App.feSet()

    On Error GoTo ErrorHandler

    Do While pick.SelectID(C_FT_ELEM, "Select rigid element to switch reference and first listed node:", elID) = C_FE_OK
        rc = el.Get(elID)
        If rc = C_FE_OK Then
            If (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) And (Not el.RigidInterpolate) Then
                rc = el.GetNodeList(0, ndcount, nds, face, weight, dof)
                If (rc = C_FE_OK) And (ndcount > 0) Then
                    oldRef = el.Node(0)
                    newRef = nds(0)
                    el.Node(0) = newRef
                    nds(0) = oldRef
                    rc = el.PutNodeList(0, ndcount, nds, face, weight, dof)
                    If rc = C_FE_OK Then rc = el.Put(elID)
                    If rc = C_FE_OK Then
                        count = count + 1
                        Call App.feAppMessage(C_FCM_NORMAL, "Updated rigid element " & CStr(elID) & ": reference " & CStr(oldRef) & " -> " & CStr(newRef) & ".")
                    End If
                End If
            End If
        End If
        Call App.feViewRegenerate(1)
    Loop

    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set pick = Nothing
    Set el = Nothing
    Set App = Nothing
End Sub
