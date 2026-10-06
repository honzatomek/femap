' =============================================================================
' Description: Replaces the reference/master node of selected rigid elements.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Replaces the reference node of an RBE2/RBE3 rigid element.
' Written against the common N4W 2004 / modern FEMAP BASIC/API subset.

Const C_FE_OK As Long = -1
Const C_FT_NODE As Long = 7
Const C_FT_ELEM As Long = 8
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13
Const C_FCM_NORMAL As Long = 0

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim el As Object
    Dim epick As Object
    Dim npick As Object
    Dim elID As Long
    Dim oldNode As Long
    Dim newNode As Long
    Dim rc As Long

    Set el = App.feElem()
    Set epick = App.feSet()
    Set npick = App.feSet()

    On Error GoTo ErrorHandler

    Do While epick.SelectID(C_FT_ELEM, "Select RBE2/RBE3 element for reference-node update:", elID) = C_FE_OK
        rc = el.Get(elID)
        If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) Then
            oldNode = el.Node(0)
            rc = npick.SelectID(C_FT_NODE, "Select new reference node:", newNode)
            If rc <> C_FE_OK Then GoTo ExitHere
            el.Node(0) = newNode
            rc = el.Put(elID)
            If rc = C_FE_OK Then Call App.feAppMessage(C_FCM_NORMAL, "Updated rigid element " & CStr(elID) & ": reference node " & CStr(oldNode) & " -> " & CStr(newNode) & ".")
            Call App.feViewRegenerate(1)
        End If
    Loop

    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set npick = Nothing
    Set epick = Nothing
    Set el = Nothing
    Set App = Nothing
End Sub
