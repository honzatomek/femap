' =============================================================================
' Description: Replaces selected nodes in rigid-element node lists.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Replaces one listed node in an RBE2/RBE3 rigid element.
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
    Dim i As Long
    Dim found As Boolean
    Dim ndcount As Long
    Dim nds As Variant
    Dim face As Variant
    Dim weight As Variant
    Dim dof As Variant

    Set el = App.feElem()
    Set epick = App.feSet()
    Set npick = App.feSet()

    On Error GoTo ErrorHandler

    Do While epick.SelectID(C_FT_ELEM, "Select RBE2/RBE3 element for node update:", elID) = C_FE_OK
        rc = el.Get(elID)
        If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) Then
            rc = el.GetNodeList(0, ndcount, nds, face, weight, dof)
            If rc = C_FE_OK Then
                found = False
                Do While Not found
                    rc = npick.SelectID(C_FT_NODE, "Select listed node to replace:", oldNode)
                    If rc <> C_FE_OK Then GoTo ExitHere
                    For i = 0 To ndcount - 1
                        If nds(i) = oldNode Then
                            found = True
                            Exit For
                        End If
                    Next
                    If Not found Then MsgBox "Node " & CStr(oldNode) & " is not a listed node of rigid element " & CStr(elID) & "."
                Loop

                rc = npick.SelectID(C_FT_NODE, "Select replacement node:", newNode)
                If rc <> C_FE_OK Then GoTo ExitHere
                nds(i) = newNode
                rc = el.PutNodeList(0, ndcount, nds, face, weight, dof)
                If rc = C_FE_OK Then rc = el.Put(elID)
                If rc = C_FE_OK Then Call App.feAppMessage(C_FCM_NORMAL, "Updated rigid element " & CStr(elID) & ": node " & CStr(oldNode) & " -> " & CStr(newNode) & ".")
                Call App.feViewRegenerate(1)
            End If
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
