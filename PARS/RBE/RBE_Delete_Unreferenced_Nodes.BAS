' =============================================================================
' Description: Deletes nodes that are left unreferenced after rigid-element editing.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Removes listed nodes from selected rigid elements when that node is referenced
' by no other element. The node entity itself is not deleted.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13
Const C_FGD_ELEM_BYNODE As Long = 25
Const C_FCM_NORMAL As Long = 0

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim el As Object
    Dim elset As Object
    Dim refset As Object
    Dim rc As Long
    Dim i As Long
    Dim j As Long
    Dim elID As Long
    Dim removedTotal As Long
    Dim removedThis As Long
    Dim nodecount1 As Long
    Dim nodecount2 As Long
    Dim nds1 As Variant
    Dim face1 As Variant
    Dim weight1 As Variant
    Dim dof1 As Variant
    Dim nds2() As Long
    Dim face2() As Long
    Dim weight2() As Double
    Dim dof2() As Long

    Set el = App.feElem()
    Set elset = App.feSet()
    Set refset = App.feSet()

    On Error GoTo ErrorHandler

    rc = elset.Select(C_FT_ELEM, True, "Select rigid elements for free-node reduction:")
    If rc <> C_FE_OK Then GoTo ExitHere

    rc = elset.Reset()
    elID = elset.Next()
    Do While elID <> 0
        rc = el.Get(elID)
        If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) And el.RigidInterpolate Then
            rc = el.GetNodeList(0, nodecount1, nds1, face1, weight1, dof1)
            If (rc = C_FE_OK) And (nodecount1 > 0) Then
                ReDim nds2(nodecount1 - 1)
                ReDim face2(nodecount1 - 1)
                ReDim weight2(nodecount1 - 1)
                ReDim dof2(nodecount1 * 6 - 1)
                nodecount2 = 0
                removedThis = 0

                For i = 0 To nodecount1 - 1
                    rc = refset.Clear()
                    rc = refset.AddRule(nds1(i), C_FGD_ELEM_BYNODE)
                    If refset.Count() > 1 Then
                        nds2(nodecount2) = nds1(i)
                        face2(nodecount2) = face1(i)
                        weight2(nodecount2) = weight1(i)
                        For j = 0 To 5
                            dof2(nodecount2 * 6 + j) = dof1(i * 6 + j)
                        Next
                        nodecount2 = nodecount2 + 1
                    Else
                        removedThis = removedThis + 1
                        removedTotal = removedTotal + 1
                    End If
                Next

                If (removedThis > 0) And (nodecount2 > 0) Then
                    rc = el.PutNodeList(0, nodecount2, nds2, face2, weight2, dof2)
                    If rc = C_FE_OK Then rc = el.Put(el.ID)
                ElseIf (removedThis > 0) And (nodecount2 = 0) Then
                    Call App.feAppMessage(C_FCM_NORMAL, "Element " & CStr(el.ID) & ": all listed nodes would be removed; element left unchanged.")
                    removedTotal = removedTotal - removedThis
                    removedThis = 0
                End If

                If removedThis > 0 Then Call App.feAppMessage(C_FCM_NORMAL, "Element " & CStr(el.ID) & ": removed " & CStr(removedThis) & " singly-referenced listed nodes.")
            End If
        End If
        elID = elset.Next()
    Loop

    Call App.feAppMessage(C_FCM_NORMAL, CStr(removedTotal) & " listed nodes removed from rigid elements in total.")
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set refset = Nothing
    Set elset = Nothing
    Set el = Nothing
    Set App = Nothing
End Sub
