' =============================================================================
' Description: Updates reference-node releases and listed-node DOF participation for selected RBE2 elements.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Updates reference-node releases and listed-node DOFs of selected RBE2 elements.
' Written against the common N4W 2004 / modern FEMAP BASIC/API subset.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13

Dim gDof(5) As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim elset As Object
    Dim el As Object
    Dim rc As Long
    Dim elID As Long
    Dim i As Long
    Dim nodecount As Long
    Dim n As Variant
    Dim f As Variant
    Dim w As Variant
    Dim dof As Variant

    Set elset = App.feSet()
    Set el = App.feElem()

    For i = 0 To 5
        gDof(i) = 1
    Next

    On Error GoTo ErrorHandler

    Begin Dialog UserDialog 310,105,"RBE2 Nodes Releases Update",.Rbe2Dlg
        GroupBox 10,7,190,91,"Select DOF:",.GroupBox1
        CheckBox 20,28,90,14,"T1",.t1
        CheckBox 20,49,90,14,"T2",.t2
        CheckBox 20,70,90,14,"T3",.t3
        CheckBox 120,28,90,14,"R1",.r1
        CheckBox 120,49,90,14,"R2",.r2
        CheckBox 120,70,90,14,"R3",.r3
        OKButton 210,14,90,21
        CancelButton 210,42,90,21
    End Dialog
    Dim dlg As UserDialog

    Do While Dialog(dlg) = C_FE_OK
        rc = elset.Select(C_FT_ELEM, True, "Select RBE2 elements for DOF update:")
        If rc <> C_FE_OK Then Exit Do

        rc = elset.Reset()
        elID = elset.Next()
        Do While elID <> 0
            rc = el.Get(elID)
            If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) And (Not el.RigidInterpolate) Then
                rc = el.GetNodeList(0, nodecount, n, f, w, dof)
                If rc = C_FE_OK Then
                    For i = 0 To 5
                        el.release(0, i) = (gDof(i) <> 0)
                    Next
                    For i = 0 To nodecount - 1
                        dof(6 * i) = gDof(0)
                        dof(6 * i + 1) = gDof(1)
                        dof(6 * i + 2) = gDof(2)
                        dof(6 * i + 3) = gDof(3)
                        dof(6 * i + 4) = gDof(4)
                        dof(6 * i + 5) = gDof(5)
                    Next
                    rc = el.PutNodeList(0, nodecount, n, f, w, dof)
                    If rc = C_FE_OK Then rc = el.Put(el.ID)
                End If
            End If
            elID = elset.Next()
        Loop
        Call App.feViewRegenerate(1)
    Loop

    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set el = Nothing
    Set elset = Nothing
    Set App = Nothing
End Sub

Private Function Rbe2Dlg(DlgItem, Action, SuppValue) As Boolean
    Dim i As Long
    Select Case Action
    Case 1
        For i = 0 To 5
            Call DlgValue(i + 1, gDof(i))
        Next
    Case 2
        Rbe2Dlg = True
        Select Case DlgItem
        Case "OK"
            Rbe2Dlg = False
        Case "Cancel"
            Rbe2Dlg = False
        Case Else
            i = DlgNumber(DlgItem)
            If (i >= 1) And (i <= 6) Then gDof(i - 1) = DlgValue(i)
        End Select
    End Select
End Function
