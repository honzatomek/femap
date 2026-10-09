' =============================================================================
' Description: Sorts all rigid elements into separate RBE2 and RBE3 FEMAP groups.
'              RBE3 is identified by RigidInterpolate=True.
'
'              Nastran for Windows 2004 compatibility version:
'              - no dynamic arrays
'              - no ReDim / ReDim Preserve
'              - no Array()
'              - no temporary feSet IDs
'              - group membership uses feGroup.RangeAdd with range type 21 (Elem_ID)
'
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FGR_ELEM_ID As Long = 21
Const C_FET_RIGID As Long = 29
Const C_FGD_ELEM_BYTYPE As Long = 24
Const C_FCM_NORMAL As Long = 0
Const C_FCM_WARNING As Long = 2
Const C_FCM_ERROR As Long = 3

Sub Main
    Dim App As Object
    Dim rigidSet As Object
    Dim rbe As Object
    Dim g2 As Object
    Dim g3 As Object
    Dim rc As Long
    Dim eid As Long
    Dim gid2 As Long
    Dim gid3 As Long
    Dim rigidCount As Long
    Dim rbe2Count As Long
    Dim rbe3Count As Long
    Dim errNumber As Long
    Dim errDescription As String
    Dim msg As String

    On Error GoTo Cleanup

    Set App = feGetObject()
    Set rigidSet = App.feSet()
    Set rbe = App.feElem()
    Set g2 = App.feGroup()
    Set g3 = App.feGroup()

    rc = rigidSet.AddRule(C_FET_RIGID, C_FGD_ELEM_BYTYPE)
    If rc <> C_FE_OK Then
        Call App.feAppMessage(C_FCM_ERROR, "Could not build the rigid-element set.")
        GoTo Cleanup
    End If

    rigidCount = rigidSet.Count()
    If rigidCount <= 0 Then
        Call App.feAppMessage(C_FCM_WARNING, "No rigid elements found.")
        GoTo Cleanup
    End If

    ' Allocate and initialize the RBE2 group first.
    ' It is Put before asking for the next empty ID so RBE3 gets a different ID.
    gid2 = g2.NextEmptyID()
    g2.title = "RBE2"

    rbe2Count = 0
    rbe3Count = 0

    rc = rigidSet.Reset()
    eid = rigidSet.Next()

    Do While eid <> 0
        rc = rbe.Get(eid)
        If rc = C_FE_OK Then
            If rbe.RigidInterpolate Then
                ' RBE3 membership is added later, after g2 has been stored.
                rbe3Count = rbe3Count + 1
            Else
                rc = g2.RangeAdd(C_FGR_ELEM_ID, eid, 0, 1, 1)
                If rc = C_FE_OK Then
                    rbe2Count = rbe2Count + 1
                End If
            End If
        End If
        eid = rigidSet.Next()
    Loop

    rc = g2.Put(gid2)
    If rc <> C_FE_OK Then
        Call App.feAppMessage(C_FCM_ERROR, "Could not create the RBE2 group.")
        GoTo Cleanup
    End If

    ' Now that RBE2 exists in the database, the next empty ID is distinct.
    gid3 = g3.NextEmptyID()
    g3.title = "RBE3"

    ' Second pass: avoids storing RBE3 IDs in a dynamic array.
    rbe3Count = 0
    rc = rigidSet.Reset()
    eid = rigidSet.Next()

    Do While eid <> 0
        rc = rbe.Get(eid)
        If rc = C_FE_OK Then
            If rbe.RigidInterpolate Then
                rc = g3.RangeAdd(C_FGR_ELEM_ID, eid, 0, 1, 1)
                If rc = C_FE_OK Then
                    rbe3Count = rbe3Count + 1
                End If
            End If
        End If
        eid = rigidSet.Next()
    Loop

    rc = g3.Put(gid3)
    If rc <> C_FE_OK Then
        Call App.feAppMessage(C_FCM_ERROR, "Could not create the RBE3 group.")
        GoTo Cleanup
    End If

    msg = "Rigid elements sorted: RBE2=" & CStr(rbe2Count) & ", RBE3=" & CStr(rbe3Count) & "."
    Call App.feAppMessage(C_FCM_NORMAL, msg)

Cleanup:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If errNumber <> 0 Then
        If Len(errDescription) > 100 Then errDescription = Left$(errDescription, 100)
        If Not App Is Nothing Then
            msg = "PARS RBE group sort error " & CStr(errNumber) & ": " & errDescription
            Call App.feAppMessage(C_FCM_ERROR, msg)
        End If
    End If

    Set g3 = Nothing
    Set g2 = Nothing
    Set rbe = Nothing
    Set rigidSet = Nothing
    Set App = Nothing
    On Error GoTo 0
End Sub
