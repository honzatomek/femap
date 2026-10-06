' =============================================================================
' Description: Creates a group of elements that remain in compression across selected output-set combinations.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

'Creates a group containing elements that remain in compression for the
'dead-load output set combined with every selected additional output set.
'Uses only API calls available in the old Nastran for Windows/FEMAP API.

Const DBUG As Boolean = False
Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FT_GROUP As Long = 24
Const C_FT_OUT_CASE As Long = 28
Const C_FCM_NORMAL As Long = 0

Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim rc As Long
    Dim i As Long
    Dim j As Long
    Dim deadSetID As Long
    Dim outSetID As Long
    Dim currentGroupID As Long
    Dim nextGroupID As Long
    Dim tempGroupID As Long

    er = 0
    If Not DBUG Then On Error GoTo Cleanup

    Dim deadSet As Object
    Set deadSet = App.feOutputSet()
    Dim outSet As Object
    Set outSet = App.feOutputSet()
    Dim outSets As Object
    Set outSets = App.feSet()

    Dim deadVec As Object
    Set deadVec = App.feOutput()
    Dim outVec As Object
    Set outVec = App.feOutput()

    Dim elemSet As Object
    Set elemSet = App.feSet()
    Dim gBase As Object
    Set gBase = App.feGroup()
    Dim gTemp As Object
    Set gTemp = App.feGroup()
    Dim gResult As Object
    Set gResult = App.feGroup()

    Dim listCount As Long
    Dim elemIDs As Variant
    Dim deadValues As Variant

    '1. Select the dead-load output set.
    If DBUG Then
        deadSetID = 11
        rc = deadSet.Get(deadSetID)
    Else
        rc = outSets.SelectID(C_FT_OUT_CASE, "Select Dead Load Output Set:", deadSetID)
        If rc <> C_FE_OK Then
            er = 1
            GoTo Cleanup
        End If
        rc = deadSet.Get(deadSetID)
    End If
    If rc <> C_FE_OK Then
        er = 1
        GoTo Cleanup
    End If

    '2. Select all additional output sets. Set.Select is available in old FEMAP.
    rc = outSets.Clear()
    rc = outSets.Select(C_FT_OUT_CASE, True, "Select Output Sets other than Dead Load:")
    If (rc <> C_FE_OK) Or (outSets.Count() < 1) Then
        er = 2
        GoTo Cleanup
    End If

    '3. Read vector 3022 from the dead-load set using the old SetID/Get pattern.
    deadVec.setID = deadSet.ID
    rc = deadVec.Get(3022)
    If rc <> C_FE_OK Then
        er = 3
        GoTo Cleanup
    End If
    rc = deadVec.GetOutputList(listCount, elemIDs, deadValues)
    If rc <> C_FE_OK Then
        er = 3
        GoTo Cleanup
    End If

    'Create the first temporary group: elements in compression in dead load alone.
    rc = elemSet.Clear()
    For i = 0 To listCount - 1
        If deadValues(i) < 0# Then rc = elemSet.Add(elemIDs(i))
    Next
    gBase.ID = gBase.NextEmptyID()
    gBase.title = "Temp compression - dead load"
    rc = gBase.SetAdd(C_FT_ELEM, elemSet.ID)
    rc = gBase.Put(gBase.ID)
    If rc <> C_FE_OK Then
        er = 3
        GoTo Cleanup
    End If
    currentGroupID = gBase.ID

    '4. For each selected set, create the compression group for dead+selected
    '   and intersect it with the running result using old feGroupBoolean(And).
    rc = outSets.Reset()
    For i = 0 To outSets.Count() - 1
        outSetID = outSets.Next()
        rc = outSet.Get(outSetID)
        If rc <> C_FE_OK Then
            er = 4
            GoTo Cleanup
        End If

        outVec.setID = outSet.ID
        rc = outVec.Get(3022)
        If rc <> C_FE_OK Then
            er = 4
            GoTo Cleanup
        End If

        rc = elemSet.Clear()
        For j = 0 To listCount - 1
            If (deadValues(j) + outVec.Value(elemIDs(j))) < 0# Then
                rc = elemSet.Add(elemIDs(j))
            End If
        Next

        gTemp.ID = gTemp.NextEmptyID()
        gTemp.title = "Temp compression - combination"
        rc = gTemp.SetAdd(C_FT_ELEM, elemSet.ID)
        rc = gTemp.Put(gTemp.ID)
        If rc <> C_FE_OK Then
            er = 4
            GoTo Cleanup
        End If
        tempGroupID = gTemp.ID

        nextGroupID = gResult.NextEmptyID()
        rc = App.feGroupBoolean(0, currentGroupID, tempGroupID, nextGroupID, "Temp compression intersection")
        If rc <> C_FE_OK Then
            er = 5
            GoTo Cleanup
        End If

        If currentGroupID <> gBase.ID Then Call App.feDelete(C_FT_GROUP, currentGroupID)
        Call App.feDelete(C_FT_GROUP, tempGroupID)
        currentGroupID = nextGroupID
    Next

    '5. Keep only the final intersection and give it the requested title.
    If currentGroupID <> gBase.ID Then Call App.feDelete(C_FT_GROUP, gBase.ID)
    rc = gResult.Get(currentGroupID)
    If rc <> C_FE_OK Then
        er = 5
        GoTo Cleanup
    End If
    gResult.title = "ELEMENTS IN COMPRESSION"
    rc = gResult.Put(currentGroupID)
    If rc <> C_FE_OK Then er = 5

Cleanup:
    If er <> 0 Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    On Error Resume Next
    Set gResult = Nothing
    Set gTemp = Nothing
    Set gBase = Nothing
    Set elemSet = Nothing
    Set outVec = Nothing
    Set deadVec = Nothing
    Set outSets = Nothing
    Set outSet = Nothing
    Set deadSet = Nothing
    Set App = Nothing
End Sub
