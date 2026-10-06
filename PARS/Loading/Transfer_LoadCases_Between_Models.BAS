' =============================================================================
' Description: Transfers load cases between FEMAP models open in the same modern FEMAP session.
' Compatibility: Modern FEMAP only; the script shows an early compatibility message on N4W 2004.
' =============================================================================

Option Explicit

Const DBUG As Boolean = True
Const C_FE_OK As Long = -1
Const C_FT_NODE As Long = 7
Const C_FT_OUT_CASE As Long = 28
Const C_FCM_NORMAL As Long = 0
Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim rc As Long
    Dim i As Long, j As Long, k As Long
    Dim numModels As Long
    Dim modelIDs As Variant

    'The multi-model application API was added long after Nastran for Windows 2004.
    'Gate N4W by the legacy version property before dispatching any modern method.
    If App.Info_Version < 10# Then GoTo NotPortable

    On Error GoTo NotPortable
    rc = App.feAppGetAllModels(numModels, modelIDs)
    On Error GoTo 0

    If numModels < 2 Then
        MsgBox "Open at least two FEMAP models before running this script.", 48, "Transfer Load Cases Between Models"
        GoTo Cleanup
    End If

    er = 0
    If Not DBUG Then On Error GoTo Cleanup

    Dim sourceIndex As Long
    Dim targetIndex As Long
    sourceIndex = SelectModelIndex(App, modelIDs, numModels, 0, "Select model to copy load cases from:")
    If sourceIndex < 0 Then
        er = 1
        GoTo Cleanup
    End If

    Do
        targetIndex = SelectModelIndex(App, modelIDs, numModels, 1, "Select model to copy load cases to:")
        If targetIndex < 0 Then
            er = 2
            GoTo Cleanup
        End If
    Loop While targetIndex = sourceIndex

    Dim sourceModel As Variant
    Dim targetModel As Variant
    sourceModel = modelIDs(sourceIndex)
    targetModel = modelIDs(targetIndex)

    rc = App.feAppSetModel(sourceModel)

    Dim LCSet As Object
    Dim ndset As Object
    Set LCSet = App.feSet()
    Set ndset = App.feSet()

    rc = LCSet.Select(C_FT_OUT_CASE, True, "Select Output Sets to be transformed:")
    If (rc <> C_FE_OK) Or (LCSet.Count() < 1) Then
        er = 3
        GoTo Cleanup
    End If

    rc = ndset.Select(C_FT_NODE, True, "Select nodes for transformation:")
    If (rc <> C_FE_OK) Or (ndset.Count() < 1) Then
        er = 4
        GoTo Cleanup
    End If

    Dim ndIDs() As Long
    Dim f1() As Double, f2() As Double, f3() As Double
    Dim m1() As Double, m2() As Double, m3() As Double
    Dim nNodes As Long
    nNodes = ndset.Count()
    ReDim ndIDs(nNodes - 1)
    ReDim f1(nNodes - 1)
    ReDim f2(nNodes - 1)
    ReDim f3(nNodes - 1)
    ReDim m1(nNodes - 1)
    ReDim m2(nNodes - 1)
    ReDim m3(nNodes - 1)

    rc = ndset.Reset()
    For i = 0 To nNodes - 1
        ndIDs(i) = ndset.Next()
    Next

    Dim vecIDs(5) As Long
    vecIDs(0) = 192
    vecIDs(1) = 193
    vecIDs(2) = 194
    vecIDs(3) = 196
    vecIDs(4) = 197
    vecIDs(5) = 198

    Dim outSet As Object
    Dim vec As Object
    Dim LCnew As Object
    Dim Ld As Object
    Set outSet = App.feOutputSet()
    Set vec = App.feOutput()
    Set LCnew = App.feLoadSet()
    Set Ld = App.feLoadMesh()

    Dim outSetID As Long
    rc = LCSet.Reset()
    For i = 0 To LCSet.Count() - 1
        rc = App.feAppSetModel(sourceModel)
        outSetID = LCSet.Next()
        rc = outSet.Get(outSetID)
        If rc <> C_FE_OK Then
            er = 5
            GoTo Cleanup
        End If

        'Read force and moment vectors from the source output set using the old Output SetID/Get API.
        For j = 0 To 5
            vec.setID = outSet.ID
            rc = vec.Get(vecIDs(j))
            If rc <> C_FE_OK Then
                er = 6
                GoTo Cleanup
            End If
            For k = 0 To nNodes - 1
                Select Case j
                Case 0
                    f1(k) = vec.Value(ndIDs(k))
                Case 1
                    f2(k) = vec.Value(ndIDs(k))
                Case 2
                    f3(k) = vec.Value(ndIDs(k))
                Case 3
                    m1(k) = vec.Value(ndIDs(k))
                Case 4
                    m2(k) = vec.Value(ndIDs(k))
                Case 5
                    m3(k) = vec.Value(ndIDs(k))
                End Select
            Next
        Next

        rc = App.feAppSetModel(targetModel)

        LCnew.ID = LCnew.NextEmptyID()
        LCnew.title = outSet.title
        rc = LCnew.Put(LCnew.ID)
        If rc <> C_FE_OK Then
            er = 7
            GoTo Cleanup
        End If

        'The original script creates nodal force loads. Preserve that behavior.
        Ld.CSys = 0
        Ld.setID = LCnew.ID
        Ld.XOn = True
        Ld.YOn = True
        Ld.ZOn = True
        For k = 0 To nNodes - 1
            Ld.ID = Ld.NextEmptyID()
            Ld.meshID = ndIDs(k)
            Ld.type = 1
            Ld.load(0) = f1(k)
            Ld.load(1) = f2(k)
            Ld.load(2) = f3(k)
            rc = Ld.Put(Ld.ID)
        Next
    Next

    er = 0
    GoTo Cleanup

NotPortable:
    On Error Resume Next
    MsgBox "This script requires the FEMAP multi-model API introduced after Nastran for Windows 2004 and is not portable to that version.", 48, "Not portable to N4W 2004"
    Set App = Nothing
    Exit Sub

Cleanup:
    On Error Resume Next
    Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    Set Ld = Nothing
    Set LCnew = Nothing
    Set vec = Nothing
    Set outSet = Nothing
    Set ndset = Nothing
    Set LCSet = Nothing
    Set App = Nothing
End Sub

Private Function SelectModelIndex(ByRef App As Object, ByVal modelIDs As Variant, ByVal numModels As Long, ByVal preselect As Long, ByVal caption As String) As Long
    Dim i As Long
    Dim rc As Long
    Dim modelName As String
    Dim prompt As String
    Dim answer As String
    Dim idx As Long

    prompt = caption & vbLf & vbLf
    For i = 0 To numModels - 1
        modelName = ""
        rc = App.feAppGetModelName(modelIDs(i), modelName)
        prompt = prompt & CStr(i + 1) & " - " & BaseFileName(modelName) & vbLf
    Next

    Do
        answer = InputBox(prompt, caption, CStr(preselect + 1))
        If answer = "" Then
            SelectModelIndex = -1
            Exit Function
        End If
        If IsNumeric(answer) Then
            idx = CLng(answer) - 1
            If (idx >= 0) And (idx < numModels) Then
                SelectModelIndex = idx
                Exit Function
            End If
        End If
    Loop
End Function

Private Function BaseFileName(ByVal fileName As String) As String
    Dim i As Long
    For i = Len(fileName) To 1 Step -1
        If Mid$(fileName, i, 1) = "\\" Then
            BaseFileName = Mid$(fileName, i + 1, Len(fileName) - i)
            Exit Function
        End If
    Next
    If fileName = "" Then
        BaseFileName = "Untitled"
    Else
        BaseFileName = fileName
    End If
End Function
