' =============================================================================
' Description: Moves model nodes by the X/Y/Z displacement vectors of the active output set.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1

Sub Main
    Dim App As Object
    Dim ndObj As Object
    Dim xData As Object
    Dim yData As Object
    Dim zData As Object
    Dim view As Object
    Dim viewID As Long
    Dim setID As Long
    Dim rc As Long

    Set App = feGetObject()
    Set ndObj = App.feNode()
    Set xData = App.feOutput()
    Set yData = App.feOutput()
    Set zData = App.feOutput()
    Set view = App.feView()

    rc = App.feAppGetActiveView(viewID)
    If rc <> C_FE_OK Then GoTo Done
    rc = view.Get(viewID)
    If rc <> C_FE_OK Then GoTo Done

    setID = view.OutputSet
    If setID <= 0 Then
        MsgBox "No current output set in the active view.", 48, "Move Nodes by Deformation"
        GoTo Done
    End If

    If Not LoadVector(xData, setID, 2) Then GoTo MissingVector
    If Not LoadVector(yData, setID, 3) Then GoTo MissingVector
    If Not LoadVector(zData, setID, 4) Then GoTo MissingVector

    rc = ndObj.First()
    Do While rc = C_FE_OK
        ndObj.x = ndObj.x + xData.Value(ndObj.ID)
        ndObj.y = ndObj.y + yData.Value(ndObj.ID)
        ndObj.z = ndObj.z + zData.Value(ndObj.ID)
        rc = ndObj.Put(ndObj.ID)
        If rc <> C_FE_OK Then GoTo Done
        rc = ndObj.Next()
    Loop

    rc = App.feViewRegenerate(0)
    GoTo Done

MissingVector:
    MsgBox "Displacement output vectors 2, 3 and 4 are required in output set " & CStr(setID) & ".", 48, "Move Nodes by Deformation"
Done:
    Set view = Nothing
    Set zData = Nothing
    Set yData = Nothing
    Set xData = Nothing
    Set ndObj = Nothing
    Set App = Nothing
End Sub

Function LoadVector(ByRef outData As Object, ByVal setID As Long, ByVal vectorID As Long) As Boolean
    outData.SetID = setID
    LoadVector = (outData.Get(vectorID) = C_FE_OK)
End Function
