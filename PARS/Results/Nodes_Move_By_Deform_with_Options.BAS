' =============================================================================
' Description: Moves model nodes by selected displacement components with a user-defined scale factor.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1

Sub Main
    Begin Dialog UserDialog 300,154,"Move Nodes by Deformation" ' %GRID:10,7,1,1
        CheckBox 40,21,90,14,"Update X",.doX,1
        CheckBox 40,42,90,14,"Update Y",.doY,1
        CheckBox 40,63,90,14,"Update Z",.doZ,1
        TextBox 140,77,120,21,.scaleText
        Text 30,84,90,14,"Scale Factor:",.scaleLabel,1
        OKButton 40,112,90,21
        CancelButton 140,112,90,21
    End Dialog

    Dim dlg As UserDialog
    Dim App As Object
    Dim ndObj As Object
    Dim xData As Object
    Dim yData As Object
    Dim zData As Object
    Dim view As Object
    Dim viewID As Long
    Dim setID As Long
    Dim rc As Long
    Dim scaleFactor As Double
    Dim msg As String

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

    dlg.doX = 1
    dlg.doY = 1
    dlg.doZ = 1
    dlg.scaleText = "1.00"
    If Dialog(dlg) = 0 Then GoTo Done

    If (dlg.doX + dlg.doY + dlg.doZ) = 0 Then
        MsgBox "No components selected.", 48, "Move Nodes by Deformation"
        GoTo Done
    End If

    scaleFactor = Val(dlg.scaleText)

    If dlg.doX = 1 Then
        If Not LoadVector(xData, setID, 2) Then GoTo MissingVector
    End If
    If dlg.doY = 1 Then
        If Not LoadVector(yData, setID, 3) Then GoTo MissingVector
    End If
    If dlg.doZ = 1 Then
        If Not LoadVector(zData, setID, 4) Then GoTo MissingVector
    End If

    msg = "Output Set #" & CStr(setID) & ":"
    If dlg.doX = 1 Then msg = msg & " X"
    If dlg.doY = 1 Then msg = msg & " Y"
    If dlg.doZ = 1 Then msg = msg & " Z"
    msg = msg & ", scale = " & CStr(scaleFactor)
    Call App.feAppMessage(0, msg)

    rc = ndObj.First()
    Do While rc = C_FE_OK
        If dlg.doX = 1 Then ndObj.x = ndObj.x + scaleFactor * xData.Value(ndObj.ID)
        If dlg.doY = 1 Then ndObj.y = ndObj.y + scaleFactor * yData.Value(ndObj.ID)
        If dlg.doZ = 1 Then ndObj.z = ndObj.z + scaleFactor * zData.Value(ndObj.ID)
        rc = ndObj.Put(ndObj.ID)
        If rc <> C_FE_OK Then GoTo Done
        rc = ndObj.Next()
    Loop

    rc = App.feViewRegenerate(0)
    GoTo Done

MissingVector:
    MsgBox "One of the requested displacement output vectors is not available in output set " & CStr(setID) & ".", 48, "Move Nodes by Deformation"
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
