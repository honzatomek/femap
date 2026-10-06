' =============================================================================
' Description: Creates massObj elements at selected nodes using a selected massObj property.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' N4W/FEMAP portable constants (do not rely on type-library enum names)
Const C_FE_FAIL As Long = 0
Const C_FT_ELEM As Long = 8
Const C_FT_PROP As Long = 11
Const C_FT_GROUP As Long = 24
Const C_FGD_PROP_BYTYPE As Long = 33
Const C_FET_L_MASS As Long = 27
Const C_FTO_POINT1 As Long = 9
Const C_MSG_NORMAL As Long = 0

'Adds massObj Elements to selected nodes
Const DBUG As Boolean = False
Private er As Long
Private massObj As Object

Sub Main()
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, i As Long

    er = 0

    If DBUG Then
        On Error GoTo 0
    Else
        On Error GoTo Cleanup
    End If

    Dim ndset As Object
    Set ndset = App.feSet()

    Dim el As Object

    Dim eset As Object
    Set eset = App.feSet()

    Set massObj = App.feSet()

    Dim g As Object
    Set g = App.feGroup()

    rc = g.Get(App.Info_ActiveID(C_FT_GROUP))

    rc = massObj.AddRule(C_FET_L_MASS, C_FGD_PROP_BYTYPE)
    If (rc <> -1) Or (massObj.Count() < 1) Then
        er = 1
        GoTo Cleanup
    End If
    If DBUG Then Debug.Print "Number of massObj elements: " & massObj.Count()

    Dim pset As Object
    Set pset = App.feSet()
    Dim pID As Long

    Do While ndset.Select(7, True, "Select nodes for massObj Element:") = -1
        i = -1
        rc = -1
        Do While rc = -1
            i = i + 1
            rc = pset.Clear()
            rc = pset.SelectID(C_FT_PROP, GetMass(i), pID)
            If (rc <> -1) Or (pset.Count() < 1) Then
                er = 2
                GoTo Cleanup
            End If
            If pset.CountCommon(massObj.ID) = 1 Then rc = 0
        Loop

        rc = g.Get(App.Info_ActiveID(C_FT_GROUP))
        rc = ndset.Reset()
        Set el = App.feElem()
        el.type = C_FET_L_MASS
        el.propID = pID
        el.color = 54
        el.topology = C_FTO_POINT1

        For i = 0 To ndset.Count() - 1 Step 1
            rc = ndset.Next()
            If rc = C_FE_FAIL Then
                er = 3
                GoTo Cleanup
            End If
            el.ID = el.NextEmptyID()
            el.Node(0) = rc
            rc = el.Put(el.ID)
            If rc <> -1 Then
                er = 4
                GoTo Cleanup
            End If
            rc = eset.Add(el.ID)
        Next

        If eset.Count() > 0 Then
            rc = g.SetAdd(C_FT_ELEM, eset.ID)
            rc = g.Put(g.ID)
            rc = eset.Clear()
        End If
        App.feViewRegenerate(0)
    Loop

Cleanup:
    If DBUG Then App.feAppMessage(C_MSG_NORMAL, "The script exited with code " & er)
    On Error Resume Next
    If (Not DBUG) And (er <> 0) Then
        If eset.Count > 0 Then rc = App.feDelete(C_FT_ELEM, eset.ID)
    ElseIf DBUG And (er <> 0) Then
        If eset.Count > 0 Then rc = App.feViewShow(C_FT_ELEM, eset.ID)
    End If
    Set eset = Nothing
    Set el = Nothing
    Set pset = Nothing
    Set ndset = Nothing
    Set massObj = Nothing
    Set App = Nothing
End Sub

Private Function GetMass(ByVal i As Long) As String
    Dim txt As String
    Dim j As Long
    Dim rc As Long

    Select Case i
    Case 0
        txt = "Select massObj Property:"
    Case 1
        txt = "Select massObj Property:"
    Case 2
        txt = "Only massObj Property can be selected:"
    Case 3 To 9
        txt = "massObj massObj massObj Property MUST be SELECTED!"
    Case 10
        txt = "Just enjoy yourself.. Best by selecting massObj Property:"
    Case 20, 30, 40, 50
        txt = "This is already the " & i & "st try, press Cancel or select massObj Property:"
    Case Else
        txt = "["
        rc = massObj.Reset()
        For j = 1 To massObj.Count() Step 1
            txt = txt & massObj.Next() & ","
        Next
        txt = Left(txt, Len(txt) - 1) & "]"
        txt = "Only following massObj Properties can be selected: " & txt
    End Select
    GetMass = txt
End Function
