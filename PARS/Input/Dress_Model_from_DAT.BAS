' =============================================================================
' Description: Dresses the current FEMAP model from FEMAP comments in a Nastran
'              DAT/BDF file. Updates existing coordinate-system, load-set,
'              constraint-set and property titles, restores standard beam/bar shape metadata
'              and orientation, and creates named Groups from GROUP/SET1 metadata comments.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FT_NODE As Long = 7
Const C_FT_ELEM As Long = 8
Const C_FT_CSYS As Long = 9
Const C_FT_MATL As Long = 10
Const C_FT_PROP As Long = 11

' Group Range Types used by feGroup.RangeAdd.
Const C_FGR_CSYS_ID As Long = 0
Const C_FGR_NODE_ID As Long = 17
Const C_FGR_ELEM_ID As Long = 21
Const C_FGR_MATL_ID As Long = 26
Const C_FGR_PROP_ID As Long = 30

Sub Main()
    Dim App As Object
    Dim fileName As String
    Dim rc As Long
    Dim f As Object
    Dim lineText As String

    Dim csysNames As Long
    Dim loadNames As Long
    Dim constraintNames As Long
    Dim propNames As Long
    Dim propShapes As Long
    Dim propOrients As Long
    Dim groupsCreated As Long
    Dim groupMembers As Long
    Dim missing As Long
    Dim malformed As Long
    Dim errNum As Long
    Dim errDesc As String

    On Error GoTo ErrorHandler

    Set App = feGetObject()

    fileName = ""
    rc = App.feFileGetName("Select Nastran DAT/BDF file", "Nastran DAT/BDF Files", "*.dat;*.bdf", True, fileName)
    If rc <> C_FE_OK Then GoTo ExitHere
    If fileName = "" Then GoTo ExitHere

    Set f = App.feRead
    rc = f.Open(fileName, 100)
    If rc <> C_FE_OK Then
        MsgBox "Could not open Nastran DAT/BDF file.", 48, "Dress Model from DAT"
        GoTo ExitHere
    End If

    Do While Not f.AtEOF()
        rc = f.Read()
        If rc <> C_FE_OK Then Exit Do
        lineText = f.Line

        If IsCommentLine(lineText) Then
            If InStr(1, lineText, "Coordinate System ") > 0 Then
                rc = UpdateNamedEntity(App, lineText, "Coordinate System ", 1)
                If rc = 1 Then
                    csysNames = csysNames + 1
                ElseIf rc = -1 Then
                    missing = missing + 1
                Else
                    malformed = malformed + 1
                End If

            ElseIf InStr(1, lineText, "Load Set ") > 0 Then
                rc = UpdateNamedEntity(App, lineText, "Load Set ", 2)
                If rc = 1 Then
                    loadNames = loadNames + 1
                ElseIf rc = -1 Then
                    missing = missing + 1
                Else
                    malformed = malformed + 1
                End If

            ElseIf InStr(1, lineText, "Constraint Set ") > 0 Then
                rc = UpdateNamedEntity(App, lineText, "Constraint Set ", 3)
                If rc = 1 Then
                    constraintNames = constraintNames + 1
                ElseIf rc = -1 Then
                    missing = missing + 1
                Else
                    malformed = malformed + 1
                End If

            ElseIf InStr(1, lineText, "PropShape ") > 0 Then
                rc = ApplyPropShape(App, lineText)
                If rc = 1 Then
                    propShapes = propShapes + 1
                ElseIf rc = -1 Then
                    missing = missing + 1
                Else
                    malformed = malformed + 1
                End If

            ElseIf InStr(1, lineText, "PropOrient ") > 0 Then
                rc = ApplyPropOrient(App, lineText)
                If rc = 1 Then
                    propOrients = propOrients + 1
                ElseIf rc = -1 Then
                    missing = missing + 1
                Else
                    malformed = malformed + 1
                End If

            ElseIf InStr(1, lineText, "Property ") > 0 Then
                rc = UpdatePropertyTitle(App, lineText)
                If rc = 1 Then
                    propNames = propNames + 1
                ElseIf rc = -1 Then
                    missing = missing + 1
                Else
                    malformed = malformed + 1
                End If
            End If
        End If
    Loop

    rc = f.Close()
    Set f = Nothing

    rc = CreateGroupsFromSetComments(App, fileName, groupsCreated, groupMembers, missing, malformed)
    If rc <> C_FE_OK Then malformed = malformed + 1

    App.feViewRegenerate(0)

    MsgBox "Model dressing complete." & Chr$(13) & Chr$(10) & _
           "Coordinate systems renamed: " & CStr(csysNames) & Chr$(13) & Chr$(10) & _
           "Load sets renamed: " & CStr(loadNames) & Chr$(13) & Chr$(10) & _
           "Constraint sets renamed: " & CStr(constraintNames) & Chr$(13) & Chr$(10) & _
           "Properties renamed: " & CStr(propNames) & Chr$(13) & Chr$(10) & _
           "Property shapes dressed: " & CStr(propShapes) & Chr$(13) & Chr$(10) & _
           "Property orientations dressed: " & CStr(propOrients) & Chr$(13) & Chr$(10) & _
           "Named groups created/updated: " & CStr(groupsCreated) & Chr$(13) & Chr$(10) & _
           "Group members added: " & CStr(groupMembers) & Chr$(13) & Chr$(10) & _
           "Referenced IDs not present in current model: " & CStr(missing) & Chr$(13) & Chr$(10) & _
           "Malformed recognized comments skipped: " & CStr(malformed), _
           64, "Dress Model from DAT"

    GoTo ExitHere

ErrorHandler:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    rc = f.Close()
    Set f = Nothing
    MsgBox "Error " & CStr(errNum) & ": " & errDesc, 48, "Dress Model from DAT"

ExitHere:
    On Error Resume Next
    rc = f.Close()
    Set f = Nothing
    Set App = Nothing
End Sub

' Returns 1=updated, -1=entity does not exist, 0=malformed/failed.
Private Function UpdateNamedEntity(ByRef App As Object, ByVal lineText As String, ByVal marker As String, ByVal kind As Long) As Long
    Dim entityID As Long
    Dim entityName As String
    Dim obj As Object
    Dim rc As Long

    UpdateNamedEntity = 0
    If Not ParseIDAndText(lineText, marker, entityID, entityName) Then Exit Function

    If kind = 1 Then
        Set obj = App.feCSys()
    ElseIf kind = 2 Then
        Set obj = App.feLoadSet()
    ElseIf kind = 3 Then
        Set obj = App.feBCSet()
    Else
        Exit Function
    End If

    rc = obj.Get(entityID)
    If rc <> C_FE_OK Then
        UpdateNamedEntity = -1
        Set obj = Nothing
        Exit Function
    End If

    obj.title = entityName
    rc = obj.Put(entityID)
    If rc = C_FE_OK Then UpdateNamedEntity = 1

    Set obj = Nothing
End Function

' Returns 1=updated, -1=property does not exist, 0=malformed/failed.
Private Function UpdatePropertyTitle(ByRef App As Object, ByVal lineText As String) As Long
    Dim propID As Long
    Dim propName As String
    Dim p As Object
    Dim rc As Long

    UpdatePropertyTitle = 0
    If Not ParseIDAndText(lineText, "Property ", propID, propName) Then Exit Function

    Set p = App.feProp()
    rc = p.Get(propID)
    If rc <> C_FE_OK Then
        UpdatePropertyTitle = -1
        Set p = Nothing
        Exit Function
    End If

    p.title = propName
    rc = p.Put(propID)
    If rc = C_FE_OK Then UpdatePropertyTitle = 1

    Set p = Nothing
End Function

' PropShape format written by FEMAP:
'   $ ... PropShape <propertyID> : <shapeID>,<end>,DIM1,DIM2,DIM3,DIM4,DIM5,DIM6
'
' The legacy Property object stores standard-shape metadata as:
'   flagI(1) = shape type
'   End A dimensions = pval(40..45)
'   End B dimensions = pval(55..60)
'
' Returns 1=updated, -1=property does not exist, 0=malformed/failed.
Private Function ApplyPropShape(ByRef App As Object, ByVal lineText As String) As Long
    Dim propID As Long
    Dim dataText As String
    Dim values(7) As Double
    Dim n As Long
    Dim shapeID As Long
    Dim shapeEnd As Long
    Dim p As Object
    Dim rc As Long
    Dim i As Long
    Dim baseIndex As Long

    ApplyPropShape = 0
    If Not ParseIDAndText(lineText, "PropShape ", propID, dataText) Then Exit Function

    n = ParseCSVNumbers(dataText, values, 8)
    If n < 8 Then Exit Function

    shapeID = values(0)
    shapeEnd = values(1)
    If shapeEnd <> 0 And shapeEnd <> 1 Then Exit Function

    Set p = App.feProp()
    rc = p.Get(propID)
    If rc <> C_FE_OK Then
        ApplyPropShape = -1
        Set p = Nothing
        Exit Function
    End If

    p.flagI(1) = shapeID

    If shapeEnd = 0 Then
        baseIndex = 40
    Else
        baseIndex = 55
    End If

    For i = 0 To 5
        p.pval(baseIndex + i) = values(2 + i)
    Next

    rc = p.Put(propID)
    If rc = C_FE_OK Then ApplyPropShape = 1

    Set p = Nothing
End Function

' PropOrient format written by FEMAP:
'   $ ... PropOrient <propertyID> :
'       <shapeID>,<end>,<orientation>,SR1,SR2,SR3,SR4,RefLoc,RefY,RefZ
'
' Legacy Property mapping:
'   End A: SR1..4=pval(46..49), Orient=50, RefLoc=51, RefY=52, RefZ=53
'   End B: SR1..4=pval(61..64), Orient=65, RefLoc=66, RefY=67, RefZ=68
'
' Returns 1=updated, -1=property does not exist, 0=malformed/failed.
Private Function ApplyPropOrient(ByRef App As Object, ByVal lineText As String) As Long
    Dim propID As Long
    Dim dataText As String
    Dim values(9) As Double
    Dim n As Long
    Dim shapeID As Long
    Dim shapeEnd As Long
    Dim p As Object
    Dim rc As Long
    Dim i As Long
    Dim srBase As Long
    Dim orientIndex As Long

    ApplyPropOrient = 0
    If Not ParseIDAndText(lineText, "PropOrient ", propID, dataText) Then Exit Function

    n = ParseCSVNumbers(dataText, values, 10)
    If n < 10 Then Exit Function

    shapeID = values(0)
    shapeEnd = values(1)
    If shapeEnd <> 0 And shapeEnd <> 1 Then Exit Function

    Set p = App.feProp()
    rc = p.Get(propID)
    If rc <> C_FE_OK Then
        ApplyPropOrient = -1
        Set p = Nothing
        Exit Function
    End If

    p.flagI(1) = shapeID

    If shapeEnd = 0 Then
        srBase = 46
        orientIndex = 50
    Else
        srBase = 61
        orientIndex = 65
    End If

    For i = 0 To 3
        p.pval(srBase + i) = values(3 + i)
    Next

    p.pval(orientIndex) = values(2)
    p.pval(orientIndex + 1) = values(7)
    p.pval(orientIndex + 2) = values(8)
    p.pval(orientIndex + 3) = values(9)

    rc = p.Put(propID)
    If rc = C_FE_OK Then ApplyPropOrient = 1

    Set p = Nothing
End Function

' Extracts the integer ID between marker and the following colon, and returns
' the trimmed text after the colon.
Private Function ParseIDAndText(ByVal lineText As String, ByVal marker As String, ByRef entityID As Long, ByRef tailText As String) As Boolean
    Dim p As Long
    Dim q As Long
    Dim idText As String

    ParseIDAndText = False
    entityID = 0
    tailText = ""

    p = InStr(1, lineText, marker)
    If p = 0 Then Exit Function

    p = p + Len(marker)
    q = InStr(p, lineText, ":")
    If q = 0 Then Exit Function

    idText = Trim$(Mid$(lineText, p, q - p))
    If idText = "" Then Exit Function

    entityID = Val(idText)
    If entityID <= 0 Then Exit Function

    tailText = Trim$(Mid$(lineText, q + 1, Len(lineText) - q))
    If tailText = "" Then Exit Function

    ParseIDAndText = True
End Function

' Simple comma-separated numeric parser for the FEMAP comment payloads.
' No VBA Split()/Array() dependency is used so the code remains compatible
' with the older Nastran for Windows BASIC interpreter.
Private Function ParseCSVNumbers(ByVal textValue As String, ByRef values() As Double, ByVal maxCount As Long) As Long
    Dim startPos As Long
    Dim commaPos As Long
    Dim token As String
    Dim count As Long

    startPos = 1
    count = 0

    Do While startPos <= Len(textValue) And count < maxCount
        commaPos = InStr(startPos, textValue, ",")

        If commaPos = 0 Then
            token = Trim$(Mid$(textValue, startPos, Len(textValue) - startPos + 1))
            If token <> "" Then
                values(count) = Val(token)
                count = count + 1
            End If
            Exit Do
        Else
            token = Trim$(Mid$(textValue, startPos, commaPos - startPos))
            If token <> "" Then
                values(count) = Val(token)
                count = count + 1
            Else
                values(count) = 0#
                count = count + 1
            End If
            startPos = commaPos + 1
        End If
    Loop

    ParseCSVNumbers = count
End Function


' =============================================================================
' GROUP metadata + SET1 -> Nastran4Win Group support
'
' Metadata convention:
'   $ GROUP 101 : LOWER SUPPORTS : NODE
'   SET1        1101   ...
'
'   $ GROUP 101 : LOWER SUPPORTS : ELEMENT
'   SET1        2101   ...
'
' GROUP ID and SET1 SID are independent.  The first occurrence creates/names
' the Nastran4Win Group; later occurrences with the same GROUP ID retrieve that
' Group and append members.  Supported object names:
'   NODE, ELEMENT, CSYS, MATERIAL, PROPERTY
' =============================================================================

Private Function CreateGroupsFromSetComments(ByRef App As Object, ByVal fileName As String, _
                                             ByRef groupsCreated As Long, ByRef membersAdded As Long, _
                                             ByRef missing As Long, ByRef malformed As Long) As Long
    Dim f As Object
    Dim rc As Long
    Dim lineText As String
    Dim handled As Boolean

    Dim pendingType As Long
    Dim pendingGroupID As Long
    Dim pendingName As String

    Dim groupActive As Boolean
    Dim groupObj As Object
    Dim activeType As Long
    Dim activeGroupID As Long
    Dim activeName As String

    Dim rangeState As Long
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim lastID As Long

    CreateGroupsFromSetComments = 0

    Set f = App.feRead
    rc = f.Open(fileName, 100)
    If rc <> C_FE_OK Then
        Set f = Nothing
        Exit Function
    End If

    Do While Not f.AtEOF()
        rc = f.Read()
        If rc <> C_FE_OK Then Exit Do
        lineText = f.Line
        handled = False

        ' SET1 continuation lines belong to the current GROUP block.
        If groupActive Then
            If IsSetContinuationLine(lineText) Then
                AddSet1Members groupObj, activeType, lineText, False, _
                               rangeState, rangeStart, rangeEnd, lastID, _
                               membersAdded, missing, malformed
                handled = True
            Else
                FinalizeGroupBlock groupObj, activeType, activeGroupID, activeName, _
                                   rangeState, rangeStart, rangeEnd, lastID, _
                                   groupsCreated, membersAdded, missing, malformed
                groupActive = False
                Set groupObj = Nothing
            End If
        End If

        If Not handled Then
            If IsCommentLine(lineText) Then
                ' Only GROUP metadata comments affect pending SET1 interpretation.
                If Left$(UCase$(Trim$(lineText)), 7) = "$ GROUP" Then
                    If pendingType <> 0 Then malformed = malformed + 1
                    pendingType = 0
                    pendingGroupID = 0
                    pendingName = ""

                    If Not ParseGroupComment(lineText, pendingGroupID, pendingName, pendingType) Then
                        malformed = malformed + 1
                    End If
                End If

            ElseIf Trim$(lineText) <> "" Then
                If pendingType <> 0 Then
                    If BulkCardName(lineText) = "SET1" And Set1ID(lineText) > 0 Then
                        Set groupObj = App.feGroup()
                        activeType = pendingType
                        activeGroupID = pendingGroupID
                        activeName = pendingName

                        rc = groupObj.Get(activeGroupID)
                        If rc = C_FE_OK Then
                            If groupObj.title <> "" And groupObj.title <> activeName Then
                                ' The metadata comment is authoritative, but a
                                ' name mismatch is still worth reporting.
                                malformed = malformed + 1
                            End If
                        End If
                        groupObj.title = activeName

                        rangeState = 0
                        rangeStart = 0
                        rangeEnd = 0
                        lastID = 0

                        AddSet1Members groupObj, activeType, lineText, True, _
                                       rangeState, rangeStart, rangeEnd, lastID, _
                                       membersAdded, missing, malformed
                        groupActive = True
                    Else
                        malformed = malformed + 1
                    End If

                    pendingType = 0
                    pendingGroupID = 0
                    pendingName = ""
                End If
            End If
        End If
    Loop

    If groupActive Then
        FinalizeGroupBlock groupObj, activeType, activeGroupID, activeName, _
                           rangeState, rangeStart, rangeEnd, lastID, _
                           groupsCreated, membersAdded, missing, malformed
        Set groupObj = Nothing
    End If

    If pendingType <> 0 Then malformed = malformed + 1

    rc = f.Close()
    Set f = Nothing
    CreateGroupsFromSetComments = C_FE_OK
End Function

Private Sub FinalizeGroupBlock(ByRef groupObj As Object, ByVal entityType As Long, _
                               ByVal groupID As Long, ByVal groupName As String, _
                               ByRef rangeState As Long, ByRef rangeStart As Long, _
                               ByRef rangeEnd As Long, ByRef lastID As Long, _
                               ByRef groupsCreated As Long, ByRef membersAdded As Long, _
                               ByRef missing As Long, ByRef malformed As Long)
    Dim rc As Long

    FlushSetRange groupObj, entityType, rangeState, rangeStart, rangeEnd, lastID, _
                  membersAdded, missing, malformed

    groupObj.title = groupName
    rc = groupObj.Put(groupID)
    If rc = C_FE_OK Then
        groupsCreated = groupsCreated + 1
    Else
        malformed = malformed + 1
    End If
End Sub

Private Function ParseGroupComment(ByVal lineText As String, ByRef groupID As Long, _
                                   ByRef groupName As String, ByRef entityType As Long) As Boolean
    Dim s As String
    Dim u As String
    Dim p As Long
    Dim q As Long
    Dim r As Long
    Dim idText As String
    Dim objectName As String

    ParseGroupComment = False
    groupID = 0
    groupName = ""
    entityType = 0

    s = Trim$(lineText)
    u = UCase$(s)
    If Left$(u, 7) <> "$ GROUP" Then Exit Function

    p = 8
    q = InStr(p, s, ":")
    If q = 0 Then Exit Function

    idText = Trim$(Mid$(s, p, q - p))
    groupID = CLng(Val(idText))
    If groupID <= 0 Then Exit Function

    r = InStr(q + 1, s, ":")
    If r = 0 Then Exit Function

    groupName = Trim$(Mid$(s, q + 1, r - q - 1))
    If groupName = "" Then Exit Function

    objectName = UCase$(Trim$(Mid$(s, r + 1, Len(s) - r)))
    If objectName = "NODE" Then
        entityType = C_FT_NODE
    ElseIf objectName = "ELEMENT" Then
        entityType = C_FT_ELEM
    ElseIf objectName = "CSYS" Then
        entityType = C_FT_CSYS
    ElseIf objectName = "MATERIAL" Then
        entityType = C_FT_MATL
    ElseIf objectName = "PROPERTY" Then
        entityType = C_FT_PROP
    Else
        entityType = 0
        Exit Function
    End If

    ParseGroupComment = True
End Function

Private Function BulkCardName(ByVal lineText As String) As String
    Dim s As String
    Dim p As Long

    s = LTrim$(lineText)
    p = InStr(1, s, ",")
    If p > 0 Then
        BulkCardName = UCase$(Trim$(Left$(s, p - 1)))
    Else
        BulkCardName = UCase$(Trim$(Left$(s, 8)))
    End If
End Function

Private Function Set1ID(ByVal lineText As String) As Long
    Dim token As String
    Dim ok As Boolean

    Set1ID = 0
    If InStr(1, lineText, ",") > 0 Then
        token = CSVField(lineText, 2, ok)
        If Not ok Then Exit Function
    Else
        If Len(lineText) < 9 Then Exit Function
        token = Trim$(Mid$(lineText, 9, 8))
    End If

    If token = "" Then Exit Function
    Set1ID = CLng(Val(token))
End Function

Private Function IsSetContinuationLine(ByVal lineText As String) As Boolean
    Dim s As String

    s = LTrim$(lineText)
    IsSetContinuationLine = (Len(s) > 0 And Left$(s, 1) = "+")
End Function

Private Sub AddSet1Members(ByRef groupObj As Object, ByVal entityType As Long, _
                           ByVal lineText As String, ByVal firstLine As Boolean, _
                           ByRef rangeState As Long, ByRef rangeStart As Long, _
                           ByRef rangeEnd As Long, ByRef lastID As Long, _
                           ByRef membersAdded As Long, ByRef missing As Long, _
                           ByRef malformed As Long)
    Dim i As Long
    Dim startField As Long
    Dim token As String
    Dim ok As Boolean
    Dim p As Long

    If InStr(1, lineText, ",") > 0 Then
        If firstLine Then
            startField = 3
        Else
            startField = 2
        End If

        i = startField
        Do
            token = CSVField(lineText, i, ok)
            If Not ok Then Exit Do
            ConsumeSetToken groupObj, entityType, token, _
                            rangeState, rangeStart, rangeEnd, lastID, _
                            membersAdded, missing, malformed
            i = i + 1
        Loop
    Else
        If firstLine Then
            p = 17
        Else
            p = 9
        End If

        Do While p <= 72 And p <= Len(lineText)
            token = Trim$(Mid$(lineText, p, 8))
            ConsumeSetToken groupObj, entityType, token, _
                            rangeState, rangeStart, rangeEnd, lastID, _
                            membersAdded, missing, malformed
            p = p + 8
        Loop
    End If
End Sub

Private Function CSVField(ByVal lineText As String, ByVal wantedIndex As Long, _
                          ByRef ok As Boolean) As String
    Dim fieldIndex As Long
    Dim startPos As Long
    Dim commaPos As Long

    ok = False
    CSVField = ""
    If wantedIndex <= 0 Then Exit Function

    fieldIndex = 1
    startPos = 1

    Do
        commaPos = InStr(startPos, lineText, ",")
        If fieldIndex = wantedIndex Then
            If commaPos = 0 Then
                CSVField = Trim$(Mid$(lineText, startPos, Len(lineText) - startPos + 1))
            Else
                CSVField = Trim$(Mid$(lineText, startPos, commaPos - startPos))
            End If
            ok = True
            Exit Function
        End If

        If commaPos = 0 Then Exit Function
        startPos = commaPos + 1
        fieldIndex = fieldIndex + 1
    Loop
End Function

Private Sub ConsumeSetToken(ByRef groupObj As Object, ByVal entityType As Long, _
                            ByVal token As String, _
                            ByRef rangeState As Long, ByRef rangeStart As Long, _
                            ByRef rangeEnd As Long, ByRef lastID As Long, _
                            ByRef membersAdded As Long, ByRef missing As Long, _
                            ByRef malformed As Long)
    Dim t As String
    Dim entityID As Long
    Dim stepValue As Long

    t = UCase$(Trim$(token))
    If t = "" Then Exit Sub

    ' State 2 means "A THRU B" is complete but might still be followed by BY.
    If rangeState = 2 And t <> "BY" Then
        ExpandSetRange groupObj, entityType, rangeStart, rangeEnd, 1, _
                       lastID, membersAdded, missing, malformed
        rangeState = 0
    End If

    If rangeState = 1 Then
        entityID = CLng(Val(t))
        If entityID <= 0 Then
            malformed = malformed + 1
            rangeState = 0
        Else
            rangeEnd = entityID
            rangeState = 2
        End If
        Exit Sub
    End If

    If rangeState = 2 Then
        If t = "BY" Then
            rangeState = 3
        Else
            malformed = malformed + 1
            rangeState = 0
        End If
        Exit Sub
    End If

    If rangeState = 3 Then
        stepValue = CLng(Val(t))
        If stepValue <= 0 Then
            malformed = malformed + 1
        Else
            ExpandSetRange groupObj, entityType, rangeStart, rangeEnd, stepValue, _
                           lastID, membersAdded, missing, malformed
        End If
        rangeState = 0
        Exit Sub
    End If

    If t = "THRU" Then
        If lastID <= 0 Then
            malformed = malformed + 1
        Else
            rangeStart = lastID
            rangeState = 1
        End If
        Exit Sub
    End If

    If t = "BY" Then
        malformed = malformed + 1
        Exit Sub
    End If

    entityID = CLng(Val(t))
    If entityID <= 0 Then
        malformed = malformed + 1
        Exit Sub
    End If

    AddGroupMember groupObj, entityType, entityID, membersAdded, missing
    lastID = entityID
End Sub

Private Sub FlushSetRange(ByRef groupObj As Object, ByVal entityType As Long, _
                          ByRef rangeState As Long, ByRef rangeStart As Long, _
                          ByRef rangeEnd As Long, ByRef lastID As Long, _
                          ByRef membersAdded As Long, ByRef missing As Long, _
                          ByRef malformed As Long)
    If rangeState = 2 Then
        ExpandSetRange groupObj, entityType, rangeStart, rangeEnd, 1, _
                       lastID, membersAdded, missing, malformed
    ElseIf rangeState <> 0 Then
        malformed = malformed + 1
    End If
    rangeState = 0
End Sub

Private Sub ExpandSetRange(ByRef groupObj As Object, ByVal entityType As Long, _
                           ByVal firstID As Long, ByVal lastRangeID As Long, _
                           ByVal stepValue As Long, ByRef lastID As Long, _
                           ByRef membersAdded As Long, ByRef missing As Long, _
                           ByRef malformed As Long)
    Dim signedStep As Long
    Dim entityID As Long

    If stepValue <= 0 Then
        malformed = malformed + 1
        Exit Sub
    End If

    signedStep = stepValue
    If lastRangeID < firstID Then signedStep = -signedStep

    entityID = firstID + signedStep
    Do While (signedStep > 0 And entityID <= lastRangeID) Or _
             (signedStep < 0 And entityID >= lastRangeID)
        AddGroupMember groupObj, entityType, entityID, membersAdded, missing
        lastID = entityID
        entityID = entityID + signedStep
    Loop
End Sub

Private Function GroupRangeTypeForEntity(ByVal entityType As Long) As Long
    GroupRangeTypeForEntity = -1

    If entityType = C_FT_CSYS Then
        GroupRangeTypeForEntity = C_FGR_CSYS_ID
    ElseIf entityType = C_FT_NODE Then
        GroupRangeTypeForEntity = C_FGR_NODE_ID
    ElseIf entityType = C_FT_ELEM Then
        GroupRangeTypeForEntity = C_FGR_ELEM_ID
    ElseIf entityType = C_FT_MATL Then
        GroupRangeTypeForEntity = C_FGR_MATL_ID
    ElseIf entityType = C_FT_PROP Then
        GroupRangeTypeForEntity = C_FGR_PROP_ID
    End If
End Function

Private Sub AddGroupMember(ByRef groupObj As Object, ByVal entityType As Long, _
                           ByVal entityID As Long, ByRef membersAdded As Long, _
                           ByRef missing As Long)
    Dim rc As Long
    Dim rangeType As Long

    rangeType = GroupRangeTypeForEntity(entityType)
    If rangeType < 0 Then
        missing = missing + 1
        Exit Sub
    End If

    rc = groupObj.RangeAdd(rangeType, entityID, 0, 1, 1)
    If rc = C_FE_OK Then
        membersAdded = membersAdded + 1
    Else
        missing = missing + 1
    End If
End Sub


Private Function IsCommentLine(ByVal lineText As String) As Boolean
    Dim s As String

    s = Trim$(lineText)
    IsCommentLine = (Len(s) > 0 And Left$(s, 1) = "$")
End Function
