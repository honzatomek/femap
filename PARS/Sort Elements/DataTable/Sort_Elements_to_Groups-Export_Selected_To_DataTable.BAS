' =============================================================================
' Description: Exports selected sorted-element groups to a FEMAP DataTable.
' Compatibility: Modern FEMAP only; the script shows an early compatibility message on N4W 2004.
' =============================================================================

Option Explicit

Const C_FT_ELEM As Long = 8
Const C_FT_GROUP As Long = 24
Const C_FCT_INT As Long = 1
Const C_FCT_DOUBLE As Long = 2
Const C_FCT_STRING As Long = 3

Sub Main()
    Dim App As Object
    Dim dt As Object
    Dim beamGroups As Object
    Dim grp As Object
    Dim txt As Object
    Dim s As Object
    Dim el As Object
    Dim prp As Object
    Dim rc As Long
    Dim i As Long
    Dim j As Long
    Dim gid As Long
    Dim total As Long
    Dim row As Long
    Dim newCol As Long
    Dim l1 As Double
    Dim l2 As Double
    Dim llt As Double
    Dim beamLen As Double

    Set App = feGetObject()

    If App.Info_Version < 10# Then
        MsgBox "This DataTable exporter requires a modern FEMAP version. Use the matching script in the Excel subdirectory in Nastran for Windows 2004.", 48, "Not available in N4W 2004"
        GoTo ExitHere
    End If

    On Error Resume Next
    Err.Clear
    Set dt = App.feDataTable()
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "This DataTable exporter requires a modern FEMAP version. Use the matching script in the Excel subdirectory in Nastran for Windows 2004.", 48, "Not available in N4W 2004"
        GoTo ExitHere
    End If
    On Error GoTo ErrorHandler

    Set beamGroups = App.feSet()
    rc = beamGroups.Select(C_FT_GROUP, True, "Select BEAM groups to export:")
    If rc <> -1 Then GoTo ExitHere
    Call FilterBeamGroups(App, beamGroups)

    If beamGroups.Count() < 1 Then
        MsgBox "No BEAM groups selected.", 48, "Sort Elements"
        GoTo ExitHere
    End If

    Set grp = App.feGroup()
    Set txt = App.feText()
    Set s = App.feSet()
    Set el = App.feElem()
    Set prp = App.feProp()

    total = 0
    rc = beamGroups.Reset()
    For i = 1 To beamGroups.Count()
        gid = beamGroups.Next()
        If grp.Get(gid) = -1 Then total = total + grp.CountList(C_FT_ELEM)
    Next
    If total < 1 Then
        MsgBox "The selected BEAM groups contain no elements.", 48, "Sort Elements"
        GoTo ExitHere
    End If

    Dim elID() As Long
    Dim propID() As Long
    Dim propName() As String
    Dim grpID() As Long
    Dim grpName() As String
    Dim lengths() As Double
    Dim lcry() As Double
    Dim lcrz() As Double
    Dim lcrlt() As Double
    ReDim elID(total - 1)
    ReDim propID(total - 1)
    ReDim propName(total - 1)
    ReDim grpID(total - 1)
    ReDim grpName(total - 1)
    ReDim lengths(total - 1)
    ReDim lcry(total - 1)
    ReDim lcrz(total - 1)
    ReDim lcrlt(total - 1)

    row = 0
    rc = beamGroups.Reset()
    For i = 1 To beamGroups.Count()
        gid = beamGroups.Next()
        If grp.Get(gid) = -1 Then
            l1 = 0#: l2 = 0#: llt = 0#: beamLen = 0#
            If txt.Get(gid) = -1 Then Call ParseBeamHeader(txt.text, l1, l2, llt, beamLen)
            Set s = grp.List(C_FT_ELEM)
            rc = s.Reset()
            For j = 1 To s.Count()
                rc = el.Get(s.Next())
                rc = prp.Get(el.propID)
                elID(row) = el.ID
                propID(row) = prp.ID
                propName(row) = prp.title
                grpID(row) = grp.ID
                grpName(row) = grp.title
                lengths(row) = beamLen
                lcry(row) = l1
                lcrz(row) = l2
                lcrlt(row) = llt
                row = row + 1
            Next
        End If
    Next

    App.feAppManagePanes "Data Table", 1
    dt.Lock False
    dt.Clear
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Group ID", C_FCT_INT, row, elID, grpID, newCol)
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Group Name", C_FCT_STRING, row, elID, grpName, newCol)
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Property ID", C_FCT_INT, row, elID, propID, newCol)
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Property Name", C_FCT_STRING, row, elID, propName, newCol)
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Beam Length", C_FCT_DOUBLE, row, elID, lengths, newCol)
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Critical length 1", C_FCT_DOUBLE, row, elID, lcry, newCol)
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Critical length 2", C_FCT_DOUBLE, row, elID, lcrz, newCol)
    rc = dt.AddColumn(False, False, C_FT_ELEM, 0, "Critical length LT", C_FCT_DOUBLE, row, elID, lcrlt, newCol)
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description, 16, "Sort Elements DataTable export"

ExitHere:
    On Error Resume Next
    Set prp = Nothing
    Set el = Nothing
    Set s = Nothing
    Set txt = Nothing
    Set grp = Nothing
    Set beamGroups = Nothing
    Set dt = Nothing
    Set App = Nothing
End Sub

Private Function BuildBeamGroupSet(ByRef App As Object, ByRef beamGroups As Object) As Long
    Dim allGroups As Object
    Dim grp As Object
    Dim rc As Long
    Dim i As Long
    Dim gid As Long

    Set allGroups = App.feSet()
    Set grp = App.feGroup()
    rc = allGroups.AddAll(C_FT_GROUP)
    rc = allGroups.Reset()
    For i = 1 To allGroups.Count()
        gid = allGroups.Next()
        If grp.Get(gid) = -1 Then
            If Left$(grp.title, 4) = "BEAM" Then rc = beamGroups.Add(gid)
        End If
    Next
    BuildBeamGroupSet = beamGroups.Count()
    Set grp = Nothing
    Set allGroups = Nothing
End Function

Private Sub FilterBeamGroups(ByRef App As Object, ByRef groups As Object)
    Dim grp As Object
    Dim ids() As Long
    Dim n As Long
    Dim i As Long
    Dim rc As Long
    Dim gid As Long

    Set grp = App.feGroup()
    n = groups.Count()
    If n < 1 Then Exit Sub
    ReDim ids(n - 1)
    rc = groups.Reset()
    For i = 0 To n - 1
        ids(i) = groups.Next()
    Next
    For i = 0 To n - 1
        gid = ids(i)
        If grp.Get(gid) <> -1 Then
            rc = groups.Remove(gid)
        ElseIf Left$(grp.title, 4) <> "BEAM" Then
            rc = groups.Remove(gid)
        End If
    Next
    Set grp = Nothing
End Sub

Private Function ParseBeamHeader(ByVal textValue As String, ByRef l1 As Double, ByRef l2 As Double, ByRef llt As Double, ByRef beamLen As Double) As Boolean
    Dim line As String
    Dim p As Long
    Dim pCR As Long
    Dim pLF As Long
    Dim startPos As Long
    Dim fieldText As String
    Dim k As Long
    Dim values(3) As Double

    line = textValue
    pCR = InStr(line, Chr(13))
    pLF = InStr(line, Chr(10))
    If pCR > 0 And pLF > 0 Then
        If pCR < pLF Then p = pCR Else p = pLF
    ElseIf pCR > 0 Then
        p = pCR
    Else
        p = pLF
    End If
    If p > 0 Then line = Left$(line, p - 1)

    startPos = 1
    For k = 0 To 3
        p = InStr(startPos, line, Chr(9))
        If p = 0 Then
            fieldText = Mid$(line, startPos, Len(line) - startPos + 1)
        Else
            fieldText = Mid$(line, startPos, p - startPos)
        End If
        If Not IsNumeric(fieldText) Then
            ParseBeamHeader = False
            Exit Function
        End If
        values(k) = CDbl(fieldText)
        If k < 3 And p = 0 Then
            ParseBeamHeader = False
            Exit Function
        End If
        startPos = p + 1
    Next

    l1 = values(0)
    l2 = values(1)
    llt = values(2)
    beamLen = values(3)
    ParseBeamHeader = True
End Function
