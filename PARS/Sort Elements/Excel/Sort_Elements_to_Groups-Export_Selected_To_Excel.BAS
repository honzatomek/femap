' =============================================================================
' Description: Exports selected sorted-element groups to a new Excel workbook.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FT_ELEM As Long = 8
Const C_FT_GROUP As Long = 24

Sub Main()
    Dim App As Object
    Dim beamGroups As Object
    Dim grp As Object
    Dim txt As Object
    Dim s As Object
    Dim el As Object
    Dim prp As Object
    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim rc As Long
    Dim i As Long
    Dim j As Long
    Dim gid As Long
    Dim row As Long
    Dim l1 As Double
    Dim l2 As Double
    Dim llt As Double
    Dim beamLen As Double

    Set App = feGetObject()
    Set beamGroups = App.feSet()
    rc = beamGroups.Select(C_FT_GROUP, True, "Select BEAM groups to export:")
    If rc <> -1 Then GoTo ExitHere
    Call FilterBeamGroups(App, beamGroups)

    If beamGroups.Count() < 1 Then
        MsgBox "No BEAM groups selected.", 48, "Sort Elements"
        GoTo ExitHere
    End If

    On Error Resume Next
    Err.Clear
    Set xl = GetObject("", "Excel.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set xl = CreateObject("Excel.Application")
    End If
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "Microsoft Excel could not be opened.", 16, "Excel export"
        GoTo ExitHere
    End If
    On Error GoTo ErrorHandler

    xl.Visible = True
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)
    On Error Resume Next
    ws.Name = "BEAM Groups"
    On Error GoTo ErrorHandler

    ws.Cells(1, 1).Value = "Element ID"
    ws.Cells(1, 2).Value = "Group ID"
    ws.Cells(1, 3).Value = "Group Name"
    ws.Cells(1, 4).Value = "Property ID"
    ws.Cells(1, 5).Value = "Property Name"
    ws.Cells(1, 6).Value = "Beam Length"
    ws.Cells(1, 7).Value = "Critical length 1"
    ws.Cells(1, 8).Value = "Critical length 2"
    ws.Cells(1, 9).Value = "Critical length LT"

    Set grp = App.feGroup()
    Set txt = App.feText()
    Set s = App.feSet()
    Set el = App.feElem()
    Set prp = App.feProp()

    row = 2
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
                ws.Cells(row, 1).Value = el.ID
                ws.Cells(row, 2).Value = grp.ID
                ws.Cells(row, 3).Value = grp.title
                ws.Cells(row, 4).Value = prp.ID
                ws.Cells(row, 5).Value = prp.title
                ws.Cells(row, 6).Value = beamLen
                ws.Cells(row, 7).Value = l1
                ws.Cells(row, 8).Value = l2
                ws.Cells(row, 9).Value = llt
                row = row + 1
            Next
        End If
    Next

    On Error Resume Next
    ws.Range("A1:I1").Font.Bold = True
    ws.Columns("A:I").AutoFit
    On Error GoTo 0
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description, 16, "Sort Elements Excel export"

ExitHere:
    On Error Resume Next
    Set prp = Nothing
    Set el = Nothing
    Set s = Nothing
    Set txt = Nothing
    Set grp = Nothing
    Set beamGroups = Nothing
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
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
