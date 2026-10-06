' =============================================================================
' Description: Sums multipoint grid-point forces for selected RBE2 elements and exports the reference-node resultants to Excel.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Sums MultiPoint Grid Point Force result vectors over the listed nodes of
' selected RBE2 elements and transfers the resultant force/moment to the
' RBE2 reference node.

Const C_FE_OK As Long = -1
Const C_FT_NODE As Long = 7
Const C_FT_ELEM As Long = 8
Const C_FT_OUT_CASE As Long = 28
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13
Const C_FCM_NORMAL As Long = 0

' Portable Excel alternative to the DataTable exporter.

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim elset As Object
    Dim el As Object
    Dim nd As Object
    Dim osets As Object
    Dim oset As Object
    Dim ovec As Object
    Dim vectors As Object
    Dim rc As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim si As Long
    Dim rowIndex As Long
    Dim nrows As Long
    Dim vecIDs() As Long
    Dim vecTitles(5) As String
    Dim ndcount As Long
    Dim nds As Variant
    Dim face As Variant
    Dim weight As Variant
    Dim dof As Variant
    Dim x0 As Double
    Dim y0 As Double
    Dim z0 As Double
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    Dim val As Double

    Dim colNode() As Long
    Dim colX() As Double
    Dim colY() As Double
    Dim colZ() As Double
    Dim colSetID() As Long
    Dim colSetTitle() As String
    Dim colT1() As Double
    Dim colT2() As Double
    Dim colT3() As Double
    Dim colR1() As Double
    Dim colR2() As Double
    Dim colR3() As Double

    Set elset = App.feSet()
    Set el = App.feElem()
    Set nd = App.feNode()
    Set osets = App.feSet()
    Set oset = App.feOutputSet()
    Set ovec = App.feOutput()
    Set vectors = App.feSet()


    vecTitles(0) = "T1 MultiPoint GPForce"
    vecTitles(1) = "T2 MultiPoint GPForce"
    vecTitles(2) = "T3 MultiPoint GPForce"
    vecTitles(3) = "R1 MultiPoint GPMoment"
    vecTitles(4) = "R2 MultiPoint GPMoment"
    vecTitles(5) = "R3 MultiPoint GPMoment"

    rc = osets.Select(C_FT_OUT_CASE, True, "Select Output Sets to Process:")
    If (rc <> C_FE_OK) Or (osets.Count() = 0) Then GoTo ExitHere

    ReDim vecIDs(osets.Count() * 6 - 1)
    rc = osets.Reset()
    si = 0
    Do While osets.Next() <> 0
        rc = oset.Get(osets.CurrentID())
        If rc = C_FE_OK Then
            For i = 0 To 5
                vecIDs(si * 6 + i) = 0
            Next
            rc = vectors.Clear()
            rc = oset.OutputVectors(vectors.ID)
            rc = vectors.Reset()
            Do While vectors.Next() <> 0
                ovec.SetID = oset.ID
                rc = ovec.Get(vectors.CurrentID())
                If rc = C_FE_OK Then
                    For i = 0 To 5
                        If InStr(1, ovec.title, vecTitles(i)) <> 0 Then vecIDs(si * 6 + i) = ovec.ID
                    Next
                End If
            Loop
        End If
        si = si + 1
    Loop

    rc = elset.Select(C_FT_ELEM, True, "Select RBE2 spiders to sum:")
    If (rc <> C_FE_OK) Or (elset.Count() = 0) Then GoTo ExitHere

    nrows = osets.Count() * elset.Count()
    ReDim colNode(nrows - 1)
    ReDim colX(nrows - 1)
    ReDim colY(nrows - 1)
    ReDim colZ(nrows - 1)
    ReDim colSetID(nrows - 1)
    ReDim colSetTitle(nrows - 1)
    ReDim colT1(nrows - 1)
    ReDim colT2(nrows - 1)
    ReDim colT3(nrows - 1)
    ReDim colR1(nrows - 1)
    ReDim colR2(nrows - 1)
    ReDim colR3(nrows - 1)

    rowIndex = 0
    rc = osets.Reset()
    si = 0
    Do While osets.Next() <> 0
        rc = oset.Get(osets.CurrentID())
        rc = elset.Reset()
        Do While elset.Next() <> 0
            rc = el.Get(elset.CurrentID())
            If rc = C_FE_OK Then
                colNode(rowIndex) = el.Node(0)
                rc = nd.Get(colNode(rowIndex))
                If rc = C_FE_OK Then
                    x0 = nd.x
                    y0 = nd.y
                    z0 = nd.z
                    colX(rowIndex) = x0
                    colY(rowIndex) = y0
                    colZ(rowIndex) = z0
                End If
                colSetID(rowIndex) = oset.ID
                colSetTitle(rowIndex) = oset.title

                If (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) And (Not el.RigidInterpolate) Then
                    rc = el.GetNodeList(0, ndcount, nds, face, weight, dof)
                    If rc = C_FE_OK Then
                        For i = 0 To 5
                            If vecIDs(si * 6 + i) <> 0 Then
                                ovec.SetID = oset.ID
                                rc = ovec.Get(vecIDs(si * 6 + i))
                                If rc = C_FE_OK Then
                                    For k = 0 To ndcount - 1
                                        val = ovec.Value(nds(k))
                                        If i = 0 Then colT1(rowIndex) = colT1(rowIndex) + val
                                        If i = 1 Then colT2(rowIndex) = colT2(rowIndex) + val
                                        If i = 2 Then colT3(rowIndex) = colT3(rowIndex) + val
                                        If i = 3 Then colR1(rowIndex) = colR1(rowIndex) + val
                                        If i = 4 Then colR2(rowIndex) = colR2(rowIndex) + val
                                        If i = 5 Then colR3(rowIndex) = colR3(rowIndex) + val

                                        If i <= 2 Then
                                            rc = nd.Get(nds(k))
                                            If rc = C_FE_OK Then
                                                dx = nd.x - x0
                                                dy = nd.y - y0
                                                dz = nd.z - z0
                                                If i = 0 Then
                                                    colR2(rowIndex) = colR2(rowIndex) + val * dz
                                                    colR3(rowIndex) = colR3(rowIndex) - val * dy
                                                ElseIf i = 1 Then
                                                    colR1(rowIndex) = colR1(rowIndex) - val * dz
                                                    colR3(rowIndex) = colR3(rowIndex) + val * dx
                                                Else
                                                    colR1(rowIndex) = colR1(rowIndex) + val * dy
                                                    colR2(rowIndex) = colR2(rowIndex) - val * dx
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                End If
            End If
            rowIndex = rowIndex + 1
        Loop
        si = si + 1
    Loop

    Dim xl As Object
    Dim wb As Object
    Dim ws As Object
    Dim headers(11) As String
    Dim c As Long

    headers(0) = "Node ID"
    headers(1) = "X"
    headers(2) = "Y"
    headers(3) = "Z"
    headers(4) = "Set Value"
    headers(5) = "Set Title"
    headers(6) = "T1 MultiPoint GPForce"
    headers(7) = "T2 MultiPoint GPForce"
    headers(8) = "T3 MultiPoint GPForce"
    headers(9) = "R1 MultiPoint GPMoment"
    headers(10) = "R2 MultiPoint GPMoment"
    headers(11) = "R3 MultiPoint GPMoment"

    On Error Resume Next
    Set xl = GetObject("", "Excel.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set xl = CreateObject("Excel.Application")
    End If
    If Err.Number <> 0 Then
        MsgBox "Could not start or connect to Microsoft Excel: " & Err.Description
        Err.Clear
        On Error GoTo ErrorHandler
        GoTo ExitHere
    End If
    On Error GoTo ErrorHandler

    xl.Visible = True
    Set wb = xl.Workbooks.Add()
    Set ws = wb.Worksheets(1)
    ws.Name = "RBE2 GPForce"

    For c = 0 To 11
        ws.Cells(1, c + 1).Value = headers(c)
        ws.Cells(1, c + 1).Font.Bold = True
    Next

    For i = 0 To nrows - 1
        ws.Cells(i + 2, 1).Value = colNode(i)
        ws.Cells(i + 2, 2).Value = colX(i)
        ws.Cells(i + 2, 3).Value = colY(i)
        ws.Cells(i + 2, 4).Value = colZ(i)
        ws.Cells(i + 2, 5).Value = colSetID(i)
        ws.Cells(i + 2, 6).Value = colSetTitle(i)
        ws.Cells(i + 2, 7).Value = colT1(i)
        ws.Cells(i + 2, 8).Value = colT2(i)
        ws.Cells(i + 2, 9).Value = colT3(i)
        ws.Cells(i + 2, 10).Value = colR1(i)
        ws.Cells(i + 2, 11).Value = colR2(i)
        ws.Cells(i + 2, 12).Value = colR3(i)
    Next
    ws.Columns.AutoFit

    Call App.feAppMessage(C_FCM_NORMAL, "Processed " & CStr(nrows) & " RBE2/output-set rows.")
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set ws = Nothing
    Set wb = Nothing
    Set xl = Nothing
    Set vectors = Nothing
    Set ovec = Nothing
    Set oset = Nothing
    Set osets = Nothing
    Set nd = Nothing
    Set el = Nothing
    Set elset = Nothing
    Set App = Nothing
End Sub
