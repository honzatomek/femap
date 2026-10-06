' =============================================================================
' Description: Exports selected nodal constraint/reaction result data to Excel.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FT_ENTITY As Long = 7

Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long
    Dim entities As Object
    Set entities = App.feSet()
    Dim os As Object
    Set os = App.feOutputSet()
    Dim outData As Object
    Set outData = App.feOutput()
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim attached As Boolean
    Dim entityID As Long, setID As Long, row As Long, col As Long, vecID As Long
    Dim value As Double
    Dim hasBase As Boolean
    Dim zero As Double
    Dim fileName As String

    zero = 0.001
    fileName = GetFilePath$("Constraints.xls", "xls", CurDir$, "Save Excel output", 3)

    On Error GoTo Failed
    rc = entities.Select(C_FT_ENTITY, True, "Select nodes to export")
    If rc <> C_FE_OK Then GoTo Cleanup

    Call OpenExcel(xlApp, xlBook, xlSheet, attached)
    xlSheet.Cells(1, 1).Value = "Reakcni sily v uzlu"
    xlSheet.Cells(2, 1).Value = "ID uzlu"
    xlSheet.Cells(2, 3).Value = "T.C.Force"
    xlSheet.Cells(2, 4).Value = "T1"
    xlSheet.Cells(2, 5).Value = "T2"
    xlSheet.Cells(2, 6).Value = "T3"
    xlSheet.Cells(2, 7).Value = "T.C.Moment"
    xlSheet.Cells(2, 8).Value = "R1"
    xlSheet.Cells(2, 9).Value = "R2"
    xlSheet.Cells(2, 10).Value = "R3"
    xlSheet.Cells(2, 12).Value = "Output Set Title"
    row = 4

    rc = entities.Reset()
    entityID = entities.Next()
    Do While entityID <> 0
        setID = os.First()
        Do While setID <> 0
            If os.analysis <> 0 Then
                hasBase = ReadOutputValue(outData, setID, 51, entityID, value)
                If hasBase Then
                    xlSheet.Cells(row, 1).Value = entityID
                    For col = 3 To 10
                        Select Case col
        Case 3
            vecID = 51
        Case 4
            vecID = 52
        Case 5
            vecID = 53
        Case 6
            vecID = 54
        Case 7
            vecID = 55
        Case 8
            vecID = 56
        Case 9
            vecID = 57
        Case 10
            vecID = 58
                        End Select
                        If ReadOutputValue(outData, setID, vecID, entityID, value) Then
                            If Abs(value) < zero Then
                                xlSheet.Cells(row, col).Value = ""
                            Else
                                xlSheet.Cells(row, col).Value = value
                            End If
                        Else
                            xlSheet.Cells(row, col).Value = ""
                        End If
                    Next
                    xlSheet.Cells(row, 12).Value = os.title
                    row = row + 1
                End If
            End If
            setID = os.Next()
        Loop
        entityID = entities.Next()
    Loop

    If row > 4 Then xlSheet.Range("C4:J" & CStr(row - 1)).NumberFormat = "General"
    xlSheet.Columns("A").Font.FontStyle = "Bold"
    xlSheet.UsedRange.AutoFit
    If fileName <> "" Then
        Call SaveWorkbook(xlBook, fileName)
        Print "Output saved to " & fileName
    End If
    GoTo Cleanup

Failed:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description
Cleanup:
    Set outData = Nothing
    Set os = Nothing
    Set entities = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Set App = Nothing
End Sub


Sub OpenExcel(ByRef xlApp As Object, ByRef xlBook As Object, ByRef xlSheet As Object, ByRef attached As Boolean)
    attached = False
    On Error Resume Next
    ' N4W 2004 accepts GetObject when the first argument is an explicit empty string.
    ' Attach to a running Excel instance first; CreateObject is the fallback.
    Set xlApp = GetObject("", "Excel.Application")
    If Err.Number = 0 Then attached = True
    Err.Clear
    On Error GoTo 0

    If Not attached Then
        Set xlApp = CreateObject("Excel.Application")
    End If

    xlApp.Visible = True
    Set xlBook = xlApp.Workbooks.Add()
    Set xlSheet = xlBook.Worksheets(1)
End Sub

Function ReadOutputValue(ByRef outData As Object, ByVal setID As Long, ByVal vecID As Long, ByVal entityID As Long, ByRef value As Double) As Boolean
    Dim rc As Long
    ReadOutputValue = False
    value = 0#
    outData.SetID = setID
    rc = outData.Get(vecID)
    If rc <> -1 Then Exit Function
    On Error Resume Next
    value = outData.Value(entityID)
    If Err.Number = 0 Then ReadOutputValue = True
    Err.Clear
    On Error GoTo 0
End Function

Sub SaveWorkbook(ByRef xlBook As Object, ByVal fileName As String)
    Dim xlApp As Object
    Set xlApp = xlBook.Application
    On Error Resume Next
    xlApp.DisplayAlerts = False
    xlBook.SaveAs fileName, 56
    xlApp.DisplayAlerts = True
    On Error GoTo 0
    Set xlApp = Nothing
End Sub

