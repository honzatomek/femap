' =============================================================================
' Description: Exports selected rod-element result data to Excel.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FT_ENTITY As Long = 8

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

    zero = 0.1
    fileName = GetFilePath$("Rods.xls", "xls", CurDir$, "Save Excel output", 3)

    On Error GoTo Failed
    rc = entities.Select(C_FT_ENTITY, True, "Select elements to export")
    If rc <> C_FE_OK Then GoTo Cleanup

    Call OpenExcel(xlApp, xlBook, xlSheet, attached)
    xlSheet.Cells(1, 1).Value = "Vnitrni sily a napeti prvku Rod"
    xlSheet.Cells(2, 1).Value = "ID prvku"
    xlSheet.Cells(2, 5).Value = "Axial Force"
    xlSheet.Cells(2, 6).Value = "Torque Force"
    xlSheet.Cells(2, 7).Value = "Axial Stress"
    xlSheet.Cells(2, 8).Value = "Torsional Stress"
    xlSheet.Cells(2, 9).Value = "Axial Stress M.S."
    xlSheet.Cells(2, 10).Value = "Tors. Stress M.S."
    xlSheet.Cells(2, 12).Value = "Output Set Title"
    row = 4

    rc = entities.Reset()
    entityID = entities.Next()
    Do While entityID <> 0
        setID = os.First()
        Do While setID <> 0
            If os.analysis <> 0 Then
                hasBase = ReadOutputValue(outData, setID, 3036, entityID, value)
                If hasBase Then
                    xlSheet.Cells(row, 1).Value = entityID
                    For col = 5 To 10
                        Select Case col
        Case 5
            vecID = 3036
        Case 6
            vecID = 3038
        Case 7
            vecID = 3183
        Case 8
            vecID = 3186
        Case 9
            vecID = 3185
        Case 10
            vecID = 3187
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

    If row > 4 Then xlSheet.Range("E4:J" & CStr(row - 1)).NumberFormat = "General"
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

