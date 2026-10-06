' =============================================================================
' Description: Exports the first beam-result report layout to Excel.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8

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
    Dim value As Double, zero As Double
    Dim fileName As String
    Dim hasBase As Boolean

    zero = 0.1
    fileName = GetFilePath$("Beams1.xls", "xls", CurDir$, "Save Excel output", 3)
    On Error GoTo Failed
    rc = entities.Select(C_FT_ELEM, True, "Select elements to export")
    If rc <> C_FE_OK Then GoTo Cleanup
    Call OpenExcel(xlApp, xlBook, xlSheet, attached)
    xlSheet.Cells(1, 1).Value = "Vnitrni sily prvku Beam"
    xlSheet.Cells(2, 1).Value = "ID prvku"
    xlSheet.Cells(2, 3).Value = "My"
    xlSheet.Cells(2, 4).Value = "Mz"
    xlSheet.Cells(2, 5).Value = "Mx"
    xlSheet.Cells(2, 6).Value = "Nx"
    xlSheet.Cells(2, 7).Value = "Qy"
    xlSheet.Cells(2, 8).Value = "Qz"
    xlSheet.Cells(2, 9).Value = "Warp.Torq."
    xlSheet.Cells(2, 2).Value = "End"
    xlSheet.Cells(2, 12).Value = "Output Set Title"
    row = 4

    rc = entities.Reset()
    entityID = entities.Next()
    Do While entityID <> 0
        setID = os.First()
        Do While setID <> 0
            If os.analysis = 1 Or os.analysis = 2 Then
                hasBase = ReadOutputValue(outData, setID, 3014, entityID, value)
                If hasBase Then
                    xlSheet.Cells(row, 1).Value = entityID
                    xlSheet.Cells(row, 2).Value = "A"
                    For col = 3 To 9
                        Select Case col
                Case 3: vecID = 3014
                Case 4: vecID = 3015
                Case 5: vecID = 3024
                Case 6: vecID = 3022
                Case 7: vecID = 3018
                Case 8: vecID = 3019
                Case 9: vecID = 3026
                        End Select
                        If ReadOutputValue(outData, setID, vecID, entityID, value) Then
                            If Abs(value) < zero Then xlSheet.Cells(row, col).Value = "" Else xlSheet.Cells(row, col).Value = value
                        End If
                    Next
                    xlSheet.Cells(row, 12).Value = os.title
                    row = row + 1
                End If
            End If
            If os.analysis = 1 Then
                hasBase = ReadOutputValue(outData, setID, 3016, entityID, value)
                If hasBase Then
                    xlSheet.Cells(row, 1).Value = entityID
                    xlSheet.Cells(row, 2).Value = "B"
                    For col = 3 To 9
                        Select Case col
                Case 3: vecID = 3016
                Case 4: vecID = 3017
                Case 5: vecID = 3025
                Case 6: vecID = 3023
                Case 7: vecID = 3020
                Case 8: vecID = 3021
                Case 9: vecID = 3026
                        End Select
                        If ReadOutputValue(outData, setID, vecID, entityID, value) Then
                            If Abs(value) < zero Then xlSheet.Cells(row, col).Value = "" Else xlSheet.Cells(row, col).Value = value
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

