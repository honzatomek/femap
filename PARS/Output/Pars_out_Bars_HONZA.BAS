' =============================================================================
' Description: Exports bar-element results to Excel using the alternate HONZA report layout.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit
Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, entityID As Long, setID As Long, row As Long, col As Long, vecID As Long
    Dim entities As Object, os As Object, outData As Object
    Set entities = App.feSet()
    Set os = App.feOutputSet()
    Set outData = App.feOutput()
    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim attached As Boolean
    Dim value As Double, zero As Double
    Dim fileName As String
    zero = 0.1
    fileName = GetFilePath$("Bars_HONZA.xls", "xls", CurDir$, "Save Excel output", 3)
    On Error GoTo Failed
    rc = entities.Select(C_FT_ELEM, True, "Select bar elements to export")
    If rc <> C_FE_OK Then GoTo Cleanup
    Call OpenExcel(xlApp, xlBook, xlSheet, attached)
    xlSheet.Cells(1,1).Value = "Vnitrni sily prvku Bar"
    xlSheet.Cells(3,1).Value = "ID prvku"
    xlSheet.Cells(3,2).Value = "Analysis Type"
    xlSheet.Cells(3,3).Value = "Output Set ID"
    xlSheet.Cells(3,4).Value = "Output Set Title"
    xlSheet.Cells(3,5).Value = "3008..Bar EndA Axial Force [N]"
    xlSheet.Cells(3,6).Value = "3004..Bar EndA Pl1 Shear Force [N]"
    xlSheet.Cells(3,7).Value = "3005..Bar EndA Pl2 Shear Force [N]"
    xlSheet.Cells(3,8).Value = "3010..Bar EndA Torque [Nmm]"
    xlSheet.Cells(3,9).Value = "3000..Bar EndA Plane1 Moment [Nmm]"
    xlSheet.Cells(3,10).Value = "3001..Bar EndA Plane2 Moment [Nmm]"
    xlSheet.Cells(3,11).Value = "3002..Bar EndB Plane1 Moment [Nmm]"
    xlSheet.Cells(3,12).Value = "3003..Bar EndB Plane2 Moment [Nmm]"
    xlSheet.Cells(3,13).Value = "3110..Bar EndA Min Comb Stress [MPa]"
    xlSheet.Cells(3,14).Value = "3109..Bar EndA Max Comb Stress [MPa]"
    xlSheet.Cells(3,15).Value = "3112..Bar EndB Min Comb Stress [MPa]"
    xlSheet.Cells(3,16).Value = "3111..Bar EndB Max Comb Stress [MPa]"
    xlSheet.Range("A3:P3").Orientation = 90
    row = 4
    rc = entities.Reset()
    entityID = entities.Next()
    Do While entityID <> 0
        setID = os.First()
        Do While setID <> 0
            If os.analysis = 1 Or os.analysis = 10 Then
                If ReadOutputValue(outData, setID, 3000, entityID, value) Then
                    xlSheet.Cells(row,1).Value = entityID
                    xlSheet.Cells(row,2).Value = os.analysis
                    xlSheet.Cells(row,3).Value = setID
                    xlSheet.Cells(row,4).Value = os.title
                    For col = 5 To 16
                        Select Case col
                Case 5: vecID = 3008
                Case 6: vecID = 3004
                Case 7: vecID = 3005
                Case 8: vecID = 3010
                Case 9: vecID = 3000
                Case 10: vecID = 3001
                Case 11: vecID = 3002
                Case 12: vecID = 3003
                Case 13: vecID = 3110
                Case 14: vecID = 3109
                Case 15: vecID = 3112
                Case 16: vecID = 3111
                        End Select
                        If ReadOutputValue(outData, setID, vecID, entityID, value) Then
                            If Abs(value) < zero Then value = 0#
                            xlSheet.Cells(row,col).Value = value
                        End If
                    Next
                    row = row + 1
                End If
            End If
            setID = os.Next()
        Loop
        entityID = entities.Next()
    Loop
    If row > 4 Then
        xlSheet.Range("E4:L" & CStr(row-1)).NumberFormat = "# ##0"
        xlSheet.Range("M4:P" & CStr(row-1)).NumberFormat = "# ##0.0"
    End If
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

