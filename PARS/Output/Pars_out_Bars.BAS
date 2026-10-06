' =============================================================================
' Description: Exports selected bar-element result data to Excel.
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
    Dim fileName As String, endName As String
    Dim hasBase As Boolean

    zero = 0.1
    fileName = GetFilePath$("Bars.xls", "xls", CurDir$, "Save Excel output", 3)
    On Error GoTo Failed
    rc = entities.Select(C_FT_ELEM, True, "Select bar elements to export")
    If rc <> C_FE_OK Then GoTo Cleanup
    Call OpenExcel(xlApp, xlBook, xlSheet, attached)
    xlSheet.Cells(1,1).Value = "Vnitrni sily prvku Bar"
    xlSheet.Cells(2,1).Value = "ID prvku"
    xlSheet.Cells(2,2).Value = "Analysis Type"
    xlSheet.Cells(2,3).Value = "Output Set ID"
    xlSheet.Cells(2,4).Value = "Output Set Title"
    xlSheet.Cells(2,5).Value = "My"
    xlSheet.Cells(2,6).Value = "Mz"
    xlSheet.Cells(2,7).Value = "Mx"
    xlSheet.Cells(2,8).Value = "Nx"
    xlSheet.Cells(2,9).Value = "Qy"
    xlSheet.Cells(2,10).Value = "Qz"
    xlSheet.Cells(2,11).Value = "Sig.min"
    xlSheet.Cells(2,12).Value = "Sig.max"
    xlSheet.Cells(2,13).Value = "End"
    row = 3

    rc = entities.Reset()
    entityID = entities.Next()
    Do While entityID <> 0
        setID = os.First()
        Do While setID <> 0
            If os.analysis = 1 Or os.analysis = 2 Then
                If ReadOutputValue(outData, setID, 3000, entityID, value) Then
                    Call WriteBarRow(xlSheet, outData, os, entityID, setID, row, "A", zero)
                    row = row + 1
                End If
            End If
            If os.analysis = 1 Then
                If ReadOutputValue(outData, setID, 3002, entityID, value) Then
                    Call WriteBarRow(xlSheet, outData, os, entityID, setID, row, "B", zero)
                    row = row + 1
                End If
            End If
            setID = os.Next()
        Loop
        entityID = entities.Next()
    Loop

    If row > 3 Then xlSheet.Range("E3:L" & CStr(row - 1)).NumberFormat = "# ##0"
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

Sub WriteBarRow(ByRef sh As Object, ByRef outData As Object, ByRef os As Object, ByVal entityID As Long, ByVal setID As Long, ByVal row As Long, ByVal endName As String, ByVal zero As Double)
    Dim col As Long, vecID As Long, value As Double
    sh.Cells(row,1).Value = entityID
    sh.Cells(row,2).Value = os.analysis
    sh.Cells(row,3).Value = setID
    sh.Cells(row,4).Value = os.title
    sh.Cells(row,13).Value = endName
    For col = 5 To 12
        If endName = "A" Then
            Select Case col
                Case 5: vecID = 3000
                Case 6: vecID = 3001
                Case 7: vecID = 3010
                Case 8: vecID = 3008
                Case 9: vecID = 3004
                Case 10: vecID = 3005
                Case 11: vecID = 3110
                Case 12: vecID = 3109
            End Select
        Else
            Select Case col
                Case 5: vecID = 3002
                Case 6: vecID = 3003
                Case 7: vecID = 3010
                Case 8: vecID = 3008
                Case 9: vecID = 3004
                Case 10: vecID = 3005
                Case 11: vecID = 3112
                Case 12: vecID = 3111
            End Select
        End If
        If ReadOutputValue(outData, setID, vecID, entityID, value) Then
            If Abs(value) < zero Then sh.Cells(row,col).Value = 0 Else sh.Cells(row,col).Value = value
        End If
    Next
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

