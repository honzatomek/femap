' =============================================================================
' Description: Template demonstrating late-bound attachment to an existing Excel instance or creation of a new one.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit On

Const DBUG As Boolean = True
Const C_FCM_NORMAL As Long = 0
Private er As Long

Sub Main
    Dim App As Object
    Dim appExcel As Object
    Dim wbkReport As Object
    Dim wksReport As Object

    Set App = feGetObject()
    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    ' Attach to an already-running Excel instance when possible.
    On Error Resume Next
    Set appExcel = GetObject("", "Excel.Application")
    If appExcel Is Nothing Then
        Err.Clear
        Set appExcel = CreateObject("Excel.Application")
    End If
    On Error GoTo 0

    If appExcel Is Nothing Then
        MsgBox "Microsoft Excel could not be opened."
        er = 1
        GoTo Cleanup
    End If

    Set wbkReport = appExcel.Workbooks.Add
    Set wksReport = wbkReport.Worksheets(1)
    wksReport.Cells(1, 1) = "This is an export template to automate Excel from FEMAP/N4W."
    appExcel.Visible = True

Cleanup:
    On Error Resume Next
    If Not App Is Nothing Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    Set wksReport = Nothing
    Set wbkReport = Nothing
    Set appExcel = Nothing
    Set App = Nothing
End Sub
