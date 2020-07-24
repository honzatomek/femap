'#Reference {00020813-0000-0000-C000-000000000046}#1.8#0#C:\Program Files\Microsoft Office 15\Root\Office15\EXCEL.EXE#Microsoft Excel 15.0 Object Library#Excel
'#Reference {00020813-0000-0000-C000-000000000046}#1.5#0#C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE#Microsoft Excel 11.0 Object Library#Excel

Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

	Dim appExcel As Object
	On Error Resume Next
	Set appExcel = GetObject(, "Excel.Application")
	If Err <> 0 Then
		Set appExcel = CreateObject("Excel.Application")
		Err.Clear
	End If
	If Not DBUG Then
		On Error GoTo Cleanup
	Else
		On Error GoTo 0
	End If

	'Dim wbkReport As Excel.Workbook
	Dim wbkReport As Object
	'Dim wksReport As Excel.Worksheet
	Dim wksReport As Object

	Set wbkReport = appExcel.Workbooks.Add
	Set wksReport = wbkReport.Worksheets(1)

	wksReport.Cells( 1, 1 ) = "This is an export template to automate excel from Femap."


	appExcel.Visible = True

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	If er <> 0 Then appExcel.Quit
	Set wksReport = Nothing
	Set wbkReport = Nothing
	Set appExcel = Nothing
	Set App = Nothing
End Sub
