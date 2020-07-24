'#Reference {00020813-0000-0000-C000-000000000046}#1.5#0#C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE#Microsoft Excel 11.0 Object Library#Excel
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim appExcel As Excel.Application
	Set appExcel =  New Excel.Application

	Dim wbkReport As Excel.Workbook
	Dim wksReport As Excel.Worksheet

	Set wbkReport = appExcel.Workbooks.Add
	Set wksReport = wbkReport.Worksheets(1)

	Dim outSet As femap.OutputSet
    Set outSet = App.feOutputSet

    Dim outSetID As Long

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim colIndex As Long
	Dim rowIndex As Long

	Dim ndID As Long

    Dim feOutput As femap.output
    Set feOutput = App.feOutput
	colIndex = 1

	' Prompt the user to select nodes
    rc = ndSet.Select( FT_NODE, True, "Select node(s) to send data to Excel" )

	' If there were any selected
	If ndSet.count > 0 Then
		'Put the Node Numbers in the A
		rowIndex = 1
		wksReport.Cells( rowIndex, colIndex ) = "Node ID/Time"

		rowIndex = rowIndex + 1

		ndID = ndSet.First()

		While ndID > 0
			wksReport.Cells( rowIndex, colIndex ) = ndID
			rowIndex = rowIndex + 1
			ndID = ndSet.Next()
		Wend

		' Walk each output set
		While outSet.Next()

			colIndex = colIndex + 1
			rowIndex = 1

			'Tell the user we are doing something
			Msg = "Processing Output Set " + Str$( outSet.ID )
			rc = App.feAppMessage( FCM_NORMAL, Msg )
			
			'Use time value of the Output Set as column header
			wksReport.Cells( rowIndex, colIndex ) = outSet.value
			rowIndex = rowIndex + 1

			'Use Output Set ID as column header
			'wksReport.Cells( rowIndex, colIndex ) = outSet.ID
			'rowIndex = rowIndex + 1

			'Use Output Set Title as column header
			'wksReport.Cells( rowIndex, colIndex ) = outSet.title
			'rowIndex = rowIndex + 1

			'Get the data vector for this set
			feOutput.setID = outSet.ID
			rc = feOutput.Get( 31 )

			' now all the nodal data for this set
			ndID = ndSet.First()
			While ndID > 0
				wksReport.Cells( rowIndex, colIndex ) = feOutput.value(ndID)
				rowIndex = rowIndex + 1
				ndID = ndSet.Next()
			Wend
		Wend
	End If

	'Make Excel visible

	appExcel.visible = True

End Sub
