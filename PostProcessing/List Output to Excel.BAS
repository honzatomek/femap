'#Reference {00020813-0000-0000-C000-000000000046}#1.6#0#C:\Program Files\Microsoft Office\Office12\EXCEL.EXE#Microsoft Excel 12.0 Object Library#Excel
'Written by Adrian Jensen
'Predictive Engineering 2008
'Predictive Engineering Assumes No Responsibility For Results Obtained From API
'Written for FEMAP 9.31
'This API writes the translation output vectors for selected nodes to a XLSX file from a selected output set

Sub Main

'Attach To the model In a femap session that Is already running

Dim App As femap.model
Set App = feFemap()

'Dimension an Excel Application, Workbook and Worksheet. Be sure to activate the Microsoft Excel 12.0 Obeject Library in References

Dim appExcel As Excel.Application
Set appExcel =  New Excel.Application

Dim wbkReport As Excel.Workbook
Set wbkReport = appExcel.Workbooks.Add

Dim wksReport As Excel.Worksheet
Set wksReport = wbkReport.Worksheets(1)

Dim Row As Long
Dim Col As Long

'Dimension all other objects and variables

Dim s As Object
Dim v As Object
Dim e As Object
Set e = App.feSet

Dim ov As femap.output
Set ov = App.feOutput

Dim minID As Long
Dim maxID As Long
Dim minVAL As Double
Dim maxVAL As Double

Dim Count As Long
Dim ID As Variant
Dim IDen As Long
Dim Title As Variant

Dim ouSetID As Long
Dim ouVec As femap.output
Set ouVec = App.feOutput
Dim CurrentData As Double

'Ask the API to create the titles in the first row of the worksheet

wksReport.Cells( 1, 1 ) ="Output Set ID"
wksReport.Cells( 1, 2 ) ="Element ID"

i = 3

'Select Output Set and Output Vectors of interest

If App.feSelectOutput( "Select Output Vectors", 0,  FOT_ANY, FOC_ANY, FT_ELEM, False, s, v ) = FE_OK  Then

'Select Elements of interest

	If e.Select( FT_ELEM, True, "Select Elements" ) = FE_OK Then

'Cycle through all selected Output Sets and Output Vectors

		While s.Next
      	v.Reset
		Col=3

			While v.Next
			e.Reset
        	ov.GetFromSet( s.CurrentID, v.CurrentID )

'Get the Titles and IDs of the selected Output Vectors

            ov.GetTitleIDList(False, v.CurrentID, v.CurrentID, Count, ID, Title)
            Row = i

'Print the Output Set IDs and Element IDs and write into first two columns of the Excel file

                While e.Next()
				wksReport.Cells( Row, 1 ).value =s.CurrentID
				wksReport.Cells( Row, 2 ).value =e.CurrentID

'Print the Output Vector data into coresponding cells of the Excel file
'Note: The Get command requires an INT4 input so you must identify the first value of the ID array with (0)
				ouVec.setID = s.CurrentID
				rc = ouVec.Get(v.CurrentID)
               	CurrentData = ouVec.value( e.CurrentID )
       			wksReport.Cells( Row, Col ) =CurrentData

                Row=Row+1
				Wend

'Print the Titles and IDs of the selected Output Vectors into the first two rows of the Excel file

			wksReport.Cells( 1, Col ).value =Title
			wksReport.Cells( 2, Col ).value =ID
			Col=Col+1

        Wend
        i = Row
    Wend
  End If
End If

'Make the Excel spreadsheet visible
appExcel.visible = True

End Sub
