Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim data As femap.DataTable
	Set data = App.feDataTable

	Dim w1 As Long

	Dim w2 As Long

 	rc = data.GetColumnWidth (4,w1)

 	w2 = (w1/2)

	rc = data.SetColumnWidth (4, w2)

	
End Sub
