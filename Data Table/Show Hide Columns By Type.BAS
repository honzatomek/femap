Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim data As femap.DataTable
	Set data = App.feDataTable

	Dim colnum As Long
	Dim coltype As Variant
	Dim Title As Variant

	data.GetColumnInfo (colnum, coltype, Title)

	For j = 0 To (colnum -1)

	If coltype (j) <> 1 Then

	data.SetColumnVisible (j, True)

	'Mess = Title(j) + ", " + "Column Type" + Str(coltype (j))

	'App.feAppMessage (FCM_HIGHLIGHT, mess)

	End If

	If coltype (j) = 12 Then

	data.SetColumnVisible (j, False)

	End If

	If coltype (j) = 13 Then

	data.SetColumnVisible (j, False)

	End If

	If coltype (j) = 14 Then

	data.SetColumnVisible (j, False)

	End If

	If coltype (j) = 8 Then

	data.SetColumnVisible (j, False)

	End If

	Next j
	
End Sub
