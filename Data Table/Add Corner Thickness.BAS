Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim data As femap.DataTable
	Set data = App.feDataTable
	data.Lock ( False )

	Dim p As femap.Prop
	Set p = App.feProp

	Dim colnum As Long
	Dim coltype As Variant
	Dim Title As Variant

	Dim PID As Long

	Dim colID1 As Long
	Dim colID2 As Long
	Dim colID3 As Long
	Dim colID4 As Long

	Dim n As Long
	Dim D As Variant

	Dim ID As Variant

	Dim T1 As Variant
	Dim T2 As Variant
	Dim T3 As Variant
	Dim T4 As Variant
	
	Dim msg As String
 
	rows = data.NumberOfRows

	If rows = 0 Then 	
		msg = "No items exist in the Data Table, add items and try again."
		App.feAppMessage (FCM_ERROR, msg)
		App.feAppMessageBox( 0, msg)
	GoTo WRAPUP
	End If

	
	ReDim T1(rows) As Double
	ReDim T2(rows) As Double
	ReDim T3(rows) As Double
	ReDim T4(rows) As Double

	'data.GetColumnInfo (colnum, coltype, Title)

	For j = 0 To (rows -1)

	data.GetColumnValue (data.FindColumn ("Prop ID"),-1, j, n, D)

	data.GetColumnValue (data.FindColumn ("ID"),-1, j, n, ID)

	PID = D(j)

	p.Get (PID)

 	T1(j) = p.pval (0)
	T2(j) = p.pval (1)
	T3(j) = p.pval (2)
	T4(j) = p.pval (3)

	Next j

	data.AddColumn (False, False , FT_ELEM, 0, "Corner thick 1", FCT_DOUBLE, rows, ID, T1, colID1)

	data.AddColumn (False, False , FT_ELEM, 0, "Corner thick 2", FCT_DOUBLE , rows, ID, T2, colID2)

	data.AddColumn (False, False , FT_ELEM, 0, "Corner thick 3", FCT_DOUBLE , rows, ID, T3, colID3)

	If data.FindColumn ("C4") = -999 Then

	GoTo Three

	End If

	data.AddColumn (False, False , FT_ELEM, 0, "Corner thick 4", FCT_DOUBLE , rows, ID, T4, colID4)

	data.SetColumnPosition (colID4, data.FindColumn ("C4") , True)

	Three:

	data.SetColumnPosition (colID1, data.FindColumn ("C1"), True)

	data.SetColumnPosition (colID2, data.FindColumn ("C2"), True)

	data.SetColumnPosition (colID3, data.FindColumn ("C3"), True)
	
WRAPUP:

End Sub
