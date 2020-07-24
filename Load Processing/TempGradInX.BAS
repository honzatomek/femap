Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim xLow As Double
	Dim xHigh As Double

	Dim feNode As femap.Node
	Set feNode = App.feNode

	xLow = 99999999.00
	xHigh = -99999999.00

	While feNode.Next
		If feNode.x < xLow Then
			xLow = feNode.x
		End If

		If feNode.x > xHigh Then
			xHigh = feNode.x
		End If
	Wend

	rc = feNode.Reset

	Dim ldSetID As Long
	ldSetID = App.Info_ActiveID( FT_LOAD_DIR )

    If ldSetID = 0 Then
    	rc = MsgBox( "No Active Load Set, Exiting.", vbOkOnly, "Temp Gradient Program")
        GoTo JUMP_OUT
    End If

	Dim ndTemp As femap.LoadNTemp
	Set ndTemp = App.feLoadNTemp
	Dim baseTemp As Double
	Dim gradTemp As Double

	ndTemp.setID = ldSetID

	rc = App.feGetReal( "Enter Base Temperature", -10000, 10000, baseTemp )
	rc = App.feGetReal( "Enter Temperature Gradient", -10000,10000,gradTemp )
	While feNode.Next
		ndTemp.temp = baseTemp + gradTemp * (Abs( feNode.x - xLow ) / Abs(xHigh - xLow ) )
		rc = ndTemp.Put( feNode.ID )
	Wend

	rc = App.feViewRegenerate( 0 )

JUMP_OUT:

End Sub
