Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim zLow As Double
	Dim zHigh As Double

	Dim feNode As femap.Node
	Set feNode = App.feNode

	zLow = 99999999.00
	zHigh = -99999999.00

	While feNode.Next
		If feNode.z < zLow Then
			zLow = feNode.z
		End If

		If feNode.z > zHigh Then
			zHigh = feNode.z
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
		ndTemp.temp = baseTemp + gradTemp * (Abs( feNode.z - zLow ) / Abs(zHigh - zLow ) )
		rc = ndTemp.Put( feNode.ID )
	Wend

	rc = App.feViewRegenerate( 0 )

JUMP_OUT:

End Sub
