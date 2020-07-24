Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim readFile As femap.Read
    Set readFile = App.feRead

    Dim fePoint As femap.Point
    Set fePoint = App.fePoint
	Dim s As String
	s = MacroDir + "\NACA2411.txt"

    rc = readFile.Open( s, 80 )
	If rc = FE_OK Then
		While Not readFile.AtEOF
			rc = readFile.Read
			fePoint.x = readFile.RealField( 1, 0.0 )
			fePoint.y = readFile.RealField( 2, 0.0 )
			fePoint.z = 0.0
			rc = fePoint.Put( fePoint.NextEmptyID )
		Wend

		rc = readFile.Open( "NACA2411.txt", 80 )
	Else
		MsgBox( "Unable to open text file",vbOkOnly)
	End If

End Sub
