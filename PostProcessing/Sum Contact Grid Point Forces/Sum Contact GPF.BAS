Rem File: SumContactGPF.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ouVec1 As femap.output
	Set ouVec1 = App.feOutput

	Dim ouVec2 As femap.output
	Set ouVec2 = App.feOutput

	Dim ouVec3 As femap.output
	Set ouVec3 = App.feOutput

	rc = ouVec1.Get(162)
	rc = ouVec2.Get(163)
	rc = ouVec3.Get(164)

	Dim suSet As femap.Set
	Set suSet = App.feSet

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim ouVal1 As Double
	Dim ouVal2 As Double
	Dim ouVal3 As Double

	Dim totalx As Double
	Dim totaly As Double
	Dim totalz As Double

	totalx = 0.0
	totaly = 0.0
	totalz = 0.0


	Dim ndID As Long

	rc = suSet.Select( FT_SURFACE, True, "Select Surfaces")

	If suSet.count > 0 Then
		rc = ndSet.AddSetRule( suSet.ID, FGD_NODE_ATSURFACE )
		rc = ndSet.Reset

		ndID = ndSet.Next()

		While ndID > 0
			ouVal1 = ouVec1.value( ndID )
			ouVal2 = ouVec2.value( ndID )
			ouVal3 = ouVec3.value( ndID )

			totalx = totalx + ouVal1
			totaly = totaly + ouVal2
			totalz = totalz + ouVal3

			ndID = ndSet.Next()
		Wend

		Msg = "Total X = " + Str$(totalx) + " Total Y = " + Str$(totaly) + " Total Z = " + Str$(totalz)

		rc = App.feAppMessage( FCL_BLACK, Msg)


	End If


	
End Sub
