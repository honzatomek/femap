Rem File: SumAdvNonlinContactGPFOnConnector.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ouVec1 As femap.output
	Set ouVec1 = App.feOutput

	Dim ouVec2 As femap.output
	Set ouVec2 = App.feOutput

	Dim ouVec3 As femap.output
	Set ouVec3 = App.feOutput

	rc = ouVec1.Get(231)
	rc = ouVec2.Get(232)
	rc = ouVec3.Get(233)

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

	Dim count As Long
	Dim ltype As Variant
	Dim eID As Variant
	Dim facnID As Variant
	Dim expand As Variant

	Dim conReg As ConnectionRegion
	Set conReg = App.feConnectionRegion

	rc = suSet.Select( FT_CONNECTION, True, "Select Connector")

	If suSet.count > 0 Then
		suSet.First
		conReg.Get(suSet.ID)

		conReg.GetEntities(FT_NODE, True)
		conReg.GetContactList(count, ltype, nID, faceID, expand)

		rc = ouVec1.Get(162)
		rc = ouVec2.Get(163)
		rc = ouVec3.Get(164)

For i = 0 To count-1

			ouVal1 = ouVec1.value( nID(i) )
			ouVal2 = ouVec2.value( nID(i) )
			ouVal3 = ouVec3.value( nID(i) )

			totalx = totalx + ouVal1
			totaly = totaly + ouVal2
			totalz = totalz + ouVal3

Next i

		Msg = "Total X = " + Str$(totalx) + " Total Y = " + Str$(totaly) + " Total Z = " + Str$(totalz)

		rc = App.feAppMessage( FCL_BLACK, Msg)


	End If


	
End Sub
