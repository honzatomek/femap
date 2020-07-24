Rem File: ListBeamAxialandTotalShear.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim ouSetID As Long
	Dim ouVec1 As femap.Output
	Set ouVec1 = App.feOutput
	Dim ouVec2 As femap.Output
	Set ouVec2 = App.feOutput
	Dim ouVec3 As femap.Output
	Set ouVec3 = App.feOutput

	Dim feView As femap.View
	Set feView = App.feView
	Dim viewID As Long

	Dim totalShear As Double
	Dim shear1 As Double
	Dim shear2 As Double
	Dim axial As Double

	rc = App.feAppGetActiveView( viewID )

	rc = feView.Get( viewID )

	ouSetID = feView.OutputSet

	ouVec1.setID = ouSetID
	ouVec2.setID = ouSetID
	ouVec3.setID = ouSetID

	rc = ouVec1.Get( 3018 ) 'Beam Shear 1
	rc = ouVec2.Get( 3019 ) 'Beam Shear 2
	rc = ouVec3.Get( 3022 ) 'Beam Axial Force

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	While feElem.Next
		If feElem.type = FET_L_BEAM Then

			shear1 = ouVec1.Value( feElem.ID )
			shear2 = ouVec2.Value( feElem.ID )
			axial = ouVec3.Value( feElem.ID )

			totalShear = Sqr( shear1 * shear1 + shear2 * shear2 )

			Msg = "For Element " + Str$(feElem.ID) + " Axial, Total Shear " + Str$( axial ) + " " + Str$( totalShear )
			rc = App.feAppMessage( FCM_NORMAL, Msg )
		End If
	Wend

	
End Sub
