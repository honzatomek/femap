Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim outSetID As Long
	Dim TopX As femap.output
	Set TopX = App.feOutput
	Dim TopY As femap.output
	Set TopY = App.feOutput
	Dim BotX As femap.output
	Set BotX = App.feOutput
	Dim BotY As femap.output
	Set BotY = App.feOutput
	Dim MidX As femap.output
	Set MidX = App.feOutput
	Dim MidY As femap.output
	Set MidY = App.feOutput

	Dim feView As femap.View
	Set feView = App.feView

	Dim viewID As Long

	rc = App.feAppGetActiveView( viewID )
	rc = feView.Get( viewID )
	outSetID = feView.OutputSet

	If outSetID > 0 Then

	MidX.InitScalarAtElem( outSetID, 9000000, "Mid X-Stress", FOT_STRESS, False )
	MidY.InitScalarAtElem( outSetID, 9000001, "Mid Y-Stress", FOT_STRESS, False )

	'7020 Top X Stress
	'7021 Top Y Stress
	'7420 Bot X Stress
	'7421 Bot Y Stress

	TopX.setID = outSetID
	TopY.setID = outSetID
	BotX.setID = outSetID
	BotY.setID = outSetID

	rc = TopX.Get( 7020 )
	rc = TopY.Get( 7021 )
	rc = BotX.Get( 7420 )
	rc = BotY.Get( 7421 )

	Dim feElem As femap.Elem
	Set feElem = App.feElem
	Dim elemID As Long

	While feElem.Next
		If feElem.type = FET_L_PLATE Then
			elemID = feElem.ID
			MidX.value(elemID) = (TopX.value(elemID) + BotX.value(elemID) ) / 2.0
			MidY.value(elemID) = (TopY.value(elemID) + BotY.value(elemID) ) / 2.0
		End If
	Wend

	rc = MidX.Put( 9000000 )
	rc = MidY.Put( 9000001 )

	End If

End Sub
