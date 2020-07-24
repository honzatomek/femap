Rem File: LaminateMaxFiberStrain.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feOutVec As femap.output
	Set feOutVec = App.feOutput

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	Dim ret_code As femap.zReturnCode

	Dim feView As femap.View
	Set feView = App.feView

	Dim viewID As Long
	Dim outSetID As Long
	Dim vectorID As Long
	Dim lamCount As Long

	Dim feOutputVec As femap.output
	Set feOutputVec = App.feOutput
	Dim newVecID As Long

	vectorID = 0

	If App.Info_Version <= 9.2 Then
		'v9.2
		vectorID = 7665
		newVecID = 299999
	End If

	If App.Info_Version >= 9.3 Then
		'v9.3
		vectorID = 1000065   '1000065
		newVecID = 8999999
	End If

	If vectorID = 0 Then
		GoTo Jumping_Out
	End If

	rc = App.feAppGetActiveView( viewID )

	rc = feView.Get( viewID )

	If feView.OutputSet > 0 Then

		outSetID = feView.OutputSet
		feOutputVec.setID = outSetID

		lamCount = 0
		While feOutputVec.Get( vectorID + lamCount * 200 )
			lamCount = lamCount + 1
		Wend

		If lamCount > 0 Then
				ReDim fromSetIDs(lamCount) As Long
				ReDim toSetIDs(lamCount) As Long
				ReDim toVecID(lamCount) As Long
				ReDim vecIDs(lamCount) As Long
				ReDim grID(lamCount) As Long
				ReDim factor(lamCount) As Double
				ReDim overwrite(lamCount) As Boolean
				ReDim ply_summary(lamCount) As Boolean
				ReDim set_info(lamCount) As Boolean
				ReDim approach(lamCount) As Long

				feOutputVec.ID =newVecID
				newVecID = feOutputVec.NextEmptyID

				For j=0 To lamCount
					approach(j) = 0
					vecIDs(j) = vectorID + j * 200
					fromSetIDs(j) = outSetID
					toSetIDs(j) = outSetID
					toVecID(j) = newVecID
					grID(j) = 0
					factor(j) = 1.0
					overwrite(j) = True
					ply_summary(j) = False
					set_info(j) = False
				Next J

				rc = App.feOutputProcess( lamCount, False, 5, approach, fromSetIDs, vecIDs, toSetIDs, toVecID, grID, factor, overwrite, ply_summary, set_info )
				rc = feOutputVec.Get( newVecID )
				feOutputVec.title = "Maximum Fiber Strain"
				rc = feOutputVec.Put( newVecID )

				Msg = "The maximum fiber strain in all elements is " + Str$( feOutputVec.maxval ) + ", in element " + Str$(feOutputVec.maxvalID )
				rc = App.feAppMessage( FCM_NORMAL, Msg )


			End If

	End If

Jumping_Out:

End Sub
