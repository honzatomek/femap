Rem File: EnvelopeLaminateFailureIndices.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim feOutVec As femap.Output
	Set feOutVec = App.feOutput

	Dim feElem As femap.Elem
	Set feElem = App.feElem

	Dim ret_code As femap.zReturnCode

	Dim feView As femap.View
	Set feView = App.feView

	Dim viewID As Long
	Dim outSetID As Long

	Dim fromSetIDs(180) As Long
	Dim toSetIDs(180) As Long
	Dim toVecID(180) As Long
	Dim vecIDs(180) As Long
	Dim grID(180) As Long
	Dim factor(180) As Double
	Dim overwrite(180) As Boolean
	Dim ply_summary(180) As Boolean
	Dim set_info(180) As Boolean
	Dim approach(180) As Long

	rc = App.feAppGetActiveView( viewID )

	rc = feView.Get( viewID )

	If feView.OutputSet > 0 Then

		outSetID = feView.OutputSet

		feOutVec.setID = outSetID

		'Determine Laminate Failure Criteria
		'For Hill
		rc = feOutVec.Exist (1000090)
		If rc = -1 Then
			outvecID = 1000090
			GoTo Calc
		End If
		'For Hoffman
		rc = feOutVec.Exist (1000092)
		If rc = -1 Then
			outvecID = 1000092
			GoTo Calc
		End If
		'For Tsai-Wu
		rc = feOutVec.Exist (1000094)
		If rc = -1 Then
			outvecID = 1000094
			GoTo Calc
		End If
		'For Max Strain
		rc = feOutVec.Exist (1000096)
		If rc = -1 Then
			outvecID = 1000096
			GoTo Calc
		Else
			GoTo Done
		End If

		Calc:

		For j=0 To 179
			approach(j) = 0
			vecIDs(j) = outvecID + j * 200
			fromSetIDs(j) = outSetID
			toSetIDs(j) = outSetID
			toVecID(j) = 9000000
			grID(j) = 0
			factor(j) = 1.0
			overwrite(j) = True
			ply_summary(j) = False
			set_info(j) = False
		Next j

		rc = App.feOutputProcess( 180, False, 5, approach, fromSetIDs, vecIDs, toSetIDs, toVecID, grID, factor, overwrite, ply_summary, set_info )

	End If

	Done:

End Sub
