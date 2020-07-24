Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim rc As Long
    Dim fastdelete As Long
    fastdelete = App.Pref_FastOutputDelete

    Dim activeLoadSetID As Long
    activeLoadSetID = App.feLoadSet.Active

    Dim activeOutSetID As Long
    activeOutSetID = App.feOutputSet.Active
	Dim outSetID As Long

	outSetID = App.Info_MaxID( FT_OUT_CASE ) + 1

	Dim ouVec As femap.Output
	Set ouVec = App.feOutput
	Dim activeOutVecID As Long
	activeOutVecID = App.feOutput.Active

	Dim ouSet As femap.OutputSet
	Set ouSet = App.feOutputSet

	ouSet.title = "Contour Temp Data from Load Set" + Str$(activeLoadSetID)
	ouSet.analysis = FAT_UNKNOWN
	ouSet.Value = 0.0

	rc = ouSet.Put( outSetID )

	Dim ldSetID As Long

	ldSetID= App.Info_ActiveID( FT_LOAD_DIR )

    If ldSetID = 0 Then
    	GoTo Jumping_Out
    End If

	Dim nodalTemp As femap.LoadNTemp
	Set nodalTemp = App.feLoadNTemp

	rc =  ouVec.InitScalarAtNode( outSetID, 9000000, "Nodal Temps", FOT_TEMP , False )

	While nodalTemp.Next
		If nodalTemp.setID = ldSetID Then
			ouVec.Value( nodalTemp.ID ) = nodalTemp.temp
		End If
	Wend

	rc = ouVec.Put( 9000000)

	Dim viewID As Long
	rc = App.feAppGetActiveView( viewID )

	Dim feView As femap.View
	Set feView = App.feView

	rc = feView.Get( viewID )
	feView.Contour = FCF_CONTOUR
	feView.OutputSet = outSetID
	feView.ContourData = 9000000
	rc = feView.Put( viewID )

	rc = feView.Regenerate

	rc = App.feAppMessageBox ( 2, ("Keep Temperature Output?" + Chr(10) + _
		"Pressing No deletes the Output Set") )
	If rc = 0 Then
		App.Pref_FastOutputDelete = 1
		rc = App.feDeleteOutput( outSetID, -9000000 )
		rc = App.feDelete( FT_OUT_CASE, -outSetID )
		feView.Contour = FCF_MODEL_COLOR
		App.feOutputSet.Active = activeOutSetID
		App.feOutput.Active = activeOutVecID
		App.Pref_FastOutputDelete = fastdelete
		rc = feView.Put( viewID )
		rc = feView.Regenerate
	End If

Jumping_Out:

End Sub
