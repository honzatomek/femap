Rem File: RenumberNodesAndElements.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim elSet As femap.Set
	Set elSet = App.feSet

	Dim ndSet As femap.Set
	Set ndSet = App.feSet

	Dim CSys As femap.CSys
	Set CSys = App.feCSys
	Dim CSysID As Long
	CSysID = App.feCSys.Active
	Dim ncsys As Long
	Dim csysTitles As Variant
	CSys.GetTitleList (0, 0, ncsys, csysIDs ,csysTitles )
	For nc = 0 To ncsys-1
		csysTitles(nc) = Str$(csysIDs(nc)) + ".." + csysTitles(nc)
	Next nc

	Dim csysList$()
	csysList$() = csysTitles

	Dim reOrderOpt As Long

	Begin Dialog UserDialog 320,280,"Renumber Elements and Nodes" ' %GRID:10,7,1,1
		GroupBox 10,7,290,168,"Select Renumbering Order",.GroupBox1
		OptionGroup .Group1
			OptionButton 40,28,110,14,"XYZ or RTZ",.OptionButton1
			OptionButton 40,49,110,14,"XZY or RZT",.OptionButton2
			OptionButton 40,126,110,14,"YXZ or TRZ",.OptionButton3
			OptionButton 40,77,110,14,"ZXY or ZRT",.OptionButton4
			OptionButton 40,98,110,14,"ZYX or ZTR",.OptionButton5
			OptionButton 40,147,110,14,"YZX or TZR",.OptionButton6
		text 20,182,280,14,"Select Renumbering Coordinate System",.Text1
		DropListBox 20,203,280,21,csysList(),.csysList
		OKButton 20,245,80,21
		CancelButton 220,245,80,21
	End Dialog
	Dim dlg As UserDialog

	Select Case Dialog (dlg)
		Case 0		'User Pressed Cancel
			GoTo EndMacro
		Case -1		'User Pressed Ok
			Select Case dlg.Group1
				Case 0
					reOrderOpt = 1
				Case 1
					reOrderOpt = 2
				Case 2
					reOrderOpt = 3
				Case 3
					reOrderOpt = 4
				Case 4
					reOrderOpt = 5
				Case 5
					reOrderOpt = 6
			End Select
	End Select

	CSysID = csysIDs ( dlg.csysList )

	'Renumber
	rc = ndSet.AddAll( FT_NODE )
	rc = elSet.AddAll( FT_ELEM )

	rc = App.feAppLock

	Select Case reOrderOpt
			Case 1 'XYZ 456
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 4, CSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 5, CSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 6, CSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 4, CSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 5, CSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 6, CSysID, False, False, False )
			Case 2 ' XZY 465
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 5, cSysID, False, False, False )
			Case 3 ' YXZ 546
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 6, cSysID, False, False, False )
			Case 4 'ZXY 645
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 5, cSysID, False, False, False )
			Case 5 'ZYX 654
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 4, cSysID, False, False, False )
			Case 6 'YZX 564
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_NODE, ndSet.ID, 1, 4, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 5, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 6, cSysID, False, False, False )
				rc = App.feRenumberOpt( FT_ELEM, elSet.ID, 1, 4, cSysID, False, False, False )
		End Select

		EndMacro:
		rc = App.feAppUnlock
		rc = App.feViewRegenerate( 0 )

End Sub
