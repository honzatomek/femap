Sub Main
	Dim App As femap.model
    Set App = feFemap()

    Dim rigidSet As femap.Set
	Set rigidSet = App.feSet

    rc = rigidSet.AddRule( femap.FET_L_RIGID, FGD_ELEM_BYTYPE )
	If rc = 2 Or rc = 4 Then
		GoTo Done
	End If

	Dim rbe As femap.Elem
	Set rbe = App.feElem
	Dim eid As Long'

	Begin Dialog UserDialog 240,161,"Rigid Element Colors" ' %GRID:10,10,1,1
		PushButton 20,15,200,30,"Select RBE2 Color",.rbe2Color_Botton
		PushButton 20,65,200,30,"Select RBE3 Color",.rbe3Color_Button
		OKButton 20,120,80,25
		CancelButton 140,120,80,25
	End Dialog
	Dim dlg As UserDialog

	Dim color2 As Long
	Dim color2Changed As Boolean
	color2Changed = False
	Dim color3 As Long
	Dim color3Changed As Boolean
	color3Changed = False

	Continue:
	Select Case Dialog(dlg)
		Case 0		'Cancel Button Pushed
			GoTo DONE
		Case 1		'Select RBE2 Color Button Pushed
			rc = App.feAppColorPalette ( FCL_BLACK, color2 )
			color2Changed = True
			GoTo Continue
		Case 2		'Select RBE3 Color Button Pushed
			rc = App.feAppColorPalette ( FCL_BLACK, color3 )
			color3Changed = True
			GoTo Continue
	End Select

	'Modify Rigid Element Colors
	rigidSet.Reset
	eid = rigidSet.First
	For i = 0 To rigidSet.Count-1
		rbe.Get ( eid )
		If rbe.RigidInterpolate And color3Changed Then
			rbe.color = color3
		ElseIf rbe.RigidInterpolate = False And color2Changed Then
			rbe.color = color2
		End If
		rbe.Put ( eid )
		eid = rigidSet.Next
	Next i

	App.feAppMessage ( FCM_HIGHLIGHT, "Rigid Element Colors Modified" )

    DONE:
    rc = App.feViewRegenerate ( 0 )

End Sub
