Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim grp As Object
    Set grp = App.feGroup()

    Dim oset As Object
    Set oset = App.feOutputSet()
    Dim settitles() As String

    Dim ovec As Object
    Set ovec = App.feOutput()
    Dim vectitles() As String

    Dim osets As Object
    Set osets = App.feSet()

    Dim ovecs As Object
    Set ovecs = App.feSet()

'1
    rc = osets.AddAll(FT_OUT_CASE)
    If (rc <> -1) And (osets.Count() < 1) Then
    	er = 1
		Call App.feAppMessage(FCM_NORMAL, "There are no Output Sets")
    	GoTo Cleanup
    End If

    rc = osets.Reset()
    ReDim settitles(osets.Count() - 1)
	For i = 0 To osets.Count() - 1 Step 1
		rc = oset.Get(osets.Next())
		settitles(i) = oset.ID & ".." & oset.title()
	Next

	rc = oset.Get(osets.First())
	rc = oset.OutputVectors(ovecs.ID)

	rc = ovecs.Reset()
	ReDim vectitles(ovecs.Count() - 1)
	For i = 0 To ovecs.Count() - 1 Step 1
		rc = ovec.Get(ovecs.Next())
		vectitles(i) = ovec.ID() & ".." & ovec.title()
	Next

	Begin Dialog UserDialog 800,420,"Select Based on Output",.outputselect ' %GRID:10,7,1,1
		GroupBox 10,7,340,63,"Output Set",.GroupBox1
		DropListBox 110,42,230,21,settitles(),.oslistbox
		OptionGroup .Group1
			OptionButton 20,28,90,14,"All Sets",.asets
			OptionButton 20,49,90,14,"Set:",.selectSet
		GroupBox 10,77,340,42,"Output Vector",.GroupBox2
		DropListBox 20,91,320,21,vectitles(),.selectVector
		GroupBox 10,126,340,105,"Selection Criteria",.GroupBox3
		OptionGroup .Group2
			OptionButton 20,147,90,14,"Above",.above4
			OptionButton 20,210,90,14,"Outside",.outside
			OptionButton 20,168,90,14,"Below",.below
			OptionButton 20,189,90,14,"Between",.between
		text 140,161,30,14,"Min:",.Text1
		text 140,182,30,14,"Max:",.Text2
		TextBox 180,161,160,21,.min
		TextBox 180,182,160,21,.max
		PushButton 250,203,90,21,"AutoFill",.autofill
		GroupBox 10,238,340,112,"Entity Selection",.GroupBox4
		OptionGroup .Group3
			OptionButton 20,259,90,14,"Nodal",.nodal
			OptionButton 20,280,90,14,"Elemental",.nodal2
		TextBox 120,259,120,84,.TextBox1,1
		PushButton 250,259,90,21,"Add",.Add
		PushButton 250,280,90,21,"Remove",.Remove
		PushButton 250,301,90,21,"Clear",.Clear
		GroupBox 10,357,340,63,"Select If",.GroupBox5
		OptionGroup .Group4
			OptionButton 20,378,310,14,"All Criteria Met (And)",.Ifand
			OptionButton 20,399,310,14,"This Criteria Met (Or)",.Ifor
		GroupBox 360,7,440,413,"Group Rules",.GroupBox6
		text 370,28,40,14,"Title",.Text3
		TextBox 410,25,380,21,.title
		TextBox 370,77,420,266,.SelectionCriteria,1
		text 370,56,300,14,"Specified Selection Criteria",.Text4
		PushButton 370,350,130,28,"Add Criteria",.AddCriteria
		PushButton 510,350,130,28,"Delete Criteria",.DeleteCriteria
		PushButton 650,350,140,28,"Reset Criteria",.ResetCriteria
		PushButton 370,385,130,28,"Update Criteria",.UpdateCriteria
		OKButton 510,385,130,28
		CancelButton 650,385,140,28
	End Dialog
	Dim dlg As UserDialog
	dlg.Group1 = 0
	dlg.Group2 = 0

	rc = Dialog(dlg)



Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub


Rem See DialogFunc help topic for more information.
Private Function outputselect(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
		DlgEnable("oslistbox", False)
	Case 2 ' Value changing or button pressed
		Select Case DlgItem$
		Case "Group1"
			DlgEnable("oslistbox", Not DlgEnable("oslistbox"))
		End Select
		Rem outputselect = True ' Prevent button press from closing the dialog box
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : outputselect = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function
