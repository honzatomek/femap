'Predictive Engineering 2006
'Predictive Engineering Assumes No Responsibility For Results Obtained From API
'API written by Jared Ellefson
'Jared.Ellefson@PredictiveEngineering.com
'Phone # 541-760-2955
'Written for FEMAP 9.2
'This API will reorient selected entities from one arbitrary plane to another

Sub Main

'First we are going to connect to the active FEMAP session

Dim App As femap.model
Set App = feFemap()

Dim entitySet As femap.Set
Set entitySet = App.feSet

'Now we are going to request from the user what type of entity they would like to reorient

Dim entityType As Long
Dim messageString As String
Dim lists$(9)
lists$(0) = "Points"
lists$(1) = "Curves"
lists$(2) = "Surfaces"
lists$(3) = "Solids"
lists$(4) = "Nodes"
lists$(5) = "Elements"

Begin Dialog UserDialog 280,133,"Entity Information" ' %GRID:10,7,1,1
	text 50,7,160,35,"What type of entity would you like to reorient?"
	DropListBox 90,42,90,63,lists(),.list1
	OKButton 40,84,90,21
	CancelButton 150,84,90,21
End Dialog
Dim dlg As UserDialog

If Dialog(dlg) = 0 Then
GoTo FAIL
End If

If dlg.list1 = 3 Then
	dlg.list1 = 36
End If

entityType = dlg.list1 + 3

If dlg.list1 = 36 Then
	dlg.list1 = 3
End If

'Now we are going to request of the user the ID's of the entites to move

messageString = "Enter the " + lists$(dlg.list1) + " you would like to reorient"

Dim setCount As Long
Dim setID As Long
Dim vecMove(3) As Double

rc = entitySet.Select(entityType,True,messageString)

setCount = entitySet.Count

setID = entitySet.ID

If setCount = 0 Then
	App.feAppMessageBox(0,"No Entities Were Specified")
	GoTo FAIL
End If

'The next bit of code is fairly complicated, we are going to do the rotations to align the two planes

Dim vecLength As Double
Dim vecAngle1 As Double
Dim vecAngle2 As Double
Dim angleNormal As Double

'We will now prompt the user to imput the two planes

rc = App.fePlanePick( "Enter the Initial Plane", pl1Base, pl1Normal, pl1Axis )

If rc = FE_CANCEL Then
	App.feAppMessageBox(0,"No Plane Specified")
	GoTo FAIL
End If

rc = App.fePlanePick( "Enter the Final Plane", pl2Base, pl2Normal, pl2Axis )

If rc = FE_CANCEL Then
	App.feAppMessageBox(0,"No Plane Specified")
	GoTo FAIL
End If

'Now we will move the planes so that their origins are coincident

vecMove(0) = pl2Base(0) - pl1Base(0)
vecMove(1) = pl2Base(1) - pl1Base(1)
vecMove(2) = pl2Base(2) - pl1Base(2)

rc = App.feVectorLength(vecMove,vecLength)

rc = App.feMoveBy( entityType, setID, False, vecLength, vecMove)

'Now we will do rotate the planes to align them,
'the rotations are basically based upon the euler angles which define them relative to each other

rc = App.feVectorCrossProduct(pl1Normal, pl2Normal, crossNormal)

rc = App.feVectorAngle(pl1Normal, pl2Normal, angleNormal)

rc = App.feVectorPlaneIntersect(pl2Base, pl1Normal, pl2Base, pl2Normal, interBase, interDir)

rc = App.feAlignTo( entityType, setID, pl2Base, pl1axis, pl2Base, interDir)

rc = App.feRotateBy( entityType, setID, pl2Base, crossNormal, angleNormal, 0)

rc = App.feAlignTo( entityType, setID, pl2Base, interDir, pl2Base, pl2Axis)

App.feAppMessage(0,"Entities have been reoriented")

FAIL:

	
End Sub
