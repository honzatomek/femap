Rem FILE: GetBarOutputData.bas
Sub Main

'Connect to FEMAP
Dim App As femap.model
Set App = feFemap()

'Get the current view
Dim v As femap.View
Set v = App.feView
Dim viewID As Long
rc = App.feAppGetActiveView(viewID)

rc = v.Get(viewID)

'Create and Output Set and Output Vector
Dim outset  As femap.OutputSet
Set outset = App.feOutputSet
Dim outdata As Object
Dim cCount As Long

Dim OutputSetID As Long
Dim elemID As Long
Dim fElem As femap.Elem
Dim eCount As Long

Dim elSet As femap.Set
Set elSet = App.feSet

'=============================
StartVec = 3113   ' Worksheets(1).Cells(1, 2)
vecCount = 2       '  Worksheets(1).Cells(2, 2)
eType = 2            ' Worksheets(1).Cells(3, 2)
'=============================

'Create an Element Object
Set fElem = App.feElem()

' Find out if there is an active output set and if so use it
If (v.OutputSet > 0) Then

    'Set OutputSetID to the current active Output Set of the active Window
    OutputSetID = v.OutputSet
    rc = outset.Get(OutputSetID)

    cCount = -1
 Dim ev As Double

    For J = StartVec To StartVec + vecCount-1
        eCount = 0
        cCount = cCount + 1

        'Load the Output Vector in question
    	Set outdata = outset.vector(J)

    	'Loop through all the elements in the model
        'This line resets the loop to the first element
        rc = fElem.First()

		Dim Msg As String
		Msg = "Elem Id    '"+ Str$(J) +" " + outdata.title
		App.feAppMessage( FCL_BLACK, Msg )

        While (fElem.ID < 100000)
            'If the Element is a Bar (2) then put it's data in the spreadsheet
				 ev = outdata.value(fElem.ID)

            If ( fElem.type = eType Or eType = 0) Then
                eCount = eCount + 1
				 ev = outdata.value(fElem.ID)
				Msg = Str$(fElem.ID) + "           " + Format$( ev, "#0.00000")
				App.feAppMessage( FCL_BLACK, Msg )
               ' Worksheets(1).Cells(eCount + 5, 1).value = fElem.ID
               ' Worksheets(1).Cells(eCount + 5, 2 + cCount).value = outdata.value(fElem.ID)
            End If
            rc = fElem.Next()
        Wend
    Next J
End If

End Sub

