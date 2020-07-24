Sub Main
'femap API       Update Material Angle using Delta Value.BAS
'author          Peter Pich
'Time            April 2008
'femap version   9.31

'Short Description       Increases 2D element material orientation angle by a constant value

'Detailed Description
'User selects elements And sets the angle incremental value delta.

'For Each selected element where material orientation Is already Set, the value of the orientation angle Is increased by delta.
'If material orientation Is Not already Set, no action Is done.

'Incompatible elements (1D, 3D, plot only) are ignored.

         Dim App As femap.model
    	 Set App = feFemap()

         Dim set_e1 As Set
         Set set_e1 = App.feSet

         Dim element As Elem
         Set element = App.feElem

Rem             variables
         Dim p(3) As Double

         Dim angle As Double
         Dim delta As Double

         Dim element_id As Long
     Dim i_good As Long
         Dim i_not_set As Long
     Dim i_bad As Long

	Dim v As femap.View
	Set v = App.feView
	Dim vid As Long
	App.feAppGetActiveView ( vid )

Rem             program

set_e1.Select( FT_ELEM,True,"select elements to change material angle")
If set_e1.Count() = 0 Then
         rc=App.feAppMessage(2, "Exiting")
     Exit Sub
End If

	v.Get ( vid )
	Dim emd As Boolean
	emd = v.Draw ( 115 )
	v.Draw ( 115 ) = True
	v.Put ( vid )
	App.feViewRegenerate ( vid )

	UpdateAngle:
rc = App.feGetReal("set delta :     angle = angle + delta", -180, 180, delta)
If rc = FE_CANCEL Then
         rc=App.feAppMessage(2, "Exiting")
        GoTo EndMacro
End If

element_id = set_e1.First()
While element_id <> FE_FAIL
element.Get(element_id)
i = element.type
If i > 10 And i < 23 Then
         If element.MaterialAngleFlag Then
                 angle = element.MaterialAngle
         angle += delta
                 rc = App.feModifyMaterialAngle(-element_id,3,p,0,0,angle)
         i_good += 1
         Else
         i_not_set += 1
         End If

Else
         i_bad +=1
End If

element_id = set_e1.Next
Wend


If i_bad > 0 Then
         Msg = Str$(i_bad) + " elements of incompatible type ignored"
         App.feAppMessage(2, Msg)
End If

If i_not_set > 0 Then
         Msg = Str$(i_not_set) + " elements : no material orientation set"
         App.feAppMessage(0, Msg)
End If

Msg = Str$(i_good) + " elements : material orientation changed"
App.feAppMessage(0, Msg)

	EndMacro:
	v.Get ( vid )
	v.Draw ( 115 ) = emd
	v.Put ( vid )
	App.feViewRegenerate ( vid )

End Sub
