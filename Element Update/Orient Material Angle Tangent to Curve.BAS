Sub Main
'femap API       mat_orient_tange.BAS
'author          Peter Pich
'Time            April 2008
'femap version   9.31

'Updated July 2013 by Chip Fricke to add temporary display of material angle and option to reverse angle


'Short Description       Sets 2D element material orientation tangential To a selected Curve

'Detailed Description
'User selects a leading Curve And elements.
'centroid of Each selected element Is projected onto the leading Curve.
'Curve tangent Is computed at the projection Point.
'The tangent vector Is used To Set material orientation of the element.

'Incompatible elements (1D, 3D, plot only) are ignored.

'If the tangent vector Is perpendicular Or almost perpendicular To the element, the material direction Is Not Set.
'These elements are Put In a Group "Normal_direction".
'The "perpendicular or almost perpendicular" Is defined using eps parameter :
 '        dot<-1+eps      Or      dot>1-eps
'where dot Is a dot product of the element normal And the constructed tangent.
'User can Set eps inside the API program. Default eps = 0,01.

'If eps = 0 material orientation Is Set For All compatible elements.
'If eps = 1 material orientation Is Set only For the compatible elements
'which have normal exactly perpendicular To the constructed tangent.
         Dim App As femap.model
    	 Set App = feFemap()

         Dim set_c1 As Set
         Set set_c1 = App.feSet

         Dim curve_1 As Curve
         Set curve_1 = App.feCurve

         Dim set_e1 As Set
         Set set_e1 = App.feSet

         Dim element As Elem
         Set element = App.feElem

         Dim set_fail As Set
         Set set_fail = App.feSet
         Dim g As Group
         Set g = App.feGroup

Rem             variables
         Dim curve_id As Long
         Dim element_id As Long

     Dim dot As Double
     Dim eps As Double
         eps=0.01

         Dim centroid As Variant
         Dim point_on_curve As Variant
         Dim tangent As Variant
         Dim normal As Variant

     Dim i_good As Long
         Dim i_bad As Long
         Dim i_fail As Long

         Dim es As femap.Set	'Set of elements With material direction modified
         Set es = App.feSet
         Dim eid As Long
         Dim matangle As Double

         Dim v As femap.View
         Set v = App.feView
         Dim vid As Long

Rem             _________________________


set_c1.SelectID(FT_CURVE,"select leading curve", curve_id)
rc=curve_1.Get(curve_id)
If rc = FE_FAIL Then
         rc=App.feAppMessage(2, "Exiting")
         Exit Sub
End If

set_e1.Select( FT_ELEM,True,"select elements to set material angle tangent To the leading Curve")
If set_e1.Count() = 0 Then
         rc=App.feAppMessage(2, "Exiting")
     Exit All
End If

element_id = set_e1.First()
While element_id <> FE_FAIL
         element.Get(element_id)

         i = element.type
         If i > 10 And i < 23 Then
             element.GetFaceCentroid(1,centroid)
                 curve_1.Project(centroid,point_on_curve)
                 curve_1.Tangent(False,point_on_curve,tangent)

                 element.GetFaceNormal(1,normal)
         App.feVectorDotProduct(normal,tangent,dot)

             If dot<-1+eps Or dot>1-eps Then
                         i_fail += 1
                         set_fail.Add(element_id)
                 Else
                         rc = App.feModifyMaterialAngle(-element_id,1,tangent,0,0,0)
                         i_good += 1
                         es.Add ( element_id )
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

' Code added to display Element Material Directions and to optionally reverse the angle
	App.feAppGetActiveView ( vid )
	v.Get ( vid )
	Dim emd As Boolean
	emd = v.Draw ( 115 )
	v.Draw ( 115 ) = True
	v.Put ( vid )
	App.feViewRegenerate ( vid )

	rc = App.feAppMessageBox ( 2, "Reverse Element Material Angles?" )
	If rc = -1	Then  'User pressed Yes
		es.Reset
		eid = es.First
		For en = 0 To es.Count-1
			element.Get ( eid )
			matangle = element.MaterialAngle
			element.MaterialAngle = matangle + 180.0
			element.Put ( eid )
			eid = es.Next
		Next  en
	End If

	If emd = False Then
		v.Get ( vid )
		v.Draw ( 115 ) = False
		v.Put ( vid )
	End If

	App.feViewRegenerate ( vid )


Msg = Str$(i_good) + " elements : material orientation set - success"
App.feAppMessage(0, Msg)



If i_fail > 0 Then
         Msg1 = Str$(i_fail) + " elements : material orientation not Set - material direction normal To the element"
         Msg2 = "        these elements stored in a group Normal_direction"
         App.feAppMessage(2, Msg1)
         App.feAppMessage(2, Msg2)
         g.SetAdd(FT_ELEM, set_fail.ID)
         g.title="Normal_direction"
         i = g.NextEmptyID()
         g.Put(i)
End If

End Sub
