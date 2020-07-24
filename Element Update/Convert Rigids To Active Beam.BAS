Rem File: ConvertRigidsToActiveBeam.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

Dim bProp As femap.Prop
Set bProp = App.feProp
Dim bpID As Long
bpID = App.Info_ActiveID ( FT_PROP )
bProp.Get ( bpID )
If bProp.type <> 5 And bProp.type <> 37 Then
	App.feAppMessageBox ( 0,	"Activate or Create and Activate a New Beam Property Before Running this API Script" )
	GoTo EndMacro
End If
bProp.Put ( bpID )

Dim fElem As femap.Elem
Dim fSet As femap.Set
Dim nodeSet As femap.Set
Dim nodeCount As Long
Dim elemID As Long
Dim nd1 As femap.Node
Dim nd2 As femap.Node

Dim vec1(3) As Double
Dim vec2(3) As Double

Dim nodeArray As Variant
Dim faceArray As Variant
Dim weightArray As Variant
Dim dofArray As Variant

Dim bElem As femap.Elem
Set bElem = App.feElem
bElem.type = 5 'Beam
bElem.topology = 0
bElem.propID = bpID

Set fElem = App.feElem
Set fSet = App.feSet
Set nodeSet = App.feSet
Set nd1 = App.feNode
Set nd2 = App.feNode
' Select a rigid element
rc = fSet.Select(8, 1, "Select Rigid Elements to Convert")

elemID = fSet.First()

While elemID > 0
    rc = fElem.Get(elemID)
    If fElem.type = 29 Then
        If rc = -1 Then
            rc = fElem.GetNodeList(0, nodeCount, nodeArray, faceArray, weightArray, dofArray)
            If rc = -1 Then
                ' Need to make a new element
                For j = 0 To nodeCount - 1
                    bElem.Node(0) = fElem.Node(0)
                    bElem.Node(1) = nodeArray(j)

                    ' Find a perpindicular vector
                    rc = nd1.Get(fElem.Node(0))
                    rc = nd2.Get(nodeArray(j))

                    vec1(0) = nd2.x - nd1.x
                    vec1(1) = nd2.y - nd1.y
                    vec1(2) = nd2.z - nd1.z

                    Call FindPerpindicularVector(vec1(0), vec1(1), vec1(2), vec2(0), vec2(1), vec2(2))

                    bElem.orient(0) = vec2(0)
                    bElem.orient(1) = vec2(1)
                    bElem.orient(2) = vec2(2)

                    PrintString = "Creating Beam Element " + Str$(bElem.NextEmptyID)
                    rc = App.feAppMessage(4, PrintString)
                    bElem.Put (bElem.NextEmptyID)
                 Next j
                rc = fElem.Delete(fElem.ID)
            End If
        End If
    End If
    elemID = fSet.Next()
Wend

EndMacro:
App.feViewRegenerate ( 0 )
End

End Sub

Sub FindPerpindicularVector(ByRef v10 As Double, ByRef v11 As Double, ByRef v12 As Double, ByRef v20 As Double, ByRef v21 As Double, ByRef v22 As Double)

    Dim length As Double
    Dim v41 As Double
    Dim v42 As Double
    Dim v43 As Double
    Dim v31 As Double
    Dim v32 As Double
    Dim v33 As Double
    length = Sqr(v10 * v10 + v11 * v11 + v12 * v12)

    v10 = v10 / length
    v11 = v11 / length
    v12 = v12 / length

    ' check for alignment with global axes

    If Abs(v10) < 0.00000001 Then
        'yz return x
        v20 = 1#
        v21 = 0#
        v22 = 0#
        GoTo done
    End If

    If Abs(v11) < 0.00000001 Then
        'xz return y
        v20 = 0#
        v21 = 1#
        v22 = 0#
        GoTo done
    End If

    If Abs(v12) < 0.00000001 Then
        'yz return z
        v20 = 0#
        v21 = 0#
        v22 = 1#
        GoTo done
    End If



    If Abs(v10 - 1#) < 0.00000001 Then
        v20 = 0#
        v21 = 1#
        v22 = 0
        GoTo done
    End If

    If Abs(v11 - 1#) < 0.00000001 Then
        v20 = 1#
        v21 = 0#
        v22 = 0
        GoTo done
    End If

    If Abs(v12 - 1#) < 0.00000001 Then
        v20 = 0#
        v21 = 1#
        v22 = 0#
        GoTo done
    End If

    If v10 > 0.9 Then 'Mostly x cross with z
        v41 = v10
        v42 = v11
        v43 = v12
        v31 = 0#
        v32 = 0#
        v33 = 1#
        v20 = v42 * v33 - v32 * v43
        v21 = v31 * v43 - v41 * v33
        v22 = v41 * v32 - v31 * v42
        GoTo done
    Else
        v41 = v10
        v42 = v11
        v43 = v12
        v31 = 1#
        v32 = 0#
        v33 = 0#
        v20 = v42 * v33 - v32 * v43
        v21 = v31 * v43 - v41 * v33
        v22 = v41 * v32 - v31 * v42
        GoTo done
    End If

done:

length = Sqr(v20 * v20 + v21 * v21 + v22 * v22)

v20 = v20 / length
v21 = v21 / length
v22 = v22 / length


End Sub





	
