Rem File: ConvertRBE2stoRBE3s.BAS
Sub Main
	Dim App As femap.model
    Set App = feFemap()

	Dim fElem As femap.Elem
	Dim fSet As femap.Set
	Dim nodeSet As femap.Set
	Dim nodeCount As Long
	Dim elemID As Long
	Dim DOF(6) As Long

	Dim nodeArray As Variant
	Dim faceArray As Variant
	Dim weightArray As Variant
	Dim dofArray As Variant
	Dim releaseArray As Variant

	Set fElem = App.feElem
	Set fSet = App.feSet
	Set nodeSet =App.feSet

	' Select a rigid element
	rc = fSet.Select(8, 1, "Select Rigid Elements to Convert to RBE3")

	If rc = -1 Then

    elemID = fSet.First()

    While elemID <> 0
        'Get the Element
        rc = fElem.Get(elemID)
        If rc = -1 Then
            If fElem.type = 29 Then 'Rigid
                ' Make sure it is not an RBE3 already
                If fElem.RigidInterpolate = 0 Then

                   releaseArray = fElem.vrelease

                    For j = 0 To 6
                        If releaseArray(j) Then
                            DOF(j) = 1
                        Else
                            DOF(j) = 0
                        End If
                    Next j
                    'Get the Arrays
                    rc = fElem.GetNodeList(0, nodeCount, nodeArray, faceArray, weightArray, dofArray)
                    For j = 0 To nodeCount - 1
                        weightArray(j) = 1#
                        dofArray(6 * j) = DOF(0)
                        dofArray((6 * j) + 1) = DOF(1)
                        dofArray((6 * j) + 2) = DOF(2)
                        dofArray((6 * j) + 3) = DOF(3)
                        dofArray((6 * j) + 4) = DOF(4)
                        dofArray((6 * j) + 5) = DOF(5)
                    Next j
                    rc = fElem.PutNodeList(0, nodeCount, nodeArray, faceArray, weightArray, dofArray)
                    fElem.RigidInterpolate = 1
                    rc = fElem.Put(elemID)
                End If
            End If
        End If

        elemID = fSet.Next()
    Wend
End If

End

	
End Sub
