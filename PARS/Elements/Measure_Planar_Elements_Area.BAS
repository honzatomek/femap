' =============================================================================
' Description: Calculates and reports the total area of selected planar elements without using newer GetFaceArea APIs.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' N4W/FEMAP portable constants (do not rely on type-library enum names)
Const C_MSG_NORMAL As Long = 0

'Portable implementation for old Nastran for Windows/FEMAP Basic.
'Does not use Elem.GetFaceArea, which was introduced in later FEMAP versions.

Sub Main()
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long, i As Long

    Dim eset As Object
    Set eset = App.feSet()

    Dim el As Object
    Set el = App.feElem()

    Dim areatotal As Double
    Dim area As Double
    areatotal = 0#

    rc = eset.Select(8, True, "Select planar elements for area inquiry:")
    If rc <> -1 Then GoTo Cleanup
    Call App.feAppMessage(C_MSG_NORMAL, eset.Count() & " elements selected..")

    rc = eset.Reset()
    For i = 0 To eset.Count() - 1 Step 1
        rc = el.Get(eset.Next())
        If rc <> -1 Then GoTo Cleanup

        'Original script accepted FEMAP planar element types 10..22.
        If (el.type > 9) And (el.type < 23) Then
            area = PlanarElementArea(App, el)
            areatotal = areatotal + area
        End If
    Next

    Call App.feAppMessage(C_MSG_NORMAL, "Accumulated Planar Element Area: " & areatotal)

Cleanup:
    On Error Resume Next
    Set el = Nothing
    Set eset = Nothing
    Set App = Nothing
End Sub

Private Function PlanarElementArea(ByRef App As Object, ByRef el As Object) As Double
    Dim n1 As Object, n2 As Object, n3 As Object, n4 As Object
    Dim rc As Long
    Dim a As Double

    Set n1 = App.feNode()
    Set n2 = App.feNode()
    Set n3 = App.feNode()
    Set n4 = App.feNode()

    rc = n1.Get(el.Node(0))
    If rc <> -1 Then GoTo Failed
    rc = n2.Get(el.Node(1))
    If rc <> -1 Then GoTo Failed
    rc = n3.Get(el.Node(2))
    If rc <> -1 Then GoTo Failed

    Select Case el.topology
    Case 2, 3 ' 2=Tri3, 3=Tri6
        a = TriangleAreaNodes(n1, n2, n3)
    Case 4, 5 ' 4=Quad4, 5=Quad8
        rc = n4.Get(el.Node(3))
        If rc <> -1 Then GoTo Failed
        'Split the quadrilateral along diagonal 1-3. For planar elements this is exact.
        a = TriangleAreaNodes(n1, n2, n3)
        a = a + TriangleAreaNodes(n1, n3, n4)
    Case Else
        a = 0#
    End Select

    PlanarElementArea = a
    GoTo Done

Failed:
    PlanarElementArea = 0#

Done:
    Set n4 = Nothing
    Set n3 = Nothing
    Set n2 = Nothing
    Set n1 = Nothing
End Function

Private Function TriangleAreaNodes(ByRef n1 As Object, ByRef n2 As Object, ByRef n3 As Object) As Double
    Dim ax As Double, ay As Double, az As Double
    Dim bx As Double, by As Double, bz As Double
    Dim cx As Double, cy As Double, cz As Double

    ax = n2.x() - n1.x()
    ay = n2.y() - n1.y()
    az = n2.z() - n1.z()
    bx = n3.x() - n1.x()
    by = n3.y() - n1.y()
    bz = n3.z() - n1.z()

    cx = ay * bz - az * by
    cy = az * bx - ax * bz
    cz = ax * by - ay * bx

    TriangleAreaNodes = 0.5# * Sqr(cx * cx + cy * cy + cz * cz)
End Function
