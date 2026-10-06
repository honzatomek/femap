' =============================================================================
' Description: Sets element colors to match the solids on which the elements are located.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FGD_SURFACE_ONSOLID As Long = 78
Const C_FGD_ELEM_ATSOLID As Long = 45

Sub Main
    Dim App As Object
    Dim solidObj As Object
    Dim surf As Object
    Dim elObj As Object
    Dim surfSet As Object
    Dim elSet As Object
    Dim rc As Long
    Dim surfID As Long
    Dim eid As Long
    Dim solidColor As Long

    Set App = feGetObject()
    Set solidObj = App.feSolid()
    Set surf = App.feSurface()
    Set elObj = App.feElem()
    Set surfSet = App.feSet()
    Set elSet = App.feSet()

    rc = solidObj.First()
    Do While rc = C_FE_OK
        Call surfSet.Clear()
        Call surfSet.AddRule(solidObj.ID, C_FGD_SURFACE_ONSOLID)
        Call surfSet.Reset()
        surfID = surfSet.Next()

        If surfID <> 0 Then
            If surf.Get(surfID) = C_FE_OK Then
                solidColor = surf.color
                Call elSet.Clear()
                Call elSet.AddRule(solidObj.ID, C_FGD_ELEM_ATSOLID)
                Call elSet.Reset()
                eid = elSet.Next()
                Do While eid <> 0
                    If elObj.Get(eid) = C_FE_OK Then
                        elObj.color = solidColor
                        Call elObj.Put(eid)
                    End If
                    eid = elSet.Next()
                Loop
            End If
        End If
        rc = solidObj.Next()
    Loop

    rc = App.feViewRegenerate(0)
    Set elSet = Nothing
    Set surfSet = Nothing
    Set elObj = Nothing
    Set surf = Nothing
    Set solidObj = Nothing
    Set App = Nothing
End Sub
