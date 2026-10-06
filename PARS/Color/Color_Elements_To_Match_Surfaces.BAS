' =============================================================================
' Description: Sets element colors to match the surfaces on which the elements are located.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FGD_ELEM_ATSURFACE As Long = 44

Sub Main
    Dim App As Object
    Dim surf As Object
    Dim elObj As Object
    Dim elSet As Object
    Dim rc As Long
    Dim eid As Long

    Set App = feGetObject()
    Set surf = App.feSurface()
    Set elObj = App.feElem()
    Set elSet = App.feSet()

    rc = surf.First()
    Do While rc = C_FE_OK
        Call elSet.Clear()
        Call elSet.AddRule(surf.ID, C_FGD_ELEM_ATSURFACE)
        Call elSet.Reset()
        eid = elSet.Next()
        Do While eid <> 0
            If elObj.Get(eid) = C_FE_OK Then
                elObj.color = surf.color
                Call elObj.Put(eid)
            End If
            eid = elSet.Next()
        Loop
        rc = surf.Next()
    Loop

    rc = App.feViewRegenerate(0)
    Set elSet = Nothing
    Set elObj = Nothing
    Set surf = Nothing
    Set App = Nothing
End Sub
