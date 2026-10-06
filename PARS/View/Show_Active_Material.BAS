' =============================================================================
' Description: Shows only entities associated with the active material in the current view.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

'Shows active material
Const C_FT_ELEM As Long = 8
Const C_FT_MATL As Long = 10
Const C_FGD_ELEM_BYMATL As Long = 22

Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long

    On Error GoTo ExitHere

    Dim maID As Long
    Dim elset As Object
    Set elset = App.feSet()

    maID = App.Info_ActiveID(C_FT_MATL)
    If maID <= 0 Then GoTo ExitHere

    rc = elset.AddRule(maID, C_FGD_ELEM_BYMATL)
    If elset.Count() > 0 Then Call App.feViewShow(C_FT_ELEM, elset.ID)

ExitHere:
    On Error Resume Next
    Set elset = Nothing
    Set App = Nothing
End Sub
