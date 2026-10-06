' =============================================================================
' Description: Shows only entities associated with the active property in the current view.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

'Shows active property
Const C_FT_ELEM As Long = 8
Const C_FT_PROP As Long = 11
Const C_FGD_ELEM_BYPROP As Long = 23

Sub Main
    Dim App As Object
    Set App = feGetObject()
    Dim rc As Long

    On Error GoTo ExitHere

    Dim pID As Long
    Dim elset As Object
    Set elset = App.feSet()

    pID = App.Info_ActiveID(C_FT_PROP)
    If pID <= 0 Then GoTo ExitHere

    rc = elset.AddRule(pID, C_FGD_ELEM_BYPROP)
    If elset.Count() > 0 Then Call App.feViewShow(C_FT_ELEM, elset.ID)

ExitHere:
    On Error Resume Next
    Set elset = Nothing
    Set App = Nothing
End Sub
