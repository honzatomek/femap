' =============================================================================
' Description: Sorts all rigid elements into separate RBE2 and RBE3 FEMAP groups.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Sort all rigid elements into RBE2 and RBE3 groups.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FET_RIGID As Long = 29
Const C_FGD_ELEM_BYTYPE As Long = 24

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim elset As Object
    Dim rbe2set As Object
    Dim rbe3set As Object
    Dim el As Object
    Dim g2 As Object
    Dim g3 As Object
    Dim rc As Long
    Dim elID As Long
    Dim gid As Long

    Set elset = App.feSet()
    Set rbe2set = App.feSet()
    Set rbe3set = App.feSet()
    Set el = App.feElem()
    Set g2 = App.feGroup()
    Set g3 = App.feGroup()

    On Error GoTo ErrorHandler

    rc = elset.AddRule(C_FET_RIGID, C_FGD_ELEM_BYTYPE)
    If (rc <> C_FE_OK) Or (elset.Count() = 0) Then GoTo ExitHere

    rc = elset.Reset()
    elID = elset.Next()
    Do While elID <> 0
        rc = el.Get(elID)
        If rc = C_FE_OK Then
            If el.RigidInterpolate Then
                rc = rbe3set.Add(elID)
            Else
                rc = rbe2set.Add(elID)
            End If
        End If
        elID = elset.Next()
    Loop

    gid = g2.NextEmptyID()
    g2.title = "RBE2"
    rc = g2.Put(gid)
    If rbe2set.Count() > 0 Then rc = g2.SetAdd(C_FT_ELEM, rbe2set.ID)
    rc = g2.Put(g2.ID)

    gid = g3.NextEmptyID()
    g3.title = "RBE3"
    rc = g3.Put(gid)
    If rbe3set.Count() > 0 Then rc = g3.SetAdd(C_FT_ELEM, rbe3set.ID)
    rc = g3.Put(g3.ID)
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set g3 = Nothing
    Set g2 = Nothing
    Set el = Nothing
    Set rbe3set = Nothing
    Set rbe2set = Nothing
    Set elset = Nothing
    Set App = Nothing
End Sub
