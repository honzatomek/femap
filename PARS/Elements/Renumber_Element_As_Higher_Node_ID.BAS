' =============================================================================
' Description: Renumbers selected line elements using the higher of their two end-node IDs.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' N4W/FEMAP portable constants (do not rely on type-library enum names)
Const C_FT_ELEM As Long = 8

'Renumbers line elements with the number of their higher end-node ID.
'Uses feRenumberOpt rather than the newer feRenumberOpt2 for old FEMAP compatibility.

Sub Main()
    Dim App As Object
    Set App = feGetObject()
    Dim el As Object
    Set el = App.feElem()
    Dim eset As Object
    Set eset = App.feSet()
    Dim eID As Long
    Dim nID As Long
    Dim i As Long
    Dim ecount As Long
    Dim rc As Long

    rc = eset.Select(8, True, "Select Elements to be renumbered by Node ID:")
    If rc <> -1 Then GoTo Cleanup
    ecount = eset.Count()
    rc = eset.Reset()

    For i = 0 To ecount - 1
        eID = eset.Next()
        rc = el.Get(eID)
        If rc = -1 Then
            '1=Rod, 5=Beam
            If (el.type = 1) Or (el.type = 5) Then
                If el.Node(0) >= el.Node(1) Then
                    nID = el.Node(0)
                Else
                    nID = el.Node(1)
                End If
                rc = App.feRenumberOpt(C_FT_ELEM, -eID, nID, 0, 0, False, False, False)
            End If
        End If
    Next

Cleanup:
    Set el = Nothing
    Set eset = Nothing
    Set App = Nothing
End Sub
