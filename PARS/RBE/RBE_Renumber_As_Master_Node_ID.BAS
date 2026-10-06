' =============================================================================
' Description: Renumbers rigid elements using their reference/master node IDs.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

' Renumbers rigid elements to the ID of their reference node.
' Uses legacy feRenumberOpt for N4W 2004 compatibility.

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FET_RIGID As Long = 29
Const C_FTO_RIGID As Long = 13

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim el As Object
    Dim eset As Object
    Dim eID As Long
    Dim nID As Long
    Dim rc As Long

    Set el = App.feElem()
    Set eset = App.feSet()

    On Error GoTo ErrorHandler

    rc = eset.Select(C_FT_ELEM, True, "Select rigid elements to renumber by reference-node ID:")
    If rc <> C_FE_OK Then GoTo ExitHere

    rc = eset.Reset()
    eID = eset.Next()
    Do While eID <> 0
        rc = el.Get(eID)
        If (rc = C_FE_OK) And (el.type = C_FET_RIGID) And (el.topology = C_FTO_RIGID) Then
            nID = el.Node(0)
            rc = App.feRenumberOpt(C_FT_ELEM, -eID, nID, 0, 0, False, False, False)
        End If
        eID = eset.Next()
    Loop
    GoTo ExitHere

ErrorHandler:
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description

ExitHere:
    Set eset = Nothing
    Set el = Nothing
    Set App = Nothing
End Sub
