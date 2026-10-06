' =============================================================================
' Description: Sums forces and moments for selected nodes for every load case using legacy feCheckSumForces.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit On

Const DBUG As Boolean = False
Const C_FE_OK As Long = -1
Const C_FT_NODE As Long = 7
Const C_FT_ELEM As Long = 8
Const C_FT_LOAD_DIR As Long = 12
Const C_FCM_NORMAL As Long = 0
Const C_FCM_WARNING As Long = 2
Const C_FCM_ERROR As Long = 3

Private er As Long

Sub Main
    Dim App As Object
    Set App = feGetObject()

    Dim rc As Long
    Dim expandGEOM As Boolean
    Dim doLIST As Boolean
    Dim useSETS As Boolean
    Dim nodeSET As Long
    Dim elemSET As Long
    Dim basePOINT(2) As Double
    Dim csysID As Long
    Dim summedFORCES(14) As Double

    er = 0

    Dim ndset As Object
    Set ndset = App.feSet()
    rc = ndset.Select(C_FT_NODE, True, "Select Nodes for Force Summation:")
    If rc <> C_FE_OK Then GoTo Cleanup

    expandGEOM = True
    doLIST = True
    useSETS = True
    nodeSET = ndset.ID()
    elemSET = 0
    basePOINT(0) = 0#
    basePOINT(1) = 0#
    basePOINT(2) = 0#
    csysID = 0

    Dim ldsets As Object
    Dim LdSet As Object
    Dim activeLD As Long
    Set ldsets = App.feSet()
    Set LdSet = App.feLoadSet()

    rc = ldsets.AddAll(C_FT_LOAD_DIR)
    activeLD = App.Info_ActiveID(C_FT_LOAD_DIR)

    rc = ldsets.Reset()
    Do While ldsets.Next() <> 0
        rc = LdSet.Get(ldsets.CurrentID())
        If rc = C_FE_OK Then
            App.Info_ActiveID(C_FT_LOAD_DIR) = LdSet.ID()
            Call App.feAppMessage(C_FCM_ERROR, Chr(10) & "Check Sum Of Forces")
            Call App.feAppMessage(C_FCM_ERROR, "Load Set: [" & LdSet.ID() & "].." & LdSet.title())
            rc = App.feCheckSumForces(expandGEOM, doLIST, useSETS, nodeSET, elemSET, basePOINT, csysID, summedFORCES)
        End If
    Loop

    App.Info_ActiveID(C_FT_LOAD_DIR) = activeLD

Cleanup:
    If DBUG Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    On Error Resume Next
    Set ndset = Nothing
    Set LdSet = Nothing
    Set ldsets = Nothing
    Set App = Nothing
End Sub
