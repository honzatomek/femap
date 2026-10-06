' =============================================================================
' Description: Sums whole-model forces and moments for every load case and reports the results.
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

    Dim ldsets As Object
    Dim LdSet As Object
    Dim activeLD As Long
    Dim eSET As Object

    Dim printResults As Boolean
    Dim printDetails As Boolean
    Dim meshLength As Double
    Dim meshArea As Double
    Dim meshVolume As Double
    Dim structMass As Double
    Dim nonstructMass As Double
    Dim totalMass As Double
    Dim structCG(2) As Double
    Dim nonstructCG(2) As Double
    Dim totalCG(2) As Double
    Dim inertia(5) As Double
    Dim inertiaCG(5) As Double

    Dim forceOverview As String
    Dim massOverview As String
    Dim forceLine As String
    Dim massLine As String

    er = 0
    expandGEOM = True
    doLIST = True
    useSETS = False
    nodeSET = 0
    elemSET = 0
    basePOINT(0) = 0#
    basePOINT(1) = 0#
    basePOINT(2) = 0#
    csysID = 0

    Set ldsets = App.feSet()
    Set LdSet = App.feLoadSet()
    Set eSET = App.feSet()

    rc = ldsets.AddAll(C_FT_LOAD_DIR)
    If ldsets.Count() = 0 Then GoTo Cleanup

    rc = eSET.AddAll(C_FT_ELEM)
    printResults = True
    printDetails = False
    activeLD = App.Info_ActiveID(C_FT_LOAD_DIR)

    forceOverview = "Forces Overview (daN):" & Chr(13) & Chr(10) & "Load Case | Fx | Fy | Fz"
    massOverview = "Mass Overview (kg):" & Chr(13) & Chr(10) & "Load Case | Mass"

    rc = ldsets.Reset()
    Do While ldsets.Next() <> 0
        rc = LdSet.Get(ldsets.CurrentID())
        If rc = C_FE_OK Then
            App.Info_ActiveID(C_FT_LOAD_DIR) = LdSet.ID()
            Call App.feAppMessage(C_FCM_NORMAL, "[" & LdSet.ID() & "].." & LdSet.title())
            Call App.feAppMessage(C_FCM_WARNING, "Check Sum Of Forces")

            rc = App.feCheckSumForces(expandGEOM, doLIST, useSETS, nodeSET, elemSET, basePOINT, csysID, summedFORCES)

            forceLine = "[" & LdSet.ID() & "].." & LdSet.title() & " | " & _
                        CStr(Round(summedFORCES(9) / 10#, 1)) & " | " & _
                        CStr(Round(summedFORCES(10) / 10#, 1)) & " | " & _
                        CStr(Round(summedFORCES(11) / 10#, 1))
            forceOverview = forceOverview & Chr(13) & Chr(10) & forceLine
            Call App.feAppMessage(C_FCM_NORMAL, "Total: Fx = " & CStr(Round(summedFORCES(9) / 10#, 1)) & _
                                  " daN, Fy = " & CStr(Round(summedFORCES(10) / 10#, 1)) & _
                                  " daN, Fz = " & CStr(Round(summedFORCES(11) / 10#, 1)) & " daN")

            If LdSet.BodyAccelOn() Then
                Call App.feAppMessage(C_FCM_WARNING, "Mass Properties")
                rc = App.feMeasureMeshMassProp(eSET.ID(), csysID, printResults, printDetails, _
                                                meshLength, meshArea, meshVolume, structMass, _
                                                nonstructMass, totalMass, structCG, nonstructCG, _
                                                totalCG, inertia, inertiaCG)
                massLine = "[" & LdSet.ID() & "].." & LdSet.title() & " | " & CStr(Round(totalMass * 1000#, 1))
                massOverview = massOverview & Chr(13) & Chr(10) & massLine
                Call App.feAppMessage(C_FCM_NORMAL, "Total: " & CStr(Round(totalMass * 1000#, 1)) & " kg")
            End If
            Call App.feAppMessage(C_FCM_NORMAL, "")
        End If
    Loop

    Call App.feAppMessage(C_FCM_NORMAL, "")
    Call App.feAppMessage(C_FCM_NORMAL, forceOverview)
    If Len(massOverview) > Len("Mass Overview (kg):" & Chr(13) & Chr(10) & "Load Case | Mass") Then
        Call App.feAppMessage(C_FCM_NORMAL, "")
        Call App.feAppMessage(C_FCM_NORMAL, massOverview)
    End If

    App.Info_ActiveID(C_FT_LOAD_DIR) = activeLD

Cleanup:
    On Error Resume Next
    App.Info_ActiveID(C_FT_LOAD_DIR) = activeLD
    If DBUG Then Call App.feAppMessage(C_FCM_NORMAL, "The script exited with code: " & er)
    Set eSET = Nothing
    Set LdSet = Nothing
    Set ldsets = Nothing
    Set App = Nothing
End Sub
