' =============================================================================
' Description: Creates a linearly varying distributed element load between
'              user-defined values at the minimum and maximum global X of the
'              selected elements.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FT_ELEM As Long = 8
Const C_FCM_NORMAL As Long = 0

Private gLoadDirection As Long

Sub Main
    Dim App As Object
    Dim elset As Object
    Dim elObj As Object
    Dim ndObj As Object
    Dim ld As Object
    Dim rc As Long
    Dim i As Long
    Dim dof(2) As Long
    Dim vals(5) As Double
    Dim func(4) As Long
    Dim xMin As Double
    Dim xMax As Double
    Dim x0 As Double
    Dim x1 As Double
    Dim xMid As Double
    Dim spanX As Double
    Dim leftValue As Double
    Dim rightValue As Double
    Dim loadValue As Double
    Dim firstElement As Boolean
    Dim keepRunning As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Cleanup

    Set App = feGetObject()
    Set elset = App.feSet()
    Set elObj = App.feElem()
    Set ndObj = App.feNode()
    Set ld = App.feLoadMesh()

    dof(1) = 0
    dof(2) = 0
    For i = 0 To 5
        vals(i) = 0#
    Next
    For i = 0 To 4
        func(i) = 0
    Next

    keepRunning = True
    Do While keepRunning
        rc = elset.Select(C_FT_ELEM, True, "Select elements to load:")
        If (rc <> C_FE_OK) Or (elset.Count() < 1) Then
            keepRunning = False
        Else
            firstElement = True
            rc = elset.Reset()
            Do While elset.Next() <> 0
                rc = elObj.Get(elset.CurrentID())
                If rc = C_FE_OK Then
                    rc = ndObj.Get(elObj.Node(0))
                    If rc = C_FE_OK Then x0 = ndObj.x
                    rc = ndObj.Get(elObj.Node(1))
                    If rc = C_FE_OK Then x1 = ndObj.x

                    If firstElement Then
                        xMin = x0
                        xMax = x0
                        firstElement = False
                    End If
                    If x0 < xMin Then xMin = x0
                    If x1 < xMin Then xMin = x1
                    If x0 > xMax Then xMax = x0
                    If x1 > xMax Then xMax = x1
                End If
            Loop

            spanX = xMax - xMin
            If Abs(spanX) <= 0.000000000001 Then
                MsgBox "The selected elements have no measurable span in global X.", 48, "Linear Distributed Load"
                keepRunning = False
            ElseIf Not PromptNumber("Input the leftmost value of Load:", "Linear Distributed load", leftValue) Then
                keepRunning = False
            ElseIf Not PromptNumber("Input the rightmost value of Load:", "Linear Distributed load", rightValue) Then
                keepRunning = False
            Else
                dof(0) = GetDirection()
                If dof(0) = 0 Then
                    keepRunning = False
                Else
                    rc = elset.Reset()
                    Do While elset.Next() <> 0
                        rc = elObj.Get(elset.CurrentID())
                        If rc = C_FE_OK Then
                            rc = ndObj.Get(elObj.Node(0))
                            If rc = C_FE_OK Then x0 = ndObj.x
                            rc = ndObj.Get(elObj.Node(1))
                            If rc = C_FE_OK Then x1 = ndObj.x

                            xMid = (x0 + x1) / 2#
                            loadValue = leftValue + (rightValue - leftValue) * ((xMid - xMin) / spanX)
                            vals(0) = loadValue
                            vals(1) = loadValue
                            rc = ld.Add(-elObj.ID, 41, 0, dof, vals, func)
                        End If
                    Loop
                    Call App.feViewRegenerate(0)
                End If
            End If
        End If
    Loop

Cleanup:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If errNumber <> 0 Then
        If Len(errDescription) > 120 Then errDescription = Left$(errDescription, 120)
        If Not App Is Nothing Then Call App.feAppMessage(3, "PARS load error " & CStr(errNumber) & ": " & errDescription)
    End If
    Set ld = Nothing
    Set ndObj = Nothing
    Set elObj = Nothing
    Set elset = Nothing
    Set App = Nothing
    On Error GoTo 0
End Sub

Private Function PromptNumber(promptText As String, dialogTitle As String, ByRef valueOut As Double) As Boolean
    Dim answer As String

    PromptNumber = False
    Do
        answer = InputBox(promptText, dialogTitle)
        If answer = "" Then Exit Function
        If IsNumeric(answer) Then
            valueOut = CDbl(answer)
            PromptNumber = True
            Exit Function
        End If
        MsgBox "Please enter a numeric value.", 48, dialogTitle
    Loop
End Function

Private Function GetDirection() As Long
    Begin Dialog UserDialog 320,91,"Get Load Direction",.LoadDirectionDialog ' %GRID:10,7,1,1
        GroupBox 10,7,210,84,"Direction",.GroupBox1
        OptionGroup .Group1
            OptionButton 20,28,90,14,"Element X",.Ex
            OptionButton 20,49,90,14,"Element Y",.Ey
            OptionButton 20,70,90,14,"Element Z",.Ez
            OptionButton 130,28,90,14,"Global X",.Gx
            OptionButton 130,49,90,14,"Global Y",.Gy
            OptionButton 130,70,90,14,"Global Z",.Gz
        OKButton 230,14,90,21
        CancelButton 230,42,90,21
    End Dialog
    Dim dlg As UserDialog

    gLoadDirection = 0
    dlg.Group1 = 5
    If Dialog(dlg) = 0 Then gLoadDirection = 0
    GetDirection = gLoadDirection
End Function

Private Function LoadDirectionDialog(DlgItem, Action, SuppValue) As Boolean
    LoadDirectionDialog = True

    Select Case Action
    Case 2
        Select Case DlgItem
        Case "OK"
            gLoadDirection = DlgValue("Group1") + 1
            LoadDirectionDialog = False
        Case "Cancel", "Esc"
            gLoadDirection = 0
            LoadDirectionDialog = False
        End Select
    End Select
End Function
