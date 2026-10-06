' =============================================================================
' Description: Lets the user choose separate FEMAP palette colors for RBE2 and
'              RBE3 rigid elements. Modern FEMAP uses feAppColorPalette;
'              Nastran for Windows 2004 falls back to a numeric color-ID prompt.
' Compatibility: Nastran for Windows 2004 and modern FEMAP.
' =============================================================================

Option Explicit

Const C_FE_OK As Long = -1
Const C_FE_CANCEL As Long = 2
Const C_FET_RIGID As Long = 29
Const C_FGD_ELEM_BYTYPE As Long = 24
Const C_FCM_NORMAL As Long = 0
Const C_RBE2_COLOR_DEFAULT As Long = 72
Const C_RBE3_COLOR_DEFAULT As Long = 122
Const C_COLOR_MIN As Long = 0
Const C_COLOR_MAX As Long = 124
Const C_COLOR_API_UNAVAILABLE As Long = -999
Const C_COLOR_PICK_CANCEL As Long = -998
Const C_DIALOG_ACTION_NONE As Long = 0
Const C_DIALOG_ACTION_RBE2 As Long = 1
Const C_DIALOG_ACTION_RBE3 As Long = 2

Private gColor2 As Long
Private gColor3 As Long
Private gColor2Changed As Boolean
Private gColor3Changed As Boolean
Private gDialogAction As Long

Sub Main
    Dim App As Object
    Dim rigidSet As Object
    Dim rbe As Object
    Dim rc As Long
    Dim eid As Long
    Dim pickedColor As Long
    Dim keepDialog As Boolean
    Dim dialogCancelled As Boolean
    Dim accepted As Boolean
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Cleanup

    Set App = feGetObject()

    gColor2 = C_RBE2_COLOR_DEFAULT
    gColor3 = C_RBE3_COLOR_DEFAULT
    gColor2Changed = False
    gColor3Changed = False
    gDialogAction = C_DIALOG_ACTION_NONE

    ' N4W cannot safely open another modal dialog from inside a UserDialog
    ' callback. The callback therefore only records which picker was requested
    ' and closes this dialog. Main opens the picker, then reopens this dialog.
    Begin Dialog UserDialog 240,161,"Rigid Element Colors",.ColorDialog ' %GRID:10,10,1,1
        PushButton 20,15,200,30,"Select RBE2 Color",.rbe2Color_Button
        PushButton 20,65,200,30,"Select RBE3 Color",.rbe3Color_Button
        OKButton 20,120,80,25
        CancelButton 140,120,80,25
    End Dialog
    Dim dlg As UserDialog

    keepDialog = True
    accepted = False

    Do While keepDialog
        gDialogAction = C_DIALOG_ACTION_NONE
        dialogCancelled = False

        If Dialog(dlg) = 0 Then dialogCancelled = True

        Select Case gDialogAction
        Case C_DIALOG_ACTION_RBE2
            pickedColor = PickColor(App, gColor2, _
                                    "Enter the FEMAP palette color for RBE2 elements (0 to 124):", _
                                    "RBE2 Color")
            If pickedColor <> C_COLOR_PICK_CANCEL Then
                gColor2 = pickedColor
                gColor2Changed = True
            End If

        Case C_DIALOG_ACTION_RBE3
            pickedColor = PickColor(App, gColor3, _
                                    "Enter the FEMAP palette color for RBE3 elements (0 to 124):", _
                                    "RBE3 Color")
            If pickedColor <> C_COLOR_PICK_CANCEL Then
                gColor3 = pickedColor
                gColor3Changed = True
            End If

        Case Else
            If dialogCancelled Then
                keepDialog = False
            Else
                accepted = True
                keepDialog = False
            End If
        End Select
    Loop

    If accepted Then
        If gColor2Changed Or gColor3Changed Then
            Set rigidSet = App.feSet()
            Set rbe = App.feElem()

            rc = rigidSet.AddRule(C_FET_RIGID, C_FGD_ELEM_BYTYPE)
            If rc = C_FE_OK Then
                rc = rigidSet.Reset()
                eid = rigidSet.Next()

                Do While eid <> 0
                    rc = rbe.Get(eid)
                    If rc = C_FE_OK Then
                        If rbe.RigidInterpolate Then
                            If gColor3Changed Then
                                rbe.color = gColor3
                                rc = rbe.Put(eid)
                            End If
                        Else
                            If gColor2Changed Then
                                rbe.color = gColor2
                                rc = rbe.Put(eid)
                            End If
                        End If
                    End If
                    eid = rigidSet.Next()
                Loop

                Call App.feAppMessage(C_FCM_NORMAL, "Rigid element colors modified.")
            End If
        End If
    End If

Cleanup:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If Not App Is Nothing Then rc = App.feViewRegenerate(0)

    If errNumber <> 0 Then
        If Len(errDescription) > 100 Then errDescription = Left$(errDescription, 100)
        If Not App Is Nothing Then
            Call App.feAppMessage(3, "PARS Color error " & CStr(errNumber) & ": " & errDescription)
        End If
    End If

    Set rbe = Nothing
    Set rigidSet = Nothing
    Set App = Nothing
    On Error GoTo 0
End Sub

Private Function ColorDialog(DlgItem, Action, SuppValue) As Boolean
    ColorDialog = True

    Select Case Action
    Case 2
        Select Case DlgItem
        Case "rbe2Color_Button"
            gDialogAction = C_DIALOG_ACTION_RBE2
            ColorDialog = False

        Case "rbe3Color_Button"
            gDialogAction = C_DIALOG_ACTION_RBE3
            ColorDialog = False
        End Select
    End Select
End Function

Private Function PickColor(App As Object, defaultColor As Long, promptText As String, dialogTitle As String) As Long
    Dim rc As Long
    Dim selectedColor As Long
    Dim answer As String
    Dim pickedColor As Long
    Dim validAnswer As Boolean

    selectedColor = defaultColor

    If IsModernFemap(App) Then
        rc = C_COLOR_API_UNAVAILABLE
        On Error Resume Next
        Err.Clear
        rc = App.feAppColorPalette(defaultColor, selectedColor)
        If Err.Number <> 0 Then
            Err.Clear
            rc = C_COLOR_API_UNAVAILABLE
        End If
        On Error GoTo 0

        If rc = C_FE_OK Then
            defaultColor = selectedColor
            PickColor = selectedColor
            Exit Function
        ElseIf rc = C_FE_CANCEL Then
            PickColor = C_COLOR_PICK_CANCEL
            Exit Function
        End If
    End If

    validAnswer = False
    Do While Not validAnswer
        answer = InputBox(promptText, dialogTitle, CStr(defaultColor))

        If answer = "" Then
            PickColor = C_COLOR_PICK_CANCEL
            Exit Function
        End If

        If IsNumeric(answer) Then
            pickedColor = CLng(answer)
            If pickedColor >= C_COLOR_MIN And pickedColor <= C_COLOR_MAX Then
                validAnswer = True
            End If
        End If

        If Not validAnswer Then
            MsgBox "Please enter an integer FEMAP color ID from 0 to 124.", 48, dialogTitle
        End If
    Loop

    defaultColor = pickedColor
    PickColor = pickedColor
End Function

Private Function IsModernFemap(App As Object) As Boolean
    Dim versionNumber As Double

    ' Info_Version is used here because this is the legacy version property
    ' available in the N4W/FEMAP object model. N4W 2004 must stay on the
    ' numeric-input path and must never dispatch feAppColorPalette.
    versionNumber = App.Info_Version
    IsModernFemap = (versionNumber >= 10#)
End Function
