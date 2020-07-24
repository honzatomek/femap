'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime#Scripting
Dim m As Integer
Dim msg() As Variant
Dim cmbfiles$()

Sub Main
    Dim App As femap.model
    Set App = feFemap()
	Dim rc As Long

    Dim fName As String

    rc = App.feFileGetName("Select a File - will look at all in the same directory", "Select Directory", "*.cmb", True, fName)
    If rc = 0 Then Exit Sub

	fldr = Left(fName, InStrRev(fName, "\"))

	ChDir fldr
	fName = Dir$("*.cmb")
	Dim a As Long
	a = 0

	While fName <> ""
		ReDim Preserve cmbfiles(a)
		cmbfiles(a) = fName
		fName = Dir$()
		a = a + 1
	Wend

	Begin Dialog UserDialog 100,100,550,378,"Import .cmb Results:",.DlgFunc ' %GRID:10,7,1,1
        MultiListBox 10,49,520,301,cmbfiles(),.List
        OKButton 170,350,80,21
        CancelButton 280,350,90,21
        CheckBox 10,21,90,14,"Select All",.CheckBox1
    End Dialog
    Dim dlg As UserDialog
    Dialog dlg ' show dialog (wait for ok)

    Select Case dlg.CheckBox1
    Case 1

    Case Else
    	For i = 0 To UBound(msg)
    		Debug.Print msg(i)
    	Next
    End Select

    
End Sub

'See DialogFunc help topic for more information.
Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
        'Beep
    Case 2 ' Value changing or button pressed
        Select Case DlgItem$
            Case "OK"
                For m = 0 To UBound(DlgValue("List"))
                    ReDim Preserve msg(m)
                    msg(m) = cmbfiles( DlgValue("List")(m) )
                Next m
            Case "Cancel","Esc"
                End
        End Select
    Case 3 ' TextBox or ComboBox text changed
        rc = 1
    Case 4 ' Focus changed
        Debug.Print "DlgFocus="""; DlgFocus(); """"
        DlgFunc = True
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function
