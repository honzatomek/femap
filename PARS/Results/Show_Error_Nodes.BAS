'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\system32\scrrun.dll#Microsoft Scripting Runtime#Scripting
Dim m As Integer
Dim msg As Variant
Dim fls$()
Const  filetype = "*.f06"

Sub Main()
	Debug.Print Join(GetFileText(GetFilename()), vbCrLf)
End Sub

Function GetFileText(ByVal filename As String) As String()
	Dim lFileNumber As Integer
	Dim txt As String
	lFileNumber = FreeFile()

	Open filename For Input As #lFileNumber
	txt = Input(LOF(lFileNumber), lFileNumber)
	Close #lFileNumber
	Debug.Print txt
	GetFileText = Split(txt, vbCrLf)
End Function


Function GetFilename() As String
    Dim App As femap.model
    Set App = feFemap()
	Dim rc As Long

    Dim fName As String

    rc = App.feFileGetName("Select a File - will look at all in the same directory", "Select Directory", filetype, True, fName)
    If rc <> -1 Then
		GetFilename = ""
		Exit Function
    End If

	fldr = Left(fName, InStrRev(fName, "\"))

	ChDir fldr
	fName = Dir$(filetype)
	Dim a As Long
	a = 0

	While fName <> ""
		ReDim Preserve fls(a)
		fls(a) = fName
		fName = Dir$()
		a = a + 1
	Wend

	Begin Dialog UserDialog 100,100,550,378,"Show erroneous nodes from *.f06 file:",.DlgFunc ' %GRID:10,7,1,1
		OKButton 170,350,80,21
		CancelButton 280,350,90,21
		ListBox 10,7,530,336,fls(),.List
	End Dialog
    Dim dlg As UserDialog
    rc = Dialog(dlg) ' show dialog (wait for ok)


    GetFilename = fldr & msg(0)
End Function

'See DialogFunc help topic for more information.
Function DlgFunc(DlgItem$, Action%, SuppValue?) As Boolean
    Select Case Action%
    Case 1 ' Dialog box initialization
        'Beep
    Case 2 ' Value changing or button pressed
        Select Case DlgItem$
            Case "OK"
                    msg = Array(fls( DlgValue("List") ))
            Case "Cancel","Esc"
                End
        End Select
    Case 3 ' TextBox or ComboBox text changed
        rc = 1
    Case 4 ' Focus changed
        'Debug.Print "DlgFocus="""; DlgFocus(); """"
        DlgFunc = True
    Case 5 ' Idle
    Case 6 ' Function key
    End Select
End Function
