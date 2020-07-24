Dim m As Integer
Dim msg() As Variant
Dim cmbfiles$()

Sub Main ()
	Dim rc As Long
	Dim lastOS As Long
	Dim i As Long, j As Long

	Dim femap As Object
	Set femap = GetObject(, "femap.model")

	Dim Snum As Long
	Dim Title As String
	Dim Pozn As String

	'promenne pro kombinaci
	Dim Radku As Variant
	Dim ov (100)  As Variant
	Dim ply (100) As Variant
	Dim info (100) As Variant
	Dim gr (100) As Variant
	Dim OsProc (100) As Variant
	Dim fact  (100) As Variant
	Dim vyslOS (100) As Variant

	Dim os As Object
	Dim aa As Long

	'nacti jmeno souboru
	Dim fName As String
	Dim fldr As String

	rc = femap.feFileGetName("Vyber souborem slozku s kombinacemi:", "Linear Comb. File","*.cmb", True, fName)
	If rc=0 Then GoTo Konec

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
		GroupBox 120,7,410,35,"Process Combinations",.GroupBox1
		CheckBox 140,21,80,14,"Max",.CheckBoxMax
		CheckBox 250,21,80,14,"Min",.CheckBoxMin
		CheckBox 350,21,80,14,"MaxAbs",.CheckBoxMaxAbs
	End Dialog
	Dim dlg As UserDialog
	Dialog dlg ' show dialog (wait for ok)

	Select Case dlg.CheckBox1
	Case 1
		For i = 0 To UBound(cmbfiles)
			ReDim Preserve msg(i)
			msg(i) = cmbfiles(i)
			Debug.Print msg(i)
		Next
	Case Else
		On Error Resume Next
		If IsError(UBound(msg)) Then
			Exit Sub
		End If
		On Error GoTo 0
	End Select

	Dim process(0 To 2) As Boolean
	process(0) = dlg.CheckBoxMax
	process(1) = dlg.CheckBoxMin
	process(2) = dlg.CheckBoxMaxAbs

	Dim ss As Object
    Set ss = femap.feSet()

    rc = ss.Clear()

	For i = 0 To UBound(msg)
		fName = msg(i)

		Call femap.feAppMessage(FCM_WARNING, "Opening file: " & fName)

		'GoTo Skip
		' vytor objekt file
		Dim f As Object
		Set f = femap.feRead

		rc = f.Open(fName,100) 'otevre soubor
		'rc = f.ShowStatus(True)

		rc = f.SetFreeFormat() 'volny format, oddelovac carka

		Do 'do konce souboru

			rc = f.Read()     '1 radek
			'print f.line
			If chyba(rc) Then GoTo Konec
			Snum=f.IntField( 2, 0)
			'print Snum

			rc = f.Read()     '2 radek
			If chyba(rc) Then GoTo Konec
			Title = f.Line

			rc = f.Read()     '3 radek
			If chyba(rc) Then GoTo Konec
			Pozn = f.Line

			Radku = 0
			Do
				rc = f.Read()
				If chyba(rc) Then GoTo Konec
				OsProc(Radku)=f.IntField( 1, 1)  'kombinuj stav cislo

				fact(Radku)=f.RealField( 2, 0.0) ' s faktorem

				'Print OsProc(Radku), fact(Radku)
				Radku=Radku+1

			Loop Until OsProc(Radku-1)=-999 'do konce kombinace
			rc = f.Read() 'nacte prazdny radek

			Set OsProc(Radku-1) = Empty
			Set fact(Radku-1) = Empty

			Set os = femap.feOutputSet
			rc = os.Delete(Snum) 'vymaze Output set s danym cislem (pokud existuje)

			os.title = Title
			os.notes = Pozn

			os.program = 4 'Nastran
			os.analysis = 1 'linearni statika
			os.value =0.0

			os.Put (Snum) ' ulozi set pod cislem Snum
			rc = ss.Add(Snum)

			'Kombinace zat stavu
			vyslOS(0)=Snum

			rc= femap.feOutputProcess( Radku-1 , True, 2, gr, OsProc, gr , vyslOS ,gr , gr , fact, ov, ply, info )
			If rc<>-1 Then MsgBox ("Chyba pri vytvareni kombinace")
		Loop Until f.AtEOF() 'do konce souboru

		rc = f.Close()

		For j = 0 To UBound(process)
			If process(j) And ss.Count <> 0 Then
				Call femap.feAppMessage(FCM_NORMAL, "Creating Envelope " & j)
				rc = femap.feOutputProcessEnvelopeFromSets( True, j, 2, False, 2, True, ss.ID, 0, 0 )
			End If
		Next
Skip:
		rc = ss.Clear()
	Next

Konec:
	' Print "konec"
End Sub

Function chyba(rc As Long)
   If rc<>-1 Then MsgBox ("Chyba pri nacitani souboru")
   chyba = 0
End Function

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
