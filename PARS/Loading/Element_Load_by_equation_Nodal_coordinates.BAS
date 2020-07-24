Option Explicit On

Const DBUG As Boolean = False
Private er As Long
Private ldir As Long

'load elements by function with variables x0, y0, z0, x1, y1, z1, A, L

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim elset As Object
    Set elset = App.feSet()

    Dim el As Object
    Set el = App.feElem()

    Dim nd As Object
    Set nd = App.feNode()

    Dim vecn(8) As String
    Dim v(8) As Double

    Dim A0 As Double
    Dim A1 As Double
    Dim L As Double

    vecn(0) = "X0"
    vecn(1) = "Y0"
    vecn(2) = "Z0"
    vecn(3) = "X1"
    vecn(4) = "Y1"
    vecn(5) = "Z1"
    vecn(6) = "A0"
    vecn(7) = "A1"
    vecn(8) = "L"

    Dim load As Double
    Dim ld As Object
    Set ld = App.feLoadMesh()

    Dim DOF(2) As Long
    DOF(0) =6
    DOF(1) = 0
    DOF(2) = 0

	Dim vals(5) As Double
	For i = 0 To 5 Step 1
		vals(i) = 0
	Next
	Dim func(4) As Long
	For i = 0 To 4 Step 1
		func(i) = 0
	Next

    rc = elset.Select(8, True, "Select elements to load:")
	If (rc <> -1) Or (elset.Count() < 1) Then
		er = 1
		GoTo Cleanup
	End If

	Dim eq As String, eq2 As String
	eq = UCase(InputBox("Input the equation:" & vbLf & "variables: [x0,y0,z0] [x1,y1,z1] [A0,A1] [L]", "Load by equation"))
	If eq = "" Then
		er = 2
		GoTo Cleanup
	End If

	DOF(0) = GetDirection()
	If DOF(0) = 0 Then
		er = 3
		GoTo Cleanup
	End If

	rc = elset.Reset()
	Do While elset.Next() <> 0
		rc = el.Get(elset.CurrentID())
		rc = nd.Get(el.Node(0))
		v(0) = nd.x
		v(1) = nd.y
		v(2) = nd.z
		rc = nd.Get(el.Node(1))
		v(3) = nd.x
		v(4) = nd.y
		v(5) = nd.z
		eq2 = eq
		rc = el.Area(A0, A1)
		v(6) = CDbl(A0)
		v(7) = CDbl(A1)
		rc = el.Length(L)
		v(8) = CDbl(L)
		For i = 0 To UBound(v) Step 1
			If InStr(1, eq2, vecn(i)) <> 0 Then
				eq2 = Replace(eq2, vecn(i), CStr(v(i)), 1,-1)
			End If
		Next
		eq2 = Replace(eq2, ",", ".")
		If DBUG Then
			Debug.Print eq & " -> " & eq2 & " = " & Eval(eq2)
		End If
		load = CDbl(Eval(eq2))
		vals(0) = load
		vals(1) = load
		rc = ld.Add(-el.ID, 41, 0, DOF, vals, func)
		'Call App.feAppMessage(FCM_NORMAL, Eval(eq2))
	Loop

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set ld = Nothing
	Set nd = Nothing
	Set el = Nothing
	Set elset = Nothing
	Set App = Nothing
End Sub

Private Function GetDirection() As Long
	Dim rc As Long
	ldir = 0

	Begin Dialog UserDialog 320,91,"Get Load Direction",.loaddir ' %GRID:10,7,1,1
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
	dlg.Group1 = 5
	rc = Dialog(dlg)

	GetDirection = ldir
End Function

Rem See DialogFunc help topic for more information.
Private Function loaddir(DlgItem$, Action%, SuppValue?) As Boolean
	Select Case Action%
	Case 1 ' Dialog box initialization
	Case 2 ' Value changing or button pressed
		Rem loaddir = True ' Prevent button press from closing the dialog box
		loaddir = True
		Select Case DlgItem$
		Case "OK"
			ldir = DlgValue("Group1") + 1
			loaddir = False
		Case "Cancel"
			ldir = 0
			loaddir = False
		End Select
	Case 3 ' TextBox or ComboBox text changed
	Case 4 ' Focus changed
	Case 5 ' Idle
		Rem Wait .1 : loaddir = True ' Continue getting idle actions
	Case 6 ' Function key
	End Select
End Function


'Ludolfs number pi
Private Function pi() As Double
	'pi = 52163/16604
	pi = 3.14159265358979323846264338327950288419716939937510
End Function

'Eulers number
Private Function e() As Double
	e = 2.71828182845904523536028747135266249775724709369995
End Function

'Logarithm of x with the base a
Private Function LogA(ByVal x As Double, Optional ByVal a As Double = 10) As Double
	LogA = Log(x)/Log(a)
End Function

'Logarithm of x with the base 10
Private Function Log10(ByVal x As Double) As Double
	Log10 = Log(x)/Log(10)
End Function

'Logarithm of x with the base e = 2,718
Private Function LogN(ByVal x As Double) As Double
	LogN = Log(x)
End Function

'Tangent x
Private Function tg(ByVal x As Double) As Double
	tg = Sin(x)/Cos(x)
End Function

'Cotangent x
Private Function cotg(ByVal x As Double) As Double
	cotg = Cos(x)/Sin(x)
End Function

'Inverse sine x
Private Function arcsin(ByVal x As Double) As Double
	arcsin = Atn(x/Sqr(-x*x+1))
End Function

'Inverse cosine x
Private Function arccos(ByVal x As Double) As Double
	arccos = Atn(-x/Sqr(-x*x+1))+2*Atn(1)
End Function

'Inverse tangent x
Private Function arctg(ByVal x As Double) As Double
	arctg = Atn(x)
End Function

'Inverse cotangent x
Private Function arccotg(ByVal x As Double) As Double
	arccotg = 2 * Atn(1) - Atn(x)
End Function

'Angle between x axis and point (x,y) in 2D
Private Function ArcTg2(ByVal x As Double, ByVal y As Double) As Double
    Select Case x
        Case Is > 0
            ArcTg2 = Atn(y / x)
        Case Is < 0
            ArcTg2 = Atn(y / x) + pi() * Sgn(y)
            If y = 0 Then ArcTg2 = ArcTg2 + pi()
        Case Is = 0
            ArcTg2 = pi()/2 * Sgn(y)
    End Select
End Function

'Secant x
Private Function sec(ByVal x As Double) As Double
	sec = 1/Cos(x)
End Function

'Cosecant x
Private Function cosec(ByVal x As Double) As Double
	cosec = 1/Sin(x)
End Function

'Inverse secant x
Private Function arcsec(ByVal x As Double) As Double
	arcsec = 2*Atn(1) - Atn(Sgn(x)/Sqr(x*x - 1))
End Function

'Inverse cosecant x
Private Function arccosec(ByVal x As Double) As Double
	arccosec = Atn(Sgn(x)/Sqr(x*x-1))
End Function

'Hyperbolic sine x
Private Function hsin(ByVal x As Double) As Double
	hsin = (Exp(x)-Exp(-x))/2
End Function

'Hyperbolic cosine x
Private Function hcos(ByVal x As Double) As Double
	hcos = (Exp(x) + Exp(-x)) / 2
End Function

'Hyperbolic tangent x
Private Function htg(ByVal x As Double) As Double
	htg = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function

'Hyperbolic cotangent x
Private Function hcotg(ByVal x As Double) As Double
	hcotg = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
End Function

'Hyperbolic secant x
Private Function hsec(ByVal x As Double) As Double
	hsec = 2 / (Exp(x) + Exp(-x))
End Function

'Hyperbolic cosecant x
Private Function hcosec(ByVal x As Double) As Double
	hcosec = 2 / (Exp(x) - Exp(-x))
End Function

'Inverse hyperbolix sine x
Private Function harcsin(ByVal x As Double) As Double
	harcsin = Log(x + Sqr(x * x + 1))
End Function

'Inverse hyperbolix cosine x
Private Function harccos(ByVal x As Double) As Double
	harccos = Log(x + Sqr(x * x - 1))
End Function

'Inverse hyperbolix tangent x
Private Function harctg(ByVal x As Double) As Double
	harctg = Log((1 + x) / (1 - x)) / 2
End Function

'Inverse hyperbolix cotangent x
Private Function harccotg(ByVal x As Double) As Double
	harccotg = Log((x+1)/(x-1))/2
End Function

'Inverse hyperbolix secant x
Private Function harcsec(ByVal x As Double) As Double
	harcsec = Log((Sqr(-x * x + 1) + 1) / x)
End Function

'Inverse hyperbolix cosecant x
Private Function harccosec(ByVal x As Double) As Double
	harccosec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function

'Maximum of values, max 12 variables
Private Function Max(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
	Dim tmp() As Double
	Dim t As Double
	Dim temp As Variant
	Dim z As Long
	temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
	tmp = EvalArgs(temp)

	t = tmp(0)
	For z = 1 To UBound(tmp) Step 1
		If tmp(z) > t Then t = tmp(z)
	Next

	Max = t
End Function

'Minimum of values, max 12 variables
Private Function Min(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
	Dim tmp() As Double
	Dim t As Double
	Dim temp As Variant
	Dim z As Long
	temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
	tmp = EvalArgs(temp)

	t = tmp(0)
	For z = 1 To UBound(tmp) Step 1
		If tmp(z) < t Then t = tmp(z)
	Next

	Min = t
End Function

'Maximum distance from 0 of a set of values, max 12 variables
Private Function MaxAbs(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
	Dim tmp() As Double
	Dim t As Double
	Dim temp As Variant
	Dim z As Long
	temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
	tmp = EvalArgs(temp)

	t = tmp(0)
	For z = 1 To UBound(tmp) Step 1
		If Abs(tmp(z)) > Abs(t) Then t = tmp(z)
	Next

	MaxAbs = t
End Function

'Sum of values
Private Function Sum(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
	Dim tmp() As Double
	Dim t As Double
	Dim temp As Variant
	Dim z As Long
	temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
	tmp = EvalArgs(temp)

	t = tmp(0)
	For z = 1 To UBound(tmp) Step 1
		 t = t + tmp(z)
	Next

	Sum = t
End Function

'Geometric average of values
Private Function Avg(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
	Dim tmp() As Double
	Dim t As Double
	Dim temp As Variant
	Dim z As Long
	temp = Array(a,b,c,d,e,f,g,h,i,j,k,l)
	tmp = EvalArgs(temp)

	t = tmp(0)
	For z = 1 To UBound(tmp) Step 1
		 t = t + tmp(z)
	Next

	Avg = t / (UBound(tmp) + 1)
End Function

'Evaluation function for arguments - if they were passed or not
Private Function EvalArgs(ByRef x() As Variant) As Variant
	Dim tmp() As Double
	Dim i As Long, j As Long, k As Long
	k = -1
	For i = 0 To UBound(x) Step 1
		If Not IsMissing(x(i)) Then
			If IsArray(x(i)) Then
				For j = 0 To UBound(x(i)) Step 1
					k = k + 1
					ReDim Preserve tmp(k)
					tmp(k) = CDbl(x(i)(j))
				Next
			Else
				k = k + 1
				ReDim Preserve tmp(k)
				tmp(k) = CDbl(x(i))
			End If
		End If
	Next
	EvalArgs = tmp
End Function
