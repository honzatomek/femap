Option Explicit On

Const DBUG As Boolean = True
Private er As Long

Sub Main
    Dim App As femap.model
    Set App = feFemap()
    Dim rc As Long, i As Long

    er = 0

    If Not DBUG Then On Error GoTo Cleanup

    Dim t(2) As Double

	t(0) = 0
	t(1) = 1
	t(2) = 10

    Debug.Print Max(t)

	Debug.Print Max(1, 3, t)
    Debug.Print Eval("Abs(-10)")

Cleanup:
	Call App.feAppMessage(FCM_NORMAL, "The script exited with code: " & er)
	On Error Resume Next
	Set App = Nothing
End Sub

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
