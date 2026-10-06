' =============================================================================
' Description: Shared math helper library for PARS scripts, including array-safe scalar/vector utility functions.
' Compatibility: Modern FEMAP BASIC only; uses Optional arguments not guaranteed by N4W 8.3.
' =============================================================================

Option Explicit On

'Portable FEMAP/N4W math helper library.
'Use from another BAS file with:  '#Uses "Math.bas"

Public Function pi() As Double
    pi = 3.14159265358979323846
End Function

Public Function e() As Double
    e = 2.71828182845904523536
End Function

Public Function LogA(ByVal x As Double, Optional ByVal a As Double = 10) As Double
    LogA = Log(x) / Log(a)
End Function

Public Function Log10(ByVal x As Double) As Double
    Log10 = Log(x) / Log(10)
End Function

Public Function LogN(ByVal x As Double) As Double
    LogN = Log(x)
End Function

Public Function tg(ByVal x As Double) As Double
    tg = Sin(x) / Cos(x)
End Function

Public Function cotg(ByVal x As Double) As Double
    cotg = Cos(x) / Sin(x)
End Function

Public Function arcsin(ByVal x As Double) As Double
    arcsin = Atn(x / Sqr(-x * x + 1))
End Function

Public Function arccos(ByVal x As Double) As Double
    arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Public Function arctg(ByVal x As Double) As Double
    arctg = Atn(x)
End Function

Public Function arccotg(ByVal x As Double) As Double
    arccotg = 2 * Atn(1) - Atn(x)
End Function

Public Function ArcTg2(ByVal x As Double, ByVal y As Double) As Double
    Select Case x
    Case Is > 0
        ArcTg2 = Atn(y / x)
    Case Is < 0
        ArcTg2 = Atn(y / x) + pi() * Sgn(y)
        If y = 0 Then ArcTg2 = ArcTg2 + pi()
    Case Else
        ArcTg2 = pi() / 2 * Sgn(y)
    End Select
End Function

Public Function sec(ByVal x As Double) As Double
    sec = 1 / Cos(x)
End Function

Public Function cosec(ByVal x As Double) As Double
    cosec = 1 / Sin(x)
End Function

Public Function arcsec(ByVal x As Double) As Double
    arcsec = 2 * Atn(1) - Atn(Sgn(x) / Sqr(x * x - 1))
End Function

Public Function arccosec(ByVal x As Double) As Double
    arccosec = Atn(Sgn(x) / Sqr(x * x - 1))
End Function

Public Function hsin(ByVal x As Double) As Double
    hsin = (Exp(x) - Exp(-x)) / 2
End Function

Public Function hcos(ByVal x As Double) As Double
    hcos = (Exp(x) + Exp(-x)) / 2
End Function

Public Function htg(ByVal x As Double) As Double
    htg = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Function

Public Function hcotg(ByVal x As Double) As Double
    hcotg = (Exp(x) + Exp(-x)) / (Exp(x) - Exp(-x))
End Function

Public Function hsec(ByVal x As Double) As Double
    hsec = 2 / (Exp(x) + Exp(-x))
End Function

Public Function hcosec(ByVal x As Double) As Double
    hcosec = 2 / (Exp(x) - Exp(-x))
End Function

Public Function harcsin(ByVal x As Double) As Double
    harcsin = Log(x + Sqr(x * x + 1))
End Function

Public Function harccos(ByVal x As Double) As Double
    harccos = Log(x + Sqr(x * x - 1))
End Function

Public Function harctg(ByVal x As Double) As Double
    harctg = Log((1 + x) / (1 - x)) / 2
End Function

Public Function harccotg(ByVal x As Double) As Double
    harccotg = Log((x + 1) / (x - 1)) / 2
End Function

Public Function harcsec(ByVal x As Double) As Double
    harcsec = Log((Sqr(-x * x + 1) + 1) / x)
End Function

Public Function harccosec(ByVal x As Double) As Double
    harccosec = Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function

Public Function Max(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
    Dim tmp() As Double
    Dim n As Long
    Dim z As Long
    n = 0
    Call AppendEvalArg(a, tmp, n)
    If Not IsMissing(b) Then Call AppendEvalArg(b, tmp, n)
    If Not IsMissing(c) Then Call AppendEvalArg(c, tmp, n)
    If Not IsMissing(d) Then Call AppendEvalArg(d, tmp, n)
    If Not IsMissing(e) Then Call AppendEvalArg(e, tmp, n)
    If Not IsMissing(f) Then Call AppendEvalArg(f, tmp, n)
    If Not IsMissing(g) Then Call AppendEvalArg(g, tmp, n)
    If Not IsMissing(h) Then Call AppendEvalArg(h, tmp, n)
    If Not IsMissing(i) Then Call AppendEvalArg(i, tmp, n)
    If Not IsMissing(j) Then Call AppendEvalArg(j, tmp, n)
    If Not IsMissing(k) Then Call AppendEvalArg(k, tmp, n)
    If Not IsMissing(l) Then Call AppendEvalArg(l, tmp, n)
    Max = tmp(0)
    For z = 1 To n - 1
        If tmp(z) > Max Then Max = tmp(z)
    Next
End Function

Public Function Min(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
    Dim tmp() As Double
    Dim n As Long
    Dim z As Long
    n = 0
    Call AppendEvalArg(a, tmp, n)
    If Not IsMissing(b) Then Call AppendEvalArg(b, tmp, n)
    If Not IsMissing(c) Then Call AppendEvalArg(c, tmp, n)
    If Not IsMissing(d) Then Call AppendEvalArg(d, tmp, n)
    If Not IsMissing(e) Then Call AppendEvalArg(e, tmp, n)
    If Not IsMissing(f) Then Call AppendEvalArg(f, tmp, n)
    If Not IsMissing(g) Then Call AppendEvalArg(g, tmp, n)
    If Not IsMissing(h) Then Call AppendEvalArg(h, tmp, n)
    If Not IsMissing(i) Then Call AppendEvalArg(i, tmp, n)
    If Not IsMissing(j) Then Call AppendEvalArg(j, tmp, n)
    If Not IsMissing(k) Then Call AppendEvalArg(k, tmp, n)
    If Not IsMissing(l) Then Call AppendEvalArg(l, tmp, n)
    Min = tmp(0)
    For z = 1 To n - 1
        If tmp(z) < Min Then Min = tmp(z)
    Next
End Function

Public Function MaxAbs(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
    Dim tmp() As Double
    Dim n As Long
    Dim z As Long
    n = 0
    Call AppendEvalArg(a, tmp, n)
    If Not IsMissing(b) Then Call AppendEvalArg(b, tmp, n)
    If Not IsMissing(c) Then Call AppendEvalArg(c, tmp, n)
    If Not IsMissing(d) Then Call AppendEvalArg(d, tmp, n)
    If Not IsMissing(e) Then Call AppendEvalArg(e, tmp, n)
    If Not IsMissing(f) Then Call AppendEvalArg(f, tmp, n)
    If Not IsMissing(g) Then Call AppendEvalArg(g, tmp, n)
    If Not IsMissing(h) Then Call AppendEvalArg(h, tmp, n)
    If Not IsMissing(i) Then Call AppendEvalArg(i, tmp, n)
    If Not IsMissing(j) Then Call AppendEvalArg(j, tmp, n)
    If Not IsMissing(k) Then Call AppendEvalArg(k, tmp, n)
    If Not IsMissing(l) Then Call AppendEvalArg(l, tmp, n)
    MaxAbs = tmp(0)
    For z = 1 To n - 1
        If Abs(tmp(z)) > Abs(MaxAbs) Then MaxAbs = tmp(z)
    Next
End Function

Public Function Sum(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
    Dim tmp() As Double
    Dim n As Long
    Dim z As Long
    n = 0
    Call AppendEvalArg(a, tmp, n)
    If Not IsMissing(b) Then Call AppendEvalArg(b, tmp, n)
    If Not IsMissing(c) Then Call AppendEvalArg(c, tmp, n)
    If Not IsMissing(d) Then Call AppendEvalArg(d, tmp, n)
    If Not IsMissing(e) Then Call AppendEvalArg(e, tmp, n)
    If Not IsMissing(f) Then Call AppendEvalArg(f, tmp, n)
    If Not IsMissing(g) Then Call AppendEvalArg(g, tmp, n)
    If Not IsMissing(h) Then Call AppendEvalArg(h, tmp, n)
    If Not IsMissing(i) Then Call AppendEvalArg(i, tmp, n)
    If Not IsMissing(j) Then Call AppendEvalArg(j, tmp, n)
    If Not IsMissing(k) Then Call AppendEvalArg(k, tmp, n)
    If Not IsMissing(l) Then Call AppendEvalArg(l, tmp, n)
    Sum = 0#
    For z = 0 To n - 1
        Sum = Sum + tmp(z)
    Next
End Function

Public Function Avg(ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant, Optional ByVal d As Variant, Optional ByVal e As Variant, Optional ByVal f As Variant, Optional ByVal g As Variant, Optional ByVal h As Variant, Optional ByVal i As Variant, Optional ByVal j As Variant, Optional ByVal k As Variant, Optional ByVal l As Variant) As Double
    Dim tmp() As Double
    Dim n As Long
    Dim z As Long
    n = 0
    Call AppendEvalArg(a, tmp, n)
    If Not IsMissing(b) Then Call AppendEvalArg(b, tmp, n)
    If Not IsMissing(c) Then Call AppendEvalArg(c, tmp, n)
    If Not IsMissing(d) Then Call AppendEvalArg(d, tmp, n)
    If Not IsMissing(e) Then Call AppendEvalArg(e, tmp, n)
    If Not IsMissing(f) Then Call AppendEvalArg(f, tmp, n)
    If Not IsMissing(g) Then Call AppendEvalArg(g, tmp, n)
    If Not IsMissing(h) Then Call AppendEvalArg(h, tmp, n)
    If Not IsMissing(i) Then Call AppendEvalArg(i, tmp, n)
    If Not IsMissing(j) Then Call AppendEvalArg(j, tmp, n)
    If Not IsMissing(k) Then Call AppendEvalArg(k, tmp, n)
    If Not IsMissing(l) Then Call AppendEvalArg(l, tmp, n)
    Avg = 0#
    For z = 0 To n - 1
        Avg = Avg + tmp(z)
    Next
    Avg = Avg / n
End Function

Private Sub AppendEvalArg(ByVal value As Variant, ByRef values() As Double, ByRef count As Long)
    Dim j As Long
    If IsArray(value) Then
        For j = 0 To UBound(value)
            Call AppendEvalScalar(CDbl(value(j)), values, count)
        Next
    Else
        Call AppendEvalScalar(CDbl(value), values, count)
    End If
End Sub

Private Sub AppendEvalScalar(ByVal value As Double, ByRef values() As Double, ByRef count As Long)
    If count = 0 Then
        ReDim values(0)
    Else
        ReDim Preserve values(count)
    End If
    values(count) = value
    count = count + 1
End Sub
