Attribute VB_Name = "GeneralMonte"
Option Private Module
Option Base 1: Option Explicit

Sub GaussCorrel(ByVal X#, ByVal sigmaX#, ByVal y#, _
  ByVal SigmaY#, ByVal Rho#, Xstar#, Ystar#)
Attribute GaussCorrel.VB_ProcData.VB_Invoke_Func = " \n14"
Dim A#, b#, c#, Ar#
Dim xx#, yy#, EpSqrR#, dx#, dY#
Const Eps = 0.0001
If Rho = 0 Then
  Xstar = Gaussian(X, sigmaX)
  Ystar = Gaussian(y, SigmaY)
  Exit Sub
End If
Ar = Abs(Rho)
' Use a small value for the errors, otherwise the ratios a/c & b/c won't
'  have a normal distr.  Mean xX, yY, a, b, & c are all 1.
EpSqrR = Eps * Sqr(1 - Ar)
A = Gaussian(1, EpSqrR)  ' A, B are randomly & normally distributed
b = Gaussian(1, EpSqrR)  '   about 1 with sigma=EpSqrR.
c = Gaussian(1, Eps * Sqr(Ar))
xx = A / c:     yy = b / c
dx = (xx - 1) * sigmaX / Eps
Xstar = X + dx
Ystar = y + Sgn(Rho) * (yy - 1) * SigmaY / Eps
' From input X-Y, errors, & error correl, return Xstar,Ystar which have
'  the appropriate Gaussian distr.
' Uses the fact that if X=a/c & Y=b/c & RhoXY>=0, then
' (SigmaC/C)^2 = (SigmaX/X)*(SigmaY/Y)*RhoXY,
' (SigmaA/A)^2 = (SigmaX/X)^2-(SigmaC/C)^2
' (SigmaB/B)^2 = (SigmaY/Y)^2-(SigmaC/C)^2
End Sub

Function DblExpDev(Optional Dummy) As Double
Attribute DblExpDev.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns a double-exponentially distributed, random deviate with mean zero and std deviation 1.
' 68.3%  conf. limit is +-0.81
' 95.0%  conf. limit is +-2.12
' 99.73% conf. limit is +-4.2
' Pass Dummy as Rand() to force recalculation
Const Sqrt2 = 1.4142135623731
DblExpDev = -Sgn(0.5 - Rnd) * Log(Rnd) / Sqrt2
End Function

Function PoiDev(ByVal Counts As Long, Optional Dummy) As Double
Attribute PoiDev.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns a random number from a Poisson distribution with mean Counts
' Modified from Numerical Recipes (1986) p. 207-208
' Pass Dummy as Rand() to force recalculation
Dim G#, EM#, T#, ALXM#, y#, SQ#
Static OldCounts As Boolean, Started As Boolean
With Application
  If Counts = 0 Then
    EM = 0
  ElseIf Counts < 12 Then
    If Counts <> OldCounts Then
      OldCounts = Counts: G = Exp(-Counts)
    End If
    EM = -1: T = 1
    Do
      EM = 1 + EM: T = T * Rnd
    Loop Until T <= G
  Else
    If Counts <> OldCounts Then
      OldCounts = Counts: SQ = Sqr(2 * Counts)
      ALXM = Log(Counts)
      G = Counts * ALXM - .GammaLn(Counts + 1)
    End If
    Do
      Do
        y = Tan(pi * Rnd)
        EM = SQ * y + Counts
      Loop Until EM >= 0
      EM = Int(EM)
      T = 0.9 * (1 + y * y) * Exp(EM * ALXM - .GammaLn(EM + 1) - G)
    Loop Until Rnd <= T
  End If
End With
PoiDev = EM
End Function

Function RandomT(ByVal Sigma, ByVal DegFree)
' Returns random number with a t distribution using Excel's Tinv function
RandomT = Application.TInv(Rnd, DegFree) * Sgn(0.5 - Rnd) * Sigma
End Function
