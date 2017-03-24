Attribute VB_Name = "Bateman"
Option Explicit: Option Base 1
DefDbl A-Z

Dim Nbd%(2), Mbd%(2), Ly#(2, 6), BatemanConst#(3, 2, 5)

Private Sub Bateman()
' Calculate the Bateman constants for U/Pb ratios of a system not in secular equilibrium.

Dim i%, j%, k%, H%, Numer, Denom

' 238 -> 234 -> 230 -> 206
' 235 -> 231 -> 207
Nbd(1) = 4: Nbd(2) = 3  ' #nuclides to keep track of in each decay scheme
Mbd(1) = 3: Mbd(2) = 2  ' #radioactive daughters permitted to be nonzero
Ly(1, 1) = Lambda238 / 1000000#: Ly(1, 2) = Lambda234
Ly(1, 3) = Lambda230 ': Ly(1, 4) = Lambda226:'Ly(1, 5) = Lambda210
Ly(2, 1) = Lambda235 / 1000000#
Ly(2, 2) = Lambda231 ': Ly(2, 3) = Lambda227
Ly(1, 6) = 0: Ly(2, 4) = 0:
Ly(2, 5) = 0: Ly(2, 6) = 0
Ly(2, 3) = 0: Ly(1, 5) = 0: Ly(1, 4) = 0

For H = 1 To 3               '  The BatemanConst(1,i,j) are for U238 and U235;

  For i = 1 To 2             '  the BatemanConst(2,i,j) are for U234 and Pa231,

    For j = H To Nbd(i) - 1  '  the BatemanConst(3,i,j) are for Th230.
      Numer = 1: Denom = 1

      For k = H To Nbd(i) - 1    ' See for example Kirby (1973), Atom. Energy
        Numer = Numer * Ly(i, k) '  Comm. Res. & Devel. Rept. MLM-2036 p. 25-
      Next k                     '  43; or Ivanovich & Harmon, Uranium-Series

      For k = H To Nbd(i)        '  Disequilibrium, Oxford, 2nd ed., p. 18.

        If k <> j Then
          Denom = Denom * (Ly(i, k) - Ly(i, j))
        End If

      Next k

      BatemanConst(H, i, j) = Numer / Denom
Next j, i, H

End Sub

Public Function DisEq76ratio(AgeMa, Init234238ar, Init230238ar, Init231235ar)
GetConsts ' Returns 207Pb*/206Pb* ratio for a non-secular-equilibrium U-Pb system

DisEq76ratio = DisEq75Ratio(AgeMa, Init231235ar) / _
  DisEq68Ratio(AgeMa, Init234238ar, Init230238ar) / Uratio
End Function

Public Function DisEq68Age(Pb206U238, Init234238ar, Init230238ar)
Dim T# ' Returns age in Ma for a non-secular-equilibrium 238U-206Pb system

GetConsts
Yrat(1, 1) = Init234238ar: Yrat(1, 2) = Init230238ar
'Yrat(2, 1) = Init231235
SecularEquil = False
Bateman
DisEq68Age = DisEqAge(1, (Pb206U238)) / 1000000#
End Function

Public Function DisEq68Ratio(AgeMa, Init234238ar, Init230238ar)
GetConsts ' Returns 206Pb/238U for a non-secular-equilibrium 238U-206Pb system
Yrat(1, 1) = Init234238ar: Yrat(1, 2) = Init230238ar
'Yrat(2, 1) = Init231235
SecularEquil = False
Bateman
DisEq68Ratio = DisEqRatio(False, 1, AgeMa * 1000000#)
End Function

Public Function DisEq75Age(Pb207U235, Init231235ar)
Dim T# ' Returns age in Ma for a non-secular-equilibrium 235U-207Pb system
GetConsts
'Yrat(1, 1) = Init234238ar: Yrat(1, 2) = Init230238ar
Yrat(2, 1) = Init231235ar
SecularEquil = False
Bateman
DisEq75Age = DisEqAge(2, (Pb207U235)) / 1000000#
End Function

Public Function DisEq75Ratio(AgeMa, Init231235ar)
GetConsts ' Returns 207Pb/235U for a non-secular-equilibrium 235U-207Pb system
'Yrat(1, 1) = Init234238ar: Yrat(1, 2) = Init230238ar
Yrat(2, 1) = Init231235ar
SecularEquil = False
Bateman
DisEq75Ratio = DisEqRatio(False, 2, AgeMa * 1000000#)
End Function

Private Function DisEqRatio(Deriv As Boolean, WhichParent%, T#)
' Deriv=False for Ratio, True for Deriv; WhichParent=1 for 206/238, 2 for 207/235;
'  t is age in yrs; DisEqRatio is 206/238 or 207/235 atomic ratio (RadDeriv=1),
'  or 1st derivative with respect to time (RatDeriv=2).
' No subs/functions used.

Dim Bad As Boolean, i%, j%, M%, m1%
Dim Result, Lambda, ExpT, Exp1, Mterm, Sterm, Eterm
Dim s1(3), s2(3), Aterm(3)

i = WhichParent: Result = 0

If SecularEquil Then
  If i = 1 Then Lambda = Lambda238 Else Lambda = Lambda235
  DisEqRatio = Exp(Lambda * T) - 1

  Exit Function

End If

Aterm(1) = 1
For j = 2 To Mbd(i)
  Aterm(j) = Ly(i, 1) / Ly(i, j) * Yrat(i, j - 1)
Next j

Eterm = Ly(i, 1) * T
EtermHandle Eterm, Exp1, Bad

If Bad Then GoTo done

For M = 1 To 3            ' m=1 for decay of U238 or U235
  s1(M) = 0: s2(M) = 0    ' m=2 for decay of initial U234 or Pa231
  m1 = M - 1              ' m=3 for decay of initial Th230

  For j = 1 To Nbd(i) - 1 ' m=4 would be for initial Ra226 (assumed to be zero)
    Eterm = -Ly(i, j) * T
    EtermHandle Eterm, ExpT, Bad

    If Bad Then GoTo done

    If j > m1 Then
      Sterm = BatemanConst(M, i, j) * ExpT
      s1(M) = s1(M) + Sterm

      If Deriv Then
        s2(M) = s2(M) + Ly(i, j) * Sterm
      End If

    End If

  Next j

Next M

For M = 1 To 3
  Mterm = Aterm(M) * Exp1

  If Deriv Then
    Result = Result + Mterm * (Ly(i, 1) * (1 + s1(M)) - s2(M))
  Else
    Result = Result + Mterm * (1 + s1(M))
  End If

Next M

done: DisEqRatio = Result
End Function

Private Function DisEqAge(Which%, Ratio#)
' Which=1 for 206/238 age, 2 for 207/235 age; Ratio is 206/238 or 207/235
'  (atomic); Returns age in years.

Dim Bad As Boolean, i%, j%, k%, M%, m1%, Iter%
Dim T, Delta, test, Eterm, Sterm, Mterm, ExpT, Exp1, tL, Rat, Deriv
Dim Aterm(3), f(3), Fp(3), s(3, 2)

SecularEquil = False
i = Which

If Ratio <= -1 Or Ratio > 1000000000# Then DisEqAge = 0: Exit Function

If Ly(i, 1) = 0 Then
  If Which = 1 Then tL = Lambda238 Else tL = Lambda235
  Ly(i, 1) = tL / 1000000#
End If

T = Log(1 + Ratio) / Ly(i, 1) ' Start with age calculated assuming secular equilibrium.

If Not SecularEquil Then
  Aterm(1) = 1

  For j = 2 To Mbd(i)
    Aterm(j) = Ly(i, 1) / Ly(i, j) * Yrat(i, j - 1)
  Next j

  Iter = 0

  Do
    Iter = 1 + Iter: Eterm = Ly(i, 1) * T
    EtermHandle Eterm, Exp1, Bad
    If Bad Then GoTo done
    'If Abs(Eterm) < MAXEXP Then Exp1 = Exp(Eterm) Else t = 0: GoTo Done

    For M = 1 To 3              ' m=1 for decay of U238 or U235
      s(M, 1) = 0: s(M, 2) = 0  ' m=2 for decay of initial U234 or Pa231
      m1 = M - 1                ' m=3 for decay of initial Th230

      For j = 1 To Nbd(i) - 1   ' m=4 would be for initial Ra226 (assumed to be zero)
        Eterm = -Ly(i, j) * T
        EtermHandle Eterm, ExpT, Bad

        If Bad Then GoTo done

        If j > m1 Then
          Sterm = BatemanConst(M, i, j) * ExpT
          s(M, 1) = s(M, 1) + Sterm
          s(M, 2) = s(M, 2) + Ly(i, j) * Sterm

        End If

      Next j


    Next M
    Rat = 0: Deriv = 0

    For M = 1 To 3
      Mterm = Aterm(M) * Exp1
      Rat = Rat + Mterm * (1 + s(M, 1))
      Deriv = Deriv + Mterm * (Ly(i, 1) * (1 + s(M, 1)) - s(M, 2))

    Next M

    If Abs(Deriv) < 1E-30 Then DisEqAge = 0: Exit Function

    Delta = (Rat - Ratio) / Deriv        ' Newton's method.
    T = T - Delta                        '     "       "

    If Abs(Ratio) > 0.000001 Then
      test = Abs((Rat - Ratio) / Ratio)  ' Converge on ratio, not age.
    Else
      test = Abs(Delta)

    End If

  Loop Until test < 0.000001

End If

done: DisEqAge = Max(0, Drnd(T, 6))
End Function

Public Function DisEqPbPbAge(R76, Init234238ar, Init230238ar, Init231235ar) As Variant
' Solve for both possible solutions of 207Pb/206Pbage if used as a 1x2 array function, or youngest if not.

Dim t1, t2, T, s, r1, r2, M, Outp(1, 2)

SecularEquil = False: M = 1000000#
DisEqPb76Age R76, 0, t1, 0, Init234238ar, Init230238ar, Init231235ar
t2 = t1
r1 = DisEq76ratio(t1 / M, Init234238ar, Init230238ar, Init231235ar)
r2 = DisEq76ratio(t1 / M + 0.001, Init234238ar, Init230238ar, Init231235ar)
s = (r2 - r1) / 0.001 ' Ratio-AgeMa slope

If s < 0 Then ' Find second solution
  T = t1

  Do
    T = T + (M / 10)
    r2 = DisEq76ratio(T / M, Init234238ar, Init230238ar, Init231235ar)

  Loop Until r2 > r1 Or T > (10 * M)

  If T < (10 * M) Then t2 = T
End If

Outp(1, 1) = t1 / M

If Selection.Columns.Count > 1 Then Outp(1, 2) = t2 / M
DisEqPbPbAge = Outp
End Function

Sub DisEqPb76Age(R76, R76Error, T, Terror, Init234238ar, Init230238ar, Init231235ar)
' Solve a radiogenic Pb-207/206 age.  R76 is the 207*/206* ratio, t1 is
'  a known age (0 for a single-stage 7/6 age).  Solve for T & T error.
' NOTE: This routine will only find the younger of the two solutions!  Setting the trial age to
'  an older value will often result in only the younger solution as well.

Dim Iter%
Dim Toler, R68, R75, R68p, R75p, R76t, Deriv, NewT, test, M
SecularEquil = False
GetConsts
Yrat(1, 1) = Init234238ar: Yrat(1, 2) = Init230238ar
Yrat(2, 1) = Init231235ar: M = 1000000#
Bateman
Toler = 0.001  ' fractional
T = 10000#     ' in years!

Do

  If Abs(T / M) > 50000# Then GoTo No76Age

  Iter = 1 + Iter
  If Iter > 50 Then GoTo No76Age
  R68 = DisEqRatio(False, 1, (T))
  R75 = DisEqRatio(False, 2, (T))
  R68p = DisEqRatio(True, 1, (T))
  R75p = DisEqRatio(True, 2, (T))

  If R68 = 0 Or R75 = 0 Or R68p = 0 Or R75p = 0 Then GoTo No76Age

  R76t = R75 / R68 / Uratio
  Deriv = (R75p - R75 / R68 * R68p) / R68 / Uratio
  NewT = T - (R76t - R76) / Deriv
  If NewT < 0 Then NewT = 100
  test = (NewT - T) / T
If Abs(test) < Toler Then Exit Do
  T = NewT
Loop

Terror = Abs(R76Error / Deriv)
Exit Sub

No76Age: T = 0: Terror = 0
End Sub
