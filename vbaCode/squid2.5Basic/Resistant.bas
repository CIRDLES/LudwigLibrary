Attribute VB_Name = "Resistant"
'
'   ****************************************************************************
'
'   SQUID2 is a program for processing SHRIMP data
'
'   Program author: Dr Ken Ludwig (Berkeley Geochronology Center)
'
'   Supporters (members of the SQUID2 Development Group):
'       - Geoscience Australia
'       - United States Geological Survey
'       - Berkeley Geochronology Center
'       - All-Russian Institute of Geological Research
'       - Australian Scientific Instruments
'       - Geological Survey of Canada
'       - John de Laeter School of Mass Spectrometry (Curtin University)
'       - National Institute of Polar Research (Japan)
'       - Research School of Earth Sciences (Australian National University)
'       - Stanford University
'
'   ****************************************************************************
'
'   Copyright (C) 2009, the Commonwealth of Australia represented by Geoscience
'                 Australia, GPO box 378, Canberra ACT 2601, Australia
'   All rights reserved.
'   (http://www.ga.gov.au/minerals/research/methodology/geochron/index.jsp)
'
'   This file is part of SQUID2.
'
'   SQUID2 is free software. Permission to use, copy, modify, and distribute
'   this software for any purpose without fee is hereby granted under the terms
'   of the GNU General Public License as published by the Free Software
'   Foundation, either version 3 of the License, or (at your option) any later
'   version, provided that this notice is included in all copies of any
'   software which is, or includes, a copy or modification of this software and
'   in all copies of the supporting documentation for such software.
'
'   SQUID2 is distributed in the hope that it will be useful, but WITHOUT ANY
'   WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
'   FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
'   details.
'
'   You should have received a copy of the GNU General Public License along
'   with this program; if not see <http://www.gnu.org/licenses/gpl.html>.
'
'   ****************************************************************************
'
' 09/03/02 -- All lower array bounds explicit

Option Explicit
Option Base 1

Sub SecularTrend(LowessVars As Lowess, Optional AutoWindow As Boolean = False)

Dim tB As Boolean, LoopCt%, i%, N%, df%, Window%, ActualWindow%
Dim SumRobustWts#, InternalWtdResidSq#, SumInternalWtdResidsSq#, MedSig#
Dim SumRobustWtdLowessY#, SumInternalWts#, InvV#, SumInvV#, DataPtErrsOnlyMSWD#
Dim MeanRobustWt#, Mean#, SigmaMean#, ExtSigma#, MSWD#, Probfit#, MinExternal#
Dim SumInvVw#, InitialMswd#, InitialProbFit#, WtdDataPtErrsOnlyMSWD#
Dim SumRobustWtdResidSq#, SumRobustWtdResidSqY#, temp1#, temp2#
Dim FractPoints#, ExternalPerr#, v#, InvVr#, SumInvVr#
Dim SigmaTot#(), Y#(), RobustWts#(), Resids#(), yy#(), LastY#()
Dim Ycel As Range, ErCel As Range, xv() As Variant

N = 0
If AutoWindow Then MinExternal = foUser("MinUPbErr")
With LowessVars

  For i = 1 To .rY.Rows.Count
    Set Ycel = .rY(i, 1)
    Set ErCel = .rYsig(i, 1)
    tB = Not Ycel.Font.Strikethrough

    If tB Then

      If IsNumeric(Ycel) And IsNumeric(ErCel) Then

        If Ycel.Value <> 0 And ErCel.Value > 0 Then
          N = 1 + N
          ReDim Preserve Y(1 To N), .daYsig(1 To N), xv(1 To N)
          xv(N) = .rX(i)
          Y(N) = Ycel
          v = ErCel

          If .bPercentErrs Then
            .daYsig(N) = v / 100 * Y(N)
          Else
            .daYsig(N) = v
          End If

        End If

      End If

    End If

  Next i

  ReDim Sorted&(1 To N), LastY(1 To N)

  For i = 1 To N: Sorted(i) = i: Next
  QuickIndxSort xv, Sorted
  ReDim SigmaTot(1 To N), .daY(1 To N), RobustWts(1 To N), Resids(1 To N), .daX(1 To N), yy(1 To N)

  For i = 1 To N
    .daX(i) = xv(i)
    yy(i) = Y(Sorted(i))
  Next

  .iWindow = IIf(AutoWindow, 3, .iWindow - 1)

  Do
    .iWindow = 1 + .iWindow
    FractPoints = .iWindow / N
    FractPoints = fvMin(FractPoints, 1)
    .iActualWindow = FractPoints * N

    Lowess .daX, yy, N, FractPoints, 2, 0, .daY, RobustWts, Resids

    LoopCt = 0: .dExtSigma = 0
    df = .iActualWindow - 3

    Do
      LoopCt = 1 + LoopCt

      For i = 1 To N
        SigmaTot(i) = sqR(.daYsig(i) ^ 2 + .dExtSigma ^ 2)
      Next i

      SumInternalWts = 0:           SumRobustWts = 0
      SumInternalWtdResidsSq = 0:   SumRobustWtdLowessY = 0
      SumRobustWtdResidSqY = 0:     SumRobustWtdResidSq = 0
      SumInvVr = 0: SumInvVw = 0: SumInvV = 0

      For i = 1 To N
        ' wtd by data-pt errs only
        InvV = 1 / SigmaTot(i) ^ 2
        SumInvV = SumInvV + InvV
        SumInvVw = SumInvVw + RobustWts(i) * InvV
        SumRobustWts = SumRobustWts + RobustWts(i)
        SumInternalWts = SumInternalWts + InvV
        InternalWtdResidSq = InvV * Resids(i) ^ 2
        SumInternalWtdResidsSq = SumInternalWtdResidsSq + InternalWtdResidSq

        InvVr = RobustWts(i) * (Resids(i) / .daYsig(i)) ^ 2
        SumInvVr = InvVr + SumInvVr
        SumRobustWtdResidSq = SumRobustWtdResidSq + InvVr
        SumRobustWtdResidSqY = SumRobustWtdResidSqY + .daY(i) * InvVr

        ' wtd by data-pt errs AND robust wts

        SumRobustWtdLowessY = SumRobustWtdLowessY + .daY(i) * RobustWts(i) 'RobustTotsigWt
      Next i

      DataPtErrsOnlyMSWD = SumInternalWtdResidsSq / (N - 1)

      temp1 = SumRobustWtdResidSq / SumRobustWts
      temp2 = .iActualWindow / (.iActualWindow - 3)
      WtdDataPtErrsOnlyMSWD = temp2 * temp1 * (N / SumRobustWts)

      .dMSWD = DataPtErrsOnlyMSWD * fdMSWDmult(.iActualWindow, N)

      If LoopCt = 1 Then
        .dProbfit = ChiSquare(.dMSWD, df)

        ' ??? should be just "= WtdDataPtErrsOnlyMSWD" ????
        .dInitialMSWD = foAp.Min(.dMSWD, WtdDataPtErrsOnlyMSWD)

        InitialProbFit = ChiSquare(.dInitialMSWD, df)
        If .dMSWD <= 1.1 Or InitialProbFit > 0.1 Then
          Exit Do
        Else
          MedSig = foAp.Median(.daYsig)
          .dExtSigma = MedSig
        End If

      ElseIf LoopCt = 99 Then
        MsgBox "Loop-locked in sub SecularTrend - calculation unreliable", , pscSq
        Exit Do

      ElseIf .dMSWD < 1.001 And .dMSWD > 0.999 Then
        Exit Do

      Else
        .dExtSigma = .dExtSigma * sqR(.dMSWD)
      End If
    Loop

    MeanRobustWt = SumRobustWts / N
    .dMean = SumRobustWtdLowessY / SumRobustWts
    .dSigmaMean = 1 / sqR(SumInvVw * MeanRobustWt)

    Mean = .dMean:                 SigmaMean = .dSigmaMean
    ExtSigma = .dExtSigma:         MSWD = .dMSWD
    Probfit = .dProbfit:           Window = .iWindow
    ActualWindow = .iActualWindow: InitialMswd = .dInitialMSWD
    ExternalPerr = 100 * ExtSigma / Mean
  Loop Until Not AutoWindow Or ExternalPerr > MinExternal Or .iWindow = N

  If AutoWindow And LoopCt > 1 And .iWindow > 4 Then
    .dMean = Mean:                 .dSigmaMean = SigmaMean
    .dExtSigma = ExtSigma:         .dMSWD = MSWD
    .dProbfit = Probfit:           .iWindow = Window
    .iActualWindow = ActualWindow: .dInitialMSWD = InitialMswd
  End If

End With
End Sub

Sub Lowess(X#(), Y#(), ByVal Npts%, ByVal FractPoints, ByVal Nsteps%, _
            ByVal Delta, LowessY#(), RobustWts#(), Resids#())

Dim OK As Boolean, Iter%, Last%, Nleft%, Nright%, NLocalPts%, i%, j%
Dim Indx&()
Dim d1#, d2#, m1#, m2#, r#, ColOne#, c9#, Cmad#, Denom#, Alpha#, Cut#

If Npts < 2 Then GoTo Bad
NLocalPts = FractPoints * Npts

For Iter = 1 To Nsteps + 1
  ' robustness iterations
  Nleft = 1: Nright = NLocalPts: Last = 0       ' index of prev estimated point
  i = 1           ' index of current point

  Do
    Do While Nright < Npts
      ' move nleft, nright to right if radius decreases
      d1 = X(i) - X(Nleft)
      d2 = X(Nright + 1) - X(i)
      ' if d1<=d2 with x(nright+1)==x(nright), lowest fixes
      If d1 <= d2 Then Exit Do
       ' radius will not decrease by move right
       Nleft = Nleft + 1
       Nright = Nright + 1
    Loop

    Lowest X, Y, Npts, X(i), LowessY(i), Nleft, Nright, Resids, Iter > 1, RobustWts, OK

    ' fitted value at x(i)
    If Not OK Then LowessY(i) = Y(i)

    ' all weights zero - copy over value (all rw==0)
    If Last < (i - 1) Then  ' skipped points -- interpolate
      Denom = X(i) - X(Last) ' non-zero - proof?

      For j = Last + 1 To i - 1
        j = j + 1
        Alpha = (X(j) - X(Last)) / Denom
        LowessY(j) = Alpha * LowessY(i) + (1 - Alpha) * LowessY(Last)
      Next j

    End If

    Last = i    ' last point actually estimated
    Cut = X(Last) + Delta ' x coord of close points

    For i = Last + 1 To Npts - 1
     i = i + 1 ' find close points
     If X(i) > Cut Then Exit For ' i one beyond last pt within Cut

     If X(i) = X(Last) Then ' exact match in x
       LowessY(i) = LowessY(Last)
       Last = i
     End If

    Next i

     i = fvMax(Last + 1, i - 1)
    ' back 1 point so interpolation within delta, but always go forward
  Loop Until Last >= Npts

  For i = 1 To Npts   ' residuals
    Resids(i) = Y(i) - LowessY(i)
  Next i

  If Iter > Nsteps Then Exit For ' compute robustness weights except last time

  For i = 1 To Npts
    RobustWts(i) = Abs(Resids(i))
  Next i

  QuickSort RobustWts
  m1 = 1 + Npts / 2: m2 = Npts - m1 + 1
  Cmad = 3# * (RobustWts(m1) + RobustWts(m2)) ' 6 median abs resid
  c9 = 0.999 * Cmad: ColOne = 0.001 * Cmad

  For i = 1 To Npts
    r = Abs(Resids(i))

    If r <= ColOne Then
      RobustWts(i) = 1# ' near 0, avoid underflow
    ElseIf r > c9 Then
      RobustWts(i) = 0# ' near 1, avoid underflow
    Else
      RobustWts(i) = (1# - (r / Cmad) ^ 2) ^ 2
    End If

  Next i

Next Iter

Exit Sub
Bad: ComplainCrash "err in Lowess"
End Sub

Sub Lowest(X#(), Y#(), ByVal Npts%, xs#, LowessY#, Nleft%, Nright%, w#(), _
  ByVal UserW As Boolean, RobustWts#(), OK As Boolean)
Dim j%, Nrt%
Dim rRange#, c#, h#, h9#, h1#, a#, b#, r#

ReDim w(1 To Npts)

rRange = X(Npts) - X(1)
h = fvMax(xs - X(Nleft), X(Nright) - xs)
h9 = 0.999 * h
h1 = 0.001 * h
a = 0#   ' sum of weights

For j = Nleft To Npts
  ' compute weights (pick up all ties on right)
  w(j) = 0#
  r = Abs(X(j) - xs)

  If r <= h9 Then ' small enough for non-zero weight

    If (r > h1) Then
      w(j) = (1# - (r / h) ^ 2) ^ 2 ' Bisquare    '3) ^ 3 ' Tricube
    Else
      w(j) = 1#
    End If

    If UserW Then w(j) = RobustWts(j) * w(j)
    a = a + w(j)
  ElseIf X(j) > xs Then
    Exit For  ' get out at first zero wt on right
  End If

Next j

Nrt = j - 1   ' rightmost pt (may be greater than nright because of ties)

If a <= 0# Then
  OK = False
Else          ' weighted least squares
  OK = True

  For j = Nleft To Nrt
    w(j) = w(j) / a       ' make sum of w(j) =1
  Next j

  If h > 0# Then          ' use linear fit
    a = 0#

    For j = Nleft To Nrt
      a = a + w(j) * X(j) ' weighted center of x values
    Next j

    b = xs - a
    c = 0#

    For j = Nleft To Nrt
      c = c + w(j) * (X(j) - a) ^ 2
    Next j

    If sqR(c) > (0.001 * rRange) Then
      ' points are spread out enough to compute slope
      b = b / c

      For j = Nleft To Nrt
        w(j) = w(j) * (1# + b * (X(j) - a))
      Next j

    End If

    LowessY = 0#

    For j = Nleft To Nrt
       LowessY = LowessY + w(j) * Y(j)
    Next j

  End If

End If

End Sub

Public Function sqBiweight(Vals As Variant, Optional Tuning = 6, _
  Optional MeanOnly As Boolean = False, _
  Optional AllComers As Boolean = False, _
  Optional IsoplotStyle As Boolean = False)

If VarType(Vals) = vbString Then Set Vals = Range(Vals)
On Error GoTo 1
sqBiweight = BiWt(Vals, Tuning, MeanOnly, AllComers, IsoplotStyle)
Exit Function
1: On Error GoTo 0
End Function

Sub GetMAD(X#(), ByVal N&, MedianVal#, Madd#, Err95#)
' Determine the Median Absolute Deviation (MAD) from the median for the first
'   N values in vector X() with median MedianVal.
Dim i&, Tstar#, AbsDev#()
ReDim AbsDev(1 To N)

For i = 1 To N
  AbsDev(i) = Abs(X(i) - MedianVal)
Next i

Madd = foAp.Median(AbsDev())

Select Case N ' KRL-derived numerical approx., valid for normal distr. w. Tuning=9
  Case Is < 2: Tstar = 0
  Case 2: Tstar = 12.7
  Case 3: Tstar = 15.3
  Case Else
     Tstar = 3.54 / sqR(N) - 3.92 / N + 70.9 / (N * N) - 60.6 / N ^ 3
End Select

Err95 = Tstar * Madd
End Sub

Public Function sqMad(v As Variant) As Variant ' Return the median absolute deviation from the median
Dim i&, N&, MadVal#, MedianVal#, vv#()
If TypeName(v) = "Range" Then N = v.Count Else N = UBound(v)
On Error GoTo 1
i = sqR(-1)
MedianVal = foAp.Median(v)
ReDim vv(1 To N)

For i = 1 To N: vv(i) = v(i): Next i

GetMAD vv, N, MedianVal, MadVal, 0
sqMad = MadVal
Exit Function

1: On Error GoTo 0
sqMad = "#NUM!"
End Function

Function fdMedianConfLevel#(ByVal N&) ' Confidence limit (%) of error on median
Dim Conf#, a As Variant

If N > 25 Then
  Conf = 95
ElseIf N > 2 Then
' Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
  a = Array(75, 87.8, 93.8, 96.9, 98.4, 93#, 96.1, 97.9, 93.5, 96.1, _
    97.8, 94.3, 96.5, 97.9, 95.1, 96.9, 93.6, 95.9, 97.3, 94.8, 96.5, 97.7, 95.7)
  Conf# = a(N - 2)
End If

fdMedianConfLevel = Drnd(Conf, 5)
End Function

Function fdMedianUpperLim#(v As Variant, Optional N) ' Upper error on median of V()
Dim u&
Const q = "11111222333444556667778"

If fbIM(N) Then
  If TypeName(v) = "Range" Then N = v.Count Else N = UBound(v)
End If

If N > 25 Then
  u = 0.5 * (N + 1 - 1.96 * sqR(N))
Else
' Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
  u = Val(Mid$(q, (N - 3) + 1, 1))  ' High rank (U-th largest)
End If

fdMedianUpperLim = foAp.Large(v, u)
End Function

Function fdMedianLowerLim#(v As Variant, Optional N) ' Lower error on median of V()
Dim Lwr&, Uppr&
Const Lr = "0304050607070809091011111213131414151616171818"

If fbIM(N) Then
  If TypeName(v) = "Range" Then N = v.Count Else N = UBound(v)
End If

If N > 25 Then
  Uppr = 0.5 * (N + 1 - 1.96 * sqR(N))
  Lwr = N + 1 - Uppr
Else
' Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
  Lwr = Val(Mid$(Lr, 2 * (N - 3) + 1, 2))  ' Low  rank (Lwr-th largest)
End If

fdMedianLowerLim = foAp.Large(v, Lwr)
End Function

Public Function fdNmad#(v#())
Dim i%, N%, Nn%, med#, Nm#, yr2#(), Resid#()
Const nMadConst = 1.4826

N = UBound(v): med = foAp.Median(v)
Nn = fvMax(3, N)
ReDim yr2(1 To N)

For i = 1 To N
  yr2(i) = (v(i) - med) ^ 2
Next i

fdNmad = nMadConst * (1 + 5 / (Nn - 2)) * sqR(foAp.Median(yr2()))
End Function

Sub QuickIndxSort(Vect(), Indx&(), _
  Optional ByVal LeftInd& = -2, Optional ByVal RightInd& = -2)
Dim i&, j&, MidInd&, TestVal As Variant, tmp As Variant

If LeftInd = -2 Then LeftInd = LBound(Vect)
If RightInd = -2 Then RightInd = UBound(Vect)

If LeftInd < RightInd Then
  MidInd = (LeftInd + RightInd) \ 2
  TestVal = Vect(MidInd)
  i = LeftInd: j = RightInd

  Do

    Do While Vect(i) < TestVal
      i = i + 1
    Loop

    Do While Vect(j) > TestVal
      j = j - 1
    Loop

    If i <= j Then
      tmp = Vect(i): Vect(i) = Vect(j): Vect(j) = tmp
      tmp = Indx(i): Indx(i) = Indx(j): Indx(j) = tmp
      i = i + 1:  j = j - 1
    End If

  Loop Until i > j

  ' Optimize sort by sorting smaller segment first
  If j <= MidInd Then
    QuickIndxSort Vect, Indx(), LeftInd, j
    QuickIndxSort Vect, Indx(), i, RightInd
  Else
    QuickIndxSort Vect, Indx(), i, RightInd
    QuickIndxSort Vect, Indx(), LeftInd, j
  End If

End If
End Sub

Sub QuickSort(Vect As Variant, _
  Optional ByVal LeftInd& = -2, Optional ByVal RightInd& = -2)
Dim i&, j&, MidInd&, TestVal#

If LeftInd = -2 Then LeftInd = LBound(Vect)
If RightInd = -2 Then RightInd = UBound(Vect)

If LeftInd < RightInd Then
  MidInd = (LeftInd + RightInd) \ 2
  TestVal = Vect(MidInd)
  i = LeftInd: j = RightInd

  Do

    Do While Vect(i) < TestVal
      i = i + 1
    Loop

    Do While Vect(j) > TestVal
      j = j - 1
    Loop

    If i <= j Then
      SwapElements Vect, i, j
      i = i + 1
      j = j - 1
    End If

  Loop Until i > j

  ' Optimize sort by sorting smaller segment first
  If j <= MidInd Then
    QuickSort Vect, LeftInd, j
    QuickSort Vect, i, RightInd
  Else
    QuickSort Vect, i, RightInd
    QuickSort Vect, LeftInd, j
  End If

End If
End Sub

' Used in QuickSort function
Private Sub SwapElements(Items As Variant, ByVal Item1&, ByVal Item2&)
Dim temp#
temp = Items(Item2)
Items(Item2) = Items(Item1)
Items(Item1) = temp
End Sub

Public Function RobReg(Xrange, Yrange, Optional WithYinter As Boolean = False, _
  Optional ByVal InclErr As Boolean = False)
Dim xdim1 As Boolean, ydim1 As Boolean
Dim i%, Nx%
Dim xyD#(), Yint#, Slope#, PlusSlope#, MinusSlope#, PlusYint#
Dim MinusYint#, SlopeErr#, YintErr#, XYclean#()
Dim v As Variant

xdim1 = False: ydim1 = False
On Error GoTo 5
v = Xrange(1, 1)
xdim1 = True
v = Yrange(1, 1)
ydim1 = True

5: On Error GoTo 1

If TypeName(Xrange) = "Range" And TypeName(Yrange) = "Range" And xdim1 And ydim1 Then
  Nx = Xrange.Rows.Count
  If Yrange.Rows.Count <> Nx Then GoTo 1
  If Yrange.Columns.Count <> 1 Or Xrange.Columns.Count <> 1 Then GoTo 1
ElseIf IsArray(Xrange) And IsArray(Yrange) Then
  Nx = UBound(Xrange, 1)
  If UBound(Yrange) <> UBound(Xrange) Then GoTo 1
Else
  RobReg = 0
  Exit Function
End If

If Nx < 3 Then GoTo 1
ReDim xyD(1 To Nx, 1 To 2)

For i = 1 To Nx
  xyD(i, 1) = Xrange(i).Value
  xyD(i, 2) = Yrange(i).Value
Next i

CleanData xyD, XYclean, Nx
Isoplot3.RobustReg2 XYclean, Slope, PlusSlope, MinusSlope, Yint, , _
                     MinusYint, PlusYint, , , , , True

If InclErr Then
  SlopeErr = Abs(PlusSlope - MinusSlope) / 2
  If WithYinter Then YintErr = Abs(PlusYint - MinusYint) / 2
End If

If WithYinter Then

  If InclErr Then
    RobReg = Array(Slope, SlopeErr, Yint, YintErr)
  Else
    RobReg = Array(Slope, Yint)
  End If

ElseIf InclErr Then
  RobReg = Array(Slope, SlopeErr)
Else
  RobReg = Slope
End If

Exit Function
1: On Error GoTo 0
RobReg = "#VALUE"
End Function

Function fdMSWDmult#(ByVal SmoothingWindow%, ByVal NumberOfPts%)

Dim rw1%, rw2%, colmn1%, c2%, wr1%, wr2%, wc1%, wc2%
Dim fr#, fC#, Rr1#, Rr2#, m#
Dim NumPts As Range, Windo As Range, MSWDexpected As Range

With ThisWorkbook.Sheets("Lowess")
  Set MSWDexpected = .[MSWDexpected]
  Set Windo = .[Window]
  Set NumPts = .[NumPts]
End With

SmoothingWindow = fvMinMax(SmoothingWindow, 5, 100)
NumberOfPts = fvMinMax(NumberOfPts, 5, 100)

wr2 = 1 ' find smallest window-entry >=SmoothingWindow
Do While Windo(wr2) < SmoothingWindow
  wr2 = 1 + wr2
Loop

wr1 = wr2 ' find largest window-entry <=SmoothingWindow
Do While Windo(wr1) > SmoothingWindow
  wr1 = wr1 - 1
Loop

wc2 = 1 ' find smallest window-entry >=NumberOfPts
Do While NumPts(wc2) < NumberOfPts
  wc2 = 1 + wc2
Loop

wc1 = wc2 ' find largest window-entry <=NumberOfPts
Do While NumPts(wc1) > NumberOfPts
  wc1 = wc1 - 1
Loop

rw1 = Windo(wr1):     rw2 = Windo(wr2)
colmn1 = NumPts(wc1): c2 = NumPts(wc2)

If rw1 = rw2 Then
  fr = 1
Else
  fr = (SmoothingWindow - rw1) / (rw2 - rw1)
End If

If colmn1 = c2 Then
  fC = 1
Else
  fC = (NumberOfPts - colmn1) / (c2 - colmn1)
End If

Rr1 = (1 - fr) * MSWDexpected(wr1, wc1) + fr * MSWDexpected(wr2, wc1)
Rr2 = (1 - fr) * MSWDexpected(wr1, wc2) + fr * MSWDexpected(wr2, wc2)
m = (1 - fC) * Rr1 + fC * Rr2
fdMSWDmult = 1 / m
End Function
