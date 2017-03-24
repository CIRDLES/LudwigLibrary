Attribute VB_Name = "Imain"
'ISOPLOT module Imain
Option Private Module
Option Explicit: Option Base 1

Sub Solve(ByVal N&, BadYork As Boolean, Np#(), xyProj#(), Failed As Boolean)
' Invoke various regressions or other major numeric tasks
Dim i&, j&, k&, ConcInt#(6), Xn#, Yn#, AgeErrWLE#
Dim TriplePass%, Inverse0 As Boolean, Normal0 As Boolean
Dim Isotype0%, CfInt#(2, 2), v#, Fpass%, Age#
Dim PbInt#(2), Resid#(), yf0 As Yorkfit, AgeErr#, cvX#
Dim cvY#, cvA#, DP() As DataPoints, xy#()
Dim cvB#, RhoAB#, RhoXY#, SlpEq#, SlpEqEr#, IntEq#
Dim IntEqEr#, Epsilon#, CovTW#, twSferr#, twIferr#
Dim CnvIferr#, SlopeLwr#, SlopeUppr#, InterLwr#, InterUppr#
Dim CnvSferr#, tSlp#, tYint#, Txb#, tSlpEr#, tYintErr#
Dim tProb#, tMSWD#, Uage#, Gamma0#, LwrAge#, UpprAge#
Dim LwrGamma0#, UpprGamma0#, UGrho#, DoErrs As Boolean
Dim xx#(), yy#(), xe#(), ye#(), Rr#()
ReDim DP(N), yfResid(N), yf.WtdResid(N), yf0.WtdResid(N)
Const Small = 0.00000001
AgeErrWLE = 0: SigYinit = 0
If Isotype = 14 And WtdAvXY Then
  Dim Xbar#, Ybar#, SumsXY#, ErrX#, ErrY#
  Dim EM#, MswdXY#, ProbEquiv#
  WtdXYmean InpDat(), N, Xbar, Ybar, SumsXY, ErrX, ErrY, RhoXY, Failed
  If Not Failed Then
    ShowXYwtdMean Xbar, ErrX, Ybar, ErrY, RhoXY, SumsXY, MswdXY, ProbEquiv, N, EM
    xyProj(0) = EM
    xyProj(1) = Xbar: xyProj(2) = EM * ErrX
    xyProj(3) = Ybar: xyProj(4) = EM * ErrY
    xyProj(5) = RhoXY
  End If
  Exit Sub
End If
For i = 1 To N
  With DP(i)
    If False And ArgonPlot And Inverse And Not Dim3 Then
      ' Transform inverse Ar-Ar data to normal before regression
      .X = InpDat(i, 1) / InpDat(i, 3) ' 39/36
      .y = 1 / InpDat(i, 3)            ' 40/36
      cvA = Abs(InpDat(i, 2) / InpDat(i, 1)) ' 39/40 fracterr
      cvB = Abs(InpDat(i, 4) / InpDat(i, 3)) ' 36/40 fracterr
      RhoAB = InpDat(i, 5)                   ' rho 39/40-36/40
      cvX = Sqr(cvA * cvA + cvB * cvB - 2 * cvA * cvB * RhoAB) ' 39/36 fracterr
      cvY = cvB ' 40/36 fracterr
      RhoXY = (cvB - cvA * RhoAB) / cvX ' 39/36-40/36 rho
      .Xerr = cvX * .X:  .Yerr = cvY * .y:   .RhoXY = RhoXY
    Else
      .X = InpDat(i, 1):       .Xerr = InpDat(i, 2)
      .y = InpDat(i, 3):       .Yerr = InpDat(i, 4)
      If Dim3 Then
        .z = InpDat(i, 5):     .Zerr = InpDat(i, 6)
        .RhoXY = InpDat(i, 7): .RhoXZ = InpDat(i, 8)
        .RhoYZ = InpDat(i, 9)
      Else
        .RhoXY = InpDat(i, 5)
        If ConcPlot And Not ConcAge And Inverse Then
          If .Xerr < 0.0001 And .Yerr < 0.001 Then
            If .X < 0.001 And .y > 0.5 And .y < 2 Then
              Anchored = True: PbAnchor = True: AnchorErr = 0.000001
              AnchorPt = i
              Exit For
            End If
          End If
        End If
      End If
    End If
  End With
Next i
If Dim3 Then
  Erase Crs
  If Linear3D Then
    If Robust Then
      If UseriesPlot Then
        RobustReg3 N, DP(), yf, Uage, Gamma0, UGrho
      Else
        'RobustReg3 N, DP(), yf
        'Show3dLine 0, 0
      End If
    Else
      Useries3Diso DP(), Np(), xyProj(), N, Failed
    End If
    Exit Sub
  Else
    PlanarFit N, DP(), Np(), Failed
    If Failed Then Exit Sub
  End If
Else
  Xn = DP(N).X: Yn = DP(N).y
  ' If a concordia plot with a forced intercept & nonzero assigned err
  '  in the forced intercept age, do 3 Yorkfit+concordia-intercept age solns.
  TriplePass = IIf(ConcPlot And AgeAnchor And AnchorErr <> 0, 2, 0)
End If
For Fpass = TriplePass To 0 Step -1
  If Not Dim3 Then
    With DP(N)
      If Fpass Then
        .X = ConcX(AnchorAge + (2 * Fpass - 3) * AnchorErr)
        .y = ConcY(AnchorAge + (2 * Fpass - 3) * AnchorErr)
      Else
        .X = Xn: .y = Yn
      End If
    End With
    If ConcAge Then
      ConcordiaAges xyProj(), N, Failed
      If Failed Then ConcAgePlot = False
      Exit Sub
    ElseIf Robust Then
      ReDim xy(N, 2)
      For i = 1 To N: With DP(i): xy(i, 1) = .X: xy(i, 2) = .y: End With: Next i
      With yf
        If N > 362 Then
          ' Doubly-resistant Siegel's algorithm, requires bootstrap errors but doesn't seem
          '  to perform better than RobustReg2
          RobustReg1 xy, .Slope, .Intercept
          If N > 6 Then
            If MsgBox("Calculate the regression-line errors?" & viv$ & "(may take a while)", vbYesNo, Iso) = vbYes Then
              .Ntrials = 2500
              BootRob xy, .Ntrials, .LwrSlope, .UpprSlope, .LwrInter, .UpprInter
            End If
          End If
        Else ' valid only for N<=362
          RobustReg2 xy(), .Slope, .LwrSlope, .UpprSlope, .Intercept, _
              .Xinter, .LwrInter, .UpprInter, .LwrXinter, .UpprXinter, , , True
        End If
        .SlopeError = (.UpprSlope - .LwrSlope) / 2
        .InterError = (.UpprInter - .LwrInter) / 2
        .XinterErr = (.UpprXinter - .LwrXinter) / 2
      End With
    Else
      If Fpass = TriplePass Then Call PointDispersion(N)
      York_Fit N, DP(), BadYork
      If False And ArgonPlot And Inverse And yf.Model = 1 And Not BadYork Then
        ReDim xx(N), yy(N), xe(N), ye(N), Rr(N)
        For i = 1 To N ' Swap x-y to determine symmetrical x-intercept
          xx(i) = InpDat(i, 3): xe(i) = InpDat(i, 4)
          yy(i) = InpDat(i, 1): ye(i) = InpDat(i, 2): Rr(i) = InpDat(i, 5)
        Next i
        ShortYork N, tSlp, tYint, Txb, tSlpEr, tYintErr, tProb, _
          False, False, tMSWD, xx(), xe(), yy(), ye(), Rr()
        Erase xx(), yy(), xe(), ye(), Rr()
        With yf
          If .Prob < MinProb Then tYintErr = tYintErr / 1.96 * Sqr(tMSWD) * StudentsT(N - 2)
          If ArgonPlot And Inverse And .Model = 1 Then _
            .Xinter = tYint: .XinterErr = tYintErr
        End With
      End If
    End If
    If BadYork Then Crs(23) = yf.Model: Exit Sub
  End If
  If ConcPlot And Not (Dim3 And Linear3D) Then
    Inverse0 = Inverse: Normal0 = Normal
    If Inverse Then  ' Convert to Conv. concordia ratios so can solve for negative ages
      Inverse = False: Normal = True: yf0 = yf
      With yf0 ' yf0 = T-W; yf = Conv.
        yf.Slope = 1 / .Intercept / Uratio
        yf.Intercept = -.Slope / .Intercept
        yf.Xbar = .Ybar / .Xbar * Uratio
        yf.Ybar = 1 / .Xbar
        twSferr = .SlopeError / .Slope
        twIferr = .InterError / .Intercept
        yf.SlopeError = Abs(twIferr * yf.Slope)
        CovTW = .RhoInterSlope * .SlopeError * .InterError
        yf.InterError = Sqr(SQ(twSferr) + SQ(twIferr) _
          - 2 * twSferr * twIferr * .RhoInterSlope) * Abs(yf.Intercept)
        CnvIferr = yf.InterError / yf.Intercept
        CnvSferr = yf.SlopeError / yf.Slope
        yf.RhoInterSlope = (SQ(twIferr) - twSferr * twIferr * .RhoInterSlope) / (CnvSferr * CnvIferr)
      End With
    End If
    ConcordiaIntercepts yf.Slope, yf.Intercept, ConcInt()
    Dim t1err#(2), t2err#(2), Bad(2) As Boolean
    ConcIntAgeErrors (ConcInt(1)), (ConcInt(2)), t1err(), t2err(), Bad()
    If Inverse0 Then
      Inverse = Inverse0: Normal = Normal0: yf = yf0
    End If
    If Fpass Then
      For k = 1 To 2: CfInt(Fpass, k) = ConcInt(k): Next k
    Else
      With yf
        If Dim3 Then
          ' Don't use old ConcInterErrors algorithm because unclear how to transform the appropriate Xbar.
          'If Inverse Then .Xbar = Crs(30)  ' ProjXYmin
          If Inverse Then .Xbar = yf0.Xbar
        Else
          For i = 1 To N: yfResid(i) = .WtdResid(i): Next i
          ' Find asymmetric 1st-deriv-exp. errors using error-hyperbola
          ConcInterErrors .Slope, .Intercept, .SlopeError, .InterError, .Xbar, ConcInt()
        End If
      End With
      If TriplePass Then
        If AnchorAge < ConcInt(2) Then ' Lower age forced
          ConcInt(3) = AnchorAge - AnchorErr: ConcInt(4) = AnchorAge + AnchorErr
          If ConcInt(3) = 0 Then ConcInt(3) = Small ' to avoid inf. errs inferred in
          If ConcInt(4) = 0 Then ConcInt(4) = Small ' SetupIsoRes.
          v = SQ(ConcInt(2) - ConcInt(5)) + SQ(ConcInt(2) - Min(CfInt(1, 2), CfInt(2, 2)))
          ' Upper Intercept age: (UpperInt-SigmaAnal)^2 + (UpperInt-SigmaAnchor)^2
          ConcInt(5) = ConcInt(2) - Sqr(v)
          v = SQ(ConcInt(2) - ConcInt(6)) + SQ(ConcInt(2) - Max(CfInt(1, 2), CfInt(2, 2)))
          ConcInt(6) = ConcInt(2) + Sqr(v)
        Else                           ' Upper age forced
          ConcInt(5) = AnchorAge - AnchorErr: ConcInt(6) = AnchorAge + AnchorErr
          If ConcInt(5) = 0 Then ConcInt(5) = Small
          If ConcInt(6) = 0 Then ConcInt(6) = Small
          v = SQ(ConcInt(1) - ConcInt(3)) + SQ(ConcInt(1) - Min(CfInt(1, 1), CfInt(2, 1)))
          ConcInt(3) = ConcInt(1) - Sqr(v)
          v = SQ(ConcInt(1) - ConcInt(4)) + SQ(ConcInt(1) - Max(CfInt(1, 1), CfInt(2, 1)))
          ConcInt(4) = ConcInt(1) + Sqr(v)
        End If
      End If
    End If
    For i = 1 To 6
      If ConcInt(i) = BadT Then ConcInt(i) = 0
    Next i
    If Abs(ConcInt(1) - ConcInt(2)) < 0.001 Then
      ConcInt(1) = 0:   ConcInt(3) = 0:    ConcInt(4) = 0
    End If
  ElseIf (Not OtherXY And Not StackedUseries) Or (Not Dim3 And iLambda(OtherIndx) <> 0) Then
    If Not PbPlot Or PbType <> 2 Then
      IsochronAges Age, AgeErr, AgeErrWLE, PbInt(), SlpEq, SlpEqEr, IntEq, IntEqEr, Epsilon
    End If
  End If
  If Fpass = 0 Then
    If Not Dim3 Then Erase Crs
    With yf
      Crs(1) = .Slope:       Crs(2) = .SlopeError
      Crs(3) = .Intercept:   Crs(4) = .InterError
      Crs(5) = .Xbar:        Crs(6) = .MSWD
      Crs(7) = .Prob:
      Crs(16) = .ErrSlApr:   Crs(17) = .ErrIntApr
      Crs(18) = .ErrSlincSc: Crs(19) = .ErrIntincSc
      Crs(20) = .Ybar:       Crs(21) = .Xinter
      Crs(22) = .XinterErr:  Crs(23) = .Model
    End With
    For i = 8 To 15: Crs(i) = 0: Next i
    If ConcPlot Then
      Crs(8) = ConcInt(1)  ' Lower intercept age
      Crs(9) = ConcInt(2)  ' Upper intercept age
      Crs(12) = ConcInt(3) ' Lower limit of lower-intercept age
      Crs(13) = ConcInt(4) ' Upper   "    "   "      "       "
      Crs(14) = ConcInt(5) ' Lower limit of upper-intercept age
      Crs(15) = ConcInt(6) ' Upper   "    "   "      "       "
      If ConcInt(3) <> 0 And ConcInt(4) <> 0 Then
        Crs(10) = (ConcInt(4) - ConcInt(3)) / 2 ' Average error on lower intercept
      End If
      If ConcInt(5) <> 0 And ConcInt(6) <> 0 Then
        Crs(11) = (ConcInt(6) - ConcInt(5)) / 2 ' Average error on upper intercept
      End If
    Else
      Crs(8) = yf.Intercept:  Crs(10) = yf.InterError
      Crs(9) = Age:           Crs(11) = AgeErr
      Crs(12) = 2 * SigYinit: Crs(15) = yf.RhoInterSlope
      If PbPlot And PbType = 1 Then '~!
        Crs(13) = PbInt(1): Crs(14) = PbInt(2)
        If AgeErrWLE > 0 Then Crs(26) = AgeErrWLE
      End If
      If IntEq <> 0 Then Crs(36) = IntEq: Crs(37) = IntEqEr
      If SmNdIso Then Crs(34) = Epsilon
    End If
    If ConcPlot And Not (Dim3 And Linear3D) And Not Bad(1) Then
      ' Define errors on intercept ages both with & without lambda errors
      If Not Anchored Or AnchorErr = 0 Then
        Crs(38) = t1err(1): Crs(39) = t2err(1)
        If Not Bad(2) Then Crs(40) = t1err(2): Crs(41) = t2err(2)
      Else ' Anchored chord with nonzero error on the anchor age
        Crs(38) = Crs(10): Crs(39) = Crs(11)
        If Not Bad(2) Then ' Quadratically put in error due solely to lambda error
          Crs(40) = Sqr(SQ(Crs(10)) + SQ(t1err(2)) - SQ(t1err(1)))
          Crs(41) = Sqr(SQ(Crs(11)) + SQ(t2err(2)) - SQ(t2err(1)))
        End If
      End If
    End If
  End If
Next Fpass
If Dim3 And Planar3D Then   ' Calculate Pb-Pb ages of common-Pb plane
  Isotype0 = Isotype  '~!
  If Inverse Then
    Inverse0 = Inverse: Inverse = False: Normal = True: yf0 = yf
    Isotype = 8: PbPlot = True: PbType = 1
    With yf
      .Slope = yf0.Intercept: .SlopeError = yf0.InterError
      .Intercept = Crs(24):  .InterError = Crs(25)
    End With
  End If
  If Not OtherXY Then
    IsochronAges Age, AgeErr, AgeErrWLE, PbInt()
    If Inverse0 Then
      Inverse = Inverse0: Normal = False:  yf = yf0
    End If
    Crs(32) = Age:      Crs(33) = AgeErr
    Crs(34) = PbInt(1): Crs(35) = PbInt(2)
  End If
  Isotype = Isotype0
  PlotIdentify
End If
End Sub

Sub York_Fit(ByVal N&, DP() As DataPoints, BadYork As Boolean, Optional Model_1_only = False)
Dim i&, Iter%, nU&, m3Count%, NextM%
Dim xs#, YS#, u#, v#, c#, uU#, Tpr#
Dim d#, e#, WeightSum#, WtSqr#, test#, vv#
Dim uv#, Sums#, SumX2Z#, SumU2Z#, SumZ#, SumXZ#
Dim CovInterSlope#, TrueXval#, Students#, VarSlApr#, VarIntApr#
Dim tmp#, Tslope#, Slope0#, Titt As Boolean, NM$
Dim Mswd1#, ErrSlApr1#, ErrIntApr1#, Numer#, Denom#
Dim Xwt#(), Ywt#(), MeanWt#()
Dim Weight#(), yRho#(), X#(), y#(), txy#()
ReDim Xwt(N), Ywt(N), MeanWt(N), yf.WtdResid(N)
ReDim yfResid(N), Weight(N), yRho(N), X(N), y(N), txy(N, 2)
Const Toler = 0.00001, MaxIter = 100, m3iterMax = 50
' 09/06/20 -- change Toler from 0.000001 to 0.00001
If MinProb = 0 Then MinProb = Val(Menus("MinProb"))
ViM Model_1_only, False
BadYork = True: yf.Model = 1
Titt = True  ' Use Titterington error-algorithm
For i = 1 To N
  With DP(i)
    X(i) = .X:      y(i) = .y
    txy(i, 1) = .X: txy(i, 2) = .y
  End With
Next i
Select Case N ' Initial estimate of slope
  Case 2
    If X(2) = X(1) Then MsgBox "Duplicate x-values: cannot regress.": BadYork = True: Exit Sub
    Tslope = (y(2) - y(1)) / (X(2) - X(1))
  Case Is < 363
    RobustReg2 txy(), Tslope, SlopeOnly:=True
  Case Else
    RobustReg1 txy(), Tslope, SlopeOnly:=True
End Select
Erase txy
NextModel:
With yf
  If .Model = 3 Then
    m3Count = 1 + m3Count
    If m3Count > m3iterMax Then Exit Sub
    Slope0 = .Slope
    If m3Count = 1 Then
      SigYinit = .ErrIntincSc
    Else
      SigYinit = SigYinit * Sqr(.MSWD)
    End If
  End If
End With
Iter = 0
Do
  yf.Slope = Slope0
  Iter = 1 + Iter
  If Iter = 1 Then
    For i = 1 To N
      With DP(i)
        If yf.Model = 1 Then
          Xwt(i) = 1 / SQ(.Xerr): Ywt(i) = 1 / SQ(.Yerr)
          yRho(i) = .RhoXY
        ElseIf yf.Model = 2 Then ' Equally-wtd pts
          Xwt(i) = 1: Ywt(i) = 1 / SQ(yf.Slope): yRho(i) = 0
        ElseIf yf.Model = 3 Then
          ' Model-3 weighting: analytical errors plus a normally-distrib-
          '  uted  variation in the initial-ratio.  The initial-ratio
          '  variation is unknown & must be estimated by the algorithm
          tmp = SQ(.Yerr)
          yRho(i) = .RhoXY * Sqr(tmp / (tmp + SQ(SigYinit)))
          Ywt(i) = 1 / (tmp + SQ(SigYinit))
        End If
        MeanWt(i) = Sqr(Xwt(i) * Ywt(i))
      End With
    Next i
  End If
  With yf
    For i = 1 To N
      Weight(i) = Xwt(i) * Ywt(i) / _
        (SQ(.Slope) * Ywt(i) + Xwt(i) - 2 * .Slope * yRho(i) * MeanWt(i))
    Next i
    WeightSum = Sum(Weight())
    xs = SumProduct(Weight(), X())
    YS = SumProduct(Weight(), y())
    .Xbar = xs / WeightSum: .Ybar = YS / WeightSum
    c = 0: d = 0: e = 0
    For i = 1 To N
      WtSqr = SQ(Weight(i))
      u = DP(i).X - .Xbar: v = DP(i).y - .Ybar
      uU = u * u: vv = v * v: uv = u * v
      c = c + (uU / Ywt(i) - vv / Xwt(i)) * WtSqr
      d = d + (uv / Xwt(i) - yRho(i) * uU / MeanWt(i)) * WtSqr
      e = e + (uv / Ywt(i) - yRho(i) * vv / MeanWt(i)) * WtSqr
    Next i
  End With
  ' Test for SQRT of neg# (usually from roundoff error?)
  test = c * c + 4 * d * e
  If test < 0 Or d = 0 Then Stop 'Exit Sub
  Slope0 = (Sqr(test) - c) / (2 * d)
  If Slope0 = 0 Then Slope0 = 1E-30
  Iter = 1 + Iter
  If Iter = MaxIter Then
    Exit Sub
  End If
Loop Until Abs((Slope0 - yf.Slope) / Slope0) < Toler
With yf
  .Intercept = .Ybar - .Slope * .Xbar
  Sums = 0
  For i = 1 To N
    yfResid(i) = .Intercept + .Slope * DP(i).X - DP(i).y ' Unwtd Y-resids
    .WtdResid(i) = Sqr(Weight(i)) * yfResid(i)           ' Wtd Y-yfResid
    Sums = Sums + SQ(.WtdResid(i))                       ' Sum of squares of wtd Y-resids)
  Next i
  If Titt Then ' Titterington algorithm
    SumX2Z = 0:  SumXZ = 0
    For i = 1 To N
      Numer = Weight(i) * yfResid(i) * (yRho(i) * MeanWt(i) - Ywt(i) * .Slope)
      TrueXval = DP(i).X + Numer / SQ(MeanWt(i))
      SumXZ = SumXZ + TrueXval * Weight(i)
      SumX2Z = SumX2Z + TrueXval * TrueXval * Weight(i)
    Next i
    Denom = SumX2Z * WeightSum - SumXZ * SumXZ
    VarSlApr = WeightSum / Denom
    VarIntApr = SumX2Z / Denom
    CovInterSlope = -SumXZ / Denom
    .RhoInterSlope = CovInterSlope / Sqr(VarSlApr * VarIntApr)
  Else        ' York algorithm
    SumX2Z = 0: SumU2Z = 0
    For i = 1 To N
      SumX2Z = SumX2Z + SQ(DP(i).X) * Weight(i)
      SumU2Z = SumU2Z + SQ(DP(i).X - .Xbar) * Weight(i)
    Next i
    VarSlApr = 1 / SumU2Z
    VarIntApr = SumX2Z / (WeightSum * SumU2Z)
  End If
  Dim SlpErr#(6), YintErr#(6), Xinterr#(6)
  'MCyorkfit N, 10000, SlpErr(), YintErr(), XintErr()
  .ErrSlApr = Sqr(VarSlApr): .ErrIntApr = Sqr(VarIntApr)
  If .Model = 1 Then
    ErrSlApr1 = .ErrSlApr: ErrIntApr1 = .ErrIntApr
  End If
  nU = N - 2
  If nU Then
    .MSWD = Sums / nU
    If .Model = 1 Then .Prob = ChiSquare(.MSWD, nU)
    Students = StudentsT(nU)
  Else
    .MSWD = 0: .Prob = 1
  End If
  If .Model = 1 Then Mswd1 = .MSWD
  .ErrSlincSc = .ErrSlApr * Sqr(.MSWD)   ' 1-sigma "incl scatter" slope-err
  .ErrIntincSc = .ErrIntApr * Sqr(.MSWD) ' 1-sigma "incl scatter" inter-err
  If .Prob > MinProb And .Model = 1 Then ' If prob of scatter from assigned
    .Emult = 1.96                        '  errs>MinProb, use a priori errs.
    .SlopeError = .Emult * .ErrSlApr
    .InterError = .Emult * .ErrIntApr
  Else   ' If <MinProb, multiply by SQRT(MSWD) & Student's-T for N-2 d.f.
    .Emult = Students * Sqr(.MSWD)
    .SlopeError = .Emult * .ErrSlApr  'Students * .ErrSlincSc
    .InterError = .Emult * .ErrIntApr 'Students * .ErrIntincSc
  End If
  If Not Model_1_only Then
    If .Model = 1 Then
      If .Prob <= MinProb And Not Anchored Then
        NextM = 3 + (ConcPlot Or ArgonPlot Or PbPlot Or UseriesPlot Or OtherXY)
        Tpr = Prnd(.Prob, -4)
        NM$ = "Probability of fit is "
        If Tpr > 0 Then NM$ = NM$ & "only "
        NM$ = NM$ & sn$(Tpr) & " - Do a Model-" & sn$(NextM) & " Fit?" & vbLf
        If NextM = 2 Then
          NM$ = NM$ & vbLf & "(Model 2 weights the points equally, ignoring the data-point errors)"
        ElseIf NextM = 3 Then
          NM$ = NM$ & "(Model 3 assumes the excess scatter results from" & vbLf & _
            "  a uniform error in the Y parameter, but also" & vbLf & _
            "  takes data-point errors into account)"
        End If
        If .Prob < 0.05 And ConcPlot Then NM$ = NM$ & viv$ & vbLf & "NOTE: " _
          & "Age-errors from low-probability U-Pb discordia may be unreliable."
        i = MsgBox(NM$, vbYesNoCancel + vbQuestion, Iso)
        If i = vbCancel Then
          ExitIsoplot
        ElseIf i = vbYes Then
          .Model = NextM
          If Iter > 100 Then .Slope = 0.1
          GoTo NextModel
        End If
      End If
    ElseIf .Model = 3 Then
      If m3Count > m3iterMax Then Exit Sub
      If Abs(.MSWD - 1) > 0.001 Then GoTo NextModel
    End If
  End If
  Xintercept .Xinter, .XinterErr, .Intercept, .InterError, .Slope, .SlopeError, .Xbar
  If .Model > 1 Then
    .ErrSlApr = ErrSlApr1:  .ErrIntApr = ErrIntApr1
    .MSWD = Mswd1
  End If
End With
BadYork = False
End Sub

Sub Xintercept(Xinter#, XinterErr#, ByVal Yinter#, ByVal YinterErr#, _
   ByVal Slope#, ByVal SlopeErr#, ByVal Xbar#)
Dim A#, b#, c#, q#, root1#, root2#, discr#
Xinter = -Yinter / Slope
A = SQ(Slope) - SQ(SlopeErr)
b = 2 * (Slope * Yinter + SQ(SlopeErr) * Xbar)
c = SQ(Yinter) - SQ(YinterErr)
discr = b * b - 4 * A * c
XinterErr = 0
If discr >= 0 Then
  q = -(b + Sgn(b) * Sqr(discr)) / 2
  If A <> 0 Then
    root1 = q / A
    If q <> 0 Then
      root2 = c / q
      XinterErr = Abs(root2 - root1) / 2
    End If
  End If
End If
End Sub

Sub Swap(A, b)
Dim c
c = A: A = b: b = c
End Sub

Sub IsochronAges(Age#, AgeErr#, AgeErrWLE#, PbInt#(), _
  Optional SlpEq#, Optional SlpEqEr#, Optional IntEq#, _
  Optional IntEqEr#, Optional Epsilon_#, Optional NoJay As Boolean = False)
' Calculate the isochron age & age-error for a regression line, &
'   display/print-out these values.  If a Pb-Pb plot, also calculate &
'   show the growth-curve .intercept-ages of the regression line.  If an Sm-Nd
'   plot, calculate & show the epsilon-CHUR value for the isochron.
Dim u#, v#, UpprIntEq#, LwrIntEq#, UpprSlpEq#, LwrSlpEq#
Dim K40BetaEcRatio#, test#, r0#, SL0%
Dim CHUR#, Chur0#, ChurSmNd#, Pb76#, Pb76err#, s$
Dim SourceR#(2), SourcePD#(2), a40_k40#, a40_a36#
Dim a40_k40err#, a40_a36err#
ViM NoJay, False
AgeErrWLE = 0: SlpEq = 0: IntEq = 0
With yf
  If Not ConcPlot And Not PbPlot And Not (Dim3 And UseriesPlot) Then  '~!
    ' A normal Rb/Sr-type isochron
    If 1 + .Slope < 3E+38 Then
      If KCaIso Then
        Do
          Do
            s$ = InputBox("Beta/Electron-Capture ratio for K-40?")
            If Len(s$) = 0 Then KwikEnd
          Loop Until IsNumeric(s$)
          K40BetaEcRatio = Val(s$)
        Loop Until K40BetaEcRatio > 0
        r0 = K40BetaEcRatio / (1 + K40BetaEcRatio)  ' K-Ca: L(Beta)/L(Tot)
        'CASE xx
        '  r0 = 1 / (1 + K40BetaEcRatio)            ' K-Ar: L(Ec)/L(Tot)
      ElseIf ClassicalIso Or UThPbIso Or OtherXY Then
        r0 = 1
      ElseIf ArgonPlot Then
        If Not NoJay Then
          ArgonJ
          If Jay = 0 Then Exit Sub
        End If
        If Dim3 Then
          Select Case ArType
            Case 1
              a40_a36 = Crs(24): a40_a36err = Crs(25)
              a40_k40 = .Slope: a40_k40err = .SlopeError ' at 2-sigma/95%conf
            Case 2
              a40_a36 = 1 / Crs(24)
              a40_a36err = Crs(25) / Crs(24) * a40_a36
              a40_k40 = -.Slope / Crs(24)
            Case 3
              a40_a36 = 1 / .Slope: a40_a36err = .SlopeError / (.Slope * .Slope)
              a40_k40 = -Crs(24) / .Slope
          End Select
          If ArType > 1 Then
             u = SQ(.SlopeError / .Slope) + SQ(Crs(25) / Crs(24))
             v = 2 * .SlopeError / .Slope * .SlopeError / .Slope * Crs(26)
             a40_k40err = Sqr(u - v) * a40_k40
          End If
        ElseIf Normal Then
          a40_a36 = .Intercept: a40_a36err = .InterError
          a40_k40 = .Slope:     a40_k40err = .SlopeError
        Else
          a40_a36 = 1 / .Intercept: a40_a36err = .InterError / .Intercept * a40_a36
          a40_k40 = 1 / .Xinter:   a40_k40err = .XinterErr / .Xinter * a40_k40
        End If
        ArgonAge a40_k40, a40_k40err, Age, AgeErr ' Age error at 2-sigma/95%
      End If
      If Not ArgonPlot And Not UseriesPlot Then
        If Normal Then ' eg 206/204 vs 238/204
          SlpEq = .Slope: IntEq = .Intercept
          If Robust Then
            UpprSlpEq = .UpprSlope: LwrSlpEq = .LwrSlope
            UpprIntEq = .UpprInter: LwrIntEq = .LwrInter
          Else
            SlpEqEr = .SlopeError: IntEqEr = .InterError
          End If
        Else   ' Must be Isotype of 10 to 12; Inverse U/Pb or Th/Pb (robust sol'n inhibited)
          SlpEq = 1 / .Intercept: IntEq = 1 / .Xinter
          SlpEqEr = .InterError / SQ(.Intercept)
          IntEqEr = .XinterErr / SQ(.Xinter)
          If Robust Then
            UpprSlpEq = 1 / .LwrInter: LwrSlpEq = 1 / .UpprInter
            UpprIntEq = 1 / .LwrXinter: LwrIntEq = 1 / .UpprXinter
          End If
        End If
        test = 1 + SlpEq / r0
        If test > MINLOG And test < MAXLOG And iLambda(Isotype) > 0 Then
          ' Std-isochron age
          Age = Log(test) / iLambda(Isotype)
          If Robust Then
            .LwrAge = 0: .UpprAge = 0
            On Error Resume Next
            .LwrAge = Log(1 + LwrSlpEq / r0) / iLambda(Isotype)
            .UpprAge = Log(1 + UpprSlpEq / r0) / iLambda(Isotype)
            On Error GoTo 0
          Else
            AgeErr = Abs(SlpEqEr / (iLambda(Isotype) * test)) ' at 2-sigma/95%conf
          End If
          If SmNdIso Then Epsilon_ = Epsilon(Age, .Intercept)
        End If
      End If
    End If
  ElseIf PbPlot And PbType = 1 Then
    If Normal Then
      Pb76 = .Slope
      If Robust Then
        UpprSlpEq = .UpprSlope: LwrSlpEq = .LwrSlope
      Else
        Pb76err = .SlopeError
      End If
    Else
      Pb76 = .Intercept
      If Robust Then
        UpprSlpEq = .UpprInter: LwrSlpEq = .LwrInter
      Else
        Pb76err = .InterError
      End If
    End If
    Age = PbPbAge(Pb76)
    If Robust Then
      .UpprAge = Age: .LwrAge = Age
      If Normal Then
        UpprSlpEq = .UpprSlope: LwrSlpEq = .LwrSlope
      Else
        UpprSlpEq = .UpprInter: LwrSlpEq = .LwrInter
      End If
      On Error Resume Next
      .UpprAge = PbPbAge(UpprSlpEq): .LwrAge = PbPbAge(LwrSlpEq)
      On Error GoTo 0
    Else
      AgeErr = PbPbAge(Pb76, 0, Age, Pb76err, False)
      If Lambda235err > 0 And Lambda238err > 0 Then
        SL0 = SigLev: SigLev = 2
        AgeErrWLE = PbPbAge(Pb76, 0, Age, Pb76err, True)
        SigLev = SL0
      End If
    End If
    GrowthInters .Slope, .Intercept, PbInt()
  Else
    If Age = BadT Or Age = 0 Then
      MsgBox "No age-solution for this isochron", , Iso
      Regress = False
    End If
  End If
End With
End Sub

Function tSt(ByVal v) As String ' return a # or numeric string as a space-trimmed string
Dim vv$
If IsNumeric(v) Then vv$ = Str(v) Else vv$ = v
tSt = Trim(vv$)
End Function

Sub GetOpSys()
Dim i%, d$, u As Boolean, X$, DS As Object, tB As Object
qq = Chr(34): Sqrt = Chr(214): viv$ = vbLf & vbLf: Dsep = "."
Mac = (InStr(OpSys, "Macintosh") > 0)
ExcelVersion = Version(True)
XL2007 = (ExcelVersion >= 12)
If XL2007 Then
  MsgBox "Sorry, this and earlier versions of Isoplot are not compatible with Excel 2007.", , "Isoplot"
  KwikEnd
End If
MacExcelX = (Mac And Int(ExcelVersion) >= 10)
Windows = Not Mac: ShapesOK = True
'Mac = True: Windows = False: MacExcelX = False
If Windows And ExcelVersion >= 10 Then
  With App
    d$ = .DecimalSeparator
    X$ = .International(xlDecimalSeparator)
    u = .UseSystemSeparators
  End With
  Dsep = X$
  If False And d$ <> "." Or X$ <> "." Then
    MsgBox "ISOPLOT requires that the Windows decimal separator be " _
      & qq & "." & qq & " to function reliably." & viv$ & _
      "To do this, open 'Regional and Language Options' from the Control" _
      & " Panel and select English.", , Iso
  End If
End If
On Error Resume Next
'App.Calculation = xlcalculationmanual
App.DisplayStatusBar = True
Set IsoPlotTypes = Menus("IsoTypes")
For i = 1 To IsoPlotTypes.Rows.Count
  IsoPlotTypes(i, 1) = IsoPlotTypes(i, 3 - Mac)
Next i
If Windows Then
  For Each DS In ThisWorkbook.DialogSheets
    For Each tB In DS.TextBoxes
      tB.Visible = True
    Next tB
  Next DS
End If
End Sub

Function IsRangeVal(ByVal T$) As Double
On Error GoTo NotArange
IsRangeVal = Range(T$).Value: Exit Function
NotArange: On Error GoTo NoVal
IsRangeVal = Val(T$): Exit Function
NoVal: IsRangeVal = 0
End Function
