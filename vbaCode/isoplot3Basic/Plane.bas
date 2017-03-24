Attribute VB_Name = "Plane"
'Isoplot module Plane
Option Private Module
Option Explicit: Option Base 1

Private Sub SimplePlaneFit(xv#(), Yv#(), Zv#(), _
  ByVal N&, Coef#(), Failed As Boolean)
' Calculate sol'n to the planar equation z = a + bx + cy
Dim Obs#(), z#()
ReDim Obs(N, 3), z(N, 1)
Dim ObsT As Variant, InvObsTObs As Variant, ObsTObs As Variant, ObsTZ As Variant
Dim i&, Param As Variant
For i = 1 To N
  z(i, 1) = Zv(i):   Obs(i, 3) = 1
  Obs(i, 1) = xv(i): Obs(i, 2) = Yv(i)
Next i
With App
  ObsT = .Transpose(Obs)
  ObsTObs = .MMult(ObsT, Obs)
  InvObsTObs = .MInverse(ObsTObs)
  Failed = IsError(InvObsTObs)
  If Failed Then Exit Sub
  ObsTZ = .MMult(ObsT, z)
  Param = .MMult(InvObsTObs, ObsTZ)
  For i = 1 To 3: Coef(i) = Param(i, 1): Next i
End With  ' NOTE: a = Coef(3), b = Coef(1), c = Coef(2)
End Sub

Private Sub KentFit(ByVal N&, DP() As DataPoints, Param#(), _
  Resid#(), MSWD#, Failed As Boolean)
' Best-Fit plane using n-dimensional maximum-likelihood approach &
'   algorithms of Kent, Watson, & Onstott, EPSL v. 97 1990, p. 1-17.
' Equation is  Yi(4)= Alpha + Beta*Yi(2) + Gamma*Yi(3)
'              Z    = Alpha + Beta*X     + Gamma*Y
' To fit the eqn. Y = a*X + b + c*Y, must swap the Z & Y parameters,
'   errors, & error-correlations.
Dim i&, j&, k&, P%, Count&
Dim bi#(4, 4), Yi#(4, 1), Delta#(3, 1), Gamma#(4, 1)
Dim RegrErrors As Variant, Partial#(3, 1), Partial2#(3, 3), Sqr3#(4, 4)
Dim GammaT As Variant, Numer As Variant, Denom As Variant, YiT As Variant ', DeltaT As Variant
Dim BiGamma As Variant, BiGammaT As Variant, GammaYiT As Variant
Dim Coef#(3), SB$, CovAB#, CovAC#, CovBC#, Sums#
Dim SumWeight#, Term#, Resid2#, Epsilon As Variant
Dim SumWtdX#, SumWtdY#, SumWtdZ#, Ct1&, Xbar#
Dim Ybar#, Zbar#, TimeIn#, SumEps#, Scalar1#
Dim Scalar2#, Scalar3#, Scalar4#, Scalar5#
Dim Sqr0 As Variant, Sqr1 As Variant, Sqr2 As Variant, Sqr4 As Variant
Dim Weight#(), zT#(), GammaTYi#(), GammaTBiGamma#()
Dim sigmaX#(), SigmaY#(), SigmaZ#()
Dim Xt#(), Yt#(), Rho#()
ReDim Weight(N), zT(N), GammaTYi(N), GammaTBiGamma(N), sigmaX(N), SigmaY(N), SigmaZ(N)
ReDim Xt(N), Yt(N), zT(N), Rho(3, N)
Const MaxCount = 20, MaxSumEps = 0.000000001, MaxTime = 12
SB$ = "regressing, iteration": Sums = 0: Count = 0: TimeIn = Timer()
' Must swap Y- & Z-values, errs, & rhos to fit Kent et al's equation
'   Y = aX + b + cZ.
For i = 1 To N
  With DP(i)
    Xt(i) = .X: Yt(i) = .z: zT(i) = .y
    sigmaX(i) = .Xerr
    SigmaY(i) = .Zerr  'Swap Y-Z
    SigmaZ(i) = .Yerr  '  "   "
    Rho(1, i) = .RhoXZ '  "   "
    Rho(2, i) = .RhoXY '  "   "
    Rho(3, i) = .RhoYZ
  End With
Next i
' Calculate first estimate of Delta with simple linear regression
SimplePlaneFit Xt(), Yt(), zT(), N, Coef(), Failed
If Failed Then GoTo FailedFit
Delta(1, 1) = Coef(3) ' = Gamma
Delta(2, 1) = Coef(1) ' = Beta
Delta(3, 1) = Coef(2) ' = Alpha
Do
  Ct1 = 1 + Count
  If Ct1 Mod 5 = 0 Then StatBar SB$ & Str(Ct1)
  If Count Mod 10 = 0 Then TooLongCheck TimeIn, MaxTime
  'DeltaT = app.Transpose(Delta)
  Gamma(1, 1) = Delta(1, 1)
  Gamma(2, 1) = Delta(2, 1)
  Gamma(3, 1) = Delta(3, 1)
  Gamma(4, 1) = -1#
  GammaT = App.Transpose(Gamma)
  For j = 1 To 3             ' Partial() is partial-L/partial-Delta
    Partial(j, 1) = 0
  Next j
  For i = 1 To N             ' Calculate Partial()
    For j = 1 To 4           ' /
      Yi(1, 1) = 1#          ' |
      Yi(2, 1) = DP(i).X     ' |  Transfer data into 1-point arrays
      Yi(3, 1) = Yt(i)       ' |
      Yi(4, 1) = zT(i)       ' |
    Next j                   ' \
    For j = 1 To 4           ' /
      bi(1, j) = 0           ' |
      bi(j, 1) = 0           ' |
    Next j                   ' |
    bi(2, 2) = SQ(sigmaX(i))  ' |
    bi(3, 3) = SQ(SigmaY(i))  ' | Variance-Covariance matrix for data-point
    bi(4, 4) = SQ(SigmaZ(i)) ' |
    ' Rho(1,i)=Rho(X,Y); Rho(2,i)=Rho(X,Z); Rhoxy(3,i)=Rho(Y,Z)
    bi(2, 3) = sigmaX(i) * SigmaY(i) * Rho(1, i)
    bi(3, 2) = bi(2, 3)       ' |
    bi(2, 4) = sigmaX(i) * SigmaZ(i) * Rho(2, i)
    bi(4, 2) = bi(2, 4)       ' |
    bi(3, 4) = SigmaY(i) * SigmaZ(i) * Rho(3, i)
    bi(4, 3) = bi(3, 4)       ' \
    ' Calc. terms in the -PartialLp/PartialDelta sum
    YiT = App.Transpose(Yi)
    Numer = App.MMult(GammaT, Yi)
    GammaTYi(i) = Numer(1) '(1, 1)
    BiGamma = App.MMult(bi, Gamma)
    Denom = App.MMult(GammaT, BiGamma)
    GammaTBiGamma(i) = Denom(1) '(1, 1)   ' Save for later
    Scalar1 = Numer(1) / Denom(1) 'Numer(1, 1) / Denom(1, 1)
    Scalar2 = SQ(Numer(1) / Denom(1))   ' Note error in Kent et al.
    For j = 1 To 3
       ' Note error in Kent et al.
       Partial(j, 1) = Partial(j, 1) - Scalar1 * Yi(j, 1) + Scalar2 _
         * BiGamma(j, 1) ' krl
    Next j
  Next i
  ' Partial2() is minus (partial^2-L)/(partial-Delta*partial-DeltaT)
  For i = 1 To 3
    For j = 1 To 3: Partial2(i, j) = 0: Next j, i
  For i = 1 To N              ' Calculate Partial2()
     For j = 2 To 4           '   /
       Yi(2, 1) = DP(i).X     '   |
       Yi(3, 1) = Yt(i)       '   |
       Yi(4, 1) = zT(i)       '   |
     Next j                   '   |
     bi(2, 2) = SQ(sigmaX(i))  '   | Need to fill data-point arrays again
     bi(3, 3) = SQ(SigmaY(i))  '   |
     bi(4, 4) = SQ(SigmaZ(i)) '   |
     bi(2, 3) = sigmaX(i) * SigmaY(i) * Rho(1, i)
     bi(3, 2) = bi(2, 3)      '   |
     bi(2, 4) = sigmaX(i) * SigmaZ(i) * Rho(2, i)
     bi(4, 2) = bi(2, 4)      '   |
     bi(3, 4) = SigmaY(i) * SigmaZ(i) * Rho(3, i)
     bi(4, 3) = bi(3, 4)      '   \
     YiT = App.Transpose(Yi)
     Sqr0 = App.MMult(Yi, YiT)
     Scalar3 = GammaTYi(i) / SQ(GammaTBiGamma(i)) ' Note error in Kent et al
     BiGamma = App.MMult(bi, Gamma)
     BiGammaT = App.Transpose(BiGamma)
     Sqr1 = App.MMult(Yi, BiGammaT)
     YiT = App.Transpose(Yi)
     GammaYiT = App.MMult(Gamma, YiT)
     Sqr2 = App.MMult(bi, GammaYiT)
     For j = 1 To 4
       For k = 1 To 4
         Sqr3(j, k) = Sqr1(j, k) + Sqr2(j, k)
     Next k, j
     Scalar4 = SQ(GammaTYi(i)) / GammaTBiGamma(i) ^ 3
     Sqr4 = App.MMult(BiGamma, BiGammaT)
     Scalar5 = SQ(GammaTYi(i) / GammaTBiGamma(i))
    For j = 1 To 3               ' Do sums
      For k = 1 To 3
        Term = Sqr0(j, k) / GammaTBiGamma(i)
        Term = Term - 2 * Scalar3 * Sqr3(j, k) + 4 * Scalar4 * Sqr4(j, k)
        Term = Term - Scalar5 * bi(j, k)
        Partial2(j, k) = Partial2(j, k) + Term
    Next k, j
  Next i
  ' The inverse of Partial2() is RegrErrors(), which contains the variances
  '   & covariances of the regression parameters.
  RegrErrors = App.MInverse(Partial2)
  If IsError(RegrErrors) Then GoTo FailedFit
  ' Epsilon is the estimated change in the regression parameters to be added
  '   (to Delta()) to improve the estimate of the parameters.
  Epsilon = App.MMult(RegrErrors, Partial)
  SumEps = 0
  Count = 1 + Count
  For i = 1 To 3
    Delta(i, 1) = Delta(i, 1) + Epsilon(i, 1)
    SumEps = SumEps + Abs(Epsilon(i, 1) / Delta(i, 1))
  Next i
Loop Until SumEps < MaxSumEps Or Count > MaxCount
If Count > MaxCount Then GoTo FailedFit
SumWtdX = 0:  Sums = 0
' GammaTYi(i) are the observed residuals;
' GammaTBiGamma(i) are the squares of the predicted residuals.
For i = 1 To N
  Resid(i) = GammaTYi(i) / Sqr(GammaTBiGamma(i)) ' Wtd residual
  Resid2 = SQ(Resid(i))                     ' Square of wtd residual
  Sums = Sums + Resid2                      ' Sums of squares of weighted residuals
  Weight(i) = 1 / Sqr(GammaTBiGamma(i))
Next i
For i = 1 To N
   SumWtdX = SumWtdX + DP(i).X * Weight(i)
Next i
SumWeight = Sum(Weight())
SumWtdY = SumProduct(Yt(), Weight())
SumWtdZ = SumProduct(zT(), Weight())
Xbar = SumWtdX / SumWeight
Ybar = SumWtdZ / SumWeight  ' remember Y-Z swapped
Zbar = SumWtdY / SumWeight  '    "           "
' Param matrix:
' Y = aX + b + cZ
' 1,1  2,1  3,1  are values for app, b, & c
' 1,2  2,2  3,2   "    "     "  sigma-app, sigma-b, & sigma-c
' 1,3  2,3  3,3   "    "     "  rho(ab), rho(ac), & rho(bc)
' Errors are 1-sigma, absolute, app priori
Param(1, 1) = Delta(2, 1) ' Kent's Beta
Param(2, 1) = Delta(1, 1) '   "    Alpha
Param(3, 1) = Delta(3, 1) '   "    Gamma
If RegrErrors(1, 1) < 0 Or RegrErrors(2, 2) < 0 Or RegrErrors(3, 3) < 0 Then
  GoTo FailedFit
Else
  Param(1, 2) = Sqr(RegrErrors(2, 2))
  Param(2, 2) = Sqr(RegrErrors(1, 1))
  Param(3, 2) = Sqr(RegrErrors(3, 3))
  CovAB = RegrErrors(3, 2): CovAC = RegrErrors(2, 1)
  CovBC = RegrErrors(1, 3)
  Param(1, 3) = CovAB / (Param(1, 2) * Param(3, 2))
  Param(2, 3) = CovAC / (Param(2, 2) * Param(1, 2))
  Param(3, 3) = CovBC / (Param(2, 2) * Param(3, 2))
  Param(1, 4) = Xbar: Param(2, 4) = Ybar: Param(3, 4) = Zbar
  If N < 4 Then MSWD = 0 Else MSWD = Sums / (N - 3)
  StatBar
  Exit Sub
End If
' Equation for error-hyperboloid is [where Sa=Sigma(app)]:
' Sy^2 = (X*Sa)^2 + (Sb)^2 + (Z*Sc)^2
' + 2[X*rho(app,b)*Sa*Sb + X*Z*rho(app,c)*Sa*Sc + Z*rho(b,c)*Sb*Sc]
FailedFit: Failed = True
StatBar
End Sub

Sub PlanarFit(ByVal N&, DP() As DataPoints, Np#(), Failed As Boolean)
Attribute PlanarFit.VB_ProcData.VB_Invoke_Func = " \n14"
' Invoke error-weighted planar regresson & handle data therefrom
Dim ParErr#(3), ConcInt#(6), Param#(3, 4), Resid#()
Dim Prob#, MSWD#, ProjXYmin#, ProjYZmin#
Dim ScatterFactor#, i&, j&, k&
ReDim Resid(N), yf.WtdResid(N)
If MinProb = 0 Then MinProb = Val(Menus("MinProb"))
KentFit N, DP(), Param(), Resid(), MSWD, Failed
If Failed Then
  MsgBox "Can't fit a plane to these data", , Iso
  Exit Sub
End If
If N > 3 Then
  Prob = ChiSquare(MSWD, N - 3)
  ScatterFactor = StudentsT(N - 3) * Sqr(MSWD)
Else
  Prob = 1:  ScatterFactor = 1
End If
yf.MSWD = MSWD: yf.Prob = Prob
For i = 1 To 3
  If Prob > MinProb Then
    ' If assigned errors have >MINPROB prob. of accounting for scatter, use
    '   propagated assigned errors to calc. 95%-conf. plane errors.
    ParErr(i) = 1.96 * Param(i, 2)
  Else
    ' Otherwise, include a Student's-t factor & expand errors to match
    '   the observed scatter.
    ParErr(i) = ScatterFactor * Param(i, 2)
  End If
Next i
If Prob < 0.0001 Then Prob = 0
ProjXYmin = -Param(2, 3) * Param(2, 2) / Param(1, 2)
' X-,value at minimum Y-error at intersection of error-hyperboloid w. XY plane.
ProjYZmin = -Param(3, 3) * Param(2, 2) / Param(3, 2)
' ditto, Z-value on YZ plane.
With yf
  .Slope = Param(1, 1):     .SlopeError = ParErr(1)
  .Intercept = Param(2, 1): .InterError = ParErr(2)
  .Xbar = Param(1, 4):      .Ybar = Param(2, 4)
  .RhoInterSlope = Param(1, 3)
  For i = 1 To N: .WtdResid(i) = Resid(i): Next i
  .Model = 1
End With
Crs(24) = Param(3, 1): Crs(25) = ParErr(3)
Crs(30) = ProjXYmin:   Crs(31) = ProjYZmin
Crs(26) = Param(2, 3): Crs(27) = Param(1, 3)
Crs(28) = Param(3, 3): Crs(29) = Param(3, 4)
If DoPlot Then ProjectPlanar DP(), Np(), Param()
End Sub

Private Sub ProjectPlanar(DP() As DataPoints, Np#(), paR#())
' Determine values & errors for data-pts projected to the X-Y plane along
'  an XYZ-planar regression
Dim i&, Ydelt#, Ypred#, j&, P As Object, PA As Object
Set P = DlgSht("ProjPts")
Set PA = P.OptionButtons("oProjPar")
ProjProc
'PlotProj = False
Do
  ShowBox P, True
  If AskInfo Then ShowHelp "ProjPtsHelp"
Loop Until Not AskInfo
'If IsOff(P.CheckBoxes("cPlotProj")) Then Exit Sub
If Not PlotProj Then Exit Sub
'PlotProj = True
ProjZ = EdBoxVal(P.EditBoxes("eWhat46"))
For i = 1 To N
  With DP(i)
    If ProjZ Then
      Ydelt = .y - (.X * paR(1, 1) + paR(2, 1) + .z * paR(3, 1)) ' Y-increment above plane
      Np(i, 1) = .X * ProjZ / (ProjZ - .z) ' X projected thru ProjZ to XY plane
      Ypred = .X * paR(1, 1) + paR(2, 1)   ' Predicted Y for this np(i,1) on XY plane
      Np(i, 3) = Ypred + Ydelt ' Y projected along plane for this np(i,1)
     Else
      Np(i, 1) = .X
      Np(i, 3) = .y - paR(3, 1) * .z
    End If
    Np(i, 2) = .Xerr: Np(i, 4) = .Yerr
    Np(i, 5) = .RhoXY
  End With
Next i
End Sub

Sub KentResProc()
Attribute KentResProc.VB_ProcData.VB_Invoke_Func = " \n14"
Dim k As Object, L As Object, c As Object, i&, tmp$, s$, vv$, ee$
Set k = DlgSht("KentRes"): Set L = k.Labels: Set c = k.CheckBoxes("cAddWtdResids")
tmp$ = "Best-Fit Plane Solution"
L(2).Text = "X = " & AxX$: L(3).Text = "Y = " & AxY$: L(4).Text = "X = " & AxZ$
If ConcPlot Then
  tmp$ = tmp$ & " to 3-D Concordia"
ElseIf ArgonPlot Then
  tmp$ = tmp$ & " to Argon-Argon Isochron-Data"
End If
For i = 2 To 4: L(i).Visible = Not OtherXY: Next i
With c
  .Visible = OtherXY '
  .Value = xlOff
End With
k.DialogFrame.Text = tmp$
L(5).Text = "a = " & VandE(Crs(1), Crs(2), 2)
L(6).Text = "b = " & VandE(Crs(3), Crs(4), 2)
L(7).Text = "c = " & VandE(Crs(24), Crs(25), 2)
L(11).Text = "X = " & Sd$(Crs(5), 6, -1)
L(12).Text = "Y = " & Sd$(Crs(20), 6, -1)
L(13).Text = "Z = " & Sd$(Crs(29), 6, -1)
L("lErrLev").Text = "errors are " & IIf(yf.Prob < MinProb, "95% conf.", "2 sigma")
For i = 17 To 19
  L(i).Text = RhoRnd(Crs(i + 9))
Next i
L(20).Text = Sd$(Crs(30), 6, -1): L(21).Text = Sd$(Crs(31), 6, -1)
L("lMSWD").Text = "MSWD = " & Mrnd(yf.MSWD) & ",    Probability of fit = " & ProbRnd(yf.Prob)
ShowBox k, True
If IsOn(k.CheckBoxes("cShowRes")) Then
  s$ = "Y = aX + b + cZ" & vbLf & L(5).Text & vbLf & _
       L(6).Text & vbLf & L(7).Text & vbLf
  s$ = s$ & "rho(ab) =" & L(17).Text & vbLf & "rho(ac) =" & L(18).Text & vbLf & _
    "rho(bc) =" & L(19).Text & vbLf & "centroid " & L(11).Text & vbLf & _
    "centroid " & L(12).Text & vbLf & "centroid " & L(13).Text _
      & vbLf & "MSWD = " & Mrnd(yf.MSWD) & ",    Prob. fit = " & ProbRnd(yf.Prob) _
      & vbLf & L("lErrLev").Text
  AddResBox s$, -1, 1, LightGreen
  DetailsShown = True
End If
InsertWtdResids k
End Sub
