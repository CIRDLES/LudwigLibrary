Attribute VB_Name = "cMC"
' Isoplot module cMC
Option Private Module
Option Explicit: Option Base 1

Const MaxCt = 100

Dim tLambda235#, tLambda238#, Ldiff#, Lrat#
Dim X#(), sigmaX#(), y#(), SigmaY#(), Rho#(), TrueYorkX#(), TrueYorkY#()

Sub ConcMcHistClick()
Attribute ConcMcHistClick.VB_ProcData.VB_Invoke_Func = " \n14"

Dim o As Object, c As Object, b As Boolean, s As Object, L As Object, e As Object, G As Object

AssignD "IsoRes", ResBox, e, c, o, L, G, , , s

b = IsOn(c("cHisto")) And IsOn(c("cMC"))
o("oLower").Enabled = b:         o("oUpper").Enabled = b
'g("gPlotWhere").Enabled = b
o("oSeparateSheet").Enabled = b: o("oDataSheet").Enabled = b
If o("oLower") = xlOff And o("oUpper") = xlOff Then o("oUpper") = xlOn
If o("oSeparateSheet") = xlOff And o("oDataSheet") = xlOff Then o("oDataSheet") = xlOn
e("eNbins").Enabled = b: s("sNbins").Enabled = b: L("lNbins").Enabled = b
End Sub

Sub ConcMcBoxClick() ' "MonteCarloErrors" Box clicked in IsoRes
Attribute ConcMcBoxClick.VB_ProcData.VB_Invoke_Func = " \n14"

Dim dc As Boolean
Dim o As Object, c As Object, L As Object, e As Object, s As Object, G As Object, b As Object

AssignD "IsoRes", ResBox, e, c, o, L, G, , b, s

DoMC = IsOn(c("cMC")): dc = (DoMC And ConcPlot)

With c("cWLE"): .Enabled = ConcPlot: .Visible = (DoPlot And .Enabled): End With

c("cWLE_MC").Enabled = (DoMC And ConcPlot)
L("lNtrials").Enabled = DoMC: s("sNtrials").Enabled = DoMC
e("eNtrials").Enabled = DoMC: c("cLIgtZero").Enabled = dc: c("cLIgtZero").Visible = ConcPlot
G("gHisto").Enabled = dc: c("cHisto").Enabled = dc: G("gPlotWhere").Enabled = dc
G("gplotwhere").Visible = 0: G("gwhichinter").Visible = 0
o("oSeparateSheet").Enabled = dc: o("oDataSheet").Enabled = dc
s("sNbins").Enabled = dc: e("eNbins").Enabled = dc: L("lNbins").Enabled = dc
o("oLower").Enabled = dc:       o("oUpper").Enabled = dc
If dc Then Cdecay = IsOn(c("cWLE")): c("cWLE") = Cdecay
If ConcPlot Then ConcMcHistClick
End Sub

Sub ConcMcSpinClick()  ' #trials spinner clicked in IsoRes
Attribute ConcMcSpinClick.VB_ProcData.VB_Invoke_Func = " \n14"
With DlgSht("IsoRes"): .EditBoxes("eNtrials").Text = .Spinners("sNtrials").Value: End With
End Sub

Sub StartMC_click()
Attribute StartMC_click.VB_ProcData.VB_Invoke_Func = " \n14"
DoMC = True
End Sub

Sub ConcMcNbinsClick()  ' #bins spinner clicked in IsoRes
Attribute ConcMcNbinsClick.VB_ProcData.VB_Invoke_Func = " \n14"
With DlgSht("IsoRes")
  .EditBoxes("eNbins").Text = .Spinners("sNbins").Value
End With
End Sub

Sub MonteCarloConcInterErrs(ByVal Ntrials&, ByRef N&, Optional ContinuousDistr = False)
Attribute MonteCarloConcInterErrs.VB_ProcData.VB_Invoke_Func = " \n14"
' Do Monte Carlo concordia-intercept age sol'n, OkAge. or OkAge/o
'  decay-const errs.  Propagates only analytical errs.

Dim Bad As Boolean, SmallChart As Boolean, ConstrainedInters As Boolean
Dim tB As Boolean, eAnch(2) As Boolean, Anch(2) As Boolean
Dim q$, ss$, ax0$, Ay0$, cn$, s$, SB$, vv$, zLprob$, ZprobS$
Dim Lerr$(2), Uerr$(2), Rmean$(2), rs$(2)
Dim c%, TbxN%, zf1%, zf2%, zCt%
Dim MaxCt&, Clr&
Dim ww!, hh!
Dim i&, j&, k&, HiLim95&, LowLim95&, HiLimSig&, LowLimSig&
Dim Tloc&, BadInters&, TotTrials&
Dim Xbar#, Slope0#, Slope#, ErrSl95#, ErrInt95#, Intercept#
Dim Prob#, fProb#, tmp$, Xanch#, Yanch#, v#, zProb#, x1#, x2#, y1#, y2#
Dim Lambda235_0#, Lambda238_0#, TyY#, temp#, tAnch#
Dim BestFitAge#(2), Err95#(2), t_1#(), t_2#()
Dim MeanAge#(2), Lower95#(2), Upper95#(2), Asum#(2)
Dim Rr As Range
Dim DBX As Object, L As Object, Tsht As Object, Csht As Object
Dim Eb As Object, Op As Object, cb As Object, CA As Object, tbx As Object

ReDim X(N), sigmaX(N), y(N), SigmaY(N), Rho(N), xx#(N), yy#(N)
ReDim TrueYorkX#(N), TrueYorkY#(N)
Const Small = 0.000000001, MaxAge = 4600

ViM ContinuousDistr, False
If Len(PlotDat$) = 0 Then PlotDat$ = Ash.Name
Set Rr = Selection ' input data-range

For i = 1 To N
  xx(i) = InpDat(i, 1):     yy(i) = InpDat(i, 3)
  sigmaX(i) = InpDat(i, 2): SigmaY(i) = InpDat(i, 4)
  Rho(i) = InpDat(i, 5)
  ' Xx, Yy are the original (Conv.-concordia) data
  ' X,  Y  will be perturbed by their errors
  ' Convert the data to Conv.-concordia data

  If Inverse Then
    ConcConvert xx(i), sigmaX(i), yy(i), SigmaY(i), Rho(i), True, Bad
    If Bad Then ExitIsoplot
  End If

  X(i) = xx(i): y(i) = yy(i)
Next i

rs$(2) = "Lower": rs$(1) = "Upper"
tLambda238 = Lambda238: tLambda235 = Lambda235
Ldiff = Lambda238 - Lambda235: Lrat = Lambda238 / Lambda235
Slope = 0.1

ShortYork N, Slope, Intercept, Xbar, ErrSl95, ErrInt95, Prob, False, Bad
If Bad Then MsgBox "No age solution for these data", , Iso: Exit Sub

Slope0 = Slope
BestFitAge(1) = 9999: BestFitAge(2) = -9999

For c = 1 To 2
  mConcInter Slope, Intercept, IIf(c = 1, 6000, -999), BestFitAge(c), Bad
  eAnch(c) = (Anchored And Abs(BestFitAge(c) - AnchorAge) < 0.1)
  Anch(c) = (eAnch(c) And AnchorErr = 0)
Next c

ReDim Age#(Ntrials, 2)
ConstrainedInters = IsOn(ResBox.CheckBoxes("cLIgtZero"))

If ConstrainedInters Then
  zLprob$ = "Probability of finite intercept-ages between 0 and" & Str(MaxAge) & " Ma is low"
End If

If Anchored And AnchorErr = 0 Then Xanch = ConcX(AnchorAge): Yanch = ConcY(AnchorAge)

For i = 1 To Ntrials
  TotTrials = 1 + TotTrials
  s$ = "Trials remaining:" & Str(Ntrials - i)

  If ConstrainedInters Then
    If i Mod 10 = 0 Or (BadInters > 0 And BadInters Mod 10 = 0) Then _
      StatBar s$ & "    Failed trials:" & Str(BadInters)
  ElseIf i Mod 20 = 0 Then
    StatBar s$
  End If

  If Cdecay Then
    tLambda235 = Gaussian(Lambda235, Lambda235err)
    tLambda238 = Gaussian(Lambda238, Lambda238err)
    Ldiff = tLambda238 - tLambda235: Lrat = tLambda238 / tLambda235
  End If

  For j = 1 To N

    If Anchored And j = N Then ' Add anchor point

      If AnchorErr > 0 Then

        Do

          If ContinuousDistr Then
            tAnch = (0.5 - Rnd) * 2 * AnchorErr + AnchorAge
          Else
            tAnch = Gaussian((AnchorAge), AnchorErr / 2)

          End If
        Loop Until Not ConstrainedInters Or (tAnch >= 0 And tAnch <= MaxAge)

        Lambda235_0 = Lambda235: Lambda238_0 = Lambda238
        Lambda235 = tLambda235: Lambda238 = tLambda238
        X(j) = ConcX(tAnch, False): y(j) = ConcY(tAnch, False)
        Lambda235 = Lambda235_0: Lambda238 = Lambda238_0

      Else  ' Anchor point fixed at anchor age
        X(j) = Xanch: y(j) = Yanch
      End If

      sigmaX(j) = Small: SigmaY(j) = Small

    ElseIf Rho(j) <> 0 Then
      GaussCorrel xx(j), sigmaX(j), yy(j), SigmaY(j), Rho(j), X(j), y(j)
    Else
      X(j) = Gaussian(xx(j), sigmaX(j)): y(j) = Gaussian(yy(j), SigmaY(j))
    End If

  Next j
  Slope = Slope0

  ShortYork N, Slope, Intercept, 0, 0, 0, 0, True, Bad

  If Not Bad And ConstrainedInters Then

    For c = 1 To 2
      Bad = (Intercept < 0 Or Slope < 0) ' so lower-intercept age is negative

      If Not Bad Then
        mConcInter Slope, Intercept, IIf(c = 1, 6000, -999), Age(i, c), Bad
        If Bad Or Age(i, c) > MaxAge Then Bad = True: Exit For

      End If

    Next c

  End If

  If Bad Then
    Age(i, 1) = 9999: Age(i, 2) = -9999
    BadInters = 1 + BadInters

    If BadInters = Ntrials Then
      zProb = Prob * (TotTrials - BadInters) / TotTrials
      If zProb < 0.05 Then Call MsgBox(zLprob$, , Iso): ExitIsoplot

    End If
    i = i + ConstrainedInters

  ElseIf Not ConstrainedInters Then

    For c = 1 To 2

      If Not Anch(c) Then

        If Not Bad Then mConcInter Slope, Intercept, IIf(c = 1, 6000, -1000), Age(i, c), Bad

        If Bad Then
          Age(i, 1) = 9999: Age(i, 2) = -9999
          BadInters = 1 + BadInters
          Exit For

        End If

      End If

    Next c

  End If

Next i

zProb = Prob * (TotTrials - BadInters) / TotTrials
ZprobS$ = Sd(zProb, 2, , True)

If zProb < 0.025 Then
  tmp$ = "The constraint" & IIf(ConstrainedInters, "s", "") & " of regressable data" & vbLf & _
    IIf(ConstrainedInters, " plus geologically possible age-intercepts", "") & vbLf & _
    " reduces the overall probability-of-fit to only " & ZprobS$
  MsgBox tmp$, , Iso

  If DoPlot Then Exit Sub
  ExitIsoplot

End If

ReDim t_1(Ntrials), t_2(Ntrials)

For i = 1 To Ntrials
  t_1(i) = Age(i, 1): t_2(i) = Age(i, 2)
Next i

With App

  For c = 1 To 2

    If Anch(c) Then
      MeanAge(c) = AnchorAge
    Else
      MeanAge(c) = BestFitAge(c)
    End If

  Next c

End With

If Cdecay Or ConstrainedInters Then

  If DoPlot And ConstrainedInters Then

    If MeanAge(1) = MeanAge(2) Then
      Crs(1) = 0: Crs(3) = 0

    Else ' Calc slope/inter to construct regression line
      x1 = ConcX(MeanAge(1), False, True): x2 = ConcX(MeanAge(2), False, True)
      If x1 = x2 Then MsgBox "Error in Monte Carlo solution for age", , Iso: ExitIsoplot
      y1 = ConcY(MeanAge(1), False, True): y2 = ConcY(MeanAge(2), False, True)
      Crs(1) = (y1 - y2) / (x1 - x2): Crs(3) = y1 - Slope * x1

    End If

  End If

End If

QuickSort t_1()
QuickSort t_2()

For i = 1 To Ntrials
  Age(i, 1) = t_1(i): Age(i, 2) = t_2(i)
Next i

Erase t_1, t_2

For c = 1 To 2

  If Anch(c) Then
    Rmean$(c) = Sd$(AnchorAge, 3)
    Lerr$(c) = "-" & Sd$(AnchorErr, 3): Uerr$(c) = "+" & Sd$(AnchorErr, 3)
  ElseIf eAnch(c) And Not ContinuousDistr Then
    Lower95(c) = AnchorAge - AnchorErr
    Upper95(c) = AnchorAge + AnchorErr
    Lerr$(c) = tSt(-AnchorErr)
    Uerr$(c) = "+" & tSt(AnchorErr)
  Else
    LowLim95 = 0.025 * Ntrials:   HiLim95 = Ntrials - LowLim95
    Lower95(c) = Age(LowLim95, c):  Upper95(c) = Age(HiLim95, c)
    Err95(c) = (Upper95(c) - Lower95(c)) / 2
    NumAndErr MeanAge(c), Err95(c), 2, Rmean$(c), ""
    Lerr$(c) = "-" & ErFo(MeanAge(c), Lower95(c) - MeanAge(c), 2)
    Uerr$(c) = "+" & ErFo(MeanAge(c), Upper95(c) - MeanAge(c), 2)
  End If

  If eAnch(c) Then Rmean$(c) = tSt(AnchorAge)
Next c

AssignD , ResBox, Eb, cb, Op
StatBar "Wait"

If cb("cHisto") = xlOn And Not AddToPlot Then
  SmallChart = (Op("oDataSheet") = xlOn)
  Sheets.Add

  If Not SmallChart Then
    PlotDat$ = "PlotDat"
    MakeSheet PlotDat$, ChrtDat
  End If

  Set Tsht = Ash
  j = 1 - IsOn(Op("oLower")): k = 0
  ax0$ = AxX$: Ay0$ = AxY$: AxY$ = ""
  AxX$ = IIf(j = 1, "Upper", "Lower") & " intercept age (Ma)"
  Nbins = MinMax(10, 1200, Val(Eb("eNbins").Text))
  Clr = IIf(j = 1, vbRed, vbGreen)
  tB = (IsOn(Op("oLower")) And IsOn(cb("cLIgtZero")))
  GaussCumProb Ntrials, False, SmallChart, Age(), , Clr, tB, , , 1 - IsOn(Op("oLower"))
  AxX$ = ax0$: AxY$ = Ay0$
  Set Csht = Ach: Set CA = Csht.ChartArea
  ActiveWindow.Zoom = 100

  With Csht

    If SmallChart Then

      With .PlotArea
        .Interior.ColorIndex = ClrIndx(vbWhite)
        .Border.LineStyle = xlNone
        .Left = 9: .Width = CA.Width - 18
        .Top = 9 ': .Height = cA.Height - 10

      End With
      CA.Interior.ColorIndex = ClrIndx(vbWhite)
    End If

    With .Axes(xlValue)
      If SmallChart Then .Border.LineStyle = xlNone
      .HasTitle = False
      .TickLabelPosition = xlNone

      If Not SmallChart Then
        .MinorTickMark = xlTickMarkNone
        .MajorTickMark = xlTickMarkNone
      End If

    End With

    With .Axes(xlCategory)
      .HasMajorGridlines = False '.HasTitle = False:
      If SmallChart Then .MinorTickMark = xlTickMarkNone
      .Border.Weight = IIf(SmallChart, xlThick, xlMedium)
    End With
    .Select

    With .Axes(1)
      zf1 = IIf(SmallChart, .TickLabels.Font.Size, 14): zCt = 0
      zf2 = IIf(SmallChart, .AxisTitle.Characters.Font.Size, 20)

      Do ' Kluge to deal with Excel bug
        If zCt > 6 Then Call MsgBox("Unresolved Excel bug occurred - sorry", , Iso): ExitIsoplot
        RescaleOnlyShapes False, , Tsht.Name

      If .TickLabels.Font.Size = zf1 And .AxisTitle.Characters.Font.Size = zf2 Then Exit Do

      If Not .HasTitle Then Exit Do

        .TickLabels.Font.Size = zf1

        With .AxisTitle
          .Characters.Font.Size = zf2
          .Font.Bold = SmallChart

        End With
        zCt = zCt + 1

      Loop

    End With

    If SmallChart Then
      .CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
      Sheets(PlotDat$).Activate ' Switch to source-data sheet
      j = 0: k = 0
      On Error GoTo 11

      For i = 1 To UBound(RangeIn) '9
        j = Max(j, RangeIn(i).Column + RangeIn(i).Columns.Count)
        k = Max(k, RangeIn(i).Row)

      Next i
11:   On Error GoTo 0

      Ash.Pictures.Paste.Select

      With Selection ' Scale down size

        With .ShapeRange
          .Fill.ForeColor.RGB = vbWhite
          If zCt = 0 Then .PictureFormat.CropBottom = 25
        End With

        ww = .Width: hh = .Height
        .Width = 250: .Height = hh / ww * .Width
        .Left = Cells(k, j).Left
        .Top = Max(0, Cells(k, j).Top - 4 * Cells(k, j).Height)

        With .Border: .Color = vbBlack: .Weight = xlHairline: End With

      End With
      DelSheet Csht
      DelSheet Tsht
      Rr.Select

    End If

  End With

End If

StatBar
Set DBX = DlgSht("MC"): Set L = DBX.Labels
L("lProb").Text = ""
If ConstrainedInters Then L("lProb").Text = "Probability of fit = " & ZprobS$

For c = 1 To 2
  i = c + 5: j = c + 7: k = c + 3
  L(k).Text = Rmean$(c)
  L(c + 9).Text = "Ma"
  L(j).Text = Lerr$(c): L(i).Text = Uerr$(c)
Next c

s$ = "Errors are 95%-confidence, " & IIf(Cdecay, "and", "but do not") _
  & " include decay-constant errors."
q$ = "ntercepts constrained between 0 and" & Str(MaxAge) & " Ma"

If ConstrainedInters Then
  s$ = s$ & "  I" & q$ & "."
ElseIf BadInters > 0 Then
  s$ = s$ & "  The" & Str(BadInters) & " " & Plural("solution", BadInters) & " without concordia " & _
    "intercepts w" & IIf(BadInters = 1, "as", "ere") & " assigned ages of +9999 and -9999 Ma."
End If

L("lErrExp").Text = s$
ShowBox DBX, True

If IsOn(ResBox.CheckBoxes("cShowRes")) Then
  s$ = "Monte Carlo Solution on" & Str(Ntrials) & " trials" & vbLf & "WITH"
  If Not Cdecay Then s$ = s$ & "OUT"
  s$ = s$ & " decay-constant errors" & vbLf
  If ConstrainedInters Then s$ = s$ & "i" & q$ & "; " & "prob-fit=" & ZprobS$ & vbLf

  For i = 5 To 4 Step -1
    ss$ = "   Ma"

    If L(i).Text & L(i + 2).Text & L(i + 4).Text = "------" Then
      s$ = s$ & "no intercept" & vbLf
    Else
      s$ = s$ & L(i).Text & "   " & L(i + 2).Text & "   " & L(i + 4).Text & "   Ma" & vbLf

    End If

  Next i

  s$ = s$ & "at 95% confidence"
  Set tbx = Ash.DrawingObjects
  TbxN = tbx.Count    ' Put this results-box just to the rt
  Set tbx = tbx(TbxN) '  of the analytical-sol'n results-box.
  AddResBox s$, -1, 0, Mauve, tbx.Left + tbx.Width
End If

If Cdecay Or ConstrainedInters Then
  ' Replace Y'fit slope & intercept with that calculated from the
  '  mean of the MonteCarlo age sol'ns
  x1 = ConcY(MeanAge(2), Inverse, -1): x2 = ConcY(MeanAge(1), Inverse, -1)
  y1 = ConcX(MeanAge(2), Inverse, -1): y2 = ConcX(MeanAge(1), Inverse, -1)
  If y1 = y2 Then MsgBox "Error in Monte Carlo solution for age" & vbLf & _
    "(lower intercept and upper intercept have the same age)", , Iso: ExitIsoplot
  Crs(1) = (x1 - x2) / (y1 - y2)
  Crs(3) = ConcY(MeanAge(1), Inverse, -1) - Crs(1) * ConcX(MeanAge(1), Inverse, -1)
  Crs(8) = MeanAge(1): Crs(9) = MeanAge(2)
End If

Crs(12) = Lower95(1): Crs(14) = Lower95(2)
Crs(13) = Upper95(1): Crs(15) = Upper95(2)
Crs(10) = Err95(1):   Crs(11) = Upper95(2)

For c = 1 To 2
  s$ = IIf(Rmean$(c) = "--", "", Rmean$(c) & "  " & _
       L(c + 5).Text & "/" & L(c + 7).Text & "  Ma")
  If c = 1 Then Uir$ = s$ Else Lir$ = s$
Next c

End Sub

Private Sub mConcInter(ByVal Slope#, ByVal Intercept#, ByVal t0#, _
  T#, Bad As Boolean)
' Calculate intercepts of a line with Tera-Wasserbug concordia curve.

Dim ct%, test#, TrialT#, ConcX#, ConcY#, cs#, xc#

Const Toler = 0.0001
TrialT = t0: test = 99

Do
  ct = 1 + ct
  ConcX = Exp(tLambda235 * TrialT) - 1
  ConcY = Exp(tLambda238 * TrialT) - 1
  cs = Lrat * Exp(Ldiff * TrialT)
  xc = (Intercept + cs * ConcX - ConcY) / (cs - Slope)

  If xc >= -1 Then
    T = Log(1 + xc) / tLambda235
    test = Abs(TrialT - T)
    TrialT = T
  Else
    TrialT = TrialT - 10 * (cs > Slope)
  End If

Loop Until test < Toler Or ct > MaxCt

If ct > MaxCt Then Bad = True
End Sub

Sub ConcConvert(X#, sigmaX#, y#, SigmaY#, _
  RhoXY#, TWin As Boolean, Bad As Boolean, _
  Optional z, Optional SigmaZ, Optional RhoXZ, Optional RhoYZ)
' Convert T-W concordia data to Conv., or vice-versa
' eg 238/206-207/206[-204/206] to/from 207/235-206/238[-204/238]

Dim BadSqrt As Boolean, d3 As Boolean
Dim xp#, yp#, zP#, rXY#, rXZ#, rYZ#, A#, b#, c#
Dim Ap#, Bp#, Cp#, Rab#, ABp#, Rac#, Rbc#, test#, xp2#, yp2#

xp = Abs(sigmaX / X):  yp = Abs(SigmaY / y): rXY = RhoXY: Bad = False
d3 = NIM(z): xp2 = xp * xp: yp2 = yp * yp
TestSqrt xp2 + yp * yp - 2 * xp * yp * rXY, ABp, Bad

If Bad Then GoTo BadErrors

If TWin Then   ' T-W in, Conv. out
  A = y / X * Uratio ' 207/235
  b = 1 / X          ' 206/238

  If ABp <> 0 Then
    Ap = ABp: Bp = xp
    Rab = (xp - yp * rXY) / ABp

  End If

Else         ' Conv. in, T-W out
  A = 1 / y  ' 238/206
  b = X / y / Uratio ' 207/206
  Ap = yp:            Bp = ABp

  If ABp <> 0 Then
    Rab = (yp - xp * rXY) / ABp
  End If

End If

If Abs(Rab) > 1 Then GoTo BadErrors

If d3 Then
  zP = Abs(SigmaZ / z): rXZ = RhoXZ: rYZ = RhoYZ

  If TWin Then
    c = z / X   ' 204/238
    TestSqrt xp2 + zP * zP - 2 * xp * zP * rXZ, Cp, Bad
    If Bad Then GoTo BadErrors
    Rac = (xp2 - xp * yp * rXY - xp * zP * rXZ + yp * zP * rYZ) / Ap / Cp
    Rbc = (xp2 - xp * zP * rXZ) / Bp / Cp
  Else
    c = z / y   ' 204/206
    TestSqrt yp2 + zP * zP - 2 * yp * zP * rYZ, Cp, Bad
    If Bad Then GoTo BadErrors
    Rac = (Ap * Ap + Cp * Cp - zP * zP) / (2 * Ap * Cp)
    Rbc = (xp * zP * rXZ - Ap * (Ap - Bp * Rab - Cp * Rac)) / (Bp * Cp)
  End If

  If Abs(Rac) > 1 Or Abs(Rbc) > 1 Then GoTo BadErrors
End If

X = A: y = b: sigmaX = Ap * A: SigmaY = Bp * b: RhoXY = Rab

If d3 Then
  z = c
  SigmaZ = Cp * z: RhoXZ = Rac: RhoYZ = Rbc
End If

Exit Sub

BadErrors: Bad = True
MsgBox "Impossible combination of errors+error-correlations in input-data", , Iso
End Sub

Sub ShortYork(ByVal N&, Slope#, Intercept#, Xbar#, _
  SlopeErr#, InterErr#, Prob#, NoErrors As Boolean, Bad As Boolean, _
  Optional pMSWD, Optional ArX, Optional ArXsig, Optional ArY, Optional ArYsig, Optional ArRho, _
  Optional Arxint_, Optional ArXintErr_)
Attribute ShortYork.VB_ProcData.VB_Invoke_Func = " \n14"
' Slope is passed as the estimated starting slope
Dim i&, ct%, Slope0#, xs#, YS#, WeightSum#, A#, b#, c#, d#, e#, tmp1#, tmp2#
Dim u#, v#, U2#, v2#, uv#, WtSqr#, Ybar#, test#, RMW#, TrueXval#, Tslope#
Dim ArXint#, ArXintErr#, SumXZ#, SumX2Z#, Yresid#, Numer#, Denom#
Dim VarIntApr#, VarSlApr#, CovIntSl#, MSWD#, Sums#, WtdResid#

ReDim Xwt#(N), Ywt#(N), Weight#(N), MeanWt#(N), txy#(N, 2)

Const Toler = 0.000001

If NIM(ArX) Then
  ReDim X(N), sigmaX(N), y(N), SigmaY(N), Rho(N)

  For i = 1 To N
    X(i) = ArX(i): sigmaX(i) = ArXsig(i)
    y(i) = ArY(i): SigmaY(i) = ArYsig(i): Rho(i) = ArRho(i)
    txy(i, 1) = X(i): txy(i, 2) = y(i)
  Next i

End If

If N = 2 Then
  Tslope = (y(2) - y(1)) / (X(2) - X(1))
Else
  RobustReg2 txy, Tslope, SlopeOnly:=True
End If

Erase txy
If IsNumeric(Tslope) Then Slope = Tslope
Bad = False

For i = 1 To N
  Xwt(i) = 1 / SQ(sigmaX(i)): Ywt(i) = 1 / SQ(SigmaY(i))
  MeanWt(i) = 1 / (sigmaX(i) * SigmaY(i))
Next i

Do
  Slope0 = Slope: ct = 1 + ct

  For i = 1 To N
    tmp1 = Slope * Slope * Ywt(i) + Xwt(i)
    tmp2 = Slope * Rho(i) * MeanWt(i)
    Weight(i) = Xwt(i) * Ywt(i) / (tmp1 - 2 * tmp2)
  Next i

  WeightSum = Sum(Weight())
  xs = SumProduct(Weight(), X())
  YS = SumProduct(Weight(), y())
  Xbar = xs / WeightSum: Ybar = YS / WeightSum
  c = 0: d = 0: e = 0

  For i = 1 To N
    WtSqr = Weight(i) * Weight(i)
    u = X(i) - Xbar: U2 = u * u
    v = y(i) - Ybar: v2 = v * v: uv = u * v
    RMW = Rho(i) / MeanWt(i)
    c = c + (U2 / Ywt(i) - v2 / Xwt(i)) * WtSqr
    d = d + (uv / Xwt(i) - RMW * U2) * WtSqr
    e = e + (uv / Ywt(i) - RMW * v2) * WtSqr
  Next i

  test = c * c + 4 * d * e

If test < 0 Or d = 0 Then Bad = True: Exit Sub

  Slope = (Sqr(test) - c) / (2 * d)
  ct = 1 + ct
  test = Abs(Slope / Slope0 - 1)
Loop Until test < Toler Or ct > MaxCt

If ct > MaxCt Then Bad = True: Exit Sub

Intercept = Ybar - Slope * Xbar

If NoErrors Then Exit Sub

' Titterington/Halliday algorithm for regression errors
SumXZ = 0: SumX2Z = 0

For i = 1 To N
  Yresid = Intercept + Slope * X(i) - y(i) ' Unwtd Y-resids
  Sums = Sums + Weight(i) * Yresid * Yresid  ' Wtd Y-resids^2
  Numer = Weight(i) * Yresid * (Rho(i) * MeanWt(i) - Ywt(i) * Slope)
  TrueXval = X(i) + Numer / (Xwt(i) * Ywt(i))
  SumX2Z = SumX2Z + TrueXval * TrueXval * Weight(i)
  SumXZ = SumXZ + TrueXval * Weight(i)
  On Error Resume Next
  TrueYorkX(i) = TrueXval
  On Error GoTo 0
Next i

Denom = SumX2Z * WeightSum - SumXZ * SumXZ
VarIntApr = 0: VarSlApr = 0

If Denom > 0 Then
  VarIntApr = SumX2Z / Denom
  VarSlApr = WeightSum / Denom
  CovIntSl = -SumXZ / Denom
End If

SlopeErr = 1.96 * Sqr(VarSlApr): InterErr = 1.96 * Sqr(VarIntApr)

If NIM(Arxint_) Then
  Xintercept ArXint, ArXintErr, Intercept, InterErr, Slope, SlopeErr, Xbar
  Arxint_ = ArXint:  ArXintErr_ = ArXintErr
End If

If N > 2 Then
  MSWD = Sums / (N - 2)
  Prob = ChiSquare(MSWD, N - 2)

  If NIM(pMSWD) Then
    pMSWD = MSWD
  End If

Else
  MSWD = 0: Prob = 1
End If

End Sub

Sub MonteCarloThUerrs(MeanAgeYr#, MedAgeYr#, ModeAgeYr#, g0in#, ByVal ar08#, _
  ByVal ThUerr#, ByVal ar48#, ByVal GammaErr#, ByVal Rho#, e95Lt#, _
  e95Ut#, E95Lg#, E95Ug#, RhoAgeG0#, Bad As Boolean, _
  ByVal NLE As Boolean, ByVal AtomRat As Boolean, ByVal Finite As Boolean)
Attribute MonteCarloThUerrs.VB_ProcData.VB_Invoke_Func = " \n14"
' Calculate MonteCarlo errors for 230Th/U ages & Initial 234/238, as
'  well as the err-corr between the two.
' 11/06/18 -- added Median, Mode age calculations & passed variables

Dim StatDispOK As Boolean
Dim SB$, ss$
Dim i&, N&, j&, rn&, Shi&, Slo&, Chi&, Clo&, rw&, Co&, nbad&
Dim Tsum#, Gsum#, tMean#, Gmean#, LL&, t0#, s1#, s2#, s3#, AgeMode#
Dim Tres#, Gres#, Denom#, Elapsed#, L0in#, L0err#, L4in#
Dim L4err#, Rsums#, PredT#, Gsums#, Th#, u#, Age#
Dim Eb As Object

Const inf = 1.70141183460469E+38, MinTrials = 1000, MaxTrials = 10000

Bad = False
StatDispOK = False
On Error GoTo McD
App.DisplayStatusBar = True
StatDispOK = True
McD: On Error GoTo 0

If Not NLE Then
  L0in = Lambda230:     L4in = Lambda234
  L0err = Lambda230err: L4err = Lambda234err '1sigma lambda errs
End If

Set Eb = DlgSht("ThUage").EditBoxes("eNtrials")
N = MinMax(MinTrials, MaxTrials, Val(Eb.Text))
Eb.Text = tSt(N)
ReDim T#(N), G#(N)
rn = 0: Rsums = 0: Gsums = 0
Randomize Timer
rn = 0: t0 = Now
ss$ = "Trials remaining:"

Do
  rn = 1 + rn: LL = 1 + LL

  If StatDispOK Then
    j = N - LL

    If LL = 1 Or j Mod 50 = 0 Then
      SB$ = ss$ & Str(N - rn)
      If Finite Then SB$ = SB$ & "   " & Str(nbad)
      StatBar SB$
    End If

  End If

  If Rho Then
    GaussCorrel ar08, ThUerr, ar48, GammaErr, Rho, Th, u
  Else
    Th = Gaussian(ar08, ThUerr)
    If GammaErr Then u = Gaussian(ar48, GammaErr) Else u = ar48
  End If

  If Not NLE Then
    Lambda230 = Gaussian(L0in, L0err)
    Lambda234 = Gaussian(L4in, L4err)
    LambdaDiff = Lambda230 - Lambda234
    LambdaK = Lambda230 / LambdaDiff

    If AtomRat Then              ' That is, if calculating activity ratios by multi-
      Th = Th * Lambda230 / L0in '  plying atomic ratios by decay-constant ratios.
      u = u * Lambda234 / L4in
    End If

  End If

  ThUage Th, u, Age

  If (Finite And Age > 0) Or (Not Finite And Age <> 0) Then
    T(rn) = Age
    G(rn) = 1 + (u - 1) * Exp(Lambda234 * Age)

    If Finite And G(rn) <= 0 Then
      rn = rn - 1
      nbad = 1 + nbad
    Else
      Tsum = Tsum + T(rn)
      Gsum = Gsum + G(rn)
    End If

  ElseIf Not Finite And Age = 0 Then
    Age = inf
    T(rn) = Age: G(rn) = 0

  Else
    rn = rn - 1
    nbad = 1 + nbad
  End If

  If LL = 10000 Then
    Elapsed = (Now - t0) * 86400
    PredT = Elapsed * N / Max(rn, 0.0001)

    If Elapsed > 1 And PredT > 300 Then
      MsgBox "Not enough successful solutions to continue.  Exiting..."

      ExitIsoplot

    End If

  End If

Loop Until rn = N

If Not NLE Then
  Lambda230 = L0in
  Lambda234 = L4in
End If

If Not Finite And rn < 0.025 * N Then
  Bad = True
  Exit Sub
End If

tMean = Tsum / rn
Gmean = Gsum / rn
'Calc err-corr between Th230/U age & Gamma0
s1 = 0: s2 = 0: s3 = 0

For i = 1 To N

  If T(i) <> inf Then
    Tres = T(i) - tMean
    Gres = G(i) - Gmean
    s1 = s1 + Tres * Gres
    s2 = s2 + Tres * Tres
    s3 = s3 + Gres * Gres
  End If

Next i

Denom = s2 * s3

If GammaErr = 0 Or Denom <= 0 Then RhoAgeG0 = 0 Else RhoAgeG0 = s1 / Sqr(Denom)

If StatDispOK Then StatBar "Sorting"
QuickSort T()
QuickSort G()
Clo = 0.025 * rn: Chi = 0.975 * rn
e95Lt = T(Clo): e95Ut = T(Chi)
E95Lg = G(Clo): E95Ug = G(Chi)
If e95Lt = inf Then e95Lt = 0
If e95Ut = inf Then e95Ut = 0
If E95Lg = inf Then E95Lg = 0
If E95Ug = inf Then E95Ug = 0

If e95Lt = 0 And e95Ut = 0 Then Bad = True

MeanAgeYr = tMean
MedAgeYr = iMedian(T)       ' 11/06/18 -- added
ModeAgeYr = Imode(T(), rn)  ' 11/06/18 -- added

If Finite Then
  MeanAgeYr = tMean
  g0in = Gmean
End If

If StatDispOK Then StatBar
End Sub

Public Function Imode(X#(), ByVal Nvals&) ' X() MUST BE SORTED
' Determine the mode of a vector

Dim BinWidth#, NumBins%, MaxX#, Xct&, MaxFreq%, LastMax%, BinNum%
Dim Freq%, MinX#, MaxBin%

' 11/06/18 -- created

If Nvals < 3 Then Imode = "#VALUE": Exit Function

NumBins = Sqr(Nvals) / 2   ' arbitrary but useful

MinX = X(1)
BinWidth = (X(Nvals) - X(1)) / NumBins
MaxFreq = 0
Xct = 1

For BinNum = 1 To NumBins
  Freq = 0
  MaxX = MinX + BinNum * BinWidth

  Do While X(Xct) < MaxX
    Freq = 1 + Freq
    Xct = Xct + 1
  Loop

  If Freq > MaxFreq Then
    MaxFreq = Freq
    MaxBin = BinNum
  End If

Next BinNum

Imode = X(1) + BinWidth * (MaxBin - 0.5)
End Function

Sub MCyorkfit(ByVal N&, ByVal Ntrials&, Bad As Boolean, _
  Optional SlpErr, Optional YintErr, Optional Xinterr)
' MonteCarlo errors for Yorkfit

Dim i&, j&, k&, HiLim95&
Dim DoSlp As Boolean, DoY As Boolean, DoX As Boolean, f!
Dim xx#, yy#, Slp#, Yint#, Slp0#, Yint0#, Xint0#
Dim Lwr68&, Lwr95&, Upr68&, Upr95&

ReDim s#(Ntrials), Yi#(Ntrials), xi#(Ntrials)
ReDim X#(N), y(N), sigmaX(N), SigmaY(N), Rho(N)

DoSlp = NIM(SlpErr)
DoX = NIM(Xinterr)
DoY = NIM(YintErr)
Bad = False

For i = 1 To N
  X(i) = InpDat(i, 1): y(i) = InpDat(i, 3)
  sigmaX(i) = InpDat(i, 2): SigmaY(i) = InpDat(i, 4)
  Rho(i) = InpDat(i, 5)
Next i

ShortYork N, Slp0, Yint0, 0, 0, 0, 0, True, Bad

If Bad Then MsgBox "No Yorkfit for data as input": Exit Sub
Xint0 = -Yint0 / Slp0

For i = 1 To Ntrials

  If (i - 1) Mod Hun = 0 Then
    StatBar "Wait    " & Str(Ntrials - i + 1)
  End If

  For j = 1 To N
    GaussCorrel (InpDat(j, 1)), (InpDat(j, 2)), (InpDat(j, 3)), _
      (InpDat(j, 4)), (InpDat(j, 5)), X(j), y(j)

  Next j

  ShortYork N, Slp, Yint, 0, 0, 0, 0, True, Bad

  If Not Bad Then
    k = 1 + k
    If DoSlp Then s(k) = Slp
    If DoY Then Yi(k) = Yint
    If DoX Then xi(k) = -Yint / Slp

  End If

Next i

f = k / Ntrials

If f < 0.95 Then
  MsgBox Str(Drnd(Hun * (1 - f), 2)) & _
    "% of the Monte Carlo Yorkfit trials failed -- abandoning Monte Carlo."

  Bad = True: Exit Sub

End If

StatBar "Wait"
Ntrials = k

ReDim Preserve s(Ntrials), xi(Ntrials), Yi(Ntrials)

Lwr68 = (1 - 0.6826) / 2 * Ntrials: Upr68 = Ntrials - Lwr68
Lwr95 = 0.025 * Ntrials: Upr95 = 0.975 * Ntrials
' 1=68%conf -err   2=68%conf +err    3=68%conf +-err
' 4=95%conf -err   5=95%conf +err    5=95%conf +-err

If DoSlp And DoY Then ' Must calculate before sorting!
  yf.RhoInterSlope = App.Correl(s(), Yi())
End If

If DoSlp Then
  ReDim SlpErr(6)
  QuickSort s()
  SlpErr(1) = Slp0 - s(Lwr68): SlpErr(2) = s(Upr68) - Slp0
  SlpErr(3) = (SlpErr(2) + SlpErr(1)) / 2
  SlpErr(4) = Slp0 - s(Lwr95): SlpErr(5) = s(Upr95) - Slp0
  SlpErr(6) = (SlpErr(4) + SlpErr(5)) / 2
End If

If DoY Then
  ReDim YintErr(6)
  QuickSort Yi()
  YintErr(1) = Yint0 - Yi(Lwr68): YintErr(2) = Yi(Upr68) - Yint0
  YintErr(3) = (YintErr(2) + YintErr(1)) / 2
  YintErr(4) = Yint0 - Yi(Lwr95): YintErr(5) = Yi(Upr95) - Yint0
  YintErr(6) = (YintErr(4) + YintErr(5)) / 2
End If

If DoX Then
  ReDim Xinterr(6)
  QuickSort xi()
  Xinterr(1) = Xint0 - xi(Lwr68): Xinterr(2) = xi(Upr68) - Xint0
  Xinterr(3) = (Xinterr(2) + Xinterr(1)) / 2
  Xinterr(4) = Xint0 - xi(Lwr95): Xinterr(5) = xi(Upr95) - Xint0
  Xinterr(6) = (Xinterr(4) + Xinterr(5)) / 2
End If

StatBar
End Sub
