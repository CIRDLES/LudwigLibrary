Attribute VB_Name = "Argon"
Option Private Module
Option Base 1: Option Explicit

Sub AddArTextbox(ByVal Cap$, Optional FontSize)
Attribute AddArTextbox.VB_ProcData.VB_Invoke_Func = " \n14"
Dim fs%, T As Textbox
StatBar "adding results-box to chart"
Set T = IsoChrt.TextBoxes.Add(0, 0, 0, 0)
If NIM(FontSize) Then fs = FontSize Else fs = Opt.IsochResFontSize

With T
  .Characters.Text = Cap$:       .Font.Name = Opt.IsochResFont
  .Interior.ColorIndex = xlNone: .Font.Size = fs
  .VerticalAlignment = xlTop:    .HorizontalAlignment = xlCenter
  .Orientation = xlHorizontal:   .AutoSize = False
  .Border.Color = vbBlack
  .RoundedCorners = True:     '.Shadow = Opt.IsochResboxShadw
  .Placement = xlMoveAndSize: .Interior.ColorIndex = ClrIndx(vbWhite)
  .PrintObject = True:        .AutoSize = True
  .ShapeRange.Shadow.Type = msoShadow6
  .Name = "ArAge"
End With

IncreaseLineSpace T, 1.2
ConvertSymbols T
Superscript Phrase:=T
GetScale

With T
  .Left = PlotBoxLeft + PlotBoxWidth / 2 - .Width / 2
  .Top = PlotBoxBottom - .Height - 15
End With

FreeSpace T
End Sub

Sub ConvertArData(ByVal N&)
Attribute ConvertArData.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%

For i = 1 To N ' Put x +- y +- rho data back in 1st 5 cols

  For j = 1 To 5
    InpDat(i, j) = InpDat(i, j + 3)
  Next j

  For j = 6 To 8: InpDat(i, j) = 0: Next j

Next i

ArgonStep = False: AutoScale = True: Isotype = 2
AxX$ = IIf(Normal, "39Ar/36Ar", "39Ar/40Ar")
AxY$ = IIf(Normal, "40Ar/36Ar", "36Ar/40Ar")
Eellipse = True: WasPlat = True
End Sub

Sub CalcArgonStepAge(Age#, AgeError#, ByVal Ar4036_0#, ByVal j%)
Attribute CalcArgonStepAge.VB_ProcData.VB_Invoke_Func = " \n14"
' Calculate age of the jth argon-argon step-release step from the measured (total-gas) ratios;
' Ignore error in "J" but propagate errors & correlations from the initial 40Ar/36Ar correction.
Dim Ar3936#, Ar3936err#, Ar3936ferr#, Ar4036#, Ar4036ferr#
Dim Rho3936_4036#, Ar4039#, Rho4039_4036#, Ar3940#
Dim Ar3940err#, Ar3940ferr#, Ar3640#, Ar3640ferr#
Dim Ar4039ferr#, Rfact#, Ar4036err#, Ar3640err#, Ar40rad_Ar39#
Dim Ar40rad_Ar39err#, Rho3940_3640#, Ar40rad_k40ferr#, Afact#

If Normal Then
  Ar3936 = InpDat(j, 4): Ar3936err = InpDat(j, 5): Ar3936ferr = Ar3936err / Ar3936:
  Ar4036 = InpDat(j, 6): Ar4036err = InpDat(j, 7): Ar4036ferr = Ar4036err / Ar4036
  Rho3936_4036 = InpDat(j, 8): Ar4039 = Ar4036 / Ar3936
  Ar4039ferr = Sqr(SQ(Ar3936ferr) + SQ(Ar4036ferr) _
    - 2 * Ar3936ferr * Ar4036ferr * Rho3936_4036)
  Rho4039_4036 = (Ar3936ferr * Rho3936_4036 - Ar4036ferr) / Ar4039ferr
Else
  Ar3940 = InpDat(j, 4): Ar3940err = InpDat(j, 5): Ar3940ferr = Ar3940err / Ar3940
  Ar3640 = InpDat(j, 6): Ar3640err = InpDat(j, 7): Ar3640ferr = Ar3640err / Ar3640
  Rho3940_3640 = InpDat(j, 8): Ar4039 = 1 / Ar3940: Ar4036 = 1 / Ar3640
  Ar4039ferr = Ar3940ferr: Ar4036ferr = Ar3640ferr
  Rho4039_4036 = -Rho3940_3640
End If

Rfact = (Ar4036 - Ar4036_0) / Ar4036
Afact = Ar4036_0 / (Ar4036 - Ar4036_0)
Ar40rad_Ar39 = Rfact * Ar4039
Ar40rad_k40ferr = Sqr(SQ(Ar4039ferr) + SQ(Afact * Ar4036ferr) _
  + 2 * Afact * Ar4039ferr * Ar4036ferr * Rho4039_4036)
Ar40rad_Ar39err = Ar40rad_k40ferr * Ar40rad_Ar39

ArgonAge Ar40rad_Ar39, Ar40rad_Ar39err, Age, AgeError, True
End Sub

Sub Cumulative(Cum As Boolean, ByVal Nsteps%, GasMult#, Gas#())
Attribute Cumulative.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, tmp#
Cum = True ' Cumulative?

For i = 2 To Nsteps
  If Gas(i) < Gas(i - 1) Then Cum = False
Next i

tmp = Gas(Nsteps)
Cum = Cum And ((tmp > 0.999 And tmp < 1.001) Or (tmp > 99.9 And tmp < 100.1))
GasMult = 1

If Cum Then
  If Gas(Nsteps) > 10 Then GasMult = Hun ' To convert from %39Ar to fract. 39Ar
Else ' Not cumulative - calculate
  For i = 2 To Nsteps: Gas(i) = Gas(i) + Gas(i - 1): Next
  For i = 1 To Nsteps: Gas(i) = Gas(i) / Gas(Nsteps): Next
  GasMult = 1
End If

End Sub

Sub CheckOuterTrends(ByVal Nsteps%, ByVal ns%, Age#(), SigmaAge#(), _
  v#(), MidGas#(), ByVal MeanAge#, ByVal OuterTol#, _
  ByVal First%, ToNextLast As Boolean)

Dim Outer%, j%, k%, Dirn%
Dim Delt#, sV3#, sVX3#, sVX23#, Sl2#, Sl2err#
Dim Mean3#, Mean3err#, dx#
ToNextLast = True

For Outer = 1 To ns Step ns - 1
  Delt = Age(Outer) - MeanAge
  ' Are the outermost steps significantly different from mean of the plateau?

  If Abs(Delt) > (OuterTol * SigmaAge(Outer)) Then Exit Sub 'GoTo NextLast

  Dirn = (Outer = ns) - (Outer = 1) ' Direction of next-to-outermost step

  If ns > 8 Then                    ' Are the wtd means of the outermost 3 steps different
    sV3 = 0: sVX3 = 0: sVX23 = 0    '  (at OuterTol sigma) from the mean? (need 9 or more steps)

    For j = Outer To Outer + 2 * Dirn Step Dirn 'OuterTol * 2 Step Dirn
      sV3 = sV3 + v(j): sVX3 = sVX3 + v(j) * Age(j)
    Next j

    Mean3 = sVX3 / sV3: Mean3err = 1 / Sqr(sV3)

    If Abs(Mean3 - MeanAge) > OuterTol * Mean3err Then Exit Sub ' GoTo NextLast
  End If

  If ns > 5 And (Outer > 1 And Outer < Nsteps) Then
    k = Outer - First + 1                     ' Is the slope defined by the outer 2 pts nonzero
    dx = MidGas(Outer + Dirn) - MidGas(Outer) '  & pointing towards center of plateau? Then

    If dx Then ' In case nonsensical input
      Sl2 = (Age(Outer + Dirn) - Age(Outer)) / dx '  reject outermost (need 6 or more steps for this test)
      Sl2err = OuterTol * Sqr((SQ(SigmaAge(Outer)) + SQ(SigmaAge(Outer + Dirn))) / (dx * dx))

      If (Abs(Sl2) - Sl2err) > 0 And Sgn(Sl2) = -Sgn(Age(Outer) - MeanAge) Then Exit Sub

    End If

  End If
Next Outer

ToNextLast = False
End Sub

Sub ArgonJ()
LoadUserForm Jinput

With Jinput
  .oPercent = Menus("jpercent"): .oAbs = Not .oPercent
  ProcJ True
  Jperr = 0: Jay = 0: Jerror = 0
  If Not ArPlat Then .eJ.Text = Menus("Jay")
  .Jer.Text = Menus("Jerror")
  .Show

  If Canceled Then ExitIsoplot

  If ArPlat Then
    Menus("ArKa") = .oKa
  Else
    Jay = Val(.eJ): Menus("Jay") = .eJ
  End If

  Jerror = Val(.Jer)
  Menus("Jerror") = .Jer
  ArKa = .oKa
  ' Jerror is at SigLev sigma!

  If ArPlat Then ' Don't Change
    Jperr = Jerror: Jerror = 0
  ElseIf .oPercent Then
    Jperr = Jerror: Jerror = Jperr / 100 * Jay
  Else
    Jperr = 100 * Jerror / Jay
  End If

End With

End Sub


Sub ArgonAge(ByVal a40_k40#, ByVal a40_k40err#, Age#, AgeErr#, _
  Optional NoJayError As Boolean = False)
' a40_k40err is at 2-sigma/95%conf

Dim test#, Numer#, Denom#, Jerr#

ViM NoJayError, False
Jerr = IIf(NoJayError, 0, Jerror) ' at SigLev sigma
test = a40_k40 * Jay + 1

If test < MINLOG Or test > MAXLOG Then
  Age = BadT
Else
  Age = Log(test) / Lambda40
  ' Calculate Numer at 1-sigma
  ' Jerr is at SigLev sigma,a 40_k40 error at 2-sigma/95%conf
  Numer = SQ(Jay * a40_k40err / 2) + SQ(a40_k40 * Jerr / SigLev)
  Denom = SQ(Lambda40 * test)
  AgeErr = 2 * Sqr(Numer / Denom) ' 2-sigma/95%
  ' An approximation because mixing 2-sigma with 95% conf!!
  ' Assuming zero correlation between measured 40A/40K and 40A/36A!
End If

End Sub

Sub PlateauCalc(ByVal Nsteps%, PlateauAge#, AgeError#, _
  FirstStep%, LastStep%, GasFract, Cap$, Gas#())

' Find an argon-argon age plateau, with the criteria that the Plateau steps must:
'  1) include 30 to 99% of the 39Ar (user selectable);
'  2) have a probability-of-fit for the weighted-mean of >0.05;
'  3) have a zero Age vs GasFract slope within error of the (error-weighted) slope
'  4) have no nonzero slope or "OuterTol"-sigma outliers at the upper & lower ends of the plateau
'  5) have at least 3 steps
'  6) steps are contiguous
' Do this by finding all possible plateaus that satisfy the above criteria,then choose the
'  plateau with the largest 39Ar fraction.
' Data in bold font, however, will be used to define a plateau no matter what.

Dim Bad As Boolean, test As Boolean, Cum As Boolean, ToNextLast As Boolean, At95 As Boolean
Dim a1$, a3a$, a6a$, a6b$, a2$, a3$, a4$, a5$, a6$, a4a$, a4b$, s$, vv$, ee$, AgeUnit As String * 3
Dim SL%, i%, j%, k%, ns%, M%, Dirn%, First%, Last%, df%, Fforced%, Lforced%, pL%(), pF%()
Dim GasMult#, tmp#, sv#, svX#, sVX2#, MeanAge#, Su#, SumGas#, Tot#, Prob#, PlatSlope#
Dim PlatSlopeErr#, OuterTol#, MSWD#, Gas0#, Sl2#, dx#, Sl2err#, Delt#
Dim CumGas#(), MidGas#(), Age#(), SigmaAge#()
Dim v#(), pT#(), pG#(), pTe#(), pMWD#(), pp#(), pS#(), pSe#()
Dim PA As Object, Lb As Object, crr As Object, dfr As Object
Dim Ob As Object, cb As Object, Grp As Object

SL = IIf(Opt.AlwaysPlot2sigma Or SigLev = 2, 2, 1)

For i = 1 To Nsteps - 2: M = M + i: Next i
' There are M possible plateaus for Nsteps step-heating steps

ReDim pT(M), pG(M), pTe(M), pMWD(M), Gas(Nsteps)
ReDim pp(M), pF%(M), pL%(M), pS(M), pSe(M)

For i = 1 To Nsteps: Gas(i) = InpDat(i, 1): Next i

If ForcedPlateau Then   ' To avoid a Student's-t * SQRT(MSWD) factor.

  For i = 1 To Nsteps

    If BoldedData(i) Then
      If Fforced = 0 Then Fforced = i
      Lforced = i
    End If

  Next i

  For i = Fforced To Lforced

    If Not BoldedData(i) Then
      MsgBox "A forced plateau must consist of contiguous steps", , Iso
        ExitIsoplot
    End If

  Next i

  If (Lforced - Fforced) < 2 Then ForcedPlateau = False
End If

OuterTol = 1.8   ' Sigma tolerance on outer steps
Cumulative Cum, Nsteps, GasMult, Gas()
i = 0

For First = 1 To Nsteps - (ArMinSteps - 1)
  If ForcedPlateau And First <> Fforced Then GoTo NextFirst

  For Last = First + (ArMinSteps - 1) To Nsteps

    If ForcedPlateau And Last <> Lforced Then GoTo NextLast
    i = i + 1

    ns = Last - First + 1
    ReDim CumGas(ns), MidGas(ns), Age(ns), SigmaAge(ns), v(ns)
    svX = 0: sv = 0: sVX2 = 0: SumGas = 0

    For j = First To Last
      k = j - First + 1
      CumGas(k) = Gas(j) / GasMult
      If j = 1 Then Gas0 = 0 Else Gas0 = Gas(j - 1) / GasMult
      MidGas(k) = Gas0 + (CumGas(k) - Gas0) / 2   ' Mean GasFract of this step
      Age(k) = InpDat(j, 2): SigmaAge(k) = InpDat(j, 3)
      v(k) = 1 / SQ(SigmaAge(k))
      SumGas = SumGas + CumGas(k) - Gas0
    Next j

    sv = Sum(v())
    svX = SumProduct(v(), Age())
    sVX2 = SumProduct(v(), Age(), Age())
    ' Insufficient 39Ar in plateau?

    If Not ForcedPlateau And SumGas < (ArMinGas / Hun) Then GoTo NextLast

    MeanAge = svX / sv
    Su = sVX2 + MeanAge * MeanAge * sv - 2 * MeanAge * svX
    df = ns - 1
    pMWD(i) = Su / df: pp(i) = ChiSquare(pMWD(i), (df))

    If Not ForcedPlateau Then
      If pp(i) < ArMinProb Then GoTo NextLast     ' Steps dissimilar at 5% probability
    End If

    CheckOuterTrends Nsteps, ns, Age(), SigmaAge(), v(), MidGas(), MeanAge, _
      OuterTol, First, ToNextLast

    If ForcedPlateau Or Not ToNextLast Then
      ' Final check -- is the plateau-slope nonzero?
      SimpleWtdRegression (ns), MidGas(), Age(), SigmaAge(), PlatSlope, PlatSlopeErr, Bad
      ' PlatSlopeErr is at 95%-conf. if Bad returned as TRUE
      test = ForcedPlateau Or (Not Bad And (Abs(PlatSlope) - 2 * PlatSlopeErr) < 0)

      If test Then  ' A valid plateau  -- add to list
        pF(i) = First:   pL(i) = Last
        pT(i) = MeanAge: pTe(i) = 1 / Sqr(sv)
        pS(i) = PlatSlope:   pSe(i) = PlatSlopeErr
        pG(i) = SumGas
        If ForcedPlateau Then Exit For
      End If

    End If

NextLast:
  Next Last
NextFirst:

Next First
GasFract = 0

For i = 1 To IIf(ForcedPlateau, 1, M)  ' Find plateau w. largest 39Ar fraction

  If ForcedPlateau Or (pT(i) <> 0 And pG(i) > GasFract) Then
    GasFract = pG(i):   MSWD = pMWD(i)
    FirstStep = pF(i):  LastStep = pL(i)
    PlateauAge = pT(i): AgeError = pTe(i) ' at 1 sigma
    PlatSlope = pS(i):  PlatSlopeErr = pSe(i)
    Prob = pp(i): At95 = False
    tmp = PlateauAge * Jperr / SigLev / 100  ' 1 sigma age-error from J error

    If ForcedPlateau Then

      If pp(i) < ArMinProb Then
        AgeError = Sqr(pMWD(i)) * StudentsT(ns - 1) * AgeError ' 95%conf, no J error
        AgeError = Sqr(AgeError ^ 2 + 4 * tmp ^ 2) ' w. J error -- approx since mixing
        At95 = True                                ' 95% conf. and 2-sigma
      Else
        AgeError = SL * Sqr(AgeError ^ 2 + tmp ^ 2) ' 1 or 2 sigma
      End If

      If Not Bad Then PlatSlopeErr = PlatSlopeErr * SL  ' 1 or 2-sigma
      Exit For

    Else
      AgeError = SL * Sqr(AgeError ^ 2 + tmp ^ 2) ' 1 or 2 sigma
    End If

  End If

Next i

AssignD "arStepAge", PA, , cb, Ob, Lb, Grp, , , , , , , , dfr

If DoPlot Then
  Grp("gFocus").Enabled = True
  Ob("oPlateau").Enabled = True: Ob("oAll").Enabled = True
  Ob("oPlateau") = ArRestricted: Ob("oAll") = Not ArRestricted
  cb("ArLines").Enabled = True
  'Cb("ArLines") = IIf(Pvlines, xlOn, xlOff)
Else
  Grp("gFocus").Enabled = False: cb("ArLines").Enabled = False
  Ob("oPlateau").Enabled = False: Ob("oAll").Enabled = False
  Ob("oPlateau") = xlOff:         Ob("oAll") = xlOff
End If

Lb("lCritProb").Text = "Probability of fit of plateau is >" & tSt(ArMinProb)
Lb("lCritSlope").Text = "No resolvable slope on plateau"
Lb("lCritOutliers").Text = "No outliers or trends at upper & lower steps"

If ForcedPlateau Then
  Lb(1).Text = "Input data are in bold font"
  Grp(2).Text = "FORCED-PLATEAU AGE RESULTS"
Else
  Lb("lCritSteps").Text = "Includes at least" & Str(ArMinGas) & _
    "% of the 39Ar in" & Str(ArMinSteps) & " or more contiguous steps"
  Grp("gResults").Text = "PLATEAU AGE RESULTS"
End If

For i = 1 To Lb.Count
  Lb(i).Visible = Not (ForcedPlateau And (i = 2 Or i = 3 Or i = 8))
Next i

If PlateauAge Then
  With Lb("lInitial"): .Text = "": .Visible = False: End With
  Grp("gResults").Width = 298.5
  AgeUnit = IIf(ArKa, " ka", " Ma")
  NumAndErr PlateauAge, AgeError, 2, vv$, ee$
  a1$ = "ge = " & vv$: a2$ = pm & ee$ & AgeUnit
  a3$ = "(" & IIf(At95, "95% conf.)", tSt(SL) & " sigma")
  a3$ = a3$ & ", " & IIf(Jperr > 0, "including J-error of" & _
    Str(Jperr) & "%", "neglecting error in J") & ")"
  a4$ = "MSWD = " & Mrnd(MSWD)
  a5$ = ProbRnd(Prob)
  i = (GasFract < 1) + (GasFract > 0.99)
  a6b$ = Sp(Hun * GasFract, i) & "% of the 39Ar"
  a6$ = "Includes " & a6b$
  Lb("lAge").Text = "A" & a1$ & a2$ & "  " & a3$
  a4a$ = ",  probability of fit = ": a4b$ = ",  probability = "
  Lb("lMSWD").Text = a4$ & a4a$ & a5$
  Lb("lSlope").Text = "Error-wtd plateau slope (" & pm & _
    "95%) = " & VandE(PlatSlope, 2 * PlatSlopeErr, 2)
  a6a$ = "steps " & sn$(FirstStep) & " through " & sn$(LastStep)
  Lb("lSteps").Text = a6$ & " (" & a6a$ & " out of " & sn$(Nsteps) & " total)"
  If DoPlot Then Cap$ = "Plateau a" & a1$ & a2$ & vbLf & a3$ & vbLf & _
    a4$ & ", probability=" & a5$ & vbLf & a6$
  cb("cShowRes").Enabled = True
Else
  Lb("lAge").Text = "DATA DO NOT DEFINE A PLATEAU"
  Lb("lMSWD").Text = "": Lb("lSlope").Text = "": Lb("lSteps").Text = ""
  cb("cshowres") = xlOff: cb("cshowres").Enabled = False
End If

Lb("lJay").Visible = False
Grp("gAgeSpectrum").Visible = False: Grp("gIsochron").Visible = False
cb("cAgeSpectrum").Visible = False: cb("cIsochron").Visible = False
cb("cInset").Visible = False:
ShowBox PA, True
ArRestricted = IsOn(Ob("oPlateau"))
Pvlines = IsOn(cb("ArLines"))

If IsOn(PA.CheckBoxes("cShowRes")) Then
  AddResBox "A" & a1$ & a2$ & vbLf & a3$ & vbLf & a4$ & a4b$ & a5$ & _
    vbLf & a6b$ & ",  " & a6a$
End If

End Sub

Sub PlateauChron(ByVal Nsteps%, ArAge#, AgeError#, PlInit#, _
  plInitErr#, FirstStep%, LastStep%, GasFract#, _
  Cap$, Gas#())

' Find an Ar-Ar isochron + age plateau, the criteria being:
'  1) Adjacent steps must define an Ar-Ar isochron with a prob-of-fit >0.05;
'  2) These ages steps must, when calculated using the isochron 36Ar/40Ar intercept,
'  3) include 30 to 99% of the 39Ar (user selectable);
'  4) have a probability of fir for the isochron of >0.05;
'  5) have a slope for Age vs GasFract of zero within error of the (error-weighted) slope
'  6) have no nonzero slope or "OuterTol"-sigma outliers at the upper & lower ends of the plateau
'  7) have at least 3 steps
'  8) steps are contiguous
' Do this by finding all possible steps that satisfy the above criteria, then choose the
'  set with the largest 39Ar fraction.
' InpDat(i,j) values for j=1 to 8
'   1: GasFract  2: Age  3: AgeErr  4: Ar39/36  5: err  6: Ar40/36  7: err  8: rho  (NORMAL)
'   1: GasFract  2: Age  3: AgeErr  4: Ar39/40  5: err  6: Ar36/40  7: err  8: rho  (INVERSE)

Dim Bad As Boolean, PlateauOK As Boolean, Cum As Boolean, ToNextLast As Boolean
Dim PlateauStepsRestrict As Boolean
Dim s$, a1$, a2$, a2a$, a3$, a4$, a4a$, a4b$, a5$, a6$, a6a$, a6b$, vv$, ee$
Dim AgeUnit As String * 3
Dim i%, j%, k%, ns%, M%, First%, Last%, SL%
Dim pF%(), pL%()
Dim GasMult#, MSWD#, Gas0#, Ar4036_0#, Ar4036_0err#, Ar40rad_Ar39err#, Ar40rad_Ar39#
Dim sv#, svX#, sVX2#, MeanAge#, SumGas#, Prob#, Slope#, SlopeErr#, OuterTol#, MaxT#
Dim Xint#, Xinterr#, PlatSlope#, PlatSlopeErr#, IsoAge#, IsoAgeErr#, ySlp#, Yinter#
Dim CumGas#(), MidGas#(), Age#(), SigmaAge#(), v#(), pT#(), pG#(), pTe#()
Dim pMWD#(), pp#(), pS#(), pSe#()
Dim aX#(), aXerr#(), aY#(), aYerr#(), aRho#(), IsoInterErr#()
Dim aSlp#(), aSlpErr#, aInter#(), aInterErr#
Dim IsoProb#(), IsoInter#(), PiAge#(), PiAgeErr#()
Dim dfr As Object, crr As Object, PA As Object, Lb As Object, Ob As Object
Dim cb As Object, Grp As Object, CBA As Object

ReDim PiAge#(Nsteps), PiAgeErr#(Nsteps)
SL = IIf(Opt.AlwaysPlot2sigma Or SigLev = 2, 2, 1)
SymbRow = Max(1, SymbRow)

For i = 1 To Nsteps - 2: M = M + i: Next i
' There are M possible plateaus for Nsteps step-heating steps

ReDim pT(M), pG(M), pTe(M), pMWD(M), pp(M), IsoProb(M), IsoInter(M), IsoInterErr(M), aSlp(M), aInter(M)
ReDim pF(M), pL(M), pS(M), pSe(M), Gas(Nsteps)
PlateauStepsRestrict = True
ArgonJ ' Requested at sigma-level of data-input

If Jay = 0 Then Exit Sub

ReDim Preserve InpDat(Nsteps, 8)

For i = 1 To Nsteps

  For j = 8 To 4 Step -1
    InpDat(i, j) = InpDat(i, j - 2)
  Next j

  InpDat(i, 2) = 0: InpDat(i, 3) = 0
Next i

If SigLev = 1 Then Jerror = Jerror / 2  ' Now 1-sigma!

For i = 1 To Nsteps: Gas(i) = InpDat(i, 1): Next i

OuterTol = 1.8 ' Sigma tolerance on outer steps
Cumulative Cum, Nsteps, GasMult, Gas()
i = 0

For First = 1 To Nsteps - (ArMinSteps - 1)

  For Last = First + (ArMinSteps - 1) To Nsteps
    i = i + 1
    ns = Last - First + 1
    ReDim aX(ns), aXerr(ns), aY(ns), aYerr(ns), aRho(ns)

    For j = First To Last
      k = j - First + 1
      aX(k) = InpDat(j, 4): aXerr(k) = InpDat(j, 5)
      aY(k) = InpDat(j, 6): aYerr(k) = InpDat(j, 7)
      aRho(k) = InpDat(j, 8)
    Next j

    ShortYork (ns), aSlp(i), aInter(i), 0, aSlpErr, aInterErr, IsoProb(i), False, _
      Bad, pMWD(i), aX(), aXerr(), aY(), aYerr(), aRho(), Xint, Xinterr

    If IsoProb(i) <= ArMinProb Or Bad Then GoTo NextLast

    ' Calculate step ages using isochron 40Ar/36Ar

    If Normal Then
      Ar4036_0 = aInter(i): Ar4036_0err = aInterErr / 1.96 '1 sigma!
      Ar40rad_Ar39 = aSlp(i): Ar40rad_Ar39err = aSlpErr / 1.96 '1 sigma!
    Else
      Ar4036_0 = 1 / aInter(i)
      Ar4036_0err = aInterErr / aInter(i) ^ 2 / 1.96 '1 sigma!
      Ar40rad_Ar39 = 1 / Xint: Ar40rad_Ar39err = Xinterr / Xint ^ 2 / 1.96 '1 sigma!
    End If

    IsoInter(i) = Ar4036_0: IsoInterErr(i) = Ar4036_0err
    ReDim CumGas(ns), MidGas(ns), Age(ns), SigmaAge(ns), v(ns), StepAge(ns)
    svX = 0: sv = 0: sVX2 = 0: SumGas = 0

    For j = First To Last
      k = j - First + 1
      CumGas(k) = Gas(j) / GasMult

      If j = 1 Then Gas0 = 0 Else Gas0 = Gas(j - 1) / GasMult

      MidGas(k) = Gas0 + (CumGas(k) - Gas0) / 2   ' Mean GasFract of this step
      SumGas = SumGas + CumGas(k) - Gas0
      CalcArgonStepAge Age(k), SigmaAge(k), Ar4036_0, j
      If Age(k) = BadT Then GoTo NextLast
      v(k) = 1 / SQ(SigmaAge(k))
    Next j

    sv = Sum(v())
    svX = SumProduct(v(), Age())
    sVX2 = SumProduct(v(), Age(), Age())
    ' Insufficient 39Ar in plateau?
    If SumGas < (ArMinGas / Hun) Then GoTo NextLast
    MeanAge = svX / sv

    CheckOuterTrends Nsteps, ns, Age(), SigmaAge(), v(), MidGas(), MeanAge, OuterTol, First, ToNextLast

    If Not ToNextLast Then
      ' Final check -- is the plateau-slope nonzero?
      SimpleWtdRegression (ns), MidGas(), Age(), SigmaAge(), PlatSlope, PlatSlopeErr, Bad
      ' SlopeErr is at 95%-conf. if Bad returned as TRUE
      PlateauOK = (Not Bad And (Abs(PlatSlope) - 2 * PlatSlopeErr) < 0)

      If PlateauOK Then  ' A valid plateau  -- add to list
        ArgonAge Ar40rad_Ar39, Ar40rad_Ar39err, IsoAge, IsoAgeErr  ' 1 sigma, incl. +-J
        pF(i) = First:     pL(i) = Last
        pT(i) = IsoAge:    pTe(i) = IsoAgeErr
        pS(i) = PlatSlope: pSe(i) = PlatSlopeErr
        pG(i) = SumGas
      End If

    End If

NextLast:
  Next Last

NextFirst:
Next First

GasFract = 0

For i = 1 To M  ' Find plateau w. largest 39Ar fraction

  If pT(i) <> 0 And pG(i) > GasFract Then
    GasFract = pG(i):  MSWD = pMWD(i)
    FirstStep = pF(i): LastStep = pL(i)
    ArAge = pT(i):     AgeError = pTe(i)
    PlatSlope = pS(i): PlatSlopeErr = pSe(i)
    Prob = IsoProb(i): ySlp = aSlp(i): Yinter = aInter(i)
    Ar4036_0 = IsoInter(i): Ar4036_0err = IsoInterErr(i)

    For j = 1 To Nsteps
      CalcArgonStepAge PiAge(j), PiAgeErr(j), Ar4036_0, j
    Next j

  End If

Next i

AssignD "arStepAge", PA, , cb, Ob, Lb, Grp, , , , , , , , dfr
Set crr = Grp("gAgeSpectrum"): Set CBA = cb("ArLines")
If ArAge = 0 Then DoPlot = False
With cb("cAgeSpectrum"): .Visible = True: .Enabled = DoPlot: End With
With Grp("gIsochron"): .Visible = True: .Enabled = DoPlot: End With
With crr: .Visible = True: .Enabled = DoPlot: End With
With cb("cIsochron"): .Visible = True: .Enabled = DoPlot: End With
With cb("cInset"): .Visible = True: .Enabled = DoPlot: End With
With Grp("gFocus"): .Visible = True: .Enabled = DoPlot: End With
cIsochron_Click
Ob("oPlateau").Enabled = DoPlot: Ob("oAll").Enabled = DoPlot

If DoPlot Then
  Ob("oPlateau") = ArRestricted: Ob("oAll") = Not ArRestricted
  CBA.Enabled = True
  'CBA = IIf(Pvlines, xlOn, xlOff)
Else
  CBA.Enabled = False
  Ob("oPlateau") = xlOff: Ob("oAll") = xlOff
End If

Lb("lCritProb").Text = "Probability of fit of Isochron for Plateau Steps is >" & tSt(ArMinProb)
Lb("lCritSlope").Text = "No resolvable slope on plateau"
Lb("lCritOutliers").Text = "No outliers or trends at upper & lower steps"
Lb("lCritSteps").Text = "Includes at least" & Str(ArMinGas) & "% of the 39Ar in" & _
  Str(ArMinSteps) & " or more contiguous steps"
Grp(2).Text = "Ar Plateau-Isochron Results"

With Lb("lJay"): .Visible = True: .Enabled = True: .Text = "(at " & AtJay$ & ")": End With

If ArAge Then
  cb("cshowres").Enabled = True: Grp("gResults").Width = 325.5: AgeUnit = " Ma"
  MaxT = -1E+99

  For i = 1 To Nsteps
    InpDat(i, 2) = PiAge(i): InpDat(i, 3) = PiAgeErr(i)
    MaxT = Max(MaxT, PiAge(i))
  Next i

  ArKa = (MaxT <= 1)

  If ArKa Then ArAge = ArAge * Thou: AgeError = AgeError * Thou: AgeUnit = " ka"

  NumAndErr ArAge, AgeError * SL, 2, vv$, ee$
  a1$ = "ge = " & vv$: a2$ = pm & ee$ & AgeUnit
  a2a$ = "(" & AtJay$ & ")": a3$ = "(" & sn$(SL) & " sigma)"
  a4$ = "MSWD = " & Mrnd(MSWD)
  a5$ = ProbRnd(Prob)
  i = (GasFract < 1) + (GasFract > 0.99)
  a6b$ = Sp(Hun * GasFract, i) & "% of the 39Ar" ' (steps" & Str(FirstStep) & "-" & tSt(LastStep) & ")"
  a6$ = "Includes " & a6b$
  Lb("lAge").Text = "A" & a1$ & a2$ & "  " & a3$
  NumAndErr Ar4036_0, Ar4036_0err * SL, 2, vv$, ee$
  With Lb("lInitial"): .Visible = True: .Text = "Initial 40/36 = " & vv$ & pm & ee$: End With
  a4a$ = ",  probability of fit = ": a4b$ = ",  probability = "
  Lb("lMSWD").Text = a4$ & a4a$ & a5$
  Lb("lSlope").Text = "Error-wtd plateau slope (" & pm & "95%) = " & VandE(PlatSlope, 2 * PlatSlopeErr, 2)
  a6a$ = "steps " & sn$(FirstStep) & " through " & sn$(LastStep)
  Lb("lSteps").Text = a6$ & " (" & a6a$ & " out of " & sn$(Nsteps) & " total)"

  If DoPlot Then Cap$ = "Plateau-Isochron a" & a1$ & a2$ & "  " & a3$ & vbLf & _
    Lb("lInitial").Text & vbLf & a4$ & ", probability=" & a5$ & vbLf & a6$

  Crs(1) = ySlp: Crs(3) = Yinter
  ArChronSteps(1) = FirstStep: ArChronSteps(2) = LastStep

Else
  Lb("lAge").Text = "DATA DO NOT DEFINE A PLATEAU-ISOCHRON"
  With Lb("lInitial"): .Text = "": .Visible = False: End With

  For i = 5 To 7: Lb(i).Text = "": Next i

  With cb("cshowres"): .Value = xlOff: .Enabled = False: End With
  cb("cIsochron") = xlOff: cb("cAgeSpectrum") = xlOff: cb("cInset") = xlOff
  cb("cIsochron").Enabled = False: cb("cAgeSpectrum").Enabled = False: cb("cInset").Enabled = False
End If

ShowBox PA, True
ArRestricted = IsOn(Ob("oPlateau")): Pvlines = IsOn(CBA)
ArIso = IsOn(cb("cIsochron")): ArSpect = IsOn(cb("cAgeSpectrum"))
ArInset = IsOn(cb("cInset"))
If Not ArIso And Not ArSpect Then DoPlot = False

If IsOn(PA.CheckBoxes("cShowRes")) Then
  AddResBox "A" & a1$ & a2$ & " " & a3$ & vbLf & Lb("lInitial").Text & _
    vbLf & a4$ & a4b$ & a5$ & vbLf & a6b$ & ",  " & a6a$ & vbLf & a2a$
End If

End Sub

Sub PlotArSteps(ByVal Nsteps%, ByVal First%, ByVal Last%, _
  ByVal Page#, ByVal Cap$, Gas#())
Attribute PlotArSteps.VB_ProcData.VB_Invoke_Func = " \n14"

Dim GasPercent As Boolean, tB As Boolean, sW As Boolean
Dim ct$, Acap$, Ylab$, tmp$
Dim i%, j%, k%, BoxCol%, Pcol%, Ymult%, SL%
Dim Bclr&
Dim y1#, y2#, ey1#, ey2#, MinT#, MaxT#, Tspred#, Tik#
Dim rr1#, rr2#, MinA#, MaxA#, eY#, Bdist#, MedEr#, MedAge#
Dim T#(), Terr#(), Boxx#(5, 2), Ar#(2, 2)
Dim rBoxRange() As Range, rLr As Range
Dim TxtBox As Object, ArSc As Object
Dim NWS As Variant, eAdd As Variant, Clr As Variant
Dim PlatLine As Variant, Transp As Variant, vBoxRange As Variant

SL = IIf(Opt.AlwaysPlot2sigma Or SigLev = 2, 2, 1)
SymbRow = Max(1, SymbRow)
StatBar "Constructing plotbox"
Ymult = 1
ReDim rBoxRange(Nsteps), T(Nsteps), Terr(Nsteps), Gas(Nsteps)

For i = 1 To Nsteps
  T(i) = InpDat(i, 2): Terr(i) = InpDat(i, 3)
  Gas(i) = InpDat(i, 1)
Next i

Cumulative tB, Nsteps, 1, Gas()

If ArChron And ArKa Then
  Ymult = Thou

  For i = 1 To N
    T(i) = Ymult * T(i): Terr(i) = Ymult * Terr(i)
  Next i

End If

' Construct the plot of step age vs cumulative Ar-39 fraction for an
'  Ar-Ar step-heating diagram
MedAge = iMedian(T())

If ArRestricted Then
  MedEr = iMedian(Terr())
  eAdd = 2 * MedEr * SL
Else
  eAdd = SL * App.Max(Terr)
End If

GasPercent = (Gas(Nsteps) > 10)
BoxCol = SymbCol

For i = 1 To Nsteps
  eY = SL * Terr(i)
  Boxx(1, 2) = T(i) - eY: Boxx(2, 2) = T(i) - eY
  Boxx(3, 2) = T(i) + eY: Boxx(4, 2) = T(i) + eY
  Boxx(5, 2) = Boxx(1, 2)

  If i > 1 Then Boxx(1, 1) = Gas(i - 1) Else Boxx(1, 1) = 0

  Boxx(2, 1) = Gas(i): Boxx(3, 1) = Boxx(2, 1)
  Boxx(4, 1) = Boxx(1, 1):   Boxx(5, 1) = Boxx(1, 1)
  j = (i - 1) * 6 + SymbRow:       k = j + 4
  Set rBoxRange(i) = sR(j, BoxCol, k, 1 + BoxCol, ChrtDat)
  LineInd rBoxRange(i), "ErrBox"
  rBoxRange(i).Value = Boxx
Next i

AddSymbCol 2

If Page Then
  Pcol = SymbCol
  Ar(1, 1) = 0:  Ar(1, 2) = Page:   Ar(2, 2) = Page
  Ar(2, 1) = IIf(GasPercent, Hun, 1)
  Set PlatLine = sR(SymbRow, Pcol, 1 + SymbRow, 1 + Pcol, ChrtDat)
  PlatLine.Value = Ar
  AddSymbCol 2
End If

MaxT = -1E+99: MinT = 1E+99

For i = 1 To Nsteps
  If Not ArRestricted Or (i >= First And i <= Last) Then
    MinT = Min(MinT, T(i)) ' - eAdd * terr(i))
    MaxT = Max(MaxT, T(i)) ' + eAdd * terr(i))
  End If
Next i

MaxT = Min(MaxT, MedAge * 2)
MinT = MinT - eAdd: MaxT = MaxT + eAdd
MinT = Max(0, MinT): Tspred = MaxT - MinT
MinT = Max(0, MinT - Tspred * 0.12)
MaxT = Min(2 * MedAge, MaxT + Tspred * 0.12)
Tspred = MaxT - MinT
Tick Tspred, Tik
MinA = -Tik

Do
  MinA = MinA + Tik
Loop Until (MinA + Tik) > MinT

' If next-lower MinA has requires fewer alpha chars, use it
If (NumChars(MinA - Tik) < NumChars(MinA)) Then MinA = MinA - Tik

If NumChars(MinA - Tik) = NumChars(MinA) Then
  ' If last # of MinA is odd & next-lower is even, use next-lower
  rr1 = Val(Right$(sn$(MinA - Tik), 1))
  rr2 = Val(Right$(sn$(MinA), 1))
  If rr1 Mod 2 = 0 And rr2 Mod 2 = 1 Then MinA = MinA - Tik
End If

MinT = MinA: MaxA = MinA

Do
  MaxA = MaxA + Tik
Loop Until MaxA > MaxT

MaxT = MaxA
Charts.Add
PlotName$ = "ArStepHeat"
ct$ = "Cumulative 39Ar "
ct$ = ct$ & IIf(GasPercent, "Percent", "Fraction")
MakeSheet PlotName$, IsoChrt    ' Use the data defining the first step's box
Landscape
Ylab$ = IIf(ArKa, "ka)", "Ma)")
Set vBoxRange = rBoxRange(1)

Ach.ChartWizard vBoxRange, xlXYScatter, 1, xlColumns, 1, 0, 2, "", ct$, "Age (" & Ylab$

With IsoChrt
  Set ArSc = .SeriesCollection
  .ChartArea.Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.SheetClr, vbWhite))
  tB = (ArChron And ArInset)

  With .PlotArea
   .Height = 375 - 50 * tB: .Top = 35 + 25 * tB
   .Width = 460 - 90 * tB: .Left = 110 + 35 * tB
  End With

  With .Axes(xlValue)
    .MajorTickMark = IIf(Opt.AxisTickCross, xlCross, xlInside)
    .MinorTickMark = xlInside
    .TickLabelPosition = xlNextToAxis
    .MinimumScale = MinT:  .MaximumScale = MaxT
    .MinorUnit = 0.000000001
    .MajorUnit = 2 * Tik: .MinorUnit = Tik:
    .ReversePlotOrder = False: .ScaleType = False
    .TickLabels.NumberFormat = TickFor(MinT, MaxT, Tik)
    .MinorUnitIsAuto = Opt.AxisAutoTikSpace: .MajorUnitIsAuto = Opt.AxisAutoTikSpace

    With .Border
      .Color = Black: .Weight = AxisLthick: .LineStyle = xlContinuous
    End With

    With .TickLabels.Font
      .Name = Opt.AxisTikLabelFont: .Size = Opt.AxisTikLabelFontSize
      .Background = xlTransparent
    End With

  End With

  With .Axes(xlCategory)
    .MajorTickMark = IIf(Opt.AxisTickCross, xlCross, xlInside)
    .MinorTickMark = xlInside: .TickLabelPosition = xlNextToAxis
    .ReversePlotOrder = False: .ScaleType = False
    .MinorUnitIsAuto = False:  .MajorUnitIsAuto = False
    Superscript Phrase:=.AxisTitle

    With .Border
      .Color = Black: .Weight = AxisLthick: .LineStyle = xlContinuous
    End With

    With .TickLabels.Font
      .Name = Opt.AxisTikLabelFont: .Size = Opt.AxisTikLabelFontSize: .Background = xlTransparent
    End With

    .MinorUnit = 0.000000001: .MinimumScale = 0

     If GasPercent Then
       .MaximumScale = Hun: .TickLabels.NumberFormat = "0"
      .MajorUnit = 20:      .MinorUnit = 10
     Else
      .MaximumScale = 1: .TickLabels.NumberFormat = "0.0"
      .MajorUnit = 0.2:  .MinorUnit = 0.1:
    End If

  End With

  With .Axes(xlValue).AxisTitle.Characters.Font
    .Name = Opt.AxisNameFont: .Size = Opt.AxisNameFontSize: .Background = xlTransparent
  End With
  With .Axes(xlCategory).AxisTitle.Characters.Font
    .Name = Opt.AxisNameFont: .Size = Opt.AxisNameFontSize: .Background = xlTransparent
  End With

  With .PlotArea.Border
    .Weight = AxisLthick: .LineStyle = xlContinuous: .Color = vbBlack
  End With

  .PlotArea.Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.PlotboxClr, vbWhite))
End With

RemoveHdrFtr IsoChrt
NWS = 1

If Page Then  ' Add the plateau-age line if defined
  IsoChrt.SeriesCollection.Add PlatLine, xlColumns, False, 1, False

  If IsoChrt.SeriesCollection.Count = NWS + 1 Then
    NWS = 1 + NWS

    With IsoChrt.SeriesCollection(NWS)
      .MarkerStyle = xlNone

      With .Border
        If ColorPlot And Not DoShape Then
          .Color = Menus("cAqua"): .Weight = xlMedium
        ElseIf ColorPlot Then
          .Color = vbBlue: .Weight = xlMedium
        Else
          .Color = vbBlack: .Weight = xlMedium
        End If
        .LineStyle = xlContinuous
      End With

    End With

  End If

End If

Transp = IIf(Page <> 0, 0.5 - 0.25 * (ArChron And ArInset And Not ColorPlot), 0)
StatBar "Adding Ar steps"
If DoShape Then GetScale

For i = 1 To Nsteps
  Bclr = Black

  If Regress Then

    If i < First Or i > Last Then

      If ColorPlot Then
        Bclr = IIf(DoShape, Menus("cCyan"), vbBlue)
      ElseIf DoShape Then
        Bclr = White
      End If

    Else

      If ColorPlot Then
        Bclr = IIf(DoShape, Menus("cPink"), vbRed)
      ElseIf DoShape Then
        Bclr = Menus("cGray60")
      End If

    End If

  Else

    If ColorPlot Then
      Bclr = IIf(DoShape, Menus("cPink"), Menus("cDkRed"))
    ElseIf DoShape Then
      Bclr = Menus("cGray60")
    End If

  End If

  If DoShape Then
    If i = 1 Then IsoChrt.SeriesCollection(1).MarkerStyle = xlNone
  ElseIf i > 1 Then
    IsoChrt.SeriesCollection.Add rBoxRange(i), xlColumns, False, 1, False
  End If

  If IsoChrt.SeriesCollection.Count = NWS + 1 Or i = 1 Then

    If i > 1 Then NWS = 1 + NWS: j = NWS Else j = 1

    If DoShape Then

      If rBoxRange(j)(1, 1) >= MaxY Then
        AddShape "ErrBox", rBoxRange(j), Bclr, Black, 0, 0, , , Transp
      Else
        AddShape "ErrBox", rBoxRange(i), Bclr, Black, False, 0, , , Transp
      End If
    Else
      IsoChrt.SeriesCollection(j).MarkerStyle = xlNone

      With IsoChrt.SeriesCollection(j).Border
        .Weight = xlThin: .Color = Bclr: .LineStyle = xlContinuous

        If Not ColorPlot And i >= First And i <= Last Then
          .Weight = IIf(ArChron And ArInset, xlThick, xlMedium)
        End If

      End With

    End If

  ElseIf DoShape Then
    AddShape "ErrBox", rBoxRange(i), Bclr, Black, False, 0, , , Transp
  End If

Next i

If Pvlines Then ' Add connecting lines to boxes if needed

  For i = 1 To N - 1
    j = Sgn(T(i + 1) - T(i))
    ey1 = SL * Terr(i): ey2 = SL * Terr(i + 1)
    y1 = T(i) + j * ey1: y2 = T(i + 1) - j * ey2
    Bdist = (y2 - y1) * j

    If Bdist > 0 Then

      With ChrtDat
        .Cells(SymbRow, SymbCol) = Str(CSng(Gas(i)))
        .Cells(1 + SymbRow, SymbCol) = .Cells(1, SymbCol)
        .Cells(SymbRow, 1 + SymbCol) = Str(CSng(y1 - j * ey1))
        .Cells(1 + SymbRow, 1 + SymbCol) = Str(CSng(y2 + j * ey2))
      End With

      Set rLr = sR(SymbRow, SymbCol, 1 + SymbRow, 1 + SymbCol, ChrtDat)
      LineInd rLr, "ShapeLine"
      AddSymbCol 2

      If DoShape Then
        AddLine rLr, PlotDat$ & Und & tSt(rLr.Column) & "|1~2_0%"

        With Selection.Border
          .Weight = xlHairline: .LineStyle = xlContinuous
        End With

      Else
        LineInd rLr, "IsoLine"
        IsoChrt.SeriesCollection.Add rLr, xlColumns, False, 1, False
        Nser = IsoChrt.SeriesCollection.Count

        With IsoChrt.SeriesCollection(Nser)
          .MarkerStyle = xlNone

          With .Border
            .Color = vbBlack: .Weight = xlHairline
            .LineStyle = xlContinuous
          End With

        End With

      End If

    End If

  Next i

End If

If Page Then AddArTextbox Cap$, IIf(ArChron And ArIso And ArInset, 10, 11)
AddArRejSymbNote 1
AddCopyButton
If ArChron And ArSpect And Not WasPlat Then PutPlotInfo
IsoChrt.Deselect: ChrtDat.Visible = False
End Sub

Sub ArgonMonteCarlo(ByVal Ntrials&, AgeLine$, InterLine$)
Attribute ArgonMonteCarlo.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Bad As Boolean
Dim i%, j%, k%
Dim Age$, eR$(2), Yinter$, YintrEr$(2)
Dim JageErr#, Ar4039#, Slp#, te#, TT#, Yint#, YintErr#
Dim T#(2), ti#(), s#(), xi#()
Dim L As Object, Yi() As Variant

Set L = DlgSht("IsoRes").Labels
ReDim s(6), Yi(6), xi(6)
MCyorkfit N, Ntrials, Bad, s(), Yi(), xi()

If Bad Then Exit Sub

With yf

  If Normal Then
    Yint = .Intercept: Slp = .Slope
  Else
    Yint = 1 / .Intercept: Slp = 1 / .Xinter
  End If

  ArgonAge Slp, 0, TT, JageErr, False

  For i = 4 To 5
    j = i - 3: k = 5 + (i = 5)

    If Normal Then
      Ar4039 = Slp - Sgn(j - 1.5) * s(k)
      YintErr = Yi(k)
    Else
      Ar4039 = 1 / (.Xinter + Sgn(j - 1.5) * xi(i))
      YintErr = Yi(i) * SQ(Yint)
    End If

    ArgonAge Ar4039, 0, T(j), 0, True
    te = Sqr(SQ(TT - T(j)) + SQ(JageErr))
    NumAndErr TT, te, 3, Age$, eR$(j)
    NumAndErr Yint, YintErr, 3, Yinter$, YintrEr$(j)
  Next i

  AgeLine$ = Age$ & "  +" & eR$(1) & " -" & eR$(2)
  InterLine$ = Yinter$ & "  +" & YintrEr$(1) & " -" & YintrEr$(2)
  L("lAgeLabel").Text = AgeLine$
  L("lInterLabel").Text = InterLine$
End With

End Sub
