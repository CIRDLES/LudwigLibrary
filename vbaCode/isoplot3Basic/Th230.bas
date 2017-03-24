Attribute VB_Name = "Th230"
' Isoplot module 230Th
Option Private Module
Option Explicit: Option Base 1
Dim Exp230#, Exp234#, ExpDiff#, L0term#, L4term#

Sub ShowUiso(ByVal MC As Boolean, ByVal WLE As Boolean, ByVal MSWD, ByRef Prob, _
  ByVal ThU, ByVal ThUerr, ByVal Gamma, ByVal GammaErr, ByVal Rho, _
  ByRef Emult, ByVal ThUage, ByVal AgeErr, ByVal UageErr, ByVal LageErr, _
  ByVal Gamma0, ByVal Gamma0err, ByVal Ug0err, ByVal Lg0err, ByVal RhoTG0, _
  ByVal Detr02_nought, ByVal Detr02_noughtErr)
' Show results of 230Th/U age calculation.
'  Input errors are at 1-sigma a priori, conversion to 95%_conf, via Emult!
Dim u As Object, L As Object, G As Object, cb As Object, tmp$, d As Object, A As String * 7
Dim tB As Object, i%, Age#, Iso2 As Boolean, Iso3 As Boolean, Isoc As Boolean
Dim f$, s$, T$, s1$, s2$, s3$, Fin As Object, Finite As Boolean, Cns$, Detr02zt#, Detr02errzt#
Dim X As String * 3, df$, WithDth As Boolean, Boo As Boolean, vv$, ee$, TT$, ShR As Object

WithDth = NIM(Detr02_nought): A = "l1sigma"
If Emult = 0 Then Emult = 2

If WithDth Then
  If Detr02_nought = 0 Then WithDth = False
End If

AssignD "Uiso", u, , cb, , L, G, tB, Dframe:=d

Set Fin = DlgSht("ThUage").CheckBoxes("Finite")
Finite = (MC And Fin.Enabled And Fin.Visible And IsOn(Fin))
Cns$ = "(solutions constrained to positive, finite ages)"
f$ = IIf(Mac, "Geneva", "Arial")
Isoc = (N > 1): Iso2 = (Isoc And Not Dim3): Iso3 = (Isoc And Dim3)
df$ = IIf(Iso3, "3D U-Series Isochron", "230Th/U Age")
If N = 1 Then df$ = df$ & " and Initial Ratio"
If Finite Then df$ = df$ & "   " & Cns$
d.Text = df$
X = "---"

'If Not Mac Then

  For i = 1 To tB.Count
    tB(i).Font.Name = f$
    tB(i).Font.Size = 10
  Next i

'End If

Boo = (Not MC Or ThUage <> 0) '(Not MC Or (ThUage <> 0 And UageErr <> 0))
tB("l95c").Visible = Boo
L("ErrCorrel").Visible = (Boo And Not Iso2)

If Isoc Then
  tB("l95c").Text = "95%-conf."

  If Not MC Then
    tB("l" & A).Text = "1-sigma int."
  End If

ElseIf MC Then
  tB("l95c").Text = "95% conf."
Else
  tB("l95c").Text = "2-sigma"
  tB("l" & A).Text = "1-sigma"
End If

cb("cAddWtdResids").Visible = (N > 1)
L("RhoAgeG0").Visible = (Boo And Not Iso2)
L("AuthRho").Visible = Not Iso2
L("AuthErrCorr").Visible = Not Iso2
u.Buttons("bDetails").Visible = Iso2
If IM(Prob) Then Prob = 1
Boo = IIf(Isoc, (Prob >= 0.05), True)

L(A & "Age").Enabled = Boo
L(A & "Gamma").Enabled = Boo
L(A & "ThU").Enabled = Boo
L(A & "Gamma0").Enabled = Boo
tB("l" & A).Visible = Not MC
L(A & "age").Visible = Not MC
L(A & "gamma0").Visible = Not MC
L(A & "gamma").Visible = Not MC
L(A & "thu").Visible = Not MC

If Isoc Then
  L("lN").Text = sn$((N)): L("lMSWD").Text = Mrnd(MSWD)
  Msw$ = L("lMSWD").Text
  L("lProb").Text = ProbRnd(Prob)

  If Not MC Then
    L(A & "Gamma").Text = ErFo(Gamma, GammaErr, 2, True)
    L(A & "ThU").Text = ErFo(ThU, ThUerr, 2, True)
    L(A & "Gamma0").Text = ErFo(Gamma0, Gamma0err, 2, True)
  End If

End If

If ThUage = 0 Then

  If MC And LageErr <> 0 Then
    L("Age").Text = ">" & tSt(Prnd(LageErr, 0))
    L("Gamma0").Text = "<" & tSt(Drnd(Ug0err, 4))
    L("Gamma0err").Text = "    " & qq
    L("RhoAgeG0").Text = ""
    L("AgeErr").Text = "at 95% conf."
  Else
    L("AgeErr").Text = X
    L("Age").Text = X
    L(A & "Age").Text = X
    L("Gamma0").Text = X
    L("Gamma0err").Text = X
    L("RhoAgeG0").Text = X
  End If

Else

  If AgeErr > 0 Then
    NumAndErr ThUage, AgeErr * Emult, 2, vv$, ee$, , True
    L("Age").Text = vv$:
    If Not MC Then L("AgeErr").Text = ee$

    If Iso2 Then
      L("Gamma0").Text = "1": L("Gamma0err").Text = pm & "0"
    Else
      NumAndErr Gamma0, Gamma0err * Emult, 2, vv$, ee$, , True
      L("Gamma0").Text = vv$
      If Not MC Then L("Gamma0err").Text = ee$
    End If

  Else
    L("Age").Text = Sp(ThUage, 0)
    If Not Iso2 Then L("Gamma0").Text = Sd$(Gamma0, 4)
  End If

  Uir$ = L("Age").Text
  TT$ = Uir$

  If MC Then
    If UageErr = 0 Then T$ = "+inf. " Else T$ = Sd$(UageErr, 2, -1, -1) & " "
    If LageErr = 0 Then T$ = T$ & "-inf." Else T$ = T$ & Sd$(LageErr, 2, -1, -1)
    L("AgeErr").Text = T$

    If Not Iso2 Then
      If Ug0err = 0 Then T$ = "+inf. " Else T$ = Sd$(Ug0err, 2, -1, -1) & " "
      If Lg0err = 0 Then T$ = T$ & "-inf." Else T$ = T$ & Sd$(Lg0err, 2, -1, -1)
      L("Gamma0Err").Text = T$
    End If

  ElseIf AgeErr = 0 Then
    L("AgeErr").Text = "inf."
    If Not Iso2 Then L(A & "Age").Text = "inf."
    Uir$ = Uir$ & " " & pm & "inf."
  Else
    L("AgeErr").Text = ErFo(ThUage, AgeErr * Emult, 2, True)
    If Not Iso2 Then L("Gamma0Err").Text = ErFo(Gamma0, Gamma0err * Emult, 2, True)
  End If

  Uir$ = Uir$ & " " & L("AgeErr").Text & " ka"
  If Not MC Then L(A & "Age").Text = ErFo(ThUage, AgeErr, 2, True)
  If Not Iso2 Then L("RhoAgeG0").Text = RhoRnd(RhoTG0)
End If

If Iso2 Then
  L("AuthGamma").Text = "1": L("AuthGammaerr").Text = pm & "0"
Else
  NumAndErr Gamma, GammaErr * Emult, 2, vv$, ee$, , True
  L("AuthGamma").Text = vv$:    L("AuthGammaerr").Text = ee$
  L("AuthRho").Text = RhoRnd(Rho)
End If

NumAndErr ThU, Emult * ThUerr, 2, vv$, ee$, , True

L("AuthThU").Text = vv$:   L("AuthThUerr").Text = ee$
G("gfit").Visible = Isoc:  G("g1sigma").Visible = Not MC
L("lln").Visible = Isoc:   L("llmswd").Visible = Isoc
L("ln").Visible = Isoc:    L("lmswd").Visible = Isoc
L("lprob").Visible = Isoc: L("llprob").Visible = Isoc

tB("l" & A).Visible = Not MC 'Isoc

L("Detr02_0").Visible = WithDth: L("Detr02err_0").Visible = WithDth
L("DetrTh_0").Visible = WithDth: L("l1sigdetr02err_0").Visible = (WithDth And Not MC)
L("Detr02_t").Visible = WithDth: L("Detr02err_t").Visible = WithDth
L("DetrTh_t").Visible = WithDth: L("l1sigdetr02err_t").Visible = (WithDth And Not MC)
G(5).Visible = WithDth

If WithDth Then
  NumAndErr Detr02_nought, Detr02_noughtErr * Emult, 2, vv$, ee$, , True
  L("Detr02_0").Text = vv$
  L("Detr02err_0").Text = ee$
  Detr02zt = Detr02_nought * Exp(Lambda230 * ThUage * Thou)
  Detr02errzt = Detr02zt * Sqr(SQ(Detr02_noughtErr / Detr02_nought) + SQ(Lambda230 * AgeErr * Thou))
  NumAndErr Detr02zt, Detr02errzt * Emult, 2, vv$, ee$, , True
  L("Detr02_t").Text = vv$: L("Detr02err_t").Text = ee$
  L("DetrTh_t").Text = "Detr. 230Th/232Th (at " & TT$ & " ka)"

  If Not MC Then
    L("l1sigdetr02err_0").Text = ErFo(Detr02_nought, Detr02_noughtErr, 2, True)
    L("l1sigdetr02err_t").Text = ErFo(Detr02zt, Detr02errzt, 2, True)
  End If

End If

ShowBox u, True

Set ShR = DlgSht("uiso").CheckBoxes("cShowRes")
If ShR <> xlOn And Not DoPlot Then Exit Sub
s1$ = "230Th/U Age = " & L("Age").Text
Boo = (Not MC Or (ThUage <> 0 And (UageErr <> 0 Or LageErr <> 0)))
If Boo Then s1$ = s1$ & " " & L("AgeErr").Text
s1$ = s1$ & "  ka"
s2$ = "  ("
If MC Then s2$ = s2$ & "Monte Carlo, "
s2$ = s2$ & "dce)"
s3$ = ""

If Not Iso2 Then
  s3$ = vbLf & L("Initial234238").Text & " = " & L("Gamma0").Text
  Boo = (Not MC Or (Gamma0 <> 0 And (Ug0err <> 0 Or Lg0err <> 0)))
  If Boo Then s3$ = s3$ & " " & L("Gamma0Err").Text
End If

If Isoc Then s3$ = s3$ & vbLf & "MSWD = " & L("lMSWD").Text & _
  ", probability = " & L("lProb").Text

AgeRes$ = s1$ & s3$
s$ = s1$ & s2$ & s3$
If ShR <> xlOn Then Exit Sub

If Isoc Then
  s$ = s$ & "  on " & L("lN").Text & " points"

  s$ = s$ & vbLf & L("Auth230238").Text & " = " & L("AuthThU").Text & _
       " " & L("AuthThUerr").Text & vbLf & L("Auth234238").Text & " = " & _
       L("AuthGamma").Text & " " & L("AuthGammaErr").Text

  If Not Iso2 Then s$ = s$ & vbLf & L("Initial234238").Text & " = " & _
    L("Gamma0").Text & " " & L("Gamma0Err").Text & vbLf

  If WithDth Then
    s$ = s$ & "Detr. 230/232 = " & L("Detr02_0").Text & " " & _
      L("Detr02err_0").Text & vbLf
  End If

  If Not Iso2 Then s$ = s$ & "Rho(230/238-234/238) = " & L("AuthRho").Text
End If

If Not Iso2 And L("RhoAgeG0").Text <> "" Then
  s$ = s$ & vbLf & "Rho(Age-Initial234/238) = " & L("RhoAgeG0").Text
End If

If Finite Then s$ = Cns$ & vbLf & s$
AddResBox s$, WLE:=WLE
End Sub

Function GammaU(ByVal Age#, ByVal Gamma0#)  ' returns initial 234/238, Age in years
Attribute GammaU.VB_ProcData.VB_Invoke_Func = " \n14"
GammaU = 1 + (Gamma0 - 1) * Exp(-Lambda234 * Age)
End Function

Sub PlotUseriesEvolution(Cv As Curves, ByVal CurveColor&, _
 ByVal CvSty%, ByVal CvWt%, SerC As Object)
Attribute PlotUseriesEvolution.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, MaxDecPts%, T#, xx#, yy#
Dim s$, Slp#, Inr#, LastXx#, LastYy#
Dim sxx!, syy!
StatBar "Adding evolution curves"
For i = 1 To Ncurves
  ChrtDat.Select
  If i > 1 Then
    ' Recalc agemin/agemax/agestep for extra U-series curves
    StoreCurveData i, Cv
    If Cv.NcurvEls(i) = 0 Then GoTo NextCurve
    Set SerC = IsoChrt.SeriesCollection
    SerC.Add CurvRange(i), xlColumns, False, True, False
    Set SerC = IsoChrt.SeriesCollection
    IsoChrt.Select
    With Last(SerC)
      .MarkerStyle = xlNone
      With .Border
        If ColorPlot Then
          .LineStyle = CvSty: .Weight = CvWt
          .Color = CurveColor
        Else
          .Weight = xlThick: .LineStyle = xlGray50
          .Color = vbBlack
        End If
      End With
    End With
  Else
    IsoChrt.Select
  End If
  If uUseTiks And i > 1 Then
    IsoChrt.SeriesCollection.Add TikRange(i), xlColumns, False, True, False
    Set SerC = IsoChrt.SeriesCollection
    With Last(SerC)
      .Border.LineStyle = xlNone
      .MarkerForegroundColor = IIf(ColorPlot, Opt.AgeTikSymbClr, vbBlack)
      If ColorPlot Then
        .MarkerBackgroundColor = Opt.AgeTikSymbFillClr
      Else
        .MarkerBackgroundColorIndex = xlNone
      End If
      .MarkerStyle = Opt.AgeTikSymbol
      .MarkerSize = Opt.AgeTikSymbSize
    End With
  End If
  If uEvoCurvLabelAge > 0 Then
    If i = 1 Then
      MaxDecPts = 0
      For j = 1 To Ncurves
        MaxDecPts = Max(MaxDecPts, NumDeci(Ugamma0(j)))
      Next j
    End If
    T = uEvoCurvLabelAge * Thou: yy = GammaU(T, Ugamma0(i))
    If yy > MinY And yy < MaxY Then ' Don't label unless within plotbox
      Th230_U238ar T, yy, xx
      If xx > MinX And xx < MaxX Then
        UconcSlope T, Ugamma0(i), Slp
        s$ = IIf(MaxDecPts > 0, "0." & String(MaxDecPts, "0"), "0")
        s$ = Format(Ugamma0(i), s$)
        StatBar "Labeling curve " & s$
        sxx = xx: syy = yy
        RotLabel s$, "Times New Roman", 9, True, sxx, syy, Slp, (Slp > 0 And Ugamma0(i) > 1)
        xx = sxx: yy = syy
      End If
    End If
  End If
NextCurve:
Next i
If uPlotIsochrons Then
  StatBar "adding isochrons"
  Sheets(PlotName$).Select
  For i = 1 To Cv.Nisocs ' Plot/label the isochrons
    IsoChrt.SeriesCollection.Add UisoRange(i), xlColumns, False, True, False
    Set SerC = IsoChrt.SeriesCollection
    With Last(SerC)
      If .MarkerStyle <> xlNone Then .MarkerStyle = xlNone
      With .Border
        If .Weight <> xlHairline Then .Weight = xlHairline
        .Color = RGB(128, 128, 128)
        If ColorPlot Then
          If .LineStyle <> xlContinuous Then .LineStyle = xlContinuous
        Else
          .LineStyle = xlDot
        End If
      End With
    End With
    If uLabelTiks Then ' Label the isochron ages
      xx = UisoRange(i).Cells(2, 1): yy = UisoRange(i).Cells(2, 2)
      Slp = (yy - UisoRange(i).Cells(1, 2)) / (xx - UisoRange(i).Cells(1, 1))
      Inr = yy - Slp * xx
      If xx > MinX And yy > MinY Then
        If xx >= MaxX Then ' UisochPos = +1 or -1
          xx = MaxX + UisochPos * 0.04 * Xspred: yy = Slp * xx + Inr
          If Abs(yy - LastYy) / Yspred < 0.06 Then xx = -1
        Else
          yy = MaxY + UisochPos * 0.05 * Yspred: xx = (yy - Inr) / Slp
          If Abs(xx - LastXx) / Xspred < (0.01 * (5 + UisochPos)) Then xx = -1
        End If
        ThUage xx, yy, T ' Get isochron age from actual plotted line-start
        T = Drnd(T / Thou, 5) ' Convert to ka
        If T > 0 And xx > 0 Then
          s$ = tSt(T)
          StatBar "Labeling isochrons " & s$
          sxx = xx: syy = yy
          RotLabel s$, "Arial", 8.5, False, (xx), (yy), Slp
          xx = sxx: yy = syy
          LastXx = xx: LastYy = yy
        End If
      End If
    End If
  Next i
End If
End Sub

Private Sub Ecalc(ByVal Age#, ByVal Which%)
' Calculate exponential values after err checking.  If Which=1, do for
'   all; if 2, for Th230 & diff only; if 3 for Th230 & U234 only.
' Return values of -1 if out of range.
Dim Test230#, Test234#, TestDiff#
GetConsts
Test230 = -Lambda230 * Age
If Which <> 2 Then Test234 = Lambda234 * Age
If Which <> 3 Then TestDiff = LambdaDiff * Age
Exp234 = -1: Exp230 = -1: ExpDiff = -1
If Abs(Test230) < MAXEXP Then
  Exp230 = Exp(Test230)
End If
If Which <> 2 And Abs(Test234) < MAXEXP Then
  Exp234 = Exp(Test234)
End If
If Which <> 3 And Abs(TestDiff) < MAXEXP Then
  ExpDiff = Exp(-TestDiff)
End If
End Sub

Sub Initial234(ByVal Age#, ByVal AgeErr#, ByVal ThRat#, _
  ByVal Gamma#, ByVal GammaErr#, ByVal ThRatErr#, ByVal RhoThGamma#, _
  ByRef Gamma0#, ByRef Gamma0err#, ByRef RhoAgeGamma0#, _
  Optional AtomRat = False, Optional NLE = True)
' Given an observed 234/238 activity ratio (Gamma), an age calc from
'  observed 230/238 ratio, & the abs errs & error-corrs in
'  these values, calc the initial 234/238 ratio (Gamma0) & err.
Dim l8#, l4#, L0#, L4err#, L#, Ld#
Dim d#, DP#, G#, Gerr#, Cov#, Gp#
Dim Tterm#, Gterm#, GtermI#, TtermI#, CovTG#
Dim Var#, WLE As Boolean, L4termI, CovAgeGamma0#, CovTL4#, L0err#
Dim ThU#, Gamm#, ThUerr#, GammErr#
ViM AtomRat, False
ViM NLE, True
WLE = Not NLE
RhoAgeGamma0 = 0: Ierror = False
Ecalc Age, 1
If Exp230 = -1 Or Exp234 = -1 Or ExpDiff = -1 Then
  Gamma0 = 0: Gamma0err = 0
  Ierror = True: Exit Sub
End If
L0 = Lambda230:       l4 = Lambda234:       l8 = Lambda238 / Million
L0err = Lambda230err: L4err = Lambda234err: Ld = LambdaDiff
L = Exp230:           d = ExpDiff:          DP = 1 - d
G = Gamma:            Gerr = GammaErr:      ThU = ThRat
ThUerr = ThRatErr
If AtomRat Then
  ThU = ThU * l8 / L0: G = G * l8 / l4
  ThUerr = ThRatErr / ThRat * ThU:  Gerr = GammaErr / Gamma * G
End If
Cov = RhoThGamma * ThUerr * Gerr
If AtomRat Then ' Calculated from atomic ratios x decay-const ratios
  Gp = G * l4 / l8 - 1
  Gterm = l4 * DP / Ld
  GtermI = l4 / l8 * Exp234
  Tterm = l8 * L + (l4 * G - l8) * d
  TtermI = l4 * Gp * Exp234
  CovTG = (Cov - Gterm * Gerr * Gerr) / Tterm
  If WLE Then L4termI = (G / l8 + Age * Gp) * Exp234
Else            ' Measured alpha or secular-equilibrium-normalized
  Gp = G - 1
  Gterm = L0 / Ld * DP
  GtermI = Exp234
  Tterm = L0 * (L + Gp * d)
  TtermI = l4 * Gp * Exp234
  CovTG = (Cov - Gterm * Gerr * Gerr) / Tterm
  If WLE Then L4termI = Age * Gp * Exp234
End If
Gamma0 = 1 + Gp * Exp234
Var = SQ(TtermI * AgeErr) + SQ(GtermI * Gerr)
Var = Var + 2 * TtermI * GtermI * CovTG
If Var < 0 Then Ierror = True: Gamma0err = 0: Exit Sub
CovAgeGamma0 = TtermI * AgeErr * AgeErr + GtermI * CovTG
If WLE Then
  CovTL4 = -L4term / Tterm * L4err * L4err
  Var = Var + SQ(L4termI * L4err) + 2 * TtermI * L4termI * CovTL4
  CovAgeGamma0 = CovAgeGamma0 + L4termI * CovTL4
End If
Gamma0err = Sqr(Var)
If Gamma0err > 0 Then RhoAgeGamma0 = CovAgeGamma0 / AgeErr / Gamma0err
End Sub

Sub InitialCorr(Detr#(), Sample#(), CarbThU#, ThUerr#, _
   CarbGamma#, GammaErr#, RhoCarbGammaThU#, Failed As Boolean)
Attribute InitialCorr.VB_ProcData.VB_Invoke_Func = " \n14"
' Correct for initial (detrital) Th230,U234,U238 using Th232 as the index.
' Algorithm looks at the data on 232/238-230/238 & 230/238-234/238 iso-
'  chron plots, & solves for the intercepts, intercept-errs, & correls.
' Detr() contains the ratios & errors for the detrital component, Sample() the ratios
'  & errors for the measured sample, in the order 232/238, err, 230/238, err,
'  234/238, err, rho(232,238,230/238), rho(232/238,234/238), rho(230/238,234/238).
' x=232/238, y=234/238, z=230/238
' subscript 1 refers to sample, subscript 2 refers to detritus
Dim i%, j%, Fpar%, Typ$
Dim x1#, x2#, y1#, y2#, z1#, z2#
Dim x1Err#, x2Err#, y1Err#, y2Err#, z1Err#, z2Err#
Dim Term0#, term1#, term2#, term3#
Dim s1#, i1#, s2#, i2#
Dim r1#, r2#, r1s#, r2s#
Dim COVx1y1#, COVx2y2#, COVx1z1#, COVx2z2#
Dim COVy1z1#, COVy1z2#, COVi1i2#, COVy2z2#
Dim VARx1#, VARx2#, VARy1#, VARy2#
Dim VARz1#, VARz2#, VARi1#, VARi2#
Dim DeltaX#, XvarTerm#, Rhox1y1#, Rhox2y2#
Dim Rhox1z1#, Rhox2z2#, Rhoy2z2#, Rhoy1z1#
x1 = Detr(1):        z1 = Detr(3):        y1 = Detr(5)
x1Err = Detr(2):     z1Err = Detr(4):     y1Err = Detr(6)
Rhox1z1 = Detr(7):   Rhox1y1 = Detr(8):   Rhoy1z1 = Detr(9)
x2 = Sample(1):      z2 = Sample(3):      y2 = Sample(5)
x2Err = Sample(2):   z2Err = Sample(4):   y2Err = Sample(6)
Rhox2z2 = Sample(7): Rhox2y2 = Sample(8): Rhoy2z2 = Sample(9)
Failed = True
DeltaX = x2 - x1
If Abs(DeltaX) < 0.000001 Then Exit Sub
s1 = (y2 - y1) / DeltaX:   s2 = (z2 - z1) / DeltaX ' Slopes
i1 = y1 - s1 * x1:         i2 = z1 - s2 * x1       ' Intercepts
VARx1 = x1Err * x1Err: VARx2 = x2Err * x2Err '/
VARy1 = y1Err * y1Err: VARy2 = y2Err * y2Err '| Data-point variances
VARz1 = z1Err * z1Err: VARz2 = z2Err * z2Err '\
COVx1y1 = Rhox1y1 * x1Err * y1Err     '/
COVx2y2 = Rhox2y2 * x2Err * y2Err     '|
COVx1z1 = Rhox1z1 * x1Err * z1Err     '| Data-point covariances
COVx2z2 = Rhox2z2 * x2Err * z2Err     '|
COVy1z1 = Rhoy1z1 * y1Err * z1Err     '|
COVy2z2 = Rhoy2z2 * y2Err * z2Err     '\
r1 = x1 / DeltaX: r1s = r1 * r1      '/
r2 = x2 / DeltaX: r2s = r2 * r2      '| useful terms
XvarTerm = r2s * VARx1 + r1s * VARx2 '\
term1 = s1 * s1 * XvarTerm + r2s * VARy1 + r1s * VARy2
term2 = -2 * s1 * (r2s * COVx1y1 + r1s * COVx2y2)
VARi1 = term1 + term2
term1 = s2 * s2 * XvarTerm + r2s * VARz1 + r1s * VARz2
term2 = -2 * s2 * (r2s * COVx1z1 + r1s * COVx2z2)
VARi2 = term1 + term2
If VARi1 < 0 Or VARi2 < 0 Then
  MsgBox "Impossible combination of errors & error-correlations", vbExclamation, Iso
  Exit Sub
End If
Term0 = Sqr(VARi1 * VARi2): term1 = s1 * s2 * XvarTerm
term2 = r2s * (COVy1z1 - s2 * COVx1y1 - s1 * COVx1z1)
term3 = r1s * (COVy2z2 - s2 * COVx2y2 - s1 * COVx2z2)
COVi1i2 = term1 + term2 + term3
CarbGamma = i1:         CarbThU = i2
If CarbGamma = 0 Then CarbGamma = 0.000000001
If CarbThU = 0 Then CarbThU = 0.000000001
GammaErr = Sqr(VARi1):  ThUerr = Sqr(VARi2)
RhoCarbGammaThU = COVi1i2 / Sqr(VARi1 * VARi2)
If Abs(RhoCarbGammaThU) > 1 Then
  MsgBox "Roundoff error in error-correlation solution", , Iso
  Exit Sub
End If
Failed = False
End Sub

Private Sub SimpleLinear(X#(), y#(), z#())
' Simple linear regression, Y on X, & also Z on X.
With App
  IntSl(1) = .Intercept(y, X): IntSl(2) = .Slope(y, X)
  IntSl(3) = .Intercept(z, X): IntSl(4) = .Slope(z, X)
End With
End Sub

Private Function Th230AgeErr(ByVal ThRatErr#, ByVal GammaErr#, ByVal Age#, _
  ByVal ThRat#, ByVal Gamma#, ByVal Rho#, _
  Optional AtomRat = False, Optional NLE = True)
' Returns error in Th230/U age, propagated from the abs err in the
'  230/238 or 230/234 activity ratio (ThRatErr), the 234/238 activity ratio,
'  & the 230/234-234/238 or 230/238-234/238 err-correl (Rho).
' Gamma is the measured U234/U238 activity ratio.
Dim Cov#, GG#, k1#, k2#, d#, Var#, Numer#
Dim L0#, l4#, l8#, Ld#, L#, DP#
Dim tmp1#, tmp2#, Tterm#, Gterm#, G#, Gp#
Dim Gerr#, WLE As Boolean, L0err#, L4err#, ThUerr#, ThU#
ViM AtomRat, False
ViM NLE, True
Ecalc Age, 2
L0 = Lambda230:       l4 = Lambda234:       l8 = Lambda238 / Million
L0err = Lambda230err: L4err = Lambda234err: Ld = LambdaDiff
L = Exp230:           d = ExpDiff:          DP = 1 - d
Ierror = False
WLE = Not NLE
If Exp230 <> -1 And ExpDiff <> -1 Then
  If AtomRat Then
    ThU = ThRat * l8 / L0:           G = Gamma * l8 / l4
    ThUerr = ThRatErr / ThRat * ThU: Gerr = GammaErr / Gamma * G
    Cov = Rho * ThUerr * Gerr
    Gp = l4 * G - l8
    Tterm = l8 * L + Gp * d:        Gterm = l4 * DP / Ld
    If WLE Then
      tmp1 = l8 / L0 * ((Age + 1 / L0) * L - 1 / L0)
      tmp2 = Gp / Ld * ((Age + 1 / Ld) * d - 1 / Ld)
      L0term = tmp1 + tmp2
      L4term = (Gp * (DP / Ld - Age * d) + G * DP) / Ld
    End If
  Else
    ThU = ThRat:          G = Gamma
    ThUerr = ThRatErr:    Gerr = GammaErr
    Cov = Rho * ThUerr * Gerr
    Gp = G - 1
    Gterm = L0 / Ld * DP: Tterm = L0 * (L + Gp * d)
    If WLE Then
      L0term = Age * L - Gp / Ld * (DP * l4 / Ld - L0 * Age * d)
      L4term = Gp * L0 / Ld * (DP / Ld - Age * d)
    End If
  End If
  Numer = ThUerr * ThUerr + SQ(Gterm * Gerr) - 2 * Gterm * Cov
  If WLE Then
    Numer = Numer + SQ(L0term * Lambda230err) + SQ(L4term * Lambda234err)
  End If
  Th230AgeErr = Sqr(Numer / SQ(Tterm))
Else
  Th230AgeErr = 0: Ierror = True
End If
End Function

Private Function Th230U238ar_X(ByVal AgeKyr#, ByVal U234238ar#)  ' Re-enable if seems useful at some point
' Input is Age in kiloyears, present-day 234/238 activity ratio (=Gamma).
' Returns activity ratio of Th230/U238 for that age.
Dim Ar#, Gamma0#
Th230_U238ar AgeKyr * Thou, U234238ar, Ar
Th230U238ar_X = Ar
End Function

Sub Th230_U238ar(ByVal AgeYr#, ByVal GammaT#, ByRef ThUar#)
Attribute Th230_U238ar.VB_ProcData.VB_Invoke_Func = " \n14"
' Input is Age in kyr, initial 234/238 activity ratio (=Gamma).
' Returns activity ratio of Th230/U238 for that age (=ThUar).
Ierror = False
Ecalc AgeYr, 2

If Exp230 <> -1 And ExpDiff <> -1 Then
  ThUar = 1 - Exp230 + LambdaK * (GammaT - 1) * (1 - ExpDiff)
Else
  Ierror = True: ThUar = 0
End If
End Sub

Private Function Th230238Deriv(ByVal Age#, ByVal Gamma#)
' Input is Age in years & 234/238 activity ratio (=Gamma).
' Returns first derivative of the Th230/U238 activity-ratio function at that age.
Ierror = False
Ecalc Age, 2
If Exp230 <> -1 And ExpDiff <> -1 Then
  Th230238Deriv = Lambda230 * (Exp230 + (Gamma - 1) * ExpDiff)
Else
  Th230238Deriv = 0:   Ierror = True
End If
End Function

Sub UconcSlope(ByVal T#, ByVal Gamma0#, UcurvSlope#)
Attribute UconcSlope.VB_ProcData.VB_Invoke_Func = " \n14"
' Return slope of the U-series evolution curve at t (in years) for Gamma0
Dim dY#, dx#, Gamma#, x1#, x2#
Dim y1#, y2#, t1#, t2#
'dY = -Lambda234 * (Gamma0 - 1) * Exp(-Lambda234 * t)
'Gamma = 1 + (Gamma0 - 1) * Exp(-Lambda234 * t)
'dX = Th230238Deriv(t, Gamma)
'If dX = 0 Then dX = 1
t1 = T - 100: t2 = T + 100
y1 = GammaU(t1, Gamma0): y2 = GammaU(t2, Gamma0)
x1 = Th230U238ar(t1 / Thou, Gamma0): x2 = Th230U238ar(t2 / Thou, Gamma0)
dY = y2 - y1: dx = x2 - x1
UcurvSlope = dY / dx
End Sub

Sub ThUage(ByVal ThU#, ByVal Gamma#, ThU_age#)
Attribute ThUage.VB_ProcData.VB_Invoke_Func = " \n14"
' From observed Th230/U238 & U234/U238 activity ratios (ThU & Gamma),
'   calculate age.
' This algorithm will not give sol'ns for ages old enough to be in the
'  "ratio-reversal" part of the Th230/U238 vs Age curve (if the initial
'  234/238 is >1, at some relatively old age-interval, the Th230/U238 activity
'  ratio will be >1, then decline back to 1 as the age increases).
Dim Deriv#, Converged%, Iter%, Ratio#, Delta#, test#
Dim Age#, i%, MaxIter%, TooOld%, TimeIn#, MaxTime#, Gamma0#
Age = 5000: Iter = 1: MaxIter = Hun: TimeIn = Timer(): MaxTime = 12

Do
  If Iter Mod 40 = 0 Then TooLongCheck TimeIn, MaxTime
  Th230_U238ar Age, Gamma, Ratio
If Ierror Then Exit Do
  Deriv = Th230238Deriv(Age, Gamma)
If Ierror Or Deriv = 0 Then Exit Do
  Delta = (Ratio - ThU) / Deriv
  Age = Age - Delta
  test = Abs(Delta / Age)
  Converged = (test < 0.0001)
  Iter = 1 + Iter
Loop Until Converged Or Iter > MaxIter
TooOld = (Abs(Age / (Log_2 / Lambda230#)) > 15)
If Ierror Or Iter > MaxIter Or TooOld Then Age = 0
ThU_age = Age
End Sub

Sub U234_Age(ByVal Gamma#, ByVal InitialGamma#, ByVal GammaErr#, _
  Age#, UpperAgeErr#, LowerAgeErr#)
' Given measured U234/U238 activity ratio (=Gamma), calculates age @ absolute
'  error, assuming an initial 234/238 activity ratio of Seawater234.  Assumes
'  Assumes no uncertainty in the U234 decay const. or in Seawater234.
Dim temp#, UpperAgeLimit#, LowerAgeLimit#
Ierror = False
Age = 0: UpperAgeErr = 0: LowerAgeErr = 0
temp = (Gamma - 1) / (InitialGamma - 1)
If temp > MAXLOG Or temp < MINLOG Then Ierror = True: Exit Sub
Age = -Log(temp) / Lambda234
If GammaErr <> 0 And Age <> 0 Then
  U234_Age Gamma - GammaErr, InitialGamma, 0, UpperAgeLimit, 0, 0
  U234_Age Gamma + GammaErr, InitialGamma, 0, LowerAgeLimit, 0, 0
  UpperAgeErr = IIf(UpperAgeLimit, UpperAgeLimit - Age, 0)
  LowerAgeErr = IIf(LowerAgeLimit, Age - LowerAgeLimit, 0)
End If
' AgeErr = GammaErr / (Lambda234 * (Gamma - 1))  ' Symmetric expresion
End Sub

Private Function aSums(Coef#(), X#(), y#(), z#(), _
  Omega#(), ByVal N&, Failed As Boolean)
' Return sums of squares of weighted residuals for XYX regression (either un-
'  constrained or concordia-constrained).
' Coef() contains the XY-XZ slope,inter,slope,inter; Omega() the inverted
'  variance-covariance matrix for each point.
Dim i&, SlopeXY#, SlopeXZ#, InterXY#, InterXZ#
Dim Ry#, Rz#, By#, Bz#, T#
Dim term1#, term2#, term3#, Term4#
Dim Alpha#, Beta#, Gamma#, WtdResidSq#, SumWtdResidSq#
InterXY = Coef(1)
If ConcConstr Then ' Coef(1)=(207/206)common  Coef(2)=(204/206)common  Coef(3)=t
  InterXZ = Coef(2)
  T = Coef(3)
  SlopeXY = ConcX(T, False, True) / Uratio - InterXY * ConcY(T, False, True)
  SlopeXZ = -InterXZ * ConcY(T, False, True)
Else
  SlopeXY = Coef(2): SlopeXZ = Coef(4)
  InterXZ = Coef(3)
End If
For i = 1 To N
  Ry = y(i) - InterXY - SlopeXY * X(i)
  Rz = z(i) - InterXZ - SlopeXZ * X(i)
  term1 = Omega(i, 1, 1) + Omega(i, 2, 2) * SlopeXY * SlopeXY
  term2 = 2 * Omega(i, 1, 2) * SlopeXY
  term3 = 2 * Omega(i, 1, 3) * SlopeXZ + Omega(i, 3, 3) * SlopeXZ * SlopeXZ
  Term4 = 2 * Omega(i, 2, 3) * SlopeXY * SlopeXZ
  Alpha = term1 + term2 + term3 + Term4
  term1 = Omega(i, 1, 2) + SlopeXY * Omega(i, 2, 2) + SlopeXZ * Omega(i, 2, 3)
  term2 = Omega(i, 1, 3) + SlopeXY * Omega(i, 2, 3) + SlopeXZ * Omega(i, 3, 3)
  Beta = Ry * term1 + Rz * term2
  term1 = Omega(i, 2, 2) * Ry * Ry + Omega(i, 3, 3) * Rz * Rz
  term2 = 2 * Omega(i, 2, 3) * Ry * Rz
  Gamma = term1 + term2
  WtdResidSq = Gamma - Beta * Beta / Alpha
  If WtdResidSq < 0 Then aSums = -1: Failed = True: Exit Function
  SumWtdResidSq = SumWtdResidSq + WtdResidSq
Next i
aSums = SumWtdResidSq
End Function

Private Sub Amoeba(P#(), y#(), aX#(), aY#(), _
  aZ#(), ByVal Ndim&, Omega#(), ByVal N&, _
  ByVal ErrorMag#, ByVal Ftol#, ByVal IterMax%, Failed As Boolean)
' Simplex multidimensional Fn-min finder, to be used to find the
'   coefs for a Fn that will yield the least sums-of-squares.
' Must be used with the user-defined Fn aSums, which will return the
'   sums-of-squares for any given set of Fn-coefs & X-Y-Z data.
' Adapted from Press & others, Numerical Recipes, 1986, Cambridge Univ. Press, 292-293
' Ndim is the number of coefs (the Coef vector) in the function to be
'   minimized, aX(),aY(),aZ() the X-Y-Z values for the Fn, N the #
'   of X-Y-Z pts, Ftol is the fractional convergence tolerance to be
'   achieved in the Fn (should not be smaller than 1E-15 or larger
'   than, say, 1E-6), IterMax the maximum # of iters permitted.
' P() must contain a starting simplex for the Fn when passed to this
'   sub - that is, P() must contain Ndim+1 rows whose columns are guessed-at
'   values of Coef().  Y() must contain the Fn-responses for these
'   Ndim+1 guesses of Coef().
Dim i&, j&, Mpts&, Disp%
Dim Iter%, Ihi%, Ilo%, inHi%
Dim MSWD#, Mfact#, TimeIn#, SB$, Iter1%
Dim Ypr#, Yprr#, Rtol#, Rtol1#, Rtol2#
Dim pr#(), Prr#(), Pbar#()
ReDim pr(Ndim), Prr(Ndim), Pbar(Ndim)
Const Tiny = 0.000000000001, Alpha = 1#, Beta = 0.5, Gamma = 2#, MaxTime = 12
SB$ = "Simplex iteration"
Mfact = ErrorMag * ErrorMag
Failed = False
TimeIn = Timer()
Mpts = Ndim + 1: Iter = 0: Ilo = 1
Do
  Iter1 = 1 + Iter
  If Iter1 Mod 10 = 0 Then StatBar SB$ & Str(Iter1)
  Ilo = 1
  If y(1) > y(2) Then
    Ihi = 1: inHi = 2
  Else
    Ihi = 2: inHi = 1
  End If
  For i = 1 To Mpts
    If y(i) < y(Ilo) Then Ilo = i
    If y(i) > y(Ihi) Then
      inHi = Ihi: Ihi = i
    ElseIf y(i) > y(inHi) Then
      If i <> Ihi Then inHi = i
    End If
  Next i
  'StatBar str(Min(Y(Ihi), Y(Ilo)) / (2 * N - 4 - ConcConstr))
   Rtol = 2 * Abs(y(Ihi) - y(Ilo)) / (Abs(y(Ihi)) + Abs(y(Ilo)))
  ' Converged?
If (Rtol < Ftol Or (Rtol = Rtol1 And Rtol1 = Rtol2)) Then Exit Do
If Ypr <> 0 And Yprr <> 0 And (Ypr < Tiny Or Yprr < Tiny) Then Exit Do
  ' Keep track of last 2 Rtol values - if absolutely no change,  conclude
  '   that the algorithm is stalled & call it converged.
  Rtol1 = Rtol: Rtol2 = Rtol1
  If Iter = IterMax Then Failed = -1: Exit Do
  Iter = 1 + Iter
  If Iter Mod 30 = 0 Then TooLongCheck TimeIn, MaxTime
  For j = 1 To Ndim: Pbar(j) = 0#: Next j
  For i = 1 To Mpts
    If i <> Ihi Then
      For j = 1 To Ndim
        Pbar(j) = Pbar(j) + P(i, j)
      Next j
    End If
  Next i
  For j = 1 To Ndim
    Pbar(j) = Pbar(j) / Ndim
    pr(j) = (1# + Alpha) * Pbar(j) - Alpha * P(Ihi, j)
  Next j
  Ypr = aSums(pr(), aX(), aY(), aZ(), Omega(), N, Failed)
  If Failed Then Exit Do
  If Ypr <= y(Ilo) Then
    For j = 1 To Ndim
      Prr(j) = Gamma * pr(j) + (1# - Gamma) * Pbar(j)
    Next j
    Yprr = aSums(Prr(), aX(), aY(), aZ(), Omega(), N, Failed)
    If Failed Then Exit Do
    If Yprr < y(Ilo) Then
      For j = 1 To Ndim: P(Ihi, j) = Prr(j): Next j
      y(Ihi) = Yprr
    Else
      For j = 1 To Ndim: P(Ihi, j) = pr(j): Next j
      y(Ihi) = Ypr
    End If
  ElseIf Ypr >= y(inHi) Then
    If Ypr < y(Ihi) Then
      For j = 1 To Ndim: P(Ihi, j) = pr(j): Next j
      y(Ihi) = Ypr
    End If
    For j = 1 To Ndim
      Prr(j) = Beta * P(Ihi, j) + (1# - Beta) * Pbar(j)
    Next j
    Yprr = aSums(Prr(), aX(), aY(), aZ(), Omega(), N, Failed)
    If Failed Then Exit Do
    If Yprr < y(Ihi) Then
      For j = 1 To Ndim: P(Ihi, j) = Prr(j): Next j
      y(Ihi) = Yprr
    Else
      For i = 1 To Mpts
        If i <> Ilo Then
          For j = 1 To Ndim
            pr(j) = 0.5 * (P(i, j) + P(Ilo, j))
            P(i, j) = pr(j)
          Next j
          y(i) = aSums(pr(), aX(), aY(), aZ(), Omega(), N, Failed)
          If Failed Then Exit Do
        End If
      Next i
    End If
  Else
    For j = 1 To Ndim: P(Ihi, j) = pr(j): Next j
    y(Ihi) = Ypr
  End If
Loop
StatBar
End Sub

Sub Th230age_Gamma0(Results#(), ByVal Th230U238#, ByVal Th230Err#, _
  ByVal U234U238#, ByVal U234err#, Optional ByVal RhoThU# = 0, _
  Optional PercentErrs = False, Optional SigmaLevel = 1, Optional WithLambdaErrs = False, _
  Optional AtomRat = False)
' Returns Results() as 230Th/U age (ka), err, Gamma0, err, Rho T-Gamma0
' Sigma level must be 1 or 2; error output is 1-sigma
' Default is absolute, **1-sigma** input errors, no lambda errors, activity ratios, zero err-correl.
Dim NLE As Boolean, ThErr#, Uerr#
ViM RhoThU, 0
ViM PercentErrs, False
ViM SigmaLevel, 1
ViM WithLambdaErrs, False
ViM AtomRat, False
NLE = Not WithLambdaErrs
SigLev = MinMax(1, 2, SigmaLevel)
If PercentErrs Then
  ThErr = Th230Err / Hun * Th230U238
  Uerr = U234err / Hun * U234U238
Else
  ThErr = Th230Err: Uerr = U234err
End If
ThErr = ThErr / SigLev: Uerr = Uerr / SigLev
SigLev = 1   ' Input errors are now changed to 1-sigma, absolute
GetConsts
ThUage_Gamma0 Th230U238, ThErr, U234U238, Uerr, RhoThU, False, NLE, (AtomRat), Results()
' Output errs from above sub are 2-sigma/95%conf!
If NIM(SigmaLevel) Then
  If SigmaLevel = 1 Then
    Results(2) = Results(2) / 2
    Results(4) = Results(4) / 2
  ElseIf SigmaLevel <> 2 Then
    Results(2) = 0: Results(4) = 0: Results(5) = 0
  End If
End If
End Sub

Sub ThUage_Gamma0(ByVal Th230U238#, ByVal Th230Err#, ByVal U234U238#, _
  ByVal U234err#, ByVal RhoThU#, ByVal MonteCarlo As Boolean, _
  ByVal NLE As Boolean, ByVal AtomRat As Boolean, Res#())
' Analytical: Returns 230Th/U age (ka), err, Gamma0, err, Rho T-Gamma0
' Monte Carlo: Returns 230Th/U age (ka), +err, -err, Gamma0, +err, -err, Rho T-Gamma0
' MonteCarlo errors are 95%-conf; input errors are 1-sigma absolute.
' Returns errors as 95% conf (MonteCarlo) or 2-sigma.  INPUT ERRORS MUST BE 1-SIGMA
Dim ThErr#, Uerr#, AgeErrYr#, Bad As Boolean
Dim MeanAgeYr#, ModeAgeYr#, MedAgeYr# ', MeanAgeYr#
Dim i%, E95L#, E95U#, E95Lg#, E95Ug#
Dim RhoAgeG0#, Finite As Boolean, Gamma0#

ThUage Th230U238, U234U238, MeanAgeYr
Res(1) = MeanAgeYr / Thou
On Error Resume Next
Finite = IsOn(DlgSht("ThUage").CheckBoxes("Finite"))
On Error GoTo 0

If MeanAgeYr = 0 And Not MonteCarlo Then

  For i = 1 To 5 - 2 * MonteCarlo
    Res(i) = 0
  Next i

  Exit Sub
End If

If MonteCarlo Then
  ' 11/06/18 -- added median & mode ages
  MonteCarloThUerrs MeanAgeYr, MedAgeYr, ModeAgeYr, Gamma0, _
    (Th230U238), (Th230Err), (U234U238), (U234err), (RhoThU), _
    E95L, E95U, E95Lg, E95Ug, RhoAgeG0, Bad, NLE, AtomRat, Finite

  If Bad Then

    For i = 1 To 7
      Res(i) = 0
    Next i

    Exit Sub
  ElseIf Finite Then
    Res(1) = MeanAgeYr / Thou
  End If

  Res(2) = 0: Res(3) = 0: Res(5) = 0: Res(6) = 0

  If MeanAgeYr <> 0 Then

    If Finite Then
      Res(4) = Gamma0
    Else
      Res(4) = InitU234U238(Res(1), U234U238)
    End If

    Res(7) = RhoAgeG0
  End If

  If E95U <> 0 And MeanAgeYr <> 0 Then
    Res(2) = (E95U - MeanAgeYr) / Thou
    Res(3) = (E95L - MeanAgeYr) / Thou
    Res(5) = E95Ug - Res(4)
    Res(6) = E95Lg - Res(4)
    Res(8) = MedAgeYr / Thou     ' 11/06/18 -- added
    Res(9) = ModeAgeYr / Thou   '    "
  ElseIf E95L <> 0 Then
    Res(3) = E95L / Thou
    Res(5) = E95Ug
    If Res(5) = 0 Then Res(5) = E95Lg
  End If

Else
  ' Find age error in years 1-sigma/68%-conf.
  AgeErrYr = Th230AgeErr(Th230Err, U234err, MeanAgeYr, Th230U238, U234U238, RhoThU, AtomRat, NLE)
  Res(2) = AgeErrYr / Thou * 2  ' Convert to 2sigma/95%conf. kyr
  ' Find 1-sigma/68%-conf. Gamma0 error
  Initial234 MeanAgeYr, AgeErrYr, Th230U238, U234U238, U234err, Th230Err, _
      RhoThU, Res(3), Res(4), Res(5), AtomRat, NLE
  Res(4) = Res(4) * 2   ' Convert to 2sigma/95%conf. kyr
End If

End Sub

Private Sub Calc3Dline(DP() As DataPoints, X#(), y#(), z#(), _
  ByVal N&, Resid#(), ByVal Toler#, Failed As Boolean, _
  MSWD#, VarCov#())
' Do the numeric scut-work to find app best-fit XYZ line using either the algorithm
'  in Ludwig & Titterington, 1994, or in Ludwig, 1998.
Dim ErrMag  As Boolean, CheckMLE As Boolean
Dim i&, j&, k&, Solid%, Ndim&
Dim sigmaX#(), SigmaY#(), SigmaZ#()
Dim RhoXY#(), RhoXZ#(), RhoYZ#(), MeanFractErr#
Dim Dum, Xt2#, SumFractErr#, Xbar#, Ybar#, Zbar#
Dim a1#, a2#, b1#, b2#, tmp#
Dim OmegaTerm As Variant, Omega#(), vc#(3, 3)
Dim T#, f As Variant, G As Variant, Finv As Variant
Dim Xt#(), Scoef#(), Simplex#(), sY#()
Dim Comm46#, Comm76#, Sxy#, Sxz#, Pred86#
Dim PredT#, wt#, SumWt#, Tsum#, DeltaInv As Variant
Dim bbT As Variant, tmp1 As Variant, tmp2 As Variant, M As Variant, Bb As Variant
Dim Delta() As Variant
Const Nines = 0.999999
Ndim = 4 + ConcConstr
ReDim f(Ndim, Ndim), Finv(Ndim, Ndim), Xt(N), G(Ndim, Ndim), VarCov(Ndim, Ndim)
ReDim Simplex(1 + Ndim, Ndim), sY(1 + Ndim), Scoef(Ndim)
ReDim sigmaX(N), SigmaY(N), SigmaZ(N), RhoXY(N), RhoXZ(N), RhoYZ(N), Omega(N, 3, 3)
Failed = False
For i = 1 To N
  With DP(i)
    If .X = 0 Or .y = 0 Or .z = 0 Then
      Failed = True
    ElseIf .Xerr = 0 Or .Yerr = 0 Or .Zerr = 0 Then
      Failed = True
    ElseIf Abs(.RhoXY) > 1 Or Abs(.RhoXZ) > 1 Or Abs(.RhoYZ) > 1 Then
      Failed = True
    End If
    If Failed Then Exit Sub
      X(i) = .X: sigmaX(i) = .Xerr
      RhoYZ(i) = .RhoYZ
      y(i) = .y: SigmaY(i) = .Yerr
      z(i) = .z: SigmaZ(i) = .Zerr
      RhoXY(i) = .RhoXY:  RhoXZ(i) = .RhoXZ
      ' Error correlations of exactly 1 can result in divisions-by-zero
      If Abs(RhoXY(i)) = 1 Then RhoXY(i) = Sgn(RhoXY(i)) * Nines
      If Abs(RhoXZ(i)) = 1 Then RhoXZ(i) = Sgn(RhoXZ(i)) * Nines
      If Abs(RhoYZ(i)) = 1 Then RhoYZ(i) = Sgn(RhoYZ(i)) * Nines
    End With
Next i
CheckMLE = False ' Check the MLE boundary conditions?
If Not ConcConstr Then
  ErrMag = True '  Rescale errors so their mean (fractional) errs are ~100%?
  ' Transform X,Y,Z to have mean of 0 to avoid subsequent ill-conditioned matrices
  Xbar = iAverage(X): Ybar = iAverage(y): Zbar = iAverage(z)
  For i = 1 To N
    If X(i) = Xbar Or y(i) = Ybar Or z(i) = Zbar Then
      GoSub NoErMag
      Exit For
    Else
      X(i) = X(i) - Xbar: y(i) = y(i) - Ybar: z(i) = z(i) - Zbar
    End If
  Next i
  If ErrMag Then ' Expand errs uniformly to avoid later roundoff problems
    For i = 1 To N
      SumFractErr = SumFractErr + Abs(sigmaX(i) / X(i)) _
        + Abs(SigmaY(i) / y(i)) + Abs(SigmaZ(i) / z(i))
    Next i
    MeanFractErr = SumFractErr / N / 3
    For i = 1 To N
      sigmaX(i) = sigmaX(i) / MeanFractErr
      SigmaY(i) = SigmaY(i) / MeanFractErr
      SigmaZ(i) = SigmaZ(i) / MeanFractErr
    Next i
  End If
Else
  GoSub NoErMag
End If
For i = 1 To N ' Create the variance/covariance matrices
  vc(1, 1) = SQ(sigmaX(i))
  vc(2, 2) = SQ(SigmaY(i))
  vc(3, 3) = SQ(SigmaZ(i))
  vc(1, 2) = sigmaX(i) * SigmaY(i) * RhoXY(i)
  vc(2, 1) = vc(1, 2)
  vc(1, 3) = sigmaX(i) * SigmaZ(i) * RhoXZ(i)
  vc(3, 1) = vc(1, 3)
  vc(2, 3) = SigmaY(i) * SigmaZ(i) * RhoYZ(i)
  vc(3, 2) = vc(2, 3)
  OmegaTerm = App.MInverse(vc)
  If IsError(vc) Then Exit Sub
  For j = 1 To 3
    For k = 1 To 3
      Omega(i, j, k) = OmegaTerm(j, k)
Next k, j, i
Erase sigmaX, SigmaY, SigmaZ
' Initial estimates of Intercepts & Slopes from simple linear regression
If ConcConstr Then
  ' Estimate age By calculating the wtd avg of the ages for each
  '  pt using the obs. 204/206-238/206 & assuming "typical"
  '  crustal common Pb.  Use the radiogenic 206 as the weighting factor.
  Comm46 = 1 / TuPbAlpha0: Comm76 = TuPbBeta0 * Comm46
  For i = 1 To N
    Sxz = (z(i) - Comm46) / X(i)
    Sxy = (y(i) - Comm76) / X(i)
    Pred86 = -Comm46 / Sxz
    If Pred86 = 0 Then Pred86 = -0.1
    tmp = 1 + 1 / Pred86
    If tmp < MINLOG Or tmp > MAXLOG Then
      PredT = 1000
    Else
      PredT = Log(tmp) / Lambda238
    End If
    wt = 1 / z(i) - Comm46
    SumWt = SumWt + wt
    Tsum = Tsum + PredT * wt
  Next i
  IntSl(1) = Comm76 ' Common-Pb 207/206
  IntSl(2) = Comm46 ' Common-Pb 204/206
  IntSl(3) = Tsum / SumWt ' Age
Else
  SimpleLinear X(), y(), z()    ' XY & XZ Intercept/Slope
End If
Randomize Timer
For i = 1 To 1 + Ndim ' Initialize simplex w. values scattered about the initial guesses.
  For j = 1 To Ndim
    If i = 1 Then
      Scoef(j) = IntSl(j)
    Else
      If ConcConstr Then
        k = Sgn(0.5 - Rnd)
        Scoef(j) = IntSl(j) * (1 + 2 * (0.5 - Rnd)) ^ k
      Else
        If IntSl(j) > 1 Then  ' Range of 2x to -1
          Scoef(j) = IntSl(j) * (0.5 + 3 * (0.5 - Rnd))
        Else
          Scoef(j) = 4 * (0.5 - Rnd) ' Range of +-2
        End If
      End If
    End If
    Simplex(i, j) = Scoef(j)
  Next j
  sY(i) = aSums(Scoef(), X(), y(), z(), Omega(), N, Failed)
  If Failed Then Exit Sub
Next i
Amoeba Simplex(), sY(), X(), y(), z(), Ndim, Omega(), N, 1 / MeanFractErr, _
   Toler, 9999, Failed
If Failed Then Exit Sub
For i = 1 To Ndim:  IntSl(i) = Simplex(1, i): Next i
Erase vc, Simplex, Scoef, sY
TrueX MSWD, Xt(), X(), y(), z(), Resid(), Omega(), N
' Now determine the elements of the Fisher Information Matrix (from the
'  2nd derivs of the expectations of -S/2).
a1 = IntSl(1)
If ConcConstr Then
  a2 = IntSl(2): T = IntSl(3)
Else
  b1 = IntSl(2): a2 = IntSl(3): b2 = IntSl(4)
End If
ReDim Delta(N, N) As Variant, M(Ndim, Ndim) As Variant, DeltaInv(N, N) As Variant
ReDim Bb(N, Ndim) As Variant, tmp1(Ndim, Ndim) As Variant, tmp2(Ndim, Ndim) As Variant
For i = 1 To N: For j = 1 To N: Delta(i, j) = 0: Next j, i ' Must assign zero values for matrix inversion
If ConcConstr Then
  FisherConstr M, Delta, Bb, Omega, Xt, N   ', bbT()
Else
  FisherUnconstr M, Delta, Bb, Omega, Xt, N ', bbT()
End If
With App
  bbT = .Transpose(Bb)
  DeltaInv = .MInverse(Delta) ' Can't do by calculating reciprocals 'cause changes variable type
  If IsError(DeltaInv) Then GoTo Failed
  ' f^-1 = (M - Bb^T Delta^-1 Bb)^-1
  tmp1 = .MMult(bbT, DeltaInv) ' Temp1= B^T Delta^-1
  tmp2 = .MMult(tmp1, Bb)      ' Temp1= B^T Delta^-1
  Finv = MatAddV(M, tmp2, -1, (Ndim), (Ndim))  ' Finv = M - B^T Delta^-1 B
  Erase Bb, bbT, DeltaInv
  MatCopy Finv, G
  f = App.MInverse(G)  ' f =(M - B^T Delta^-1 B)^-1
End With
If IsError(f) Then GoTo Failed
Tcheck Xt(), X(), y(), z(), Omega(), N, CheckMLE, Failed
If Failed Then Exit Sub
G = App.Transpose(f)
If Not ConcConstr Then
  ' if Y' = app' + b*X' where Y' = Y - Ybar & X' = X - Xbar, then
  '    Y  = app  + b*X  where app = app' + Ybar - b*Xbar
  ' Transform the intercepts to correspond to the original XYZ data
  a1 = a1 + Ybar - b1 * Xbar: a2 = a2 + Zbar - b2 * Xbar
  IntSl(1) = a1: IntSl(3) = a2
  ' Transform the variances/covariances to account for initial transformation
  '  (slope variances/covariances are not affected)
  ' g() is the variance/covariance matrix for the transformed XYZ data
  ' f()  will contain the variance/covariance matrix for the original data
  f(1, 1) = G(1, 1) + Xbar * Xbar * G(2, 2) - 2 * G(1, 2) * Xbar '_ app
  f(3, 3) = G(3, 3) + Xbar * Xbar * G(4, 4) - 2 * G(3, 4) * Xbar '_ app
  f(1, 2) = G(1, 2) - Xbar * G(2, 2) 'rho ab
  f(2, 1) = f(1, 2)                  ' "  "
  f(3, 4) = G(3, 4) - Xbar * G(4, 4) ' "  bB
  f(4, 3) = f(3, 4)                  ' "  "
  f(1, 3) = G(1, 3) - Xbar * (G(1, 4) + G(2, 3) - Xbar * G(2, 4)) 'rho aA
  f(3, 1) = f(1, 3)                                               ' "  "
  f(1, 4) = G(1, 4) - Xbar * G(2, 4) ' "  aB
  f(4, 1) = f(1, 4)                  ' "  "
  f(2, 3) = G(2, 3) - Xbar * G(2, 4) ' "  bA
  f(3, 2) = f(2, 3)                  ' "  "
End If
For i = 1 To Ndim
  If f(i, i) < 0 Then Failed = True: Exit Sub
Next i
If ErrMag Then
  MSWD = MSWD / SQ(MeanFractErr)
  For i = 1 To N
    Resid(i) = Resid(i) / MeanFractErr
  Next i
End If
For i = 1 To Ndim
  For j = 1 To Ndim
    If i = j Then
      VarCov(i, j) = f(i, j) * SQ(MeanFractErr)
      ErrRho(i, j) = Sqr(VarCov(i, j)) ' 1-sigma error
    Else
      ErrRho(i, j) = f(i, j) / Sqr(f(i, i) * f(j, j)) ' Error correl.
    End If
Next j, i
For i = 1 To Ndim
  For j = 1 To Ndim
    If i <> j Then VarCov(i, j) = ErrRho(i, j) * ErrRho(i, i) * ErrRho(i, j)
Next j, i
If Not ConcConstr Then
  For i = 1 To N ' Transform back to original coords
    X(i) = X(i) + Xbar: y(i) = y(i) + Ybar: z(i) = z(i) + Zbar
  Next i
End If
Exit Sub

NoErMag:  ErrMag = False
Xbar = 0: Ybar = 0: Zbar = 0
MeanFractErr = 1
Return

Failed: On Error GoTo 0
MsgBox "Sorry, unable to find a solution for these data", , Iso
ExitIsoplot
End Sub

Sub Useries3Diso(DP() As DataPoints, Np#(), xyProj#(), ByVal N&, Failed As Boolean)
Attribute Useries3Diso.VB_ProcData.VB_Invoke_Func = " \n14"
' Given set of X-Y-Z data pts, determine the MLE X-Y & X-Z regression-
'  line slopes & intercepts, as well as their errors & error correlations.
' Algorithm developed by Titterington & Ludwig in April, 1993.
Dim ErP%, Emult#, AP95%, MSWD#, tbx$
Dim i&, j&, k&, P&
Dim Count&, Iter%, Ap1%, Style%
Dim Solid%, Tsums#, uDF&
Dim z1#, z2#, Zs#, Xdelt#, Ydelt#
Dim xStart#, yStart#, xEnd#, yEnd#, Gamma#
Dim GammaErr#, Toler#, ThAge#
Dim AgeErr#, UageErr#, LageErr#, Prob#
Dim ThUpperAge#, ThLowerAge#, GammaUpperAge#
Dim GammaLowerAge#, u#, v#, VarCov#()
Dim Detr02_0#, Detr02err_0#
Dim X#(), y#(), z#(), Resid#()
'Dim Pslope#, PslopeErr#, Pinter#, Pintererr#
Dim Uis As Object, predX#, is2#, is3#, is4#
i = 4 + ConcConstr
ReDim X(N), y(N), z(N), Resid(N), IntSl(i), ErrRho(i, i)
Style = 0
If UseriesPlot Then
  If UsType = 3 Then Style = 1 '230/238-234/238-232/238 (3D Useries evolution)
  If UsType = 2 Then Style = 2 '232/238-234/238-230/238 (Osmond Type II)
  If UsType = 1 Then Style = 3 '232/238-230/238-234/238 (mod Osmond Type II)
  ' Style = 4 -> 238/232-234/232-230/232 (Rosholt Type II)
  ' Style = 5 -> 238/232-230/232-234/232 (mod. Rosholt Type II)
End If
Failed = False: Toler = 1E-16
If N < 3 Then Toler = Toler / Hun
Calc3Dline DP(), X(), y(), z(), N, Resid(), Toler, Failed, MSWD, VarCov()
If Failed Then
  If MsgBox("Unable to fit a line to these data." & viv$ & _
    "You may want to try again, since there is some randomness" & vbLf & _
    "in the starting values of the Simplex solution.", _
    vbOKCancel, Iso) <> vbOK Then ExitIsoplot
  Exit Sub
End If
uDF = 2 * N - 4 - ConcConstr
If MinProb = 0 Then MinProb = Val(Menus("MinProb"))
Prob = ChiSquare(MSWD, uDF)
If Prob > MinProb Then Emult = 1.96 Else Emult = StudentsT(uDF) * Sqr(MSWD)
ErP = (Prob < 0.1 Or Prob > 0.99) + (Prob < 0.01 Or Prob > 0.999) - 2
For i = 1 To 4 + ConcConstr           ' Convert errors to 95%-conf.
  ErrRho(i, i) = Emult * ErrRho(i, i)
Next i
For i = 1 To 4 + ConcConstr
  For j = 1 To 4 + ConcConstr
    If i = j Then
      VarCov(i, j) = SQ(ErrRho(i, j)) ' Note that VarCov now contains the
    Else                              '  "95%-conf." Var-Cov, not 1-sigma!
      VarCov(i, j) = ErrRho(i, j) * ErrRho(i, i) * ErrRho(j, j)
    End If
Next j, i
xyProj(0) = 2 ' Call "1sigma" half of 95% conf
If UseriesPlot Then
  ' xyProj(1 to 5) = Detrital-free 230/238,err,234/238,err,rho
  Select Case Style
    Case 1  ' X=230/238 Y=234/238 Z=232/238
      XYplaneInter xyProj(), VarCov()
    Case 2  ' X=232/238 Y=234/238 Z=230/238
      i = 3: j = 1: k = 4
    Case 3  ' X=232/238 Y=230/238 Z=234/238
      i = 1: j = 3: k = 2
'   Case 4  ' X=238/232 Y=234/232 Z=230/232
'     i = 4: j = 2: k = 3
'   Case 5  ' X=238/232 Y=230/232 Z=234/232
'     i = 2: j = 4: k = 1
  End Select
  If Style > 1 Then
    xyProj(1) = IntSl(i)     ' 230/238
    xyProj(3) = IntSl(j)     ' 234/238
    xyProj(2) = ErrRho(i, i) ' 2-sigma 230/238 err
    xyProj(4) = ErrRho(j, j) ' 2-sigma 234/238 err
    xyProj(5) = ErrRho(i, j) ' rho 230/238-234/238
    Detr02_0 = IntSl(k)        ' 230/238-232/238 slope
    Detr02err_0 = ErrRho(k, k)
  End If
  SinglePointThUage xyProj(1), xyProj(2) / Emult, xyProj(3), xyProj(4) / Emult, xyProj(5), _
    MSWD, Prob, Emult, Detr02_0, Detr02err_0 / Emult ' &&
ElseIf OtherXY Then
  Show3dLine MSWD, Prob
ElseIf (ConcPlot And Not ConcAge) Then
  Show3dConcLin MSWD, Prob, VarCov(), xyProj()
  If DoPlot And ConcPlot And ConcConstr Then
    Lir$ = VandE(IntSl(3), ErrRho(3, 3), 2)
    Msw$ = Mrnd(MSWD)
  End If
End If
If DoPlot Then
  If Style = 1 Or (ConcPlot And Not ConcAge) Then ' Optional projection of XYZ values to the XY plane
    If PlotProj Then                              '  parallel to the XYZ regression line.
      For i = 1 To N
        If ConcPlot And ConcConstr Then
          u = ConcX(IntSl(3), True, True): v = ConcY(IntSl(3), True, True)
          is3 = IntSl(2): is4 = -is3 / u
          is2 = ConcX(IntSl(3), 0, -1) / Uratio - IntSl(1) * ConcY(IntSl(3), 0, -1)
        Else
          u = -IntSl(3) / IntSl(4)    ' X at concordia-plane intercept
          v = IntSl(1) + IntSl(2) * u ' Y "    "
          is2 = IntSl(2): is3 = IntSl(3): is4 = IntSl(4)
        End If
        predX = (z(i) - is3) / is4
        Xdelt = X(i) - predX
        Ydelt = y(i) - (IntSl(1) + is2 * predX)
        Np(i, 1) = u + Xdelt:  Np(i, 3) = v + Ydelt
        Np(i, 2) = DP(i).Xerr: Np(i, 4) = DP(i).Yerr
        Np(i, 5) = DP(i).RhoXY
      Next i
    End If
  End If
End If
ReDim yf.WtdResid(N)
For i = 1 To N: yf.WtdResid(i) = Resid(i): Next i 'SQ(Resid(i)): Next i
End Sub

Private Sub Tcheck(Xt0#(), X#(), y#(), z#(), _
  Omega#(), ByVal N&, ByVal Check As Boolean, Failed As Boolean)
' Check the XY-XZ regression sol'n
Dim i&, k%, L%, a1#, a2#, Mult#
Dim b1#, b2#, Denom#, Numer1#, Numer2#
Dim NumerXY#, NumerXZ#, Numer#, cc#
Dim Diff#, Resid#, T#, H#(), G#()
Dim YS#(), xs#(), Zs#()
Dim Xt(), Left#(), Right#()
ReDim Xt(N)
If Check Then
  ReDim G(4, 2), YS(4), xs(4, 2), Zs(4), H(N, 2), Left(4), Right(4)
End If
a1 = IntSl(1)
If ConcConstr Then
  a2 = IntSl(2): T = IntSl(3)
  b1 = ConcX(T, False, True) / Uratio - a1 * ConcY(T, False, True)
  b2 = -a2 * ConcY(T, False, True)
Else
  b1 = IntSl(2)
  a2 = IntSl(3): b2 = IntSl(4)
End If
Failed = False
For i = 1 To N
  Denom = Omega(i, 1, 1) + b1 * b1 * Omega(i, 2, 2) + b2 * b2 * Omega(i, 3, 3)
  Denom = Denom + 2 * b1 * Omega(i, 1, 2) + 2 * b2 * Omega(i, 1, 3)
  Denom = Denom + 2 * b1 * b2 * Omega(i, 2, 3)
  Numer1 = Omega(i, 1, 2) + b1 * Omega(i, 2, 2) + b2 * Omega(i, 2, 3)
  Numer2 = Omega(i, 1, 3) + b1 * Omega(i, 2, 3) + b2 * Omega(i, 3, 3)
  NumerXY = (y(i) - a1 - b1 * X(i)) * Numer1
  NumerXZ = (z(i) - a2 - b2 * X(i)) * Numer2
  Numer = NumerXY + NumerXZ
  Xt(i) = X(i) + Numer / Denom  ' Must be same as Xt0(i)
  If Abs((Xt(i) - Xt0(i)) / Xt0(i)) > 0.000001 Then Failed = True:   Exit Sub
  If Check Then
    H(i, 1) = Numer1 / Denom
    H(i, 2) = Numer2 / Denom
  End If
Next i
If Check Then
  For k = 1 To 4
    L = IIf(k = 1 Or k = 3, 2, 3)
    For i = 1 To N
      cc = Omega(i, 1, L) + b1 * Omega(i, 2, L) + b2 * Omega(i, L, 3)
      If k < 3 Then Mult = 1# Else Mult = Xt(i)
      G(k, 1) = G(k, 1) + Mult * (Omega(i, 2, L) - H(i, 1) * cc)
      G(k, 2) = G(k, 2) + Mult * (Omega(i, L, 3) - H(i, 2) * cc)
      xs(k, 1) = xs(k, 1) + Mult * X(i) * (Omega(i, 2, L) - H(i, 1) * cc)
      xs(k, 2) = xs(k, 2) + Mult * X(i) * (Omega(i, L, 3) - H(i, 2) * cc)
      YS(k) = YS(k) + Mult * y(i) * (Omega(i, 2, L) - H(i, 1) * cc)
      Zs(k) = Zs(k) + Mult * z(i) * (Omega(i, L, 3) - H(i, 2) * cc)
    Next i
  Next k
  For k = 1 To 4
    Left(k) = G(k, 1) * a1 + G(k, 2) * a2
    Right(k) = YS(k) - b1 * xs(k, 1) + Zs(k) - b2 * xs(k, 2)
    Resid = Left(k) - Right(k)
    Diff = Diff + Abs(Resid)
  Next k
  If Diff > 0.01 Then Failed = True
End If
End Sub

Private Sub TrueX(MSWD#, Xt#(), X#(), y#(), z#(), _
  Resid#(), Omega#(), ByVal N&)
' Return "true" x-values for XY-XZ regression.  IntSl contains the XY,XZ
'  slope,inter,slope,inter; Omega() contains the inverted variance-covariance
'  matrix for each point.
Dim i&, a1#, a2#, Ry#, Rz#
Dim term1#, term2#, term3#, Term4#
Dim Alpha#, Beta#, Gamma#, SumWtdResidSq#
Dim b1#, b2#, T#, ba#
a1 = IntSl(1)
If ConcConstr Then
  a2 = IntSl(2): T = IntSl(3)
  b1 = ConcX(T, False, True) / Uratio - a1 * ConcY(T, False, True)
  b2 = -a2 * ConcY(T, False, True)
Else
  b1 = IntSl(2): b2 = IntSl(4): a2 = IntSl(3)
End If
For i = 1 To N
  Ry = y(i) - a1 - b1 * X(i)
  Rz = z(i) - a2 - b2 * X(i)
  term1 = Omega(i, 1, 1) + Omega(i, 2, 2) * b1 * b1
  term2 = 2 * Omega(i, 1, 2) * b1
  term3 = 2 * Omega(i, 1, 3) * b2 + Omega(i, 3, 3) * b2 * b2
  Term4 = 2 * Omega(i, 2, 3) * b1 * b2
  Alpha = term1 + term2 + term3 + Term4
  term1 = Omega(i, 1, 2) + b1 * Omega(i, 2, 2) + b2 * Omega(i, 2, 3)
  term2 = Omega(i, 1, 3) + b1 * Omega(i, 2, 3) + b2 * Omega(i, 3, 3)
  Beta = Ry * term1 + Rz * term2
  Gamma = Omega(i, 2, 2) * Ry * Ry + Omega(i, 3, 3) * Rz * Rz
  Gamma = Gamma + 2 * Omega(i, 2, 3) * Ry * Rz
  ba = -Beta / Alpha
  Xt(i) = X(i) - ba
  Resid(i) = Sqr(Gamma - Beta * Beta / Alpha) * Sgn(ba)
  SumWtdResidSq = SumWtdResidSq + SQ(Resid(i))
Next i
If N < 3 Then MSWD = 0 Else MSWD = SumWtdResidSq / (2 * N - 4 - ConcConstr)
End Sub

Private Function MatAddV(A As Variant, b As Variant, ByVal Add%, _
  ByVal Nrows&, ByVal Ncols%)
Dim c As Variant
ReDim c(Nrows, Ncols)
 ' C()=A()+Add*B()
Dim i&, j%
For i = 1 To Nrows
  For j = 1 To Ncols
    c(i, j) = A(i, j) + Add * b(i, j)
Next j, i
MatAddV = c
End Function

Private Sub FisherUnconstr(M As Variant, Delta As Variant, Bb As Variant, _
  Omega#(), Xt#(), ByVal N&) ', bbT#(),
' Determine the elements of the Fisher Information Matrix (from the
'  second derivatives of the expectations of -S/2) & thus the elements of the
'  Fisher Information Matrix. -- for calculation of an unconstrained 3-D regression
'  line a la Ludwig & Titterington, 1994.
Dim i&, j%
Dim Eta(), Om11, Om12, Om13, Om22, Om23, Om33, Xt2, Dum
Dim m111, m112, m122, m222, m211, m212, p11, p12, p21, p22
Dim a1, a2, b1, b2
ReDim Eta(2, N)
a1 = IntSl(1): a2 = IntSl(3): b1 = IntSl(2): b2 = IntSl(4)
For i = 1 To N
  Om11 = Omega(i, 1, 1): Om12 = Omega(i, 1, 2): Om13 = Omega(i, 1, 3)
  Om22 = Omega(i, 2, 2): Om23 = Omega(i, 2, 3): Om33 = Omega(i, 3, 3)
  Xt2 = SQ(Xt(i))
  m111 = m111 + Om22:       m112 = m112 + Xt(i) * Om22
  m122 = m122 + Xt2 * Om22: m222 = m222 + Xt2 * Om33
  m211 = m211 + Om33:       m212 = m212 + Xt(i) * Om33
  p11 = p11 + Om23:         p12 = p12 + Xt(i) * Om23
  p21 = p21 + Xt(i) * Om23: p22 = p22 + Xt2 * Om23
  Eta(1, i) = b1 * Om22 + Om12 + b2 * Om23
  Eta(2, i) = Om13 + b1 * Om23 + b2 * Om33
  Dum = Om11 + 2 * b1 * Om12 + 2 * b2 * Om13 + b1 * b1 * Om22 + b2 * b2 * Om33
  Delta(i, i) = Dum + 2 * b1 * b2 * Om23
  Bb(i, 1) = Eta(1, i): Bb(i, 2) = Xt(i) * Eta(1, i)
  Bb(i, 3) = Eta(2, i): Bb(i, 4) = Xt(i) * Eta(2, i)
  'For j = 1 To 4: bbT(j, i) = bB(i, j): Next j
Next i
Erase Eta
M(1, 1) = m111: M(1, 2) = m112: M(1, 3) = p11:  M(1, 4) = p12
M(2, 1) = m112: M(2, 2) = m122: M(2, 3) = p21:  M(2, 4) = p22
M(3, 1) = p11:  M(3, 2) = p21:  M(3, 3) = m211: M(3, 4) = m212
M(4, 1) = p12:  M(4, 2) = p22:  M(4, 3) = m212: M(4, 4) = m222
End Sub

Private Sub FisherConstr(M As Variant, Delta As Variant, Bb As Variant, _
  Omega#(), Xt#(), ByVal N&)  ', bbT#()
' Determine the elements of the Fisher Information Matrix (from the
'  second derivatives of the expectations of -S/2) & thus the elements of the
'  Fisher Information Matrix. -- for calculation of the concordia-constrained
'  3D regression line (ie a "Total U/Pb Isochron").
' See GCA v62, p665-676, 1998 for details
Dim aY#, aZ#, T#, test#
aY = IntSl(1): aZ = IntSl(2):  T = IntSl(3)
' Equations are: Y = Ay - [Cy(t) - a*Cx(t)]X
'                Z = Az - Az*Cx(t)*X
' where
'  Y=207/206,        X=238/206.         Z=204/206
'  Ay=Common 207/206, Az=Common 204/206
'  Cx(t)=EXP(Lambda238*T)-1
'  Cy(t)=EXP(Lambda235*T)-1
Dim i&, j%
Dim Bz, BzP, By, ByP, x2, Eta, Eta2, SumAyAy, SumAyAz, SumTT
Dim SumAzAz, SumAyT, SumAzT, Tm1, Tm2, Cx, cY, CxPrime, CyPrime
Dim Om11, Om12, Om13, Om21, Om22, Om23, Om33, X
' May run out of array-memory (generating SubscriptOutOfRange errors)
'  for Large N (>80 or 90).
test = Lambda235 * T
If Abs(test) > MAXEXP Then BadExp
Cx = ConcY(T, False, True):   cY = ConcX(T, False, True) / Uratio
CxPrime = Lambda238 * Exp(Lambda238 * T)
CyPrime = Lambda235 / Uratio * Exp(test)
Bz = -aZ * Cx:     BzP = -aZ * CxPrime
By = cY - aY * Cx: ByP = CyPrime - aY * CxPrime
' By, Bz are the X-Y & X-Z slopes of the regression line.
' Calculate the negative expectations of the second derivatives of the
'  likelihood function.
For i = 1 To N
  Om11 = Omega(i, 1, 1): Om12 = Omega(i, 1, 2)
  Om13 = Omega(i, 1, 3): Om22 = Omega(i, 2, 2)
  Om23 = Omega(i, 2, 3): Om33 = Omega(i, 3, 3)
  X = Xt(i): x2 = X * X
  Eta = 1# - Cx * X: Eta2 = Eta * Eta
  SumAyAy = SumAyAy + Eta2 * Om22
  SumAyAz = SumAyAz + Eta2 * Om23
  SumAyT = SumAyT + Eta * X * (ByP * Om22 + BzP * Om23)
  SumAzAz = SumAzAz + Eta2 * Om33
  SumAzT = SumAzT + Eta * X * (ByP * Om23 + BzP * Om33)
  SumTT = SumTT + x2 * (ByP * ByP * Om22 + 2 * ByP * BzP * Om23 + BzP * BzP * Om33)
  Tm1 = Om11 + 2 * By * Om12 + 2 * Bz * Om13
  Tm2 = By * By * Om22 + 2 * By * Bz * Om23 + Bz * Bz * Om33
  Delta(i, i) = Tm1 + Tm2
  Bb(i, 1) = Eta * (Om12 + By * Om22 + Bz * Om23)
  Bb(i, 2) = Eta * (Om13 + By * Om23 + Bz * Om33)
  Tm1 = ByP * Om12 + BzP * Om13 + By * ByP * Om22
  Tm2 = (ByP * Bz + By * BzP) * Om23 + Bz * BzP * Om33
  Bb(i, 3) = X * (Tm1 + Tm2)
  'For j = 1 To 3: bbT(j, i) = bB(i, j): Next j
Next i
M(1, 1) = SumAyAy
M(1, 2) = SumAyAz: M(2, 1) = M(1, 2)
M(1, 3) = SumAyT:  M(3, 1) = M(1, 3)
M(2, 2) = SumAzAz
M(2, 3) = SumAzT:  M(3, 2) = M(2, 3)
M(3, 3) = SumTT
End Sub

Sub Show3dLine(ByVal MSWD#, ByVal Prob#)
Attribute Show3dLine.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As Object, L As Object, G As Object, T As Object, nR As Boolean
Dim i%, j%, k%, s$, vv$, ee$
AssignD "3dLinRes", d, , , , L, G
nR = Not Robust
For i = 1 To 4
  NumAndErr IntSl(i), (ErrRho(i, i)), 2, vv$, ee$, True
  s$ = Chr$(96 + i) & " =" & vv$
  If Robust Then
    s$ = s$ & "   "
    With yf
      Select Case i
        Case 1
          s$ = s$ & Sd(.UpprInter - .Intercept, 2, True, True) & "/" & Sd(.LwrInter - .Intercept, 2, True, True)
        Case 2
          s$ = s$ & Sd(.UpprSlope - .Slope, 2, True, True) & "/" & Sd(.LwrSlope - .Slope, 2, True, True)
        Case 3
          s$ = s$ & Sd(.UpprZinter - .Zinter, 2, True, True) & "/" & Sd(.LwrZinter - .Zinter, 2, True, True)
        Case 4
          s$ = s$ & Sd(.UpprSlopeXZ - .SlopeXZ, 2, True, True) & "/" & Sd(.LwrSlopeXZ - .SlopeXZ, 2, True, True)
      End Select
    End With
  Else
    s$ = s$ & pm & ee$
  End If
  L(i).Text = s$
Next i
For j = 1 To 3
  For k = j + 1 To 4
     L(i).Text = Chr$(96 + j) & "," & Chr$(96 + k) & " = " & RhoRnd(ErrRho(j, k))
     i = i + 1
Next k, j
L("lprobcap").Visible = nR: L("lprob").Visible = nR
L("lMSWDcap").Visible = nR: L("lMSWD").Visible = nR
G("gMprob").Visible = nR
L(i + 1).Text = IIf(nR, Mrnd(MSWD), "")
L(i + 3).Text = IIf(nR, ProbRnd(Prob), "")
L("lErrLev").Text = "errors are " & IIf(Robust Or Prob < MinProb, "95% conf", "2 sigma")
ShowBox d, True
If IsOff(d.CheckBoxes("cShowRes")) Then Exit Sub
Set T = d.TextBoxes
s$ = T(1).Text & "         " & T(2).Text & vbLf
For i = 1 To 4: s$ = s$ & L(i).Text & vbLf: Next i
For i = 5 To 7
  j = i + 3
  vv$ = Space(Max(2, 17 - Len(L(i).Text)))
  s$ = s$ & "rho " & L(i).Text & vv$ & "rho " & L(j).Text
  If nR Or i < 7 Then s$ = s$ & vbLf
Next i
If nR Then
  For i = 1 To 3 Step 2
    s$ = s$ & L(i + 10).Text & "=" & L(i + 11).Text
    If i = 1 Then s$ = s$ & ",  "
  Next i
  s$ = s$ & vbLf & L("lErrLev").Text
End If
AddResBox s$, Clr:=RGB(0, 255, 255), FixedFont:=True
End Sub

Private Sub Show3dConcLin(ByVal MSWD#, ByVal Prob#, VarCov#(), _
  xyProj#())
' Put the results of the concordia-constrained 3D-linear regression results
'  in a dialog box.
Dim Comm64#, Comm74#, Ferr46#, Ferr76#, Ferr74#
Dim Err64#, Err74#, Rho6474#, tmp$
Dim Age#, PbAge#, Mu#, X#, y#, s#(4)
Dim P#(4), c As Object, d As Object, T As Object
Dim v#, SigRho%, i%, j%, tB$, df As Object
For i = 1 To 4 + ConcConstr
  P(i) = IntSl(i): s(i) = ErrRho(i, i)
Next i
i = 3 + ConcConstr
Comm64 = 1 / P(i):                  Comm74 = Comm64 * P(1)
Ferr46 = s(i) / P(i):               Ferr76 = s(1) / P(1)
Ferr74 = Sqr(Ferr46 * Ferr46 + Ferr76 * Ferr76 - 2 * Ferr46 * Ferr76 * ErrRho(1, i))
Err64 = Ferr46 * Comm64:            Err74 = Ferr74 * Comm74
Rho6474 = (Ferr46 - Ferr76 * ErrRho(1, i)) / Ferr74
If Comm64 > 12 And Comm64 < 100 And Comm74 < 20 And Comm74 > 12 Then
  CalcPbgrowthParams True
  SingleStagePbAgeMu Comm64, Comm74, PbAge, Mu
Else
  Age = 0: Mu = 0
End If
AssignD "ConcLinRes", d, , , , c, , T, Dframe:=df
df.Text = IIf(ConcConstr, "Total Pb/U Isochron Solution", "Unconstrained 3D Linear U-Pb Concordia Solution")
tmp$ = IIf(ConcConstr, "", "Unconstrained ")
d.GroupBoxes("gAge").Text = tmp$ & "Concordia-Plane Intercept"
For i = 1 To 3: c(i).Visible = Not ConcConstr: Next i
T("tAge").Visible = ConcConstr:   T("tDCE").Visible = ConcConstr
If ConcConstr Then
  Age = P(3)
  T("tAge").Text = "Age = " & VandE(Age, s(3), 2) & " Ma"
Else
  XYplaneInter xyProj(), VarCov()
  c(1).Text = "238U/206Pb = " & VandE(xyProj(1), xyProj(2), 2)
  c(2).Text = "207Pb/206Pb = " & VandE(xyProj(3), xyProj(4), 2)
  c(3).Text = "error correl. = " & RhoRnd(xyProj(5))
End If
c(4).Text = "206Pb/204Pb = " & VandE(Comm64, Err64, 2)
c(5).Text = "207Pb/204Pb = " & VandE(Comm74, Err74, 2)
c(6).Text = "error correl. = " & RhoRnd(Rho6474)
c(8).Text = Mrnd(MSWD): c(10).Text = ProbRnd(Prob)
If PbAge Then
  c(11).Text = "Stacey-Kramers Age = " & sn$(Int(PbAge)) & " Ma"
  c(12).Text = "Stacey-Kramers Mu  = " & Sp(Mu, -2)
Else
  c(11).Text = "": c(12).Text = ""
End If
Do
  ShowBox d, True
If Not AskInfo Then Exit Do
  ShowHelp "TotalPbUHelp"
Loop
If IsOn(d.CheckBoxes("cShowRes")) Then
  tB$ = IIf(ConcConstr, "Concordia-", "Un") & "constrained linear 3-D isochron" & vbLf
  If ConcConstr Then tB$ = tB$ & T("tAge").Text & vbLf
  tB$ = tB$ & "MSWD = " & c(8).Text & ",  probability =" & c(10).Text & vbLf
  If Not ConcConstr Then
    tB$ = tB$ & "Concordia plane intercepts at" & vbLf
    tB$ = tB$ & c(1).Text & vbLf & c(2).Text & vbLf & c(3).Text & vbLf
  End If
  tB$ = tB$ & "Common-Pb plane intercepts at" & vbLf
  tB$ = tB$ & c(4).Text & vbLf
  tB$ = tB$ & c(5).Text & vbLf & c(6).Text
  If PbAge Then tB$ = tB$ & vbLf & c(11).Text & vbLf & c(12).Text
  AddResBox tB$
End If
End Sub

Private Sub XYplaneInter(xyProj#(), vc#())
' Calculate the err-ellipse of the intercept of an unconstrained
'  3D line with the XY plane.
Dim k1#, k2#, k3#, k4#, k5#, k6#
Dim i%, j%, P#(4), p22#, p32#, p42#
Const Tiny = 1E-32
For i = 1 To 4: P(i) = IntSl(i): Next i
If P(4) = 0 Then P(4) = Tiny
p22 = SQ(P(2)): p32 = SQ(P(3)): p42 = SQ(P(4))
xyProj(1) = -P(3) / P(4)
xyProj(3) = P(1) + P(2) * xyProj(1)
k1 = vc(3, 3) + p32 / p42 * vc(4, 4) - 2 * P(3) / P(4) * vc(3, 4)
xyProj(2) = Max(Tiny, Sqr(k1 / p42))
k1 = vc(1, 1) + p32 / p42 * vc(2, 2) + p22 / p42 * vc(3, 3)
k2 = SQ(P(2) * P(3) / p42) * vc(4, 4)
k3 = -2 * P(3) / P(4) * vc(1, 2) - 2 * P(2) / P(4) * vc(1, 3)
k4 = 2 * P(2) * P(3) / p42 * vc(1, 4) + 2 * P(3) * P(2) / p42 * vc(2, 3)
k5 = -2 * P(2) * p32 / P(4) ^ 3 * vc(2, 4)
k6 = -2 * p22 * P(3) / P(4) ^ 3 * vc(3, 4)
xyProj(4) = Max(Tiny, Sqr(Max(0, k1 + k2 + k3 + k4 + k5 + k6)))
k1 = -vc(1, 3) / P(4) + P(3) / p42 * vc(3, 2) + P(2) / p42 * vc(3, 3)
k2 = -P(2) * P(3) / P(4) ^ 3 * vc(3, 4)
k3 = P(3) / p42 * vc(1, 4) - p32 / P(4) ^ 3 * vc(2, 4)
k4 = -P(2) * P(3) / P(4) ^ 3 * vc(3, 4) + P(2) * p32 / SQ(p42) * vc(4, 4)
xyProj(5) = (k1 + k2 + k3 + k4) / xyProj(2) / xyProj(4)
End Sub

Sub SinglePointThUage(Optional ThU, Optional ThUerr, Optional Gamma, _
  Optional GammaErr, Optional Rho, Optional MSWD, Optional Prob = -1, _
  Optional Emult = 2, Optional Detr02_nought, Optional Detr02_noughtErr)
Attribute SinglePointThUage.VB_ProcData.VB_Invoke_Func = " \n14"
'Calculate 230Th/U age and initial gamma from a single set of 230/238-234/238 ratios &errs.
' Errors input at 1-sigma a priori, with conversion to 95%-conf. via Emult.
Dim c As Object, o As Object, L As Object, EM#, Omc As Object
Dim MC As Boolean, Ntrials&, NLE As Boolean, AtomRat As Boolean, rU#()
Dim ThAge#, ThAgeErr#, UageErr#, LageErr#, Gamma0#
Dim Gamma0err#, Ug0err#, Lg0err#, RhoTG0#
ViM Emult, 2
AssignD "ThUage", , , c, o
Set Omc = o("MonteCarloU")
If N = 1 Or IM(Prob) Or Prob = -1 Then
  Omc.Enabled = True
Else
  Omc.Enabled = (Prob > 0.05)
End If
If Not Omc.Enabled Then Omc = xlOff: o("FirstDeriv") = xlOn
ThUAgeWLE
ShowBox DlgSht("ThUage"), True
MC = IsOn(o("MonteCarloU"))
c("Finite").Enabled = MC
ReDim rU(5 - 4 * MC)
NLE = IsOff(c("InclDCU")): AtomRat = IsOn(o("AtomRat"))
If IM(ThU) Then
  ThU = InpDat(1, 1): ThUerr = InpDat(1, 2)
  Gamma = InpDat(1, 3): GammaErr = InpDat(1, 4)
  Rho = InpDat(1, 5)
End If
' ThUerr, GammaErr at 1-sigma a priori!

ThUage_Gamma0 ThU, ThUerr, Gamma, GammaErr, Rho, MC, NLE, AtomRat, rU()
' Age error returned as 95%-conf (MC) or 2-sigma!
' 11/06/18 -- ru(8) and ru(9) are the median & mode ages

ThAge = rU(1): Gamma0 = rU(3 - MC)
RhoTG0 = rU(5 - 2 * MC)

If MC Then
  ThAgeErr = 0: Gamma0err = 0
  If rU(2) <> 0 And rU(3) <> 0 Then ThAgeErr = (rU(2) - rU(3)) / 2
  If rU(5) <> 0 And rU(6) <> 0 Then Gamma0err = (rU(5) - rU(6)) / 2
  UageErr = rU(2)
  LageErr = rU(3)
  Ug0err = rU(5)
  Lg0err = rU(6)

  If UageErr = 0 And ThAge <> 0 Then
    LageErr = LageErr - ThAge
    Ug0err = Ug0err - Gamma0
  End If

Else
  ThAgeErr = rU(2) / 2
  Gamma0err = rU(4) / 2  ' Convert back to 1-sigma
End If

ShowUiso MC, Not NLE, MSWD, Prob, ThU, ThUerr, Gamma, GammaErr, Rho, _
  Emult, ThAge, ThAgeErr, UageErr, LageErr, Gamma0, Gamma0err, Ug0err, Lg0err, _
  RhoTG0, Detr02_nought, Detr02_noughtErr
End Sub

Sub UevoProcT()
Attribute UevoProcT.VB_ProcData.VB_Invoke_Func = " \n14"
Dim b, P, s$, o As Object

With UevoT
  b = .cMultCurves:  P = .cLabelCurves: .cLabelCurves.Enabled = b
  .eLabelKa.Enabled = (b And P): .lKa.Enabled = (b And P)
  .lAtThe.Enabled = (b And P)
  uMultipleEvos = b
  LabelUcurves = (b And .cLabelCurves)
  b = Not .oNeither
  .gIsochrons.Enabled = b: .lTickInterval.Enabled = b
  .eTickInterval.Enabled = b: .cLabelTicks.Enabled = b
  .oInside.Enabled = (b And .cLabelTicks)
  .oOutside.Enabled = .oInside.Enabled
  If .oInside.Enabled Then UisochPos = .oInside - .oOutside
  uPlotIsochrons = .oIsochrons: uLabelTiks = .cLabelTicks
  LabelUcurves = (.cMultCurves And .cLabelCurves)
  .eLabelKa.Enabled = LabelUcurves: .lKa.Enabled = LabelUcurves
  uEvoCurvLabelAge = 0
  If LabelUcurves And IsNumeric(.eLabelKa.Text) Then
    If .eLabelKa <> "" And .eLabelKa <> "0" Then uEvoCurvLabelAge = Val(.eLabelKa)
  End If
  uUseTiks = .oAgeTicks:
  uFirstGamma0 = Val(.eGamma0_1): MaxAge = Val(.eMaxAge)
  If Not .oNeither Then
    s$ = Trim(.eTickInterval)
    If Len(s$) = 0 Or Not IsNumeric(s$) Then
      CurvTikInter = 0
    ElseIf s$ = "0" Then
      .eTickInterval = ""
    Else
      CurvTikInter = Val(s$)
    End If
  End If
End With
End Sub

Public Sub ThUAgeWLE() ' 230Th/U age-errors are affected by way in which the isotope
' ratios are calculated only if decay-const errs are considered, so don't show
' the option of choosing which type of ratios unless WLE
Dim b1 As Boolean, b2 As Boolean, o As Object, H%, spi As Object, Eb As Object, ff As Object
Dim La As Object, cb As Object, Grp As Object

AssignD "ThUage", , Eb, cb, o, La, Grp, , , spi, Dframe:=ff

b1 = IsOn(cb("InclDCU"))
Grp("gWithDCerrs").Visible = b1
o("SecEquil").Visible = b1
o("AtomRat").Visible = b1
b2 = IsOn(o("MonteCarloU"))
spi(1).Visible = b2
Eb(1).Visible = b2
La(1).Visible = b2
cb("Finite").Enabled = b2

With Grp("gWithDCerrs")
  ff.Height = .Top + IIf(b1, .Height + 3 - ff.Top, -5 - ff.Top)
End With

End Sub

Sub test()
Dim r, ct
'r = 1
Do
  r = r + 1
  ct = ct + 1
  If Trim(Cells(r, 1)) = "" Then
    Rows(r).Delete
    r = r - 1
  End If
Loop Until ct = 1700

End Sub
