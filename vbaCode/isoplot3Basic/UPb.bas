Attribute VB_Name = "UPb"
Option Private Module
Option Base 1: Option Explicit

Private Sub ConcordiaTicks()
Dim i%, NT%, MaxTix%
Dim FtOdd As Boolean, FnOdd As Boolean, ft$, fn$
Dim AgeRat#, k#, Zer#, c#, X#, y#
Dim Xfract#, Yfract#, Ftt#, Fnn#, tmp#
If Cdecay Then
  tmp = (AgeSpred / ((MaxAge + MinAge) / 2)) / (Lambda238err / Lambda238)
  MaxTix = 12 + 4 * (tmp < 60) + 2 * (tmp < 35)
Else
  MaxTix = 20
End If
Tick AgeSpred, CurvTikInter
CurvTikInter = Drnd(CurvTikInter, 2)
If MinAge <= 0 Then MinAge = 0.001
If Cdecay Then
  If AgeSpred / CurvTikInter > MaxTix Then Call Tick(AgeSpred * 2, CurvTikInter)
ElseIf AgeSpred / CurvTikInter < 8 And (Normal Or AgeSpred <= 2000) Then
  Tick AgeSpred / 2, CurvTikInter
End If
AgeSpred = MaxAge - MinAge: AgeRat = MaxAge / MinAge
If Inverse Then ' Fewer ticks for T-W if crowded
  If AgeRat > 5 And MaxAge > 1 Then Call Tick(AgeSpred / 4, CurvTikInter)
  Do
    If AgeSpred / CurvTikInter > MaxTix Then CurvTikInter = CurvTikInter * 2
  Loop Until AgeSpred / CurvTikInter <= MaxTix
End If
Xfract = Abs((ConcX(MaxAge) - ConcX(MinAge)) / Xspred)
Yfract = Abs((ConcY(MaxAge) - ConcY(MinAge)) / Yspred)
k = 0.6
For i = 1 To 4
  k = k / 2
  If Xfract < k And Yfract < k Then CurvTikInter = CurvTikInter * 2
Next i
If AutoScale Or True Then
  FirstCurvTik = Drnd(Int(MinAge / CurvTikInter) * CurvTikInter, 5)
  If FirstCurvTik = 0 Or (Inverse And FirstCurvTik < 0) Then FirstCurvTik = CurvTikInter
  If NumChars(FirstCurvTik + CurvTikInter) < NumChars(FirstCurvTik) Then
    FirstCurvTik = FirstCurvTik + CurvTikInter
  End If
End If
Ftt = Drnd(FirstCurvTik, 5): Fnn = Drnd(FirstCurvTik + CurvTikInter, 5)
ft$ = Str(Ftt): fn$ = Str(Fnn)
If Len(fn$) <= Len(ft$) Then
  ft$ = Right$(ft$, 1): fn$ = Right$(fn$, 1)
  If ft$ <> "0" And fn$ = "0" Then
    FirstCurvTik = Fnn
  Else
    FtOdd = False: FnOdd = False
    For i = 1 To 9 Step 2
      If Val(ft$) = i Then FtOdd = True
      If Val(fn$) = i Then FnOdd = True
    Next i
    If FtOdd And Not FnOdd Then FirstCurvTik = Fnn
  End If
End If
End Sub

Sub ConcordiaCurveData(Cv As Curves)
Attribute ConcordiaCurveData.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, r%, DoCurve As Boolean, db As Object
Dim StartX#, EndX#, StartAge#, EndAge#, Xinter#
Dim StartY#, EndY#, u#, v#, AgeStep#, T#
StatBar "creating Concordia curve"
' If an auto-scaled concordia age plot, use the Cdecay status set from
'  the "show age with..." optionbuttons.
If (ConcPlot And (Lambda235err > 0 Or Lambda238err > 0) And Cdecay0) Then
  If Not ConcAge Or Not AutoScale Then
    Set db = DlgSht
    If AutoScale And Regress Then
      Set db = db("isores").CheckBoxes("cWLE")
    Else
      Set db = db("concscale").CheckBoxes("cWLE")
    End If
    Cdecay = IsOn(db)
  End If
Else
  Cdecay = False
End If
StartX = MinX
If Inverse And StartX <= 0 Then StartX = ConcX(6000)
StartAge = ConcXage(StartX)
EndX = MaxX:  EndAge = ConcXage(EndX)
If AutoScale Then
  If Inverse Then StartAge = Min(StartAge, ConcYage(MaxY))
  StartY = ConcY(StartAge): EndY = ConcY(EndAge)
  If Inverse Then
    If StartY > MaxY Then
      StartY = MaxY: StartAge = ConcYage(StartY)
      StartX = ConcX(StartAge)
    End If
    If EndY < MinY Then
      EndY = MinY: EndAge = ConcYage(EndY)
      EndX = ConcX(EndAge)
    End If
  Else
    If StartY < MinY Then
      StartY = MinY: StartAge = ConcYage(StartY)
      StartX = ConcX(StartAge)
    End If
    If EndY > MaxY Then
      EndY = MaxY: EndAge = ConcYage(EndY)
      EndX = ConcX(EndAge)
    End If
  End If
End If
EndAge = Min(6000, EndAge)
u = EndX - StartX: v = Abs(EndY - StartY)
DoCurve = (u / Xspred > 0.05)
If Not DoCurve Then Ncurves = 0
If AutoScale Then MinAge = StartAge: MaxAge = EndAge
If Inverse And AutoScale Then Swap MinAge, MaxAge
MinAge = Max(0, MinAge)
AgeSpred = MaxAge - MinAge
If DoCurve Then
  ConcordiaTicks ' If Autoscale then
  StoreCurveData 1, Cv
End If
StatBar
End Sub

Sub ConcordiaIntercepts(ByVal Slope#, ByVal Intercept#, ConcInt#(), _
  Optional JustYoung As Boolean = False, Optional JustOld As Boolean = False)
Attribute ConcordiaIntercepts.VB_ProcData.VB_Invoke_Func = " \n14"
' Calculate intercepts of a best-fit line with the concordia curve.
' ConcInt(1) is lower intercept, ConcInt(2) is upper intercept.
Dim i%, j%, NoInter%, TT#, X#, T#
Dim Delta#, Numer#, Denom#, cs#
ViM JustYoung, False
ViM JustOld, False
If JustYoung And JustOld Then MsgBox "Error in call to ConcordiaIntercepts": ExitIsoplot
For i = 0 - JustOld To 1 + JustYoung
 If Normal Then
    j = i + 1
    TT = IIf(j = 1, -500, 5500)
  Else
    j = 2 - i
    TT = IIf(j = 1, 10, 5500)
  End If
  NoInter = 0
  Do
    cs = ConcSlope(TT)
    Numer = Intercept + cs * ConcX(TT) - ConcY(TT)
    Denom = cs - Slope
    X = Numer / Denom
  If X < -1 Then ConcInt(j) = BadT: NoInter = True: Exit Do
    T = ConcXage(X)
    Delta = Abs(T - TT)
  If Delta < 0.01 Then Exit Do
    TT = T
  Loop
  If Not NoInter Then ConcInt(j) = T
Next i
End Sub

Sub ConcInterErrors(ByVal Slope#, ByVal Inter#, ByVal SlopeErr#, _
  ByVal InterErr#, ByVal Xbar#, ConcInt#())
' Determine intercepts of hyperbolic error-envelope about the regression
'  line with the concordia curve.  The slope & intercept values & their errs
'  are given in the passed parameters.  Xbar is the wtd mean of the regression's
'  X-values; Normal=TRUE (Inverse=FALSE) if a Conv. concordia-plot, opposite
'  if a T-W plot.
' Returns the elements 3 to 6 of ConcInt() as, in order, the intercepts of
'  the lower & upper error-envelope arms about the younger regression-line
'  intercept, & the intercepts of the lower & upper error-envelope arms
'  about the older regression-line intercept.
Dim i%, j%, A%, TT#, T#, Del#, Delta#
Dim cs#, b#, c#, d#, e#, v#, discr#, x1#
Dim BadSqrt As Boolean
GetConsts
For i = 3 To 6
  j = IIf(Normal, i, i + IIf(i < 5, 2, -2))
  If (j < 5 And ConcInt(1) = BadT) Or (j > 4 And ConcInt(2) = BadT) Then _
    ConcInt(j) = BadT: GoTo NextArm
  Select Case j
    Case 3: TT = -1000      ' Lower intercept, younger age-limit
            If Inverse Then TT = 1
    Case 4: TT = ConcInt(1) ' Lower intercept, older      "
    Case 5: TT = ConcInt(2) ' Upper intercept, younger    "
    Case 6: TT = 6000       ' Upper intercept, older      "
  End Select
  A = IIf(j = 3 Or j = 5, -1, 1)
  Delta = 1E+37
  Do
    cs = ConcSlope(TT)
    b = ConcY(TT) - cs * ConcX(TT)
    d = 2 * ((b - Inter) * (cs - Slope) + Xbar * SlopeErr * SlopeErr)
    e = SQ(cs - Slope) - SlopeErr * SlopeErr
    v = SQ(b - Inter) - InterErr * InterErr
    TestSqrt d * d - 4 * e * v, discr, BadSqrt
    If BadSqrt Then ConcInt(j) = BadT: GoTo NextArm
    x1 = (A * discr - d) / (2 * e)
    If Normal Then
      If x1 < -1 Then GoTo BadInt
    Else
      If x1 = 0 Then GoTo BadInt
      If (1 / x1) < -1 Then GoTo BadInt
    End If
    T = ConcXage(x1)
    Del = Abs(T - TT)
    If Del > Delta Then ConcInt(j) = BadT: Exit Do
    If Del < 0.01 Then Exit Do
    Delta = Del:    TT = T
  Loop
  If ConcInt(j) <> BadT Then ConcInt(j) = T
NextArm:
Next i
If Inverse Then
  Swap ConcInt(3), ConcInt(4)
  Swap ConcInt(5), ConcInt(6)
End If
For i = 3 To 4  ' Same intercepts for both arms?  algorithm failed.
  If CSng(ConcInt(i)) = CSng(ConcInt(i + 2)) Then
    ConcInt(i) = BadT: ConcInt(i + 2) = BadT
  End If
Next i
Exit Sub
BadInt: ConcInt(j) = BadT: GoTo NextArm
End Sub

Sub PlotConcordiaBand(ByVal CurveColor&)
Attribute PlotConcordiaBand.VB_ProcData.VB_Invoke_Func = " \n14"
Dim tC&, i%, j%, CvWt&, CvSty&
StatBar "plotting concordia band"
If DoShape Then
  'IsoChrt.Select
  tC = IIf(ColorPlot, RGB(140, 0, 0), RGB(120, 120, 120))
  i = 0: On Error Resume Next
  i = CurvRange(2).Rows.Count
  On Error GoTo 0
  If i > 0 Then
    AddShape "ConcBand", CurvRange(2), tC, tC, False, 2 + 2 * BandBehind, , , , -1
    With Last(IsoChrt.Shapes)
      .ZOrder msoSendToBack
      If Not BandBehind Then .Fill.Transparency = 0.3
    End With
  End If
Else
  For i = 2 To 3
    'ChrtDat.Select
    Ach.SeriesCollection.Add CurvRange(i), xlColumns, False, True, False
    'IsoChrt.Select
    With Last(Ach.SeriesCollection)
      If .MarkerStyle <> xlNone Then .MarkerStyle = xlNone
      j = Opt.ConcLineThick
      If j = xlGray50 Then
        CvWt = xlThick: CvSty = j
      Else
        CvWt = xlThin: CvSty = xlContinuous 'In case corrupt MenutItems cell
        If j = xlHairline Or j = xlMedium Or j = xlThick Then CvWt = j
      End If
      .Border.Color = CurveColor
      With .Border
       If .Weight <> CvWt Then .Weight = CvWt
       If .LineStyle <> CvSty Then .LineStyle = CvSty
      End With
    End With
  Next i
End If ' (Doshape)
End Sub

Sub Proc3dU()
Attribute Proc3dU.VB_ProcData.VB_Invoke_Func = " \n14"
Dim W As Object, c As Object, o As Object, i%, d As Object, FC&
Dim b As Boolean, ft As Object, ce As Object, N1 As Boolean, G As Object
AssignD "3dU", W, , c, o, , G
Set d = W.DrawingObjects: Set ce = c("cInclEvo")
N1 = (N = 1 And Not DoPlot)
For i = 1 To 3
  If IsOn(o(i)) Then UsType = i: Dim3 = True: Inverse = True
Next i
Set ft = d("Isoch2d").Font
FC = IIf(Dim3, Gray50, Black)
ft.Color = FC
ft.Size = IIf(Mac, 11, 9)
With d("rec2Dbox")
  With .Border: .Color = FC: .Weight = xlThin: End With
  .Visible = Not N1
End With
o("oType1twoD").Enabled = Not N1
o("oType4").Enabled = Not N1
G("gOptions").Enabled = Not N1
d("Isoch2d").Visible = Not N1
d("Isoch2d").Visible = Not N1
If IsOn(o("oTypeMinus1")) Then
  UsType = -1: Dim3 = False: Inverse = True ' 2D 234/238 vs 230/238
ElseIf IsOn(o("oType1twoD")) Then
  UsType = 1: Dim3 = False: Inverse = True  ' 2D 230/238 vs 232/238
  AxY$ = "230Th/238U": AxX$ = "232Th/238U"
ElseIf IsOn(o("oType4")) Then
  UsType = 4: Dim3 = False: Inverse = False ' 2D 230/232 vs 238/232
  AxY$ = "230Th/232Th": AxX$ = "238U/232Th"
End If
b = ((UsType = -1 Or UsType = 3) And DoPlot And Not AddToPlot)
ce.Enabled = (b And Not N1)
If ce.Enabled Then uEvoCurve = IsOn(ce)
'W.GroupBoxes("gOther").Enabled = b  ?????????? no such Groupbox xcept in IsoSetup
ce = (b And uEvoCurve)
c("cPlotProj").Enabled = (DoPlot And Regress And UsType = 3 And Not N1)
If Not c("cPlotProj").Enabled Then c("cPlotProj") = xlOff
Normal = Not Inverse
End Sub

Private Sub ConcAgeNLE_click()
ConcAgeWLE_click
End Sub

Sub ConcIntAgeErrors(ByVal t1#, ByVal t2#, t1err#(), t2err#(), Bad() As Boolean)
Attribute ConcIntAgeErrors.VB_ProcData.VB_Invoke_Func = " \n14"
' Propagate (conventional) Concordia chord slope-intercept errs into (symmetric) concordia-
' intercept age errs using 1st-deriv expansion, both w. & w/o decay-const errs.
Dim e51#, e52#, e81#, e82#, Delta5#
Dim Slope#, SlopeErr#, Inter#, InterErr#
Dim CovSlopeInter#, m1#, m2#, b1#, b2#
Dim Coef#(3, 3), CoefI As Variant, yy#(3, 1), Tcoef As Variant
Dim e52m1#, Covt1t2#(2), L5e#, L8e#
Dim b3#, b4#, M3#, M4#, i%
Const MxT = 10000#, Niner = 0.99999999
Bad(1) = True: Bad(2) = True
If Abs(t1 - t2) < 0.001 Then Exit Sub
If t1 < -MxT Or t1 > MxT Or t2 < -MxT Or t2 > MxT Then Exit Sub
With yf
  Slope = .Slope:     SlopeErr = .SlopeError
  Inter = .Intercept: InterErr = .InterError
  If Abs(.RhoInterSlope) >= 1 Then
    .RhoInterSlope = Niner * Sgn(.RhoInterSlope)
  End If
  ' If passed as +1 or -1, can end up with negative variance sol'ns.
  CovSlopeInter = .RhoInterSlope * SlopeErr * InterErr
End With
e51 = Exp(Lambda235 * t1): e52 = Exp(Lambda235 * t2)
e81 = Exp(Lambda238 * t1): e82 = Exp(Lambda238 * t2)
Delta5 = e51 - e52:  e52m1 = e52 - 1
m1 = (Lambda238 * e81 - Slope * Lambda235 * e51) / Delta5
m2 = -(Lambda238 * e82 - Slope * Lambda235 * e52) / Delta5
b1 = -e52m1 * m1
b2 = -(Delta5 + e52m1) * m2
Coef(1, 1) = m1 * m1: Coef(1, 2) = m2 * m2: Coef(1, 3) = 2 * m1 * m2
Coef(2, 1) = b1 * b1: Coef(2, 2) = b2 * b2: Coef(2, 3) = 2 * b1 * b2
Coef(3, 1) = m1 * b1: Coef(3, 2) = m2 * b2: Coef(3, 3) = m1 * b2 + m2 * b1
yy(1, 1) = SlopeErr * SlopeErr: yy(2, 1) = InterErr * InterErr: yy(3, 1) = CovSlopeInter
For i = 1 To 2
  If i = 2 Then
    If Lambda235err = 0 Or Lambda238err = 0 Then Exit Sub
    ' Include decay-const errs (at 2-sigma)
    L5e = 2 * Lambda235err: L8e = 2 * Lambda238err
    M3 = Slope * (t2 * e52 - t1 * e51) / Delta5
    M4 = (t1 * e81 - t2 * e82) / Delta5
    b3 = -Slope * t2 * e52 - e52m1 * M3
    b4 = t2 * e82 - e52m1 * M4
    yy(1, 1) = yy(1, 1) + SQ(M3 * L5e) + SQ(M4 * L8e)
    yy(2, 1) = yy(2, 1) + SQ(b3 * L5e) + SQ(b4 * L8e)
    yy(3, 1) = yy(3, 1) + M3 * b3 * L5e * L5e + M4 * b4 * L8e * L8e
  End If
  With App          ' Solve the simultaneous eqns for t1-t2 variance/covariance
    CoefI = .MInverse(Coef) '  as a Fn of Slope-Inter variance/covariance.
    If IsError(CoefI) Then Exit Sub
    Tcoef = .MMult(CoefI, yy)
  End With
  If Tcoef(1, 1) < 0 Or Tcoef(2, 1) < 0 Then Exit Sub
  t1err(i) = Sqr(Tcoef(1, 1))
  t2err(i) = Sqr(Tcoef(2, 1))
  Covt1t2(i) = Tcoef(3, 1)
  Bad(i) = False
Next i
End Sub

Function ConcSlope(ByVal T#) ' Slope of U/Pb concordia curve at age T
Attribute ConcSlope.VB_ProcData.VB_Invoke_Func = " \n14"
Dim cs#, temp#, Eterm#, e5#
Eterm = (Lambda238 - Lambda235) * T
If Abs(Eterm) > MAXEXP Then ConcSlope = BadT: Exit Function
cs = Lambda238 * Exp(Eterm) / Lambda235
If Normal Then
  ConcSlope = cs
Else
  e5 = Lambda235 * T
  If Abs(e5) > MAXEXP Then BadExp
  temp = Exp(e5) - 1 - (Exp(Lambda238 * T) - 1) / cs
  ConcSlope = temp / Uratio
End If
End Function

Function ConcX(ByVal Age#, Optional TeWa, Optional NoGet = False)
Attribute ConcX.VB_ProcData.VB_Invoke_Func = " \n14"
' X-value for U/Pb concordia curve at specified age.
ViM NoGet, False
Dim Eterm#
If IM(TeWa) Then
  TeWa = Inverse
ElseIf Not NoGet Then
  GetConsts
End If
If TeWa Then
  Eterm = Lambda238 * Age
  If Abs(Eterm) > MAXEXP Then
    BadExp
  ElseIf Eterm = 0 Then
    ConcX = 1E+32
  Else
    ConcX = 1 / (Exp(Eterm) - 1)
  End If
Else
  Eterm = Lambda235 * Age
  If Abs(Eterm) > MAXEXP Then BadExp
  ConcX = Exp(Eterm) - 1
End If
End Function

Function ConcY(ByVal Age#, Optional TeWa, Optional NoGet = False)
Attribute ConcY.VB_ProcData.VB_Invoke_Func = " \n14"
Dim e8#, e5# ' Y-value for U/Pb concordia curve at specified age.
ViM NoGet, False
If IM(TeWa) Then TeWa = Inverse
If Not NoGet Then GetConsts
e8 = Lambda238 * Age
If TeWa Then
  e5 = Lambda235 * Age
  If Abs(e5) > MAXEXP Then
    BadExp
  ElseIf e5 = 0 Then  ' L'H�pital's rule
    ConcY = Lambda235 / Lambda238 / Uratio
  Else
    ConcY = (Exp(e5) - 1) / (Exp(e8) - 1) / Uratio
  End If
Else
  If Abs(e8) > MAXEXP Then BadExp
  ConcY = Exp(e8) - 1
End If
End Function

Function ConcXage(ByVal r#, Optional TeWa) ' Age for X-value R on U/Pb concordia curve
Attribute ConcXage.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Lterm#
If IM(TeWa) Then TeWa = Inverse
GetConsts
If TeWa Then Lterm = 1 + 1 / r Else Lterm = 1 + r
If Lterm < MINLOG Or Lterm > MAXLOG Then BadLog
Lterm = Log(Lterm)
If TeWa Then
  ConcXage = Lterm / Lambda238
Else
  ConcXage = Lterm / Lambda235
End If
End Function

Function ConcYage(ByVal r#, Optional TeWa) ' Age for Y-value R on U/Pb concordia curve
Attribute ConcYage.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Lterm#
If IM(TeWa) Then TeWa = Inverse Else Call GetConsts
If TeWa Then
  ConcYage = PbPbAge(r)
Else
  Lterm = 1 + r
  If Lterm < MINLOG Or Lterm > MAXLOG Then BadLog
  ConcYage = Log(Lterm) / Lambda238
End If
End Function

Function PbPbAge(ByVal Pb#, Optional t1 = 0, Optional iAge, _
  Optional Err76, Optional WithLambdaErrs = False) As Double
' Calculates age in Ma from radiogenic Pb-207/206 to t1 (=0 if not passed);
'  but if 2 more params (iAge & Err76) are passed, param Pb is the
'  7/6 ratio, iAge the age for that ratio, & Err76 is the absolute
'  error in the 7/6 ratio.  Uses Newton's method.
' If WithLambdaErrs=TRUE, include decay-constant errors (at global sigma-level)
'  in the age error.
Dim Exp5#, Exp8#, Numer#, Denom#, Func#
Dim T#, term1#, term2#, Deriv#, Delta#
Dim Pb76#, BadSqrt As Boolean, Exp5t1#, Exp8t1#
Dim CalcErr As Boolean, Test5#, Test8#, test#, P#
Const Toler = 0.00001
ViM t1, 0
ViM WithLambdaErrs, False
If NIM(Err76) And NIM(iAge) Then CalcErr = True
Pb76 = Pb
GetConsts
If CalcErr Then
  T = iAge
ElseIf Pb76 > (Lambda235 / Lambda238 / Uratio) Then ' 7/6 @t=0
  T = 1000
Else
  T = -4000 ' Need a trial age to start
End If
Test5 = Lambda235 * t1: Test8 = Lambda238 * t1
If Abs(Test5) > MAXEXP Or Abs(Test8) > MAXEXP Then GoTo PbFail
Exp5t1 = Exp(Test5): Exp8t1 = Exp(Test8)
Do
  Test5 = Lambda235 * T: Test8 = Lambda238 * T
  If Abs(Test5) > MAXEXP Or Abs(Test8) > MAXEXP Then GoTo PbFail
  Exp5 = Exp(Test5):  Exp8 = Exp(Test8)
  Numer = Exp5t1 - Exp5: Denom = Exp8t1 - Exp8
  If Denom = 0 Then GoTo PbFail
  Func = Numer / Denom / Uratio
  term1 = -Lambda235 * Exp5
  term2 = Lambda238 * Exp8 * Numer / Denom
  Deriv = (term1 + term2) / Denom / Uratio
  If Deriv = 0 Then GoTo PbFail
  If CalcErr Then
    If WithLambdaErrs And t1 = 0 Then
      Numer = SQ((Exp8 - 1) * Err76) + SQ(T * Exp5 * SigLev * Lambda235err / Uratio) + _
       SQ(Pb76 * T * Exp8 * SigLev * Lambda238err)
      Denom = SQ(Pb76 * Lambda238 * Exp8 - Lambda235 * Exp5 / Uratio)
      If Denom = 0 Then GoTo PbFail
      TestSqrt Numer / Denom, P, BadSqrt
      If BadSqrt Then GoTo PbFail
      PbPbAge = P
    Else
      PbPbAge = Abs(Err76 / Deriv)
    End If
    Exit Function
  ElseIf Deriv = 0 Then
    GoTo PbFail
  End If
  Delta = (Pb76 - Func) / Deriv
  T = T + Delta
Loop Until Abs(Delta) < Toler
PbPbAge = T
Exit Function
PbFail: PbPbAge = BadT
End Function

Sub ShowConcAge(ByVal T#, ByVal ErrT#, ByVal MswdOne#, _
  ByVal MswdMany#, ByVal tNLE#, ByVal ErrTNLE#, _
  ByVal MswdOneNLE#, ByVal MswdManyNLE#, ByVal Npts&)
' Show Concordia-Age results in a dialog box
Dim i&, L As Object, tmp, Mult95, Mult95nle, Prob, Mult
Dim p1#, p1NLE#, pMy#, pMyNle#
Dim s1%, c9%, dfT&
Dim s1N%, c9N%, Op As Object, AgeDispOK As Boolean
Dim tB As Object, NL As Boolean, CA$, caA$, caE$, caM$, caPr$, caEL%
Dim Grp As Object, ts$, ShowDat As Boolean, vv$, ee$, o As Object, cb As Object
Const Ptol = 0.001
AssignD "ConcAge", , , cb, Op, L, Grp, tB
AgeRes$ = ""
NL = (Lambda235err = 0 And Lambda238err = 0)
dfT = 2 * Npts - 1    ' for the weighted-mean concordant age
p1 = ChiSquare(MswdOne, 1):       p1NLE = ChiSquare(MswdOneNLE, 1)
pMy = ChiSquare(MswdMany, dfT): pMyNle = ChiSquare(MswdManyNLE, dfT)
c9 = (pMy < 0.3):   c9N = (pMyNle < 0.3)
s1 = (pMy >= 0.05): s1N = (pMyNle >= 0.05)
For Each o In L: o.Visible = True: o.Enabled = True: Next
For Each o In Grp: o.Enabled = True: o.Visible = True: Next
tB(1).Text = "1sigma a priori": tB(2).Text = "2sigma a priori"
tB(3).Text = "tsigmaSqrtMSWD"
With tB("tWLE"): .Text = "With lambda errors": .Visible = True: End With
With tB("tNLE"): .Text = "Without lambda errors": .Visible = True: End With
If Not Mac Then
  For i = 1 To 5
    With tB(i)
      .Font.Size = 11: .Font.Bold = (i > 3)
      .Font.Name = IIf(Mac, "Geneva", "Arial")
      ConvertSymbols tB(i)
    End With
  Next i
  On Error Resume Next
  For i = 1 To tB.Count
      With tB(i)
        If .Font.Name <> "Null" Then .Font.ColorIndex = xlNone
        .Visible = True
    End With
  Next i
End If
On Error GoTo 0
'1 age_WLE        2 age_NLE         3 MSWD =
'4 Probability =  5 1sigNLE         6 1sig WLE
'7 2sigNLE        8 2sig WLE        9 95 NLE
'10 95 WLE        11 m_wle         12 m_NLE
'13 p_wle         14 p_NLE         15 MSWD =
'16 Probability = 17 Age =         18 Age =
'19 MSWD =        20 Probability = 21 MSWD C+E_WLE
'22 Prob C+E_WLE  23 MSWD =        24 Probability =
'25 NSWD C+E_NLE  26 Prob C+E_NLE
If Npts = 1 Then
  Grp(4).Visible = False:  Grp(6).Visible = False
  L(23).Visible = False: L(24).Visible = False
  L(19).Visible = False: L(20).Visible = False
  L("lMswdConcEqWLE").Visible = False: L("lProbConcEqWLE").Visible = False
  L("lMswdConcEqNLE").Visible = False: L("lProbConcEqNLE").Visible = False
End If
If Not NL Then
  L("lMswdConcWLE").Text = Mrnd(MswdOne):    L("lProbConcWLE").Text = ProbRnd(p1)
  If Npts > 1 Then
    L("lMswdConcEqWLE").Text = Mrnd(MswdMany): L("lProbConcEqWLE").Text = ProbRnd(pMy)
  End If
End If
If Not Mac Then tB("tWLE").Font.ColorIndex = xlAutomatic
If pMy >= Ptol And Not NL Then ' Prob OK with lambda errors, can use lambda errors
  L("lAgeWLE").Visible = True:  L(6).Enabled = s1:  L("lAge2sigWLE").Enabled = s1
  NumAndErr T, 2 * ErrT, 2, vv$, ee$, , True
  L("lAgeWLE").Text = vv$ & " Ma":    L("lAge2sigWLE").Text = ee$
  L(6).Text = ErFo(T, ErrT, 2, True)
  'NumAndErr T, ErrT, 2, vv$, ee$, , True
  'l("lAgeWLE").Text = vv$ & " Ma":    L(6).Text = ee$
  'l("lAge2sigWLE").Text = ErFo(T, 2 * ErrT, 2, True)
  Mult95 = StudentsT(dfT) * Sqr(MswdMany)
  L("lAge95WLE").Visible = c9
  If c9 Then L("lAge95WLE").Text = ErFo(T, Mult95 * ErrT, 2, True)
Else                          ' Prob too low with lambda errors or no lambda errs only
  If NL Then
    L("lAgeWLE").Visible = 0
    L(6).Visible = 0:   L("lAge2sigWLE").Visible = 0
    L("lAge95WLE").Visible = 0:  L("lProbConcWLE").Visible = 0
    L("lProbConcEqWLE").Visible = 0:  L("lMswdConcWLE").Visible = 0
    L("lMswdConcEqWLE").Visible = 0:  L(3).Visible = 0
    L(4).Visible = 0:   L(17).Visible = 0
    L(19).Visible = 0:  L(20).Visible = 0
    Grp(3).Enabled = 0:   Grp(4).Enabled = 0
    If Not Mac Then tB("tWLE").Font.Color = Menus("cGray50")
  Else
    L("lAgeWLE").Text = "DISCORDANT"
  End If
End If
L("lMswdConcNLE").Text = Mrnd(MswdOneNLE):    L("lProbConcNLE").Text = ProbRnd(p1NLE)
If Npts > 1 Then
  L("lProbConcEqNLE").Text = ProbRnd(pMyNle):   L("lMswdConcEqNLE").Text = Mrnd(MswdManyNLE)
End If
If pMyNle < Ptol Then
  L("lAgeNLE").Text = "DISCORDANT"
  L("lAge1sigNLE").Visible = 0: L("lAge2sigNLE").Visible = 0: L("lAge95NLE").Visible = 0
Else
  NumAndErr tNLE, 2 * ErrTNLE, 2, vv$, ee$, , True
  L("lAgeNLE").Text = vv$ & " Ma": L("lAge2sigNLE").Text = ee$
  L("lAge1sigNLE").Text = ErFo(tNLE, ErrTNLE, 2, True)
  'NumAndErr tNLE, ErrTNLE, 2, vv$, ee$, , True
  'l("lAgeNLE").Text = vv$ & " Ma":      l("lAge1sigNLE").Text = ee$
  'l("lAge2sigNLE").Text = ErFo(tNLE, 2 * ErrTNLE, 2, True)
  L("lAge1sigNLE").Enabled = s1N:   L("lAge2sigNLE").Enabled = s1N
  Mult95nle = StudentsT(dfT) * Sqr(MswdManyNLE)
  L("lAge95NLE").Text = ErFo(tNLE, Mult95nle * ErrTNLE, 2, True)
  L("lAge95NLE").Visible = c9N
End If
If Not Mac Then
  For i = 1 To 2
    tB(i).Font.Color = IIf(pMy < Ptol And pMyNle < Ptol, Menus("cGray50"), vbBlack)
  Next i
  tB(3).Font.Color = IIf(Not c9 And (Not c9N Or NL Or Cmisc.NoLerr), Menus("cGray50"), vbBlack)
End If
For i = 5 To 9 Step 2
  If pMyNle < Ptol Or tNLE = 0 Then L(i).Visible = False
  If pMy < Ptol Or Cmisc.NoLerr Then L(i + 1).Visible = False
Next i
If tNLE = 0 Then
  L("lAgeNLE").Visible = False:  L("lMswdConcNLE").Visible = False: L("lProbConcNLE").Visible = False
  L("lMswdConcEqNLE").Visible = False: L("lProbConcEqNLE").Visible = False
End If
If Cmisc.NoLerr Then
  L("lAgeWLE").Visible = False:  L("lMswdConcWLE").Visible = False: L("lProbConcWLE").Visible = False
  L("lMswdConcEqWLE").Visible = False: L("lProbConcEqWLE").Visible = False
End If
'for i=1 to op.count:?i,op(i).text:next i
' 1  with decay-const errs   2  w/o decay-const errs
' 3  at 1 - Sigma            4  at 2 - Sigma
' 5  at 95%-conf.
'For i = 1 To Tb.Count: Print i, Tb(i).Text: Next i
' 1  1s a priori             2  2s a priori
' 3  tsigma�MSWD             4  With lambda errors
' 5  Without lambda errors
AgeDispOK = (DoPlot And ((pMy >= Ptol And Not NL And Not Cmisc.NoLerr) Or pMyNle >= Ptol))
For i = 7 To 8
  Grp(i).Visible = True 'AgeDispOK
Next i
For i = 1 To 5
  Op(i).Visible = True 'AgeDispOK
Next i
If AgeDispOK Then
  For i = 1 To 5: Op(i).Enabled = True: Next i
  If NL Or Cmisc.NoLerr Then
    Op(1).Enabled = False: Op(1) = xlOff
  Else
    Op(1) = xlOn
  End If
  If (NL Or Cmisc.NoLerr) And IsOff(Op(1)) Then Op(2).Enabled = True: Op(2) = xlOn
  If (NL Or Cmisc.NoLerr And Not c9N) Or (Not NL And Not Cmisc.NoLerr And Not c9) Then
    ' Prob >0.3 -- disable 95%-conf, select SigLev sigma level
    Op(2 + SigLev) = xlOn: Op(5 - SigLev) = xlOff: Op(5) = xlOff: Op(5).Enabled = False
  ElseIf (NL Or Cmisc.NoLerr And s1N) Or (Not NL And Not Cmisc.NoLerr And s1) Then
    ' Prob >.05 but <0.3 -- enable 95%-conf, select SigLev sigma level
    Op(2 + SigLev) = xlOn: Op(5 - SigLev) = xlOff: Op(5) = xlOff:
  Else
    ' Prob <.05 -- disable 1&2 sigma, enable & select 95%-conf
    Op(3).Enabled = False: Op(4).Enabled = False:  Op(5) = xlOn
  End If
End If
If (pMy >= Ptol And Not NL And Not Cmisc.NoLerr) Or pMyNle >= Ptol Then
  For i = 1 To 3: tB(i).Visible = True: Next i
End If
With Op("oBehind")
  .Visible = (DoShape And Cdecay0 And DoPlot And AutoScale)
  Op("oFront").Visible = .Visible: Grp("gConcBand").Visible = .Visible
  .Enabled = (IsOn(Op("oWLE")) And DoShape And Cdecay0)
  Op("oFront").Enabled = .Enabled
  Grp("gConcBand").Enabled = .Visible
End With
Do
  ShowBox DlgSht("ConcAge"), True
If Not AskInfo Then Exit Do
  Caveat_ConcAge
Loop
ShowDat = IsOn(cb(1))
If Not ShowDat And Not AgeDispOK Then Exit Sub
BandBehind = (IsOn(Op("oBehind")) And IsOn(Op("oWLE")) And DoShape And Cdecay0)
If IsOn(Op(1)) Then
  CA$ = L("lAgeWLE").Text
  Cdecay = True
  If IsOn(Op(3)) Then
    caE$ = L(6).Text
  ElseIf IsOn(Op(4)) Then
    caE$ = L("lAge2sigWLE").Text
  Else
    caE$ = L("lAge95WLE").Text
  End If
  caM$ = L("lMswdConcWLE").Text: caPr$ = L("lProbConcWLE").Text
Else
  CA$ = L("lAgeNLE").Text
  Cdecay = False
  If IsOn(Op(3)) Then
    caE$ = L("lAge1sigNLE").Text
  ElseIf IsOn(Op(4)) Then
    caE$ = L("lAge2sigNLE").Text
  Else
    caE$ = L("lAge95NLE").Text
  End If
  caM$ = L("lMswdConcNLE").Text: caPr$ = L("lProbConcNLE").Text
End If
caEL = -IsOn(Op(3)) - 2 * IsOn(Op(4)) - 3 * IsOn(Op(5))
If CA$ = "DISCORDANT" Then
  ts$ = "Data are not concordant":  CA$ = ts$ & vbLf & "(":  caE$ = ""
Else
  i = InStr(CA$, " Ma")
  CA$ = "Concordia Age = " & Left$(CA$, i - 1) & " " & caE$ & " Ma"
  CA$ = CA$ & vbLf & "("
  CA$ = CA$ & IIf(caEL = 3, "95% confidence, ", sn$(caEL) & "-sigma, ")
End If
CA$ = CA$ & "decay-const. errs " & IIf(IsOn(Op(1)), "included", "ignored")
CA$ = CA$ & ")" & vbLf & "MSWD (of concordance) = " & caM$ & "," _
  & vbLf & "Probability (of concordance) = " & caPr$
ts$ = CA$: AgeRes$ = CA$
If IsOn(cb(1)) Then AddResBox ts$
End Sub

Sub CreateConcordiaBandShapes(Cv As Curves, ByVal CurveColor&, AgeEllipseLimits#())
Attribute CreateConcordiaBandShapes.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, k%, v1#, v2#, EllRange As Range, tC&, tB As Boolean, AEL#(3), H%
If DoShape And Cv.Ncurvtiks > 0 Then ReDim AgeEllipseLimits(Cv.Ncurvtiks, 3)
GetScale
For i = 1 To Cv.Ncurvtiks
  If Not Cv.CurvTikPresent(i) Then GoTo NextTik
  StatBar "plotting concordia age-ellipse " & sn$(i)
  On Error GoTo NextTik ' In case wasn't within plotbox, hence never created
  Set EllRange = ChrtDat.Range("ConcAgeTik" & sn$(i))
  On Error GoTo 0
  tB = True
  If DoShape Then
    For j = 1 To EllRange.Rows.Count
      v1 = EllRange(j, 1): v2 = EllRange(j, 2)
      If v1 < MinX Or v1 > MaxX Or v2 < MinY Or v2 > MaxY Then tB = False
    Next j
    If tB Then
      tC = IIf(ColorPlot, vbWhite, vbWhite)
      With EllRange
        tB = IIf(.Cells(.Rows.Count + 2, 1) = "ClippedEllipse", False, True)
      End With
      AddShape "ConcTikEll", EllRange, tC, Black, tB, 2 + 2 * BandBehind, , , , 0.25, AEL()
      For j = 1 To 3: AgeEllipseLimits(i, j) = AEL(j): Next j
    End If
  Else
    IsoChrt.Select
    Ach.SeriesCollection.Add EllRange, xlColumns, False, True, False
    Nser = Ach.SeriesCollection.Count
    With Ach.SeriesCollection(Nser)
      .Border.Color = CurveColor
      With .Border
        .LineStyle = xlContinuous
        H = Opt.ConcLineThick
        .Weight = IIf(H = xlHairline, xlHairline, xlThin)
      End With
      .MarkerStyle = xlNone: .Smooth = True
      If Opt.ClipEllipse Then Call EllipseClip(Ach.SeriesCollection(Nser), EllRange)
      LineInd EllRange, "ConcTikEll"
    End With
  End If
NextTik: On Error GoTo 0
Next i
End Sub

Function PbT(ByVal PbX_Pb204ratio#, ByVal WhichRatio%)
Attribute PbT.VB_ProcData.VB_Invoke_Func = " \n14"
Dim T#, wr% ' Age from Pb-isotope ratio, assuming single-stage growth.
wr = WhichRatio
T = PbExp(wr) - (PbX_Pb204ratio - PbR0(wr)) / MuIsh(wr)
If T < MINLOG Or T > MAXLOG Then PbT = 0 Else PbT = Log(T) / PbLambda(wr)
End Function

Function PbR(ByVal Age, ByVal WhichRatio%)
Attribute PbR.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Eterm#, wr%  ' Pb-isotope ratio from age, assuming single-stage growth
wr = WhichRatio
Eterm = PbLambda(wr) * Age
If Abs(Eterm) > MAXEXP Then BadExp
PbR = PbR0(wr) + MuIsh(wr) * (PbExp(wr) - Exp(Eterm))
End Function

Sub SingleStagePbAgeMu(ByVal Alpha#, ByVal Beta#, Age#, Mu#)
Attribute SingleStagePbAgeMu.VB_ProcData.VB_Invoke_Func = " \n14"
' Age-Mu from Pb-isotope ratio, assuming single-stage growth.
Dim k5#, k8#, Pslope#, Pinter#, test#
Dim Alpha1#, Beta1#, Alpha2#, Beta2#, Tangent#
Dim T#, t1#, TangentInter#, Count%
k5 = Lambda235 * pbStartAge
If k5 > (MAXEXP) Then GoTo NoConv
k8 = Exp(Lambda238 * pbStartAge): k5 = Exp(k5)
Pslope = (Beta - pbBeta0) / (Alpha - pbAlpha0) ' Line slope to initial mantle Pb
Pinter = pbBeta0 - Pslope * pbAlpha0           ' Intercept ...
Alpha1 = Alpha
test = k8 - (Alpha1 - pbAlpha0) / pbMu
If test > MAXLOG Or test < MINLOG Then GoTo NoConv
t1 = (Log(test)) / Lambda238
test = Lambda235 * t1
If Abs(test) > MAXEXP Then GoTo NoConv
Beta1 = pbBeta0 + pbMu / Uratio * (k5 - Exp(test))
Count = 0
Do
  Count = 1 + Count
  test = (Lambda235 - Lambda238) * t1
  If Count > 50 Or Abs(test) > MAXEXP Then GoTo NoConv
  Tangent = Exp(test) / Uratio
  TangentInter = Beta1 - Tangent * Alpha1
  Alpha2 = (TangentInter - Pinter) / (Pslope - Tangent)
  Beta2 = Pslope * Alpha2 + Pinter
  test = k8 - (Alpha2 - pbAlpha0) / pbMu
  If test > MAXLOG Or test < MINLOG Then GoTo NoConv
  T = Log(test) / Lambda238
If Abs(t1 - T) < 0.1 Then Exit Do
  Alpha1 = Alpha2
  Beta1 = pbBeta0 + pbMu / Uratio * (k5 - Exp(Lambda235 * T))
  t1 = T
Loop
Mu = (Alpha - pbAlpha0) / (k8 - Exp(Lambda238 * T))
Age = T
Exit Sub
NoConv: Age = 0: Mu = 0
End Sub

Sub GrowthInters(ByVal Slope#, ByVal Inter#, PbInt#())
Attribute GrowthInters.VB_ProcData.VB_Invoke_Func = " \n14"
' Calculate intercepts of a Pb 207/204-206/204 line with a single-stage
'   Pb-growth curve.
Dim d(0 To 1), i%, j%, Count%, Ninters%
Dim TrialT#, c#, X#, T#, Eterm#
If MuIsh(0) = 0 Or PbLambda(0) = 0 Then CalcPbgrowthParams True
Ninters = 0
For j = 0 To 1
  TrialT = 6500 * j - 1000: Count = 0
  Do
    For i = 0 To 1
      Eterm = PbLambda(i) * TrialT
      If Abs(Eterm) > MAXEXP Then PbInt(j + 1) = BadT: GoTo NextJ
      d(i) = -MuIsh(i) * PbLambda(i) * Exp(Eterm)
    Next i
    c = d(1) / d(0)
    X = (Inter + c * PbR(TrialT, 0) - PbR(TrialT, 1)) / (c - Slope)
    Count = 1 + Count
    If X < -1 Or Count > 20 Then PbInt(j + 1) = BadT: GoTo NextJ
    T = PbT(X, 0)
    If Abs(T - TrialT) < 0.01 Then Exit Do
    TrialT = T
  Loop
  PbInt(j + 1) = T
  Ninters = 1 + Ninters
NextJ:
Next j
If Ninters = 1 Then
  If PbInt(1) = BadT Then Swap PbInt(1), PbInt(2)
ElseIf Ninters = 2 Then
  If Int(PbInt(1)) = Int(PbInt(2)) Then Ninters = 1
End If
End Sub

Sub ConcBandShapeRange(ShpRange As Range, ByVal Nlo%, ByVal Nhi%, _
  CurveLo#(), CurveHi#())
Attribute ConcBandShapeRange.VB_ProcData.VB_Invoke_Func = " \n14"
' Create the range defining the Concordia band shape
Dim i%, j%, k%
Dim Corner(2) As Boolean, x1(2), y1(2), x2(2), y2(2), Lc(2), rc(2), RevCurvHi(), Cend(2)
ReDim RevCurvHi(Nhi, 2)  ' Do corner-points need to be added?
x1(1) = CurveLo(1, 1):   y1(1) = CurveLo(1, 2)
x1(2) = CurveLo(Nlo, 1): y1(2) = CurveLo(Nlo, 2)
x2(1) = CurveHi(1, 1):   y2(1) = CurveHi(1, 2)
x2(2) = CurveHi(Nhi, 1): y2(2) = CurveHi(Nhi, 2)
Cend(1) = x1(1): Cend(2) = y1(1)   ' ending point
CornerPoints x1(), y1(), x2(), y2(), Lc(), rc(), Corner()
For i = 1 To Nhi   ' Reverse the order of the second curve,
  j = Nhi + 1 - i  '  then add to the tail of the first.
  For k = 1 To 2
    RevCurvHi(j, k) = CurveHi(i, k)
Next k, i
SymbRow = Max(1, SymbRow)
sR(SymbRow, SymbCol, Nlo - 1 + SymbRow, 1 + SymbCol) = CurveLo
k = SymbRow + Nlo
If Corner(2) Then sR(k, SymbCol, k, 1 + SymbCol) = rc: k = 1 + k
j = k: k = j + Nhi - 1
sR(j, SymbCol, k, 1 + SymbCol) = RevCurvHi: k = 1 + k
If Corner(1) Then sR(k, SymbCol, k, 1 + SymbCol) = Lc: k = 1 + k
sR(k, SymbCol, k, 1 + SymbCol) = Cend ' Close the curve
Set ShpRange = sR(SymbRow, SymbCol, k, 1 + SymbCol)
End Sub

Sub CalcPbgrowthParams(Optional SetObjects = False)       ' Handle Single-stage Pb-growth plot dialog-box click
Attribute CalcPbgrowthParams.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, sk As Object, Eb As Object, test
ViM SetObjects, False
If SetObjects Then
  GetConsts
  Set skV = Menus("StaceyKramers").Cells
Else
  With DlgSht("PbGrowth").EditBoxes
    pbAlpha0 = Val(.Item(1).Text): pbBeta0 = Val(.Item(2).Text)
    pbGamma0 = Val(.Item(3).Text)
    pbMu = Val(.Item(4).Text):     pbKappaMu = Val(.Item(5).Text)
    pbStartAge = Val(.Item(6).Text)
  End With
End If
If PbLambda(2) = Empty Then
  PbLambda(0) = Lambda238: PbLambda(1) = Lambda235
  PbLambda(2) = Menus("Lambda232").Cells.Value * Million
End If
If pbAlpha0 = Empty Then
  pbAlpha0 = Val(skV(1)): pbBeta0 = Val(skV(2)):   pbGamma0 = Val(skV(3))
  pbMu = Val(skV(4)):     pbKappaMu = Val(skV(5)): pbStartAge = Val(skV(6))
End If
ParRat(0) = 1: ParRat(1) = Uratio: ParRat(2) = 1 / pbKappaMu
PbR0(0) = pbAlpha0: PbR0(1) = pbBeta0: PbR0(2) = pbGamma0
For i = 0 To 2: MuIsh(i) = pbMu / ParRat(i): Next i
For i = 0 To 2: PbExp(i) = Exp(pbStartAge * PbLambda(i)): Next i
End Sub

Private Sub ProcPbGrowthTicks()
Dim P As Object: Set P = DlgSht("PbGrowth").CheckBoxes
If P(2).Enabled Then PbTickLabels = IsOn(P(2))
P(2).Enabled = IsOn(P(1)): P(2) = xlOff
If P(2).Enabled And PbTickLabels Then P(2) = xlOn
End Sub

Sub MNBRAK(aX#, BX#, Cx#, _
  FA#, FB#, FC#, MNBAD As Boolean)
' Given initial ages AX & BX, search in the downhill direction for new
'  ages AX, BX, CX whose ConcordSums FA, FB, FC bracket a minimum.
'  Adapted from Numerical Recipes, p. 281.
Dim r#, q#, test#, tmp1#, tmp2#, u#
Dim ULIM#, FU#, Bad As Boolean
Const Gold = 1.618034, GLIMIT = 100, Tiny = 1E-20, Big = 1E+38
MNBAD = False
FA = ConcordSums(aX, Bad): If Bad Then FA = Big
FB = ConcordSums(BX, Bad): If Bad Then FB = Big
If FB > FA Then
  Swap aX, BX
  Swap FA, FB
End If
Cx = BX + Gold * (BX - aX)
FC = ConcordSums(Cx, Bad)
If Bad Then
  If FA = Big And FB = Big Then MNBAD = True: Exit Sub
Else
  FC = Big
End If
Mnbrak1:
If FB >= FC Then
  r = (BX - aX) * (FB - FC)
  q = (BX - Cx) * (FB - FA)
  test = q - r
  tmp1 = BX - ((BX - Cx) * q - (BX - aX) * r)
  tmp2 = 2# * SIGN(Max(Abs(q - r), Tiny), q - r)
  u = tmp1 / tmp2
  ULIM = BX + GLIMIT * (Cx - BX)
  If ((BX - u) * (u - Cx)) > 0 Then
    FU = ConcordSums(u, Bad): If Bad Then FU = Big
    If FU < FC Then
      aX = BX: FA = FB
      BX = u:  FB = FU
      GoTo Mnbrak1
    ElseIf FU > FB Then
      Cx = u:  FC = FU
      GoTo Mnbrak1
    End If
    u = Cx + Gold * (Cx - BX)
    FU = ConcordSums(u, Bad): If Bad Then FU = Big
  ElseIf ((Cx - u) * (u - ULIM)) > 0 Then
    FU = ConcordSums(u, Bad): If Bad Then FU = Big
    If FU < FC Then
      BX = Cx: Cx = u
      u = Cx + Gold * (Cx - BX)
      FB = FC: FC = FU
      FU = ConcordSums(u, Bad): If Bad Then FU = Big
    End If
  ElseIf ((u - ULIM) * (ULIM - Cx)) >= 0 Then
    u = ULIM
    FU = ConcordSums(u, Bad): If Bad Then FU = Big
  Else
    u = Cx + Gold * (Cx - BX)
    FU = ConcordSums(u, Bad): If Bad Then FU = Big
  End If
  aX = BX: BX = Cx: Cx = u
  FA = FB: FB = FC: FC = FU
  GoTo Mnbrak1
End If
End Sub

Sub BRENT(ByVal aX#, ByVal BX#, ByVal Cx#, ByVal Tol#, _
  Xmin#, Fmin#, Bad As Boolean)
' Given a bracketing triplet of abscissas AX, BX, CX such that BX is
'  between AX & CX & Sums(BX) is less than both Sums(AX) & Sums(BX),
'  this routine isolates the minimum to a fractional precision of
'  about TOL using Brent's method.  The abscissa of the minimum is
'  returned as XMIN, & the minimum function value is returned as BRENT.
' Numerical Recipes, p. 285-286.
' (check tranlation of the FORTRAN statement SIGN)
Dim Iter%
Const ITMAX = 100, CGOLD = 0.381966, ZEPS = 0.00000000000001
Dim X#, v#, W#, d#, e#, A#
Dim b#, u#, fx#, FU#, Fv#, FW#
Dim Xm#, TOL1#, TOL2#, r#, q#, P#
Dim Etemp#, T#, TimeIn#, MaxTime#
A = Min(aX, Cx): b = Max(aX, Cx)
v = BX: W = v: X = v: e = 0
fx = ConcordSums((X), Bad)
If Bad Then Exit Sub
Fv = fx: FW = fx
TimeIn = Timer(): MaxTime = 12
For Iter = 1 To ITMAX
  If Iter Mod 20 = 0 Then TooLongCheck TimeIn, MaxTime
  Xm = (A + b) / 2
  TOL1 = Tol * Abs(X) + ZEPS
  TOL2 = 2 * TOL1
  If Abs(X - Xm) <= (TOL2 - (b - A) / 2) Then
    Xmin = X: Fmin = fx
    Exit Sub
  End If
  If Abs(e) > TOL1 Then
    r = (X - W) * (fx - Fv)
    q = (X - v) * (fx - FW)
    P = (X - v) * q - (X - W) * r
    q = 2 * (q - r)
    If q > 0 Then P = -P
    q = Abs(q)
    Etemp = e: e = d
    If Abs(P) >= Abs(q * Etemp / 2) Or P <= (q * (A - X)) Or P <= (q * (A - X)) _
       Or P >= (q * (b - X)) Then GoTo Br1
    d = P / q:        u = X + d
    If (u - A) < TOL2 Or (b - u) < TOL2 Then d = SIGN(TOL1, Xm - X)
    GoTo Br2
  End If
Br1:
  If X >= Xm Then e = A - X Else e = b - X
  d = CGOLD * e
Br2:
  If Abs(d) >= TOL1 Then
    u = X + d
  Else
    u = X + SIGN(TOL1, d)
  End If
  FU = ConcordSums((u), Bad)
  If Bad Then Exit Sub
  If FU <= fx Then
    If u >= X Then A = X Else b = X
    v = W: Fv = FW
    W = X: FW = fx
    X = u: fx = FU
  Else
    If u < X Then A = u Else b = u
    If FU <= FW Or W = X Then
      v = W: Fv = FW
      W = u: FW = FU
    ElseIf (FU <= Fv Or v = X Or v = W) Then
      v = u: Fv = FU
    End If
  End If
Next Iter
Bad = True: T = 0
End Sub

Sub VarTcalc(ByVal T#, SigmaT#, Bad As Boolean)
' Calculate the variance in age for a single assumed-concordant data point
'  on the Conv. U/Pb concordia diagram (with or without taking into
'  account the uranium decay-constant errors).
' See GCA v62, p665-676, 1998 for explanation.
Dim e5#, e8#, Q5#, Q8#, Xvar#, Yvar#
Dim Cov#, Om11#, Om12#, Om22#, Fisher#
SigmaT = 0
e5 = Lambda235 * T
If Abs(e5) > MAXEXP Then Exit Sub
e5 = Exp(e5):        e8 = Exp(Lambda238 * T)
Q5 = Lambda235 * e5: Q8 = Lambda238 * e8
Xvar = vcXY(1, 1): Yvar = vcXY(2, 2)
If Not Cmisc.NoLerr Then
  Xvar = Xvar + SQ(T * e5 * Lambda235err)
  Yvar = Yvar + SQ(T * e8 * Lambda238err)
End If
Cov = vcXY(1, 2)
Inv2x2 Xvar, Yvar, Cov, Om11, Om22, Om12, Bad
If Bad Then SigmaT = 0: Exit Sub
Fisher = Q5 * Q5 * Om11 + Q8 * Q8 * Om22 + 2 * Q5 * Q8 * Om12
' Fisher is the expected second derivative with respect to T of the
'  sums-of-squares of the weighted residuals.
If Fisher > 0 Then SigmaT = Sqr(1 / Fisher)
End Sub

Sub ShowXYwtdMean(ByVal X#, ByVal ErrX#, ByVal y#, ByVal ErrY#, _
  ByVal Rho#, ByVal Sums#, ByRef MSWD#, ByRef Prob#, _
  ByVal Npts&, ByRef ErrMult#)
' Show X-Y weighted-mean results
Dim df&, SfP%, SfX%, SfY%, i%
Dim r, Mult95, ErrX95, ErrY95, xym As Object, pl3 As Boolean, pl5 As Boolean
Dim ss$, Op As Object, tB As Object, Grp As Object, Chk As Object, s$, L As Object
Dim Proba$, el$, j%, vv$, ee$, tBo As Boolean
AssignD "xyWtdAv", xym, , Chk, Op, L, Grp, tB
Set Chk = Chk("cShowRes")
For i = 1 To tB.Count: tB(i).Visible = True: Next
tB(1).Text = "1sigma a priori": tB(2).Text = "2sigma a priori"
tB(3).Text = tB(1).Text:    tB(5).Text = tB(2).Text
tB(4).Text = "tsigmaSqrtMSWD"
tB(6).Text = tB(4).Text
If Not Mac Then
  For i = 1 To tB.Count
    With tB(i).Font
      .Name = IIf(Mac, "Geneva", "Arial")
      .Size = 11 + Windows
    End With
    tB(i).Visible = True
    ConvertSymbols tB(i)
  Next i
End If
For i = 1 To Grp.Count: Grp(i).Visible = True: Next i
df = 2 * Npts - 2
If df <= 0 Then
  MSWD = 0: Prob = 1
Else
  MSWD = Sums / df
  Prob = ChiSquare(MSWD, df)
End If
If Prob >= 0.001 Then
  For i = 1 To L.Count: L(i).Enabled = True: Next i
  NumAndErr X, ErrX, 2, vv$, ee$, , True
  L(1).Text = "X = " & vv$:     L(3).Text = ee$
  NumAndErr y, ErrY, 2, vv$, ee$, , True
  L(2).Text = "Y = " & vv$:     L(12).Text = ee$
  L(5).Text = "X-Y error correlation = " & RhoRnd(Rho)
  L(6).Text = "MSWD = " & Mrnd(MSWD)
  Proba$ = ProbRnd(Prob)
  L(7).Text = "Probability of X-Y equivalence = " & Proba$
  L(8).Text = ErFo(X, 2 * ErrX, 2, True)
  L(14).Text = ErFo(y, 2 * ErrY, 2, True)
  L(4).Text = "(" & ErFo(X, ErrX, 2, , True) & ")"
  L(13).Text = "(" & ErFo(y, ErrY, 2, , True) & ")"
  L(9).Text = "(" & ErFo(X, 2 * ErrX, 2, , True) & ")"
  L(15).Text = "(" & ErFo(y, 2 * ErrY, 2, , True) & ")"
  Mult95 = StudentsT((df)) * Sqr(MSWD)
  L(10).Text = ErFo(X, Mult95 * ErrX, 2, True)
  L(11).Text = "(" & ErFo(X, Mult95 * ErrX, 2, , True) & ")"
  L(16).Text = ErFo(y, Mult95 * ErrY, 2, True)
  L(17).Text = "(" & ErFo(y, Mult95 * ErrY, 2, , True) & ")"
  pl3 = (Prob < 0.3):   pl5 = (Prob >= 0.05)
  For i = 1 To 17
    With L(i)
      If i > 2 And (i < 5 Or i > 7) Then
        .Enabled = IIf(i = 10 Or i = 11 Or i = 16 Or i = 17, pl3, pl5)
      Else
        .Enabled = True
      End If
    End With
  Next i
  If Not Mac Then
    For i = 1 To 6
      tBo = IIf(i = 4 Or i = 6, pl3, pl5)
      tB(i).Font.Color = IIf(tBo, vbBlack, Menus("cGray50"))
    Next i
  End If
  Grp(6).Visible = True
  Grp(6).Enabled = True '(DoPlot Or IsOn(Chk))
  For i = 1 To 3
    With Op(i)
      .Visible = True
      .Enabled = True 'DoPlot Or IsOn(Chk))
      .Value = xlOff
    End With
  Next i
  Op(4).Enabled = DoPlot
  Op(4).Visible = True
  'If (DoPlot Or IsOn(Chk)) Then
    If Prob > MinProb Then
      Op(SigLev).Value = xlOn
      If Prob > 0.3 Then Op(3).Enabled = False
    Else
      Op(3).Value = xlOn
      If Prob < 0.05 Then Op(1).Enabled = False: Op(2).Enabled = False
    End If
  'End If
  ShowBox xym, True
  'If (DoPlot Or IsOn(Chk)) Then
    If IsOn(Op(1)) Then
      ErrMult = 1
    ElseIf IsOn(Op(2)) Then
      ErrMult = 2
    ElseIf IsOn(Op(3)) Then
      ErrMult = Mult95
    Else
      ErrMult = 0
    End If
  'End If
Else
  ss$ = "Data points are not equivalent"
  If ConcAge Then ss$ = ss$ & " -- cannot calculate a Concordia Age."
  ss$ = ss$ & String(2, vbLf) & "(MSWD=" & Mrnd(MSWD)
  ss$ = ss$ & ",  Probability-of-fit = " & ProbRnd(Prob) & ")"
  If MsgBox(ss$, vbOKCancel, Iso) <> vbOK Or Not DoPlot Then ExitIsoplot
End If
If IsOn(Chk) Then
  If IsOn(Op(1)) Then
    i = 3:  j = 12:  el$ = "1-sigma"
  ElseIf IsOn(Op(2)) Then
    i = 8:  j = 14:  el$ = "2-sigma"
  Else
    i = 10: j = 16:  el$ = "95%-conf."
  End If
  s$ = "X-Y Weighted Mean:" & vbLf + L(1).Text & L(i).Text & "  " & el$ & _
        vbLf & L(2).Text & L(j).Text & vbLf & L(5).Text & vbLf & L(6).Text & _
        ",  Probability =" & Proba$
  AddResBox s$, -1, 1, LightGreen
End If
End Sub

Private Sub ShowXYerrLevels()
Dim xyo As Object, xyc As Object, xyg As Object
Dim i%, b As Boolean
AssignD "xyWtdAv", , , xyc, xyo, , xyg
If Not DoPlot Then
  b = IsOn(xyc(1))
  xyg(6).Enabled = b
  For i = 1 To 3
    xyo(i).Enabled = b
    If Not b Then xyo(i) = xlOff
  Next i
End If
End Sub
Function ConcordSums(ByVal T#, Bad As Boolean)
 ' Calculate the sums of the squares of the weighted residuals for a single
 '  Conv.-Conc. X-Y data point, where the true value of
 '  each of the data pts is assumed to be on the same point on the
 '  concordia curve, & where the decay constants that describe the
 '  concordia curve have known uncertainties.
 ' See GCA 62, p. 665-676, 1998 for explanation.
Dim e5#, e8#, Ee5#, Ee8#, Rx#, Ry#
Dim Xbvar#, Ybvar#, Om11#, Om22#, Om12#
Bad = False
e5 = Lambda235 * T
If Abs(e5) > MAXEXP Then ConcordSums = 0: Exit Function
e5 = Exp(e5):       e8 = Exp(Lambda238 * T)
Ee5 = e5 - 1:       Ee8 = e8 - 1
Rx = Cmisc.Xconc - Ee5:   Ry = Cmisc.Yconc - Ee8
Xbvar = vcXY(1, 1): Ybvar = vcXY(2, 2)
If Not Cmisc.NoLerr Then
 Xbvar = Xbvar + SQ(T * e5 * Lambda235err)
 Ybvar = Ybvar + SQ(T * e8 * Lambda238err)
End If
Inv2x2 Xbvar, Ybvar, vcXY(1, 2), Om11, Om22, Om12, Bad
If Bad Then
  ConcordSums = 0
Else
  ConcordSums = Rx * Rx * Om11 + Ry * Ry * Om22 + 2 * Rx * Ry * Om12
End If
End Function
