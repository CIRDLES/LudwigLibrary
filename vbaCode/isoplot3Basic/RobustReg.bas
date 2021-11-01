Attribute VB_Name = "RobustReg"
Option Private Module
Option Explicit: Option Base 1
Private Const nMadConst = 1.4826

Sub RobustReg1(xy#(), Slope#, Optional Intercept, _
  Optional DoCheck = True, Optional SlopeOnly = False)
' Similar to RobustReg2, but is doubly robust -- for each point, calculates median of all pairwise
'  slopes; then calculates the median of these medians (Siegel).
' Errors must be estimated using a Bootstrap, so slow!
Dim i&, j&, k&, X#(), y#(), MedSlope#()
Dim N&, Slp#(), MedInt#(), DoInt As Boolean
Const Eps = 0.00000001
ViM DoCheck, True
ViM SlopeOnly, False
DoInt = Not SlopeOnly
MakeXY xy(), X(), y(), N, (DoCheck)
If N < 3 Then MsgBox "Invalid input-range", , Iso: ExitIsoplot
If N > 4095 Then MsgBox "Max. #points is 4095", , Iso: ExitIsoplot
ReDim MedSlope(N), Slp(N - 1)
If DoInt Then ReDim MedInt(N)
For i = 1 To N
  k = 0
  For j = 1 To N
    If i <> j Then
      k = 1 + k
      If X(i) <> X(j) Then
        Slp(k) = (y(j) - y(i)) / (X(j) - X(i))
      Else
        Slp(k) = 0
      End If
      Slp(k) = Slp(k) + (0.5 - Rnd) * Eps
    End If
  Next j
  MedSlope(i) = iMedian(Slp())
  If DoInt Then MedInt(i) = y(i) - MedSlope(i) * X(i) + (0.5 - Rnd) * Eps
Next i
Slope = iMedian(MedSlope())
If DoInt Then Intercept = iMedian(MedInt())
End Sub

Sub GetRobSlope(X#(), y#(), ByVal N&, ByRef M&, _
   ByVal RejOutliers As Boolean, ByVal wXinter As Boolean, Slope#, _
  Yint#, Xint#, Slp#(), Yinter#(), Xinter#())
Attribute GetRobSlope.VB_ProcData.VB_Invoke_Func = " \n14"
' Note: If RejectOutliers=TRUE than x() and y() are returned minus any outliers!
Dim i&, j&, k&, Vs#, u#, s#
Dim Resid#(), Resid2#(), vx#, Vy#
Dim xx#(), yy#()
Const Eps = 0.00000000001
M = N * (N - 1) / 2
If M > EndRow Then Exit Sub
ReDim Slp(M), Yinter(M)
If wXinter Then ReDim Xinter(M)
For i = 1 To N - 1
  For j = i + 1 To N
    k = 1 + k
    If X(i) <> X(j) Then
      Vs = (y(j) - y(i)) / (X(j) - X(i))
    Else
      Vs = 0
    End If
    Vs = Vs + (0.5 - Rnd) * Eps
    Slp(k) = Vs: Yinter(k) = Vy
    Vy = y(i) - Vs * X(i) + (0.5 - Rnd) * Eps
    Yinter(k) = Vy
    If wXinter Then
      vx = -Vy / Vs
      Xinter(k) = vx
    End If
Next j, i
Slope = iMedian(Slp()): Yint = iMedian(Yinter())
If wXinter Then Xint = iMedian(Xinter())
If RejOutliers Then ' Reject outliers based on MAD using approach outlined in Powell et al.
  ReDim Resid(N), Resid2(N), xx(N), yy(N)
  For i = 1 To N
    Resid(i) = y(i) - Slope * X(i) - Yint
    Resid2(i) = Resid(i) * Resid(i)
  Next i
  Erase Slp, Yinter
  s = nMadConst * (1 + 5 / (N - 2)) * Sqr(iMedian(Resid2()))
  k = 0
  For i = 1 To N
    u = Abs(Resid(i) / s)
    If u < 2.5 Then
      k = 1 + k
      xx(k) = X(i): yy(k) = y(i)
    End If
  Next i
  If k < N Then
    N = k
    ReDim X(N), y(N)
    For i = 1 To N
      X(i) = xx(i): y(i) = yy(i)
    Next i
  End If
End If
End Sub

Sub RobustReg3(ByVal N&, DP() As DataPoints, yf As Yorkfit, _
  Optional Uage, Optional Ugamma0, Optional UGrho)
Attribute RobustReg3.VB_ProcData.VB_Invoke_Func = " \n14"
' Robust 3-D linear regression using median of all pairwise slopes/intercepts (Theil),
' modified after 2-D regression from Hoaglin, Mosteller & Tukey, Understanding
' Robust & Exploratory Data Analysis, John Wiley & Sons, 1983, p. 160, with errors
' from code in Rock & Duffy, 1986 (Comp. Geosci. 12, 807-818), derived from Vugrinovich
' (1981), J. Math. Geol. 13, 443-454).
Dim i&, j&, k&, US As Boolean, s1$, s2$, s3$
Dim UpprAge#, LwrAge#, UpprGamma0#, LwrGamma0#
Dim Age#(), Gamma0#(), X#(), y#(), z#()
Dim M&, SlpXY#(), SlpXZ#(), Th230U238#, Gamma#
Dim Yinter#(), Zinter#(), c$, ThU#, Gfree#
Dim LwrInd&, UpprInd&, VsXY#, VsXZ#
Dim ViXY#, ViXZ#, r As Range, Denom#
Const Big = 1E+32, Eps = 0.00000000001
' Max m = 65536!
If N < 3 Then MsgBox "Need 3 or more x-y pairs", , Iso: ExitIsoplot
US = NIM(Uage)
ReDim X(N), y(N), z(N)
For i = 1 To N
  With DP(i): X(i) = .X: y(i) = .y: z(i) = .z: End With
Next i
M = N * (N - 1#) / 2#
If M > EndRow Then
  MsgBox "Can't do robust xyz regression for N>360"
  KwikEnd
End If
ReDim SlpXY(M), SlpXZ(M), Yinter(M), Zinter(M)
If US Then ReDim Age#(M), Gamma0(M)
k = 0
For i = 1 To N - 1
  For j = i + 1 To N
    k = 1 + k
    Denom = X(j) - X(i)
    If Denom <> 0 Then
      VsXY = (y(j) - y(i)) / Denom
      VsXZ = (z(j) - z(i)) / Denom
    Else
      VsXY = 0: VsXZ = 0
    End If
    VsXY = VsXY + (0.5 - Rnd) * Eps
    VsXZ = VsXZ + (0.5 - Rnd) * Eps
    SlpXY(k) = VsXY:    SlpXZ(k) = VsXZ
    ViXY = y(i) - VsXY * X(i) + (0.5 - Rnd) * Eps
    ViXZ = z(i) - VsXZ * X(i) + (0.5 - Rnd) * Eps
    Yinter(k) = ViXY:    Zinter(k) = ViXZ
    If US Then
      Select Case UsType
        Case 1: ThU = ViXY: Gfree = ViXZ
        Case 2: ThU = ViXZ: Gfree = ViXY
        Case 3
          ThU = -ViXZ / VsXZ
          Gfree = ViXY + VsXY * ThU
      End Select
      ThUage ThU, Gfree, Age(k)
      If Age(k) = 0 Then
        If ThU > Gfree Then
          Age(k) = Big: Gamma0(k) = -Big
        Else
          Age(k) = -Big: Gamma0(k) = Big
        End If
      Else
        Age(k) = Age(k) / Thou
        Gamma0(k) = InitU234U238(Age(k), Gfree)
      End If
    End If
  Next j
Next i
Erase X, y, z
Conf95 N, (M), LwrInd, UpprInd
'ReDim IntSl(4) ', ErrRho(4, 4)
QuickSort SlpXY()
QuickSort SlpXZ()
QuickSort Yinter()
QuickSort Zinter()
With yf
  .Slope = iMedian(SlpXY(), -1):      .SlopeXZ = iMedian(SlpXZ(), -1)
  .LwrSlope = SlpXY(LwrInd):          .UpprSlope = SlpXY(UpprInd)
  .LwrSlopeXZ = SlpXZ(LwrInd):        .UpprSlopeXZ = SlpXZ(UpprInd)
  .Intercept = iMedian(Yinter(), -1): .Zinter = iMedian(Zinter(), -1)
  .LwrInter = Yinter(LwrInd):         .UpprInter = Yinter(UpprInd)
  .LwrZinter = Zinter(LwrInd):        .UpprZinter = Zinter(UpprInd)
End With
If US Then
  QuickSort Age()
  QuickSort Gamma0()
  LwrAge = Age(LwrInd):       UpprAge = Age(UpprInd)
  LwrGamma0 = Gamma0(LwrInd): UpprGamma0 = Gamma0(UpprInd)
End If
With App
  If US Then
    Uage = iMedian(Age()): Ugamma0 = iMedian(Gamma0())
    UGrho = .Correl(Age, Gamma0)
    s1$ = "Robust 3-D Linear Isochron Solution:" & vbLf
    s2$ = "Age = " & Sd(Uage, 4, , True) & "  " & _
      Sd(UpprAge - Uage, 2, True, True) & _
      "/" & Sd(LwrAge - Uage, 2, True, True) & "  ka  " & vbLf & _
      "Initial 234U/238U = " & Sd(Ugamma0, 4, , True) & "  " & _
      Sd(UpprGamma0 - Ugamma0, 2, True, True) & "/" & _
      Sd(LwrGamma0 - Ugamma0, 2, True, True) & vbLf & _
      "Err-correl for Age-Initial234/238  = " & RhoRnd(UGrho) & vbLf
    s3$ = "(approx. 95% conf.-limit errors)"
    MsgBox s1$ & vbLf & s2$ & vbLf & s3$, , Iso
    AddResBox s1$ & s2$ & s3$
  End If
End With
End Sub

Sub BootRob(xy#(), ByVal Ntrials&, SlopeLwr#, SlopeUppr#, _
  InterLwr#, InterUppr#)
' Get errs for Siegel-algorithm robust regr'n
Dim i&, j&, k&, N&, M&, MxCt&
Dim Slope#(), Inter#(), NT&, SB$, ct&, txy#()
N = UBound(xy, 1):  MxCt = Ntrials + 2000
ReDim txy(N, 2), Slope(Ntrials), Inter(Ntrials)
SB$ = "bootstrapping the regression errors"
Randomize Timer
NT = Ntrials + 50
i = 1
Do
  For j = 1 To N
    k = 1 + (N - 1) * Rnd
    txy(j, 1) = xy(k, 1)
    txy(j, 2) = xy(k, 2)
  Next j
  If i Mod 10 = 0 Then StatBar SB$ & Str(NT - i)
  RobustReg1 txy(), Slope(i), Inter(i), False, False
  ct = 1 + ct
  If ct > MxCt Then MsgBox "Unable to calculate robust regression errors", , Iso: Exit Sub
  i = i + 1
Loop Until i > Ntrials
StatBar
QuickSort Slope()
QuickSort Inter()
SlopeUppr = Slope(0.975 * Ntrials)
SlopeLwr = Slope(0.025 * Ntrials)
InterUppr = Inter(0.975 * Ntrials)
InterLwr = Inter(0.025 * Ntrials)
StatBar
End Sub

Sub ShowRobust()
Attribute ShowRobust.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, L As Object, N&, e#, xs$, aer$, uer$, ler$, Uerr#, Lerr#
Set L = DlgSht("RobustRes").Labels
With L("Lntrials")
  N = yf.Ntrials: .Visible = (N > 0)
  If .Visible Then .Text = IIf(N > 0, "Bootstrap errors from" & Str(N) & " trials", _
      "(at least 8 points required for error calc.)")
End With
With yf
  s$ = "Slope = "
  If .UpprSlope <> .LwrSlope Then
    Uerr = .UpprSlope - .Slope: Lerr = .LwrSlope - .Slope
    e = (Uerr - Lerr) / 2
    NumAndErr .Slope, e, 2, xs$, aer$
    NumAndErr .Slope, Uerr, 2, "", uer$
    NumAndErr .Slope, Lerr, 2, "", ler$
    s$ = s$ & xs$ & "  +" & uer$ & "/-" & ler$
  Else
    s$ = s$ & tSt(.Slope)
  End If
  L("Lslope").Text = s$
  s$ = "Intercept = "
  If .UpprInter <> .LwrInter Then
    Uerr = .UpprInter - .Intercept: Lerr = .LwrInter - .Intercept
    e = (Uerr - Lerr) / 2
    NumAndErr .Intercept, e, 2, xs$, aer$
    NumAndErr .Intercept, Uerr, 2, "", uer$
    NumAndErr .Intercept, Lerr, 2, "", ler$
    s$ = s$ & xs$ & "  +" & uer$ & "/-" & ler$
  Else
    s$ = s$ & tSt(.Intercept)
  End If
  L("Linter").Text = s$
End With
Do
  ShowDialog "RobustRes"
If Not AskInfo Then Exit Do
  Caveat_RobustRes
Loop
If IsOn(DlgSht("RobustRes").CheckBoxes("cShowRes")) Then
  s$ = "Robust Regression   (~95% conf.)" & vbLf & L("Lslope").Text & vbLf & L("Linter").Text
  AddResBox s$
End If
End Sub

Sub Conf95(ByVal Npts&, ByVal Nmedians&, LowInd&, UpprInd&)
Attribute Conf95.VB_ProcData.VB_Invoke_Func = " \n14"
' Finds sorting-indexes to get 95%-conf. limits for repeated pairwise slope/inter medians using
' algorithm coded in Rock & Duffy, 1986 (Comp. Geosci. 12, 807-818), derived from Vugorinovich
' (1981, J. StrictMath. Geol. 13, 443-454).
Dim i&, c$, X#, Star95&
Select Case Npts
  Case Is < 5
    LowInd = 1: UpprInd = Npts
  Case Is < 14
    c$ = "081012141719222528"
    Star95 = Val(Mid$(c$, 2 * Npts - 9, 2))
  Case Else
    X = Sqr(Npts * (Npts - 1#) * (2# * Npts + 5#) / 18#)
    Star95 = 1.96 * X
End Select
If Npts > 4 Then
  LowInd = (Nmedians - Star95) / 2
  UpprInd = (Nmedians + Star95) / 2
End If
End Sub

Sub Wtd(X#(), Sigma#(), ByVal Np&, WtdMean#, _
  SigmaMean#, MSWD#, Prob#)
' Simple weighted averages with errors and prob-of-fit.
Dim i&, j&, W#, SumW#, SumX#, s#
For i = 1 To Np
  W = 1 / SQ(Sigma(i))
  SumW = SumW + W
  SumX = SumX + X(i) * W
Next i
WtdMean = SumX / SumW
SigmaMean = 1 / Sqr(SumW)
For i = 1 To Np
  s = s + SQ((X(i) - WtdMean) / Sigma(i))
Next i
MSWD = s / (Np - 1)
Prob = ChiSquare(MSWD, Np - 1)
End Sub

Public Function nMAD(ByVal N, ByVal Yresid)
Dim i&, yr2#()
ReDim yr2(N)
For i = 1 To N: yr2(i) = SQ(Yresid(i)): Next i
nMAD = nMadConst * (1 + 5 / (N - 2)) * Sqr(iMedian(yr2()))
End Function

Private Sub ExtractCoherentGroup(X#(), Sigma#(), LargeErr() As Boolean, _
  ByVal N&, ByVal Ncontig&, gFirst&, gLast&, gN&)
' Extracts largest & best-defined coherent age-group from a list of age-sorted analyses.
' Only analyses whose ages are yield an acceptable MSWD and Prob-Of-Fit can be selected.
Dim i&, j&, k&, Ng&, MaxN&, Ngroups%
Dim First&, Last&, MaxProbInd&, MaxNind&, nU&
Dim Nn&, GrpFirst&(), GrpLast&(), GrpN&(), tB As Boolean
Dim gProb#, LastgProb#, AvErr#, nMadd#, W#, Wav#
Dim GrpProb#(), q#, Tx#(), tSig#(), Sum1#, Sum2#
Dim gx#(), gSig#(), Yresid#()
ReDim gx(N), gSig(N), Yresid(N), Tx(N), tSig(N)
ReDim GrpFirst(N), GrpLast(N), GrpProb(N), GrpN(N)
Const MinP = 0.025, MaxMSWD = 2.25 ' ie 1.5 x analytical errors
AvErr = iAverage(Sigma())
Nn = N: j = 0: k = 0
ReDim LargeErr(N)
nMadd = nMAD(N, Sigma())
If nMadd < AvErr Then nMadd = AvErr
nMadd = 1.5 * nMadd
For i = 1 To N ' Clean of large-error outliers
  Tx(i) = X(i): tSig(i) = Sigma(i)
  LargeErr(i) = (Sigma(i) > nMadd)
Next i
For i = 1 To Nn - 1
  If Not LargeErr(i) Then
    First = i: Last = First: Ng = 1: gx(1) = Tx(i): gSig(1) = tSig(i)
    Do
      Last = 1 + Last
      If Not LargeErr(Last) Then
        Ng = 1 + Ng
        gx(Ng) = Tx(Last): gSig(Ng) = tSig(Last)
        Sum1 = 0: Sum2 = 0: q = 0
        For j = 1 To Ng
          W = 1 / SQ(gSig(j))
          Sum1 = Sum1 + gx(j) * W
          Sum2 = Sum2 + W
        Next j
        Wav = Sum1 / Sum2
        For j = 1 To Ng
          q = q + SQ((gx(j) - Wav) / gSig(j))
        Next j
        nU = Ng - 1
        gProb = ChiSquare(q / nU, nU)
        If gProb >= MinP Or q <= MaxMSWD Then LastgProb = gProb
      End If
    Loop Until (Not LargeErr(Last) And (gProb < MinP And q > MaxMSWD)) Or Last = Nn
    If gProb < MinP Then
      Do
        Last = Last - 1
      Loop Until Not LargeErr(Last)
      Ng = Ng - 1
    End If
    If Ngroups = 0 Then tB = True Else tB = (Last <> GrpLast(Ngroups))
    If Ng > 1 And tB Then
      Ngroups = 1 + Ngroups
      GrpFirst(Ngroups) = First: GrpLast(Ngroups) = Last
      GrpN(Ngroups) = Ng: GrpProb(Ngroups) = LastgProb
      If Last = Nn Then Exit For
    End If
  End If
Next i
If Ngroups = 0 Then Exit Sub
gFirst = 0: gLast = 0:  MaxN = 0
For i = 1 To Ngroups
  Ng = GrpN(i)
  If Ng > MaxN Then MaxN = Ng: MaxNind = i
Next i
If MaxN < Ncontig Then Exit Sub
For i = 1 To Ngroups    ' If more than 1 group has MaxN,
  If i <> MaxNind Then  '   select highest prob-of-fit.
    If GrpN(i) = MaxN Then
      If GrpProb(i) > GrpProb(MaxNind) Then MaxNind = i
    End If
  End If
Next i
gFirst = GrpFirst(MaxNind): gLast = GrpLast(MaxNind): gN = GrpN(MaxNind)
End Sub

Sub ZirconAgeExtractor(Optional FromMenu As Boolean = False)
Attribute ZirconAgeExtractor.VB_ProcData.VB_Invoke_Func = " \n14"
Dim r As Range, N&, k&, i&, dFirst&, dLast&, AlphaBox As Boolean
Dim Ng&, j&, Age#, PlusErr#, MinusErr#, Sp As Object, Ncontig&
Dim ConfLevel#, Rd As Object, o As Object, Chrt As Object, tbx As Object, sh1 As Object, sh2 As Object
Dim Bad As Boolean, s$, s1$, s2$, s3$, e As Object, Na%, W As wWtdAver, Rej(), ChtSht As Object
Dim T#(), SigmaTi#(), SigmaT#(), tmpInp#(), L As Object
Dim NM#, LargeErr() As Boolean, IndX&(), gT#(), gSig#()
Const MinVals = 8, MinInGrp = 5
ViM FromMenu, False

If FromMenu Then
  N = UBound(InpDat, 1)
  ReDim tmpInp(N, 2), T#(N)

  For i = 1 To N
    For j = 1 To 2
      tmpInp(i, j) = InpDat(i, j)
  Next j, i

Else
  AgeExtract = True: Isotype = 24
  Set DatSht = Ash
  DatSheet$ = DatSht.Name
  DoPlot = True: ColorPlot = True
  RangeCheck 0, 0, Na
  NoUp
  Set r = Selection: Irange$ = r.Address
  ParseAgeRange r, N, Na, T(), SigmaTi(), TopRow, RightCol, Bad
  If Bad Or N < MinVals Then
    MsgBox "Invalid input range" & vbLf & "(need 2 columns and" & Str(MinVals) & _
      " or more rows of numeric data)", , Iso
    ExitIsoplot
  End If
  NoUp False
  AbsErrs = True '(o("oAbs") = xlOn)
  SigLev = 1 '- (o("o2sigma") = xlOn)
  Ncontig = MinInGrp ' MinMax(MinInGrp, n, Val(e.Text))
End If
NoUp
ReDim IndX(N), SigmaT#(N)

If Not FromMenu Then
  ReDim InpDat(N, 2), tmpInp(N, 2)

  For i = 1 To N
    tmpInp(i, 1) = T(i)
    tmpInp(i, 2) = SigmaTi(i) / SigLev
    If Not AbsErrs Then tmpInp(i, 2) = tmpInp(i, 2) / 100 * T(i)
  Next i

End If

SortCol tmpInp(), T(), N, 1, IndX()
j = 0

For i = N To 1 Step -1
  j = 1 + j
  SigmaT(j) = tmpInp(IndX(i), 2)
  InpDat(j, 1) = T(i)
  InpDat(j, 2) = SigmaT(j)
Next i

For j = 1 To N
  T(j) = InpDat(j, 1)
  SigmaT(j) = InpDat(j, 2)
Next j

' t(), SIgmat(), and InpDat() are now all sorted in descending age-order
ExtractCoherentGroup T(), SigmaT(), LargeErr(), N, Ncontig, dFirst, dLast, Ng

If Ng < 2 Then
  MsgBox "Couldn't find a coherent group of " & Ncontig & " or more ages", , Iso
  ExitIsoplot
End If

ReDim gT(Ng), gSig(Ng)
j = 0

For i = dFirst To dLast
  If Not LargeErr(i) Then
    j = 1 + j
    gT(j) = T(i)
    gSig(j) = SigmaT(i)
  End If
Next i

Age = iMedian(gT())
PlusErr = MedianUpperLim(gT) - Age
MinusErr = Age - MedianLowerLim(gT)
ConfLevel = MedianConfLevel(Ng)
s1$ = "TuffZirc Age  = " & Format(Age, "0.00") & "   +" & Format(PlusErr, "0.00") _
   & "   -" & Format(MinusErr, "0.00") & "  Ma"
s2$ = "(" & tSt(ConfLevel) & "% conf, from coherent group of" & Str(Ng) & ")"
s$ = s1$ & vbLf & vbLf & s2$
s3$ = vbLf & s '1$ & viv & s2$

Do
  Load TuffZirc 'LoadUserForm TuffZirc
  With TuffZirc
    .cAddChart = DoPlot
    With .tbResults '.lResults
      .Text = s3$: .Enabled = True: .AutoSize = True
      .AutoSize = False: .Height = .Height + 10: .Width = .Width + 10
    End With
    .Show
    AlphaBox = .cShowWithData
  End With
  If Canceled Then ExitIsoplot
If Not AskInfo Then Exit Do
  Caveat_TuffZirc
Loop

If DoPlot Then
  With W
    .IntMean = Age: .ExtMean = Age + PlusErr: .BiwtMean = Age - MinusErr
    .ChosenMean = dFirst: .ChosenErr = dLast
  End With
  SymbCol = 3: SymbRow = 1: DoShape = True
  Sheets.Add:  PlotDat$ = "PlotDat"
  AssignIsoVars
  MakeSheet PlotDat$, ChrtDat
  WtdAverPlot N, W, Rej(), 0, s$, LargeErr()
  Set ChtSht = ActiveSheet
  AddErrSymbSizeNote True, 18
  Last(Ach.Shapes).ZOrder msoSendToBack
  PutPlotInfo

  If True Then
    ActiveChart.PlotArea.Select
    CopyPicture
    Set Chrt = Last(ActiveSheet.Shapes)
    DelSheet ChtSht
    DelSheet ChrtDat
  End If

  Set sh1 = Last(Ash.Shapes)

ElseIf AlphaBox Then
  AddResBox s$, , , RGB(200, 255, 255), Italics:=7

  If DoPlot Then
    Set sh2 = Last(Ash.Shapes)
    sh2.Top = Bottom(sh1) + 2
  End If

End If

'AddCopyButton "TuffZirc_Help", "Caveat_TuffZirc", True
ExitIsoplot
End Sub

Sub TuffZircSpinClick()
Attribute TuffZircSpinClick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, e As Object, s As Object
With DlgSht("TuffZirc")
  Set e = .EditBoxes("eSpin"): Set s = .Spinners("sSpin")
End With
e.Text = Trim(Str(s.Value))
End Sub

Sub ParseAgeRange(r As Range, ByRef N&, ByVal Na%, _
  T#(), SigmaTi#(), TopRow&, RightCol%, Bad As Boolean)
Attribute ParseAgeRange.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i&, j&, k&, nCC%, nCCz%
Dim c As Range, v#, rc%, nc%, nR&, nRa&()
Dim y#()
ReDim nRa(Na)
For i = 1 To Na: nRa(i) = r.Areas(i).Rows.Count: Next i
nR = Sum(nRa)
ReDim T(nR), SigmaTi(nR)
N = 0
For i = 1 To Na
  With r.Areas(i)
    TopRow = IIf(.Row > TopRow, .Row, TopRow)
    nCC = .Columns.Count: rc = .Column + nCC - 1
    RightCol = IIf(rc > RightCol, rc, RightCol)
    If i = 1 Then
      nCCz = nCC
    ElseIf nCC <> r.Areas(1).Columns.Count Then
      Bad = True: Exit Sub
    End If
    If nCC > 2 Then Bad = True: Exit Sub
    For j = 1 To nRa(i)
      For k = 1 To nCC
        Set c = .Cells(j, k)
        If Not IsEmpty(c) Then
          If IsNumeric(c) Then
            v = c.Value
            If Not c.Font.Strikethrough Then
              If c.Column = r.Areas(1).Column Then
                N = N + 1: T(N) = v
              Else
                nc = 1 + nc: SigmaTi(nc) = v
              End If
            End If
          End If
        End If
      Next k
    Next j
  End With
Next i
ReDim Preserve T(N), SigmaTi(N)
End Sub