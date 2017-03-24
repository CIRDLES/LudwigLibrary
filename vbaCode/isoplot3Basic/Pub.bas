Attribute VB_Name = "Pub"
' ISOPLOT module Pub
Option Explicit: Option Base 1

Public Function AgePb76(ByVal Pb76rad, Optional ByVal t2)
Attribute AgePb76.VB_Description = "Age (Ma) from radiogenic 207Pb/206Pb: t2=0 if unspecified."
Attribute AgePb76.VB_ProcData.VB_Invoke_Func = " \n14"
Dim test#, tmp1, tmp2
tmp1 = Menus("NumErr")
If Not IsNumeric(Pb76rad) Then GoTo 1
If Pb76rad < 0.03 Or Pb76rad > 10 Then GoTo 1
IMN t2, 0
If Not IsNumeric(t2) Then GoTo 1
If t2 < 0 Or t2 > 10000000000# Then GoTo 1
test = PbPbAge(Pb76rad, t2)
If Not IsNumeric(test) Then GoTo 1
If test = BadT Then GoTo 1
tmp1 = test
1:
AgePb76 = tmp1
End Function

Public Function AgeErPb76(ByVal Pb76rad, ByVal Pb76er, Optional t2 = 0, Optional WithLambdaErrs = False, _
  Optional SigmaLevel, Optional PercentErrsIn = False)
Attribute AgeErPb76.VB_Description = "Error in Pb7/6 age, input err is abs. 2sigma unless SigmaLevel=1 or PercentErrsIn=True.  t2 is end of Pb growth (default=0); output errs always 2sigma"
Attribute AgeErPb76.VB_ProcData.VB_Invoke_Func = " \n14"
' Return error in Pb-207/206 age. t2 (=end of Pb evolutions) is zero unless wpecified; decay-const
'  errs ignored unless withLambdaErrs=TRUE; SigmaLevel=SigLev unless specified (relevant only if
'  decay-const errs included).  Errs always returned at input Sigma-level!.
Dim Age, wLambdaErrs As Boolean, tmp1, SL0
ViM t2, 0
ViM WithLambdaErrs, False
SL0 = SigLev
ViM SigmaLevel, SigLev
If SigmaLevel <> 2 Or Not IsNumeric(SigmaLevel) Then SigmaLevel = 1
ViM PercentErrsIn, False
tmp1 = Menus("NumErr")
If Not IsNumeric(Pb76rad) Or Not IsNumeric(Pb76er) Then GoTo 1
IMN t2, 0
If PercentErrsIn Then Pb76er = Pb76er / Hun * Pb76rad
' NO   'Pb76er = Pb76er / IIf(SigmaLevel = 1, 1, 2) ' convert to 1-sigma if needed
' Pb76er is at SigmaLevel
Age = AgePb76(Pb76rad, t2)
If Not IsNumeric(Age) Then GoTo 1
If Age = BadT Then GoTo 1
tmp1 = PbPbAge(Pb76rad, t2, Age, Pb76er, WithLambdaErrs)
1:
SigLev = SL0
AgeErPb76 = tmp1
End Function

Public Function AgePb6U8(ByVal Pb6U8) ' 206Pb/238U age
Attribute AgePb6U8.VB_Description = "Age (Ma) from radiogenic 206Pb/238U"
Attribute AgePb6U8.VB_ProcData.VB_Invoke_Func = " \n14"
Static La238#
If IsNumeric(Pb6U8) Then
  If La238 = 0 Then
    GetConsts
    La238 = Lambda238
  End If
  AgePb6U8 = Log(1 + Pb6U8) / La238
Else
  AgePb6U8 = Menus("ValueErr")
End If
End Function

Public Function AgePb7U5(ByVal Pb7U5)  '207Pb/235U age
Attribute AgePb7U5.VB_Description = "Age (Ma) from radiogenic 207Pb/235U"
Attribute AgePb7U5.VB_ProcData.VB_Invoke_Func = " \n14"
Static La235#
If IsNumeric(Pb7U5) Then
  If La235 = 0 Then
    GetConsts
    La235 = Lambda235
  End If
  AgePb7U5 = Log(1 + Pb7U5) / La235
Else
  AgePb7U5 = Menus("ValueErr")
End If
End Function

Public Function AgePb8Th2(Pb208Th232)  '208Pb/232Th age
Attribute AgePb8Th2.VB_Description = "Age (Ma) from radiogenic 208Pb/232Th"
Attribute AgePb8Th2.VB_ProcData.VB_Invoke_Func = " \n14"
Static La232#
If IsNumeric(Pb208Th232) Then
  If La232 = 0 Then
    GetConsts
    La232 = Lambda232
  End If
  AgePb8Th2 = Log(1 + Pb208Th232) / La232
Else
  AgePb8Th2 = Menus("ValueErr")
End If
End Function

Public Function Age4corr(TotPb6U8, TotPb64, CommonPb64)
Attribute Age4corr.VB_Description = "Input is the total 206Pb/238U, total 206Pb/204Pb, and the common 206Pb/204Pb; ouput is the 206Pb/238U age corrected for common Pb using 204Pb as the index."
Dim r
r = TotPb6U8 * (1 - CommonPb64 / TotPb64)
Age4corr = Log(1 + r) / Lambda(238) / Million
End Function

Public Function AgeEr4Corr(TotPb6U8, TotPb6U8err, _
  TotPb64, TotPb64err, CommPb64, CommPb64err)
Attribute AgeEr4Corr.VB_Description = "Input is the total 206Pb/238U, error, total  206Pb/204Pb, error, and common 206Pb/204Pb, error; ouput is the error in the  204Pb-corrected age."
Dim CommonPbCorrErr#, CommPbCorrFact#, tmp#, Age#
Dim RatioFerr#, Rad68#, Rad68ferr#, TotPb6U8ferr#
Dim CommPbCorrFacterr, Rad68err, CommPbCorrFactFerr#
CommPbCorrFact = 1 - CommPb64 / TotPb64
Rad68 = TotPb6U8 * CommPbCorrFact
TotPb6U8ferr = TotPb6U8err / TotPb6U8
Age = Log(1 + Rad68) / Lambda(238) / Million
tmp = CommPb64err ^ 2 + (TotPb64err / TotPb64) ^ 2
CommPbCorrFactFerr = Sqr(tmp) / TotPb64 / CommPbCorrFact
Rad68ferr = Sqr(CommPbCorrFactFerr ^ 2 + TotPb6U8ferr ^ 2)
Rad68err = Rad68ferr * Rad68
AgeEr4Corr = Rad68err / (1 + Rad68) / Lambda(238) / Million
End Function

Public Function Age7corr(TotPb6U8, TotPb76, Comm76)
Attribute Age7corr.VB_Description = "Age from uncorrected Tera-Wasserburg ratios, assuming the specified common-Pb 207/206"
Attribute Age7corr.VB_ProcData.VB_Invoke_Func = " \n14"
' Given a total 206Pb/238U & 207Pb/206Pb, plus the common 207Pb/206Pb,
'  calculate the age assuming that the sample's radiogenic 206/238 &
'  207/235 ages are concordant.
Dim s#, T#, t1#, Deriv#, Delta#
Dim f#, e5#, e8#, Ee5#, Ee8#, TotPb7U5#
Dim Iter%, MaxIter%
If IsNumeric(TotPb6U8) And IsNumeric(TotPb76) And IsNumeric(Comm76) Then
  MaxIter = 999
  GetConsts
  TotPb7U5 = TotPb76 * Uratio * TotPb6U8
  T = 1000     ' Solve using Newton's method, using 1000 Ma as trial age.
  Do
    Iter = 1 + Iter
    e5 = Lambda235 * T
    If Abs(e5) > MAXEXP Then Age7corr = 0: Exit Function
    e8 = Exp(Lambda238 * T): e5 = Exp(e5)
    Ee8 = e8 - 1:            Ee5 = e5 - 1
    f = Uratio * Comm76 * (TotPb6U8 - Ee8) - TotPb7U5 + Ee5
    Deriv = Lambda235 * e5 - Uratio * Comm76 * Lambda238 * e8
    If Deriv = 0 Then Age7corr = NumErr: Exit Function
    Delta = -f / Deriv
    t1 = T + Delta
  If Abs(Delta) < 0.001 Then Exit Do
  If Iter > MaxIter Then Age7corr = NumErr: Exit Function
    T = t1
  Loop
  Age7corr = t1
Else
  Age7corr = Menus("ValueErr")
End If
End Function

Public Function AgeEr7Corr(Age, TotPb6U8, TotPb6U8err, _
  TotPb76, TotPb76err, CommPb76, CommPb76err)
' Calculation of 207-corrected age error.
Dim Numer#, Numer1#, Numer2#, Numer3#, Denom#
Dim e5#, e8#, Ee5#, Ee8#, TotPb7U5var, test#
If IsNumeric(Age) And IsNumeric(TotPb6U8) And IsNumeric(TotPb6U8err) And IsNumeric(TotPb76) _
  And IsNumeric(TotPb76err) And IsNumeric(CommPb76) And IsNumeric(CommPb76err) Then
  GetConsts
  TotPb7U5var = Uratio * Uratio * (SQ(TotPb6U8 * TotPb76err) + SQ(TotPb76 * TotPb6U8err))
  e5 = Lambda235 * Age
  If Abs(e5) > MAXEXP Then AgeEr7Corr = NumErr: Exit Function
  e8 = Exp(Lambda238 * Age): e5 = Exp(e5)
  Ee8 = e8 - 1:              Ee5 = e5 - 1
  Denom = SQ(Uratio * CommPb76 * Lambda238 * e8 - Lambda235 * e5)
  Numer1 = SQ(Uratio * (TotPb6U8 - Ee8) * CommPb76err)
  Numer2 = Uratio * Uratio * CommPb76 * (CommPb76 - 2 * TotPb76) * SQ(TotPb6U8err)
  Numer3 = TotPb7U5var
  Numer = Numer1 + Numer2 + Numer3
  AgeEr7Corr = Sqr(Numer / Denom)
Else
  AgeEr7Corr = Menus("ValueErr")
End If
End Function

Public Function Age8Corr(TotPb6U8, TotPb8Th2, Th2U8, CommPb68)
Attribute Age8Corr.VB_Description = "Age from uncorrected U-Pb ratios, assuming the specified common-Pb 208/206 & Th/Pb-U/Pb concordance."
Attribute Age8Corr.VB_ProcData.VB_Invoke_Func = " \n14"
' Pb208-corrected age -- that is, given the 206Pbtot/238U, 208Pbtot/232Th,
'  232Th/238U, & (206Pb/208Pb)common, calculate the sample age assuming
'  that the sample's radiogenic 206Pb/238U & 208Pb/232Th ages are concordant.
Dim s#, T#, t1#, Deriv#, Delta#
Dim f#, e2#, e8#, Iter%, MaxIter%
If IsNumeric(TotPb6U8) And IsNumeric(TotPb8Th2) And IsNumeric(Th2U8) And IsNumeric(CommPb68) Then
  MaxIter = 999
  GetConsts
  T = 1000  ' Solve using Newton's method, using 1000 Ma as trial age.
  Do
    Iter = 1 + Iter
    e8 = Lambda238 * T
    If Abs(e8) > MAXEXP Then Age8Corr = 0: Exit Function
    e2 = Exp(Lambda232 * T): e8 = Exp(e8)
    f = TotPb6U8 - e8 + 1 - Th2U8 * CommPb68 * (TotPb8Th2 - e2 + 1)
    Deriv = Th2U8 * CommPb68 * Lambda232 * e2 - Lambda238 * e8
    If Deriv = 0 Then Age8Corr = NumErr: Exit Function
    Delta = -f / Deriv
    t1 = T + Delta
  If Abs(Delta) < 0.001 Then Exit Do
  If Iter > MaxIter Then Age8Corr = NumErr: Exit Function
    T = t1
  Loop
  Age8Corr = t1
Else
  Age8Corr = Menus("ValueErr")
End If
End Function

Public Function AgeEr8Corr(T, TotPb6U8, TotPb6U8err, TotPb8Th2, TotPb8Th2err, _
  Th2U8, Th2U8err, CommPb68, CommPb68err)
Attribute AgeEr8Corr.VB_Description = "Error in 208-corrected age (input-ratio errors are absolute).  See Age8corr."
Attribute AgeEr8Corr.VB_ProcData.VB_Invoke_Func = " \n14"
' Error in Pb208-corrected age.

Dim e2#, e8#, P#, c2#, t1#, t2#, t3#, H#, G#, k#, SigmaA#, PsiI#, SigmaPsiI#, SigmaG#
Dim Numer#, Numer1#, Numer2#, Denom#, SigmaH#

If IsNumeric(T) And IsNumeric(TotPb6U8) And IsNumeric(TotPb6U8err) And _
  IsNumeric(TotPb8Th2) And IsNumeric(TotPb8Th2err) And IsNumeric(Th2U8) And _
  IsNumeric(Th2U8err) And IsNumeric(CommPb68) And IsNumeric(CommPb68err) Then
  GetConsts
  e8 = Lambda238 * T

  G = TotPb8Th2
  SigmaG = TotPb8Th2err
  H = Th2U8
  SigmaH = Th2U8err
  SigmaA = TotPb6U8err
  PsiI = CommPb68
  SigmaPsiI = CommPb68err

  If Abs(e8) > MAXEXP Then AgeEr8Corr = NumErr: Exit Function
  e2 = Exp(Lambda232 * T): e8 = Exp(e8)
  P = G + 1 - e2
  c2 = H * CommPb68

  t1 = SQ(H * SigmaG)
  t2 = SQ(P * SigmaH)
  t3 = SQ(H * P * SigmaPsiI / PsiI)
  k = Lambda238 * e8 - H * PsiI * Lambda232 * e2

  Numer = SQ(SigmaA) + SQ(PsiI) * (t1 + t2 + t3)
  Denom = k * k
  AgeEr8Corr = Sqr(Numer / Denom)

'  Numer1 = P * P * (SQ(H * SigmaPsiI) + SQ(CommPb68 * SigmaH))
'  Numer2 = SQ(c2 * SigmaG) + SQ(SigmaA)
'  Numer = Numer1 + Numer2
'  Denom = SQ(c2 * Lambda232 * e2 - Lambda238 * e8)
'  If Denom = 0 Then AgeEr8Corr = NumErr: Exit Function
'  AgeEr8Corr = Sqr(Numer / Denom)
Else
  AgeEr8Corr = Menus("ValueErr")
End If
Rcalc
End Function

Public Function ChiSquare(ByVal MSWD, ByVal DegFree&)
Attribute ChiSquare.VB_Description = "Probability of fit from MSWD and degrees of freedom"
Attribute ChiSquare.VB_ProcData.VB_Invoke_Func = " \n14"
' Determine probability that the observed MSWD or less will have been
'  generated by the assigned errors only.  Nu is the degrees of freedom.
If IsNumeric(MSWD) And IsNumeric(DegFree) Then
  If MSWD = 0 Or DegFree = 0 Then
    ChiSquare = 1
  Else
    ChiSquare = App.FDist(MSWD, DegFree, 1000000000#)
    'ChiDist(MSWD * DegFree, DegFree)
  End If
Else
  ChiSquare = Menus("ValueErr")
End If
End Function

Public Function Drnd(ByVal Number#, ByVal Sigfigs%)       ' Return Number rounded to SigFigs
Attribute Drnd.VB_Description = "Round a number to specified # of significant figures."
Attribute Drnd.VB_ProcData.VB_Invoke_Func = " \n14"
Dim A#, P&, d&, q, z% '  significant figures.
If IsNumeric(Number) And IsNumeric(Sigfigs) Then
  If Number = 0 Then
    Drnd = 0
  Else
    z = Int(Log10(Abs(Number)))
    q = 10# ^ z: A = Number / q
    P = 10# ^ (Sigfigs - 1): d = A * P
    Drnd = 1# * d * q / P
  End If
Else
  Drnd = Menus("ValueErr")
End If
End Function

Public Function StudentsT(ByVal DegFree, Optional ConfLimit = 95) ' ConfLimit in percent
Attribute StudentsT.VB_Description = "Students-t for specified degrees of freedom at 95%-confidence"
Attribute StudentsT.VB_ProcData.VB_Invoke_Func = " \n14"
IMN ConfLimit, 95
If IsNumeric(DegFree) Then
  StudentsT = App.TInv(1 - ConfLimit / Hun, DegFree)
Else
  StudentsT = Menus("ValueErr")
End If
End Function

Sub test()
Dim A(), i, j, v
ReDim A(2, 2)
For i = 1 To 2: For j = 1 To 2: A(i, j) = Rnd(3): Next j, i
v = WtdAv(A)
End Sub

Public Function WtdAv(ValuesAndErrors, Optional PercentOut = False, Optional PercentIn = False, _
  Optional SigmaLevelIn = 2, Optional CanReject = False, _
  Optional ConstantExternalErr = False, Optional SigmaLevelOut = 2)
' Array function to calculate weighted averages.  ValuesErrs is a 2-column range
'  with values & errors, PercentOut specifies whether output errors are absolute
'  (default) or %; PercentIn specifies input errors as absolute (default) or %;
'  SigmaLevel is sigma-level of input errors (default is 1-sigma); OUTPUT ALWAYS @2-SIGMA
'  CanRej permits rejection of outliers (default is False); If Probability<15%,
'  ConstExtErr specifies weighting by assigned errors plus a constant-external error
'  (default is weighting by assigned errors only).
' Errors are expanded by t-sigma-Sqrt(MSWD) if probability<10%.
' Output is a 2 columns of 5 (ConstExtErr=FALSE) or 6 (ConstExtErr=TRUE) rows, where
'  the left column contains values & the right captions.
Dim W(7, 2), SL0, s$, v, Nareas%, TotCols%, i%
s = TypeName(ValuesAndErrors)
If s = "Range" Then
  With ValuesAndErrors
    Nareas = .Areas.Count
    If Nareas = 1 Then
      TotCols = 2
    Else
      TotCols = 0
      For i = 1 To Nareas: TotCols = 1 + .Areas(i).Columns.Count: Next
    End If
    If .Rows.Count < 2 Or TotCols <> 2 Then WtdAv = Null: Exit Function
  End With
ElseIf InStr(s, "()") = 0 Then
  WtdAv = Null: Exit Function
Else
  v = 0
  On Error GoTo 1
  v = UBound(ValuesAndErrors, 2)
1:  On Error GoTo 0
  If v = 0 Then WtdAv = Null: Exit Function
End If

SL0 = SigLev
ViM PercentIn, False
ViM PercentOut, False
ViM SigmaLevelIn, 2
ViM SigmaLevelOut, 2
ViM CanReject, False
ViM ConstantExternalErr, False
On Error Resume Next
WeightedAv W, ValuesAndErrors, PercentOut, PercentIn, SigmaLevelIn, _
  CanReject, ConstantExternalErr, True, SigmaLevelOut
WtdAv = W
SigLev = SL0
End Function

Public Function SingleStagePbR(Age, WhichRatio%)
Attribute SingleStagePbR.VB_Description = "Stacey-Kramers single-stage 206/204, 207/204, or 208/204 (Which=0,1,2) from specified age"
Attribute SingleStagePbR.VB_ProcData.VB_Invoke_Func = " \n14"
' Pb-isotope ratio from age, assuming single-stage growth.
If IsNumeric(Age) And IsNumeric(WhichRatio) Then
  CalcPbgrowthParams True
  SingleStagePbR = PbR(Age, WhichRatio)
Else
  SingleStagePbR = Menus("ValueErr")
End If
End Function

Public Function SingleStagePbT(Pb206Pb204, Pb207Pb204)
Attribute SingleStagePbT.VB_Description = "Stacey-Kramers single-stage-growth age (Ma) from 206Pb/204 Pb and 207Pb/204Pb"
Attribute SingleStagePbT.VB_ProcData.VB_Invoke_Func = " \n14"
' Single-stage Pb-evolution age
Dim Age#, Mu#
If IsNumeric(Pb206Pb204) And IsNumeric(Pb207Pb204) Then
  CalcPbgrowthParams True
  SingleStagePbAgeMu Pb206Pb204, Pb207Pb204, Age, Mu
  SingleStagePbT = Age
Else
  SingleStagePbT = Menus("ValueErr")
End If
End Function

Public Function SingleStagePbMu(Pb206Pb204, Pb207Pb204)
Attribute SingleStagePbMu.VB_Description = "Stacey-Kramers single-stage Mu [=(238/204)today] from 206Pb/204 Pb and 207Pb/204Pb"
Attribute SingleStagePbMu.VB_ProcData.VB_Invoke_Func = " \n14"
' Single-stage Pb-evolution 238U/204Pb
Dim Age#, Mu#
If IsNumeric(Pb206Pb204) And IsNumeric(Pb207Pb204) Then
  CalcPbgrowthParams True
  SingleStagePbAgeMu Pb206Pb204, Pb207Pb204, Age, Mu
  SingleStagePbMu = Mu
Else
  SingleStagePbMu = Menus("ValueErr")
End If
End Function

Public Sub SuperRadIso() ' Superscript all numbers in the cells
Attribute SuperRadIso.VB_ProcData.VB_Invoke_Func = " \n14"
Dim c As Object          '   with mixed numbers-alpha.
For Each c In Selection.Cells
  Superscript Phrase:=c, AllNukes:=False
Next
End Sub

Public Sub SuperIso() ' Superscript all numbers in the cells
Attribute SuperIso.VB_Description = "Converts all numbers to superscripts."
Attribute SuperIso.VB_ProcData.VB_Invoke_Func = "S\n14"
Dim c As Object       '   with mixed numbers-alpha.
If TypeName(Selection) <> "Range" Then MsgBox "You must select a Worksheet range", , Iso: ExitIsoplot
For Each c In Selection.Cells
  Superscript Phrase:=c, AllNukes:=True
Next
End Sub

Public Function U234age(U234238ar#, InitialAR#) 'Returns 234U/238U age in kyr
Attribute U234age.VB_Description = "Age (ka) from present-day 234U/238U and initial 234U/.238U (as activity ratios)"
Attribute U234age.VB_ProcData.VB_Invoke_Func = " \n14"
Dim AgeYr#
If IsNumeric(U234238ar) And IsNumeric(InitialAR) Then
  GetConsts
  U234_Age U234238ar, InitialAR, 0, AgeYr, 0, 0
  U234age = AgeYr / Thou
Else
  U234age = Menus("ValueErr")
End If
End Function

Public Function U234ageAndErr(U234238ar#, U234err#, _
  InitialAR#, InitialARerr#, Optional PercentErrs As Boolean = False)
Attribute U234ageAndErr.VB_Description = "Array function that returns the 234U/238U age and error ( in kyr) as a  1x2 cell output.  Isotope ratios are activity ratios, errors are absolute unless PERCENT is specified as TRUE."
Attribute U234ageAndErr.VB_ProcData.VB_Invoke_Func = " \n14"
'Returns 234U/238U age and error in kyr
Dim AgeYr#, tmp#, Outp#(1, 2)
If IsNumeric(U234238ar) And IsNumeric(InitialAR _
  And IsNumeric(U234err) And IsNumeric(InitialARerr)) Then
  GetConsts
  U234_Age U234238ar, InitialAR, 0, AgeYr, 0, 0
  Outp(1, 1) = AgeYr / Thou
  If PercentErrs Then
    U234err = U234err / Hun * U234238ar
    InitialARerr = InitialARerr / Hun * InitialAR
  End If
  tmp = (U234err / (U234238ar - 1)) ^ 2 + (InitialARerr / (InitialAR - 1)) ^ 2
  Outp(1, 2) = Sqr(tmp) / Lambda234 / Thou
  U234ageAndErr = Outp
Else
  U234ageAndErr = Menus("ValueErr")
End If
End Function

Public Function Biweight(NumRange, Optional Tuning = 9) ' Tukey's biweight
Attribute Biweight.VB_Description = "Array function for Tukey's Biweight robust mean & error.  Input range is a column of values, output is a 3-row x 2-column range."
Attribute Biweight.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Tbi#, Sbi#, Err95#, xv#(), i%, j%, N&, M&, bw(3, 2)
IMN Tuning, 9
If Tuning <> 6 And Tuning <> 9 Then Tuning = 9
N = 0
If IsObject(NumRange) Then
  M = NumRange.Count
Else
  M = UBound(NumRange)
End If
ReDim xv(M)
For i = 1 To M
  If IsNumber(NumRange(i)) Then
    N = N + 1
    xv(N) = NumRange(i)
  End If
Next i
If N < 2 Then Exit Function
ReDim Preserve xv(N)
TukeysBiweight xv(), N, Int(Tuning), Tbi, Sbi, Err95
bw(1, 1) = Tbi:   bw(1, 2) = "Biweight Mean"
bw(2, 1) = Sbi:   bw(2, 2) = "Biweight Sigma"
bw(3, 1) = Err95: bw(3, 2) = "95%-conf. error"
Biweight = bw
Rcalc
End Function

Public Function CorrThU(Detritus As Object, Sample As Object, _
  Optional PercentIn = False, Optional PercentOut = False)
Attribute CorrThU.VB_Description = "Input is detrital 2/8,err,0/8,err,4/8,err [rho28-08,rho28-48,rho08-48], ditto for sample; output is corr. 230/238,err,234/238,err,rho"
Attribute CorrThU.VB_ProcData.VB_Invoke_Func = " \n14"
' Array function to correct a Th230-Th232-U234-U238 analysis for detrital
'  Th & U. The Detr & Sample arrays contain:
'  232/238, err, 230,238, err, 234/238, err,[,Rho28-08,Rho28-48,Rho08-48]
'  (The rho's are optional).  Input & Output errors can either absolute or percent
'  percent (default is absolute). Use the optional PercentIn & PercentOut
'   variables to specify.  Output is a 1 to 5 cell row with the detritus-corrected
' 230/238, err, 234/238, err, Rho.  ALL RATIOS ARE ACTIVITY RATIOS.
Dim dc%, sc%, s$, Bad As Boolean
Dim Detr#(9), Carb#(9), i%, cc#(5)
ViM PercentIn, False
ViM PercentOut, False
GetConsts
dc = Detritus.Columns.Count: sc = Sample.Columns.Count
If (dc <> 6 And dc <> 9) Or (sc <> 6 And sc <> 9) Then CorrThU = "Error": Exit Function
For i = 7 To 9
  Detr(i) = 0: Carb(i) = 0
Next i
For i = 1 To dc
  If i Mod 2 > 0 Or i > 6 Or Not PercentIn Then
    Detr(i) = Detritus(i)
  Else
    Detr(i) = Detritus(i) / Hun * Detritus(i - 1)
  End If
Next i
For i = 1 To sc
  If i Mod 2 > 0 Or i > 6 Or Not PercentIn Then
    Carb(i) = Sample(i)
  Else
    Carb(i) = Sample(i) / Hun * Sample(i - 1)
  End If
Next i
InitialCorr Detr(), Carb(), cc(1), cc(2), cc(3), cc(4), cc(5), Bad
If Bad Then CorrThU = "ERR": Exit Function
If PercentOut Then
  For i = 2 To 4 Step 2
    cc(i) = cc(i) / cc(i - 1) * Hun
  Next i
End If
CorrThU = cc
End Function

Function ModelAge(RockParent, RockRad, Optional Mtype = 2, Optional Depleted = False, _
  Optional SourceParent, Optional SourceRad, Optional DecayConst)
Attribute ModelAge.VB_Description = "Rb/Sr, Sm/Nd, Re/Os, or Lu/Hf model age (Mtype =1,2,3,4) from rock ratios (eg 147/144, 143/144), Source ratios"
Attribute ModelAge.VB_ProcData.VB_Invoke_Func = " \n14"
' Calculate  Rb-Sr, Sm-Nd, Lu-Hf, or Re-Os model ages.
' Required arguments are the Parent-isotope & radiogenic-isotope ratios of the
'   rock, relative to the daughter-element normalizing-isotope (eg 147Sm/144Nd
'   & 143Nd/144Nd); Mtype specifies what system (default is Sm/Nd); Depleted
'   specifies calculation of a depleted-mantle age; SourceParent & SourceRad are the
'   corresponding ratios for the source; DecayConst is the decay constant of the
'   parent isotope, in decays per year.
' If Depleted is specified, Mtype must be 2, the DePaolo quadratic constants
'   are used, & the function becomes an array function whose 1x3-cell output is
'   the depleted-mantle age, initial ratio, & Epsilon CHUR (assuming that the
'   Source ratios are CHUR ratios).
' If SourceRad, SourceParent, or DecayConst are unspecified, the stored Isoplot
'   values are used.
Dim test#, Tsimple#, Count%, t0#
Dim TT#, f#, mLambda, M As Object
Const Dcurve1 = 0.25, Dcurve2 = -3#, Dcurve3 = 8.5
Set M = Menus("ModelAgeParams")
ViM Mtype, 2
ViM Depleted, False
If IM(DecayConst) Then mLambda = M(Mtype, 1).Value Else mLambda = DecayConst
ViM SourceParent, Val(M(Mtype, 2))
ViM SourceRad, Val(M(Mtype, 3))
If IsNumeric(RockParent) And IsNumeric(RockRad) And IsNumeric(Mtype) And _
  IsNumeric(SourceParent) And IsNumeric(SourceRad) And IsNumeric(mLambda) Then
  mLambda = mLambda * Million ' Convert to decays/myr
  test = ((RockRad - SourceRad) / (RockParent - SourceParent) + 1)
  If test > MAXLOG Or test < MINLOG Then ModelAge = 0: Exit Function
  Tsimple = 1 / mLambda * Log(test)
  If Mtype = 2 And Depleted Then       ' Solve for Sm/Nd depleted-source model-age
    Dim Ma(3), Dcurve(3), Epsilon, Tdepl, DeplNd
    Tdepl = Tsimple
    Do
      t0 = Tdepl: TT = t0 / Thou
      Epsilon = TT * (Dcurve1 * TT + Dcurve2) + Dcurve3
      ' Epsilon(depl. source) @ T0
      f = 1 + Epsilon / 10000#
      test = (RockRad - f * SourceRad) / (RockParent - f * SourceParent) + 1
      If test > MAXLOG Or test < MINLOG Then ModelAge = "ERROR": Exit Function
      Tdepl = 1 / mLambda * Log(test)  ' Depleted-source model age
      Count = 1 + Count
      If Count > 100 Then ModelAge = "ERROR": Exit Function
    Loop Until Abs(Tdepl - t0) < 0.01
    test = mLambda * Tdepl
    If Abs(test) > MAXEXP Then ModelAge = "ERROR": Exit Function
    DeplNd = RockRad - RockParent * (Exp(test) - 1)
    '= CHUR 143/144 at depleted-source model-age
    Ma(1) = Tdepl: Ma(2) = DeplNd: Ma(3) = Epsilon
    ModelAge = Ma
  Else
    ModelAge = Tsimple
  End If
Else
  ModelAge = Menus("ValueErr")
End If
End Function

Public Function ConcordiaTW(U8Pb6, U8Pb6err, Pb76, Pb76err, Optional ErrCorrel, _
  Optional WithLambdaErrs, Optional PercentErrs, Optional SigmaLevel)
Attribute ConcordiaTW.VB_Description = "Basic input is 238/206,err,207/206,err; output is a 1x4-cell range of age, error, MSWD, probability.  Default input-errs are abs. 2-sigma; output err is 2sigma a priori."
Attribute ConcordiaTW.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns Concordia age for T-W concordia data
' See Concordia function for usage.
Dim X#, y#, Xerr#, Yerr#, Bad As Boolean, SL0, eC#
SL0 = SigLev
ViM ErrCorrel, 0
ViM WithLambdaErrs, False
ViM PercentErrs, False
ViM SigmaLevel, 2
If IsNumeric(U8Pb6) And IsNumeric(U8Pb6err) And IsNumeric(Pb76) Then
  If U8Pb6 > 0 And Pb76 > 0 Then
    X = U8Pb6: Xerr = U8Pb6err: y = Pb76: Yerr = Pb76err
    GetConsts
    If PercentErrs Then
      Xerr = Xerr / Hun * X: Yerr = Yerr / Hun * y
      PercentErrs = False:   AbsErrs = True
    End If
    ConcConvert X, Xerr, y, Yerr, eC, True, Bad
    If Bad Then Exit Function
    ConcordiaTW = Concordia(X, Xerr, y, Yerr, eC, WithLambdaErrs, _
      False, SigmaLevel)
  Else
    ConcordiaTW = Menus("ValueErr")
  End If
Else
  ConcordiaTW = Menus("ValueErr")
End If
SigLev = SL0
End Function

Public Function Concordia(Pb7U5, Pb7U5err, Pb6U8, Pb6U8err, ErrCorrel, _
  Optional WithLambdaErrs = False, Optional PercentErrs = False, Optional SigmaLevel = 2)
Attribute Concordia.VB_Description = "Basic input is 207/235,err,206/238,err,err-correl; output is a 1x4-cell range of age, error, MSWD, probability.  Default input-errs are abs. 2-sigma; output err is 2sigma a priori."
Attribute Concordia.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns Concordia age for Conv.-concordia data; Input the Concordia X,err,Y,err,RhoXY
' Output is 1 range of 4 values -- t, t-error (1-sigma apriori),MSWD,Prob-of-fit
' If a second row is included in the output range, include names of the 4 result-values.
' Output errors are always 2-sigma.
Dim i%, Dummy#(), CncA(2, 4), SL0
SL0 = SigLev
ReDim InpDat(1, 5)
ViM WithLambdaErrs, False
ViM PercentErrs, False
ViM SigmaLevel, 2
If SigmaLevel <> 1 Or SigmaLevel <> 1 Then SigmaLevel = 2
If IsNumeric(Pb7U5) And IsNumeric(Pb7U5err) And IsNumeric(Pb6U8) And _
  IsNumeric(Pb6U8err) And IsNumeric(ErrCorrel) Then
  If Pb7U5 > 0 And Pb6U8 > 0 Then
    InpDat(1, 1) = Pb7U5:  InpDat(1, 2) = Pb7U5err
    InpDat(1, 3) = Pb6U8:  InpDat(1, 4) = Pb6U8err
    InpDat(1, 5) = ErrCorrel
    For i = 2 To 4 Step 2
      If SigmaLevel = 2 Then InpDat(1, i) = InpDat(1, i) / 2
      If PercentErrs Then
        InpDat(1, i) = InpDat(1, i) / Hun * InpDat(1, i - 1)
      End If
    Next i
    Inverse = False:        Normal = True
    SigLev = SigmaLevel:    AbsErrs = True
    GetConsts
    ConcordiaAges Dummy(), 1, False, CncA(1, 1), CncA(1, 2), CncA(1, 3), CncA(1, 4), WithLambdaErrs
    CncA(2, 1) = "Age":  CncA(2, 2) = "2sig a priori err"
    CncA(2, 3) = "MSWD": CncA(2, 4) = "Probability"
    CncA(1, 2) = 2 * CncA(1, 2)   ' Always at 2-sigma
    Concordia = CncA()
  Else
    Concordia = Menus("ValueErr")
  End If
Else
  Concordia = Menus("ValueErr")
End If
SigLev = SL0
End Function

Public Function xyWtdAv(XYrange As Object, Optional PercentErrs = False, Optional SigmaLevel = 1)
Attribute xyWtdAv.VB_Description = "Array function Input an nx5 range of x,sigmax,y,sigmay,rhoxy; output is a 2x7 range (or less) with the wtd x-y mean & errors"
Attribute xyWtdAv.VB_ProcData.VB_Invoke_Func = " \n14"
' Input is a Nx5 range containing X,Xerr,Y,Yerr,RhoXY, where the
'  errors are 1-sigma absolute.  Output is a either 1 or 2 rows of up to 7 cells, where
'  the second (optional) row contains descriptions of the output cells, which are:
'  X,2-sigma a priori Y-error,Y,2-sigma a priori Y-erorr, RhoXY, MSWD, Prob-of-fit
' Output errors are always 2-sigma, a priori!
Dim i&, j&, k&, N&, df&, Bad As Boolean, OK() As Boolean, SL0
Dim SumsXY#, MSWD#, W#(2, 7), ww(2, 7) As Variant, nR&
SL0 = SigLev
ViM PercentErrs, False
ViM SigmaLevel, 1
N = 0: k = 0
nR = XYrange.Rows.Count
ReDim OK(nR)
For i = 1 To nR
  OK(i) = True
  For j = 1 To 5
    OK(i) = IsNumber(XYrange(i, j))
    If Not OK(i) Then Exit For
  Next j
  If OK(i) Then N = N + 1
Next i
If N = 0 Then Exit Function
ReDim InpDat(N, 5)
For i = 1 To nR
  If OK(i) Then
    k = k + 1
    For j = 1 To 5
      InpDat(k, j) = XYrange(i, j)
      If j = 2 Or j = 4 Then
        InpDat(k, j) = InpDat(k, j) / SigmaLevel
        If PercentErrs Then InpDat(k, j) = InpDat(k, j) / Hun * InpDat(k, j - 1)
      End If
    Next j
  End If
Next i
WtdXYmean InpDat(), N, W(1, 1), W(1, 3), SumsXY, W(1, 2), W(1, 4), W(1, 5), Bad
For i = 1 To 3 Step 2: ww(1, i) = W(1, i): Next i
ww(1, 2) = W(1, 2) * 2  ' Convert to 2-sigma a priori
ww(1, 4) = W(1, 4) * 2  '   "
ww(1, 5) = W(1, 5)
If Bad Then Exit Function
df = 2 * N - 2
If df <= 0 Then MSWD = 0 Else MSWD = SumsXY / df
ww(1, 6) = MSWD:     ww(1, 7) = ChiSquare(MSWD, df)
ww(2, 1) = "X mean": ww(2, 2) = "a priori err"
ww(2, 3) = "Y mean": ww(2, 4) = "a priori err"
ww(2, 5) = "correl": ww(2, 6) = "MSWD"
ww(2, 7) = "Probability"
xyWtdAv = ww()
SigLev = SL0
End Function

Public Function Gaussian(ByVal Mean#, ByVal Sigma#, Optional PercentErr = False)
Attribute Gaussian.VB_Description = "Returns a Normally-distributed random number with the specified mean and sigma."
Attribute Gaussian.VB_ProcData.VB_Invoke_Func = " \n14"
' Return a ramdom number with a mean of zero & a Gaussian distribution (from Numerical Recipes)
Static Iset%, Gset#
Dim GasDev#, v1#, v2#, r#, Fac#, aSigma#
ViM PercentErr, False
If PercentErr Then aSigma = Sigma / Hun * Mean Else aSigma = Sigma
If Iset Then
 GasDev = Gset
Else
 Do
  v1 = 2 * Rnd - 1: v2 = 2 * Rnd - 1
  r = v1 * v1 + v2 * v2
 Loop Until r <= 1
 Fac = Sqr(-2 * Log(r) / r)
 Gset = v1 * Fac: GasDev = v2 * Fac
End If
Iset = Not Iset
Gaussian = Mean + aSigma * GasDev
End Function

Public Function GaussCorr(X, sigmaX, y, SigmaY, RhoXY)
Attribute GaussCorr.VB_Description = "Array function that returns a pair of normally-distributed random numbers with the specified means , sigmas, and error correlation."
Attribute GaussCorr.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns a 2-cell range with gaussian-distributed, correlated X-Y values.
Dim W#(2)
GaussCorrel X, sigmaX, y, SigmaY, RhoXY, W(1), W(2)
GaussCorr = W
End Function

Sub Cleanup(Optional Quitting As Boolean)
Attribute Cleanup.VB_Description = "Remove hidden worksheets originally used for (now-deleted) Isoplot charts."
Attribute Cleanup.VB_ProcData.VB_Invoke_Func = " \n14"
' Remove hidden Isoplot PLOTDAT sheets for which the associated charts have been deleted.
Dim i%, j%, k%, Nused%, eX%
Dim s$, Wbk As Object, UsedNames$()
Dim Used As Boolean, DelName$(), Ndel%, T$, Sh As Object, c As Object
Dim nc%, d As Object, Co As Object, CanDelete As Boolean
IMN Quitting, False
If Not Quitting Then
  Xcalc
  NoUp
  GetOpSys
End If
Set Wbk = Awb
On Error Resume Next
For Each Sh In Wbk.Sheets
  If Sh.Type = xlXYScatter Then ' Isoplot chart?
    Set c = Sh.SeriesCollection
    GoSub GetSheetName
  ElseIf Sh.Type = xlWorksheet Then  ' Look for embedded charts moved into the worksheet
    For Each Co In Sh.ChartObjects
      Set c = Co.Chart.SeriesCollection
      GoSub GetSheetName
    Next
  End If
Next
On Error GoTo 0
' Need the "ucase" because moving a chart changes the formulas to upper-case.
For Each Sh In Wbk.Sheets
  On Error GoTo NextSheet  ' Will crash if more than one error.
  With Sh
    If UCase(Left$(.Name, 7)) = "PLOTDAT" And .Visible = False Then
      If .Cells(1, 1) = "Source sheet" Then
        Used = False
        For j = 1 To Nused
          If UsedNames$(j) = UCase(.Name) Then Used = True: Exit For
        Next j
        If Not Used Then
          Ndel = 1 + Ndel
          ReDim Preserve DelName$(Ndel)
          DelName$(Ndel) = .Name
        End If
      End If
    End If
  End With
NextSheet: On Error GoTo 0
Next
For i = 1 To Ndel
  If (i - 1) Mod 4 = 0 Then s$ = s$ & vbLf
  s$ = s$ & "  " & DelName$(i)
Next i
CanDelete = True
If Ndel = 0 And Not Quitting Then
  MsgBox "No unused hidden worksheets to delete.", , Iso
ElseIf CanDelete Then
  NoAlerts
  For i = 1 To Ndel
    With Sheets(DelName$(i)): .Visible = True: .Delete: End With
  Next i
  If Not Quitting Then
    T$ = IIf(Ndel > 1, "s", "")
    MsgBox "Hidden worksheet" & T$ & vbLf & s$ & viv$ & " deleted.", , Iso
  End If
End If
If Not Quitting Then Rcalc
Exit Sub

GetSheetName:               ' Parse-out the sheet name of the data for the
If c.Count = 0 Then Return  '  chart from the seriescollection formula.
T$ = c(1).Formula
eX = InStr(T$, "!"): k = eX
Do
  k = k - 1
Loop Until Mid$(T$, k, 1) = ","
PlotDat$ = Mid$(T$, k + 1, eX - k - 1)
If Len(PlotDat$) Then
  Nused = 1 + Nused
  ReDim Preserve UsedNames$(Nused)
  UsedNames$(Nused) = UCase(PlotDat$)
  Set ChrtDat = Sheets(PlotDat$)
End If
Return
End Sub

Public Function InitU234U238(AgeKyr, U234238ar)
Attribute InitU234U238.VB_Description = "Return the initial 234U/238U activity ratio for the specified age (ka) and present-day 23U/238U AR"
Attribute InitU234U238.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns initial U234/U238 activity ratio (onlyInitU234U238(AgeKyr, U234238ar))
Dim test#
If IsNumeric(AgeKyr) And IsNumeric(U234238ar) Then
  GetConsts
  test = AgeKyr * Thou * Lambda234
  If Abs(test) < MAXEXP Then
    InitU234U238 = 1 + (U234238ar - 1) * Exp(test)
  Else
    InitU234U238 = Menus("NumErr")
  End If
Else
  InitU234U238 = Menus("ValueErr")
End If
Rcalc
End Function

Public Function Th230age(Th230U238ar#, U234238ar#)
Attribute Th230age.VB_Description = "Return the 230Th/U age for the specified present-day 230Th/238U and 234U/238U activity ratios"
Attribute Th230age.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns 230Th/U age in kiloyears
Dim AgeYr#
If IsNumeric(Th230U238ar) And IsNumeric(U234238ar) Then
  If Th230U238ar >= 0 And U234238ar >= 0 Then
    GetConsts
    ThUage Th230U238ar, U234238ar, AgeYr
    Th230age = AgeYr / Thou
  Else
    Th230age = Menus("NumErr")
  End If
Else
  Th230age = Menus("ValueErr")
End If
End Function

Public Function Th230AgeAndInitial(Th230U238#, Th230Err#, _
  U234U238#, U234err#, _
  Optional RhoThU# = 0, Optional PercentErrs = False, _
  Optional SigmaLevel = 2, Optional WithLambdaErrs = False, Optional AtomRatios = False)
Attribute Th230AgeAndInitial.VB_Description = "Arran Fn, output is a 1x5 range with age,err,Gamma0,err,rho age-gamma0; default inputs errs are 2sigma abs., output errs same sigma as input"
Attribute Th230AgeAndInitial.VB_ProcData.VB_Invoke_Func = " \n14"
' 230Th/U age, initial 234/238, & errors from activity-ratio input.
' Returns Results() as 230Th/U age (ka), err, Gamma0, err, Rho T-Gamma0
' Sigma level must be 1 or 2; error output is 1-sigma
' Default is absolute, **2-sigma** input errors, no lambda errors, activity ratios, zero err-correl.
Dim Res#(5), SL0
SL0 = SigLev
ViM RhoThU, 0
ViM PercentErrs, False
ViM SigmaLevel, 2
ViM WithLambdaErrs, False
ViM AtomRatios, False
If SigmaLevel <> 2 Then SigmaLevel = 1

If IsNumeric(Th230U238) And IsNumeric(Th230Err) And IsNumeric(U234U238) _
  And IsNumeric(U234err) And IsNumeric(RhoThU) Then

  If Th230U238 > 0 And U234U238 > 0 And Abs(RhoThU) <= 1 Then
    Th230age_Gamma0 Res(), Th230U238, Th230Err, U234U238, U234err, _
      RhoThU, PercentErrs, SigmaLevel, WithLambdaErrs, AtomRatios
    Th230AgeAndInitial = Res
  Else
    Th230AgeAndInitial = Menus("NumErr")
  End If

Else
  Th230AgeAndInitial = Menus("ValueErr")
End If

SigLev = SL0
End Function

Public Function Th230U238ar(ByVal AgeKyr, ByVal InitU234U238ar)
Attribute Th230U238ar.VB_Description = "230Th/238U activity ratio for specified age (in ka) and initial 234U/238U AR"
Attribute Th230U238ar.VB_ProcData.VB_Invoke_Func = " \n14"
'Input is Age in kiloyears, initial 234/238 activity ratio (=Gamma0).
' Returns activity ratio of Th230/U238 for that age.
Dim Gamma#, T#, r#

If IsNumeric(AgeKyr) And IsNumeric(InitU234U238ar) Then
  If AgeKyr > 0 And InitU234U238ar > 0 Then
    T = AgeKyr * Thou
    GetConsts
    Gamma = 1 + (InitU234U238ar - 1) * Exp(-Lambda234 * T)
    Th230_U238ar T, Gamma, r
    Th230U238ar = r
  Else
    Th230U238ar = Menus("NumErr")
  End If
Else
  Th230U238ar = Menus("ValueErr")
End If
End Function

Public Function U234U238ar(ByVal AgeKyr, ByVal InitU234U238ar)
Attribute U234U238ar.VB_Description = "U234/U238 activity ratio for specified are (in ka) and initial U234/U238."
Attribute U234U238ar.VB_ProcData.VB_Invoke_Func = " \n14"
'Input is Age in kiloyears, initial 234/238 activity ratio.
' Returns activity ratio of U234/U238 for that age.
If IsNumeric(AgeKyr) And IsNumeric(InitU234U238ar) Then
  If AgeKyr >= 0 And InitU234U238ar >= 0 Then
    GetConsts
    U234U238ar = 1 + (InitU234U238ar - 1) * Exp(-Lambda234 * AgeKyr * Thou)
  Else
    U234U238ar = Menus("NumErr")
  End If
Else
  U234U238ar = Menus("ValueErr")
End If
End Function

Function ConvertConc(InputRange, Optional TeraWassIn = False, Optional PercentErrs = False)
Attribute ConvertConc.VB_Description = "Converts TW to conv Concordia data or vice-versa. Input range is X er Y er Rxy [Z er Rxz Ryz] or X er Y er [Z er] (TW only); output is 5 or 9-cell row: X er Y er Rxy [Z er Rxz Ryz]"
Attribute ConvertConc.VB_ProcData.VB_Invoke_Func = " \n14"
' Convert T-W concordia data to Conv., or vice-versa.
' Input is a row of 4,5,6,7, or 9 cols containing:
' (T-W):
' X er Y er  OR  X er Y er Rxy  OR  X er Y er Z er  OR  X er Y er Rxy Z er Rxz Ryz
' (Conv.):
' X er Y er Rxy  OR  X er Y er Rxy Z er Rxz Ryz
' Output is either   X er Y er Rxy Z er Rxz Ryz  OR  X er Y er Rxy
Dim d3 As Boolean, c%, TW6col As Boolean, tRw As Boolean, Pct As Boolean, Out#()
Dim X#, eX#, y#, eY#, rXY#, z#
Dim eZ#, rXZ#, rYZ#, Bad As Boolean, i%, r#()
ViM TeraWassIn, False
ViM PercentErrs, False
Pct = PercentErrs: tRw = TeraWassIn
c = InputRange.Columns.Count
If InputRange.Rows.Count <> 1 Or c < 4 Or c > 9 Then Exit Function
If Not tRw Then
  If c < 5 Then Exit Function
  If InputRange(5).Value = 0 Then Exit Function
End If
ReDim r(c)
For i = 1 To c: r(i) = InputRange(i).Value: Next i
d3 = (c > 5):       TW6col = (tRw And c = 6)
X = r(1): eX = r(2)
y = r(3): eY = r(4)
If Pct Then eX = eX / Hun * X: eY = eY / Hun * y
If c > 4 And Not TW6col Then
  rXY = r(5)
  If Abs(rXY) > 1 Then Exit Function
End If
GetConsts
If d3 Then
  z = r(6 + TW6col): eZ = r(7 + TW6col)
  If Pct Then eZ = eZ / Hun * z
  If Not TW6col Then
    If c > 7 Then rXZ = r(8)
    If c > 8 Then rYZ = r(9)
  End If
  ConcConvert X, eX, y, eY, rXY, tRw, Bad, z, eZ, rXZ, rYZ
Else
  ConcConvert X, eX, y, eY, rXY, tRw, Bad
End If
If Bad Then Exit Function
ReDim Out(5 - 4 * d3)
Out(1) = X:   Out(3) = y:  Out(5) = rXY
If Pct Then
  Out(2) = eX / X * Hun
  Out(4) = eY / y * Hun
Else
  Out(2) = eX: Out(4) = eY
End If
If d3 Then
  Out(6) = z:  Out(8) = rXZ: Out(9) = rYZ
  If Pct Then Out(7) = eZ / z * Hun Else Out(7) = eZ
End If
ConvertConc = Out
End Function

Function AlphaMS(Ratios, Optional PercentErrs = True)
Attribute AlphaMS.VB_Description = "Convert X=232/238 Y=230/238 Z=234/238 to X=238/232 Y=230/232 Z=234/232 & viceversa.  Input/output ranges are X,er,Y,er,Z,er[,rhoXY,rhoXZ,rhoYZ]; %errs default"
Attribute AlphaMS.VB_ProcData.VB_Invoke_Func = " \n14"
' Convert 232/238-230/238-234/238 ratios to 238/232-230/232-234/232 ratios
'   & vice-versa.  Input is a 6- or 9-cell range with
'   X,er,Y,er,Z,er [,Rxy,Rxz,Ryz].  Default errors are percent.
' ALPHA DATA MUST INCLUDE ERROR CORRELATIONS FOR THE CONVERSION TO BE VALID
Dim i%, Ncells%, Alpha(9)
Dim X#, y#, z#, A#, b#, c#
Dim Sx#, sY#, sZ#, sa#, SB#, sc#
Dim rXY#, rXZ#, rYZ#, Rab#
Dim Rac#, Rbc#, sx2#, sy2#, sz2#
ViM PercentErrs, True
Ncells = Ratios.Count
If Ncells <> 6 And Ncells <> 9 Then Exit Function
X = Ratios(1):  y = Ratios(3):  z = Ratios(5)
Sx = Ratios(2): sY = Ratios(4): sZ = Ratios(6)
If Ncells = 9 Then
  rXY = Ratios(7): rXZ = Ratios(8): rYZ = Ratios(9)
End If
If Not PercentErrs Then
  Sx = Hun * Sx / X: sY = Hun * sY / y * sZ = Hun * sZ / z
End If
A = 1 / X: b = y / X: c = z / X: sx2 = Sx * Sx: sy2 = sY * sY: sz2 = sZ * sZ
sa = Sx
SB = Sqr(sx2 + sy2 - 2 * Sx * sY * rXY)
sc = Sqr(sx2 + sz2 - 2 * Sx * sZ * rXZ)
Rab = (Sx - sY * rXY) / SB
Rac = (Sx - sZ * rXZ) / sc
Rbc = (sx2 - Sx * sY * rXY - Sx * sZ * rXZ + sY * sZ * rYZ) / (SB * sc)
If Not PercentErrs Then
  Sx = Sx / Hun * X: sY = sY / Hun * y: sZ = sZ / Hun * z
End If
Alpha(1) = A:   Alpha(3) = b:   Alpha(5) = c
Alpha(2) = sa:  Alpha(4) = SB:  Alpha(6) = sc
Alpha(7) = Rab: Alpha(8) = Rac: Alpha(9) = Rbc
AlphaMS = Alpha
End Function

Public Function Mad(v As Variant) ' Return the median absolute deviation from the median
Attribute Mad.VB_Description = "Returns MAD (Median Absolute Deviation from the Median) of the input range."
Attribute Mad.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i&, N&, MadVal#, MedianVal#, vv#()
If IsObject(v) Then N = v.Count Else N = UBound(v)
MedianVal = App.Median(v)
ReDim vv(N)
For i = 1 To N: vv(i) = v(i): Next i
GetMAD vv, N, MedianVal, MadVal, 0
Mad = MadVal
End Function

Public Function Lambda(NuclideMass%)
Attribute Lambda.VB_Description = "Returns decay const. for the long-lived nuclide whose mass is 'NuclideMass'."
Attribute Lambda.VB_ProcData.VB_Invoke_Func = " \n14"
Lambda = Val(Menus("Lambda" & Trim(Str(NuclideMass))))
End Function
Public Function HalfLife(NuclideMass%, Optional InMyr)
Attribute HalfLife.VB_Description = "Returns the half life for the long-lived radionuclide whose integral mass = NuclideMass.  Output is in years, unless InMyr is specified as TRUE, in which case output is in Myr."
Attribute HalfLife.VB_ProcData.VB_Invoke_Func = " \n14"
Dim test
ViM InMyr, False
test = Val(Menus("Lambda" & Trim(Str(NuclideMass))))
If test > 0 Then
  If InMyr Then test = test * 1000000#
  HalfLife = Log_2 / test
Else
  HalfLife = test
End If
End Function
Function La147#()
Attribute La147.VB_Description = "Returns the decay constant (decays/atom/yr) for Sm-147"
Attribute La147.VB_ProcData.VB_Invoke_Func = " \n14"
La147 = Lambda(147)
End Function
Function La230#()
Attribute La230.VB_Description = "Returns the decay constant (decays/atom/yr) for Th-230"
Attribute La230.VB_ProcData.VB_Invoke_Func = " \n14"
La230 = Lambda(230)
End Function
Function La232#()
Attribute La232.VB_Description = "Returns the decay constant (decays/atom/yr) for Th-232"
Attribute La232.VB_ProcData.VB_Invoke_Func = " \n14"
La232 = Lambda(232)
End Function
Function La234#()
Attribute La234.VB_Description = "Returns the decay constant (decays/atom/yr) for U-234"
Attribute La234.VB_ProcData.VB_Invoke_Func = " \n14"
La234 = Lambda(234)
End Function
Function La235#()
Attribute La235.VB_Description = "Returns the decay constant (decays/atom/yr) for U-235"
Attribute La235.VB_ProcData.VB_Invoke_Func = " \n14"
La235 = Lambda(235)
End Function
Function La238#()
Attribute La238.VB_Description = "Returns the decay constant (decays/atom/yr) for U-238"
Attribute La238.VB_ProcData.VB_Invoke_Func = " \n14"
La238 = Lambda(238)
End Function
Function La87#()
Attribute La87.VB_Description = "Returns the decay constant (decays/atom/yr) for Rb-87"
Attribute La87.VB_ProcData.VB_Invoke_Func = " \n14"
La87 = Lambda(87)
End Function
Function U238235#()
Attribute U238235.VB_Description = "Returns the present-day, natural 238U/235U atomic ratio."
Attribute U238235.VB_ProcData.VB_Invoke_Func = " \n14"
U238235 = Val(Menus("uratio"))
End Function
Function Pb76(AgeMa) ' Return radiogenic 207Pb/206Pb (secular equilbrium)
Attribute Pb76.VB_Description = "Returns the radiogenic Pb-207/Pb-206 for the specified age in Ma."
Attribute Pb76.VB_ProcData.VB_Invoke_Func = " \n14"
GetConsts
If AgeMa = 0 Then
  Pb76 = Lambda235 / Lambda238 / Uratio
Else
  Pb76 = (Exp(Lambda235 * AgeMa) - 1) / (Exp(Lambda238 * AgeMa) - 1) / Uratio
End If
End Function

Sub GetConsts(Optional ForceGet = False) ' Retrieve Isoplot consts from the MenuItems data-sheet
Attribute GetConsts.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, Lambdas As Range, dce As Range, rAr As Range, rUr As Range
Dim rMg As Range, rAM As Range, rMP As Range, InitAR As Range, dc As Range
LambdaRef = Array(0, 10, 4, 5, 6, 9, 10, 0, 0, 1, 2, 3, 0, 0, 0, 0, 0, 10, 10)
' Maps decay constants in the DecayConsts range or the Menu Sheet to the Lambdas range,
'  which is ordered by Isotype.
ViM ForceGet, False
'App.ReferenceStyle = xlA1  'CAUSES CRASH IF ENABLED
If ForceGet Or Lambda238 = 0 Or Uratio = 0 Then
  pm = Chr(177): pmm = " " & pm & " "
  With MenuSht
    Set rUr = .Range("Uratio")
    Set rAr = .Range("_Air4036"): Set rMg = .Range("ArMinGas")
    Set rAM = .Range("ArMinSteps"): Set rMP = .Range("ArMinProb")
    Set Lambdas = .Range("Lambdas").Cells
    Set dc = .Range("DecayConsts").Cells
    Set dce = .Range("DecayConstPerrs").Cells
    ReDim iLambda(Lambdas.Count)
    On Error Resume Next
    For i = 1 To Lambdas.Count ' iLambda is indexed to plot-type, and is in 1/myr
      If LambdaRef(i) <> 0 Then
        iLambda(i) = dc(LambdaRef(i)) * Million
        Lambdas(i) = iLambda(i) ' can't do this if called from a function
      End If
    Next i
    On Error GoTo 0
    U_1.AssignDirectLambdaNames
    If IsEmpty(rUr) Or rUr.Value = 0 Then rUr = 137.88
    Uratio = rUr.Value
    MinProb = MinMax(0.05, 0.3, .Range("MinProb"))
    Set InitAR = .Range("PbUinitAR")
    NumErr = .Range("NumErr")
    If IsEmpty(rAr) Then rAr = 295.5
    If rAr <= 0 Then rAr = 295.5
    Air4036 = rAr.Value
    If IsEmpty(rAM) Or rAM.Value <= 1 Then rAM = 3
    ArMinSteps = Max(2, rAM.Value)
    If IsEmpty(rMP) Then rMP = 0.05
    ArMinProb = MinMax(0.01, 30, rMP.Value)
    If IsEmpty(rMg) Or rMg.Value = 0 Then rMg = 60
    ArMinGas = rMg.Value
  End With
  Yrat(1, 1) = InitAR(1):  Yrat(1, 2) = InitAR(2):  Yrat(2, 1) = InitAR(3)
  ' Only for symbol font!
  GetOpSys
  On Error GoTo 1
  If Workbooks.Count = 0 Then Workbooks.Add
  ActiveSheet.PageSetup.PaperSize = xlPaperLetter
  Exit Sub
1: On Error Resume Next
ActiveSheet.PageSetup.PaperSize = xlPaperA4
End If
Exit Sub
CantCalc: On Error GoTo 0
ExitIsoplot
End Sub

Sub PutPb76Age(Optional ShiftOutput = 0, Optional NoCheck = False)
Attribute PutPb76Age.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Rin As Range, i&, Rr As Range, N&, j%, OK As Boolean, f$, Ro As Range
Dim Fr&, Lr&  ' Put the Pb-Pb age 1 column to right of selected (Pb76*) range
ViM ShiftOutput, 0
ViM NoCheck, False
NoUp
If TypeName(Selection) <> "Range" Then MsgBox "Must start from a worksheet range", , Iso: ExitIsoplot
Set Rin = Selection
If Rin.Columns.Count > 1 And Not NoCheck Then _
  MsgBox "Selection must be 1 column of ratios", , Iso: ExitIsoplot
If Rin.Rows.Count = EndRow Then Set Rin = SelectFromWholeCols(Rin) ' selection is complete cols
N = Rin.Rows.Count: j = 2 + ShiftOutput
Set Ro = Range(Rin(1, j), Rin(N, j))
If Not NoCheck Then
  OK = True
  For i = 1 To N
    If Not IsEmpty(Ro(i)) Then OK = False: Exit For
  Next i
  If Not OK Then
    f$ = "This will cause the selected cells (see worksheet)"
    f$ = f$ & vbLf & "to be over-written.  Do you want to proceed?"
    Ro.Select
    If MsgBox(f$, vbOKCancel, Iso) = vbCancel Then Rin.Select: KwikEnd
  End If
End If
f$ = "=" & TW.Name & "!AgePb76("
For i = 1 To N
  OK = False
  Set Rr = Rin(i)
  If IsNumber(Rr, True) Then
    Rin(i, j).Formula = f$ & Rr.Address & ")"
  Else
    Rin(i, j) = ""
  End If
Next i
Ro.NumberFormat = "0"
Rin(1).Select
End Sub

Sub PutPb76AgeAndErr() ' Put the Pb-Pb age & err 2 columns to right of selected (Pb76*) range
Attribute PutPb76AgeAndErr.VB_ProcData.VB_Invoke_Func = " \n14"
' Assumes Pb76err is 1-col to right of ratio
Dim pc$, Rat, eR, i&, j%, f$, Rin As Range, N&, OK As Boolean, Ro As Range
Dim Fr&, Lr&
NoUp
If TypeName(Selection) <> "Range" Then MsgBox "You must select a Worksheet range", , Iso: ExitIsoplot
Set Rin = Selection
If Rin.Columns.Count <> 2 Then _
  MsgBox "Selected range must include both ratios and errors", , Iso: ExitIsoplot
If Rin.Rows.Count = EndRow Then Set Rin = SelectFromWholeCols(Rin) ' selection is complete cols
N = Rin.Rows.Count
Set Ro = Range(Rin(1, 3), Rin(N, 4))
OK = True
For i = 1 To N
  For j = 1 To 2
    If Not IsEmpty(Ro(i, j)) Then OK = False
Next j, i
If Not OK Then
  f$ = "This will cause the selected cells (see worksheet)"
  f$ = f$ & vbLf & "to be over-written.  Do you want to proceed?"
  Ro.Select
  If MsgBox(f$, vbOKCancel, Iso) = vbCancel Then Rin.Select: KwikEnd
End If
Range(Rin(1, 1), Rin(N, 1)).Select
PutPb76Age 1, True
f$ = "=" & TW.Name & "!AgeErPb76("
For i = 1 To N
  Set Rat = Rin(i, 1): Set eR = Rin(i, 2)
  If IsNumber(Rat, True) And IsNumber(eR, True) Then
    If Len(pc$) = 0 Then
      pc$ = IIf(eR / Rat > 0.3, "True", "False")
    End If
    Rin(i, 4).Formula = f$ & Rat.Address & "," & eR.Address & ",,,," & pc$ & ")"
  Else
    Rin(i, 4) = ""
  End If
Next i
HA Range(Rin(1, 3), Rin(N, 3)), xlRight
With Range(Rin(1, 4), Rin(N, 4))
  .NumberFormat = pm & "0": .HorizontalAlignment = xlLeft
End With
Rin(1, 1).Select
End Sub

Sub PutWtdAv()
Attribute PutWtdAv.VB_ProcData.VB_Invoke_Func = " \n14"
Dim pc As Boolean, OK As Boolean, tr%, k%, SL%, c%
Dim i&, j&, f$, N&, Br&
Dim Ro As Range, vi As Range, Rin As Range, RV As Range
NoUp
If TypeName(Selection) <> "Range" Then MsgBox "You must select a Worksheet range", , Iso: ExitIsoplot
Set Rin = Selection
N = Rin.Rows.Count
If Rin.Columns.Count <> 2 Or N < 2 Then _
  MsgBox "Must select 2 columns (value, error) and at least 2 rows for auto Wtd Av", , Iso: ExitIsoplot
If N = EndRow Then Set Rin = SelectFromWholeCols(Rin)
With Rin
  N = .Rows.Count: c = .Column
  tr = .Row: Br = tr + N - 1
End With
j = N
For i = 1 To Rin.Rows.Count
  j = j + (Not IsNumber(Rin(i, 1)) And Not IsNumber(Rin(i, 2), True))
Next i
If j < 2 Then MsgBox "Need at least 2 value-error pairs for Wtd Av", , Iso: ExitIsoplot
For i = 1 To Rin.Rows.Count
  If Not IsNumber(Rin(i, 1)) Or Not IsNumber(Rin(i, 2), True) Then _
    MsgBox "Need both value and error for each item", , Iso: ExitIsoplot
Next i
Set Ro = sR(Br + 2, c, Br + 7, c + 1)
Set vi = Range(Ro(8, 1), Ro(12, 2)): Set RV = App.Union(Ro, vi)
OK = True
For i = 1 To 12
  For j = 1 To 2
    If Not IsEmpty(Ro(i, j)) Then OK = False
Next j, i
If Not OK Then
  NoUp False
  RV.Select
  If MsgBox("Output will overwrite selected area", vbOKCancel, Iso) <> vbOK _
   Then Rin.Select: KwikEnd
End If
NoUp
pc = False: SL = 2
If tr > 1 Then
  f$ = LCase(Rin(0, 2).Text)
  If InStr(f$, "%") > 0 Or InStr(f$, "percent") > 0 Then pc = True
  If InStr(f$, "1 sig") Or InStr(f$, "1s") Or InStr(f$, "1-sig") Then SL = 1
End If
f$ = "=" & TW.Name & "!WtdAv("
On Error GoTo ArrPres
vi(0, 1) = "User specifies"
vi(1, 1) = False: vi(1, 2) = "% out"
vi(2, 1) = pc:    vi(2, 2) = "% in"
vi(3, 1) = SL:    vi(3, 2) = "Sigma level"
vi(4, 1) = True:  vi(4, 2) = "Can reject"
vi(5, 1) = False: vi(5, 2) = "Const. ext. err"
With RV: .NumberFormat = General: .Font.Bold = False: .Font.Underline = False: End With
vi(3, 2).NumberFormat = General
HA Range(Ro(1, 1), Ro(13, 1)), xlRight
HA Range(Ro(1, 2), Ro(13, 2)), xlLeft
With vi(0, 1): .HorizontalAlignment = xlLeft: .Font.Underline = True: End With
Ro.FormulaArray = f$ & Rin.Address & "," & vi(1, 1).Address & "," & vi(2, 1).Address & _
  "," & vi(3, 1).Address & "," & vi(4, 1).Address & "," & vi(5, 1).Address & ")"
Ro(1, 1).Select
Exit Sub
ArrPres: If Err = 1004 Then
  MsgBox "Output range occupied by existing Array Function", , Iso
Else
  MsgBox "Error in weighted-average function", , Iso
End If
Ro(1, 1).Select
End Sub

Sub PutYorkfit()
Attribute PutYorkfit.VB_ProcData.VB_Invoke_Func = " \n14"
Dim pc As Boolean, i&, j&, f$, Rin As Range, N&, Nn&
Dim tr%, Br&, c%, OK As Boolean, Ro As Range, SL%
Dim nc%, Sx, sY, rXY, Bad As Boolean, te, s1$, s2$, v, k%
Dim d() As DataPoints
NoUp
If TypeName(Selection) <> "Range" Then MsgBox "You must select a Worksheet range", , Iso: KwikEnd
Set Rin = Selection
nc = Rin.Columns.Count: N = Rin.Rows.Count
If nc < 4 Or nc > 5 Or N < 2 Then
  MsgBox "Must select 4 or 5 columns" & viv$ & _
    "x   x-err   y   y-err   [xy err-correl]" & viv$ & _
    "and at least 2 rows for Yorkfit", , Iso
  KwikEnd
End If
If N = EndRow Then Set Rin = SelectFromWholeCols(Rin)
With Rin
  N = .Rows.Count: c = .Column
  tr = .Row: Br = tr + N - 1
End With
If N > 9999 Then MsgBox "Too many data-points", , Iso: KwikEnd
Nn = N
pc = True: SL = 2
For i = -(tr = 1) To 1 ' Looks for header row both as 1st row in selection & in row above
  For j = 2 To 4 Step 2
    f$ = LCase(Rin(i, j).Text)
    If InStr(f$, "%") = 0 And InStr(f$, "percent") = 0 And _
      (InStr(f$, "abs") > 0 Or InStr(f$, "err") > 0) Then pc = False
    If InStr(f$, "1 sig") > 0 Or InStr(f$, "1s") > 0 Or InStr(f$, "1-s") > 0 Then SL = 1
  Next j
Next i
k = 0
ReDim d(Nn)
For i = 1 To Nn
  OK = True
  For j = 1 To nc
    OK = True: v = Rin(i, j)
    If Not IsNumber(v) Then OK = False: Exit For
    If ((j = 1 Or j = 3) And v = 0) Or ((j = 2 Or j = 4) And v < 0) _
      Or (j = 5 And Abs(v) > 1) Then OK = False: Exit For
  Next j
  If OK Then
    k = 1 + k
    With d(k)
      For j = 1 To nc
        v = Rin(i, j)
        Select Case j
          Case 1: .X = v
          Case 2: Sx = v / SL
          Case 3: .y = v
          Case 4: sY = v / SL
          Case 5: .RhoXY = v
        End Select
      Next j
      If pc Then Sx = Sx / Hun * .X: sY = sY / Hun * .y
      .Xerr = Sx: .Yerr = sY
      If nc < 5 Then .RhoXY = 0
    End With
  End If
Next i
Nn = k
If Nn < 2 Then MsgBox "Need at least 2 valid value-error pairs for Yorkfit", , Iso: ExitIsoplot
ReDim Preserve d(Nn)
Set Ro = sR(Br + 2, c, Br + 12, c + 1)
OK = True
For i = 1 To 11
  For j = 1 To 2
    If Not IsEmpty(Ro(i, j)) Then OK = False
Next j, i
If Not OK Then
  NoUp False
  Ro.Select
  If MsgBox("Output will overwrite selected area", vbOKCancel, Iso) <> vbOK _
   Then Rin.Select: KwikEnd
End If
NoUp
York_Fit Nn, d, Bad, True
If Bad Then MsgBox "Yorkfit failed", , Iso: ExitIsoplot
On Error GoTo ArrPres
Ro.Clear
With Ro: .NumberFormat = General: .Font.Bold = False: .Font.Underline = False: End With
HA Range(Ro(1, 1), Ro(11, 1)), xlRight
HA Range(Ro(1, 2), Ro(11, 2)), xlLeft
With yf
  Ro(1, 1) = .Slope:         Ro(2, 1) = Drnd(.ErrSlApr, 2)
  Ro(3, 1) = Drnd(.SlopeError, 3)
  Ro(4, 1) = .Intercept:     Ro(5, 1) = Drnd(.ErrIntApr, 3)
  Ro(6, 1) = Drnd(.InterError, 3)
  Ro(7, 1) = Drnd(.RhoInterSlope, 5)
  Ro(8, 1) = Drnd(.MSWD, 3): Ro(9, 1) = Drnd(.Prob, 3)
End With
Ro(1, 2) = "Slope": Ro(2, 2) = pm & "1sigma a priori": Ro(3, 2) = pm & "95% conf."
Ro(4, 2) = "Intercept": Ro(5, 2) = Ro(2, 2): Ro(6, 2) = Ro(3, 2)
Ro(7, 2) = "Rho(slope-inter)"
Ro(8, 2) = "MSWD": Ro(9, 2) = "Prob. of fit"
s1$ = IIf(pc, "percent", "absolute")
s2$ = tSt(SL)
Ro(11, 1) = "for " & s2$ & "-sigma " & s1$ & " input errors"
HA Ro(11, 1), xlLeft
Rin.Select
Exit Sub
ArrPres: f$ = IIf(Err = 1004, "Output range occupied by existing Array Function", "Error in Yorkfit")
Ro(1, 1).Select
End Sub

Sub PutRobustRegr()
Attribute PutRobustRegr.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, Rin As Range, N&, f$
Dim tr%, Br&, c%, OK As Boolean, Ro As Range
Dim nc%, Slope#, Lslope#, Uslope#
Dim Intercept#, Lint#, Uint#, tB As Boolean
NoUp
If TypeName(Selection) <> "Range" Then MsgBox "You must select a Worksheet range", , Iso: ExitIsoplot
Set Rin = Selection
With Rin
  If .Areas.Count > 1 Then MsgBox "Input range must be contiguous", , Iso: KwikEnd
  N = .Rows.Count: c = .Column: nc = .Columns.Count
End With
If N >= EndRow Then Set Rin = SelectFromWholeCols(Rin)
With Rin
  N = .Rows.Count
  If N > 9999 Then MsgBox "Too many data-points", , Iso: KwikEnd
  tr = .Row: Br = tr + N - 1
End With
Set Ro = sR(Br + 2, c, Br + 8, c + 1)
OK = True
For i = 1 To 6
  For j = 1 To 2
    If Not IsEmpty(Ro(i, j)) Then OK = False
Next j, i
If Not OK Then
  NoUp False
  Ro.Select
  If MsgBox("Output will overwrite selected area", vbOKCancel, Iso) <> vbOK _
   Then Rin.Select: KwikEnd
End If
RobustReg2 Rin, Slope, Lslope, Uslope, Intercept, , Lint, Uint, DoCheck:=True
Ro.Clear
With Ro: .NumberFormat = General: .Font.Bold = False: .Font.Underline = False: End With
HA Range(Ro(1, 1), Ro(6, 1)), xlRight
HA Range(Ro(1, 2), Ro(6, 2)), xlLeft
Ro(1, 2) = "Slope":     Ro(1, 1) = Slope
Ro(2, 2) = "+error": Ro(2, 1) = Drnd(Uslope - Slope, 2)
Ro(3, 2) = "-error": Ro(3, 1) = Drnd(Slope - Lslope, 2)
Ro(4, 2) = "Intercept": Ro(4, 1) = Intercept
Ro(5, 2) = "+error": Ro(5, 1) = Drnd(Uint - Intercept, 2)
Ro(6, 2) = "-error": Ro(6, 1) = Drnd(Intercept - Lint, 2)
Ro(7, 1) = "(errors are 95% conf.)"
HA Range(Ro(1, 1), Ro(6, 1)), xlRight
HA Range(Ro(1, 2), Ro(6, 2)), xlLeft
HA Ro(7, 1), xlLeft
Range(Rin.Address).Select
End Sub

Sub ButtonInvokedPlot(Itype%, Inver As Boolean, DimThree As Boolean)
InitializePlotTypes
Isotype = Itype: Inverse = Inver
StP.DropDowns("dIsotype") = Itype
Isoplot , Isotype, Inverse, DimThree
End Sub
Sub InvokeConcPlotNormal()
ConcPlot = True
ButtonInvokedPlot 1, False, False
End Sub
Sub InvokeConcPlotInverse()
ConcPlot = True
ButtonInvokedPlot 1, True, False
End Sub
Sub InvokeAnyXY()
OtherXY = True
Isoplot , 14, False, False
End Sub
Sub InvokeWtdAvPlot()
WtdAvPlot = True
Isoplot , 15, False, False
End Sub
Sub InvokeArArIsochronNormal()
Isoplot , 2, False, False
End Sub
Sub InvokeArArIsochronInverse()
Isoplot , 2, True, False
End Sub

Sub InvokeRbSrPlot()
Attribute InvokeRbSrPlot.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 3, False, False
End Sub
Sub InvokeSmNdPlot()
Attribute InvokeSmNdPlot.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 4, False, False
End Sub
Sub InvokeReOsPlot()
Attribute InvokeReOsPlot.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 5, False, False
End Sub
Sub InvokeArPlateau()
Attribute InvokeArPlateau.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 18, False, False
End Sub
Sub InvokePbPbPlotNormal()
Attribute InvokePbPbPlotNormal.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 8, False, False
End Sub
Sub InvokePbPbPlotInverse()
Attribute InvokePbPbPlotInverse.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 8, True, False
End Sub
Sub InvokeMuAlphaNormal()
Attribute InvokeMuAlphaNormal.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 10, False, False
End Sub
Sub InvokeMuAlphaInverse()
Attribute InvokeMuAlphaInverse.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 10, True, False
End Sub
Sub InvokeProbDensity()
Attribute InvokeProbDensity.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 16, False, False
End Sub

Sub InvokeLinearizedProb(Optional ForResids)
Attribute InvokeLinearizedProb.VB_ProcData.VB_Invoke_Func = " \n14"
' User invoked linear-probability plot.  Must be from a single-column range.
' Alternatively, called from Isoplot to to a LinProbPlot of the weighted residuals of a just-completed regression.
Dim nR%, nc%, Na%, dum1 As wWtdAver, dum2(), dum3&
Dim s$, cc As Object, c As Object, i%, j%, f As Object, tbx As Object
Dim X As Object, P As Object, y As Object, xw!, xl!, sn%
Dim Clr0 As Boolean, plot0 As Boolean, ax0$, tmp!, Yn As Object, InpDat0#()
Dim pRegr0 As Boolean, w0 As Boolean
IMN ForResids, False
If ForResids Then ' Cache the x-y input data
  On Error GoTo 1
  For i = 1 To N: tmp = tmp + Abs(yf.WtdResid(i)): Next i
  If tmp = 0 Then Exit Sub
  On Error GoTo 0
  MatCopy InpDat(), InpDat0()
  With yf
    For i = 1 To N: InpDat(i, 1) = .WtdResid(i): Next i
  End With
  Clr0 = ColorPlot: ax0$ = AxX$: pRegr0 = pRegress: w0 = WtdAvPlot
Else
  RangeCheck nR, ndCols, Na
  N = nR
  If N = 0 Then Exit Sub
  ReDim InpDat(N, 5)
  For i = 1 To N
    InpDat(i, 1) = 1 * Selection(i)
    tmp = tmp + Abs(InpDat(i, 1))
  Next i
  If tmp = 0 Then Exit Sub
End If
If ndCols = 1 Or nR < 3 Or ForResids Then
  NoUp
  ProbPlot = True: WtdAvPlot = False: pDots = True: pRegress = True: ColorPlot = True
  If Not ForResids Then
    Set DatSht = Ash: DatSheet$ = DatSht.Name
  End If
  Set cc = DlgSht("ProbPlot").CheckBoxes
  cc("cInclStats") = xlOn: cc("cRegression") = xlOn
  SymbCol = 3: SymbRow = 1: DoShape = True
  Sheets.Add:  PlotDat$ = "PlotDat"
  AssignIsoVars
  MakeSheet PlotDat$, ChrtDat
  pFirst = 1: pLast = N
  WtdAverPlot N, dum1, dum2(), 0, ""
  PutPlotInfo
  If ForResids Then MatCopy InpDat0(), InpDat()
  Set f = ActiveChart
  With f
    Set P = .PlotArea: Set c = .ChartArea: Set y = .Axes(xlValue)
    Set X = .Axes(xlCategory)
    .SeriesCollection("ProbX").DataLabels.Font.Size = 15
    With .SeriesCollection(1)
      sn = .Points.Count
      .MarkerSize = 11 + (sn > 12) + (sn > 36) + (sn > 80)
    End With
    y.HasTitle = ForResids
    If ForResids Then
      Set Yn = y.AxisTitle
      With Yn
        .Caption = "Weighted residuals": .Font.Size = 30: .Font.Bold = False
        If ConcAge Then .Caption = .Caption & " of equivalence"
      End With
    End If
  End With
  c.Interior.Color = RGB(212, 255, 212)
  'p.Interior.Color = RGB(212, 212, 212)
  y.TickLabels.Font.Size = 16
  If ForResids Then tmp = X.Left - Yn.Left Else tmp = P.Width - X.Width
  P.Width = y.Height + tmp + 20  ' So plotbox width=height
  If ForResids Then P.Left = Yn.Font.Size + 2 + 30 Else P.Left = 30
  P.Top = 30 ' Reduce margins
  Set tbx = f.TextBoxes("ProbLine")
  xw = X.Width: xl = X.Left
  c.Width = Right_(X) + 30
  If ForResids Then tmp = X.Left - Yn.Left Else tmp = P.Width - X.Width
  P.Width = y.Height + tmp + 20   ' So box width=box height
  If ForResids Then P.Left = Yn.Font.Size + 2 + 30 Else P.Left = 30
  P.Top = 20 ' Reduce margins
  X.AxisTitle.Top = c.Height
  With tbx ' Increase font size, put back at lower right of plotbox
    .Characters.Font.Size = 17
    .Left = Right_(X) - .Width - 2
    .Top = Bottom(X) - .Height - 8
    .HorizontalAlignment = xlCenter
  End With
  c.Border.LineStyle = xlContinuous
  P.Select
  CopyPicture
  DelSheet f
  DelSheet ChrtDat
  ProbPlot = False: WtdAvPlot = w0: ColorPlot = Clr0
  AxX$ = ax0$
Else
  Isoplot , 17, 0, 0
End If
1: On Error GoTo 0
End Sub
Sub InvokeUseriesEvolution()
Attribute InvokeUseriesEvolution.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 13, True, False
End Sub
Sub InvokeUseriesIsochron()
Attribute InvokeUseriesIsochron.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 13, True, True
End Sub
Sub InvokeArPlateauIsochron()
Attribute InvokeArPlateauIsochron.VB_ProcData.VB_Invoke_Func = " \n14"
Isoplot , 19, True, False
End Sub
Sub InvokeMix()
Attribute InvokeMix.VB_ProcData.VB_Invoke_Func = " \n14"
'Isoplot , 23, False, False
RangeCheck 0, 0, 0 'nr, Nc, Na
Mix
End Sub
Sub InvokeBracket()
Attribute InvokeBracket.VB_ProcData.VB_Invoke_Func = " \n14"
'Isoplot , 21, False, False
GetOpSys '
SetupBracket False, False
End Sub
Sub InvokeUseriesBracket()
'Isoplot , 22, False, False
GetOpSys
SetupBracket False, True
End Sub
Sub InvokePlanar3DconcPbU()
Attribute InvokePlanar3DconcPbU.VB_ProcData.VB_Invoke_Func = " \n14"
StPc("cCalculate") = xlOn
Isoplot , 1, True, True, False
End Sub
Sub InvokeTotalPbU()
Attribute InvokeTotalPbU.VB_ProcData.VB_Invoke_Func = " \n14"
StPc("cCalculate") = xlOn
Isoplot , 1, True, True, True
End Sub

Public Function YorkSlope(xyValuesAndErrs As Range, Optional PercentErrors, Optional SigmaLevel)
Attribute YorkSlope.VB_Description = "Returns the slope of a 2-error x-y regression (Yorkfit).  The xyValuesAndErrs range contains x,x-err,y,y-err,[err-correl].  Optional SigmaLevel (default= 1) & PercentErrs (default=False) describe x-y errors."
Attribute YorkSlope.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d() As DataPoints, Bad As Boolean, SL0
SL0 = SigLev
ViM SigmaLevel, 1
ViM PercentErrors, False
GetDat xyValuesAndErrs, d(), N, SigmaLevel, PercentErrors, Bad
If Not Bad Then YorkSlope = yf.Slope Else YorkSlope = NumErr
SigLev = SL0
End Function

Public Function YorkInter(xyValuesAndErrs As Range, Optional PercentErrors, Optional SigmaLevel)
Attribute YorkInter.VB_Description = "Returns the intercept of a 2-error x-y regression (Yorkfit).  The xyValuesAndErrs range contains x,x-err,y,y-err,[err-correl].  Optional SigmaLevel (default= 1) & PercentErrs (default=False) describe x-y errors."
Attribute YorkInter.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d() As DataPoints, Bad As Boolean, SL0
SL0 = SigLev
ViM SigmaLevel, 1
ViM PercentErrors, False
GetDat xyValuesAndErrs, d(), N, SigmaLevel, PercentErrors, Bad
If Not Bad Then YorkInter = yf.Intercept Else YorkInter = NumErr
SigLev = SL0
End Function

Public Function YorkMSWD(xyValuesAndErrs As Range, Optional PercentErrors, Optional SigmaLevel)
Attribute YorkMSWD.VB_Description = "Returns the MSWD of a 2-error x-y regression (Yorkfit).  The xyValuesAndErrs range contains x,x-err,y,y-err,[err-correl].  Optional SigmaLevel (default= 1) & PercentErrs (default=False) describe x-y errors."
Attribute YorkMSWD.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d() As DataPoints, Bad As Boolean, SL0
SL0 = SigLev
ViM SigmaLevel, 1
ViM PercentErrors, False
GetDat xyValuesAndErrs, d(), N, SigmaLevel, PercentErrors, Bad
If Not Bad Then YorkMSWD = yf.MSWD Else YorkMSWD = NumErr
YorkMSWD = IIf(Bad, NumErr, yf.MSWD)
SigLev = SL0
End Function

Public Function YorkProb(xyValuesAndErrs As Range, Optional PercentErrors, Optional SigmaLevel)
Attribute YorkProb.VB_Description = "Returns probability-of-fit for a 2-error x-y regression (Yorkfit).  The xyValuesAndErrs range contains x,x-err,y,y-err,[err-correl].  Optional SigmaLevel (default= 1) & PercentErrs (default=False) describe x-y errors."
Attribute YorkProb.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d() As DataPoints, Bad As Boolean, SL0
SL0 = SigLev
ViM SigmaLevel, 1
ViM PercentErrors, False
GetDat xyValuesAndErrs, d(), N, SigmaLevel, PercentErrors, Bad
If Not Bad Then YorkProb = yf.Prob Else yf.Prob = NumErr
SigLev = SL0
End Function

Public Function YorkSlopeErr95(xyValuesAndErrs As Range, Optional PercentErrors, Optional SigmaLevel)
Attribute YorkSlopeErr95.VB_Description = "Returns 2-sigma (prob-fit>0.15) or 95%conf. (prob-fit<0.15) slope error of a Yorkfit.  The xyValuesAndErrs range contains x,x-err,y,y-err,[err-correl]. The Optional SigmaLevel default= 1, PercentErrs default=False."
Attribute YorkSlopeErr95.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d() As DataPoints, Bad As Boolean, SL0
SL0 = SigLev
ViM SigmaLevel, 1
ViM PercentErrors, False
GetDat xyValuesAndErrs, d(), N, SigmaLevel, PercentErrors, Bad
If Not Bad Then YorkSlopeErr95 = yf.SlopeError Else yf.SlopeError = NumErr
SigLev = SL0
End Function

Public Function YorkInterErr95(xyValuesAndErrs As Range, Optional PercentErrors, Optional SigmaLevel)
Attribute YorkInterErr95.VB_Description = "Returns 2-sigma (prob-fit>0.15) or 95%conf. (prob-fit<0.15) intercept error of a Yorkfit.  The xyValuesAndErrs range contains x,x-err, y,y-err, [err-correl]. The Optional SigmaLevel default= 1, PercentErrs default=False."
Attribute YorkInterErr95.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d() As DataPoints, Bad As Boolean, SL0
ViM SigmaLevel, 1
ViM PercentErrors, False
GetDat xyValuesAndErrs, d(), N, SigmaLevel, PercentErrors, Bad
If Not Bad Then YorkInterErr95 = yf.InterError Else yf.InterError = NumErr
End Function

Private Sub GetDat(r As Variant, DP() As DataPoints, ByVal N&, _
  ByVal SigmaLevel%, ByVal Percent As Boolean, Bad As Boolean)
  ' 09/06/19 -- Change passed parameter R from Range to Variant
Dim i&, c%, SL0
Bad = True: SL0 = SigLev

If TypeName(r) = "Range" Then         ' /
  With r                              '|
    If .Areas.Count > 1 Then GoTo 1   '| 09/06/19
    c = .Columns.Count                '| Mods to deal with possibility of R
    If c < 4 Or c > 5 Then GoTo 1     '|   being either a Range or numeric
    N = .Rows.Count                   '|   array.
  End With                            '|
Else                                  '|
  N = UBound(r, 1)                    '|
  c = UBound(r, 2)                    '|
End If                                ' \

ReDim DP(N)
For i = 1 To N
  DP(i).X = r(i, 1): DP(i).y = r(i, 3)
  DP(i).Xerr = r(i, 2) / SigmaLevel
  DP(i).Yerr = r(i, 4) / SigmaLevel
  If c = 5 Then DP(i).RhoXY = r(i, 5)
  If Percent Then
    DP(i).Xerr = DP(i).Xerr / Hun * DP(i).X
    DP(i).Yerr = DP(i).Yerr / Hun * DP(i).y
  End If
Next i
York_Fit N, DP(), Bad, True
1: SigLev = SL0
End Sub

Sub WtdXYmean(Pts#(), ByVal Npts&, Xbar#, Ybar#, _
  SumsXY#, ErrX#, ErrY#, RhoXY#, Bad As Boolean)
' Calculate the position, errs, & err corr of the weighted mean
'  of a suite of X-Y data pts.
Dim i&, j&, Xvar#, Yvar#, Cov#, WtdResid#
Dim Om11#, Om12#, Om22#, A#, b#, c#
Dim Alpha#, Beta#, Denom#, Rx#, Ry#
Dim s1#, s2#, Oh#()
ReDim Oh(Npts, 3)
For i = 1 To Npts
  Xvar = SQ(Pts(i, 2)): Yvar = SQ(Pts(i, 4))
  Cov = Pts(i, 5) * Sqr(Xvar * Yvar)
  Inv2x2 Xvar, Yvar, Cov, Om11, Om22, Om12, Bad
  If Bad Then GoTo BadXY
  A = A + Om11: b = b + Om22: c = c + Om12
  Alpha = Alpha + Pts(i, 1) * Om11 + Pts(i, 3) * Om12
  Beta = Beta + Pts(i, 3) * Om22 + Pts(i, 1) * Om12
  Oh(i, 1) = Om11: Oh(i, 2) = Om22: Oh(i, 3) = Om12
Next i
Denom = A * b - c * c
If Denom = 0 Then GoTo BadXY
Xbar = (b * Alpha - Beta * c) / Denom
Ybar = (A * Beta - Alpha * c) / Denom
' Can calculate Sums without specifically calculating the residuals, & so
'  without storing the Om11, Om22, Om12, but requires slightly more
'  calculational effort & suspect is more vulnerable to roundoff errors.
SumsXY = 0
ReDim yf.WtdResid(Npts)
For i = 1 To Npts
 Rx = Pts(i, 1) - Xbar: Ry = Pts(i, 3) - Ybar
 s1 = Rx * Rx * Oh(i, 1) + Ry * Ry * Oh(i, 2)
 s2 = 2 * Rx * Ry * Oh(i, 3)
 WtdResid = s1 + s2
 SumsXY = SumsXY + WtdResid
 yf.WtdResid(i) = Sqr(WtdResid)
Next i
' Now calculate the variance-covariance matrix of Xbar,Ybar
Inv2x2 A, b, c, vcXY(1, 1), vcXY(2, 2), vcXY(1, 2), Bad
If Bad Then GoTo BadXY
vcXY(2, 1) = vcXY(1, 2)
ErrX = Sqr(vcXY(1, 1)): ErrY = Sqr(vcXY(2, 2))
RhoXY = vcXY(1, 2) / (ErrX * ErrY)
Bad = 0
Exit Sub

BadXY: Bad = -1
MsgBox "Can't determine the X-Y Weighted Mean for these data", , Iso
KwikEnd
End Sub

Sub ConcordiaAges(xyProj#(), ByVal Npts&, Bad As Boolean, Optional pubT, _
  Optional pubTerr, Optional pubMSWD, Optional pubProb, Optional WLE = True)
' Calculate the weighted X-Y mean of the data pts (including error
'  correlation) & the "Concordia Age" & age-error of Xbar, Ybar.
' The "Concordia Age" is the most probable age of a data point if one can assume
'  that the U/Pb ages of the true data point are precisely concordant.  Calcu-
'  lates the age & error both with & without uranium decay-constant errors.
' See GCA 62, p. 665-676, 1998 for explanation.
Dim BadRhos%, RhoBad As Boolean, dfT%, Pub As Boolean
Dim i%, k%
Dim ProbEquiv#, TrialAge#, MswdXY#, r#, Mult95#, Mult95xy#
Dim SumsAge1#, MswdAgeOne#, SumsAgeOneNLE#, MswdAgeMany#
Dim rXY#, xp#, yp#, b#, Ap#, Bp#, Rab#, MswdAgeOneNLE#
Dim ErrX95#, ErrY95#, SumsAgeManyNLE#, MswdAgeManyNLE#
Dim AgeBrak1#, AgeBrak2#, EM#, SumsXY#, ErrX#, ErrY#, T#
Dim SigmaAge#, TT#, ss#, Xbar#, Ybar#, RhoXY#
Dim SumsAgeMany#, SumsAgeOne#, tNLE#, SigmaAgeNLE#
Const BrentToler = 0.0000001
' Xbar, ErrX are the input X +-1sigma
' Ybar, ErrY are the input Y +-1sigma
' Xconc, Yconc are the Conv. Conc. X-Y
ViM WLE, True
Pub = NIM(pubT)
Xbar = 0: Ybar = 0
dfT = 2 * Npts - 1 ' for the weighted-mean concordant age
If Npts = 1 Then
  Xbar = InpDat(1, 1):  Ybar = InpDat(1, 3)
  ErrX = InpDat(1, 2):  ErrY = InpDat(1, 4)
  RhoXY = InpDat(1, 5): ProbEquiv = 1
Else
  WtdXYmean InpDat(), Npts, Xbar, Ybar, SumsXY, ErrX, ErrY, RhoXY, Bad
  If Bad Then Exit Sub  ' plus error message???
  ShowXYwtdMean Xbar, ErrX, Ybar, ErrY, RhoXY, SumsXY, _
    MswdXY, ProbEquiv, Npts, EM
  xyProj(0) = EM
  xyProj(1) = Xbar: xyProj(2) = EM * ErrX
  xyProj(3) = Ybar: xyProj(4) = EM * ErrY
  xyProj(5) = RhoXY
End If
If ProbEquiv < 0.001 Then
  If Pub Then Bad = True: Exit Sub
  GoTo XYdone
End If

If Inverse Then Call ConcConvert(Xbar, ErrX, Ybar, ErrY, RhoXY, True, Bad)
If Bad Then ExitIsoplot

Cmisc.Xconc = Xbar: Cmisc.Yconc = Ybar
vcXY(1, 1) = ErrX * ErrX: vcXY(2, 2) = ErrY * ErrY
vcXY(1, 2) = RhoXY * ErrX * ErrY
vcXY(2, 1) = vcXY(1, 2)
TrialAge = 1 + Cmisc.Yconc
If TrialAge < MINLOG Or TrialAge > MAXLOG Or Cmisc.Yconc <= 0 Then
  tNLE = 0: T = 0
  GoTo BadTnle
Else
  tNLE = Log(TrialAge) / Lambda238
  TrialAge = tNLE: Cmisc.NoLerr = True
  If tNLE > 0 Then AgeNLE Cmisc.Xconc, Cmisc.Yconc, vcXY(), tNLE
  If tNLE Then
    SumsAgeOneNLE = ConcordSums(tNLE, Bad)
    MswdAgeOneNLE = SumsAgeOneNLE
    VarTcalc tNLE, SigmaAgeNLE, Bad
    SumsAgeManyNLE = SumsXY + SumsAgeOneNLE
    MswdAgeManyNLE = SumsAgeManyNLE / dfT
  End If
End If
If (Lambda235err > 0 Or Lambda238err > 0) And (Not Pub Or WLE) Then
  Cmisc.NoLerr = False
  ' Find optimum age for Xbar,Ybar, & MSWD for concordance of Xbar,Ybar.
  AgeBrak1 = 1000:  AgeBrak2 = 1100
  MNBRAK AgeBrak1, TrialAge, AgeBrak2, 0, 0, 0, Bad
  If TrialAge = 0 Then Bad = True
  T = 0
  If Not Bad Then BRENT AgeBrak1, TrialAge, AgeBrak2, 0.0000001, TT, ss, Bad
  If Not Bad Then
    T = TT:   SumsAgeOne = ss
    VarTcalc T, SigmaAge, Bad
    MswdAgeOne = SumsAgeOne
    SumsAgeMany = SumsXY + SumsAgeOne
    MswdAgeMany = SumsAgeMany / dfT
  End If
End If
BadTnle:
If tNLE = 0 And T = 0 Then
  If Pub Then
    Bad = True: Exit Sub
  Else
    MsgBox "Unable to solve for a Concordia Age", , Iso
  End If
Else
  If Pub Then
    If Not WLE Then
      pubT = tNLE: pubTerr = SigmaAgeNLE
      pubMSWD = MswdAgeOneNLE: pubProb = ChiSquare(pubMSWD, 1)
    Else
      pubT = T: pubTerr = SigmaAge
      pubMSWD = MswdAgeOne: pubProb = ChiSquare(pubMSWD, 1)
    End If
  Else
    ShowConcAge T, SigmaAge, MswdAgeOne, MswdAgeMany, tNLE, SigmaAgeNLE, _
      MswdAgeOneNLE, MswdAgeManyNLE, Npts
  End If
End If
If Pub Then Exit Sub
XYdone:
ConcAgePlot = (Npts > 1 And Xbar <> 0 And Ybar <> 0 And ProbEquiv >= 0.001)
End Sub

Sub RobustReg2(xy As Variant, Slope#, Optional Lslope#, Optional Uslope#, _
  Optional Yint#, Optional Xint#, Optional Lyint#, _
  Optional Uyint#, Optional Lxint#, Optional Uxint#, _
  Optional DoCheck = True, Optional SlopeOnly = False, Optional WithXinter As Boolean = False)
' Robust linear regression using median of all pairwise slopes/intercepts,
' after Hoaglin, Mosteller & Tukey, Understanding Robust & Exploratory Data Analysis,
' John Wiley & Sons, 1983, p. 160, with errors from code in Rock & Duffy, 1986
' (Comp. Geosci. 12, 807-818), derived from Vugrinovich (1981), J. Math. Geol. 13,
'  443-454).
' Has simple, rapid solution for errors.
Dim X#(), y#()
Dim N&, M&, Slp#(), Yinter#()
Dim LwrInd&, UpprInd&, Vs#, Vy#, vx#
Dim Xinter#()
ViM SlopeOnly, False
ViM DoCheck, True
NoUp
MakeXY xy, X(), y(), N, (DoCheck)
If N < 3 Then MsgBox "Need 3 or more x-y pairs", , Iso: ExitIsoplot
If M > EndRow Then
  MsgBox "Can't do robust regression for N>360"
  KwikEnd
End If
GetRobSlope X(), y(), N, M, False, False, 0, 0, 0, Slp(), Yinter(), Xinter()
If M > EndRow Then
  MsgBox "Can't do robust regression for N>360"
  KwikEnd
End If
GetRobSlope X(), y(), N, M, False, WithXinter, Slope, Yint, Xint, Slp(), Yinter(), Xinter()
Erase X, y
If Not SlopeOnly Then Conf95 N, (M), LwrInd, UpprInd
QuickSort Slp()
QuickSort Yinter()
Slope = iMedian(Slp(), True)
If Not SlopeOnly Then
  Lslope = Slp(LwrInd): Uslope = Slp(UpprInd)
  Yint = iMedian(Yinter(), True)
  Lyint = Yinter(LwrInd): Uyint = Yinter(UpprInd)
  If WithXinter Then
    QuickSort Xinter()
    Xint = iMedian(Xinter(), True)
    Lxint = Xinter(LwrInd): Uxint = Xinter(UpprInd)
  End If
End If
End Sub

Sub NoUp(Optional Yes = True)
ViM Yes, True
App.ScreenUpdating = Not Yes
End Sub

Sub AgeNLE(ByVal X#, ByVal y#, VarCov#(), T#)
' Using a 2-D Newton's method, find the age for a presumed-concordant
'  point on the U-Pb Concordia diagram that minimizes Sums,
'  assuming no decay-constant errors.
' See GCA 62, p. 665-676, 1998 for explanation.
Dim ct%, Bad As Boolean
Dim Om11#, Om22#, Om12#, t2#, e5#, e8#
Dim Ee5#, Ee8#, Q5#, Q8#, Qq5#, Qq8#
Dim Rx#, Ry#, d1#, d2a#, d2b#, d2#
Dim Incr#, test#
Const MaxCt = 1000, Toler = 0.000001
Inv2x2 VarCov(1, 1), VarCov(2, 2), VarCov(1, 2), Om11, Om22, Om12, Bad
If Bad Then GoTo NoAgeSoln
t2 = T
Do
  ct = 1 + ct
  T = t2
  e5 = Lambda235 * T
  If ct = MaxCt Or Abs(e5) > MAXEXP Then GoTo NoAgeSoln
  e5 = Exp(e5):         e8 = Exp(Lambda238 * T)
  Ee5 = e5 - 1:         Ee8 = e8 - 1
  Q5 = Lambda235 * e5:  Q8 = Lambda238 * e8
  Qq5 = Lambda235 * Q5: Qq8 = Lambda238 * Q8
  Rx = X - Ee5:         Ry = y - Ee8
  ' First derivative of T w.r.t. S, times -0.5
  d1 = Rx * Q5 * Om11 + Ry * Q8 * Om22 + (Ry * Q5 + Rx * Q8) * Om12
  ' Second derivative of T w.r.t. S, times +0.5
  d2a = (Q5 * Q5 + Qq5 * Rx) * Om11 + (Q8 * Q8 + Qq8 * Ry) * Om22
  d2b = (2 * Q5 * Q8 + Ry * Qq5 + Rx * Qq8) * Om12
  d2 = d2a + d2b
  If d2 = 0 Then GoTo NoAgeSoln
  Incr = d1 / d2
  test = Abs(Incr / T)
  t2 = T + Incr
Loop Until test < Toler
T = t2
Exit Sub
NoAgeSoln:  'Print "Unable to solve for age of these data"
T = 0
End Sub

Sub CleanupIsoRefs(Optional Dummy)
Dim s As Worksheet, c As Object, FirstAddress$, P%, q%, f$, rf$, L%
Dim cb As Object, cn As Object, u%
App.ReferenceStyle = xlA1
For Each s In ActiveWorkbook.Worksheets
  StatBar s.Name
  With s.Cells
    Set c = .Find(What:="iso*.xla'!", after:=.Cells(1, 1), LookIn:=xlFormulas, LookAt:= _
      xlPart, SearchDirection:=xlNext)
    If Not c Is Nothing Then
      FirstAddress = c.Address
      Do
        f = LCase(c.Formula)
        If Left(f, 1) = "=" Then
          rf = RevStr(f): L = Len(f)
          P = InStr(rf, "!'alx."): q = InStr(rf, "osi")
          If q > P Then c.Formula = "=" & Mid(f, L - P + 2)
        End If
        Set c = .FindNext(c)
      If c Is Nothing Then Exit Do
      Loop While c.Address <> FirstAddress
    End If
  End With
  App.Calculate
Next s
On Error GoTo 1
For Each cb In App.CommandBars
  With cb
    For Each cn In .Controls
      On Error GoTo 2
      If cn.Type = 1 Then
        f = LCase(cn.OnAction): P = InStr(f, "!")
        q = InStr(f, "isoplot"): u = InStr(f, ".xla")
        If P > 0 And q > P And u > q Then cn.OnAction = Mid(f, P + 1)
      End If
      On Error GoTo 0
2:  Next cn
  End With
Next cb
1: On Error GoTo 0
StatBar
MsgBox "All references to Isoplot updated."
End Sub

Sub TukeysBiweight(X#(), ByVal N&, ByVal Tuning%, _
  Tbi#, Sbi#, Err95#)
' Calculates Tukey's biweight estimator of location & scale.
' Tbi is a very robust estimator of "mean", Sbi is the robust estimator of
'   "sigma".  These estimators converge to the true mean & true sigma for
'   Gaussian distributions, but are very resistant to outliers.
' The lower the "Tuning" constant is, the more the tails of the distribution
'   are effectively "trimmed" (& the more robust the estimators are against
'   outliers), with the price that more "good" data is disregarded.  pts
'   that deviate from the "mean" greater that "Tuning" times the "standard
'   deviation" are assigned a weight of zero ('rejected').
' Err95 is the 95% conf-limit on Tbi.  "Gaussian" is returned as -1 if
'   the distribution of x() appears to be normal (at 95%-conf. limit), 0 if  not.
' Adapted & inferred from Hoaglin, Mosteller, & Tukey, 1983, Understanding
'   Robust & Exploratory Data Analysis: John Wiley & Sons, pp. 341, 367,
'   376-378, 385-387, 423,& 425-427.
Dim j%, Iter%, Tuner#, MedianVal#
Dim Snsum#, Sdsum#, Tnsum#, Delta#, u#
Dim U1#, U2#, U5#, U12#, Madd#
Dim LastTbi#, LastSbi#, TbiDelt#, SbiDelt#
Dim W#, T#, TbiMatch As Boolean, SbiMatch As Boolean
Const MaxIter = 100, Small = 1E-30, ZerTest = 0.0000000001, NonzerTest = 0.0000000001
MedianVal = iMedian(X()) ' Initial estimator of location is Median.
Tbi = MedianVal
GetMAD X(), N, Tbi, Madd, 0   ' Initial estimator of scale is MAD.
Sbi = Max(Madd, Small)
Do
  Iter = Iter + 1
  Tuner = Tuning * Sbi
  Snsum = 0: Sdsum = 0: Tnsum = 0
  For j = 1 To N
    Delta = X(j) - Tbi
    If Abs(Delta) < Tuner Then
      u = Delta / Tuner
      U2 = u * u      ' U^2
      U1 = 1 - U2     ' 1-U^2
      U12 = U1 * U1   '(1-U^2)^2
      U5 = 1 - 5 * U2 ' 1-5U^2
      Snsum = Snsum + SQ(Delta * U12)
      Sdsum = Sdsum + U1 * U5
      Tnsum = Tnsum + u * U12
    End If
  Next j
  LastTbi = Tbi: LastSbi = Sbi
  Sbi = Sqr(N * Snsum) / Abs(Sdsum)
  If Sbi < Small Then Sbi = Small
  Tbi = LastTbi + Tuner * Tnsum / Sdsum ' Newton-Raphson method
  TbiDelt = Abs(Tbi - LastTbi)
  SbiDelt = Abs(Sbi - LastSbi)
  If Tbi = 0 Then
    TbiMatch = (TbiDelt < ZerTest)
  Else
    TbiMatch = ((TbiDelt / Tbi) < NonzerTest)
  End If
  If Sbi = 0 Then
    SbiMatch = (SbiDelt < ZerTest)
  Else
    SbiMatch = ((SbiDelt / Sbi) < NonzerTest)
  End If
Loop Until (TbiMatch And SbiMatch) Or Iter > MaxIter
If Sbi <= Small Then Sbi = 0
' t-approx. for near-Gaussian distr's; from Monte Carlo
'  simulations followed by Simplex fit (valid for Tuning=9).
Select Case N
  Case 2, 3: T = 47.2  ' really only for N=3
  Case 4:    T = 4.736
  Case Is >= 5
    W = N - 4.358
    T = 1.96 + 0.401 / Sqr(W) + 1.17 / W + 0.0185 / (W * W)
End Select
Err95 = T * Sbi / Sqr(N)
End Sub

Function iMedian(ByVal v As Variant, Optional Sorted As Boolean = False) As Double
Dim N&, i&, j&, X#, Lwr&, Uppr&
Lwr = LBound(v): Uppr = UBound(v)
If Not Sorted Then QuickSort v
N = Uppr - Lwr + 1
If N Mod 2 = 0 Then
  i = Lwr + N \ 2
  iMedian = (v(i) + v(i - 1)) / 2
Else
  i = Lwr + N \ 2
  iMedian = v(i)
End If
End Function

Sub GaussCumProb(ByVal N&, Optional WithCum As Boolean = True, _
  Optional ToAttach As Boolean = False, Optional Xvals, Optional Xsigs, _
  Optional BoxClr, Optional ZeroXmin As Boolean = False, _
  Optional ChartAsSheet As Boolean = False, Optional Mode#, _
  Optional XvalColNum% = 0, Optional XsigsColNum% = 0, _
  Optional NumBins, Optional StartBin, Optional EndBin, Optional BinSpan)
  ' Construct cumulative-probability curve & histogram
  ' WithCum=>include cumulative gaussian curve; ToAttach=>to be immediately moved to the data-
  ' sheet as a picture;Ta contains the 1-D histogram-data vector if not supplied by InpDat();
  ' BoxClr specifies histogram-box colors.

Dim i&, j&, k&, M&, Ncells&, P#(), mm#, sc As Object
Dim X#, y#, Cellsize#, Extreme#, Xtik#, XvalMin#
Dim Incr#, s$, tC&, SerClr&, Ncurves%
Dim MaxBin%, b#(5, 2), H&(), ti&, MinBin, tL, Ad As Object
Dim Tx#(), ts#(), MinB#, MaxB#, BinScale#
Dim hR() As Range, Hcol%, WithHisto As Boolean, Cw, Sqrt2pi, Trans!
Dim Sigma#, Mu#, Rr As Range, Gp As Object, ArOld, ArNew, CumDat As Worksheet
Dim Area#, Ymax2#, OldArea#, HasHisto As Boolean, bRat!, StRange As Range, vStRange As Variant
Dim Dummy As Range, CumSer%, HaveDummy As Boolean, Nax%, cArr, FirstCurveSer%
Dim Ytik#, tYmax#, ConcInterMC As Boolean, Ydelt%, DummyCt%
Dim rw%, cc%, ii%

cArr = Array(13893632, 395485, 1357599, 390140, 8653042, 15379200, 144, 1139712, 9437184, 3830160, 10813510)
Const Tiny = 0.000001
ViM WithCum, True
ViM ToAttach, True
ViM ZeroXmin, False
ViM ChartAsSheet, False
ConcInterMC = (ToAttach And ConcPlot And DoMC)
SymbRow = Max(1, SymbRow)
StatBar "Constructing chart"
ReDim Tx(N), ts(N)

If ToAttach Then
  Set CumDat = Ash: Set ChrtDat = CumDat
Else
  Set CumDat = ChrtDat
End If

If IM(Xvals) Then
  j = IIf(ndCols < 3, 1, 3)
  For i = 1 To N: Tx(i) = InpDat(i, j): Next i

Else

  For i = 1 To N
    If XvalColNum > 0 Then
      Tx(i) = Xvals(i, XvalColNum)
    Else
      Tx(i) = Xvals(i)
    End If

  Next i

End If

If (CumGauss Or DoMix) And Not ConcPlot And Not DoMC Then

  If IM(Xsigs) Then
    j = IIf(ndCols < 4, 2, 4)
    For i = 1 To N: ts(i) = InpDat(i, j): Next i
  Else

    For i = 1 To N

      If XsigsColNum > 0 Then
        ts(i) = Xsigs(i, XsigsColNum)
      Else
        ts(i) = Xsigs(i)
      End If

    Next i

  End If

End If

Ncells = 2000 'If smaller (say 1000) get odd artefacts in cumgauss curve because of too-few points defining
MinX = 1E+32: MaxX = -1E+32
If NIM(BinSpan) Then BinWidth = BinSpan
WithHisto = (Nbins > 0 Or BinWidth > 0)
ReDim P(Ncells, 2)
XvalMin = 1E+32

For i = 1 To N: XvalMin = Min(XvalMin, Tx(i)): Next

If WithCum Then

  For i = 1 To N
    X = Tx(i)
    Extreme = 3 * ts(i)  ' Min & Max range of plot is +-3-sigma
    MinX = Min(MinX, X - Extreme) '  from range of data-pts.
    If MinX <= 0 And XvalMin >= 0 Then MinX = 0
    MaxX = Max(MaxX, X + Extreme) '  (actually less 1 cell-size at max)
  Next i

Else

  For i = 1 To N
    MinX = Min(MinX, Tx(i))
    MaxX = Max(MaxX, Tx(i))
  Next i

End If

GoSub RoundXaxis

If WithCum Then
  Sqrt2pi = Sqr(TwoPi)

  For i = 1 To Ncells
    X = MinX + (i - 1) * Cellsize
    P(i, 1) = X

    For j = 1 To N
      Mu = Tx(j)
      Sigma = ts(j)
      Incr = Exp(-SQ((X - Mu) / Sigma) / 2) / (Sigma * Sqrt2pi)
      If Incr < Tiny Then Incr = 0
      ' Eliminating the above line causes erratic assignment of very large values
      '  to p(i,2) as soon as the Incr = ... line is executed!  Bug in Excel VBA.
      P(i, 2) = P(i, 2) + Incr
    Next j

    'Ymax = Max(Ymax, p(i, 2))
  Next i

  Area = AreaUnderCurve(P(), Ncells)

  For i = 1 To Ncells: P(i, 2) = P(i, 2) / Area: Next i

  If AddToPlot Then
    SymbCol = CumDat.Cells(4, 2)  ' 1st empty column
    If SymbCol > 250 Then SymbCol = 5: SymbRow = SymbRow + 200
    Set Rr = sR(SymbRow, SymbCol, Ncells - 1 + SymbRow, 1 + SymbCol, CumDat) ' new CumGauss range
    AddSymbCol 2
  Else
    Set Rr = sR(1, 3, Ncells, 4, CumDat)
    Rr.Name = "gauss"
    LineInd Rr
    SymbCol = 5: BinScale = 1
  End If

End If

If AddToPlot Then
  j = 0: k = 0

  For j = 1 To 20 ' Determine if original plot contained histogram boxes
    If Cells(6, j) = "ErrBox" Then k = j: Exit For
  Next j

  Set IsoChrt = Sheets(Cells(2, 2).Text)

  If k > 0 Then
    WithHisto = True
    BinWidth = Cells(2, k) - Cells(1, k)
    BinStart = Cells(1, 5) ' Match bin width & location to original plot
    MinX = Min(Axxis(1).MinimumScale, P(1, 1))
    MaxX = Max(Axxis(1).MaximumScale, P(Ncells, 1))
    GoSub RoundXaxis
    Set IsoChrt = Sheets(PlotName$)
    With Axxis(1)
      .MinimumScale = MinX
      .MaximumScale = MaxX
      .CrossesAt = xlAutomatic
      .MajorUnitIsAuto = True
      .MinorUnitIsAuto = True
    End With
    BinSpec (MaxX - MinX) / BinWidth
  End If

End If

If WithHisto Then   ' Add histogram
  HasHisto = True

  If Nbins > 0 And IM(StartBin) And Not AddToPlot Then  ' Specified as auto bin-width
    BinWidth = Xspred / Nbins
    MinBin = MinX

  ElseIf NIM(NumBins) Then   ' Bin width & bin start specified
    MinBin = StartBin
    BinWidth = BinSpan
    Nbins = NumBins
  Else

    MinBin = BinStart '  find where first bin should start.
    i = Sgn(0.5 + (MinBin > MinX))

    Do While (i * MinBin) < (i * MinX)
      MinBin = MinBin + i * BinWidth
    Loop

    BinSpec (MaxX - MinX) / BinWidth
  End If

  If Nbins = 0 Then MsgBox "Invalid bin specification", , Iso: ExitIsoplot
  ReDim H(Nbins), hR(Nbins)
  SymbCol = Max(1, SymbCol): Hcol = SymbCol: SymbRow = Max(1, SymbRow)

  For i = 1 To Nbins
    'If i Mod 10 = 0 Then StatBar "Calculating" & Str(Nbins - 1 + 1)
    MinB = Drnd(MinBin + (i - 1) * BinWidth, 9)
    MaxB = Drnd(MinB + BinWidth, 9)

    For j = 1 To N
      X = Drnd(Tx(j), 9)

      If X >= MinB And X < MaxB Then
        H(i) = 1 + H(i)
        If H(i) > MaxBin Then MaxBin = H(i)
      End If

    Next j

  Next i

  If ToAttach Then    ' Don't try to include bin-heights too small
    mm = MaxBin / Hun '  to see at upper/lower extremes of data.

    For i = 1 To Nbins
      X = MinBin + i * BinWidth

      If H(i) > mm Then
        MinX = X - BinWidth
        Exit For
      End If

    Next i

    For i = Nbins To 1 Step -1
      X = MinBin + i * BinWidth

      If H(i) > mm Then
        MaxX = X
        Exit For
      End If

    Next i

  End If

  If Not AddToPlot Then GoSub RoundXaxis
  CumDat.Activate
  ti = 0

  For i = 1 To Nbins
    b(2, 2) = 0: b(5, 2) = 0: b(1, 2) = 0

    If H(i) > 0 Then
      j = ti * 6 + SymbRow
      k = j + 3 + SymbRow
      ti = 1 + ti
      b(1, 1) = MinBin + (i - 1) * BinWidth
      b(2, 1) = b(1, 1) + BinWidth
      b(3, 1) = b(2, 1): b(4, 1) = b(1, 1): b(5, 1) = b(1, 1)
      b(3, 2) = H(i):    b(4, 2) = b(3, 2)

      If AddToPlot And HistoStacked Then
        Ydelt = 0: cc = SymbCol - 2

        Do

          For ii = cc To 5 Step -1 ' find next histo-cell column-pair down
            If Cells(6, ii) = "ErrBox" Then Exit For
          Next ii

          rw = 1 ' Find histo-cell that matches the new data cell if any

          Do While Not IsEmpty(Cells(rw, ii - 1))

            If CSng(Cells(rw, ii - 1)) = CSng(b(1, 1)) Then
              Ydelt = Ydelt + Cells(rw + 2, ii) - Cells(rw, ii)
              Exit Do ' Y-increment  for stacking
            End If

            rw = rw + 6
          Loop

          cc = ii - 2
        Loop Until cc < 5

        ' Add the y-increment
        For ii = 1 To 5: b(ii, 2) = b(ii, 2) + Ydelt: Next ii
        MaxBin = Max(MaxBin, b(4, 2) + 1)
      End If

      Set hR(ti) = sR(j, Hcol, k, 1 + Hcol, CumDat)
      hR(ti).Value = b
      LineInd hR(ti), "ErrBox"
    End If

  Next i

  Nbins = ti:  AddSymbCol 2

  If Not WithCum Then
    P(1, 2) = MaxBin * 1.05: P(2, 2) = P(1, 2)
  End If

End If

If Not AddToPlot And WithHisto Then
  ChrtDat.Activate
  Set Dummy = sR(1, SymbCol, 2, 1 + SymbCol, ChrtDat)
  HaveDummy = True
  AddSymbCol 2
End If

If WithCum Then Rr.Value = P()
StatBar "Starting chart"

If AddToPlot Then
  Nax = 1: Ncurves = 0: j = 0

  For Each sc In IsoChrt.SeriesCollection
    j = 1 + j
    If sc.AxisGroup = 2 Then Nax = 2

    If sc.Points.Count > 50 Then
      Ncurves = 1 + Ncurves
      If Ncurves = 1 Then FirstCurveSer = j
    End If

  Next sc

  HasHisto = (Nax = 2)

  If ColorPlot Then
    SerClr = cArr(Ncurves + 1)
    SerClr = IIf(SerClr >= 0, SerClr, vbBlack)
  Else
    SerClr = vbBlack
  End If

  IsoChrt.Select
  CumDat.Visible = False

  If WithCum Then
    With IsoChrt
      .SeriesCollection.Add Rr, xlColumns, False, 1, False
      Set sc = Last(.SeriesCollection)
      sc.AxisGroup = 1 ' Don't know why needed
      With sc
        .AxisGroup = Nax: .MarkerStyle = xlNone
        With .Border
          .Color = IIf(ColorPlot, SerClr, vbBlack)
          .Weight = IIf(ColorPlot, IIf(Mac, xlMedium, xlThick), xlMedium)
        End With
        .Smooth = True  ' Can yield artefacts at P=0.
      End With
      .SeriesCollection(FirstCurveSer).Border.Color = cArr(1)
      With .Axes(xlValue, 1 - HasHisto)
        .MaximumScaleIsAuto = True: .MajorUnitIsAuto = True: .MinorUnitIsAuto = True
      End With

      If Not WithHisto Then
        With .Axes(1)
          MinX = Min(.MinimumScale, Rr(1, 1))

          If .HasTitle Then
            If InStr(LCase(.AxisTitle.Text), "age") And MinX < 0 Then MinX = 0
          End If

          MaxX = Max(.MaximumScale, Rr(Ncells, 1))

          If MinX < .MinimumScale Or MaxX > .MaximumScale Then
             GoSub RoundXaxis
             .MinimumScale = MinX
             .MaximumScale = MaxX
             Tick MaxX - MinX, Ytik
            .MajorUnitIsAuto = True
            .MinorUnitIsAuto = True
          End If

        End With
      End If

    End With
  End If

  StatBar
End If

If Not AddToPlot Then
  Charts.Add
  MakeSheet "ProbDens", Gp
  Set IsoChrt = Ash
  Landscape

  If HaveDummy Then
    Dummy(1, 1) = P(1, 1): Dummy(2, 1) = P(Ncells, 1)
    Dummy(1, 2) = 0: Dummy(2, 2) = 0
    Set StRange = Dummy
  Else
    Set StRange = Rr
  End If

  Set vStRange = StRange

  IsoChrt.ChartWizard Source:=vStRange, Gallery:=xlXYScatter, Format:=6, _
                      PlotBy:=xlColumns, CategoryLabels:=1, SeriesLabels:=0, _
                      HasLegend:=False, Title:="", CategoryTitle:=AxX$, _
                      ValueTitle:=AxY$, ExtraTitle:=""
  With IsoChrt
    PlotName$ = .Name

    If HaveDummy Then
      DummyCt = .SeriesCollection.Count

      For i = 1 To DummyCt
        With .SeriesCollection(i)
          .MarkerStyle = xlNone
          .Border.LineStyle = xlLineStyleNone
        End With
      Next i

      If WithCum Then .SeriesCollection.Add Rr, xlColumns, False, 1, False
    End If

    If WithCum Then
      If HaveDummy Then Last(.SeriesCollection).AxisGroup = 2
      CumSer = .SeriesCollection.Count
    End If

    .SizeWithWindow = False

    If Not DoMix Or ChartAsSheet Then
      With .PlotArea: .Height = 375: .Top = 35: .Width = 500: .Left = 95: End With
    End If

    With .Axes(xlValue)

      If ConcInterMC Then
        .MinimumScale = 0: .MaximumScale = MaxBin * 1.1
        .TickLabelPosition = xlNone
        .MajorTickMark = xlNone: .MinorTickMark = xlNone
        .HasTitle = False
      Else
        FormatHistoProbAxis 2 + WithHisto, 1, ChartAsSheet, MaxBin, DoMix, AddToPlot
        If WithHisto And WithCum Then _
          FormatHistoProbAxis 2, 2, ChartAsSheet, MaxBin, DoMix, AddToPlot
      End If

    End With
    If Xspred / Xtik > 8 Then Xtik = 2 * Xtik
    With .Axes(xlCategory)
      .MajorTickMark = xlOutside:  .MinorTickMark = xlInside

      If ZeroXmin And (MinX / Xspred) > 2 Then
        MinX = 0: .MinimumScale = 0
      ElseIf WithCum And WithHisto Then
        MinX = P(1, 1)
      ElseIf Not WithCum Then
        MinX = MinX - Xtik: .MinimumScale = MinX
      End If

      MaxX = MaxX + Xtik: .MaximumScale = MaxX
      .MinimumScale = MinX
      .CrossesAt = .MinimumScale
      .MinorUnitIsAuto = True:     .MajorUnit = Xtik
      .Border.Weight = AxisLthick: .TickLabelPosition = xlNextToAxis
      With .TickLabels.Font
        .Name = Opt.AxisTikLabelFont

        If ConcInterMC Then
          .Size = 24
        ElseIf DoMix Then
          If ChartAsSheet Then .Size = 16 Else .Size = 22
        Else
          .Size = Opt.AxisTikLabelFontSize
        End If

        .Background = xlTransparent
      End With
      .TickLabels.NumberFormat = TickFor(MinX, MaxX, Xtik)
      .ScaleType = False:        .Crosses = xlAutomatic
      .ReversePlotOrder = False: .CrossesAt = .MinimumScale

      If .HasTitle And AxX$ <> "" Then
        With .AxisTitle.Characters.Font
          .Name = Opt.AxisNameFont
          .Size = IIf(ConcInterMC, 30, Opt.AxisNameFontSize)
          .Background = xlTransparent
          If ConcInterMC Then .Bold = False
        End With
        Superscript Phrase:=.AxisTitle
      End If

      .HasMajorGridlines = True
      With .MajorGridlines.Border
        .Color = Menus("cGray50"): .Weight = xlHairline
        .LineStyle = IIf(Nbins > 0, xlDot, xlContinuous)
      End With
     End With


    If WithCum Then
      With .SeriesCollection(CumSer)
        With .Border
          .Color = IIf(ColorPlot, vbRed, vbBlack)
          .Weight = IIf(ColorPlot Or DoMix, xlThick, xlMedium)
        End With
        .MarkerStyle = xlNone
        .Smooth = True  ' Can yield artefacts at P=0.
      End With
    End If

    With .PlotArea.Border: .Weight = xlMedium: .Color = vbBlack: End With
  End With
End If

With IsoChrt
  If WithHisto Then
    StatBar "Adding histogram"

    For i = 1 To Nbins

      If DoShape Then
        GetScale
        If ColorPlot Then

          If AddToPlot Then
            tC = SerClr: Trans = 0.75
          ElseIf IM(BoxClr) Then
            tC = IIf(DoMix, RGB(140, 240, 240), RGB(70, 210, 210))
            Trans = 0.5
          Else
            tC = BoxClr: Trans = 0
          End If

          tL = 0

        Else
          tC = RGB(192, 192, 192): tL = 0.25: Trans = 0.5
        End If

        AddShape "ErrBox", hR(i), tC, Black, False, 0, , , Trans, tL
        MaxBin = Max(MaxBin, hR(i)(3, 2))

      Else
        .SeriesCollection.Add hR(i), xlColumns, False, 1, False
        With Last(.SeriesCollection)
          .AxisGroup = 1: .MarkerStyle = xlNone: .Smooth = False
          With .Border
            .Weight = xlMedium

            If AddToPlot Then
              With IsoChrt.SeriesCollection

                For j = 1 To .Count
                  If .Item(j).Points.Count = Ncells Then Exit For
                Next j

                .Item(j).Border.Color = vbBlue
              End With
              .Color = SerClr

            Else
              .Color = IIf(ColorPlot, vbBlue, vbBlack)
            End If

          End With
        End With
      End If

    Next i

    If NIM(Mode) Then Mode = MaxBin

'     If HaveDummy Then

'      For i = 1 To DummyCt
'        .SeriesCollection(1).Delete
'      Next i

'    End If

    If AddToPlot And WithHisto Then
      tYmax = Min(1 + MaxBin, Int(1.2 * MaxBin))
      With Axxis(2)

        If tYmax > .MaximumScale Then
          .MaximumScale = tYmax
          Tick .MaximumScale, Ytik
          .MajorUnit = Max(1, IIf(.MaximumScale > 9, 2 * Ytik, Ytik))
          .MinorUnit = Max(1, .MajorUnit / 2)
          If .MinorUnit <> .MajorUnit Then .MinorTickMark = xlInside
        End If

      End With
    End If

  End If

  CumDat.Visible = False: .Activate

  If Not AddToPlot Then
    .PlotArea.Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.PlotboxClr, vbWhite))
    .ChartArea.Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.SheetClr, vbWhite))
  End If

End With

If AddToPlot Then
  RescaleOnlyShapes True
Else
  If Not ToAttach Then AddCopyButton
  RemoveHdrFtr Gp
End If

Exit Sub

RoundXaxis:
  Tick MaxX - MinX, Xtik  ' Must repeat rounding of MinX, MaxX
  MinX = Drnd(Int(Drnd(MinX / Xtik, 7)) * Xtik, 7)
  X = MinX

  Do
    X = Drnd(X + Xtik, 7)
  Loop Until X >= MaxX Or X = MinX

  MaxX = X
  Xspred = MaxX - MinX
  Cellsize = Xspred / Ncells
Return

End Sub

Sub RescaleOnlyShapes(Optional ExternalInvoked = True, Optional ChartSelected = False, Optional HiddenSheetName)
ViM ExternalInvoked, True
ViM ChartSelected, False
StatBar "rescaling"
GetOpSys
PutShapesBack False, ChartSelected, , HiddenSheetName
End Sub

Public Function AgeSingleDiscordantPt(RangeIn As Range, LowerInterceptAge, _
                Optional TeraWasserburg = False, Optional InputSigmaLevel = 1, _
                Optional OutputSigmaLevel = 1, Optional PercentErrsIn = True, _
                Optional WithLambdaErrs = False)
' 09/06/19 -- added.
' Input is a range of 5 values ((conventional concplot): ConcX,err,ConcY,err,rho,
'  Assumed lower intercdept age (ma), sigma level of input errors, desired sigma
'  level of output error, AreInputErrorsInPercent, with decay const errors.
' Output is a 1-row by 2-column range containing the upper-intercept age of a chord
'  through the x-y coordinates (first 5 parameters) and anchored by the lower-intercept
'  age on Concordia.
Dim X#, y#, DP() As DataPoints, N%, Bad As Boolean, i%, r(2, 5) As Variant, ConcInt#(6)
Dim AgeAndError(2), aBad(2) As Boolean, t1err#(2), t2err#(2), Rr(2, 5) As Variant
Dim Vr
AgeAndError(1) = "#NUM!"
AgeAndError(2) = "#NUM!"
On Error GoTo done
If TeraWasserburg Then
  Vr = ConvertConc(RangeIn, TeraWasserburg, PercentErrsIn)
  For i = 1 To 5
    r(1, i) = Vr(i)
  Next i
Else
  For i = 1 To 5
    r(1, i) = RangeIn(i)
  Next i
End If
r(2, 1) = ConcX(LowerInterceptAge): r(2, 3) = ConcY(LowerInterceptAge)
r(2, 2) = 0.001:                    r(2, 4) = 0.001
If Not PercentErrsIn Then
  For i = 2 To 4 Step 2
    r(1, i) = 0.00001 * r(1, i - 1)
  Next i
End If
GetDat r, DP, N, InputSigmaLevel, PercentErrsIn, Bad
If Bad Then GoTo done
Normal = True
ConcordiaIntercepts yf.Slope, yf.Intercept, ConcInt, , True
If ConcInt(1) >= 0 And ConcInt(2) > 0 Then
  If CSng(ConcInt(2) / IIf(ConcInt(1) = 0, -1, ConcInt(1))) <> 1 Then
    AgeAndError(1) = ConcInt(2)
    ConcIntAgeErrors ByVal LowerInterceptAge, AgeAndError(1), t1err(), t2err(), aBad
    If Not aBad(2) Then
      AgeAndError(2) = t2err(1 - WithLambdaErrs) * OutputSigmaLevel
    End If
  End If
End If
done:
AgeSingleDiscordantPt = AgeAndError
End Function
