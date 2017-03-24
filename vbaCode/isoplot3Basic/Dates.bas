Attribute VB_Name = "Dates"
Option Private Module
Option Explicit: Option Base 1
Const SqrTwoPi = 2.506628274631, PiOffs = 0.0001, Brk = "Ages of stacked, dated beds"
Dim aa#(), VarA#(), SigmaA#(), Ndat%, Ncomp%
Dim Toffs#, BadMix As Boolean, SigmaRho#(), MixConstr As Boolean

Sub Mix(Optional FromMenu = False)
Attribute Mix.VB_ProcData.VB_Invoke_Func = "m\n14"
' Implementation of Sambridge's algorithm for deconvoluting superimposed Gaussian distributions
Dim i%, j%, k%, nU%, jj%, kk%, oSL%
Dim iSL%, cProp%, Iter%, Xcalc0&, Nfm As Boolean
Dim Percent As Boolean, TestOK As Boolean, OK As Boolean, Misfit!, v!
Dim pProp!, Sp#, ms$, NoSoln As Boolean, T#(), pi#(), tmp$, Guess As Boolean
Dim d As Object, Eb As Object, cb As Object, Op As Object, La As Object, tB As Object, Bu As Object, spi As Object
Dim eb1 As Object, eb2 As Object, eP1 As Object, ep2 As Object, SigmaT#(), SigmaPi#(), ChartAsSheet As Boolean
Dim s As Range, js As String * 1, ks As String * 1, DS$, pS$, Calculated As Boolean, rT$, re As String * 34
Dim Misfit1!, Ncomp0%, t0#(1), Pi0#(1), Sh As Object, Idat#()
Dim Lower#, Upper#, Spred#, Incr#, Mx$, done As Boolean, PicName$, rn$
Dim pL!, pT!, pH!, p_L!, p_T!, pW!, G As Object
Dim pc2 As Boolean, iSL2%, ChartWithData As Boolean, NumbersWithData As Boolean, ErrRhoWithData As Boolean
Dim rc$()
ViM FromMenu, False
Nfm = Not FromMenu
If FromMenu Then
  CheckInpRange True, Idat()
Else
  Xcalc0 = Qcalc
  On Error GoTo BadRange
  Set s = Selection
  On Error GoTo 0
  Set DatSht = Ash
  DatSheet$ = DatSht.Name
  CheckInpRange False, Idat(), s
  If s.Columns.Count <> 2 Then
    MsgBox "Ages and errors must be in adjacent columns" & _
      viv$ & "(unmixing of overlapping ages)", , Iso
    KwikEnd
  End If
  Ndat = Min(LastOccupiedRow(s.Column), s.Rows.Count)
  GetConsts
  AssignIsoVars
  k = 0
  For i = 1 To Ndat
    If InpDat(i, 3) > 0 Then
      k = 1 + k
      Idat(k, 1) = InpDat(i, 1)
      Idat(k, 2) = InpDat(i, 3)
    End If
  Next i
  For i = 1 To k
    InpDat(i, 1) = Idat(i, 1)
    InpDat(i, 3) = Idat(i, 2)
  Next i
End If
Ndat = N
If Ndat < 5 Then MsgBox "Need at least 6 data points", , Mx$: KwikEnd
ms$ = "Trial proportions must total 1"
AssignD "Mix", d, Eb, cb, Op, La, G, tB, Bu, spi
Bu("bGo").Text = "Calculate"
NoSoln = True
Fit tB, d, NoSoln
ClearInput Eb
ClearOutput Eb, La, T(), pi()
ShowMix_click
Guess_click
For i = 1 To Eb.Count: Eb(i).Enabled = True: Next
For i = 1 To Op.Count: Op(i).Enabled = True: Next
For i = 1 To G.Count: G(i).Enabled = True: Next
For i = 1 To tB.Count: tB(i).Enabled = True: Next
If FromMenu Then
  Op("oI" & tSt(SigLev) & "sigma") = xlOn
  If AbsErrs Then Op("oAbs") = xlOn Else Op("oPercent") = xlOn
End If
Op("oAbs").Enabled = Nfm: Op("oPercent").Enabled = Nfm
For i = 1 To 2: Op("oI" & tSt(i) & "sigma").Enabled = Nfm: Next i
DoMix = True
spi("sGuess").Value = 2: Eb("eGuess").Text = "2"

First:
Do
Again:
  Do
    If Not Calculated Then Ncomp = 2
    Bu("bExit").Enabled = Calculated
    Percent = IsOn(Op("oPercent")): iSL = 2 + IsOn(Op("oI1sigma"))
    If Calculated Then
      Do
        Canceled = False: MixConstr = False
        done = DialogShow(d)
        If Canceled Then ExitIsoplot
      If Not AskInfo Then Exit Do
        Caveat_Mix
      Loop
      pc2 = IsOn(Op("oPercent")): iSL2 = 2 + IsOn(Op("oI1sigma"))
      If pc2 <> Percent Or iSL2 <> iSL Then Calculated = False: GoTo First
      Bu("bGo").Text = "Calculate Again"
    End If
    If Calculated And MixConstr Then
      ChartAsSheet = IsOn(cb("cChartAsSheet"))
      ChartWithData = IsOn(cb("cChartWithData"))
      NumbersWithData = IsOn(cb("cShowWithData"))
      ErrRhoWithData = IsOn(cb("cShowMatrix"))
      AttachCumGauss PicName$, (FromMenu), s, T(), pi(), NoSoln, ChartWithData, ChartAsSheet
      If Calculated And (NumbersWithData Or ErrRhoWithData) And Not NoSoln Then
        If ChartAsSheet Then DatSht.Select
        js = tSt(oSL)
        rT$ = " Age      " & pm & js & "sigma    fraction    " & pm & js & "sigma"
        For j = 1 To Ncomp
          js = tSt(j)
          re = La("lA" & js).Text
          Mid$(re, 10, 6) = La("lAsig" & js).Text
          Mid$(re, 20, 5) = La("lP" & js).Text
          Mid$(re, 29, 5) = La("lPsig" & js).Text
          rT$ = rT$ & vbLf & re
        Next j
        If Misfit1 <> 0 Then rT$ = rT$ & vbLf & "relative misfit = " & tB("tMisfit").Text
        If FromMenu Then Set s = Range(Irange$) Else Irange$ = s.Address
        If ChartWithData Then pL = Right_(Last(Ash.Shapes)) Else pL = Right_(s)
        AddResBox rT$, , , LightTurquoise, pL, , True, , , s.Top, , rn$
        With Ash.TextBoxes(rn$): pL = .Left: pT = .Top: pH = .Height: pW = .Width: End With
        p_L = pL: p_T = pT + pH + 2
        With cb("cShowMatrix")
          If .Enabled And .Value = xlOn Then
            nU = 2 * Ncomp - 1
            rT$ = Space((2 + 8 * nU - 16) / 2) & "Sigma-Rho Matrix" & vbLf & "      "
            For j = 1 To Ncomp - 1
              rT$ = rT$ & "f" & tSt(j) & "      "
            Next j
            For j = Ncomp To nU
              rT$ = rT$ & "t" & tSt(j - Ncomp + 1)
              If j < nU Then rT$ = rT$ & "      "
            Next j
            ReDim rc$(nU)
            For i = 1 To nU
              If i < Ncomp Then
                rc$(i) = "f": j = i
              Else
                rc$(i) = "t": j = i - Ncomp + 1
              End If
              rc$(i) = rc$(i) & tSt(j)
              rc$(i) = rc$(i) & Space(8 * nU)
              For j = 1 To nU
                v = SigmaRho(i, j)
                If i <> j Or i < Ncomp Then tmp$ = App.Fixed(v, 3) Else tmp$ = tSt(Drnd(v, 3))
                kk = (Left$(tmp$, 1) = "-")
                Mid$(rc$(i), 5 + (j - 1) * 8 + kk, 6) = tmp$
              Next j
              rT$ = rT$ & vbLf & rc$(i)
            Next i
            Set Sh = Last(Ash.Shapes)
            AddResBox rT$, , , LightGreen, Sh.Left, , True, , True, 6 + Bottom(Sh), , rn$
            With Ash.TextBoxes(rn$)
              p_L = .Left
              p_T = .Top + .Height + 1
            End With
          End If
        End With
      End If
      If PicName$ <> "" And pL > 0 Then
        With Ash.Shapes(PicName$): .Left = p_L: .Top = p_T: End With
      End If
      ExcelCalc = Xcalc0
      If ChartAsSheet Then IsoChrt.Select
      If Not FromMenu Then KwikEnd
      Exit Sub
    End If
    Guess = (IsOn(Op("oIsoGuess")) Or Not Calculated)
    If Not Guess Then
      Do ' Make the trial entries contiguous
        OK = True: kk = 0: cProp = 0
        For j = 1 To 5
          js = tSt(j): ks = tSt(j + 1)
          Set eb1 = Eb("eT" & js): Set eb2 = Eb("eT" & ks)
          Set eP1 = Eb("eP" & js): Set ep2 = Eb("eP" & ks)
          DS$ = Eb("eT" & js).Text: pS$ = Eb("eP" & js).Text
          If eb1.Text = "" And eb2.Text <> "" Then
            eb1.Text = eb2.Text: eb2.Text = ""
            eP1.Text = ep2.Text: ep2.Text = ""
            OK = False
          End If
        Next j
      Loop Until OK
      Sp = 0: Ncomp = 0
      ReDim T(6), pi(6)
      For j = 1 To 6
        js = tSt(j)
        DS$ = Eb("eT" & js).Text: pS$ = Eb("eP" & js).Text
        If DS$ <> "" Then
          Ncomp = 1 + Ncomp:  OK = False
          T(Ncomp) = Val(DS$): pi(Ncomp) = Val(pS$)
          kk = kk - (pi(Ncomp) > 0): Sp = Sp + pi(Ncomp)
          If pi(Ncomp) <= 0 Then
            MsgBox "Must enter relative proportions for all components", , Mx$
          Else
            OK = True
          End If
          If Not OK Then Exit For
        End If
      Next j
      ReDim Preserve T(Ncomp), pi(Ncomp)
    End If
  Loop Until Guess Or (OK And Ncomp > 0 And kk >= Ncomp)
  NoUp
  Percent = IsOn(Op("oPercent"))
  oSL = 2 + IsOn(Op("oO1sigma")): iSL = 2 + IsOn(Op("oI1sigma"))
  ReDim SigmaA(Ndat), aa(Ndat)
  For i = 1 To Ndat ' Initialize age & Sigma arrays of input data
    aa(i) = Idat(i, 1): SigmaA(i) = Idat(i, 2) / iSL
    If Percent Then SigmaA(i) = SigmaA(i) / Hun * aa(i)
  Next i
  If Guess Then
    Randomize Timer
    Lower = App.Min(aa): Upper = App.Max(aa): Spred = Upper - Lower
    If Calculated Then Ncomp = EdBoxVal(Eb("eGuess")) Else Eb("eGuess").Text = "2"
    ReDim T(Ncomp), pi(Ncomp)
    Lower = Lower + Spred / 8 / Ncomp:    Upper = Upper - Spred / 8 / Ncomp
    For j = 1 To Ncomp ' More or less evenly distributed T across the observed values
      T(j) = Lower + j * Spred / (Ncomp + 1) + (0.5 - Rnd) * Spred / 4 / Ncomp
      pi(j) = 0
    Next j
    i = 0: Sp = 0
    Do ' Create roughly equal Pi()
      Do
        j = 1 + Rnd * (Ncomp - 2)
      Loop Until pi(j) = 0
      i = i + 1
      pi(j) = (1 - Sp) / (Ncomp + 1 - i) + (0.5 - Rnd) / Ncomp
      Sp = Sp + pi(j)
    Loop Until i = Ncomp - 1
    For i = 1 To 6
      js = tSt(i)
      If i > Ncomp Then
        Eb("eT" & js).Text = ""
        Eb("eP" & js).Text = ""
      Else
        If pi(i) = 0 Then pi(i) = 1 - Sp
        Eb("eT" & js).Text = tSt(Drnd(T(i), 4))
        Eb("eP" & js).Text = tSt(Drnd(pi(i), 3))
      End If
    Next i
    Sp = 1
  End If
  For j = 1 To Ncomp: pi(j) = Prnd(pi(j) / Sp, -4): Next j
  ClearOutput Eb, La, T(), pi()
  Ncomp0 = Ncomp: Ncomp = 1
  t0(1) = iAverage(aa): Pi0(1) = 1
  UnMix t0(), SigmaT(), Pi0(), SigmaPi(), Misfit1, NoSoln
  If Calculated Or Not NoSoln Then
    If NoSoln Then
      Misfit1 = -1
      MsgBox "Couldn't calculate 1-component misfit", , Mx$
    End If
    Ncomp = Ncomp0
    UnMix T(), SigmaT(), pi(), SigmaPi(), Misfit, NoSoln
    If Calculated Or Not NoSoln Then
      If NoSoln Then
        MsgBox "Unable to find a solution", , Mx$
      Else
        Misfit = Misfit / Misfit1
        Calculated = True
        With App
          For j = 1 To Ncomp
            js = Trim(Str(j))
            La("lA" & js).Text = Ernd(T(j), SigmaT(j), True)
            La("lAsig" & js).Text = Drnd(oSL * SigmaT(j), 2)
            La("lP" & js).Text = .Fixed(pi(j), 2)
            If j < Ncomp Then
              La("lPsig" & js).Text = .Fixed(oSL * SigmaPi(j), 2)
            Else
              La("lPsig" & js).Text = " ---"
            End If
          Next j
        End With
        tB("tCalcAgeSigma").Text = pm & Trim(Str(oSL)) & "sigma"
        tB("tCalcPropSigma").Text = tB("tCalcAgeSigma").Text
        Fit tB, d, NoSoln
        tB("tMisfit").Text = IIf(Misfit1 = 0 Or Misfit = 0, "", Format(Misfit, "0.000"))
        If cProp > 0 Then
          Eb("eP" & tSt(cProp)).Text = tSt(pProp)
        End If
      End If
    End If
  End If
  Calculated = True
  If NoSoln Then
    Fit tB, d, NoSoln
    ClearInput Eb
    ClearOutput Eb, La, T(), pi()
  End If
Loop
NoSoln:   MsgBox "Unable to find a solution", , Mx$: GoTo Again
BadRange: MsgBox "Invalid selected range", , Iso
ExitIsoplot
End Sub

Sub ClearInput(Eb As Object)
Attribute ClearInput.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j As String * 1
For i = 1 To 6
  j = tSt(i)
  Eb("eT" & j).Text = "": Eb("eP" & j).Text = ""
Next i
End Sub

Sub ClearOutput(Eb As Object, La As Object, T#(), pi#())
Attribute ClearOutput.VB_ProcData.VB_Invoke_Func = " \n14"
' Clear the output section of the Mix dialog box
Dim j%, js As String * 1, ts$, pS$, eT As Object, eT1 As Object, eP As Object, eP1 As Object
If Ncomp > 0 Then ReDim Preserve pi(Ncomp), T(Ncomp)
For j = 1 To 6 ' Clear the output section of the dialog box
  js = tSt(j)
  La("lA" & js).Text = "": La("lAsig" & js).Text = "": La("lP" & js).Text = ""
  If j < 6 Then La("lPsig" & js).Text = ""
  ts$ = "": pS$ = ""
  Set eT = Eb("eT" & js): Set eP = Eb("ep" & js)
  If j = 1 Then Set eT1 = eT: Set eP1 = eP
  If j <= Ncomp Then 'And Not NoSoln Then
    If pi(j) > 0 Then
      ts$ = tSt(Drnd(T(j), 4)): pS$ = tSt(Drnd(pi(j), 3))
    End If
  End If
  'eT.Left = eT1.Left: eT.Width = eT1.Width: eT.Height = eT1.Height
  'eT.Top = eT1.Top + (j - 1) * (eT1.Height + 2)
  'eP.Height = eT1.Height: eP.Width = eT1.Width - 2
  'eP.Left = eT1.Left + eT1.Width + 4
  Eb("eT" & js).Text = ts$: Eb("eP" & js).Text = pS$
Next j
End Sub

Sub Fit(tB As Object, d As Object, NoSoln As Boolean)
Attribute Fit.VB_ProcData.VB_Invoke_Func = " \n14"
Dim OK As Boolean
OK = Not NoSoln
tB("tMisfitCaption").Visible = OK: tB("tMisfit").Visible = OK
d.GroupBoxes("gMisfit").Visible = OK
End Sub

Sub ff(pi#(), T#(), f#())
Attribute ff.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%
For j = 1 To Ncomp
  For i = 1 To Ndat
    f(i, j) = 1 / SigmaA(i) / SqrTwoPi * Exp(-SQ(aa(i) - T(j)) / (2 * VarA(i)))
Next i, j
End Sub

Function LL2(ByVal jj%, ByVal kk%, ByVal PiJoffs%, ByVal PiKoffs%, _
  ByVal tJoffs%, ByVal tKoffs%, pi#(), T#())
Attribute LL2.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, OffsT#, OffsP#
Dim PiA#(), Ta#(), FA#()
' Elements of inverse variance-covariance matrix (second derivative of Ln(L))
' jj, kk are indices of Pi or T components to diddle;
' PiJoffs, PiKoffs are signs (+1, 0, -1) of Pi offset;
' tJoffs, tKoffs are signs (+1, 0, -1) of T offset.
ReDim PiA(Ncomp), Ta(Ncomp), FA(Ndat, Ncomp)
For i = 1 To Ncomp
  PiA(i) = pi(i): Ta(i) = T(i)
  If i = jj Then
    If PiJoffs Then PiA(i) = PiA(i) + PiJoffs * PiOffs
    If tJoffs Then Ta(i) = Ta(i) + tJoffs * Toffs
  End If
  If i = kk Then
    If PiKoffs And (Not PiJoffs Or jj <> kk) Then PiA(i) = PiA(i) + PiKoffs * PiOffs
    If tKoffs And (Not tJoffs Or jj <> kk) Then Ta(i) = Ta(i) + tKoffs * Toffs
  End If
Next i
ff PiA(), Ta(), FA()
LL2 = LL(PiA(), FA())
End Function

Function LL(pi#(), f#())
Attribute LL.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, L#
LL = 0: BadMix = False
For i = 1 To Ndat
  L = 0
  For j = 1 To Ncomp
    L = L + pi(j) * f(i, j)
  Next j
  If L <= 0 Then BadMix = True: Exit Function
  LL = LL + Log(L)
Next i
End Function

Sub Bracket(ByVal Nbeds&, ByVal Ntrials&, Nbins%, Age#(), _
  SigmaA#(), T#(), Upper#(), Lower#(), _
  Delta#(), DeltaPlus#(), DeltaMinus#(), ByVal Clev!, _
  InpRange As Range, TempSht As Worksheet, Optional Useries As Boolean = False)
' Input is a 2-column range of ages of stratigraphically ordered units (oldest at bottom, youngest at top)
'  with 1-SigmaA absolute errors in the column to the right of the ages.
' Output are in the 3 adjacent columns, being, for each unit, the lower limit on the unit's age (@95%
'  conf.), the mode of the age, and the upper 95%-conf limit.
Dim i&, j%, k%, MaxTry&, ct&, LowerInd&
Dim UpperInd&, OK As Boolean, Abeds%
Dim BinWidth#, UnSucc$, ThU#, Gamma#, BinLim#()
Dim v#, u#(), uT#
Dim tB#(), Delt#(), b#(), Bins%()
NoUp
Abeds = Nbeds - (Nbeds < 8)
ReDim tB(Ntrials, Nbeds), Delt(Ntrials, 2 To Abeds)
Randomize Timer
ViM Useries, False
MaxTry = 50000: Nbins = 20
UnSucc$ = "No successful trials in" & Str(MaxTry) & " attempts"
LowerInd = (1 - Clev) / 2 * Ntrials: UpperInd = Ntrials - LowerInd
If Useries Then
  ReDim u(Nbeds, 5)
  For i = 1 To Nbeds
    For j = 1 To 5
      If j < 5 Or ndCols > 4 Then
        u(i, j) = InpRange(i, j)
      Else
        u(i, j) = 0
      End If
      If j = 2 Or j = 4 Then
        u(i, j) = u(i, j) / SigLev
        If Not AbsErrs Then u(i, j) = u(i, j) / Hun * u(i, j - 1)
      End If
  Next j, i
End If
For i = 1 To Ntrials
  Do
    OK = True
    For j = 1 To Nbeds
      If Useries Then
        GaussCorrel u(j, 1), u(j, 2), u(j, 3), u(j, 4), u(j, 5), ThU, Gamma
        ThUage ThU, Gamma, uT
        tB(i, j) = uT / Thou
      Else
        tB(i, j) = Gaussian(Age(j), SigmaA(j))
      End If
      If tB(i, j) < 0 Or Useries And tB(i, j) = 0 Then OK = False: Exit For
      If j > 1 Then
        If tB(i, j) < tB(i, j - 1) Then OK = False: Exit For
      End If
    Next j
    ct = 1 + ct
    If ct = MaxTry Then
      If i < 2 Then MsgBox UnSucc$, , Brk: KwikEnd
    End If
  Loop Until OK
  If i < Hun Or i Mod Hun = 0 Then StatBar Str(Ntrials - i) & Space(6) & _
      "fract. successful trials  " & Str(Drnd(i / ct, 2)) & "   "
Next i
StatBar "Calculating"
For i = 1 To Ntrials
  For j = 2 To Abeds
    If j > Nbeds Then
      Delt(i, j) = tB(i, Nbeds) - tB(i, 1)
    Else
      Delt(i, j) = tB(i, j) - tB(i, j - 1)
    End If
Next j, i
Sheets.Add
Set TempSht = Ash
For j = 1 To 2 * Nbeds + (Nbeds = 8)
  StatBar "Calculating" & Str((2 * Nbeds - j + 1) \ 2)
  If j > Nbeds Then
    SortCol Delt(), b(), Ntrials, j - Nbeds + 1
  Else
    SortCol tB(), b(), Ntrials, j
    Lower(j) = b(LowerInd)
    Upper(j) = b(UpperInd)
  End If
  BinWidth = (b(Ntrials) - b(1)) / Nbins
  Freq Ntrials, b(), Bins(), BinLim(), Nbins, b(1), BinWidth, v
  For i = 1 To Nbins
    Cells(i, 2 * j) = Bins(i)
    Cells(i, 2 * j - 1) = BinLim(i) + BinWidth / 2
  Next i
  If j > Nbeds Then
    k = j - Nbeds + 1
    Delta(k) = v
    DeltaMinus(k) = v - b(LowerInd)
    DeltaPlus(k) = b(UpperInd) - v
  Else
    T(j) = v
  End If
Next j
StatBar
End Sub

Sub SetupBracket(Optional FromMenu = False, Optional Useries As Boolean = True) ' User interface to Bracket
Dim i%, j%, c%, r As Range, Ob As Object, P As String * 1, cr As Range
Dim okB%, f%, d1 As String * 10, d2 As String * 8, d3 As String * 8, d4$, d$
Dim Nbins%, ba As Boolean, TempSht As Worksheet, Tx$, nU As Boolean, AvJ%, M%
Dim e As Object, o As Object, L As Object, b As Object, OK As Boolean, s As Object, Ta#, ts#
Dim AvEr, Ntrials&, OKn As Boolean, Age#(), SigmaA#(), Ch As Object, MaxB%
Dim ms1%, ms2%, ms3%, si!, tmp$, Li As Object, tB As Object
Dim nR%, nc%, Clev!, cl As String * 2, ShowPlot As Boolean, v!
Dim G As Object, tL As Object, mB As String * 1, Xcalc0&, Idat#(), Tbox$, Br As Object
Dim e2() As String * 8, e3() As String * 8, d5$, e4$(), cf$, DeltaGraf As Boolean, Bb As Boolean
ViM FromMenu, False
ViM Useries, True
MaxB = 8: mB = tSt(MaxB): tmp$ = "": nU = Not Useries
If Ash.Type <> xlWorksheet Then _
  MsgBox "This function can only be invoked from a Worksheet.", , Brk: Exit Sub
  Set DatSht = Ash: DatSheet$ = DatSht.Name
If Not FromMenu Then
  Xcalc0 = Qcalc
  On Error GoTo BadRange
  Set r = Selection
  Set cr = r.CurrentRegion
  On Error GoTo 0
  If r.Count = 1 And (cr.Columns.Count = 2 - 3 * Useries) And cr.Rows.Count > 1 Then Set r = cr: r.Select
  nR = r.Rows.Count: nc = r.Columns.Count: tmp$ = ""
  If Useries And (nR < 2 Or nc < 4 Or nc > 5) Then
    tmp$ = "Input range must contain at least 2 rows of 4 or 5 columns, e.g." & viv & _
      "    230Th/238U, err, 234U/238U, err [, err-correl]" & viv & _
      "Rows must be in order of increasing relative age."
  ElseIf Not Useries And (nR < 2 Or nc <> 2) Then
    tmp$ = "Input range must contain dates and date errors in 2 adjacent columns." _
      & viv$ & "Dates must follow actual stratigraphic order, highest stratigraphic levels at top."
  End If
  If tmp$ <> "" Then
    MsgBox tmp$, , Brk: KwikEnd
  End If
  GetOpSys
End If
AssignD "Bracket", Br, e, Ch, o, L, G, tB, b, s, , , , Li
GoSub Conf
For Each Ob In tB
  With Ob
    Select Case .Name
      Case "tTitle"
        f = 10 + 2 * Mac
        Tx$ = IIf(Useries, "U-series analyses", "dated beds")
        .Text = "Age limits on a stratigraphic"
        If MacExcelX Then .Text = .Text & "ally ordered"
        .Text = .Text & " sequence of " & Tx$
      Case "tStrat": f = 8 + 2 * Mac
      Case Else: f = 9 + 2 * Mac
    End Select
    .Font.Size = f
  End With
Next
If Not FromMenu Then N = nR: ndCols = r.Columns.Count
CheckInpRange (FromMenu), Idat(), r
If N < 2 Then
  MsgBox "Need ages/errors for at least 2 units", , Brk: Exit Sub
ElseIf N > MaxB Then
  MsgBox "Maximum # of dated units is" & Str(MaxB), , Brk: Exit Sub
End If
For i = 1 To MaxB ' Erase old output
  P = tSt(i)
  L("la" & P).Text = "": L("ua" & P).Text = "": L("ba" & P).Text = ""
  If i > 1 Then
    L("da" & P).Text = "": L("uda" & P).Text = "": L("bda" & P).Text = ""
  End If
  e("eT" & P).Text = "": e("ee" & P).Text = ""
Next i
tB("tDate").Text = IIf(Useries, "/238U", "Date")
tB("tNum").Text = IIf(Useries, "230Th", "")
tB("tError").Text = IIf(Useries, "/238U", "error")
tB("tSigLev").Text = IIf(Useries, "234U", "1 sigma")
tB("tBest").Text = IIf(Useries, "  Best age ka", "Best age")
Tx$ = IIf(Useries, "sample", "bed")
tB("tDelt").Text = "Difference from overlying " & Tx$
Set tL = L("da" & mB)
G("gTopBottom").Visible = False
For i = 1 To N ' Transfer input from selected range
  P = tSt(i)
  e("eT" & P).Text = tSt(Drnd(Idat(i, 1), 4))
  e("ee" & P).Text = tSt(Drnd(Idat(i, 2), 4))
Next i
b("bGo").Caption = "GO": b("Cancel").Visible = True
e("eTrials").Enabled = True: L("lTrials").Enabled = True: s("sTrials").Visible = True
If FromMenu Then
o("o1sigma") = IIf(SigLev = 1, xlOn, xlOff)
o("o2sigma") = IIf(SigLev = 1, xlOff, xlOn)
o("oAbs") = IIf(AbsErrs, xlOn, xlOff)
o("opercent") = IIf(AbsErrs, xlOff, xlOn)
Else
  SigLev = 1 - IsOn(o("o2sigma"))
  AbsErrs = IsOff(o("opercent"))
End If
Siglev_Click
ConfLev_click
BracketShowResClick
Do ' Show dialog box & check inputs
  For i = 1 To tB.Count: tB(i).Visible = True: Next
  If MacExcelX Then tB("tStrat").Visible = False
  With tB("tTopBottom")
    .Text = IIf(Useries, "Oldest minus youngest:", "Top-Bottom age difference:")
    If Not MacExcelX Then .Characters.Font.Bold = True 'Not Useries
    .Visible = False
  End With
  OK = True: OKn = True: okB = True
  GoSub Conf
  ba = IsOn(Ch("cShowRes")): Bb = IsOn(Ch("cShowChart"))
  Ch("cShowRes").Text = "Put results directly on worksheet"
  Ch("cShowChart").Enabled = ba
  o("oAges").Enabled = (ba And Bb): o("oDeltas").Enabled = (ba And Bb)
  Do
    If Not DialogShow(Br) Then Exit Sub
  If Not AskInfo Then Exit Do
    Caveat_Bracket
  Loop
  GoSub Conf
  SigLev = 1 - IsOn(o("o2sigma")): AbsErrs = IsOn(o("oAbs"))
  ShowPlot = IsOn(Ch("cShowChart"))
  If Useries Then
    If FromMenu Then Set r = Selection
    N = r.Rows.Count
    ReDim Age#(N), SigmaA#(N)
  Else
    N = 0
    ReDim Age#(MaxB), SigmaA#(MaxB)
    For i = 1 To MaxB ' Parse input ages/errors
      If Trim(e("eT" & tSt(i)).Text) <> "" Then
        N = N + 1
        ReDim Preserve Age#(N), SigmaA#(N)
        Age(N) = EdBoxVal(e("eT" & tSt(i)))
        SigmaA(N) = EdBoxVal(e("ee" & tSt(i))) / SigLev
        If Not AbsErrs Then SigmaA(N) = SigmaA(N) / Hun * Age(N)
        If Age(N) = 0 Or SigmaA(N) <= 0 Then OK = False
      End If
    Next i
    ReDim Preserve Age#(N), SigmaA#(N)
  End If
  Ntrials = Thou * Val(e("eTrials").Text)
  If Ntrials < 2000 Or Ntrials > 64000 Then OKn = False
  If N < 2 Then okB = False
  If Not okB Then
    MsgBox "Need at least 2 dated units", , Brk
  ElseIf Not OK Then
    MsgBox "Ages and errors must be nonzero", , Brk
  ElseIf Not OKn Then
    MsgBox "#trials must be between 2000 and 64000", , Brk
  End If
Loop Until OK And OKn And okB
If Not FromMenu Then ColorPlot = True ' In case invoked by toolbar button
ReDim T#(N), Upper#(N), Lower#(N)
' Invoke the calculation
M = N - (N < MaxB)
ReDim Delta#(2 To M), DeltaPlus#(2 To M)
ReDim DeltaMinus#(2 To M), e2(M), e3(M), e4$(M)
If FromMenu Then Set r = RangeIn(1)
If MacExcelX Then tB("tStrat").Visible = False
Bracket N, Ntrials, Nbins, Age(), SigmaA(), T(), Upper(), Lower(), _
  Delta(), DeltaPlus(), DeltaMinus(), Clev, r, TempSht, Useries
Ch("cShowChart").Enabled = True
f = N - (N < MaxB): AvJ = 0
On Error Resume Next
For i = 1 To f ' Format the output
  P = tSt(IIf(i > N, MaxB, i))
  If i <= N Then
    AvEr = (Upper(i) - Lower(i)) / 2
    j = Int(-Log(AvEr) / Log(10) + 2)
    AvJ = j + AvJ
  Else
    j = AvJ / N
  End If
  With App
    If i <= N Then
      L("la" & P).Text = .Fixed(T(i), j) ' age
      L("ua" & P).Text = Bfmt(Upper(i) - T(i), j, True) ' tmp$ & Abs(v) ' +error
      L("ba" & P).Text = Bfmt(T(i) - Lower(i), j, False) 'tmp$ & Abs(v) ' -error
    End If
    If i > 1 Then
      e2(i) = Bfmt(Delta(i), j, True) 'tmp$ & .Fixed(Delta(i), j)
      e3(i) = Bfmt(DeltaPlus(i), j, True) 'tmp$ & .Fixed(DeltaPlus(i), j)
      e4$(i) = Bfmt(DeltaMinus(i), j, False) '"-" & .Fixed(DeltaMinus(i), j)
      L("da" & P).Text = e2(i)     ' Delta age
      L("uda" & P).Text = Trim(e3(i))  ' +error
      L("bda" & P).Text = e4$(i)       ' -error
    End If
  End With
Next i
On Error GoTo 0
Siglev_Click
If IsOff(Ch("cShowRes")) Then
  b("bGo").Text = "OK": b("Cancel").Visible = False
  e("eTrials").Enabled = False: L("lTrials").Enabled = False: s("sTrials").Visible = False
  If MacExcelX Then tB("tStrat").Visible = False
  For i = 1 To tB.Count: tB(i).Visible = True: Next
  tB("tTopBottom").Visible = (N < MaxB)
  G("gTopBottom").Visible = (N < MaxB)
  Ch("cShowRes").Text = "Put results on worksheet"
  ShowDialog Br
  GoSub Conf
Else
  d$ = "Best age   " & cl & "% conf.    Difference  " & cl & "% conf."
  With App
  For i = 1 To f
    If i <= N Then
      d1 = vbLf & L("la" & tSt(i)).Text: d2 = L("ua" & tSt(i)).Text
      d3 = L("ba" & tSt(i)).Text: d4$ = ""
    Else
      d1 = "": d2 = "": d3 = ""
      d4$ = vbLf & Space(3 + 1 * Useries) & _
        IIf(Useries, "Oldest minus youngest", "Bottom-Top age delta") & ": "
    End If
    If i > 1 Then
      d4$ = d4$ & e2(i) & e3(i) & e4$(i)
    End If
    d$ = d$ & d1 & d2 & d3 & d4$
  Next i
  End With
  'DatSht.Select
  With r
    Irange$ = .Address
    TopRow = .Row - 1: RightCol = .Column + .Columns.Count - 1
    AddResBox (d$)
    'r(1, 1).Select
  End With
  Tbox$ = Last(DatSht.TextBoxes).Name
  With DatSht.TextBoxes(Tbox$)
    With .Characters(1, InStr(.Text, vbLf) - 1).Font
      .Underline = True: .Bold = True
    End With
    If N < MaxB Then
      j = Len(d$)
      .Characters(j - Len(d4$) + 2, j).Font.Italic = True
    End If
    With .Font
      .Name = "Courier New": .Size = 10 + Mac
    End With
  End With
End If
If IsOn(Ch("cShowChart")) And Ch("cShowChart").Enabled Then
  DeltaGraf = IsOn(o("oDeltas"))
  BracketChart N, Nbins, r, DatSht, True, Tbox$, DeltaGraf
End If
DelSheet TempSht, True
If Not FromMenu Then ExcelCalc = Xcalc0: KwikEnd
Exit Sub

Conf: Clev = 0.6826: cl = "68"
If o("o95") = xlOn Then Clev = 0.95: cl = "95"
cf$ = cl & "% confidence"
tB("t95").Text = cf: tB("tDelt95").Text = cf
Return
BadRange: MsgBox "Invalid input-range for isoplot", , Iso
ExitIsoplot
End Sub

Sub BracketSpin()
Attribute BracketSpin.VB_ProcData.VB_Invoke_Func = " \n14"
With DlgSht("Bracket")
  .EditBoxes("eTrials").Text = .Spinners("sTrials").Value
End With
End Sub

Sub Siglev_Click()
Attribute Siglev_Click.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As String * 1, o As Object, T As Object
AssignD "Bracket", , , , o, , , T
If T("tDate").Text = "Date" Then
  s = IIf(o("o2sigma") = xlOn, "2", "1")
  T("tSigLev").Text = s & " sigma"
  s = IIf(o("opercent") = xlOn, "%", " ")
  T("tError").Text = s & " error"
End If
End Sub

Sub ConfLev_click()
Attribute ConfLev_click.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As String * 2, cf$
With DlgSht("Bracket")
  s = IIf(.OptionButtons("o95") = xlOn, "95", "68")
  cf$ = s & "%" & vbLf & "confidence"
  .TextBoxes("t95").Text = cf$
  .TextBoxes("tdelt95").Text = cf$
End With
End Sub

Sub BracketChart(ByVal Nbeds&, ByVal Nbins%, InpRange As Range, InpSht As Worksheet, _
  AsPicture As Boolean, ByVal BelowThisBox$, Optional DeltaGraf As Boolean = True)
Attribute BracketChart.VB_ProcData.VB_Invoke_Func = " \n14"
' Creates Aa small chart showing the probability distribution for Aa series
'  of stacked, dated beds
Dim ChtDat As Object, Cht As Object, Na$, i%, Rr$(2), b As Boolean, v&, M%
Dim Noffs%, s$, rw1%, rw2%, tr As Range, tMin#, tMax#
Dim j%, k%, T$, MinVal#, MaxVal#, BinW#, aa, Bb, cc, tmp$
Dim Kv#(), Kh#(), ScCt%, sc As Object, s1%, s2%
ReDim vv#(Nbins), H#(Nbins)
ViM DeltaGraf, True
On Error GoTo 0
StatBar "Creating Graphics..."
Set ChtDat = Ash: Na$ = ",'" & ChtDat.Name & "'!"
Cells(1, 1).Select ' Essential to have a single cell selected
Charts.Add
Ach.Location Where:=xlLocationAsObject, Name:=InpSht.Name
Set Cht = Last(Ash.Shapes)
With Cht
  .Left = InpRange(1, 3).Left: .Width = 150: .Height = 150
  If Len(BelowThisBox$) = 0 Then
    .Top = InpRange(1, 1).Top:
  Else
    Cht.Top = Bottom(Ash.TextBoxes(BelowThisBox$)) + 4
  End If
End With
MinVal = 1E+99: MaxVal = -MinVal
With Ach
  .ChartType = xlXYScatterSmoothNoMarkers
 .HasLegend = False
 Set sc = .SeriesCollection
  For i = 1 - DeltaGraf To Nbeds
    sc.NewSeries
    ChtDat.Select
    For j = 1 To 2
      If DeltaGraf Then
        k = 2 * Nbeds + 2 * (i - 1) + j - 2
      Else
        k = j + 2 * (i - 1)
      End If
      Set tr = sR(1, k, Nbins)
      BinW = tr(2) - tr(1)
      If False Then
        If j = 1 And (tr(1) - BinW * 1.1) < 0 Then
          ' Quadratic extrapolation to zero
          sR(1, k, 1, k + 1).Insert Shift:=xlDown
          Nbins = 1 + Nbins
          Set tr = sR(1, k, Nbins)
          ReDim Preserve vv(Nbins), H(Nbins)
          Quad tr(2, 1), tr(3, 1), tr(4, 1), tr(2, 2), tr(3, 2), tr(4, 2), , , cc
          tr(1, 1) = 0: tr(1, 2) = cc
        End If
      End If
      With tr
        Rr$(j) = tr.Address
        If j = 1 Then
          For k = 1 To Nbins: tr(k, 1) = -tr(k, 1): Next k
          If tr(Nbins) < MinVal Then MinVal = tr(Nbins)
          If tr(1) > MaxVal Then MaxVal = tr(1)
          For k = 1 To Nbins: vv(k) = -tr(k, 1): H(k) = tr(k, 2): Next k
          Kernel BinW, vv(), H(), Kv(), Kh(), Nbins
        End If
      End With
    Next j
    InpSht.Select
    Last(.SeriesCollection).Formula = _
      "=SERIES(" & Na$ & Rr$(2) & Na$ & Rr$(1) & "," & Str(i) & ")"
  Next i
  ScCt = sc.Count
  With .PlotArea
    v = IIf(ColorPlot, Menus("cStraw"), vbWhite)
    .Interior.ColorIndex = ClrIndx(v)
    .Top = 0: .Height = Ach.ChartArea.Height
    .Border.Weight = xlHairline
  End With
  With .ChartArea
    .Border.LineStyle = 0: .Interior.ColorIndex = xlNone
  End With
  With .Axes(xlCategory)
    If .MinimumScale < 0 Then .MinimumScale = 0
    .HasMajorGridlines = False
    .HasTitle = False: .TickLabelPosition = xlNone
    .MajorTickMark = xlNone: .MinorTickMark = xlNone
  End With
  With .Axes(xlValue)
    Do ' Optimize values for min & max of age axis
      b = True
      If (.MaximumScale - .MajorUnit) > MaxVal Then
        .MaximumScale = .MaximumScale - .MajorUnit
        b = False
      End If
      If (.MinimumScale + .MajorUnit) < MinVal Then
        .MinimumScale = .MinimumScale + .MajorUnit
        b = False
      End If
    Loop Until b
    If .MaximumScale > 0 Then
      ' Force top of chart to be exactly 0 and on a major tick
      .MinimumScale = .MajorUnit * (Int((.MaximumScale - .MinimumScale) / .MajorUnit) - 1)
      .MaximumScale = 0
    End If
    .MinorTickMark = xlInside: .MajorTickMark = xlCross
    .MinorUnit = .MajorUnit / 2
    k = -Int(App.Log10(.MajorUnit))
    If k > 0 Then ' Minimize decimal places in age-tick labels
      T$ = "0" & Dsep & String(k, "0")
      T$ = T$ & ";" & T$ & ";" & "0"
    Else
      T$ = "0;0;0"
    End If
    With .TickLabels
      .NumberFormat = T$: .Font.Size = 7
    End With
    .HasTitle = True
    With .AxisTitle
      tmp$ = IIf(DeltaGraf, IIf(Mac, "Age Delta", "Age difference"), "Age")
      If Mac Then .Orientation = xlVertical
      With .Characters: .Font.Size = 8: .Text = tmp$: End With
    End With
    With .MajorGridlines.Border
      .LineStyle = xlContinuous: .ColorIndex = Menus("iGray50")
    End With
    .HasMinorGridlines = True
    With .MinorGridlines.Border
      .LineStyle = xlDot: .Color = Menus("cGray75")
    End With
  End With
  s1 = ScCt - Nbeds + 1
  For i = s1 To ScCt
    With sc(i).Border
      If ColorPlot Then
        Select Case i
          Case s1: .Color = vbRed
          Case s1 + 1: .Color = vbBlue
          Case s1 + 2: .Color = Menus("cGreenBlack")
          Case s1 + 3: .Color = Menus("cPink")
          Case s1 + 4: .Color = Menus("cDkTeal")
          Case s1 + 5: .Color = vbBlack
          Case s1 + 6: .Color = Menus("cRedBlack")
          Case s1 + 7: .Color = Menus("cBlueBlack")
        End Select
      Else
        .Color = 0
      End If
      .Weight = xlThin
    End With
  Next i
End With
If AsPicture Then ConvertToPicture
NoUp False
NoUp True
'InpRange(1, 1).Select
StatBar
End Sub

Sub UnMix(T#(), SigmaT#(), pi#(), SigmaPi#(), _
  ByRef Misfit!, NoSoln As Boolean)
Dim i%, j%, nU%, jj%, kk%, Iter%, Converged As Boolean
Dim LLp0#, LL00#, LLm0#, LLpp#, LLpm#, LLmp#
Dim LLmm#, LL0p#, LL0m#, Num#, Denom#, Pj#, Vs#
Dim f#(), ss#(), Lpi#(), Lt#()
Dim td#, Pd#, Cinv As Variant
Iter = 0: NoSoln = False
ReDim ss(Ndat), VarA(Ndat), Lt#(Ncomp), Lpi#(Ncomp), f#(Ndat, Ncomp)
For i = 1 To Ndat: VarA(i) = SQ(SigmaA(i)): Next i
Do
  ff pi(), T(), f()
  For i = 1 To Ndat
    ss(i) = 0
    For j = 1 To Ncomp
      ss(i) = ss(i) + pi(j) * f(i, j)
    Next j
    If ss(i) = 0 Then ss(i) = 1E-32
  Next i
  For j = 1 To Ncomp
    Pj = 0
    For i = 1 To Ndat
      Pj = Pj + pi(j) * f(i, j) / ss(i)
    Next i
    pi(j) = Pj / Ndat
  Next j
  For j = 1 To Ncomp
    Num = 0: Denom = 0
    For i = 1 To Ndat
      Vs = pi(j) * f(i, j) / (VarA(i) * ss(i))
      Num = Num + aa(i) * Vs
      Denom = Denom + Vs
    Next i
    If Denom = 0 Then NoSoln = True: Exit Sub
    T(j) = Num / Denom
  Next j
  If Iter > 0 Then
    Converged = True
    For j = 1 To Ncomp ' Converged?
      td = Abs((T(j) - Lt(j)) / T(j))
      Pd = Abs((pi(j) - Lpi(j)))
      Converged = Converged And td < 0.00001 And Pd < 0.00001
      ' Degenerate sol'ns?
      If j > 1 Then Converged = Converged And (Abs(T(j) / T(j - 1) - 1) > 0.0002)
    Next j
  End If
  Iter = 1 + Iter
  For j = 1 To Ncomp: Lpi(j) = pi(j): Lt(j) = T(j): Next j
Loop Until (Converged And Iter > 1) Or Iter > 99
nU = 2 * Ncomp - 1 ' Now calculate the variance-covariance matrix
ReDim c(nU, nU), SigmaT(Ncomp)
If Ncomp > 1 Then ReDim SigmaPi(Ncomp - 1)
Toffs = iAverage(T) / Ncomp / 10000
ff pi(), T(), f(): LL00 = LL(pi(), f())
If BadMix Then GoTo Cant
For i = 1 To nU
  For j = 1 To nU
    If i >= j Then
      If i < Ncomp And j < Ncomp Then ' The Pi-Pi terms
        jj = i: kk = j
        If jj = kk Then ' Pi(i) variance
          LLp0 = LL2(jj, kk, 1, 0, 0, 0, pi(), T())
          LLm0 = LL2(jj, kk, -1, 0, 0, 0, pi(), T())
          c(i, j) = -(LLp0 - 2 * LL00 + LLm0) / SQ(PiOffs)
        Else ' Pi(i)-Pi(j) covariance
          LLpp = LL2(jj, kk, 1, 1, 0, 0, pi(), T())
          LLpm = LL2(jj, kk, 1, -1, 0, 0, pi(), T())
          LLmp = LL2(jj, kk, -1, 1, 0, 0, pi(), T())
          LLmm = LL2(jj, kk, -1, -1, 0, 0, pi(), T())
          c(i, j) = -(LLpp - LLpm - LLmp + LLmm) / SQ(2 * PiOffs)
        End If
      ElseIf i >= Ncomp And j >= Ncomp Then ' The T-T terms
        kk = i - Ncomp + 1: jj = j - Ncomp + 1
        If jj = kk Then ' T(i) variance
          LL0p = LL2(jj, kk, 0, 0, 0, 1, pi(), T())
          LL0m = LL2(jj, kk, 0, 0, 0, -1, pi(), T())
          c(i, j) = -(LL0p - 2 * LL00 + LL0m) / SQ(Toffs)
        Else ' T(i)-T(j) covariance
          LLpp = LL2(jj, kk, 0, 0, 1, 1, pi(), T())
          LLpm = LL2(jj, kk, 0, 0, 1, -1, pi(), T())
          LLmp = LL2(jj, kk, 0, 0, -1, 1, pi(), T())
          LLmm = LL2(jj, kk, 0, 0, -1, -1, pi(), T())
          c(i, j) = -(LLpp - LLpm - LLmp + LLmm) / SQ(2 * Toffs)
        End If
      Else ' The Pi-T covariance terms
        If i < Ncomp Then
          jj = i: kk = j - Ncomp + 1
        Else
          jj = j: kk = i - Ncomp + 1
        End If
        LLpp = LL2(jj, kk, 1, 0, 0, 1, pi(), T())
        LLpm = LL2(jj, kk, 1, 0, 0, -1, pi(), T())
        LLmp = LL2(jj, kk, -1, 0, 0, 1, pi(), T())
        LLmm = LL2(jj, kk, -1, 0, 0, -1, pi(), T())
        c(i, j) = -(LLpp - LLpm - LLmp + LLmm) / (4 * PiOffs * Toffs)
      End If
    End If
Next j, i
For i = 1 To nU ' Copy diagonal elements
  For j = 1 To nU
    If i < j Then c(i, j) = c(j, i)
Next j, i
Cinv = App.MInverse(c)
If IsError(Cinv) Then GoTo Cant
If Ncomp = 1 Then
  SigmaT(1) = Sqr(Cinv(1))
Else
  For i = 1 To 2 * Ncomp - 1
    If Cinv(i, i) < 0 Then GoTo Cant
  Next i
  For j = 1 To Ncomp
    i = j + Ncomp - 1
    SigmaT(j) = Sqr(Cinv(i, i))
    If j < Ncomp Then SigmaPi(j) = Sqr(Cinv(j, j))
  Next j
  ReDim SigmaRho#(nU, nU)
  For i = 1 To nU ' Create the Sigma-Rho array in case needed
    For j = 1 To nU
      If i = j Then ' Sigma(i) = Sqrt[Var(i)]
        SigmaRho(i, j) = Sqr(Cinv(i, i))
      Else          ' Rho(i,j) = Cov(i,j)/[Sigma(i)*Sigma(j)]
        SigmaRho(i, j) = Cinv(i, j) / Sqr(Cinv(i, i) * Cinv(j, j))
      End If
  Next j, i
End If
Misfit = LL00
Exit Sub
Cant: NoSoln = True
End Sub

Sub MixShowWithData()
Attribute MixShowWithData.VB_ProcData.VB_Invoke_Func = " \n14"
Dim c, Bu, c1, c2
AssignD "Mix", , , c, , , , , Bu
Set c1 = c("cShowWithData")
Set c2 = c("cShowMatrix")
If IsOn(c1) Then Bu.Enabled = True
c2.Enabled = IsOn(c1)
If IsOff(c1) Then c("cShowMatrix") = xlOff
End Sub
Sub MixCancelClick()
Attribute MixCancelClick.VB_ProcData.VB_Invoke_Func = " \n14"
Canceled = True: AskInfo = False: MixConstr = False
End Sub
Sub ShowMix_click()
Attribute ShowMix_click.VB_ProcData.VB_Invoke_Func = " \n14"
With DlgSht("Mix").CheckBoxes
  .Item("cShowMatrix").Enabled = IsOn(.Item("cShowWithData"))
End With
End Sub
Sub Dummy()
Attribute Dummy.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub
Sub MixSpinClick()
Attribute MixSpinClick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim e As Object, s As Object
AssignD Name:="Mix", EditBoxes:=e, Spinners:=s
e("eGuess").Text = LTrim(s("sGuess").Value)
End Sub
Sub MixConstructClick()
Attribute MixConstructClick.VB_ProcData.VB_Invoke_Func = " \n14"
MixConstr = True: Canceled = False: AskInfo = False
End Sub
Sub Guess_click()
Attribute Guess_click.VB_ProcData.VB_Invoke_Func = " \n14"
Dim tB As Object, e As Object, o As Object, Eg As Object
Dim s As Object, La As Object, b As Boolean, j%, i As String * 1, q As Boolean
AssignD "Mix", , e, , o, La, , tB, , s
Set o = o("oIsoGuess"): Set s = s("sGuess"): Set La = La("lGuess"): Set Eg = e("eGuess")
b = IsOn(o)
Eg.Enabled = b: s.Enabled = b: La.Enabled = b
q = Not b
For j = 1 To 6
  i = tSt(j)
  'E("eT" & i).Visible = -1: E("eP" & i).Visible = -1
  e("eT" & i).Enabled = q: e("eP" & i).Enabled = q
Next j

'Tb("tTage").Visible = -1: Tb("tTProp").Visible = -1
tB("tTage").Enabled = q: tB("tTProp").Enabled = q
If b Then Ncomp = MinMax(2, 10, Val(Eg.Text))
End Sub

Sub AttachCumGauss(ByVal PicName$, ByVal FromMenu As Boolean, s As Range, T#(), pi#(), _
  NoSoln As Boolean, ByVal ChartWithData As Boolean, ByVal ChartAsSheet As Boolean)
Attribute AttachCumGauss.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ChtSht As Object, j%, MinY!, MaxY!, Yspred!, W!, v
Dim AsPicture As Boolean, i%
' The input data are in 2 adj. cols if invoked by clicking on UNMIX icon;
'   in cols 1 & 3 if from Isoplot itself.
If Not (ChartWithData Or ChartAsSheet) Then Exit Sub
AsPicture = ChartWithData
NoUp
Nbins = 8 * (1 - (N > 15) - (N > 30) - (N > 45) - (N > 60))
Sheets.Add:  PlotDat$ = "PlotDat"
MakeSheet PlotDat$, ChrtDat
DoShape = True: ColorPlot = True
SymbCol = 7: SymbRow = Max(1, SymbRow)
GaussCumProb N, True, , InpDat(), InpDat(), , , ChartAsSheet, , 1, 3
Set ChtSht = Ash
PutPlotInfo
If Ncomp > 1 And Not NoSoln Then
  With ChtSht
    With .Axes(xlValue, 1): MinY = .MinimumScale: MaxY = .MaximumScale: End With
    Yspred = MaxY - MinY
    For j = 1 To Ncomp
      With ChrtDat
        .Cells(SymbRow, SymbCol) = T(j): .Cells(1 + SymbRow, SymbCol) = T(j)
        .Cells(SymbRow, 1 + SymbCol) = MinY
        .Cells(1 + SymbRow, 1 + SymbCol) = MaxY - Yspred / 50
      End With
      ChrtDat.Activate
      Set v = sR(SymbRow, SymbCol, 1 + SymbRow, 1 + SymbCol, ChrtDat)
      .SeriesCollection.Add v, xlColumns, False, 1, False
      AddSymbCol 2
      With Last(.SeriesCollection)
        .AxisGroup = 1: .MarkerStyle = xlNone
        With .Border
          .Weight = xlThick: .LineStyle = xlContinuous: .Color = Menus("cGreen")
        End With
      End With
    Next j
    If Not ChartAsSheet Then .ChartArea.Interior.ColorIndex = xlNone
    .Select
  End With
End If
If Not FromMenu Then ChrtDat.Cells(12, 2) = s.Address
ActiveWindow.Zoom = 100
RescaleOnlyShapes False
If AsPicture Then
  CopyPicture ExternalInvoked:=False, Name:=PicName$
  If Not ChartAsSheet Then
    DelSheet ChtSht
    DelSheet ChrtDat
  End If
  With Ash.Pictures(PicName$).ShapeRange
    .Fill.Visible = True: .Line.Visible = True
    .Fill.ForeColor.RGB = RGB(255, 255, 200)
  End With
ElseIf Not ChartAsSheet Then
  Ach.ChartArea.Copy
  Sheets(ChrtDat.Cells(1, 2).Text).Select
  Range(ChrtDat.Cells(12, 2).Text).Cells(1, 3).Select
  Ash.Paste
  With Last(Ash.ChartObjects)
    W = .Width: .Width = 300: .Height = .Height * .Width / W
  End With
End If
End Sub
Sub BracketShowResClick()
Attribute BracketShowResClick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim b As Boolean
With DlgSht("Bracket").CheckBoxes
  b = IsOn(.Item("cShowRes"))
  With .Item("cShowChart")
    If Not b Then .Value = xlOff
    .Enabled = b
  End With
End With
End Sub

Sub Quad(ByVal x1, ByVal x2, ByVal x3, ByVal y1, ByVal y2, ByVal y3, _
  Optional A, Optional b, Optional c)
' Solve y=ax^2 + bx + c for 3 points
Dim Coef As Variant, y#(3, 1), X#(3, 3), InvX As Variant
y(1, 1) = y1: y(2, 1) = y2: y(3, 1) = y3
X(1, 1) = SQ(x1): X(2, 1) = SQ(x2): X(3, 1) = SQ(x3)
X(1, 2) = x1: X(2, 2) = x2: X(3, 2) = x3
X(1, 3) = 1: X(2, 3) = 1: X(3, 3) = 1
With App
  InvX = .MInverse(X)
  If IsError(InvX) Then
    MsgBox "Error in matrix inversion, sub Quad"
    ExitIsoplot
  End If
  Coef = .MMult(InvX, y)
End With
If NIM(A) Then A = Coef(1, 1)
If NIM(b) Then b = Coef(2, 1)
If NIM(c) Then c = Coef(3, 1)
End Sub

Sub ShowBracketChartClick()
Attribute ShowBracketChartClick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim c As Object, Bu As Object, o As Object, G As Object, b As Boolean
AssignD "Bracket", , , c, o, , G, , Bu
b = IsOn(c("cShowChart"))
'g("gPlot").Enabled = b
o("oAges").Enabled = b
o("oDeltas").Enabled = b
End Sub

Sub Kernel(ByVal BinWidth, v#(), H#(), Kv#(), Kh#(), ByVal N&)
Attribute Kernel.VB_ProcData.VB_Invoke_Func = " \n14"
' V() contains the data values, H() their frequencies
Dim i%, j%, MinV, MaxV, Spred, Kspred
Dim MinK, MaxK, Kinter, Kval, c, s5, T
Dim Npts%
s5 = Sqr(5): c = 0.75 / s5: Npts = 100
With App
  MinV = .Min(v): MaxV = .Max(v)
End With
Spred = MaxV - MinV
MinK = MinV - 4 * BinWidth
MaxK = MaxV + 4 * BinWidth
Kinter = (MaxK - MinK) / Npts
ReDim Kv(Npts), Kh(Npts)
For i = 1 To Npts
  Kv(i) = MinK + Kinter * (i - 1)
  Kh(i) = 0
  For j = 1 To N
    T = Abs(Kv(i) - v(j)) / BinWidth
    If T < s5 Then
      Kval = c * (1 - SQ(T) / 5)
    Else
      Kval = 0
    End If
    Kh(i) = Kh(i) + Kval * H(j)
Next j, i
i = 0
Do
  i = i + 1
Loop Until IsEmpty(Cells(1, i))
For j = 1 To Npts
  Cells(j, i) = Kv(j)
  Cells(j, i + 1) = Kh(j)
Next j
End Sub

Function Bfmt(ByVal v, ByVal f%, ByVal DefPlus As Boolean)
Attribute Bfmt.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$
If DefPlus Then
  s$ = IIf(v >= 0, "+", "-")
Else
  s$ = IIf(v > 0, "-", "+")
End If
Bfmt = s$ & App.Fixed(Abs(v), f)
End Function

Sub Freq(ByVal N&, v#(), Bins%(), BinLim#(), _
  ByVal Nbins%, ByVal Bstart, ByVal Bwidth, _
  MaxBinVal)
' Input vector V , Bins() contains bin boundaries, H is count in Bins
' The first bin contains values >=BinStart but <(BinsStart+BinWidth), et cetera.
Dim i&, j%, MaxBinNum&, MaxBinIndx&, LwrInd&
ReDim Bins(Nbins), BinLim(Nbins + 1)
QuickSort v()
For j = 1 To Nbins + 1
  If j <= Nbins Then Bins(j) = 0
  BinLim(j) = Bstart + (j - 1) * Bwidth
Next j
LwrInd = 1: MaxBinNum = 0
For j = 1 To Nbins
  For i = LwrInd To N
    If v(i) >= BinLim(j) Then
      If v(i) >= BinLim(j + 1) Then
        LwrInd = i: Exit For
      End If
      Bins(j) = 1 + Bins(j)
    End If
  Next i
Next j
For j = 1 To Nbins
  If Bins(j) > MaxBinNum Then
    MaxBinNum = Bins(j): MaxBinIndx = j
  End If
Next j
If MaxBinIndx > 0 Then
  MaxBinVal = BinLim(MaxBinIndx) + Bwidth / 2
End If
End Sub

Public Sub Detrital(Optional FromMenu = False)
' Input is a 2-column range of ages of stratigraphically ordered units (oldest at bottom, youngest at top)
'  with 1-SigmaA absolute errors in the column to the right of the ages.
' Output are in the 3 adjacent columns, being, for each unit, the lower limit on the unit's age (@95%
'  conf.), the mode of the age, and he upper 95%-conf limit.
Dim Age#(), GrafSht, PicName$, M$, Ntmp%
Dim SigmaA#(), T#(), Clev!, pi#(), A, c, d, Sht$(), Nsht%, Boo As Boolean
Dim InpRange As Range, TempSht As Worksheet, BinLim#(), LrM As Range, LrL As Range, LrU As Range
Dim i&, j%, k%, LowerInd&, InpCol%, MaxBinVal#, UpperInd, Results$, tmp#, Ord$
Dim LowerLim, AgeBins%(), TrialYoungest#(), DatSht As Worksheet, Msg$, Rr&, cc%
Dim SimAge#(), SortedSimAge#(), Ndates&, YoungestAge#, Src As SeriesCollection, DatRange As Range
Dim ww, hh, Lower95, Upper95, ModeYoungest#, LowerErr, UpperErr, PlotDatSht, Ncols%
Dim ResLeft!, ResTop!, ResWidth!, ResHt!, Nareas%, Nrows%, SigmaTi#(), TopRow&, RightCol%, Bad As Boolean
Dim SortedYoungest#(), TickInterval#, StartBin#, NextBin#, EndBin#, LastNum%, test#
Dim MinYoungestAge#, MaxYoungestAge#, MaxBinNum#, AverageYoungest#, MedianYoungest#, AgeTest#
Dim YoungthIn%, YoungthIndx%, TestN%, TestAge#(), TestSigma#(), sIndx%(), sAge#()
Const Ntrials = 40000
Nbins = 60

AssignIsoVars
NoUp
On Error GoTo 1
Set DatRange = Selection
GoTo 2

1: KwikEnd

2: On Error GoTo 0

If FromMenu Then
  Ndates = UBound(InpDat, 1)
  ReDim tmpInp(Ndates, 2), T#(Ndates)

  For i = 1 To Ndates
    For j = 1 To 2
      tmpInp(i, j) = InpDat(i, j)
  Next j, i

Else
  YoungestDetrital = True: Isotype = 25
  SigLev = 1
  AbsErrs = True
  Set DatSht = Ash
  DatSheet$ = DatSht.Name
  DoPlot = True: ColorPlot = True
  RangeCheck Nrows, Ncols, Nareas

  If Ncols = 2 Then
    WhatKindOfData Nareas, Nrows
    NoUp
    Set DatRange = Selection: Irange$ = DatRange.Address
    ParseAgeRange DatRange, Ndates, Nareas, T(), SigmaTi(), TopRow, RightCol, Bad
  End If

  If Bad Or Ndates < 2 Or Ncols <> 2 Then
    MsgBox "Invalid input range" & vbLf & "(need 2 columns and 2 or more rows of numeric data)", , Iso
    ExitIsoplot
  End If

End If

ReDim IndX(Ndates), SigmaT#(Ndates), TestAge(Ndates), TestSigma(Ndates)

If Not FromMenu Then
  ReDim InpDat(Ndates, 2), tmpInp(Ndates, 2)

  For i = 1 To Ndates
    InpDat(i, 1) = T(i)
    tmp = SigmaTi(i) / SigLev
    If Not AbsErrs Then tmp = tmp / 100 * T(i)
    InpDat(i, 2) = tmp
  Next i

End If

Clev = 95
ReDim Age(Ndates), SigmaA(Ndates), TrialYoungest(Ntrials, 1)
Randomize Timer
LowerInd = Ntrials * (100 - Clev) / 2 / 100
UpperInd = Ntrials - LowerInd
YoungthIndx = Max(1, Menus("DetritalYoungthIndx"))

For i = 1 To Ndates
  Age(i) = InpDat(i, 1)    ' DatRange(i, 1)
  SigmaA(i) = InpDat(i, 2) ' DatRange(i, 2)
Next i

SortCol Age, sAge, Ndates, 0, sIndx

' Ignore dates that are obviously too old to worry about
TestN = 0
AgeTest = sAge(YoungthIndx) + 3 * SigmaA(sIndx(YoungthIndx))

For i = 1 To Ndates
  test = Age(i) - 3 * SigmaA(i)

  If test < AgeTest Then
    TestN = 1 + TestN
    TestAge(TestN) = Age(i)
    TestSigma(TestN) = SigmaA(i)
  End If

Next i

ReDim Preserve TestAge(TestN), TestSigma(TestN)
ReDim SimAge(TestN, 1)
Set DatSht = Ash

For i = 1 To Ntrials

  If i Mod 200 = 0 Then
    StatBar "WAIT    " & Str(Ntrials - i)
  End If

  For j = 1 To TestN
    SimAge(j, 1) = Gaussian(TestAge(j), TestSigma(j))
  Next j

  SortCol SimAge(), SortedSimAge(), TestN, 1
  TrialYoungest(i, 1) = SortedSimAge(YoungthIndx)
Next i

StatBar
SortCol TrialYoungest, SortedYoungest, Ntrials, 1
MinYoungestAge = SortedYoungest(1)
MaxYoungestAge = SortedYoungest(Ntrials)
'LowerLim = SortedYoungest(LowerInd)
Tick MaxYoungestAge - MinYoungestAge, TickInterval
TickInterval = 5 * TickInterval
NextBin = 0

Do
  NextBin = NextBin + TickInterval
Loop Until NextBin > MinYoungestAge

StartBin = NextBin - TickInterval
EndBin = StartBin

Do
  EndBin = EndBin + TickInterval
Loop Until EndBin > MaxYoungestAge

BinWidth = (EndBin + TickInterval - StartBin) / Nbins
'MaxBinVal = SortedYoungest(Ntrials)
Freq Ntrials, SortedYoungest(), AgeBins(), BinLim(), Nbins, _
    StartBin, BinWidth, MaxBinVal
j = 1: ww = 0

For i = 2 To Nbins
  If AgeBins(i) > AgeBins(j) Then j = i
Next i

ww = BinLim(j) + BinWidth / 2
ModeYoungest = ww
AverageYoungest = App.Average(SortedYoungest)
MedianYoungest = iMedian(SortedYoungest, True)
Lower95 = SortedYoungest(LowerInd / 2)
Upper95 = SortedYoungest(Ntrials - LowerInd / 2)
LowerErr = ModeYoungest - Lower95
UpperErr = -ModeYoungest + Upper95
StatBar
LastNum = Right(Str(YoungthIndx), 1)

Select Case LastNum
  Case 0, Is > 3: Ord = "th "
  Case 1:  Ord = IIf(YoungthIndx = 1, "", "st ")
  Case 2:  Ord = "nd "
  Case 3:  Ord = "rd "
End Select

If YoungthIndx > 1 Then Ord = Trim(Str(YoungthIndx)) & Ord
Results = "Age of the " & Ord & "youngest grain is " & Drnd(ModeYoungest, 5) & " +" & _
          Trim(Drnd(Upper95 - ModeYoungest, 2)) & " -" & _
          Trim(Drnd(ModeYoungest - Lower95, 2)) & " Ma at 95% conf."
Load TuffZirc

With TuffZirc
  .Caption = "Suite of single detrital grains"
  With .tbResults
    .Text = vbLf & Results & vbLf & vbLf & "(assumes input errors are " & _
      IIf(AbsErrs, "absolute", "percent") & "," & Str(SigLev) & "-sigma)" & _
      vbLf & vbLf
    .AutoSize = True
    .Width = .Width * 1.1
  End With
  .lbExplain.Visible = True
  .lbYoungth.Visible = True
  With .tbYoungth
    .Visible = True
    YoungthIn = Max(1, Menus("DetritalYoungthIndx"))
    .Value = YoungthIn
  End With
  With .spYoungth
    .Visible = True
    .Value = YoungthIn
  End With
  .Height = Bottom(.lbExplain) + 25
  .bCalcNoExit.Visible = False   ' 09/12/09 -- added
  .Show
End With

If Canceled Then
  Menus("DetritalYoungthIndx") = YoungthIn
  KwikEnd

ElseIf AskInfo Then
  ShowHelp "YoungestDetritalHelp"

ElseIf Menus("DetritalYoungthIndx") <> YoungthIn Then
  GoTo 2
End If

If DoPlot Then
  ColorPlot = True: DoShape = False
  Sheets.Add
  Set PlotDatSht = Ash
  Nsht = Sheets.Count
  ReDim Sht(Nsht)
  For i = 1 To Nsht: Sht(i) = Sheets(i).Name: Next
  GaussCumProb Ntrials, 0, -1, SortedYoungest(), , RGB(192, 192, 255), , 0, _
   MaxBinNum, , , Nbins, StartBin, EndBin, BinWidth
  Set GrafSht = ActiveSheet
  PlotDatSht.Activate
  GrafSht.Activate
  With ActiveChart
    .ChartArea.Interior.Color = RGB(255, 144, 144)
    .PlotArea.Interior.Color = Straw 'RGB(225, 255, 225)
    With .Axes

      For i = 1 To 2
        With .Item(i)
         .TickLabels.Font.Size = 20 + 2 * (i = 2)
         .HasTitle = (i = 1)

        If i = 1 Then
          With .AxisTitle.Characters
           .Text = "Age (Ma)"
           .Font.Size = 36: .Font.Bold = False
          End With
        End If

        If i = 1 Then
          .AxisTitle.AutoScaleFont = False
          .TickLabels.AutoScaleFont = False
        Else
          .TickLabelPosition = xlNone
        End If

       End With
      Next i

    End With
    With .PlotArea
      .Left = .Left = 30
      .Width = ActiveChart.ChartArea.Width - 30
      .Top = 0
    End With
    PlotDatSht.Activate
    A = .Axes(2).MaximumScale * 0.97
    j = 0

    For j = 1 To 255
      j = j + 1
      If Cells(1, j) <> "" Then Exit For
    Next j

    j = j + 2
    Cells(1, j) = ModeYoungest
    Cells(1, j + 1) = 0
    Cells(2, j) = ModeYoungest
    Cells(2, j + 1) = A
    Set LrM = Range(Cells(1, j), Cells(3, j + 1))
    Cells(1, j + 2) = Lower95
    Cells(1, j + 3) = 0
    Cells(2, j + 2) = Lower95
    Cells(2, j + 3) = A
    Set LrL = Range(Cells(1, j + 2), Cells(2, j + 3))
    Cells(1, j + 4) = Lower95
    Cells(1, j + 5) = 0
    Cells(2, j + 4) = Lower95
    Cells(2, j + 5) = A
    Set LrU = Range(Cells(1, j + 2), Cells(2, j + 3))
    GrafSht.Activate
    .SeriesCollection.Add LrM, xlColumns, False, 1
    With .SeriesCollection(.SeriesCollection.Count)
     .Border.Weight = xlThick
     .Border.Color = 255
    End With

    If DoShape Then RescaleOnlyShapes False ', True

    .CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
    DatSht.Activate ' Switch to source-data sheet
    Ash.Pictures.Paste.Select
    NoAlerts
    GrafSht.Delete
    With Selection
      .Name = "Youngest" & Right(Timer, 4)
      PicName = .Name
      ' Scale down size
      With .ShapeRange
        .Fill.ForeColor.RGB = vbWhite
      End With
      ww = .Width: hh = .Height
      .Width = 250: .Height = hh / ww * .Width
      .Left = DatRange(2, 3).Left + 15
      .Top = DatRange(1, 1).Top + 15
      With .Border: .Color = vbBlack: .Weight = xlHairline: End With
    End With
  End With
End If

If DoPlot Or ShowRes Then

  If ShowRes Then
    DatSht.Activate
    ResLeft = Right_(DatRange) + 5
    ResTop = DatRange.Top
    ResWidth = 77
  Else
    Set A = Selection
    ResLeft = A.Left: ResTop = A.Top + 5
    ResWidth = A.Width + 53
  End If

  ResHt = 56
  Ash.Shapes.AddShape(msoShapeRoundedRectangle, ResLeft, ResTop, ResWidth, _
      ResHt).Select
  With Selection
      .Name = IIf(DoPlot, PicName & "_", "Youngest" & Right(Timer, 4))
      .Characters.Text = Results: .AutoSize = True
       .Height = .Height + 5
      .VerticalAlignment = xlCenter
      .HorizontalAlignment = xlCenter

      If DoPlot Then
        .Left = A.Left + A.Width / 2 - .Width / 2
        .Top = A.Top - .Height + 3
      End If

    With .ShapeRange
      .Fill.ForeColor.SchemeColor = 47
      .Line.ForeColor.SchemeColor = 8
    End With
  End With

  If DoPlot Then
    Ash.Shapes.Range(Array(PicName, PicName & "_")).Group.Select
    Selection.Top = DatRange.Top
  End If

  DatRange.Cells(1, 1).Select
End If

Menus("DetritalYoungthIndx") = YoungthIndx
StatBar
ExitIsoplot
End Sub
