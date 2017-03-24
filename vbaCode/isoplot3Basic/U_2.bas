Attribute VB_Name = "U_2"
' ISOPLOT module U_2
Option Private Module
Option Explicit: Option Base 1
Dim Rmc As Boolean

Private Sub IsAutoscale()
AutoScale = True
End Sub

Private Sub IsXyLimit()
XYlim = True: Canceled = False
End Sub

Sub SetupIsoRes(tbx$, Optional AgeLabel, Optional InterLabely) ' Prepare the "Isochron Results" dialog box
Attribute SetupIsoRes.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, LintEr$, rLint$, rUint$, NoIso As Boolean, Asymm#, b As Boolean, P%
Dim UintEr$, M$, m2$, u#, v#, s1$(2), s2$(2), s0$, Cres8#, Cres10#
Dim i%, j%
Dim Cres3#, Cres4#, Irat#, Tcov#, IratErr#
Dim tB$(20), sa$, McOK As Boolean, test#, L As Object, re As Object
Dim Ch As Object, Op As Object, df As Object, Gp As Object, Lanch As Boolean, Uanch As Boolean
Dim Lines As Object, Circles As Object, Behind As Object, Front As Object
Dim k%, vv$, ee$, LowerDist, UpperDist, YesMC As Object, Eb As Object
Dim ReA As Object, d As String * 3, spi As Object, McC As Boolean, Bu As Object
AssignD , ResBox, Eb, Ch, Op, L, Gp, YesMC, Bu, spi, , , re, , df
Set YesMC = YesMC("tSuggestMC"): Set ReA = Gp("gUpperInter")
Set Front = Op("oFront"): Set Behind = Op("oBehind")
Set Lines = Op("oLines"): Set Circles = Op("oCircles")
d = "---"
NoIso = (OtherXY And iLambda(OtherIndx) = 0)
Cres8 = Crs(8): Cres10 = Crs(10)
If ConcPlot And Crs(38) > 0 Then Cres10 = Crs(38) 'Lower-int err w/o dce
Cres3 = Crs(3): Cres4 = Crs(4)
YesMC.Visible = False
If PbPlot Then
  rLint$ = ""
ElseIf ConcPlot Then
  If Anchored Then
    Lanch = (Abs(Cres8 - AnchorAge) < 0.01)
    Uanch = (Abs(Crs(9) - AnchorAge) < 0.01)
  End If
  If Cres8 <> 0 Then ' Lower inter age
    If Cres10 <> 0 Or Lanch Then ' Lower inter age error
      NumAndErr Cres8, Cres10, 2, Lint$, ee$
      If Anchored And Abs(Cres8 - AnchorAge) < 0.01 Then
        Lint$ = tSt(AnchorAge): LintEr$ = tSt(AnchorErr)
      Else
        LintEr$ = ee$
        If Crs(40) Then
          M$ = Lint$
          NumAndErr Cres8, Crs(40), 2, M$, ee$
          If LintEr$ <> ee$ Then LintEr$ = LintEr$ & "   [" & pm & ee$ & "] "
        End If
      End If
      rLint$ = Lint$ & pm & LintEr$
    Else
      Lint$ = Sp(Cres8, 0): LintEr$ = " ***"
      rLint$ = Lint$ & " " & pm & LintEr$
    End If
    rLint$ = rLint$ & " Ma"
  Else
    Lint$ = d: LintEr$ = d: rLint$ = "None"
  End If
ElseIf Not (Dim3 And UseriesPlot) And Not (Inverse And Not ArgonPlot) Then
  If NoIso Then
    Lint$ = "": LintEr$ = "": rLint$ = ""
  Else
    If Robust Then
      With yf
        If .UpprInter = .LwrInter Then
          Lint$ = Sd(.Intercept, 5, , True)
          rLint$ = Lint$
        Else
          NumAndErr .Intercept, (.UpprInter - .LwrInter) / 2, 2, Lint$, ""
          NumAndErr .Intercept, .UpprInter - .Intercept, 2, "", s1$(1)
          NumAndErr .Intercept, .LwrInter - .Intercept, 2, "", s1$(2)
          rLint$ = Lint$ & "  +" & s1$(1) & "/-" & s1$(2)
        End If
      End With
    ElseIf ArgonPlot And Inverse Then
      NumAndErr 1 / Cres3, Cres4 / SQ(Cres3), 2, Lint$, LintEr$
      rLint$ = Lint$ & pm & LintEr$
    Else
      NumAndErr Cres3, Cres4, 2, Lint$, LintEr$
      rLint$ = Lint$ & pm & LintEr$
    End If
  End If
Else
  NumAndErr Crs(36), Crs(37), 2, Lint$, LintEr$
  rLint$ = Lint$ & pm & LintEr$
End If
If Not PbAnchor Then
  If Crs(9) <> 0 Then ' Upper int age
    If (ConcPlot And Crs(39) <> 0) Or (Not ConcPlot And Crs(11) <> 0) Or Robust Then
      If Robust Then
        With yf
          If .UpprAge = .LwrAge Then
            Uint$ = Sd(Crs(9), 4, , True)
            rUint$ = Uint$
          Else
            NumAndErr Crs(9), (.UpprAge - .LwrAge) / 2, 2, Uint$, ""
            NumAndErr Crs(9), .UpprAge - Crs(9), 2, Uint$, s1$(1)
            NumAndErr Crs(9), .LwrAge - Crs(9), 2, Uint$, s1$(2)
            rUint$ = Uint$ & "  +" & s1$(1) & "/-" & s1$(2)
          End If
        End With
      Else
        NumAndErr Crs(9), Crs(11 - 28 * ConcPlot), 2, Uint$, ee$
        If ConcPlot And Uanch Then
          Uint$ = tSt(AnchorAge): UintEr$ = tSt(AnchorErr)
        Else
          UintEr$ = ee$
          If ConcPlot And Crs(41) <> 0 Then
            NumAndErr Crs(9), Crs(41), 2, "", ee$
            If UintEr$ <> ee$ Then UintEr$ = UintEr$ & "   [" & pm & ee$ & "] "
          End If
        End If
        rUint$ = Uint$ & pm & UintEr$
      End If
      L("lAge").Enabled = True
    Else
      Uint$ = Sp(Crs(9), 0): UintEr$ = " ***"
      rUint$ = Uint$ & " " & pm & UintEr$
    End If
    rUint$ = rUint$ & IIf(UseriesPlot, " ka", " Ma")
  Else
    Uint$ = d: UintEr$ = d: rUint$ = "None "
  End If
  L("lAge").Enabled = True
Else
  L("lAge").Enabled = False: rUint$ = ""
End If
McOK = False
If Not Dim3 And Not Robust Then
  If ConcPlot Then
    If (Cres8 <> 0 Or Crs(9) <> 0) And (Crs(7) > 0.05) And _
      Not (Anchored And PbAnchor) And Not Robust Then McOK = True
  ElseIf ArgonPlot Then
    If yf.Prob > MinProb And Crs(23) = 1 Then McOK = True
  End If
End If
If McOK Then
  m2$ = IIf(Crs(23) = 1, "and errors (most reliable)", "(most reliable, but analytical errs only)")
  Ch("cMC").Text = "Monte Carlo ages " & m2$
End If
Ch("cMC").Enabled = McOK: Gp("gMC").Enabled = McOK
If ConcPlot Then ConcMcHistClick
Gp("gHisto").Enabled = (McOK And ConcPlot)
Gp("gPlotWhere").Enabled = (McOK And ConcPlot)
Ch("cHisto").Enabled = (McOK And ConcPlot)
If ConcPlot And McOK And Crs(23) = 1 Then ' For Model-1 only
  Asymm = 0
  LowerDist = Crs(5) - ConcX(Cres8)  ' Determine whether centroid of Y'fit
  UpperDist = ConcX(Crs(9)) - Crs(5) '  is closest to upper or lower inter.
  If LowerDist > UpperDist Then      ' Test upper-inter  error-asymmetry
    test = Crs(15) - Crs(9)
    If Crs(9) <> 0 And test <> 0 Then
      Asymm = Abs((Crs(9) - Crs(14)) / test - 1) ' Pos-err/Neg-err minus 1
    End If
  ElseIf Cres3 <> 0 And Cres4 <> 0 And Cres8 <> 0 Then  ' Test lower
    test = Cres4 - Cres8
    If Cres8 <> 0 And test <> 0 Then
      Asymm = Abs((Cres8 - Cres3) / test - 1)
    End If
  End If
  If Asymm > 0.25 Then
    YesMC.Visible = True
    If Not Mac Then
      With YesMC.Font
        .Name = "Arial": .Bold = Mac: .Size = 9 - 2 * Mac
      End With
    End If
  End If
  ConcMcBoxClick
ElseIf McOK Then
  ConcMcBoxClick
End If
If Anchored Then Ch("cLIgtZero") = xlOff
Ch("cLIgtZero").Enabled = False
If Not McOK Then Ch("cMC") = xlOff: Ch("cLIgtZero") = xlOff
DoMC = (McOK And IsOn(Ch("cMC"))): McC = (DoMC And ConcPlot)
With Ch("cLIgtZero"): .Enabled = (DoMC And ConcPlot And Not Anchored): .Visible = ConcPlot: End With
With Ch("cWLE"): .Enabled = (ConcPlot And DoPlot And AutoScale): .Visible = .Enabled: End With
Eb("eNtrials").Enabled = DoMC: L("lNtrials").Enabled = DoMC
spi("sNtrials").Enabled = DoMC: Eb("eNbins").Enabled = McC
L("lNbins").Enabled = McC: spi("sNbins").Enabled = McC
Op("oLower").Enabled = McC: Op("oUpper").Enabled = McC
Op("oSeparateSheet").Enabled = McC: Op("oDataSheet").Enabled = McC
Ch("cMC").Text = "Monte Carlo age errors"
If ConcPlot Then Ch("cMC").Text = Ch("cMC").Text & " (most reliable)"
If McOK Then
  If ConcPlot Then
    If Not Cdecay0 Then Ch("cWLE") = xlOff
    If Cdecay0 Then IncludeDecayConstErrsClick
  End If
  If Val(Eb("eNtrials").Text) < Thou Then Eb("eNtrials").Text = Thou
End If
L("lProbVal").Visible = Not Robust
If Not Robust Then
  Msw$ = Mrnd(Crs(6))
  If Regress And ConcPlot Then Lir$ = Lint$ & pm & LintEr$
  Uir$ = Uint$ & pm & UintEr$ & IIf(UseriesPlot, " ka", " Ma")
  L("lProbVal").Text = ProbRnd(Crs(7))
End If
With L("lTopLabel")
  .Left = ReA.Left + (ReA.Width) / 3
  If Robust Then
    .Text = "Robust Regression"
  Else
    If Not ConcPlot Or Not Dim3 Then
      .Text = "Model" & Str(Crs(23)) & " Solution   (" & pm & "95%-conf.)"
      If ConcPlot And Lambda235err > 0 And Lambda238err > 0 And _
        (InStr(LintEr$, "[") > 0 Or InStr(UintEr$, "[") > 0) Then
          .Text = .Text & "   without [with] decay-const. errs"
          .Left = ReA.Left + 5
      End If
    ElseIf ConcPlot And Dim3 And Planar3D Then
      .Text = "3-D Planar regression"
    End If
  End If
End With
ResBox.Buttons("bCancel").Visible = True
ResBox.Buttons("bDetails").Enabled = (Not (ConcPlot And Anchored) And Not Robust)
L("lNpts").Text = "# points": L("lNptsVal").Text = sn$(N + Anchored)
With L("lMSWD"): .Text = "MSWD": .Enabled = Not Robust: End With
L("lMSWDval").Text = IIf(Robust, "", Msw$)
With L("lProb"): .Text = "Probability of fit": .Enabled = Not Robust: End With
L("lInterLabel").Text = IIf(UseriesPlot And Not Dim3, "", rLint$)
L("lAgeLabel").Text = rUint$
L("lPbIsoAge").Text = "": L("lPbGrowth").Text = ""
If ConcPlot Then
  L("lIntercept").Text = "Lower intercept:":   L("lAge").Text = "Upper intercept:"
  df.Text = IIf(Dim3, "3-D Planar ", "") & "Concordia-Intercept Ages"
Else
  If Not OtherXY Then
    s0$ = IsoPlotTypes.Cells(Isotype).Text
    For k = 1 To 2
      M$ = IIf(k = 1, AxX$, AxY$)
      j = 0
      Do
        j = 1 + j
      Loop Until Not IsNum(M$, j)
      s1$(k) = "": s2$(k) = ""
      If M$ <> "" And j > 1 Then
        s1$(k) = Left$(M$, j - 1)
        P = InStr(M$, "/")
        If P > 0 Then
          s2$(k) = Mid$(M$, 1 + P, j - 1)
        End If
      End If
    Next k
    s$ = "ratio"
    If Normal And s1$(2) <> "" And s2$(2) <> "" Then
      s$ = s1$(2) & "/" & s2$(2) ' Y-numer/X-numer
    ElseIf ArgonPlot And s2$(2) <> "" And s1$(2) <> "" Then
      s$ = s2$(2) & "/" & s1$(2) ' Y-denom/Y-numer
    ElseIf Inverse And UThPbIso And s2$(1) <> "" And s1$(1) <> "" Then
      s$ = s2$(1) & "/" & s1$(1) ' X-denom/X-numer (U-Pb or Th-Pb inverse isochron)
    ElseIf s1$(2) <> "" And s1$(1) <> "" Then
      s$ = s1$(2) & "/" & s1$(1) ' Y-numer/X-numer
    End If
  Else
    s0$ = AxX$ & "-" & AxY$: s$ = AxY$
  End If
  If InStr(LCase(s0$), "isochron") = 0 Then
    df.Text = s0$ & " Isochron Ages"
  Else
    df.Text = s0$ & " Ages"
  End If
  L("lAge").Text = "Isochron age:"
  If PbPlot Or (Not Dim3 And UseriesPlot) Then
    L("lIntercept").Text = ""
  ElseIf ClassicalIso Or UThPbIso Then  '~!
    L("lIntercept").Text = "Initial " & s$ & "="
  Else
    L("lIntercept").Text = s$ & " intercept: "
  End If
  If ArgonPlot Then
    L("lPbIsoAge").Text = "       (at " & AtJay$ & ")"
  ElseIf Crs(23) = 3 Then
    L("lPbIsoAge").Text = "Initial " & s$ & " variation =" & ErFo(Cres3, Crs(12), 2) & " (2-sigma)"
  ElseIf PbPlot Then
    L("lIntercept").Text = ""
    If Crs(11) <> BadT And Crs(11) <> 0 And Crs(26) <> BadT And Crs(26) <> 0 Then
      If Abs(Crs(26) / Crs(11) - 1) > 0.1 And Crs(9) > 0 Then
        L("lIntercept").Text = "w. decay-const errs:"
        L("lInterLabel").Text = "(" & ErFo(Crs(9), Crs(26), 2, True) & " Ma)"
      End If
    End If
    If Crs(13) <> BadT Then
      M$ = "Growth-curve intercept"
      If Crs(14) <> BadT Then M$ = M$ & "s"
      M$ = M$ & " at " & Sp(Crs(13), 0)
      If Crs(14) <> BadT Then
        M$ = M$ & " and " & Sp(Crs(14), 0)
      End If
      L("lPbIsoAge").Text = M$ & " Ma"
    End If
  End If
  If SmNdIso Then
    With Menus("CHUR")
      L("lPbGrowth").Text = "Epsilon(CHUR) = " & Sp(Crs(34), -1, True) _
        & " (CHUR Sm/Nd, R =" & .Cells(2, 7) & ", " & .Cells(2, 8) & ")"
    End With
  ElseIf PbPlot And PbType = 1 And Crs(9) <> BadT Then '~!
    u = Prnd(PbR(Crs(9), 0), -2)
    v = Crs(1 - 2 * Inverse) * u + Crs(3 + 2 * Inverse)
    M$ = "207Pb/204Pb at 206Pb/204Pb ="
    L("lPbGrowth").Text = M$ & sn$(u) & " is " & Sp(v, -2)
  End If
End If
If ConcPlot And Inverse And Dim3 Then
  If Crs(32) <> BadT Then
    M$ = "Isochron age in Common-Pb Plane = " & Sp(Crs(32), 0)
    If Crs(33) <> BadT Then M$ = M$ & pm & Sp(Crs(33), 0)
    L("lPbIsoAge").Text = M$ & " Ma"
  End If
  If Crs(34) <> BadT Then
    M$ = "Pb growth-curve intercept"
    If Crs(35) <> BadT Then M$ = M$ & "s"
    M$ = M$ & " at " & Sp(Crs(34), 0)
    If Crs(35) <> BadT Then M$ = M$ & " and " & Sp(Crs(35), 0)
    L("lPbGrowth").Text = M$ & " Ma"
  End If
End If
If ConcPlot Then
  If Anchored Then
    sa$ = "Anchored at"
    If PbAnchor Then
      sa$ = sa$ & " 207/206 =" & Str(Anchor76) & pm & tSt(AnchorErr)
    Else
      sa$ = sa$ & Str(AnchorAge) & " " & pm & Str(AnchorErr) & " Ma"
    End If
    s$ = sa$
  End If
  tbx$ = L("lTopLabel").Text & " on " & L("lNptsVal").Text & " points" & vbLf
  If Len(L("lInterLabel").Text) Then
    If Not Anchored Or AnchorAge <> Drnd(Cres8, 5) Then
      s$ = L("lIntercept").Text & " " & L("lInterLabel").Text
    End If
    tbx$ = tbx$ & s$ & vbLf
  End If
  If Len(L("lAgeLabel").Text) Or (Anchored And PbAnchor) Then
    If Anchored And (PbAnchor Or AnchorAge = Drnd(Crs(9), 5)) Then
      s$ = sa$
    Else
      s$ = L("lAge").Text & " " & L("lAgeLabel").Text
    End If
    tbx$ = tbx$ & s$ & vbLf
  End If
  tbx$ = tbx$ & L("lMSWD").Text & " = " & L("lMSWDval").Text & _
    ", " & L("lProb").Text & " = " & L("lProbVal").Text
Else
  tbx$ = L("lTopLabel").Text & " on " & L("lNptsVal").Text & _
    " points" & vbLf & "Age = " & L("lAgeLabel").Text
  s$ = L("lIntercept").Text & L("lInterLabel").Text
  If Len(s$) Then tbx$ = tbx$ & vbLf & s$
  If Robust Then
    Uir$ = L("lAgeLabel").Text
  Else
    tbx$ = tbx$ & vbLf & L("lMSWD").Text & " = " & _
      L("lMSWDval").Text & ", Probability = " & L("lProbVal").Text
  End If
  If Len(LTrim(L("lPbIsoAge").Text)) Then tbx$ = tbx$ & vbLf & L("lPbIsoAge").Text
End If
b = (ConcPlot And AutoScale And DoPlot)
Gp("gAgeTicks").Visible = b: Circles.Visible = b: Lines.Visible = b
If b And Circles = xlOff And Lines = xlOff Then Circles = xlOn
If b Then Circles.Caption = ConcSymbInfo
b = (b And Cdecay0)
With Ch("cWLE_MC"): .Visible = ConcPlot: .Enabled = IsOn(Ch("cMC")): End With
Ch("cWLE").Visible = b: Behind.Visible = b  ' Show option to show concordia
Front.Visible = b: Gp("gLambdaErrs").Visible = b '  curve w. decay-const errs?
b = (ConcPlot And AutoScale)
If b Then
  b = (Lambda235err > 0 Or Lambda238err > 0)
  If Not b Then Ch("cWLE") = xlOff
  Ch("cWLE").Enabled = b
  Behind.Enabled = (IsOn(Ch("cWLE")) And DoShape)
  Front.Enabled = Behind.Enabled
  If Behind.Enabled And Behind = xlOff And Front = xlOff Then Behind = xlOn
End If
b = (Ch("cWLE") = xlOff Or Not Ch("cWLE").Enabled)
Gp("gAgeTicks").Enabled = b: Circles.Enabled = b: Lines.Enabled = b
L("lIntercept").Visible = Not NoIso Or (PbPlot And PbType = 1)
Gp("gLowerInter").Visible = L("lIntercept").Visible
L("lAge").Visible = Not NoIso: L("lAgeLabel").Visible = Not NoIso
If NoIso Then Ch("cShowRes") = xlOff
Ch("cShowRes").Enabled = Not NoIso
If Robust Then Ch("cAddWtdResids") = xlOff
Ch("cAddWtdResids").Visible = Not Robust
End Sub
Function ConcSymbInfo() As String
Attribute ConcSymbInfo.VB_ProcData.VB_Invoke_Func = " \n14"
Dim IndX%, s$
IndX = Match(Opt.AgeTikSymbol, Menus("AgeTikSymbolCode"))
s$ = LCase(Menus("AgeTikSymbol")(IndX))
If Len(s$) = 1 Then s$ = s$ & " symbol"
ConcSymbInfo = "as small " & s$ & "s with Horizontal labels"
End Function
Sub Caveat_Isores()
Attribute Caveat_Isores.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, P$
s$ = DlgSht("IsoRes").Labels("lTopLabel").Text
If Left(s$, 5) = "Model" Then
  P$ = Mid(s$, 7, 1)
  s$ = "Model" & P$
  If P$ = "1" Then s$ = s$ & IIf(yf.Prob >= MinProb, "hi", "lo")
ElseIf Left(s$, 6) = "Robust" Then
  s$ = "Robust"
ElseIf Left(s$, 1) = "3" Then
  s$ = "PlanarConc"
End If
ShowHelp s$ & "Help"
End Sub
Sub ShowHelp(ByVal Helper$)
Dim T As Object, P As Object, H As Object, W!
Set H = Help
LoadUserForm H
Set T = H.HelpText: T.Height = 400
With T
  .Text = vbLf & Caveats(Helper$) & viv$
  .AutoSize = True
  .Height = .Height + 10
  .AutoSize = False
   H.Height = .Height + 65
End With
H.Show
Unload H
End Sub
Sub Caveat_ConcAge()
Attribute Caveat_ConcAge.VB_ProcData.VB_Invoke_Func = " \n14"
ShowHelp "ConcAgeHelp"
End Sub
Sub Caveat_Bracket()
Attribute Caveat_Bracket.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$
s$ = IIf(StackedUseries, "BracketUseriesHelp", "BracketAgeHelp")
ShowHelp s$
End Sub
Sub Caveat_RobustRes()
Attribute Caveat_RobustRes.VB_ProcData.VB_Invoke_Func = " \n14"
ShowHelp "RobustHelp"
End Sub
Sub Caveat_TuffZirc()
Attribute Caveat_TuffZirc.VB_ProcData.VB_Invoke_Func = " \n14"
ShowHelp "TuffZircHelp"
End Sub
Sub Caveat_Mix()
Attribute Caveat_Mix.VB_ProcData.VB_Invoke_Func = " \n14"
ShowHelp "MixHelp"
End Sub
Sub Help_ArStepAge()
Attribute Help_ArStepAge.VB_ProcData.VB_Invoke_Func = " \n14"
ShowHelp "ArStepAgeHelp"
End Sub
Sub Help_Wtd()
Attribute Help_Wtd.VB_ProcData.VB_Invoke_Func = " \n14"
ShowHelp "WtdHelp"
End Sub
Sub Help_ExtVar()
ShowHelp "ExtVarHelp"
End Sub
Sub Help_BiWt()
ShowHelp "BiWtHelp"
End Sub
Sub Help_Median()
ShowHelp "MedianHelp"
End Sub
Private Sub OKbutton()
Canceled = False: AskInfo = False
End Sub
Private Sub CancelButton()
Canceled = True: AskInfo = False
End Sub
Private Sub InfoButton()
Canceled = False: AskInfo = True
End Sub
Private Sub PlotProjected() ' Checkbox "cPlotProj" clicked in "ProjPts" dialog-box
Dim P As Object, po As Object, Pe As Object, pL As Object, pc As Object
Set P = DlgSht("ProjPts"): Set po = P.OptionButtons
Set Pe = P.EditBoxes:   Set pc = P.CheckBoxes
PlotProj = IsOn(pc("cPlotProj"))
Pe("eWhat46").Enabled = PlotProj: po("oProjPar").Enabled = PlotProj
po("oProjSpec").Enabled = PlotProj
If PlotProj Then
  If IsOff(po("oProjPar")) And IsOff(po("oProjSpec")) Then po("oProjPar") = xlOn
  Pe("eWhat46").Enabled = IsOn(po("oProjSpec"))
Else
  po("oProjSpec") = xlOff: po("oProjPar") = xlOff
  Pe("eWhat46").Enabled = False
End If
If Pe("eWhat46").Enabled Then ProjZ = EdBoxVal(Pe("eWhat46"))
End Sub

Private Sub PlotParZ()  ' Option-button "ProjPar" clicked in "ProjPts" dialog-box
Dim P As Object
Set P = DlgSht("ProjPts")
With P.EditBoxes("eWhat46")
  .Text = "0": .Enabled = False
End With
End Sub

Private Sub PlotParWhat()  ' Option-button "ProjSpec" clicked in "ProjPts" dialog-box
Dim P As Object, Pe As Object
Set P = DlgSht("ProjPts"): Set Pe = P.EditBoxes("eWhat46")
If Not IsNumeric(Pe.Text) Then Pe.Text = "0"
Pe.Enabled = True: P.Focus = "eWhat46"
End Sub

Sub ProjProc()  ' Setup the "ProjPts" dialog box
Attribute ProjProc.VB_ProcData.VB_Invoke_Func = " \n14"
Dim P As Object, e As Object
Set P = DlgSht("ProjPts"): Set e = P.EditBoxes("eWhat46")
PlotProjected
If PlotProj And IsOn(P.OptionButtons("oProjSpec")) Then
  ProjZ = EdBoxVal(e)
Else
  e.Text = ""
End If
End Sub

Sub ResboxProc(Optional MC As Boolean = False, Optional SlpErr, Optional YintErr, Optional Xinterr)
Attribute ResboxProc.VB_ProcData.VB_Invoke_Func = " \n14"
Dim u#, v#, tmp$, YL As Object, Le$, Ue$, s$, i%, tbx As Object, mcB As Object
Dim TbxN%, dfT$, vv$, Vs$, vi$, ee$, tB As Boolean, cb As Object, P$, pL$, A
AssignD , YorkRes, , cb, , YL, , , mcB
Set mcB = mcB("MonteCarlo")
ViM MC, False
tB = (Crs(7) < 0.01)
' Round slope, inter to 1sigma sigfigs if p>.05, else to 95%conf sigfigs
NumAndErr Crs(1), Crs(16 + 14 * tB), 2, Vs$, "" ' Slope, sf's @95% or 1sig
NumAndErr Crs(3), Crs(17 + 13 * tB), 2, vi$, "" ' Inter,  "
YL("Slp").Text = Vs$
If MC Then
  YL("l1sap").Text = "68.3% conf. limit"
  YL("S95").Text = Sd(-SlpErr(4), 2, True) & "  " & Sd(SlpErr(5), 2, True)
Else
  YL("l1sap").Text = "1-sigma a priori"
  YL("S95").Text = ErFo(Crs(1), Crs(2), 2)
End If
YL("Int").Text = vi$
If MC Then
  YL("sap").Text = Sd(-SlpErr(1), 2, True) & "  " & Sd(SlpErr(2), 2, True)
Else
  YL("sap").Text = ErFo(Crs(1), Crs(16), 2) ' +- slope a-p
End If
If MC Then
  YL("l1sos").Text = "": YL("sis").Text = "": YL("iis").Text = ""
Else
  YL("l1sos").Text = "1-sigma obs. scatter"
  YL("sis").Text = ErFo(Crs(1), Crs(18), 2) ' +- slope scatt
End If
YL("lXbar").Text = Sd$(Crs(5), 6, 0, -1)  ' Xbar
YL("lYbar").Text = Sd$(Crs(20), 6, 0, -1) ' Ybar
YL("lLowerInt").Enabled = (Crs(8) <> 0)
YL("Xint").Enabled = (Crs(9) <> 0)
If PbAnchor And AnchorErr = 0 Then
  YL("iap").Text = "":  YL("iis").Text = ""
  YL("i95").Text = "": YL("Xint").Enabled = False
Else
  YL("Xint").Enabled = True
  If MC Then
    YL("iap").Text = Sd(-YintErr(1), 2, True) & "  " & Sd(YintErr(2), 2, True)
    YL("i95").Text = Sd(-YintErr(4), 2, True) & "  " & Sd(YintErr(5), 2, True)
  Else
    YL("iap").Text = ErFo(Crs(3), Crs(17), 2) ' +-inter a-p
    YL("iis").Text = ErFo(Crs(3), Crs(19), 2) ' +-inter scatt
    YL("i95").Text = ErFo(Crs(3), Crs(4), 2)  ' +-inter 95%
  End If
End If
If Crs(8) Then
  u = Crs(12) - Crs(8): v = Crs(13) - Crs(8)
  If Crs(12) <> 0 Then Le$ = Sd$(u, 2) Else Le$ = "-inf."
  If Crs(13) <> 0 Then Ue$ = Sd$(v, 2) Else Ue$ = "inf."
  tmp$ = Lint$ & "  (" & Le$ & " +" & Ue$ & ")"
  YL("LwrInterAge").Text = tmp$ & " Ma"
Else
  YL("LwrInterAge").Text = ""
End If
If Crs(9) <> 0 And Not PbAnchor Then
  u = Crs(14) - Crs(9): v = Crs(15) - Crs(9)
  If Crs(14) Then Le$ = Sd$(u, 2) Else Le$ = "-inf."
  If Crs(15) Then Ue$ = Sd$(v, 2) Else Ue$ = "inf."
  tmp$ = Uint$ & "  (" & Le$ & " +" & Ue$ & ")"
  YL("XintErr").Enabled = True:  YL("XintErr").Text = tmp$ & " Ma"
Else
  YL("XintErr").Text = "":  YL("XintErr").Enabled = False
End If
mcB.Enabled = Not MC And Not Dim3 And Not ConcPlot And yf.Prob > 0.05
If ConcPlot Then
  dfT$ = "Regression Params and Concordia Intercepts"
  With YL("NoLambdaErrs")
    .Text = "Age errors above do not include decay-constant errors"
    .Visible = True
  End With
  YL("lLowerInt").Text = "Lower Intercept":  YL("Xint").Text = "Upper Intercept"
Else
  dfT$ = "Regression Parameters"
  If MC Then dfT$ = "Monte Carlo " & dfT$
  YL("NoLambdaErrs").Visible = False
  YL("lLowerInt").Text = "                  X-Intercept ="
  NumAndErr Crs(21), Crs(22), 2, vv$, ee$, , True
  YL("Xint").Text = vv$:     YL(18).Text = ""
  YL("XintErr").Enabled = True
  If MC Then
    YL("XintErr").Text = Sd(-Xinterr(4), 2, True) & "  " & Sd(Xinterr(5), 2, True)
  Else
    YL("XintErr").Text = "   " & ee$
  End If
  YL("XintErr").Text = YL("XintErr").Text & "  (95% conf.)"
End If
YL("RhoSI").Text = "Slope-intercept error correlation = " & RhoRnd(yf.RhoInterSlope)
If ConcPlot And Anchored Then
  A = Array("sap", "iap", "sis", "iis", "s95", "i95", "YbarEq", "lXbar", "lYbar", _
       "lLowerInt", "xint", "LwrInterAge", "xinterr")
  For i = 1 To 13
    YL(A(i)).Text = ""
  Next i
  cb(1) = xlOff: cb(1).Enabled = False
Else
  cb(1).Enabled = True
End If
YorkRes.DialogFrame.Text = dfT$
YL("lMP").Text = "MSWD = " & Mrnd(Crs(6)) & ",  Probability = " & ProbRnd(Crs(7))
With cb("cPlotErrEnv")
  .Enabled = (DoPlot And Not Dim3)
  .Value = xlOff
End With
With cb("cAddWtdResids")
  .Visible = OtherXY: .Value = xlOff
End With
ShowBox YorkRes, True
PlotErrEnv = IsOn(cb("cPlotErrEnv"))
If MC Then P = "  " Else P = pm
pL$ = IIf(yf.Prob < MinProb, " (95% conf)", " (2 sigma)")
If cb(1) = xlOn Or OtherXY Then
  s$ = "Slope = " & YL("Slp").Text & P & YL("S95").Text & pL$ & vbLf & _
       "Inter = " & YL("Int").Text & P & YL("i95").Text & vbLf & _
       "Xbar = " & YL("lXbar").Text & ",  Ybar =" & YL("lYbar").Text
  If ConcPlot Then
    s$ = s$ & vbLf & "Lower inter = " & YL("LwrInterAge").Text & _
         vbLf & "Upper inter = " & YL("XintErr").Text
  ElseIf Crs(23) = 3 Then
    s$ = s$ & vbLf & ResBox.Labels("lPbIsoAge").Text
  End If
  s$ = s$ & vbLf$ & YL("lMP").Text
  AddResBox s$, -1, 0, LightGreen
  DetailsShown = True
End If
InsertWtdResids YorkRes
End Sub

Sub RegresDetailsPick()
Attribute RegresDetailsPick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Bad As Boolean, L As Object
Set L = DlgSht("IsoRes").Labels
If Dim3 And Planar3D And (ConcPlot Or ArgonPlot Or OtherXY) Then
  KentResProc
ElseIf Linear3D And ConcPlot Then
  ' Code for linear-3d display here
Else
    Rmc = False
    ResboxProc
  If Rmc Then
    ReDim s(6), Yi(6), xi(6)
    MCyorkfit N, 4000, Bad, s(), Yi(), xi()
    If Bad Then Exit Sub
    Crs(16) = s(3): Crs(17) = Yi(3): Crs(22) = xi(6)
    Crs(10) = s(6): Crs(4) = Yi(6)
    ResboxProc True, s(), Yi(), xi()
  End If
End If
End Sub

Sub AnchorProc() ' Prepare the "Anchored" dialog-box
Attribute AnchorProc.VB_ProcData.VB_Invoke_Func = " \n14"
Dim e As Object, o As Object, L As Object, e1, e2, tB As Boolean, DS As Object
Dim G As Object, df As Object, M As Object
AssignD , Anch, e, , o, L, G, Dframe:=df
Set M = G("gMConly")
L("lAge").Accelerator = "": L("lAgeErr").Accelerator = ""
If ConcPlot And Not Inverse And IsOn(o("oCommPb")) Then o("oAge") = xlOn
o("oRefChord").Enabled = DoPlot
If Not DoPlot Then o("oRefChord") = xlOff
If IsOff(o("oCommPb")) And IsOff(o("oRefChord")) Then o("oAge") = xlOn
If IsOn(o("oCommPb")) Then
  With L("lAge"): .Text = "Common-Pb 207/206": .Accelerator = "P": End With
  With L("lAgeErr"): .Text = "error": .Accelerator = "e": End With
ElseIf IsOn(o("oAge")) Then
  With L("lAge"): .Text = "Age": .Accelerator = "g": End With
  With L("lAgeErr"): .Text = "error": .Accelerator = "e": End With
ElseIf IsOn(o("oRefChord")) Then
  GoSub RefChordSetup
End If
PbAnchor = IsOn(o("oCommPb")): AgeAnchor = IsOn(o("oAge"))
RefChord = IsOn(o("oRefChord"))
e1 = EdBoxVal(e("eAge")): e2 = EdBoxVal(e("eAgeErr"))
If PbAnchor Then
  Anchor76 = e1:  AnchorErr = e2
ElseIf AgeAnchor Then
  AnchorAge = e1: AnchorErr = e2
ElseIf RefChord Then
  AnchorT1 = e1:  AnchorT2 = e2
End If
If Regress Then
  o("oCommPb").Enabled = Inverse: o("oAge").Enabled = True
  o("oRefChord").Enabled = False
  If Normal Then
    o("oAge") = xlOn: o("oCommPb") = xlOff
  ElseIf IsOff(o("oAge")) And IsOff(o("oCommPb")) Then
    o("oAge") = xlOn
  End If
End If
tB = (Regress And IsOn(o("oAge")))
M.Visible = tB
o("oContinuous").Visible = tB: o("oGaussian").Visible = tB
With df
  .Height = IIf(tB, M.Top - .Top - 5, Bottom(M) - .Top + 5)
End With
If Not Regress Then
  o("oCommPb").Enabled = False: o("oAge").Enabled = False
  o("oRefChord") = xlOn
  GoSub RefChordSetup
End If
Exit Sub
RefChordSetup:
L("lAge").Text = "Lower age": L("lAge").Accelerator = "L"
L("lAgeErr").Text = "Upper age": L("lAgeErr").Accelerator = "U"
Return
End Sub

Sub ConvertSymbols(T As Object)
Attribute ConvertSymbols.VB_ProcData.VB_Invoke_Func = " \n14"
' Changes all "sigma" or "-sigma" to greek-letter sigma, "SQRT" to square-root sign
Dim P%, G$(5), s(5) As String * 1, q%, i%, TT$, IsCaption As Boolean
G$(1) = "-sigma": G$(2) = " sigma": G$(3) = "sigma": G$(4) = "sqrt": G$(5) = "lambda"
s(1) = "s": s(2) = "s": s(3) = "s": s(4) = Sqrt: s(5) = "lt.Characters.Text"
On Error GoTo 2
TT = T.Characters.Text
'IsCaption = False
'GoTo 4
'1: On Error GoTo 2
'tt = t.Caption
'IsCaption = True
GoTo 4
2: Exit Sub
4: P = InStr(TT, "ts+")
If P > 0 And IsCaption Then
  SymbolChar T + 1, P, 2
End If
For i = 1 To 5
 q = Len(G$(i)) - 1
 Do
  P = InStr(LCase(T.Characters.Text), G$(i))
  If P > 0 Then
   T.Characters(P, 1).Text = s(i)
   SymbolChar T, P, 1
   T.Characters(P + 1, q).Text = ""
  End If
Loop Until P = 0
Next i
End Sub

Private Sub AutoBinClick() ' Auto-binwidth checkbox clicked: clear the bin-width/#bins box
DlgSht("AxLab").EditBoxes("eBinNumWidth").Text = ""
InclHistClick
End Sub

Sub InclHistClick() ' Handle histo-spec dialog in AxisLabels
Attribute InclHistClick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim v As Boolean, L As Object, e As Object, c As Object, s As Object, A As Object, Lw As Object, f As Object
AssignD "AxLab", , e, c, , L
Set A = c("cAutoBins"): Set s = e("eBinStart"):    Set Lw = L("lBinStart")
Set f = c("cFilledBins")
Set c = c("cInclHist"): Set e = e("eBinNumWidth"): Set L = L("lBinNumWidth")
v = IsOn(c)
A.Visible = v: e.Visible = v:  L.Visible = v: f.Visible = v
s.Visible = (v And IsOff(A)):  Lw.Visible = s.Visible
If v Then
 'If E.Text = "" And
 With L
  If IsOn(A) Then
    e.Text = tSt(10 * (1 - (N > 30) - 2 * (N > 60)))
   .Text = "Number of bins": .Accelerator = "N"
  Else
   .Text = "Bin width":      .Accelerator = "w"
  End If
 End With
End If
End Sub

Sub ConcAgeWLE_click()
Attribute ConcAgeWLE_click.VB_ProcData.VB_Invoke_Func = " \n14"
Dim d As Object, o As Object, WLE As Boolean
If Not DoShape Or Not Cdecay0 Then Exit Sub
Set d = DlgSht("ConcAge"): Set o = d.OptionButtons
WLE = IsOn(o("oWLE"))
o("oFront").Enabled = WLE: o("oBehind").Enabled = WLE
d.GroupBoxes("gConcBand").Enabled = WLE
End Sub

Sub ProbDiag_click() ' Live ProbPlot interaction
Attribute ProbDiag_click.VB_ProcData.VB_Invoke_Func = " \n14"
Dim c As Object, G As Object, o As Object, e As Object, L As Object
Dim Bb As Boolean, ba As Boolean, pn%, s$, sn$
sn$ = tSt(UBound(InpDat, 1))
AssignD "ProbPlot", , e, c, o, L, G
pDots = IsOn(o("oDots")): pBars = Not pDots
ba = (ndCols = 1): Bb = (pBars And ba)
pInpSig = (IsOff(o("oFromData")) And ba)
pRegress = IsOn(c("cRegression"))
G("gErrbarSig").Enabled = Bb:   o("oDots").Enabled = True
o("oFromData").Enabled = Bb:    o("oUserSpec").Enabled = Bb
o("oErrBars").Enabled = True
If Not ba Then e("eSigma").Text = "": o("oFromData") = xlOn
e("eSigma").Enabled = Bb And pInpSig
L("lSigma").Enabled = e("eSigma").Enabled
e("eFirst").Enabled = pRegress: e("eLast").Enabled = pRegress
L("lFirst").Enabled = pRegress: L("lLast").Enabled = pRegress
L("lLast").Text = "Rank of Highest point to include (N=" & sn$ & ")"
With c("cInclStats")
  If Not pRegress Then .Value = xlOff
  .Enabled = pRegress
End With
If pRegress Then
  If Len(e("eFirst").Text) = 0 Then e("eFirst").Text = "1"
  If Len(e("eLast").Text) = 0 Then e("eLast").Text = sn$
End If
If pBars Then
  s$ = IIf(AbsErrs, "abs.", "%")
  L("lSigma").Text = Chr$(48 + SigLev) & "-sigma " & s$ & " errors for all points"
End If
End Sub

Function ProbDiagOK() ' Test ProbPlot input
Attribute ProbDiagOK.VB_ProcData.VB_Invoke_Func = " \n14"
Dim N&, c As Object, e As Object, o As Object, pn&, s$
N = UBound(InpDat, 1): ProbDiagOK = False
AssignD "ProbPlot", , e, c, o
If IsOn(c("cRegression")) Then
  pFirst = Max(1, Val(e("eFirst").Text)): pLast = Min(N, Val(e("eLast").Text))
  pn = pLast - pFirst + 1
  If pn < 5 Then MsgBox "Need at least 5 points for regression": Exit Function
Else
  pFirst = 1: pLast = N
End If
If ndCols = 1 And IsOn(o("oErrBars")) And IsOff(o("oFromData")) Then
  s$ = Trim(e("eSigma").Text)
  If Len(s$) = 0 Then MsgBox "You must enter a Sigma value": Exit Function
  If Val(s$) <= 0 Then MsgBox "Sigma must be >0": Exit Function
  pSigma = Val(s$)
End If
AxY$ = e("eYaxis").Text
ProbDiagOK = True
End Function

Function NIM(ByVal v) As Boolean
Attribute NIM.VB_ProcData.VB_Invoke_Func = " \n14"
NIM = Not IM(v)
End Function

Function IM(ByVal v) As Boolean
Attribute IM.VB_ProcData.VB_Invoke_Func = " \n14"
IM = (IsMissing(v) Or IsNull(v))
End Function

Sub IMN(v, ByVal Default)
Attribute IMN.VB_ProcData.VB_Invoke_Func = " \n14"
v = IIf(IsMissing(v) Or IsNull(v) Or Not IsNumeric(v), Default, v)
End Sub

Function sR(ByVal r1&, ByVal c1%, _
  Optional r2, Optional c2, Optional HostSheet) As Range
Attribute sR.VB_ProcData.VB_Invoke_Func = " \n14"
ViM r2, r1 ' Shortcut for cell-defined range
ViM c2, c1
If IM(HostSheet) Then
  Set sR = Range(Cells(r1, c1), Cells(r2, c2))
Else
  Set sR = Range(HostSheet.Cells(r1, c1), HostSheet.Cells(r2, c2))
End If
End Function

Sub LoadUserForm(Uform As Object, Optional DoShow As Boolean = False)
Attribute LoadUserForm.VB_ProcData.VB_Invoke_Func = " \n14"
Dim r As Range, N%, Na$, i%, j%, cOff%, First%, fs!
Set r = TW.Sheets("UserFrms").Range(Uform.Name)
fs = -2 * MacExcelX
cOff = 0 '-7 * (Left(OpSys, 7) <> "Windows")
Load Uform
With Uform
  .Font.Name = r(1, 2 + cOff).Text
  .Font.Size = Val(r(1, 3 + cOff)) - fs
  .Height = Val(r(1, 4 + cOff))
  .Left = Val(r(1, 5 + cOff))
  .Top = Val(r(1, 6 + cOff))
  .Width = Val(r(1, 7 + cOff))
  N = r.Rows.Count: First = r.Row
  On Error Resume Next
  For i = 0 To .Controls.Count - 1
    Na$ = LCase(.Controls(i).Name)
    For j = 2 To N
      If LCase(r(j, 1).Text) = Na$ Then Exit For
    Next j
    If j > N Then
        GoTo 1 'MsgBox "Error in LoadUserForm sub -- no entry for '" & Na$ & "'": KwikEnd
    End If
    With .Controls(i)
      With .Font
        .Name = r(j, 2 + cOff).Text
        .Size = r(j, 3 + cOff) - fs
        .Bold = r(j, 8 + cOff)
      End With
      .Top = Val(r(j, 6 + cOff)):   .Left = Val(r(j, 5 + cOff))
      .Width = Val(r(j, 7 + cOff)): .Height = Val(r(j, 4 + cOff))
    End With
1
  Next i
  On Error GoTo 0
End With
If DoShow Then Uform.Show
End Sub

Sub ResboxDoMCpressed()
Attribute ResboxDoMCpressed.VB_ProcData.VB_Invoke_Func = " \n14"
Rmc = True
End Sub

Sub DelSheet(Optional Sht, Optional ShowWait As Boolean = False)
Attribute DelSheet.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As Object ' Delete the active sheet or active chart
ViM ShowWait, False
If IM(Sht) Then
  Set s = Ash
ElseIf IsObject(Sht) Then
  Set s = Sht
Else
  Set s = Sheets(Sht)
End If
On Error Resume Next
With App
  If s.Name <> "" Then
    .DisplayAlerts = False
    If ShowWait Then Call StatBar("Wait")
    s.Delete
    If ShowWait Then Call StatBar
  End If
End With
End Sub

Sub MakeXY(xy As Variant, X#(), y#(), N&, DoCheck As Boolean)
Attribute MakeXY.VB_ProcData.VB_Invoke_Func = " \n14"
' Extract x() and y() from either a range or a single 2D array.  If a range and Docheck,
'  extract only valid numeric x-y pairs.
Dim i%, xx, yy, M%, A As Areas
If IsObject(xy) Then M = xy.Rows.Count Else M = UBound(xy, 1)
N = 0
If DoCheck And IsObject(xy) Then
  Set A = xy.Areas
  If A.Count > 2 Then Exit Sub
  ReDim X(M), y(M)
  For i = 1 To M
    If A.Count = 2 Then
      xx = A(1).Cells(i, 1)
      yy = A(2).Cells(i, 1)
    Else
      xx = xy(i, 1): yy = xy(i, 2)
    End If
    If IsNumeric(xx) And IsNumeric(yy) And Not IsEmpty(xx) And Not IsEmpty(yy) Then
      N = 1 + N
      X(N) = xx: y(N) = yy
    End If
  Next i
  ReDim Preserve X(N), y(N)
Else
  N = M
  ReDim X(N), y(N)
  For i = 1 To N
    X(i) = xy(i, 1): y(i) = xy(i, 2)
  Next i
End If
End Sub

Sub Inv2x2(ByVal xx#, ByVal yy#, ByVal xy#, _
  iXX#, iYY#, iXY#, Bad As Boolean)
Dim Determ# ' Invert a symmetric 2x2 matrix
Bad = False
Determ = xx * yy - xy * xy
If Determ = 0 Then Bad = True: Exit Sub
iXX = yy / Determ: iYY = xx / Determ
iXY = -xy / Determ
End Sub

Function AtJay$()
Attribute AtJay.VB_ProcData.VB_Invoke_Func = " \n14"
Dim v, M$
M$ = "J= " & sn$(Jay)
If Jerror Then
  If Jperr Then v = Jperr Else v = Jerror
  M$ = M$ & pm & Sd(v, 3)  'sn$(v)
  If Jperr Then M$ = M$ & "%"
  M$ = M$ & Str(SigLev) & " sigma"
End If
AtJay = M$
End Function
Function Caveats(ByVal Help$)
Attribute Caveats.VB_ProcData.VB_Invoke_Func = " \n14"
Dim r%
With TW.Sheets("Caveats")
  Do
    r = r + 1
    If r = 999 Then MsgBox "Error in locating " & Help$: ExitIsoplot
  Loop Until LCase(.Cells(r, 1)) = LCase(Help$)
  Set Caveats = .Cells(r, 2)
End With
End Function

Sub ProcJ(Optional Init As Boolean = False)
Dim A As Boolean, b As Boolean, s$
A = ArPlat: b = Not A
With Jinput
  If Init Then
    .Caption = IIf(A, "Ar-Ar Plateau Age", "Ar-Ar Isochron")
  End If
  With .eJ
    .Visible = True: .Enabled = b
    .ForeColor = IIf(b, vbBlack, RGB(128, 128, 128))
    .BackColor = IIf(b, vbWhite, Jinput.BackColor)
  End With
    .lJ.Enabled = b 'Visible = b
  If A Then
    .eJ = "": .oKa = Menus("ArKa"): .oMa = Not .oKa
    '.oAbs = False: .oPercent = True
  Else
  End If
  .oAbs.Visible = True: .oPercent.Visible = True
  .oAbs.Enabled = True: .oPercent.Enabled = True
  .oMa.Visible = -1: .oKa.Visible = -1
  .oMa.Enabled = A: .oKa.Enabled = A
  If A Then .oKa = Menus("ArKa"): .oMa = Not .oKa
  s$ = IIf(.oPercent, "percent", "absolute")
  .lJerr = tSt(SigLev) & "-sigma " & s$ & " error in J"
End With
End Sub

Sub CornerRange(Sh As Worksheet, rMax, cMax%, Optional UpprRow = 1, Optional LeftCol = 1)
Dim r&, c%, eR&, eC%
cMax = 0: rMax = 0
For c = LeftCol To 255
  eR = Sh.Cells(65536, c).End(xlUp).Row
  If eR > rMax Then rMax = eR
Next c
For r = UpprRow To rMax
  eC = Sh.Cells(r, 256).End(xlToLeft).Column
  If eC > cMax Then cMax = eC
Next r
End Sub

Sub EtermHandle(ByVal Eterm, ExpTerm#, Bad As Boolean)
Bad = False
If Eterm < (-MAXEXP) Then
 ExpTerm = 0
ElseIf Abs(Eterm) > MAXEXP Then
 Bad = True
Else
  ExpTerm = Exp(Eterm)
End If
End Sub
