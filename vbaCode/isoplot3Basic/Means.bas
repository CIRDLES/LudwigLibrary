Attribute VB_Name = "Means"
' Isoplot module Means
Option Private Module
Option Explicit: Option Base 1
Dim wwA As wWtdAver

Sub WtdAverPlot(ByVal N&, ww As wWtdAver, Wrejected(), Nrej&, Wcap$, _
 Optional LargeErr)
' Construct weighted-averages plot or Probability diagram
Dim i&, j&, k%, nR%, EndCap%, s$, rSlp$, rEr$
Dim DP As Object, ShowStats As Boolean, rAv$, rAvEr$, Regr As Variant, IndX&()
Dim lWeight%, BarThick%, tB As Boolean, TxtBox As Textbox
Dim AvXY As Range, AvErs As Range, rRejXY As Range, RejErs As Range
Dim rRejErs As Range, vRejErs As Variant, tmp$, uosm#
Dim WP As Object, Wsc As Object, BiLine As Range, rWline As Range
Dim MinV#, MaxV#, SpredV#, TikV#, v#, eClr&, eLine&
Dim Wmean#, WmeanErr#, NWS#, WmeanPerr#, AvLineClr&
Dim aWline#(2, 2), aBiLine#(2, 2), aAvXy#(), aAvErs#()
Dim aRejXy#(), aRejErs#(), PrErs#(), tpY#(), Vr
Dim BoxRange As Object, r&, c%, vi&, vv#, e#
Dim Clr&(2), Ptrn#(2), Xoffs#, Mean#, Sigma#
Dim tx1#, tx2#, pY#(), pLab As Variant, pv#(), rRb As Range
Dim rPrt As Range, pLn&, Slp#, Intr#, Nlg%
Dim rLineReg As Range, yMean#, yMeanErr#, SlpErr#
Dim PA As Object, CA As Object, pn&, NlgPlotted%
Dim Tx#(), ty#()
lWeight = AxisLthick ' weight for Y-axis
If WtdAvPlot Then
  Wmean = ww.ChosenMean:    WmeanErr = ww.ChosenErr:    WmeanPerr = ww.ChosenErrPercent
  If Opt.AlwaysPlot2sigma And Wcap$ <> "" Then Wcap$ = Wcap$ & vbLf & "(error bars are 2-sigma)"
  aWline(1, 1) = 0.3 + 0.1 * DoShape: aWline(2, 1) = N + 0.7 - 0.1 * DoShape
  aWline(1, 2) = Wmean:   aWline(2, 2) = Wmean
  SymbRow = Max(1, SymbRow)
  Set rWline = sR(SymbRow, SymbCol, 1 + SymbRow, 1 + SymbCol, ChrtDat)
  AddSymbCol 2
  rWline.Value = aWline
ElseIf AgeExtract Then
  AddSymbCol 2
End If
Set AvXY = sR(SymbRow, SymbCol, N - Nrej - 1 + SymbRow, 1 + SymbCol, ChrtDat)
AddSymbCol 2
Set AvErs = sR(SymbRow, SymbCol, N - Nrej - 1 + SymbRow, SymbCol, ChrtDat)
AddSymbCol 1
ReDim aAvXy(N - Nrej, 2), aAvErs(N - Nrej, 1)
If WtdAvPlot And Nrej > 0 Then
  Set rRejXY = sR(SymbRow, SymbCol, Nrej - 1 + SymbRow, 1 + SymbCol, ChrtDat)
  AddSymbCol 2
  Set RejErs = sR(SymbRow, SymbCol, Nrej - 1 + SymbRow, , ChrtDat)
  AddSymbCol 1
  ReDim aRejXy(Nrej, 2), aRejErs(Nrej, 1)
  ' aRejErs must be dimensioned Nrej,1 (not Nrej) to fill RejErs range correctly
End If
MinV = 1E+99: MaxV = -1E+99: nR = 0
If ProbPlot Then
  Set DP = DlgSht("ProbPlot")
  ShowStats = IsOn(DP.CheckBoxes("cInclStats"))
  ReDim pY(N), PrErs(N)
  If pInpSig Then
    Sigma = EdBoxVal(DP.EditBoxes("eSigma")) / SigLev
    If Not AbsErrs Then Sigma = Sigma / Hun * Mean
  End If
End If
j = 1 - (Not ProbPlot And SigLev = 2 Or Opt.AlwaysPlot2sigma)
i = 0
For k = 1 To N  ' Err bars NOT always shown at 2-sigma
  If InpDat(k, 1) <> 0 Or InpDat(k, 2) <> 0 Then
    i = i + 1
    If WtdAvPlot Or AgeExtract Then aAvXy(i, 1) = k
    aAvXy(i, 2 + ProbPlot) = InpDat(k, 1)
    If WtdAvPlot Or AgeExtract Then aAvErs(i, 1) = InpDat(k, 2) * j
    If ProbPlot Then
      pY(i) = InpDat(k, 1)
      If pBars Then
        If ndCols = 1 Then InpDat(k, 2) = Sigma
      Else
        InpDat(k, 2) = 0
      End If
      PrErs(i) = InpDat(k, 2)
    End If
    MinV = Min(MinV, InpDat(k, 1) - InpDat(k, 2) * j)
    MaxV = Max(MaxV, InpDat(k, 1) + InpDat(k, 2) * j)
  ElseIf WtdAvPlot And Nrej > 0 Then
    nR = 1 + nR
    aRejXy(nR, 1) = k: aRejXy(nR, 2) = Wrejected(k, 1)
    aRejErs(nR, 1) = Wrejected(k, 2) * j
    MinV = Min(MinV, Wrejected(k, 1) - Wrejected(k, 2) * j)
    MaxV = Max(MaxV, Wrejected(k, 1) + Wrejected(k, 2) * j)
  End If
Next k
If ProbPlot Then
  InitIndex IndX(), N
  QuickIndxSort pY(), IndX()
  For i = 1 To N
    aAvXy(i, 1) = pY(i)
    aAvErs(i, 1) = PrErs(IndX(i))
  Next i
  With App
    ReDim tpY(N)
    For i = pFirst To pLast: tpY(i) = pY(i): Next i
    Mean = iAverage(tpY): Sigma = .StDev(tpY)
    pLab = Array(0.01, 0.1, 1, 5, 10, 20, 30, 50, 70, 80, 90, 95, 99, 99.9, 99.99)
    pLn = UBound(pLab)
    ReDim pv(pLn)
    For i = 1 To pLn
      pv(i) = .NormInv(pLab(i) / Hun, Mean, Sigma)
    Next i
    For i = 1 To N: aAvXy(i, 2) = pY(i): Next i
    Erase pY, PrErs, IndX, tpY
    For i = 1 To N
      If i = 1 Then
        uosm = 1 - 0.5 ^ (1 / N)
      ElseIf i = N Then
        uosm = 0.5 ^ (1 / N)
      Else
        uosm = (i - 0.3175) / (N + 0.365)
      End If
      aAvXy(i, 1) = .NormInv(uosm, Mean, Sigma)
    Next i
    If pRegress Then ' Find regression line for the data
      j = pLast - pFirst + 1
      ReDim Tx(j), ty(j)
      j = 0
      For i = pFirst To pLast ' Split out the specified "accepted" points
        If i >= pFirst And i <= pLast Then
          j = 1 + j
          Tx(j) = aAvXy(i, 1)
          ty(j) = aAvXy(i, 2)
        End If
      Next i
      pn = j
      Regr = .LinEst(ty, Tx, , True)
      If pn > 2 Then
        Slp = Regr(1, 1):    Intr = Regr(1, 2)
        SlpErr = Regr(2, 1) * StudentsT(pn - 2)
        yMean = iAverage(ty)
        yMeanErr = StudentsT(pn - 2) * .StDev(ty)
      End If
    End If
  End With
End If
AvXY.Value = aAvXy: LineInd AvXY
If WtdAvPlot Or AgeExtract Then
  AvErs.Value = aAvErs: LineInd AvErs
  If Nrej Then
    rRejXY.Value = aRejXy: RejErs.Value = aRejErs
    LineInd rRejXY:        LineInd RejErs
  End If
End If
SpredV = MaxV - MinV
MinV = MinV - SpredV / 10: MaxV = MaxV + SpredV / 10
Tick SpredV, TikV
MinV = Int(MinV / TikV) * TikV
MaxV = (Int(MaxV / TikV) + 1) * TikV
For i = 1 To 4
  v = MinV - i * TikV
  If NumChars(v) < NumChars(MinV) Then MinV = v
Next i
If ProbPlot Then ' Create x-axis probability-tick range
  Set rRb = sR(SymbRow, SymbCol, pLn - 1 + SymbRow, 1 + SymbCol, ChrtDat)
  Set rPrt = sR(SymbRow, 2 + SymbCol, pLn - 1 + SymbRow, 3 + SymbCol, ChrtDat)
  For i = 1 To pLn
    rRb(i, 1) = pv(i): rRb(i, 2) = MinV
    rPrt(i, 1) = pv(i): rPrt(i, 2) = MaxV
  Next i
  rRb(1 + pLn, 2) = "MinY": rPrt(1 + pLn, 2) = "MaxY"
  AddSymbCol 4
End If
StatBar "Assembling plot"
Charts.Add
If WtdAvPlot Then
  PlotName$ = "Average"
ElseIf ProbPlot Then
  PlotName$ = "ProbPlot"
ElseIf AgeExtract Then
  PlotName$ = "ExtrAge"
ElseIf YoungestDetrital Then
  PlotName$ = "YoungestDetr"
End If
MakeSheet PlotName$, WP
Set IsoChrt = WP
Landscape
If ProbPlot Then
  AxX$ = "Probability"
ElseIf WtdAvPlot Then
  AxY$ = AxX$: AxX$ = ""
ElseIf AgeExtract Then
  AxX$ = "": AxY$ = "Age"
ElseIf PlotName$ = "ExtrAge" Then
  AxX$ = "Age (ma)": AxY$ = ""
End If
Set Vr = AvXY
Ach.ChartWizard Vr, xlXYScatter, 1, xlColumns, 1, 0, 2, "", AxX$, AxY$, ""
Set IsoChrt = Ach ' = WP also
Set CA = WP.ChartArea: Set PA = WP.PlotArea
Set Wsc = WP.SeriesCollection: Wsc(1).Name = "IsoDat1"
CA.Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.SheetClr, vbWhite))
With PA
  If AgeExtract Then
    .Height = 430: .Top = 15: .Width = 570: .Left = 60
  Else
    .Height = 375: .Top = 35: .Width = 460: .Left = 110
  End If
End With
FormatAxes WP, True, MinV, MaxV, TikV
If WtdAvPlot Or AgeExtract Then
  FormatAxes WP, False, 0, 1 + N, 0
Else
  FormatAxes WP, False, pv(1), pv(pLn), 0
End If
With WP
  If WtdAvPlot Or AgeExtract Then
    With .Axes(xlValue)
      .HasMajorGridlines = True
      With .MajorGridlines.Border
        If ColorPlot Then
          .ColorIndex = Menus("iGray50"): .LineStyle = xlContinuous: .Weight = xlThin
        Else
          .Color = Black: .LineStyle = xlDot: .Weight = xlHairline
        End If
      End With
    End With
  End If
  With .PlotArea
    With .Border: .Weight = lWeight: .Color = Black: End With
    .Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.PlotboxClr, vbWhite))
  End With
End With
RemoveHdrFtr WP
If DoShape And (WtdAvPlot Or AgeExtract) Then
  GetScale
  Ptrn(1) = msoPattern60Percent
  If ColorPlot Then
    Clr(1) = RGB(255, 20, 20)
    Clr(2) = RGB(0, 200, 255): Ptrn(2) = 0
  Else
    Clr(1) = RGB(150, 150, 150)
    Clr(2) = Black
    Ptrn(2) = msoPatternLightUpwardDiagonal
  End If
  If AgeExtract Then ' Add best-age bar
    Set BoxRange = Range(ChrtDat.Cells(SymbRow, SymbCol), ChrtDat.Cells(4 + SymbRow, 1 + SymbCol))
    BoxRange(1, 1) = 0.02: BoxRange(1, 2) = ww.BiwtMean
    BoxRange(2, 1) = N + 0.98: BoxRange(2, 2) = BoxRange(1, 2)
    BoxRange(3, 1) = BoxRange(2, 1): BoxRange(3, 2) = ww.ExtMean
    BoxRange(4, 1) = BoxRange(1, 1): BoxRange(4, 2) = BoxRange(3, 2)
    BoxRange(5, 1) = BoxRange(1, 1): BoxRange(5, 2) = BoxRange(1, 2)
    AddSymbCol 2
    LineInd BoxRange, "ErrBox"
    eClr = IIf(ColorPlot, RGB(48, 255, 48), RGB(164, 164, 164))
    AddShape "ErrBox", BoxRange, eClr, 0, False, 0, , , 0.4, -1
    'ActiveSheet.Shapes(ActiveSheet.Shapes.Count).ZOrder msoSendToBack
  End If
  r = SymbRow: c = SymbCol: Xoffs = 0.15
  For j = 1 To 2
    For i = 1 To -(N - nR) * (j = 1) - nR * (j = 2)
      Set BoxRange = Range(ChrtDat.Cells(r, c), ChrtDat.Cells(r + 4, c + 1))
      If j = 1 Then
        vi = aAvXy(i, 1): vv = aAvXy(i, 2): e = aAvErs(i, 1)
      Else
        vi = aRejXy(i, 1): vv = aRejXy(i, 2): e = aRejErs(i, 1)
      End If
      BoxRange(1, 1) = vi - Xoffs:     BoxRange(1, 2) = vv - e
      BoxRange(2, 1) = vi + Xoffs:     BoxRange(2, 2) = BoxRange(1, 2)
      BoxRange(3, 1) = BoxRange(2, 1): BoxRange(3, 2) = vv + e
      BoxRange(4, 1) = BoxRange(1, 1): BoxRange(4, 2) = BoxRange(3, 2)
      BoxRange(5, 1) = BoxRange(1, 1): BoxRange(5, 2) = BoxRange(1, 2)
      LineInd BoxRange, "ErrBox"
      If AgeExtract Then
        tB = (i < ww.ChosenMean Or i > ww.ChosenErr) ' not in the accepted group
        If ColorPlot Then
          eClr = IIf(LargeErr(i), vbWhite, IIf(tB, RGB(64, 64, 255), RGB(192, 0, 0)))
        Else
          eClr = IIf(LargeErr(i), RGB(128, 128, 128), IIf(tB, RGB(208, 208, 208), RGB(50, 50, 50)))
        End If
        eLine = IIf(LargeErr(i), RGB(128, 128, 128), IIf(N > 15, eClr, Black))
      Else
        eClr = Clr(j)
        eLine = IIf(N > 20, eClr, Black)
      End If
      AddShape "ErrBox", BoxRange, eClr, eLine, False, 0, , Ptrn(j)
      r = r + 6
    Next i
  Next j
End If
Wsc(1).MarkerStyle = xlNone
NWS = 1
If ProbPlot Or Not DoShape Then
  With Wsc(1)
    If ProbPlot And pDots Then
      .MarkerStyle = xlCircle
      .MarkerForegroundColor = IIf(ColorPlot, Red, Black)
      .MarkerBackgroundColor = IIf(ColorPlot, White, RGB(192, 192, 192))
      .MarkerSize = 4 - (N < 50) - (N < 40) - (N < 30) - (N < 20) - (N < 10)
    Else
      EndCap = IIf(N < 25, xlCap, xlNoCap)
      Select Case N
        Case Is > 40: BarThick = xlThin
        Case Is > 8:  BarThick = xlMedium
        Case Else:    BarThick = xlThick
      End Select
      .Border.LineStyle = xlNone
      If Sigma > 0 And Not pInpSig And aAvErs(1, 1) = 0 Then
        .ErrorBar xlY, xlBoth, xlCustom, Sigma, Sigma
      Else
        .ErrorBar xlY, xlBoth, xlCustom, aAvErs, aAvErs
      End If
      With .ErrorBars
        With .Border
          .Color = IIf(ColorPlot, Red, Black)
          .Weight = BarThick
        End With
        .EndStyle = EndCap
      End With
    End If
  End With
  If ProbPlot Then
    WP.SeriesCollection.Add rRb, xlColumns, False, 1, False
    WP.SeriesCollection(2).Name = "ProbX"
    WP.SeriesCollection.Add rPrt, xlColumns, False, 1, False
    For j = 1 To 2    ' Add x-axis probability "ticks"
      With WP.SeriesCollection(1 + j) ' Both lower & upper plotbox bounds
        .Border.LineStyle = xlNone: .MarkerStyle = xlPlus: .MarkerSize = 8
        .MarkerBackgroundColorIndex = xlNone: .MarkerForegroundColor = vbBlack
        If j = 1 Then ' Label with vertical numbers
          .ApplyDataLabels Type:=xlShowLabel, LegendKey:=False
          With .DataLabels
            With .Font
              .Name = Opt.AxisTikLabelFont
              .Size = Opt.AxisTikLabelFontSize
              .Background = xlTransparent
            End With
            For i = 1 To pLn: .Item(i).Text = tSt(pLab(i)): Next i
            .Position = xlLabelPositionBelow: .Orientation = xlUpward
          End With
        End If
      End With
    Next j
    If pRegress And pn > 2 Then
      For i = 1 To N
        If i < pFirst Or i > pLast Then ' Non-included ponts -
          With WP.SeriesCollection(1).Points(i)         '  add/replace symbol with "X"
            .MarkerStyle = xlX: .MarkerSize = 6
            .MarkerBackgroundColorIndex = xlNone: .MarkerForegroundColor = vbBlack
          End With
        End If
      Next i
      tx1 = pv(1): tx2 = pv(pLn)
      If (Slp * tx1 + Intr) < MinV Then tx1 = (MinV - Intr) / Slp
      If (Slp * tx2 + Intr) > MaxV Then tx2 = (MaxV - Intr) / Slp
      With ChrtDat
        .Cells(SymbRow, SymbCol) = tx1: .Cells(1 + SymbRow, SymbCol) = tx2
        .Cells(SymbRow, 1 + SymbCol) = Slp * tx1 + Intr
        .Cells(1 + SymbRow, 1 + SymbCol) = Slp * tx2 + Intr
        Set rLineReg = Range(.Cells(SymbRow, SymbCol), .Cells(1 + SymbRow, 1 + SymbCol))
      End With
      AddSymbCol 2
      WP.SeriesCollection.Add rLineReg, xlColumns, False, 1, False
      With Last(WP.SeriesCollection)
        .MarkerStyle = xlNone
        With .Border
          .Weight = xlThin: .LineStyle = xlContinuous
          .Color = IIf(ColorPlot, vbBlue, vbBlack)
        End With
      End With
    End If
    With WP.Axes(xlCategory).AxisTitle: .Top = .Top + 15: End With
    If ShowStats Then ' Add text box showing slope, +-, & mean Y
      NumAndErr Slp, SlpErr, 2, rSlp$, rEr$, , True
      NumAndErr yMean, yMeanErr, 2, rAv$, rAvEr$, , True
      s$ = "Slope = " & rSlp$ & " " & rEr$ & vbLf & _
        "Mean Y = " & rAv$ & " " & rAvEr$ & vbLf & "(95% conf.)"
      AddResBox s$, Clr:=vbWhite, FontSize:=11, Name:="ProbLine", OnChart:=True
      With Ach.TextBoxes("ProbLine")
        .Left = Right_(PA) - .Width - 10
        .Top = Bottom(PA) - .Height - 25
      End With
    End If
  End If
  If WtdAvPlot And Nrej > 0 Then
    WP.SeriesCollection.Add rRejXY, xlColumns, False, 1, False
    If WP.SeriesCollection.Count = NWS + 1 Then
      NWS = 1 + NWS
      With WP.SeriesCollection(NWS)
        .Border.LineStyle = xlNone:  .MarkerStyle = xlNone
        .ErrorBar xlY, xlBoth, xlCustom, RejErs, RejErs
        With .ErrorBars.Border
          .Color = IIf(ColorPlot, Blue, Gray40)
          .Weight = BarThick
        End With
        .ErrorBars.EndStyle = EndCap
      End With
    End If
  End If
End If
If WtdAvPlot Or AgeExtract Then
  If WtdAvPlot Then
    AvLineClr = IIf(Not ColorPlot, Black, IIf(Mac, BrightGreen, Green))
    WP.SeriesCollection.Add rWline, xlColumns, False, 1, False
    If DoShape Or WP.SeriesCollection.Count = NWS + 1 Then
      NWS = 1 + NWS
      WP.SeriesCollection(NWS).MarkerStyle = xlNone
      With WP.SeriesCollection(NWS).Border
        .Color = AvLineClr: .Weight = xlMedium: .LineStyle = xlContinuous
      End With
    End If
  End If
  If Wcap$ <> "" Then
    StatBar "adding textbox to chart"
    Set TxtBox = WP.TextBoxes.Add(0, 0, 0, 0)
    With TxtBox
      With .Characters
        .Text = Wcap$
        With .Font
          .Name = Opt.IsochResFont
          .Size = IIf(AgeExtract, 24, Opt.IsochResFontSize)
        End With
      End With
      If AgeExtract And Left(Wcap$, 8) = "TuffZirc" Then .Characters(1, 8).Font.Italic = True
      IncreaseLineSpace TxtBox, 1.2
      .Interior.ColorIndex = xlNone
      .VerticalAlignment = xlTop:  .HorizontalAlignment = xlCenter
      .Orientation = xlHorizontal: .AutoSize = True
      With .Border
        .LineStyle = xlContinuous: .Color = Black
        .Weight = xlThin
      End With
      .ShapeRange.Shadow.Type = msoShadow6
      .RoundedCorners = True
      .Interior.Color = vbWhite: .Placement = xlMoveAndSize
      .PrintObject = True: .AutoSize = True
    End With
    ConvertSymbols TxtBox
    GetScale
    With TxtBox
      If AgeExtract Then
        .Left = WP.Axes(1).Left + WP.Axes(1).Width / 2 - .Width / 2
        .Top = CA.Height - .Height - 5
      Else
        .Left = PlotBoxLeft + PlotBoxWidth / 2 - .Width / 2
        .Top = PlotBoxBottom - .Height - 10
      End If
    End With
    FreeSpace TxtBox
  End If
End If
If Not AgeExtract Then AddCopyButton
CA.Select
ChrtDat.Visible = False
End Sub

Sub WeightedAverage(ByVal Npts&, ww As wWtdAver, Nrej&, Wrejected(), _
  ByVal CanTukeys As Boolean, Optional OneSigOut = False, _
  Optional IntErr68 = 0, Optional ExtErr68 = 0)
Dim i&, j&, Count&, nU&
Dim N0&, Nn&, Weight#, SumWtdRatios#
Dim q#, IntMean#, ExtSigma#, t68#, t95#, WtdAvg#, WtdAvgErr#
Dim IntSigmaMean#, ExtSigmaMean#, TotSigmaMean#, PointError#, TotalError#
Dim Tolerance#, Sums#, InverseVar#(), temp#, Tot1sig#, IntVar#
Dim ExtVar#, WtdR2#, TotVar#, Trial#, Resid#, WtdResid#
Dim yy#(), IvarY#(), tbx#(), r$(4), NoEvar As Boolean, ExtXbarSigma#
Nn = Npts: Count = 0: Nrej = 0
If MinProb = 0 Then MinProb = Val(Menus("MinProb"))
ReDim InverseVar(Nn), yy(Nn), IvarY(Nn), tbx(Nn), yf.WtdResid(Nn)
If CanReject Then ReDim Wrejected(Nn, 2)
For i = 1 To Nn
  InverseVar(i) = 1 / SQ(InpDat(i, 2))
Next i
Recalc:
ExtSigma = 0:  ww.Ext2Sigma = 0: Weight = 0
SumWtdRatios = 0: q = 0: Count = Count + 1
For i = 1 To Npts
  If InpDat(i, 1) <> 0 And InpDat(i, 2) <> 0 Then
    Weight = Weight + InverseVar(i)
    SumWtdRatios = SumWtdRatios + InverseVar(i) * InpDat(i, 1)
    q = q + InverseVar(i) * SQ(InpDat(i, 1))
  End If
Next i
nU = Nn - 1 ' Deg. freedom
t68 = StudentsT(nU, 68.26)
t95 = StudentsT(nU, 95)
IntMean = SumWtdRatios / Weight  ' "Internal" error of wtd average
ww.IntMean = IntMean             ' Use double-prec. var.in calcs!
'Sums = q - Weight * Sq(IntMean) ' Sums of squares of weighted deviates
'If Sums < 0 Then                ' May be better in case of roundoff probs
  Sums = 0
  For i = 1 To Npts
    If InpDat(i, 1) <> 0 And InpDat(i, 2) <> 0 Then
      Resid = InpDat(i, 1) - IntMean  ' Simple residual
      WtdResid = Resid / InpDat(i, 2) ' Wtd residual
      WtdR2 = SQ(WtdResid)            ' Square of wtd residual
      If Nn = Npts Then yf.WtdResid(i) = WtdResid 'WtdR2
      Sums = Sums + WtdR2
    End If
  Next i
  Sums = Max(Sums, 0)
'End If
With ww
  .MSWD = Sums / nU  ' Mean square of weighted deviates
  IntSigmaMean = Sqr(1 / Weight)
  .IntMeanErr2sigma = 2 * IntSigmaMean
  TotSigmaMean = IntSigmaMean * Sqr(.MSWD)
  .Probability = ChiSquare(.MSWD, (nU))
  .IntMeanErr95 = IntSigmaMean * IIf(.Probability >= 0.3, 1.96, t95 * Sqr(.MSWD))
  If OneSigOut Then
    IntErr68 = IntSigmaMean * IIf(.Probability >= 0.3, 0.9998, StudentsT(nU, 68.26) * Sqr(.MSWD))
  End If
End With
If ww.Probability < MinProb And ww.MSWD > 1 Then
  'Find the MLE constant external variance
  Nn = 0
  For i = 1 To Npts
    If InpDat(i, 1) <> 0 Then
     Nn = 1 + Nn: yy(Nn) = InpDat(i, 1)
     IvarY(Nn) = SQ(InpDat(i, 2))
  End If
  Next i
  ReDim Preserve yy(Nn), IvarY(Nn)
  IntVar = Nn / Weight        ' Mean internal variance of pts
  TotVar = IntVar * Sums / nU ' Mean total variance of pts
  WtdExtRTSEC 0, 10 * SQ(IntSigmaMean), ExtVar, ww.ExtMean, ExtXbarSigma, yy(), IvarY(), (Nn), NoEvar
  With ww
    If Not NoEvar Then
      ExtSigma = Sqr(ExtVar)
      ' Knowns: N of the x(i) plus N of the SigmaX(i)
      ' Unknowns: Xbar abd ExtVar
      .ExtMeanErr95 = StudentsT(2 * Nn - 2) * ExtXbarSigma
      If NIM(ExtErr68) Then
        ExtErr68 = StudentsT(2 * Nn - 2, 68.26) * ExtXbarSigma
      End If
      .Ext2Sigma = 2 * ExtSigma
    ElseIf .MSWD > 4 Then  ' Failure of RTSEC algorithm because of extremely high MSWD
      ExtSigma = App.StDev(yy)
      .ExtMean = iAverage(yy)
      .ExtMeanErr95 = t95 * ExtSigma / Sqr(Nn)
      If NIM(ExtErr68) Then
        ExtErr68 = StudentsT(nU, 68.26) * ExtSigma
      End If
      .Ext2Sigma = 2 * ExtSigma
    Else
      .ExtMean = 0: .ExtMeanErr95 = 0: .Ext2Sigma = 0
      If NIM(ExtErr68) Then ExtErr68 = 0
    End If
    .ExtMeanErr68 = t68 / t95 * .ExtMeanErr95
  End With
End If
If CanReject And ww.Probability < MinProb Then GoSub Reject
If CanTukeys Then
  For i = 1 To Npts: tbx(i) = InpDat(i, 1): Next i
  With ww
    TukeysBiweight tbx(), Npts, 6, .BiwtMean, .BiWtSigma, .BiWtErr95
    .Median = iMedian(tbx()): .MedianConf = MedianConfLevel(Npts)
    .MedianPlusErr = MedianUpperLim(tbx) - .Median
    .MedianMinusErr = .Median - MedianLowerLim(tbx)
  End With
End If
Exit Sub

Reject:
If ww.Ext2Sigma Then
  WtdAvg = ww.ExtMean
  WtdAvgErr = IIf(OneSigOut, ExtErr68, ww.ExtMeanErr95)
Else
  WtdAvg = ww.IntMean
  WtdAvgErr = IIf(OneSigOut, ExtErr68, ww.IntMeanErr95)
End If
N0 = Nn   ' Reject outliers
For i = 1 To Npts
  If InpDat(i, 1) <> 0 And Nn > 0.85 * Npts Then '0.7 * Npts Then ' Reject no more than 30% of ratios
  ' Start rej. tolerance at 2-sigma, increase slightly each pass.
    PointError = 2 * Sqr(SQ(InpDat(i, 2)) + ExtSigma ^ 2)
    ' 2-sigma error of point being tested
    TotalError = Sqr(PointError ^ 2 + (2 * ww.ExtMeanErr68) ^ 2)
    'Tolerance = (1 + (Count - 1) / 8) * TotalError
    ' 1st-pass tolerance is 2-sigma; 2nd is 2.25-sigma; 3rd is 2.5-sigma..
    Tolerance = (1 + (Count - 1) / 4) * TotalError
    If HardRej Then Tolerance = Tolerance * 1.25
    ' 1st-pass tolerance is 2-sigma; 2nd is 2.5-sigma; 3rd is 3-sigma...
    q = InpDat(i, 1) - WtdAvg
    If Abs(q) > Tolerance And Nn > 2 Then '!!
      Nrej = 1 + Nrej: Nn = Nn - 1
      For j = 1 To 2
        Wrejected(i, j) = InpDat(i, j)
        InpDat(i, j) = 0
      Next j
    End If
  End If
Next i
If Nn < N0 Then GoTo Recalc
Return
End Sub

Function WtdExtFunc(ByVal ExtVar#, Xbar#, XbarSigma#, _
  X#(), IntVar#(), ByVal N%)
' WtdExtFunc will be zero when the external variance, ExtVar, is chosen correctly (N>=3)
Dim i%, ff#, SumW#, SumXW#, W#()
Dim SumW2resid2#, SumXW2#, Resid# ',SumW2#
ReDim W(N)
For i = 1 To N
  W(i) = 1 / (IntVar(i) + ExtVar)
  SumW = SumW + W(i)
  SumXW = SumXW + X(i) * W(i)
  'SumW2 = SumW2 + SQ(W(i))
Next i
Xbar = SumXW / SumW
For i = 1 To N
  Resid = X(i) - Xbar
  SumW2resid2 = SumW2resid2 + SQ(W(i) * Resid)
Next i
ff = SumW2resid2 - SumW
'ExtVarErr = Sqr(Abs(1 / (0.5 * SumW2))) ' 1-sigma error in external variance
XbarSigma = Sqr(Abs(1 / SumW)) ' 1-sigma error in Xbar
WtdExtFunc = ff
End Function

Sub WtdExtRTSEC(ByVal ExtVar1#, ByVal ExtVar2#, ExtVar#, _
  Xbar#, XbarSigma#, X#(), IntVar#(), _
  N%, Failed As Boolean)
'  Using the secant method, find the root of a function WtdExtFunc thought to lie between ExtVar1 and ExtVar2.
'  The root, returned as WtdExtRTSEC, is refined until its accuracy is +-xacc.
' Press et al, 1987, p. 250-251.
Dim j%, f#, rts#, dx#, xacc#, Lastf1#, Lastf2#
Dim FL#, xl#, facc#, tmp#, rct%
Const Maxit = 99, MaxD = 100

xacc = 0.000000001: facc = 0.0000001: Failed = False
FL = WtdExtFunc(ExtVar1, Xbar, XbarSigma, X(), IntVar(), N)
f = WtdExtFunc(ExtVar2, Xbar, XbarSigma, X(), IntVar(), N)

On Error GoTo FailedExit
If CSng(f) = CSng(FL) Then GoTo FailedExit
On Error GoTo 0
If Abs(FL) < Abs(f) Then ' Pick the bound with the smaller function value as
  rts = ExtVar1: xl = ExtVar2      '  the most recent guess.
  Swap FL, f
Else
  xl = ExtVar1
  rts = ExtVar2
End If
For j = 1 To Maxit               ' Secant loop.
  If f = FL Then GoTo Succeeded
  dx = (xl - rts) * f / (f - FL) ' Increment wrt latest value.
  xl = rts
  FL = f
  rct = 0
  Do
    tmp = rts + dx
  If tmp >= 0 Then
    Exit Do
  ElseIf rct > MaxD Then
    GoTo FailedExit
  End If
  dx = dx / 2
  rct = 1 + rct
  If rct > 99 Then GoTo FailedExit
  Loop
  rts = tmp
  f = WtdExtFunc(rts, Xbar, XbarSigma, X(), IntVar(), N)
  ' Done if f<facc or f is oscillating
  If Abs(f) < facc Or Abs(f) = Abs(Lastf2) Then GoTo Succeeded
  Lastf2 = Lastf1: Lastf1 = f
Next j

Succeeded: ExtVar = rts: Exit Sub

FailedExit: Failed = True
End Sub

Public Function Ernd(ByVal Value#, PlusOrMinus#, Optional Short = False)
Attribute Ernd.VB_Description = "Round a number with specified uncertainty to a reasonable # of significant figures ('Short'=TRUE for fewer)"
Attribute Ernd.VB_ProcData.VB_Invoke_Func = " \n14"
' Given a numeric value (Value) & its absolute uncertainty (PlusOrMinus)
'   return Value rounded to a reasonable number of significant figures.
Dim d%, RoundedValue#
ViM Short, False
If Short <> False Then Short = True
PlusOrMinus = Abs(PlusOrMinus)
RoundedValue = Drnd(Value, 7)
If PlusOrMinus > 0 And Abs(Value) > 0 Then
  d = zz(PlusOrMinus) - 2 - Short
  If d > -8 Then RoundedValue = Prnd(Value, d)
End If
Ernd = RoundedValue
End Function

Sub ShowWtdAv(W As wWtdAver, ByVal N&, ByVal Nrej&, T$)
Attribute ShowWtdAv.VB_ProcData.VB_Invoke_Func = " \n14"
' Create the dialog box for Weighted-Averages results
Dim o As Object, L As Labels, s As Object, ML As Object, tB As TextBoxes, i&
Dim b As Boolean, tmp$, tmp1$, er1$, er2$, Op As Object, G As GroupBoxes, c As Object, M$, pr$
Dim t95$, HavExt As Boolean, BWE As Boolean, rj$
AssignD "WtdAv", s, , c, Op, L, G, tB
wwA = W
c("ShowWithPlot").Enabled = DoPlot
For i = 1 To tB.Count: tB(i).Visible = True: Next
For i = 1 To G.Count: G(i).Visible = True: Next
If True Then 'Not MacExcelX Then
  For i = 1 To tB.Count
    With tB(i)
      With .Font
        .Name = IIf(Mac, "Geneva", "Arial")
        .Bold = Windows: .Size = 9 + Mac
        .Italic = False: .ColorIndex = xlAutomatic
      End With
      .AutoSize = (.Name <> "IntConf95")
    End With
  Next i
End If
t95$ = "95% conf."
With W
  HavExt = (.Probability < 0.3 And .ExtMean <> 0)
  If Not Mac Then tB("ShowWdata").Font.Bold = True
  G("Rejected").Enabled = True 'CanReject
  L("Nrej").Visible = True 'CanReject
  L("Nrej").Text = sn$(Nrej) & " of" & Str(N)
  BWE = (N > 4)
  G("gBiWt").Enabled = BWE: L("BiWtMean").Enabled = BWE
  G("gMedian").Enabled = BWE: L("Median").Enabled = BWE
  tB("tBiWtConf95").Visible = BWE
  tB("tMedianConf").Text = "(" & tSt(.MedianConf) & "% conf.)"
  tB("tMedianConf").Visible = BWE
  L("IntMean2sigma").Text = VandE(.IntMean, .IntMeanErr2sigma, 2, , , True)
  L("MSWD").Text = "MSWD = " & Mrnd(.MSWD)
  pr$ = ProbRnd(.Probability)
  L("Probability").Text = "Probability of fit = " & pr$
  L("IntMean2sigma").Enabled = True
  With tB("tInternal2sigma"):
    .Text = "2-sigma internal"
    If Not Mac Then .Font.Color = Black
  End With
  With tB("tExtConf95"):
    .Enabled = (HavExt And Not MacExcelX)
    .Visible = HavExt
  End With
  G("gExternal").Enabled = HavExt
  With Op("oShowBiWt"):   .Visible = False: .Enabled = False: End With
  With Op("oShowMedian"): .Visible = False: .Enabled = False: End With
  If BWE Then
    L("BiWtMean").Text = VandE(.BiwtMean, .BiWtErr95, 2, , , True)
    NumAndErr .Median, Min(.MedianPlusErr, .MedianMinusErr), 2, tmp$, ""
    er1$ = ErFo(.Median, .MedianPlusErr, 2): er2$ = ErFo(.Median, .MedianMinusErr, 2)
    tmp1$ = "   [" & ErFo(.Median, (.MedianPlusErr + .MedianMinusErr) / 2, 2, , True) & "]"
    L("Median").Text = tmp$ & "  +" & er1$ & "/-" & er2$ & tmp1$
    With Op("oShowBiWt"): .Visible = True: .Enabled = True: End With
    With Op("oShowMedian"): .Visible = True: .Enabled = True: End With
  Else
    If IsOn(Op("oShowBiWt")) Then Op("oShowInternal95") = xlOn
    Op("oShowBiWt") = xlOff: Op("oShowMedian") = xlOff
  End If
  If Not HavExt And IsOn(Op("oShowExternal")) Then
    Op("oShowInternal95") = xlOn: Op("oShowExternal") = xlOff
  End If
  If .Probability >= 0.3 Then
    G("gExternal").Enabled = False
    L("IntMean95").Text = VandE(.IntMean, .IntMeanErr2sigma / 2 * 1.96, 2, , , True)
    tB("tIntConf95").Text = t95$ & "  (=1.96sigma)"
    L("ExtMean").Visible = False: L("ExtErr").Visible = False
    With Op("oShowExternal"): .Enabled = False: .Visible = False: End With
    tB("tExtConf95").Visible = False
    tB("tIntConf95").Visible = True
    tB("tIntConf95").Enabled = False
    tB("tInternal2sigma").Enabled = True
  Else
    With tB("tIntConf95"): .Visible = True: .Enabled = True: End With
    tB("tInternal2sigma").Enabled = (.Probability > 0.01)
    tB("tInternal2sigma").Visible = True
    Op("oShowExternal").Enabled = HavExt: Op("oShowExternal").Visible = HavExt
    G("gExternal").Visible = True:
    If HavExt Then
      L("ExtMean").Text = VandE(.ExtMean, .ExtMeanErr95, 2, , , True)
      L("ExtErr").Text = "External 2-sigma err req'd (each pt) = " & ErFo(.ExtMean, .Ext2Sigma, 2) _
        & "  [" & pm & ErFo(.ExtMean, .Ext2Sigma, 2, , True) & "]"
    End If
    With tB("tIntConf95")
      .Text = t95$ & "  (=tsigmasqrtMSWD)"
      If Not Mac Then .Characters(14, 1).Font.Italic = True
    End With
    L("IntMean95").Text = VandE(.IntMean, .IntMeanErr95, 2, , , True)
    L("ExtMean").Visible = HavExt: L("ExtErr").Visible = HavExt
    If .Probability < 0.05 Then
      L("IntMean2sigma").Enabled = False
      If IsOn(Op("oShowInternal2sigma")) Then
        Op("oShowInternal95") = xlOn: Op("oShowInternal2sigma") = xlOff
      End If
      Op("oShowInternal2sigma").Enabled = False
      If Not Mac Then tB("tInternal2sigma").Font.Color = Gray50
    End If
  End If
  ShowWtdDatClick
  ShowBox s, True
  rj$ = tSt(Nrej) & " of" & Str(N) & " rej."
  T$ = "Mean = ": tmp$ = vbLf & "Wtd by data-pt errs only, " & rj$
  If IsOn(Op("oShowInternal2sigma")) Then  ' Assigned-error weighted only
    T$ = T$ & L("IntMean2sigma").Text & "  2-sigma" & tmp$
    .ChosenMean = .IntMean: .ChosenErr = .IntMeanErr2sigma
  ElseIf IsOn(Op("oShowInternal95")) Then
    T$ = T$ & L("IntMean95").Text & "  " & t95$ & tmp$
    .ChosenMean = .IntMean: .ChosenErr = .IntMeanErr95
  ElseIf IsOn(Op("oShowExternal")) Then
    T$ = T$ & L("ExtMean").Text & "  " & t95$
    T$ = T$ & vbLf & "Wtd by data-pt + ext. errs, " & rj$
    T$ = T$ & vbLf & L("ExtErr").Text
    .ChosenMean = .ExtMean
    .ChosenErr = .ExtMeanErr95
  ElseIf IsOn(Op("oShowBiWt")) Then
    T$ = "Tukey's Biweight Mean = " & vbLf
    T$ = T$ & L("BiWtMean").Text & "  " & t95$
    .ChosenMean = .BiwtMean: .ChosenErr = .BiWtErr95
  ElseIf IsOn(Op("oShowMedian")) Then
    T$ = "Median = " & vbLf & L("Median").Text & "  " & tB("tMedianconf").Text
    .ChosenMean = .Median: .ChosenErr = (.MedianPlusErr + .MedianMinusErr) / 2
  End If
End With
W.ChosenErrPercent = W.ChosenErr / Hun * W.ChosenMean
If Op("oShowBiWt") <> xlOn And Op("oShowMedian") <> xlOn Then _
  T$ = T$ & vbLf & L(1).Text & ", probability = " & pr$
If IsOn(c("ShowWithData")) Then
  AddResBox T$
End If
If IsOff(c("ShowWithPlot")) Or Not DoPlot Then T$ = ""
End Sub

Sub SimpleWtdRegression(ByVal N&, X#(), y#(), SigmaY#(), _
  Slope#, SlopeErr#, PoorFit As Boolean)
' Linear regression weighted for Y errors only
Dim i&, Inter#, sv#, svX#, sVY#, sVXY#, sVX2#
Dim s#, MSWD#, Prob#, SlopeVar#, df&, v#()
ReDim v(N)
PoorFit = False: SlopeErr = 0
For i = 1 To N: v(i) = 1 / SQ(SigmaY(i)): Next i
' Solve eq'ns to minimize sums of wtd (Y-resids)^2 while also finding 2nd derivs of S w.r.t. slope & inter.
sv = Sum(v)
svX = SumProduct(v, X):     sVY = SumProduct(v, y)
sVXY = SumProduct(v, X, y): sVX2 = SumProduct(v, X, X)
Slope = (sv * sVXY - svX * sVY) / (sv * sVX2 - svX * svX)
Inter = (sVY - Slope * svX) / sv
' Elements of Fisher Information Matrix are:
'   sV   sVX
'   sVX  sVX2
For i = 1 To N ' Sums of squares of wtd Y-resids
  s = s + v(i) * SQ(y(i) - Slope * X(i) - Inter)
Next i
df = N - 2
If df > 0 Then
  MSWD = s / df
  Prob = ChiSquare(MSWD, df)
Else
  MSWD = 0: Prob = 1
End If
SlopeVar = sv / (sv * sVX2 - svX * svX) ' =InverseI(2,2)
SlopeErr = Sqr(SlopeVar)
If Prob < 0.05 Then PoorFit = True
' Expand to 95%-conf.error
If PoorFit Then SlopeErr = SlopeErr * Sqr(MSWD) * StudentsT(N - 2) / 2
End Sub

Private Sub ShowWtdDatClick()
Dim Op As Object, G As Object, c As Object, tB As Object, b As Boolean
AssignD "WtdAv", , , c, Op, , G, tB
b = (IsOn(c("ShowWithData")) Or IsOn(c("ShowWithPlot")))
tB("ShowWdata").Visible = b
Op("oShowInternal2sigma").Enabled = (b And wwA.Probability >= 0.05)
Op("oShowInternal95").Enabled = b
Op("oShowExternal").Enabled = (b And wwA.Probability < 0.3 And wwA.ExtMean <> 0)
Op("oShowBiWt").Enabled = (b And N > 4)
End Sub

Function AreaUnderCurve(xy#(), ByVal N)
Attribute AreaUnderCurve.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, A#, DeltaX#, AvY#
For i = 1 To N - 1 ' Numerical integration of area under a curve defined by xy()
  DeltaX = xy(i + 1, 1) - xy(i, 1)
  AvY = (xy(i, 2) + xy(i + 1, 2)) / 2
  A = A + DeltaX * AvY
Next i
AreaUnderCurve = A
End Function

Sub BinSpec(ByVal EstNbins) ' Make sure #bins for histoplot is reasonable
Attribute BinSpec.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s
s = Array("few", "many")
If EstNbins < 2 Or EstNbins > 1000 Then
  MsgBox "Can't proceed with histogram plot -- too " & _
    s(2 + (EstNbins < 2)) & " bins specified.", , Iso
  ExitIsoplot
End If
Nbins = EstNbins
End Sub

Sub FormatHistoProbAxis(ByVal HistoProb%, ByVal AxisNum%, ByVal ChartAsSheet As Boolean, _
  ByVal MaxBin%, Optional DoMix As Boolean = False, Optional AddToPlot As Boolean = False)
Attribute FormatHistoProbAxis.VB_ProcData.VB_Invoke_Func = " \n14"
ViM DoMix, False
ViM AddToPlot, False
With ActiveChart.Axes(xlValue, AxisNum)
  Select Case HistoProb ' Histogram #samples per bin
    Case 1
      .TickLabelPosition = xlNextToAxis
      .MajorTickMark = xlCross: .MinorTickMark = xlInside
      .MaximumScale = Max(1 + MaxBin, Int((1.3 * MaxBin * 5) / 5))
       With .TickLabels.Font
         .Name = Opt.AxisTikLabelFont
         .Size = IIf(DoMix And Not ChartAsSheet, 17, Opt.AxisTikLabelFontSize)
         .Background = xlTransparent
       End With
      .MajorUnit = Max(1, .MajorUnit)
      If .MajorUnit = 1 And .MinorUnit < 1 Then
        .MinorTickMark = xlNone
      Else
        .MinorUnit = Max(1, .MinorUnit)
      End If
      .TickLabels.NumberFormat = "0"
      .HasTitle = Not DoMix
    Case 2 ' Cumulative probability
      .MajorTickMark = xlNone: .MinorTickMark = xlNone
      .TickLabelPosition = xlNone: .HasTitle = Not DoMix
  End Select
  If .HasTitle Then
    With .AxisTitle
      .Text = IIf(HistoProb = 1, "Number", "Relative probability")
      .Orientation = IIf(AxisNum = 2, xlDownward, xlUpward)
      With .Characters.Font
        .Name = Opt.AxisNameFont: .Size = Opt.AxisNameFontSize
        .Background = xlTransparent
      End With
    End With
  End If
  .MinimumScale = 0
End With
End Sub

Sub WeightedAv(W(), ValuesErrs, Optional PercentOut = False, Optional PercentIn = False, _
  Optional SigmaLevelIn = 2, Optional CanRej = False, Optional ConstExtErr = False, _
  Optional AltRej = False, Optional SigmaLevelOut = 2)
' Array function to calculate weighted averages.  ValuesErrs is a 2-column range
'  with values & errors, PercentOut specifies whether output errors are absolute
'  (default) or %; PercentIn specifies input errors as absolute (default) or %;
'  SigmaLevelIn is sigma-level of inputerrors (default is 2-sigma); OUTPUT ALWAYS 2-SIGMA.
'  CanRej permits rejection of outliers (default is False); If Probability<15%,
'  ConstExtErr specifies weighting by assigned errors plus a constant-external error
'  (default is weighting by assigned errors only).
' Errors are expanded by t-sigma-Sqrt(MSWD) if probability<10%.
' Output is a 2 columns of 5 (ConstExtErr=FALSE) or 6 (ConstExtErr=TRUE) rows, where
'  the left column contains values & the right captions.
Dim i&, Nrej&, j&, nR&, v#(), wt#(), IntErr68#, ExtErr68#
Dim Sigma#, v1#, v2#, tB As Boolean, Wrejected(), Nareas%, Co%
Dim StrkThru As Boolean, ww As wWtdAver, SL$(2), Cnf$(2), rj$, NoColZero As Boolean
ViM PercentOut, False
ViM PercentIn, False
ViM SigmaLevelIn, 2
ViM CanRej, False
ViM ConstExtErr, False
ViM AltRej, False
If SigmaLevelIn <= 0 Then SigmaLevelIn = 2
NoColZero = False: CanReject = CanRej
If TypeName(ValuesErrs) = "Range" Then
  Nareas = ValuesErrs.Areas.Count
Else
  Nareas = 1
End If
If AltRej Then ' If AltRej AND can't read font attribs, rejected values are indicated
  i = 99       '  by "rej" in column to left of values; otherwise by StrkThru.
  On Error Resume Next
  i = ValuesErrs(1, 1).Font.Strikethrough
  On Error GoTo 0
  If i <> 99 Then AltRej = False
End If
If AltRej Then   ' Is values-column in column A (so no "rej" possible)?
  i = 9
  On Error Resume Next
  i = IsNumeric(ValuesErrs(1, 0))
  On Error GoTo 0
  NoColZero = (i = 9)
End If
SigLev = 2: AbsErrs = True
nR = ValuesErrs.Rows.Count
ReDim v(nR, 2)
SL$(1) = "1-sigma": SL$(2) = "2-sigma"
Cnf$(1) = "68%-conf.": Cnf$(2) = "95%-conf."
N = 0
For i = 1 To nR
  v1 = Val(ValuesErrs(i, 1))
  v2 = Val(ValuesErrs.Areas(Nareas)(i, IIf(Nareas = 2, 1, 2)))
  StrkThru = False
  On Error Resume Next
  If AltRej Then
    If Not NoColZero Then StrkThru = (ValuesErrs(i, 0) = "rej")
  Else
    StrkThru = ValuesErrs(i, 1).Font.Strikethrough
  End If
  On Error GoTo 0
  If IsNumeric(v1) And IsNumeric(v2) Then
    If v1 <> 0 And v2 <> 0 And Not StrkThru Then
      N = 1 + N
      Sigma = v2 / SigmaLevelIn
      If PercentIn Then Sigma = Abs(Sigma / Hun * v1)
      v(N, 1) = v1: v(N, 2) = Sigma
    End If
  End If
Next i
ReDim wt(N), Wrejected(N, 2), InpDat(N, 2)
For i = 1 To N
  InpDat(i, 1) = v(i, 1): InpDat(i, 2) = v(i, 2)
Next i
Erase v
WeightedAverage N, ww, Nrej, Wrejected(), False, (SigmaLevelOut = 1), IntErr68, ExtErr68
With ww
  W(1, 1) = .IntMean: W(1, 2) = "Wtd Mean (from internal errs)"
  If ConstExtErr Then
    W(3, 2) = "Required external " & SL$(SigmaLevelOut)
    W(3, 1) = 0
  End If
  If .Probability > 0.05 Then
    W(2, 1) = .IntMeanErr2sigma / 2 * SigmaLevelOut
    W(2, 2) = SL$(SigmaLevelOut)
  ElseIf ConstExtErr Then
    W(1, 1) = .ExtMean
    W(1, 2) = "Wtd Mean (using ext. err)"
    W(2, 1) = Choose(SigmaLevelOut, .ExtMeanErr68, .ExtMeanErr95)  'ExtErr68, .ExtMeanErr95)
    W(2, 2) = Cnf$(SigmaLevelOut)
    W(3, 1) = .Ext2Sigma / 2 * SigmaLevelOut
  Else
    W(2, 1) = Choose(SigmaLevelOut, IntErr68, .IntMeanErr95)
    W(2, 2) = Cnf$(SigLev)
  End If
End With
W(2, 2) = W(2, 2) & " err. of mean"
If PercentOut Then
  W(2, 1) = Abs(W(2, 1) / W(1, 1) * Hun)
  If ConstExtErr Then W(3, 1) = Abs(W(3, 1) / W(1, 1) * Hun)
  W(2, 2) = W(2, 2) & " (%)"
  For i = 2 To 3 + ConstExtErr
    j = InStr(W(i, 2), "sigma")
    If j Then W(i, 2) = Left$(W(i, 2), j - 1) & "sigma%" & Mid$(W(i, 2), j + 5)
  Next i
End If
W(3 - ConstExtErr, 1) = ww.MSWD:        W(3 - ConstExtErr, 2) = "MSWD"
W(4 - ConstExtErr, 1) = Nrej:           W(4 - ConstExtErr, 2) = "rejected"
W(5 - ConstExtErr, 1) = ww.Probability: W(5 - ConstExtErr, 2) = "Probability of fit"
If CanReject And UBound(W, 1) >= j Then
  j = 6 - ConstExtErr: rj$ = ""
  If Nrej Then      ' Create space-delimited string containing index #s
    For i = 1 To N  '  of the rejected points.
      If Not IsEmpty(Wrejected(i, 1)) Then rj$ = rj$ & Str(i)
    Next i
    W(j, 1) = rj$: W(j, 2) = "rej. item #(s)"
  Else
    W(j, 1) = "": W(j, 2) = ""
  End If
End If
End Sub

Sub WtdAvCorr(v#(), VarCov#(), ByVal N&, Vbar#, SigmaVbar#, _
  MSWD#, Prob#, SigRho As Boolean, Bad As Boolean)  ' INCORRECT MSWD EQUATION
Attribute WtdAvCorr.VB_ProcData.VB_Invoke_Func = " \n14"
' Weighted average of a single variable (V) whose values are correlated amongst themselves.
' If SigRho is True, then VarCov() contains sigma's and rho's instead of variances & covariances.
Dim i&, j&, Numer#, Denom#, Sums#, OMij#
Dim OmegaInv As Variant, Omega As Variant
ReDim OmegaInv(N, N) As Variant
For i = 1 To N ' Construct variance-covariance matrix
  For j = 1 To N
    If Not SigRho Then
      OmegaInv(i, j) = VarCov(i, j)
    ElseIf i = j Then ' convert sigma's to variances
      OmegaInv(i, i) = SQ(VarCov(i, i))
    Else              ' convert rho's to covariances
      OmegaInv(i, j) = VarCov(i, i) * VarCov(j, j) * VarCov(i, j)
    End If
Next j, i
Omega = App.MInverse(OmegaInv)
If IsError(Omega) Then Exit Sub
For i = 1 To N
  For j = 1 To N
    OMij = Omega(i, j)
    Numer = Numer + (v(i) + v(j)) * OMij
    Denom = Denom + OMij
    Sums = Sums + v(i) * v(j) * OMij ' Obviously INCORRECT
Next j, i
If Denom <= 0 Then Bad = True: Exit Sub
Vbar = Numer / Denom / 2
SigmaVbar = Sqr(1 / Denom)
MSWD = Sums / (N - 1) ' Obviously INCORRECT
Prob = ChiSquare(MSWD, N - 1)
End Sub

Sub PutMedian()
Attribute PutMedian.VB_ProcData.VB_Invoke_Func = " \n14"
' Places the median and upper/lower conf. limits on the selected values using Rock et al's method
' (Rock, Webb, McNaughton, & Bell, Chem Geol 1987, 66, 163-177)
Dim i&, N&, j&, Lrow&, Oa$, Nn$
Dim v As Range, o As Range, Va$, W$, vv(), Bad As Boolean
Set v = Selection
If v.Areas.Count > 1 Then MsgBox "Input range must be contiguous", , Iso: KwikEnd
If v.Rows.Count = EndRow Then Set v = SelectFromWholeCols(v)
Lrow = v.Row + v.Rows.Count + 1
j = v.Column: qq = Chr(34)
Set o = sR(Lrow, j, Lrow + 3, j + 1)
For i = 1 To o.Count
  If Not IsEmpty(o(i)) Then
    NoUp False
    o.Select
    If MsgBox("Output will overwrite selected area", vbOKCancel, Iso) = vbOK _
      Then Exit For Else v.Select: KwikEnd
  End If
Next i
NoUp
ConvertRange v, vv(), Bad
If Bad Then MsgBox "Nothing numeric in input range", , Iso: ExitIsoplot
If v.Rows.Count > UBound(vv) Then MsgBox "Input range must be entirely numeric", , Iso: ExitIsoplot
If v.Rows.Count < 3 Then MsgBox "3 or more values required", , Iso: ExitIsoplot
Va$ = v.Address: Oa$ = o(1, 1).Address: W$ = TW.Name & "!"
Nn$ = "COUNT(" & Va$ & ")"
On Error GoTo Bad
o(1, 2) = "Median":  o(1, 1) = "= MEDIAN(" & v.Address & ")"
On Error GoTo 0
o(2, 2) = "+error": o(3, 2) = "-error"
o(2, 1) = "=" & W$ & "MedianUpperLim(" & Va$ & "," & Nn$ & ")-" & Oa$
o(3, 1) = "=" & Oa$ & "-" & W$ & "MedianlowerLim(" & Va$ & "," & Nn$ & ")"
o(4, 1) = "=CONCATENATE(" & W$ & "MedianConfLevel(" & Nn$ & "), " & qq & "% conf." & qq & ")"
Range(o(1, 1), o(3, 1)).NumberFormat = v.NumberFormat
o.Font.Size = v.Font.Size
For i = 1 To 3
  HA o(i, 1), xlRight
  HA o(i, 2), xlLeft
Next i
HA o(4, 1), xlLeft
v.Select
Exit Sub
Bad: If Err = 1004 Then
  MsgBox "Array formula already occupies part of output range", , Iso
  o.Select
Else
  MsgBox "Error in median error calculation", , Iso
End If
End Sub

Function MedianConfLevel(ByVal N&) ' Confidence limit (%) of error on median
Attribute MedianConfLevel.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Conf!, A As Variant
If N > 25 Then
  Conf = 95
ElseIf N > 2 Then
' Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
  A = Array(75#, 87.8, 93.8, 96.9, 98.4, 93#, 96.1, 97.9, 93.5, 96.1, _
    97.8, 94.3, 96.5, 97.9, 95.1, 96.9, 93.6, 95.9, 97.3, 94.8, 96.5, 97.7, 95.7)
  Conf = A(N - 2)
End If
MedianConfLevel = Drnd(Conf, 5)
End Function

Function MedianUpperLim(v As Variant, Optional N) ' Upper error on median of V()
Attribute MedianUpperLim.VB_ProcData.VB_Invoke_Func = " \n14"
Dim u&, i&
Const uR = "11111222333444556667778"
If IM(N) Then
  If IsObject(v) Then N = v.Count Else N = UBound(v)
End If
If N > 25 Then
  u = 0.5 * (N + 1 - 1.96 * Sqr(N))
Else
' Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
  u = Val(Mid$(uR, (N - 3) + 1, 1))  ' High rank (U-th largest)
End If
MedianUpperLim = App.Large(v, u)
End Function

Function MedianLowerLim(v As Variant, Optional N) ' Lower error on median of V()
Attribute MedianLowerLim.VB_ProcData.VB_Invoke_Func = " \n14"
Dim L&, u&
Const Lr = "0304050607070809091011111213131414151616171818"
If IM(N) Then
  If IsObject(v) Then N = v.Count Else N = UBound(v)
End If
If N > 25 Then
  u = 0.5 * (N + 1 - 1.96 * Sqr(N))
  L = N + 1 - u
Else
' Table from Rock et al, based on Sign test & table of binomial probs for a ranked data-set.
  L = Val(Mid$(Lr, 2 * (N - 3) + 1, 2))  ' Low  rank (L-th largest)
End If
MedianLowerLim = App.Large(v, L)
End Function

Sub ConvertRange(Rin As Range, Rout() As Variant, NoneIn As Boolean)
Attribute ConvertRange.VB_ProcData.VB_Invoke_Func = " \n14"
Dim A, k&, r&, c%, q As Range
NoneIn = True
For Each A In Rin.Areas
  With A
    For r = 1 To .Rows.Count
      For c = 1 To .Columns.Count
        Set q = .Cells(r, c)
        If IsNumber(q) Then
          k = 1 + k
          ReDim Preserve Rout(k)
          Rout(k) = q.Value
          If k = 1 Then NoneIn = False
        End If
      Next c
    Next r
  End With
Next A
End Sub
