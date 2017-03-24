Attribute VB_Name = "Grf1"
' Isoplot Module Grf1
Option Explicit: Option Base 1: Option Private Module
Dim KeepFirstRange As Boolean

Sub ConstructPlot(ByVal N)
'xxx
Dim i%, j%, k%, v#, u#, CA As Object, PA As Object
Dim r%, T#, Emult#, Ptrn&, W#, s$, s1$
Dim rLine As Range, XYrange As Range, rRefChord As Range, tB As Boolean, H#
Dim Ly#, Xtik#, Ytik#, StartX#, EndX#
Dim StartY#, EndY#, Zs#, is2#, ss#, ii#, vv$, ee$
Dim Ellipses As Boolean, rBox As Range, vBox As Variant
Dim LineWithin As Boolean, NameFirstRange As Boolean
Dim DP() As DataPoints, xx#, yy#, temp#(), CurveColor&, tC&
Dim Cv As Curves, Slp#, Inr#, Cct%, CvWt%, CvSty%
Dim AgeEllipseLimits#(), LastXx#, LastYy#, MaxDecPts%

StatBar "creating plot framework"
SymbRow = Max(1, SymbRow)
If N > 0 Then ReDim DP(N)
CurveColor = IIf(ColorPlot, Opt.CurvClr, vbBlack)
If AddToPlot Then GoTo Scaled

If AutoScale Then
  MinX = 1E+37: MaxX = -1E+37: MinY = MinX:     MaxY = MaxX
  Emult = 2  ' Always autoscale to 2-sigma dispersion

    For i = 1 To N
      DP(i).X = InpDat(i, 1): DP(i).Xerr = InpDat(i, 2)
      DP(i).y = InpDat(i, 3): DP(i).Yerr = InpDat(i, 4)
      If i < N Or ConcAge Or Not ConcPlot Or Not Anchored Then
        u = DP(i).Xerr * Emult: v = DP(i).Yerr * Emult
        MinX = Min(MinX, DP(i).X - u): MaxX = Max(MaxX, DP(i).X + u)
        MinY = Min(MinY, DP(i).y - v): MaxY = Max(MaxY, DP(i).y + v)
      End If
    Next i

  If OtherXY Then
    Ly = -1E+37
  Else
    Ly = 0
    MinX = Max(0, MinX)
  End If

  MinY = Max(Ly, MinY)
  MaxX = IIf(MinX = MaxX, 1.1 * MinX, MaxX)
  MaxY = IIf(MinY = MaxY, 1.1 * MinY, MaxY)
  Xspred = MaxX - MinX: Yspred = MaxY - MinY

  If Not OtherXY Then
    MinX = Max(0, MinX - Xspred / 5):  MaxX = MaxX + Xspred / 5
    MinY = Max(Ly, MinY - Yspred / 5): MaxY = MaxY + Yspred / 5
  End If

  If ConcPlot And Inverse Then
    MinY = Max(0.04, MinY)

    If Yspred / MinY < 0.04 Then
      MinY = Max(0.04, MinY / 1.02)
      MaxY = 1.02 * MaxY
    End If

  End If

  Xspred = MaxX - MinX: Yspred = MaxY - MinY

Else

  If ConcPlot Then
    MinCurvAge = MinAge: MaxCurvAge = MaxAge

    If XYlim Then
      u = IIf(Normal, MinX, MaxX)
      u = ConcXage(u): v = ConcYage(MinY)
      MinAge = Max(0, Max(u, v))
      u = IIf(Normal, MaxX, MinX)
      v = ConcYage(MaxY)

      If u > 0 Then
         u = ConcXage(u)
         MaxAge = Min(6000, Min(u, v))
      Else
         MaxAge = v
      End If

      If Inverse Then MinAge = Max(MinAge, MaxAge / 20)

    Else

      If Normal Then
        MinX = ConcX(MinAge): MaxX = ConcX(MaxAge)
      Else
        MaxX = ConcX(MinAge): MinX = ConcX(MaxAge)
      End If

      MinY = ConcY(MinAge):   MaxY = ConcY(MaxAge)
    End If

    AgeSpred = MaxAge - MinAge
    CurvAgeSpred = AgeSpred

    If XYlim Then
      CurvAgeSpred = MaxCurvAge - MinCurvAge
    Else
      CurvAgeSpred = AgeSpred
      MinCurvAge = MinAge: MaxCurvAge = MaxAge
    End If

    If CurvTikInter > AgeSpred Then CurvTikInter = 0
  End If

  Xspred = MaxX - MinX: Yspred = MaxY - MinY
End If

If Not ProbPlot Then Tick MaxX - MinX, Xtik
Tick MaxY - MinY, Ytik
' Put the minx & miny values on a reasonable tick ("even" number)
MinX = Drnd(Int(Drnd(MinX / Xtik, 7)) * Xtik, 7)
MinY = Drnd(Int(Drnd(MinY / Ytik, 7)) * Ytik, 7)

If (ConcPlot And Not XYlim) Or (Not ConcPlot And AutoScale) Then
  ' Put the MaxX & MaxY values exactly on a tick
  xx = MinX: yy = MinY
  Do: xx = Drnd(xx + Xtik, 7): Loop Until xx >= MaxX Or xx = MinX: MaxX = xx
  Do: yy = Drnd(yy + Ytik, 7): Loop Until yy >= MaxY Or yy = MinY: MaxY = yy
End If

If MinX = MaxX Or MinY = MaxY Then
  MsgBox "Plot limits too close", , Iso
  StatBar
  ExitIsoplot
End If

Scaled: StatBar
Xspred = MaxX - MinX: Yspred = MaxY - MinY

If N > 0 Then
  tB = False

  For i = 1 To N
    If InBox(InpDat(i, 1), InpDat(i, 3)) Then tB = True: Exit For
  Next i

  If Not tB Then
    s$ = "None of the data points fall within the plot-box limits" _
      + vbLf & "-- Construct plot anyway?"
    If MsgBox(s$, vbYesNo, Iso) <> vbYes Then ExitIsoplot
  End If

  Dseries = 1
  ReDim Preserve DSeriesN(1)
  DSeriesN(1) = N
  ReDim temp(N, 2)

  For i = 1 To N
    temp(i, 1) = InpDat(i, 1)
    temp(i, 2) = InpDat(i, 3)
  Next i

  i = N

Else
  i = 1
End If

Set XYrange = sR(SymbRow, SymbCol, i + Anchored - 1 + SymbRow, 1 + SymbCol)
With XYrange   ' X-Y of input data
  .Name = "_gXY" & sn$(Dseries)

  If N > 0 Then
    .Value = temp: Erase temp
  Else
    .Value = 1
  End If

End With

AddSymbCol 2
LineInd XYrange
Cv.Nisocs = 0

If Not AddToPlot Or (PbGrowth And PbPlot) Then

  If ConcPlot Or PbGrowth Or uEvoCurve Then
    StatBar "creating curve data"
  End If

  If ConcPlot Then
    ConcordiaCurveData Cv
  ElseIf PbGrowth Then
    PbGrowthCurveData Cv
  ElseIf uEvoCurve Then
    StoreCurveData 1, Cv
  End If

End If

If RefChord Then
  ReferenceChord AnchorT1, AnchorT2, StartX, StartY, EndX, EndY
  i = (StartX >= MinX And StartX <= MaxX And StartY >= MinY And StartY <= MaxY)
  j = (EndX >= MinX And EndX <= MaxX And EndY >= MinY And EndY <= MaxY)

  If i Or j Then
    Cells(SymbRow, SymbCol).Value = StartX
    Cells(SymbRow, SymbCol + 1).Value = StartY
    Cells(1 + SymbRow, SymbCol).Value = EndX
    Cells(1 + SymbRow, SymbCol + 1).Value = EndY
    Set rRefChord = sR(SymbRow, SymbCol, 1 + SymbRow, SymbCol + 1)
    AddSymbCol 2
  Else
    RefChord = False
  End If

End If

' Construct the regression line (or its X-Y plane projection)
If Regress Then

  If Dim3 And (UseriesPlot Or (ConcPlot And Linear3D)) Then
    EndY = Min(MaxY, IntSl(1)) ' The Y-intercept at X=0

    If ConcPlot Then

      If ConcConstr Then       ' Start at the concordia-curve intercept
        StartX = ConcX(IntSl(3), True, True)
        StartY = ConcY(IntSl(3), True, True)
        is2 = ConcX(IntSl(3), 0, -1) / Uratio - IntSl(1) * ConcY(IntSl(3), 0, -1)
      Else                     ' Start at concordia-plane X-Y intercept
        StartX = -IntSl(3) / IntSl(4)
        StartY = IntSl(1) + IntSl(2) * StartX
        is2 = IntSl(2)         ' is2 is the X-Y slope
      End If

      EndX = (EndY - IntSl(1)) / is2
      LineWithin = InBox(StartX, StartY)

      If LineWithin Then
        PointsInBox StartX, StartY, EndX, EndY
      Else  ' That is, if can't start the line on the concordia curve
        ss = (EndY - StartY) / (EndX - StartX)
        ii = StartY - ss * StartX
        LineInBox ss, ii, StartX, StartY, EndX, EndY, LineWithin
      End If

    Else
      If UsType = 3 Then        ' 230/238 - 234/238 - 232/238
        StartX = -IntSl(3) / IntSl(4)
        StartY = IntSl(1) + IntSl(2) * StartX
        EndX = MaxX
        EndY = IntSl(1) + IntSl(2) * EndX
        v = EndY

        If EndY < MinY Then
          EndY = MinY
        ElseIf EndY > MaxY Then
          EndY = MaxY
        End If

        EndX = (EndY - IntSl(1)) / IntSl(2)

        If EndX < MinX Then
          EndX = MinX
        ElseIf EndX > MaxX Then
          EndX = MaxX
        End If

        EndY = IntSl(1) + IntSl(2) * EndX
        LineWithin = InBox(StartX, StartY)
        If Not LineWithin Then LineWithin = InBox(EndX, EndY)

      Else  ' UsType = 1 Or UsType = 2; X=232/238, Y=230/238 or 234/238
        Slp = IntSl(2): Inr = IntSl(1)
        LineInBox Slp, Inr, StartX, StartY, EndX, EndY, LineWithin
      End If

    End If

  Else

    If Crs(1) = 0 And Crs(3) = 0 Then
      LineWithin = False
    Else
      LineInBox Crs(1), Crs(3), StartX, StartY, EndX, EndY, LineWithin
    End If

  End If

  If LineWithin Then
    Cells(SymbRow, SymbCol).Value = StartX
    Cells(SymbRow, SymbCol + 1).Value = StartY
    Cells(1 + SymbRow, SymbCol).Value = EndX
    Cells(1 + SymbRow, SymbCol + 1).Value = EndY
    Set rLine = sR(SymbRow, SymbCol, 1 + SymbRow, SymbCol + 1)
    AddSymbCol 2
  End If

End If

If N > 0 Then
  Emult = 2 + (SigLev = 1 And Not Opt.AlwaysPlot2sigma)
  CreatePlotdataSource InpDat(), Emult
End If

If Ncurves > 0 And (Not AddToPlot Or PbGrowth) Then
  i = Ncurves: Ncurves = 0
  On Error Resume Next
  If Ncurves <> 1 Or Cv.NcurvEls(1) > 0 Then Ncurves = i
  On Error GoTo 0
End If

NameFirstRange = False: tB = False
On Error Resume Next
tB = (Cv.NcurvEls(1) > 0)
On Error GoTo 0

If tB Then
  tB = tB And (Not AddToPlot Or PbGrowth) And Ncurves > 0 And (ConcPlot Or PbGrowth Or uEvoCurve)
End If

KeepFirstRange = tB
Set rBox = Cells(1, 1)
On Error Resume Next

If tB Then
  Set rBox = CurvRange(1)
Else
  Set rBox = XYrange
  If Eellipse Or Ebox Then NameFirstRange = True
End If

On Error GoTo 0

If AddToPlot Then
  Sheets(PlotName$).Select
  Set sc = Sheets(PlotName$).SeriesCollection
  Nser = sc.Count
  Set IsoChrt = Ach
  GoTo DrawRefChord
Else
  StatBar "assembling plot elements"
  Charts.Add
  Landscape
  Set vBox = rBox
  Ach.ChartWizard vBox, xlXYScatter, 6, xlColumns, 1, 0, 2, , AxX$, AxY$
  Nser = 1
End If

Set IsoChrt = Ach: Set sc = IsoChrt.SeriesCollection
Set CA = IsoChrt.ChartArea: Set PA = IsoChrt.PlotArea
If NameFirstRange Then Call NameSeries
PlotName$ = IIf(ConcPlot, "Concordia", IIf(PbPlot, "PbPlot", "Isochron"))
MakeSheet PlotName$, IsoChrt

With IsoChrt
  .SizeWithWindow = False

  With PA
    .Height = Min(CA.Height - 60, 375)
    .Width = Min(CA.Width - 180, 460)
    .Top = 35: .Left = 110
     With .Border

      If FromSquid Or Opt.PlotboxBorder Then
        .LineStyle = xlContinuous
        .Weight = IIf(FromSquid, xlThick, AxisLthick)
        .Color = vbBlack
      Else
        .LineStyle = xlNone
      End If

    End With

    .Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.PlotboxClr, vbWhite))
  End With

End With

RemoveHdrFtr IsoChrt
CA.Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.SheetClr, vbWhite))
tB = False
On Error Resume Next

If Ncurves > 0 Then
  tB = (ConcPlot Or PbGrowth Or (uEvoCurve And Cv.NcurvEls(1) > 0))
End If

j = Opt.ConcLineThick
CvWt = xlThin: CvSty = xlContinuous 'In case corrupt MenutItems cell
If j = xlHairline Or j = xlMedium Or j = xlThick Then CvWt = j

If Cdecay And ConcPlot Then
  CvSty = xlNone
ElseIf UseriesPlot Then
  CvWt = xlMedium
ElseIf (ConcPlot Or PbPlot) And j = xlGray50 Then
  CvWt = xlThick: CvSty = j
End If

On Error GoTo 0
With IsoChrt.SeriesCollection(1)

  If tB Then
    StatBar "formatting curve"
    With .Border
      .LineStyle = CvSty

      If Not Cdecay Or Not ConcPlot Then
        .Weight = CvWt
        IsoChrt.SeriesCollection(1).Border.Color = CurveColor
      End If

    End With
    .MarkerStyle = xlNone: .Smooth = Not Cdecay
  Else
    .Border.LineStyle = xlNone: .MarkerStyle = xlNone
  End If

End With
StatBar "formatting Y-axis"
FormatAxes IsoChrt, True, MinY, MaxY, Ytik
StatBar "formatting X-axis"
FormatAxes IsoChrt, False, MinX, MaxX, Xtik

If FromSquid Then

  With IsoChrt
    PA.Height = CA.Height - 60
    PA.Width = 1.35 * PA.Height
    PA.Left = CA.Left + (CA.Width - PA.Width) / 2
    PA.Top = CA.Top + (CA.Height - PA.Height) / 2
  End With
Else
  CenterPlotArea
End If

If StackIso And InStr(AxY$, "/") > 0 Then PositionYaxisLabel "YaxisLabel"
If Cdecay Then CreateConcordiaBandShapes Cv, CurveColor, AgeEllipseLimits()
tB = False ' Add labels to curve age-ticks
On Error Resume Next

If Ncurves > 0 And CurvTikInter > 0 And Cv.Ncurvtiks > 0 Then

  If (ConcPlot Or PbTicks) And Cv.NcurvEls(1) > 0 Then
    tB = True
  ElseIf uEvoCurve Then
    tB = (uEvoCurve And Cv.Ncurvtiks > 0 And (uUseTiks Or uLabelTiks))
  End If

End If

On Error GoTo 0
If tB Then LabelCurveTicks Cv, AgeEllipseLimits()
If ConcPlot And Cdecay And Ncurves > 0 Then PlotConcordiaBand CurveColor
' Add additional U-series evolution curves &/or isochrons

If uEvoCurve Then _
  PlotUseriesEvolution Cv, IIf(ColorPlot, CurveColor, vbBlack), CvSty, CvWt, IsoChrt.SeriesCollection
StatBar
IsoChrt.Select

DrawRefChord:

If RefChord Then
  IsoChrt.SeriesCollection.Add rRefChord, xlColumns, False, True, False
  Nser = IsoChrt.SeriesCollection.Count
  With IsoChrt.SeriesCollection(Nser)
    With .Border
     .Color = vbBlack
     .Weight = xlThin: .LineStyle = xlContinuous
    End With
    .MarkerStyle = xlNone
  End With
End If

DrawYorkfitLine:

If Regress And LineWithin And Not ConcAge Then
  IsoChrt.SeriesCollection.Add Source:=rLine, Rowcol:=xlColumns, SeriesLabels:=False, _
    CategoryLabels:=True, Replace:=False
  Nser = IsoChrt.SeriesCollection.Count

  With IsoChrt.SeriesCollection(Nser)
    tC = IIf(ColorPlot, Opt.IsochClr, vbBlack)
    .Border.Color = tC

    With .Border
      .LineStyle = Opt.IsochStyle
      .Weight = Opt.IsochLineThick
    End With

    .MarkerStyle = xlNone
  End With

  If PlotErrEnv And Not Dim3 Then
    With yf
      PlotErrorEnvelope .Slope, .SlopeError, .Intercept, .InterError, .Xbar
    End With
  End If

  IsoChrt.Deselect
  Sheets(Ash.Name).Visible = True
End If

If DoPlot And N > 0 Then

  If ColorPlot Then
    tC = excClr
  ElseIf DoShape And excSymb = 0 And Not StraightLine And Not SplineLine Then
    tC = Menus("cGray75")
  Else
    tC = vbBlack
  End If

  PlotDataPoints xlThin, xlContinuous, tC, True
End If

DataPlotted:
AddCopyButton
If (ConcAge And Len(AgeRes$) = 0) Or FromSquid Then Exit Sub
tB = (Regress And Not (ConcPlot And Dim3 And Linear3D And Not ConcConstr))
tB = tB And (Not PbPlot Or PbType = 1) And Not (OtherXY And iLambda(OtherIndx) = 0)
If tB Then AddResultsBox
End Sub

Sub CurvPos(ByVal T#, X#, y#, Optional CurveNum = 1, Optional Cerr = 0)
Dim tmp1#, tmp2#, ConcXerr#, ConcYerr#, e5t#, e8t#
ViM CurveNum, 1
ViM Cerr, 0
If ConcPlot Then
  X = ConcX(T):  y = ConcY(T)
  If Cdecay And Cerr <> 0 Then
    e5t = Exp(Lambda235 * T):   e8t = Exp(Lambda238 * T)
    tmp1 = SQ(Lambda235err) + SQ(Lambda235 / Lambda238 * Lambda238err)
    If Normal Then
      tmp2 = SQ(Lambda238err) + SQ(Lambda238 / Lambda235 * Lambda235err)
      ConcYerr = 2 * T * e8t * Sqr(tmp2)               ' At 2-Sigma!
    Else
      ConcYerr = 2 * X * T * e5t / Uratio * Sqr(tmp1)  ' At 2-Sigma!
    End If
    y = y + Cerr * ConcYerr
  End If
ElseIf PbGrowth Then
  X = PbR(T, 0): y = PbR(T, PbType)
ElseIf uEvoCurve Then
  y = GammaU(T * Thou, Ugamma0(CurveNum))
  Th230_U238ar T * Thou, y, X
End If
End Sub

Sub CurveInBox(CurvPts#(), ByVal AgeMin#, ByVal AgeMax#, ByVal AgeStep#, _
  Nsegs%, Optional CurveNum = 1, Optional Cerr = 0)
' For U-Pb concordia curves, Pb-growth curves, or U-series evolution curves, return in CurvPts()
' the array of X-Y pts defining the segment of the curve (beween AgeMin & AgeMax) that lies
'  within the plotbox.
Dim Within As Boolean, Changed As Boolean, Nrml As Boolean
Dim T#, x1#, y1#, x2#, y2#, Ttol#
Dim Direction%, Tstep#, Bt#, Dt#, In1#, In2#
Dim ce As Boolean, bDelt#
ViM CurveNum, 1
ViM Cerr, 0
Nrml = (Not (ConcPlot And Inverse) And Not PbGrowth)
ce = (ConcPlot And Cerr <> 0)
If Nrml Then
  Direction = 1
  If ConcPlot Then
    T = Max(AgeMin - 24 * AgeStep, 0) ' 8 * AgeStep, 0)
    If Cerr <> 0 Then T = Max(0, T - 24 * AgeStep) ' 8 * AgeStep)
    T = Max(AgeMin - 8 * AgeStep, 0) ' 2 * AgeStep, 0)
  End If
Else
  Direction = -1
  If ConcPlot Then
    T = Min(6000, AgeMax + 24 * AgeStep) '8 * AgeStep)
    If Cerr <> 0 Then T = Min(6000, T + 24 * AgeStep) '8 * AgeStep)
  Else
    T = IIf(PbGrowth, AgeMax, AgeMax + 2 * AgeStep)
  End If
End If
CurvPos T, x1, y1, CurveNum, Cerr
In1 = InBox(x1, y1)
Tstep = Direction * AgeStep
Nsegs = 0
Do
  T = T + Tstep
If Nrml And T > AgeMax Then
  Exit Do
ElseIf Not Nrml And T < AgeMin Then
  Exit Do
End If
  CurvPos T, x2, y2, CurveNum, Cerr
  PointsInBox x1, y1, x2, y2, Within, Changed
  'If x2 > MaxX Then Exit Do
  If Not Within Then x1 = x2: y1 = y2
Loop Until Within Or x2 > MaxX
If Not Within Then Exit Sub
Nsegs = 2
ReDim CurvPts(2, 2)
CurvPts(1, 1) = x1:  CurvPts(2, 1) = y1
CurvPts(1, 2) = x2:  CurvPts(2, 2) = y2
x1 = x2:          y1 = y2
If ConcPlot Then
  Bt = IIf(Inverse, 0.1, 5000)
Else
  Dt = AgeStep / 10
  If Nrml Then
    Bt = MaxAge + Dt
  Else
    Bt = Max(MinAge - Dt, 0)
  End If
End If
Ttol = (MaxAge - MinAge) / 10000
If uEvoCurve Then Changed = False
Do
  T = T + Tstep
  If Abs(T - Bt) < Ttol Then T = Bt
If (Nrml And T > Bt) Or (Not Nrml And T < Bt) Then Exit Do
  If Abs(T - Bt) < Ttol Then T = Bt
  CurvPos T, x2, y2, CurveNum, Cerr
  If Not uEvoCurve Then ' because U-series evolution curve can curve back into plotbox
    PointsInBox x1, y1, x2, y2, Within, Changed
  End If
  Nsegs = 1 + Nsegs
  ReDim Preserve CurvPts(2, Nsegs)
  CurvPts(1, Nsegs) = x2: CurvPts(2, Nsegs) = y2
  x1 = x2: y1 = y2
Loop Until Changed Or T = Bt
End Sub

Sub StoreCurveData(ByVal CurveNum%, Cv As Curves)
' Calculate the pts that define the Concordia, Pb-growth, or U-series evolution curve(s)
'  & store them in the PlotDat sheet.
Dim i%, j%, k%, r%, MaxCurvEls%, NcurvTiks0%, uIsoPts#(2, 2)
Dim X#, y#, x_1#, Slope#, Inter#, Nlo%, Nhi%, Rr%, DidClip As Boolean
Dim IsoInBox As Boolean, xStart#, xEnd#, yStart#, yEnd#
Dim IsoLoc#(2, 2), T#, AgeStep#, CurveLoc#()
Dim AgeMax#, AgeMin#, Cp#(), ctAgeMin#, ctAgeMax#, Cerr%, Ccols%
Dim CurveLo#(), CurveHi#(), xXc, yYc, CdTik#(), RhoTik#, dx#, dY#, CovXY#, Ttol#
Dim tmp1#, tmp2#, eT#(1, 5), Nsegs%, ee#(), eet#(), MaxTix%, e5t#, e8t#, s$
SymbRow = Max(1, SymbRow)
If uEvoCurve Then
  ' Start curve at min 230/238
   MinAge = (Prnd(Th230age(MinX, Ugamma0(CurveNum)), 0)) ' / Thou)
ElseIf PbGrowth Then
  MaxAge = pbStartAge
End If
If MaxAge <= MinAge Then
  i = 0: On Error Resume Next: i = Cv.NcurvEls(CurveNum): On Error GoTo 0
  If i = 0 Then Exit Sub
  Cv.NcurvEls(CurveNum) = 0: Exit Sub
End If
AgeSpred = MaxAge - MinAge
MaxCurvEls = 20 * (3 / 2) ^ (Opt.CurveRes - 1)
AgeStep = AgeSpred / (MaxCurvEls - 1)
'If Cdecay And ConcPlot Then AgeStep = AgeStep * 4
If CurveNum = 1 Then ReDim Cv.NcurvEls(Ncurves)
If uEvoCurve Then
  If AgeSpred > 300 Then AgeStep = AgeStep / 2
  If AgeSpred > 600 Then AgeStep = AgeStep / 2
End If
AgeMin = MinAge: AgeMax = MaxAge
If FromSquid Then
  MaxTix = 10
Else
  MaxTix = IIf(Cdecay, IIf(AutoScale, 10, 16), IIf(UseriesPlot, 127, 30))
End If
ReDim Cp#(2, 1)
If Cdecay And ConcPlot Then ReDim CurvRange(3)
For Cerr = Cdecay To -Cdecay
  CurveInBox Cp(), AgeMin, AgeMax, AgeStep, Nsegs, CurveNum, Cerr
  If Cerr = 0 And Nsegs = 0 Then Exit Sub
  If Nsegs Then
    If Cerr < 0 Then
      ReDim CurveLo(Nsegs, 2)
      For i = 1 To Nsegs: For j = 1 To 2: CurveLo(i, j) = Cp(j, i): Next j, i
    ElseIf Cerr > 0 Then
      ReDim CurveHi(Nsegs, 2)
      For i = 1 To Nsegs: For j = 1 To 2: CurveHi(i, j) = Cp(j, i): Next j, i
    Else
      Cv.NcurvEls(CurveNum) = Nsegs
      ReDim CurveLoc(Nsegs, 2)
      For i = 1 To Nsegs: For j = 1 To 2: CurveLoc(i, j) = Cp(j, i): Next j, i
    End If
  End If
Next Cerr
If CurvTikInter > 0 Then
  Do
    If uEvoCurve Then FirstCurvTik = 0
    If ConcPlot And Inverse And XYlim Then
      ctAgeMin = MinCurvAge: ctAgeMax = MaxCurvAge
      Do
        CurvPos ctAgeMin + CurvTikInter, X, y, CurveNum
        If InBox(X, y) And ctAgeMin >= MinCurvAge Then Exit Do
        If ctAgeMin >= MaxCurvAge Then
          i = 0: On Error Resume Next: i = Cv.NcurvEls(CurveNum): On Error GoTo 0
          If i = 0 Then Exit Sub
          Cv.NcurvEls(CurveNum) = 0: Exit Sub
        End If
        ctAgeMin = ctAgeMin + CurvTikInter
      Loop
      Do
        CurvPos ctAgeMax - CurvTikInter, X, y, CurveNum
      If InBox(X, y) And ctAgeMax <= MaxCurvAge Then Exit Do
        ctAgeMax = ctAgeMax - CurvTikInter
      Loop
    Else
      ctAgeMin = AgeMin: ctAgeMax = AgeMax
      MaxCurvAge = MaxAge
    End If
    With Cv
      .Ncurvtiks = Max(0, Int((MaxCurvAge - FirstCurvTik) / CurvTikInter + 1))
      If .Ncurvtiks = 0 Then Exit Sub
      NcurvTiks0 = .Ncurvtiks
      ReDim .CurvTik(.Ncurvtiks, 0 To 2), .CurvTikPresent(.Ncurvtiks)
      i = 0: j = 0: T = FirstCurvTik
      Do
        i = i + 1
        If CurveNum = 1 Then
          CurvPos T, X, y, CurveNum
          If True Then 'InBox(X, Y)  ' OK to leave in, because (if concordia curve), will
            j = 1 + j                '  assign MarkerStyle of all ticks outside plotbox
            .CurvTik(j, 0) = T       '  as xlNone.
            .CurvTik(j, 1) = X: .CurvTik(j, 2) = y
          ElseIf .Ncurvtiks > 0 Then
            .Ncurvtiks = .Ncurvtiks - 1
          End If
        End If
        T = Drnd(T + CurvTikInter, 5)
      Loop Until i >= NcurvTiks0
    End With
  If j < MaxTix Or CurveNum > 1 Then Exit Do
    CurvTikInter = 2 * CurvTikInter
  Loop
  If Cdecay And ConcPlot Then
    i = 0: j = 0: T = FirstCurvTik
    Do
      i = i + 1
      CurvPos T, X, y, CurveNum
      If InBox(X, y) Then
        j = 1 + j
        e5t = Exp(Lambda235 * T): e8t = Exp(Lambda238 * T)
        If Normal Then
          dx = T * e5t * Lambda235err
          dY = T * e8t * Lambda238err
          RhoTik = 0
        Else
          dx = X * X * T * e8t * Lambda238err ' Lambda23xerr is 1-sigma!
          tmp1 = e5t * Lambda235err / Uratio
          tmp2 = y * e8t * Lambda238err
          dY = X * T * Sqr(SQ(tmp1) + SQ(tmp2))
          CovXY = X ^ 3 * y * SQ(T * e8t * Lambda238err)
          RhoTik = CovXY / dx / dY
        End If
        eT(1, 1) = X: eT(1, 2) = dx
        eT(1, 3) = y: eT(1, 4) = dY
        eT(1, 5) = RhoTik
        StatBar "creating " & Str(T) & " Ma age-ellipse"
        Ellipse 1, eT(), eet(), Nsegs, 2, DidClip
        If Nsegs Then
          If DoShape And Cdecay Then
            Rr = 1 + SymbRow
            With Cells(SymbRow, SymbCol)
              .Value = T:
              .Name = "ConcAgeTikAge" & sn$(j)
            End With
          Else
            Rr = SymbRow
          End If
          MTrans eet(), ee()
          Cv.NageElls = j: Cv.CurvTikPresent(j) = True
          s$ = "ConcAgeTik" & sn$(j)
          With sR(Rr, SymbCol, Nsegs + Rr - 1, SymbCol + 1)
            .Name = s$: .Value = ee
          End With
          If Opt.ClipEllipse And DidClip Then
            EllCorner Range(s$)
            LineInd Range(s$), "ClippedEllipse", 2
          End If
          AddSymbCol 2
        End If
      End If
      T = T + CurvTikInter
    Loop Until j >= Cv.Ncurvtiks Or T > MaxCurvAge
    StatBar
  End If
  If uEvoCurve And uPlotIsochrons And Cv.Nisocs = 0 Then
    T = FirstCurvTik
    Do
      T = T + CurvTikInter
      If T > 0 Then
        Th230_U238ar T * Thou, MinY, x_1
        ' Slope of isochron at age t years
        Slope = 1 / (LambdaK * (1 - Exp(-LambdaDiff * T * Thou)))
        Inter = MinY - Slope * x_1
        LineInBox Slope, Inter, xStart, yStart, xEnd, yEnd, IsoInBox
        If IsoInBox Then
          uIsoPts(1, 1) = xStart: uIsoPts(1, 2) = yStart
          uIsoPts(2, 1) = xEnd:   uIsoPts(2, 2) = yEnd
          Cv.Nisocs = 1 + Cv.Nisocs
          ReDim Preserve UisoRange(Cv.Nisocs)
          Set UisoRange(Cv.Nisocs) = sR(SymbRow, SymbCol, 1 + SymbRow, 1 + SymbCol)
          UisoRange(Cv.Nisocs).Value = uIsoPts()
          AddSymbCol 2
        End If
      End If
    Loop Until T >= MaxCurvAge
  End If
End If
i = 0: On Error Resume Next: i = Cv.NcurvEls(CurveNum): On Error GoTo 0
If i > 0 Then
  Set CurvRange(CurveNum) = sR(SymbRow, SymbCol, Cv.NcurvEls(CurveNum) - 1 + SymbRow, 1 + SymbCol, ChrtDat)
  CurvRange(CurveNum).Value = CurveLoc()
  AddSymbCol 2
  LineInd CurvRange(CurveNum)
ElseIf DoShape Then
  DoShape = False
  Exit Sub
End If
If Cdecay And ConcPlot Then
  Nlo = 0: Nhi = 0
  On Error Resume Next
  Nlo = UBound(CurveLo, 1): Nhi = UBound(CurveHi, 1)
  On Error GoTo 0
  If DoShape And Nlo > 0 And Nhi > 0 Then
    ConcBandShapeRange CurvRange(2), Nlo, Nhi, CurveLo(), CurveHi()
    AddSymbCol 2
  Else
    If Nlo > 0 And Nhi > 0 Then
      Set CurvRange(2) = sR(SymbRow, SymbCol, Nlo - 1 + SymbRow, 1 + SymbCol)
      Set CurvRange(3) = sR(SymbRow, 2 + SymbCol, Nhi - 1 + SymbRow, 3 + SymbCol)
      CurvRange(2).Value = CurveLo()
      CurvRange(3).Value = CurveHi()
      For i = 2 To 3: LineInd CurvRange(i), "ConcBand": Next i
      AddSymbCol 4
    End If
  End If
End If
If CurvTikInter > 0 And Not (UseriesPlot And uPlotIsochrons And CurveNum > 1) Then
  ' pts for concordia, U-series evolution, or Pb-growth curve ticks.
  ReDim CurveLoc(Cv.Ncurvtiks, 2)
  For i = 1 To Cv.Ncurvtiks
    CurveLoc(i, 1) = Cv.CurvTik(i, 1): CurveLoc(i, 2) = Cv.CurvTik(i, 2)
  Next i
  Set TikRange(CurveNum) = sR(SymbRow, SymbCol, Cv.Ncurvtiks - 1 + SymbRow, SymbCol + 1)
  TikRange(CurveNum).Value = CurveLoc
  LineInd TikRange(CurveNum)
  AddSymbCol 2
End If
End Sub

Sub PlotDataPoints(ByVal sWeight%, ByVal sStyle%, sColor&, _
    Dpoints As Boolean, Optional StatBarPhrase, Optional SortType)
' Plot the data pts in the PlotBox
Dim i%, rEllRange As Range, Clr&(), j%, PdN%, SB$
Dim c&, k%, tmp$, lWeight%(), FC&, bC&
Dim sType%, Ptrn&, Colr&, tr As Range, Tprop%
Dim Bt#, bh#, ShpType$, Curved As Boolean, eWt#

If NIM(StatBarPhrase) Then SB$ = StatBarPhrase
sType = IIf(NIM(StatBarPhrase), SortType, 1)
PdN = DSeriesN(Dseries) + Anchored
IsoChrt.Select: IsoChrt.Visible = True
ReDim Clr(PdN, 2), lWeight(PdN)
If Not Eellipse Then SB$ = "plotting the data points"

For i = 1 To PdN          '  Set the line-clr of the plot-symbs &, if the first cell of

  If ArChron Then

    If i < ArChronSteps(1) Or i > ArChronSteps(2) Then

      If DoShape Then ' Cyan fill (clr), white fill (b&w)
        sColor = IIf(ColorPlot, vbCyan, vbWhite)
      Else            ' Cyan outline (clr), gray outline (b&w)
        sColor = IIf(ColorPlot, vbCyan, Menus("cGray89"))
      End If

    Else

      If DoShape Then ' Red fill (clr), light gray fill (b&w)
        sColor = IIf(ColorPlot, vbRed, Menus("cGray75"))
      Else            ' Red outline (clr), black outline (b&w)
        sColor = IIf(ColorPlot, vbRed, vbBlack)
      End If

    End If

  End If

  If FromSquid Then
    c = vbCyan: Clr(i, 1) = vbCyan: Clr(i, 2) = vbCyan
  Else
    Clr(i, 1) = xlAutomatic '   the data-row is bolded & plot-symb is an xcel symb
    c = DatClr(i)           '   (as opposed to an error bar or ellipse),
    If c = 0 Then c = 1     '   the symbol-fill color (default is white).

    If sColor >= 0 Then
      Clr(i, 1) = IIf(DoShape And Not eCross And _
        Not StraightLine And Not SplineLine And _
        excSymb <> xlX And excSymb <> 9, vbBlack, sColor)
    ElseIf StraightLine Or SplineLine Then
      Clr(i, 1) = Black
    ElseIf c <> xlNone And c <> xlAutomatic Then
      Clr(i, 1) = Abs(c)
    End If

  End If

  lWeight(i) = sWeight

  If c < 0 And c <> xlNone And c <> xlAutomatic _
    And excSymb <> xlX And excSymb <> xlPlus Then

    If DoShape Then
      Clr(i, 2) = IIf(sColor >= 0, sColor, Abs(c))
    ElseIf excSymb = xlCircle Or excSymb = xlSquare Or _
      excSymb = xlTriangle Or excSymb = xlDiamond Then
      Clr(i, 2) = White
    Else
      Clr(i, 2) = Clr(i, 1)
    End If

    lWeight(i) = Min(xlThick, 1 + sWeight)

  ElseIf excSymb = xlX Or excSymb = xlPlus Then
    Clr(i, 2) = xlNone

  Else

    If Not FromSquid Then

      If DoShape Then
        Clr(i, 2) = IIf(sColor >= 0, Abs(sColor), Clr(i, 1))
      Else
        Clr(i, 2) = White
      End If

    End If

  End If

Next i

If Eellipse Or Ebox Then

  If AddToPlot Or Dpoints Then
    Set tr = ChrtDat.Range("_gXY" & sn$(Dseries))
    Ach.SeriesCollection.Add Source:=tr, Rowcol:=xlColumns, _
      SeriesLabels:=False, CategoryLabels:=True, Replace:=False
    Nser = Ach.SeriesCollection.Count

    With Ach.SeriesCollection(Nser)
      If .MarkerStyle <> xlNone Then .MarkerStyle = xlNone
     .Border.LineStyle = xlNone
    End With

    NameSeries
  End If

  If Eellipse Then
    If Len(SB$) = 0 Then SB$ = "plotting err-ellipse for point"
  End If

  If DoShape Then GetScale

  For i = 1 To PdN
    On Error GoTo nextSer

    If Eellipse Then
      StatBar SB$ & Str(i)
      tmp$ = "Ellipse": ShpType$ = "ErrEll"
    Else
      tmp$ = "_Box": ShpType$ = "ErrBox"
      StatBar SB$ & Str(i)
    End If

    Set rEllRange = ChrtDat.Range(tmp$ & sn$(Dseries) & Und & sn$(i))
    On Error GoTo 0
    If Eellipse And rEllRange.Count <= 2 Then GoTo nextSer
    IsoChrt.Select

    If DoShape Then
      Curved = False

      If Eellipse Then

        If Opt.ClipEllipse And _
         rEllRange(2 + rEllRange.Rows.Count, 1) = "ClippedEllipse" Then
          EllCorner rEllRange
        Else
          Curved = True
        End If

      End If

      eWt = IIf(ArChron, 0, 0.75)
      AddShape ShpType$, rEllRange, Clr(i, 2), _
        IIf(ForceFill, excClr, Black), Curved, sType, , , _
        IIf(ForceFill, 1, -0.5 * FromSquid), eWt

    Else
      Ach.SeriesCollection.Add Source:=rEllRange, Rowcol:=xlColumns, _
      SeriesLabels:=False, CategoryLabels:=True, Replace:=False
      Nser = Ach.SeriesCollection.Count

      With Ach.SeriesCollection(Nser)
        If .MarkerStyle <> xlNone Then .MarkerStyle = xlNone
        If .Smooth <> Eellipse Then .Smooth = Eellipse

        With .Border

          If ArChron And Not DoShape And Not ColorPlot Then
            .Weight = IIf(i < ArChronSteps(1) Or i > ArChronSteps(2), xlHairline, xlThin)
            Ach.SeriesCollection(Nser).Border.Color = Clr(i, 1)
          Else
            If .Weight <> lWeight(i) Then .Weight = lWeight(i)
            Ach.SeriesCollection(Nser).Border.Color = Clr(i, 1)
          End If

          If .LineStyle <> sStyle Then .LineStyle = sStyle
          If Eellipse And Opt.ClipEllipse Then _
            Call EllipseClip(Ach.SeriesCollection(Nser), rEllRange)
        End With

      End With

      LabelSeriesColor rEllRange, Clr(i, 1)
    End If

nextSer: On Error GoTo 0
  Next i

ElseIf excSymb <> 0 Or eCross Or StraightLine Or SplineLine Then
  StatBar SB$
  Set tr = ChrtDat.Range("_gXY" & sn$(Dseries))
  Ach.SeriesCollection.Add Source:=tr, Rowcol:=xlColumns, SeriesLabels:=False, _
    CategoryLabels:=True, Replace:=False
  Nser = Ach.SeriesCollection.Count
  NameSeries
  LineInd tr

  If SplineLine Then
    MakeSpline tr, Clr()
  Else

    With Ach.SeriesCollection(Nser)

      If StraightLine Then

        With .Border
          .LineStyle = xlContinuous: .Color = Clr(1, 1)
        End With

        '.Smooth = False
      Else
        .Border.LineStyle = xlNone
      End If

      If excSymb = xlX Or excSymb = xlPlus Then
        .MarkerBackgroundColorIndex = xlNone
      ElseIf .MarkerBackgroundColor <> vbWhite Then
        .MarkerBackgroundColor = vbWhite
      End If

      If Not eCross Then
        .MarkerStyle = excSymb: .MarkerSize = Opt.AgeTikSymbSize
      Else
        .MarkerStyle = xlNone
        .ErrorBar Direction:=xlX, include:=xlBoth, Type:=xlCustom, _
           Amount:=xErRange, MinusValues:=xErRange
        .ErrorBar Direction:=xlY, include:=xlBoth, Type:=xlCustom, _
           Amount:=yErRange, MinusValues:=yErRange

        With .ErrorBars
          .Border.Color = Clr(1, 1) ' Can't specify error-bar color for each point

          With .Border
            If .LineStyle <> sStyle Then .LineStyle = sStyle
            If .Weight <> sWeight Then .Weight = sWeight
          End With

          Tprop = IIf(Opt.EndCaps, xlCap, xlNoCap)
          If .EndStyle <> Tprop Then .EndStyle = Tprop
        End With

      End If

      If excSymb <> 0 Then

        For i = 1 To PdN
          StatBar SB$ & Str(i)

          If DoShape And excSymb <> xlX And excSymb <> xlCross And excSymb <> xlPlus Then
            bC = Clr(i, 2): FC = bC
          Else
            FC = Clr(i, 1): bC = Clr(i, 2)
          End If

          With .Points(i)
            .MarkerForegroundColor = FC

            If bC = xlNone Or bC = xlAutomatic Then
              .MarkerBackgroundColorIndex = bC
            Else
              .MarkerBackgroundColor = bC
            End If

          End With

        Next i

      End If

    End With

  End If

  If Dseries = 1 And Not AddToPlot And Not KeepFirstRange Then

    With IsoChrt ' Delete the dummy data-set used for initial Charts.ADD
      Bt = .PlotArea.Top: bh = .PlotArea.Height
      Ach.SeriesCollection(1).Delete     ' Causes chart-title of "IsoDat1" to be added, so
      Nser = Nser - 1  '  must delete, then restore plotbox dimensions.

      If .HasTitle Then
        .ChartTitle.Delete
        .PlotArea.Top = Bt: .PlotArea.Height = bh
      End If

    End With

  End If

End If

StatBar
End Sub

Sub CreatePlotdataSource(XYdat#(), ByVal Emult#, Optional StatusBarPhrase)
Dim i%, j%, k%, MaxR%, ecR%
Dim c%, r%, xCarr#(), yCarr#(), Ear#()
Dim tmp$, ns%, xs#, YS#, re As Range, eX#, eY#
Dim xx#, yy#, Boxx(0 To 4, 2)
Dim PdN%, EarT#(), SB$, gXY As Range, DidClip As Boolean
ReDim xCarr(DSeriesN(Dseries), 1), yCarr(DSeriesN(Dseries), 1)
PdN = DSeriesN(Dseries) + Anchored
SymbRow = Max(1, SymbRow)
ChrtDat.Visible = True: ChrtDat.Select
If NIM(StatusBarPhrase) Then SB$ = StatusBarPhrase

If Eellipse Then
  If Len(SB$) = 0 Then SB$ = "creating ellipse data for point"

  For i = 1 To PdN
    xs = Emult * XYdat(i, 2): YS = Emult * XYdat(i, 4)
    ns = 0
    StatBar SB$ & Str(i)

    If xs < Xspred Or YS < Yspred Then               ' If one of the dimensions of the
      Ellipse i, XYdat(), EarT(), ns, Emult, DidClip '  ellipse is larger than the plotbox,
      MTrans EarT(), Ear()                           '  ignore it by creating a dummy range.
    End If

    tmp$ = "Ellipse" & sn$(Dseries) & Und & sn$(i)

    If ns Then
      sR(SymbRow, SymbCol, ns - 1 + SymbRow, 1 + SymbCol).Name = tmp$
      Set re = Range(tmp$)
      re.Value = Ear
      If Opt.ClipEllipse And DidClip Then LineInd re, "ClippedEllipse", 2
    Else
      sR(SymbRow, SymbCol, SymbRow, 1 + SymbCol).Name = tmp$
      Range(tmp$).Formula = "=0"
    End If

    AddSymbCol 2
  Next i

  StatBar

ElseIf Ebox Then

  For i = 1 To PdN
    xx = XYdat(i, 1): yy = XYdat(i, 3)
    eX = Emult * XYdat(i, 2): eY = Emult * XYdat(i, 4)
    Boxx(0, 1) = xx - eX:    Boxx(0, 2) = yy - eY
    Boxx(1, 1) = Boxx(0, 1): Boxx(1, 2) = yy + eY
    Boxx(2, 1) = xx + eX:    Boxx(2, 2) = Boxx(1, 2)
    Boxx(3, 1) = Boxx(2, 1): Boxx(3, 2) = Boxx(0, 2)
    Boxx(4, 1) = Boxx(0, 1): Boxx(4, 2) = Boxx(0, 2)
    tmp$ = "_Box" & sn$(Dseries) & Und & sn$(i)
    r = SymbRow + 7 * (i - 1)
    sR(r, SymbCol, r + 4, 1 + SymbCol).Name = tmp$
    Range(tmp$).Value = Boxx
    LineInd Range(tmp$), "ErrBox"
  Next i

  AddSymbCol 2

ElseIf eCross Then

  For i = 1 To PdN
    xCarr(i, 1) = XYdat(i, 2) * Emult
    yCarr(i, 1) = XYdat(i, 4) * Emult
  Next i

  ecR = PdN - 1 + SymbRow
  Set xErRange = sR(SymbRow, SymbCol, ecR)
  Set yErRange = sR(SymbRow, SymbCol + 1, ecR, SymbCol + 1)
  xErRange.Value = xCarr: yErRange.Value = yCarr
  LineInd xErRange
  LineInd yErRange
  AddSymbCol 2
End If

If (Dseries > 1 And Not Eellipse) Or AddToPlot Then
  ' Create XY data series if the first dataset,
  Set gXY = sR(SymbRow, SymbCol, DSeriesN(Dseries) - 1 + SymbRow, 1 + SymbCol)

  For i = 1 To PdN
    gXY(i, 1) = XYdat(i, 1): gXY(i, 2) = XYdat(i, 3)
  Next i

  gXY.Name = "_gXY" & sn$(Dseries)
  LineInd gXY
  AddSymbCol 2
End If

End Sub

Private Function InBox(ByVal X#, ByVal y#, Optional MaxxedOK = True) As Boolean
' Is an x-y point within the plot-box?
Dim Lx#, Ux#, Ly#, Uy#, Rx#, Ry#
ViM MaxxedOK, True
InBox = False
Rx = Drnd(X, 7)
If MaxxedOK Then
  If Rx < Drnd(MinX, 7) Then Exit Function
  If Rx > Drnd(MaxX, 7) Then Exit Function
  Ry = Drnd(y, 7)
  If Ry < Drnd(MinY, 7) Then Exit Function
  If Ry > Drnd(MaxY, 7) Then Exit Function
Else
  If Rx <= Drnd(MinX, 7) Then Exit Function
  If Rx >= Drnd(MaxX, 7) Then Exit Function
  Ry = Drnd(y, 7)
  If Ry <= Drnd(MinY, 7) Then Exit Function
  If Ry >= Drnd(MaxY, 7) Then Exit Function
End If
InBox = True
End Function

Sub StatBar(Optional s)  ' Set the text in the Excel status-bar
Dim ss As String
If IM(s) Then s = "" Else s = Space(20) & s & "..."
App.StatusBar = s
End Sub

Private Sub NameSeries()
Dim i%, NdSer%
NdSer = 0
For i = 1 To sc.Count
  If Left$(sc(i).Name, 6) = "IsoDat" Then NdSer = 1 + NdSer
Next i
NdSer = 1 + NdSer
sc(Nser).Name = "IsoDat" & tSt(NdSer)
End Sub

Sub EllipseClip(Eseries As Object, eRange As Range)
Dim sme As Boolean, i%, j%, LastInB As Boolean
Dim EllPts As Object, InB As Boolean
LastInB = False: sme = True
Set EllPts = Eseries.Points
For j = 1 To EllPts.Count
  InB = InBox(eRange(j, 1), eRange(j, 2), False)
  If Not InB Then       ' Get smoothing artefacts at clipping
    If Not LastInB Then '  boundary for clipped ellipses.
      If sme Then sc(Nser).Smooth = False: sme = False
      EllPts(j).Border.LineStyle = xlNone
    End If
  End If
  LastInB = InB
Next j
End Sub

Sub FormatAxes(Cht As Object, ByVal Yaxis As Boolean, ByVal AxMin, ByVal AxMax, ByVal AxTik)
Dim aX As Object, k%, b As Boolean  ' Format X or Y axis
Dim L!, T!
Set aX = Axxis(1 - Yaxis, Cht)
With aX
  If AxMin = 0 And AxMax = 0 Then
    .MinimumScale = xlAutomatic: .MaximumScale = xlAutomatic
  Else
    .MinimumScale = Drnd(AxMin, 7): .MaximumScale = Drnd(AxMax, 7)
  End If
  If AxTik > 0 And Not FromSquid Then
    .MinorUnit = Drnd(AxTik, 5): .MajorUnit = Drnd(AxTik * 2, 5)
    .MajorTickMark = IIf(Opt.AxisTickCross, xlCross, xlInside)
    .MinorTickMark = xlInside
    If .TickLabelPosition <> xlNextToAxis Then .TickLabelPosition = xlNextToAxis
    With .TickLabels
      With .Font
        .Name = Opt.AxisTikLabelFont: .Background = xlTransparent
        If (AgeExtract Or YoungestDetrital) And Yaxis Then
          .Size = 22
        Else
          .Size = Opt.AxisTikLabelFontSize - (ProbPlot And Yaxis) * 2
        End If
      End With
      .NumberFormat = TickFor(AxMin, AxMax, AxTik)
    End With
  Else
    .MajorTickMark = xlNone: .MinorTickMark = xlNone
    .TickLabelPosition = xlNone
  End If
  .Border.Weight = AxisLthick
  .ScaleType = False
  If .Crosses <> xlAutomatic Then .Crosses = xlAutomatic
  If .ReversePlotOrder <> False Then .ReversePlotOrder = False
  If .MinimumScale < 0 Then .CrossesAt = .MinimumScale
  .MinorUnitIsAuto = Opt.AxisAutoTikSpace: .MajorUnitIsAuto = Opt.AxisAutoTikSpace
  If .HasTitle Then
    With .AxisTitle
      With .Characters.Font
        .Name = Opt.AxisNameFont: .Background = xlTransparent
        If (AgeExtract Or YoungestDetrital) And Yaxis Then
          .Size = 28
        Else
          .Size = Opt.AxisNameFontSize
        End If
      End With
    End With
    If Yaxis Then
      k = xlUpward ' Bug in Mac Excel98 - can't display non-horizontal superscripts
      If StackIso And InStr(AxY$, "/") > 0 Then
        k = xlHorizontal  'Mac And Not NoSuper And Int(ExcelVersion) = 8 And b
        GetScale , True
        .HasTitle = False
        StackYaxis AxY$, (Opt.AxisNameFontSize), True
      ElseIf Not Mac Then
        Superscript Phrase:=.AxisTitle, DidSuper:=b, CanStack:=True
        .AxisTitle.Orientation = k
      End If
    Else
      Superscript .AxisTitle
    End If
  End If
End With
End Sub

Sub StackYaxis(ByVal Na$, ByVal Fsize!, ByVal Fbold As Boolean)
' Take "ABC/DEF" and show as ABC stacked over DEF with a line as divider
Dim P%, Upper$, Lower$, c As Chart, L!, T!, u As Object
Dim q As Object, HL As Object, YL As Object, DS As Boolean
Set c = Ach: Set q = c.PlotArea
P = InStr(Na$, "/"): Upper$ = Left(Na$, P - 1): Lower$ = Mid(Na$, P + 1)
With c.TextBoxes.Add(0, 0, 0, 0)
  .AutoSize = True
  .Text = Upper$:  .Name = "YaxUpper"
  .Font.Bold = Fbold: .Font.Size = Fsize
End With
Set u = c.TextBoxes("YaxUpper")
Superscript u
T = T + u.Height
With c.TextBoxes.Add(L, T, 0, 0)
  .AutoSize = True
  .Text = Lower$:  .Name = "YaxLower"
  .Font.Bold = Fbold: .Font.Size = Fsize
End With
Superscript c.TextBoxes("YaxLower"), DS
Set HL = c.Shapes.AddLine(L, T - 2 - Not DS, L + u.Width, T - 2 - Not DS)
HL.Line.Weight = 1 - Fbold / 2: HL.Name = "Yline"
Set YL = c.Shapes.Range(Array("YaxLower", "YaxUpper", "Yline"))
YL.Align msoAlignCenters, False
YL.Group.Name = "YaxisLabel"
End Sub

Sub PositionYaxisLabel(ByVal Na$)  ' Position a stacked Y-axis label
Dim P As Object, c As Chart
Set c = Ach: Set P = c.PlotArea
With c.Shapes(Na$)
  .Top = P.Top + P.Height / 2 - .Height / 2
  .Left = P.Left - .Width * 1.2
End With
End Sub

Sub AddCopyButton(Optional Button, Optional Macro, Optional OnWksht As Boolean = False)
Dim AC As Object, ACA As Object, s As Shapes, o As Shape, ns%
Dim nss%, f As Object, Shps As Object
ViM OnWksht, False
If OnWksht Then
  Set AC = Ash: Set ACA = AC: Set s = AC.Shapes
Else
  If Ash.Type <> -4169 Or Not Menus("ShowMoveChart") Then Exit Sub
  Set AC = Ach: Set ACA = AC.ChartArea: Set s = AC.Shapes
End If
Macro = IIf(NIM(Macro), Macro, IIf(Mac, "MoveChart", "CopyPicture"))
Button = IIf(NIM(Button), Button, IIf(Mac, "ChartToData2", "ChartToData"))
For Each o In s
  If o.Name = Button Or Right(o.OnAction, Len(Macro)) = Macro Then Exit Sub
Next o
MenuSht.Shapes(Button).Copy
AC.Paste
If OnWksht Then
  Set Shps = Ash.Shapes: ns = Shps.Count
  Set f = IIf(DoPlot Or ns > 1, Shps(ns - 1), RangeIn(1)(1, 1))
  With Shps(ns)
    If DoPlot Then
      .Left = f.Left + f.Width - .Width
      .Top = f.Top + f.Height - .Height
    ElseIf ns = 1 Then
      .Left = f.Left + 30: .Top = f.Top + 5
    Else
      .Left = f.Left + 5: .Top = f.Top + f.Height + 3
    End If
  End With
  RangeIn(1)(1, 1).Select
Else
  With Ach.Shapes(Button)
    .Left = ACA.Width - .Width: .Top = ACA.Height - .Height
  End With
  AC.Select: AC.Deselect
End If
End Sub

Sub CurvAgeTikNdeci(Ndec%, Cv As Curves)
Dim i%, j%, tmp$ ' Determine max# decimal pts for age-tix
Ndec = 0
For i = 1 To UBound(Cv.CurvTik, 1) Step 2
  tmp$ = sn$(Cv.CurvTik(i, 0)): j = DecPos(tmp$)
  If j > 0 Then Ndec = Max(Ndec, Len(Mid$(tmp$, 1 + j)))
Next i
End Sub

Sub ObliqueTick(X!, y!, Xlength#, Angle#, x1!, y1!, x2!, y2!)
' Put a tick-line at a specified location, length (in logical x-units), and angle (degrees)
' All input variables are in user (logical) units.  x1,y1-x2,y2 are the logical
'  coordinates of the rotated line
Dim Radians, PlotX1!, PlotX2!, PlotY1!, PlotY2!
Dim Width!, Left1!, Left2!, Top1!, Top2!, Xinc#, Yinc#, CenterLeft!, CenterTop!
Dim PlotLength#, rotX1!, rotX2!, rotY1!, rotY2!
If PlotBoxHeight = 0 Then GetScale
Width = Xlength / Xspred * (PlotBoxRight - PlotBoxLeft) ' in points
LeftTop_XY_Convert CenterLeft, CenterTop, X, y, True
LeftTop_XY_Convert PlotX1, 0, X - Xlength / 2, 0, True
LeftTop_XY_Convert PlotX2, 0, X + Xlength / 2, 0, True
Radians = pi / 180 * Angle
PlotLength = PlotX2 - PlotX1 ' Line width in points
Xinc = PlotLength / 2 * Cos(Radians): Yinc = PlotLength / 2 * Sin(Radians) ' points
rotX1 = CenterLeft - Xinc: rotY1 = CenterTop - Yinc ' points
rotX2 = CenterLeft + Xinc: rotY2 = CenterTop + Yinc '   "
LeftTop_XY_Convert rotX1, rotY1, x1, y1, False ' convert to logical
LeftTop_XY_Convert rotX2, rotY2, x2, y2, False '   "
End Sub

Sub PutObliqueTicks(Cv As Curves, SerC As Object, ByRef Nser%)
' Put line-ticks normal to concordia curve and label them
Dim i%, TikRow%, X!, y!, T#, Slope#
Dim Ndec%, Angle#, rRange As Range, s$, TikXlen#, x1!, y1!
Dim x2!, y2!, FontSize!, j%, aa!, Ad!, l1!, l2!, Lwidth!, Loff!, Cwidth!
GetScale
SymbRow = Max(1, SymbRow): TikRow = SymbRow
FontSize = Val(Menus("CurvTikFontSize")) '* 12 / 18
CurvAgeTikNdeci Ndec, Cv
For i = 1 To UBound(Cv.CurvTik, 1) Step 2
  With Cv: X = .CurvTik(i, 1): y = .CurvTik(i, 2): T = .CurvTik(i, 0): End With
  If T > 0 And InBox(X, y) Then
    Slope = ConcSlope(T)
    LogicalSlopeToPhysicalAngle Slope, Angle
    If Inverse Then Angle = Angle + 180 ' of tick
    aa = Angle - 90: Ad = pi / 180# * aa
    TikXlen = (Xspred / 50) * Opt.AgeTikSymbSize / 6 ' X-length of tick
    ObliqueTick X, y, TikXlen, -aa, x1, y1, x2, y2
    sR(TikRow, SymbCol, , , ChrtDat) = x1
    sR(TikRow, 1 + SymbCol, , , ChrtDat) = y1
    sR(1 + TikRow, SymbCol, , , ChrtDat) = x2
    sR(1 + TikRow, 1 + SymbCol, , , ChrtDat) = y2
    Set rRange = sR(TikRow, SymbCol, TikRow + 1, SymbCol + 1, ChrtDat)
    IsoChrt.SeriesCollection.Add rRange, xlColumns, False, True, False
    Nser = SerC.Count
    With SerC(Nser)
      .MarkerStyle = xlNone
      With .Border
        .Color = Opt.CurvClr
        If Opt.ConcLineThick = xlGray50 Then
          .LineStyle = xlGray50: .Weight = xlThick
        Else
          .LineStyle = xlContinuous: .Weight = Opt.ConcLineThick
        End If
      End With
      With .Points(1)
        s$ = IIf(Ndec > 0, App.Fixed(T, Ndec), sn$(T)) & "  "
        .ApplyDataLabels Type:=xlShowLabel, LegendKey:=False
         With .DataLabel
            .Text = s$: .Font.Size = FontSize
            .Font.Name = Opt.AgeTikFont
            .Position = xlLabelPositionRight: l2 = .Left
            .Position = xlLabelPositionLeft: l1 = .Left
            .Position = xlLabelPositionCenter
            Lwidth = l2 - l1: Cwidth = Lwidth / Len(s$)
            Loff = (Lwidth - Cwidth) / 2
            .Orientation = aa
            .Left = .Left - Loff * Cos(Ad)
            .Top = .Top + Loff * Sin(Ad)
         End With
      End With
    End With
    TikRow = 2 + TikRow
  End If
Next i
'ChrtDat.Select
LineInd sR(TikRow - 1, SymbCol, , , ChrtDat)
'IsoChrt.Select
AddSymbCol 2
End Sub

Sub cIsochron_Click()
cAgeSpectrum_click
End Sub

Sub cAgeSpectrum_click()
With DlgSht("ArStepAge").CheckBoxes
  .Item("cInset").Enabled = DoPlot And ((IsOn(.Item("cIsochron")) And IsOn(.Item("cAgeSpectrum"))))
End With
End Sub

Sub LabelCurveTicks(Cv As Curves, AgeEllipseLimits#())
Dim FirstEll%, Ndec%, NcTix%, i%, j%, k%, Nser%
Dim s$, s1$, s2$, One$, Two$, l1!, l2!, fs!
Dim eR As Range, ell As Range, AtikLeft!, AtikTop!, xx!, yy!, SlopeFact#
Dim T#, TT#, TrueSlope#, DeltY#, DeltX#, Xdelt#, v1#, v2#
Dim Xoffs#, Yoffs#, SeLeft!, SeTop!, Cwidth!, CaRt!, CaBot!
StatBar "labelling curve-ticks"
fs = IIf(FromSquid, 24, Opt.AgeTikFontSize)
SymbRow = Max(1, SymbRow)
If DoShape And ConcPlot And Cdecay Then
  FirstEll = 1
  If Cv.NageElls > 6 Then ' Put best-appearing tick-label first
    With ChrtDat
      One$ = .Range("ConcAgeTikAge" & sn$(1)).Text
      Two$ = .Range("ConcAgeTikAge" & sn$(2)).Text
    End With
    If Len(One$) > Len(Two$) Then
      FirstEll = 2
    Else
      One$ = Right$(One$, 1): Two$ = Right$(Two$, 1)
      If Two$ = "0" Then
        If One$ <> "0" Then FirstEll = 2
      ElseIf Val(One$) Mod 2 > 0 Then
        If Val(Two$) Mod 2 = 0 Then FirstEll = 2
      End If
    End If
  End If
  Ndec = 0
  For i = FirstEll To Cv.NageElls Step 2    ' Determine max# decimal pts for age-tix
    s$ = ChrtDat.Range("ConcAgeTikAge" & sn$(i)).Text
    j = DecPos(s$)
    If j > 0 Then Ndec = Max(Ndec, Len(Mid$(s$, 1 + j)))
  Next i
  j = SymbRow - 1
  GetScale
  With Ach
    CaRt = Right_(.ChartArea): CaBot = Bottom(.ChartArea)
  End With
  For i = FirstEll To Cv.NageElls Step 2
    Set ell = ChrtDat.Range("ConcAgeTik" & sn$(i))
    v1 = ell(i, 1).Value: v2 = ell(i, 2).Value
    If v1 >= MinX And v1 <= MaxX And v2 >= MinY And v2 <= MaxY Then
      Set eR = ChrtDat.Range("ConcAgeTikAge" & sn$(i))
      With Ach.TextBoxes.Add(0, 0, 1, 1)
        .AutoSize = True: .Border.LineStyle = xlNone
        If Ndec > 0 Then s$ = App.Fixed(eR.Value, Ndec) Else s$ = eR.Text
        .Text = s$
        .Font.Name = Opt.AgeTikFont: .Font.Size = fs
        .HorizontalAlignment = xlRight
        AtikLeft = AgeEllipseLimits(i, 1) - .Width
        If Normal Then
          .VerticalAlignment = xlBottom
          AtikTop = AgeEllipseLimits(i, 2) - .Height
        Else
          .VerticalAlignment = xlTop
          AtikTop = AgeEllipseLimits(i, 3)
        End If
        If AtikLeft > 0 And AtikTop > 0 And AtikLeft < CaRt And AtikTop < CaBot Then
          .Left = AtikLeft: .Top = AtikTop
          LeftTop_XY_Convert AtikLeft, AtikTop, xx, yy, False
          j = 1 + j
          ChrtDat.Cells(j, SymbCol) = xx: ChrtDat.Cells(j, 1 + SymbCol) = yy
          ' Enable the line below to to have the labels rescale when the chart is rescaled
          .Name = ChrtDat.Name & Und & sn$(SymbCol) & "|" & sn$(j) & "~" & sn$(j) & "_2T"
        Else
          .Delete
        End If
      End With
    End If
  Next i
  If j > SymbRow Then AddSymbCol 2
Else
  Nser = Ach.SeriesCollection.Count
  If LineAgeTik And ConcPlot And Not Cdecay Then
    If CurvTikInter > 0 Then PutObliqueTicks Cv, Ach.SeriesCollection, Nser
  Else
    Ach.SeriesCollection.Add TikRange(1), xlColumns, False, True, False
    Nser = Ach.SeriesCollection.Count
    With Ach.SeriesCollection(Nser)
      .Border.LineStyle = xlNone
      If Cdecay Or (uEvoCurve And uPlotIsochrons) Then
        If .MarkerStyle <> xlNone Then .MarkerStyle = xlNone
      Else
        .MarkerStyle = Opt.AgeTikSymbol
        .MarkerForegroundColor = IIf(ColorPlot, Opt.AgeTikSymbClr, vbBlack)
        .MarkerBackgroundColor = IIf(ColorPlot, Opt.AgeTikSymbFillClr, vbWhite)
        If FromSquid Then
          .MarkerSize = 8
        ElseIf .MarkerSize <> Opt.AgeTikSymbSize Then
          .MarkerSize = Opt.AgeTikSymbSize
        End If
        For i = 1 To TikRange(1).Rows.Count
          With TikRange(1)
            xx = .Cells(i, 1): yy = .Cells(i, 2)
          End With
          ' Don't show ticks if outside plotbox
          If Not InBox(xx, yy) Then .Points(i).MarkerStyle = xlNone
        Next i
      End If
    End With
  End If
  If Not ConcPlot Or Cdecay Or Not LineAgeTik Then
    If CurvTikInter > 0 And Not (PbGrowth And Not PbTickLabels) _
      And Not (UseriesPlot And uPlotIsochrons) Then
      SlopeFact = Xspred / Yspred * IsoChrt.PlotArea.Height / IsoChrt.PlotArea.Width
      NcTix = 0
      CurvAgeTikNdeci Ndec, Cv
      For i = 1 To Ach.SeriesCollection(Nser).Points.Count Step 2
        xx = Cv.CurvTik(i, 1): yy = Cv.CurvTik(i, 2)
        If InBox(xx, yy) Then
          T = Cv.CurvTik(i, 0)
          NcTix = 1 + NcTix
          If UseriesPlot And T = 0 Then GoTo NextTikI ' Don't label zero-age isochron
          If i = 1 Then
            TT = T + CurvTikInter
            If Not ConcPlot Then
              j = NumChars(T): k = NumChars(TT)
              s1$ = Right$(sn$(T), 1): s2$ = Right$(sn$(TT), 1)
              If j > k Or (s1$ <> "0" And s2$ = "0") Then
                T = TT: i = i + 1
              End If
            End If
          End If
          s$ = IIf(Ndec > 0, App.Fixed(T, Ndec), sn$(T))
          j = Len(s$)
          If uEvoCurve Then
            UconcSlope T * Thou, Ugamma0(1), TrueSlope
            TrueSlope = TrueSlope * SlopeFact
          ElseIf ConcPlot Then
            TrueSlope = ConcSlope(T) * SlopeFact
          ElseIf PbTickLabels Then
            TT = IIf(T <= 0, 0.1, T)
            DeltY = (PbR(TT + 1, PbType) - PbR(TT, PbType))
            DeltX = (PbR(TT + 1, 0) - PbR(TT, 0))
            TrueSlope = DeltY / DeltX * SlopeFact
          End If
          Xdelt = (xx - MinX) / (MaxX - MinX)
          If Abs(TrueSlope) < 2 Or Xdelt > 0.1 Or UseriesPlot Then
            With Ach.SeriesCollection(Nser).Points(i)
              .ApplyDataLabels Type:=xlShowLabel, LegendKey:=False
              .DataLabel.Select
              With Selection
                .Text = s$
                .Font.Size = fs
                .Font.Name = Opt.AgeTikFont
                .Position = xlLabelPositionRight: l2 = .Left
                .Position = xlLabelPositionLeft:  l1 = .Left
                Cwidth = (l2 - l1) / Len(s$)
                Xoffs = 0: Yoffs = 0
                If UseriesPlot Then
                  Xoffs = 0 'fs * 0.25 * (TrueSlope > -0.5)
                  Yoffs = fs * 0.9 * ((TrueSlope > -1.3) + (TrueSlope > 1))  '0.25 *
                ElseIf Not Cdecay Or Normal Then
                  If Normal Then
                    If TrueSlope < 0.1 Then
                      .Position = xlLabelPositionAbove: Yoffs = fs / 2
                    ElseIf TrueSlope < 0.2 Then
                      Yoffs = -fs * 0.65: Xoffs = Cwidth / 2
                    ElseIf TrueSlope > 4 Then
                      Xoffs = Cwidth / 6
                    ElseIf FromSquid Then
                      Yoffs = -0.8 * fs: Xoffs = -Cwidth / 4
                    Else
                      Yoffs = -fs / 3: Xoffs = Cwidth / 4
                    End If
                  Else
                    If TrueSlope > -0.1 Then
                      .Position = xlLabelPositionBelow: Yoffs = -fs / 5
                    ElseIf TrueSlope < -6 Then
                      Yoffs = fs / 8: Xoffs = Cwidth / 8
                    Else
                      Yoffs = fs * 0.8: Xoffs = Cwidth / 3
                    End If
                  End If
                  'k = (TrueSlope < 0 And TrueSlope > -1)
                  'Xoffs = fs * (0.728 * (j + k / 2) + 0.5 + 0.3 * k - 0.5 * Cdecay)
                  'Select Case TrueSlope
                  '  Case Is > 1:    Yoffs = 0
                  '  Case Is > 0:    Yoffs = -0.625
                  '  Case Is > -1.5: Yoffs = 0.84
                  '  Case Else:      Yoffs = 0
                  'End Select
                  'Yoffs = Yoffs * fs
                End If
                SeLeft = .Left: SeTop = .Top
                .Left = SeLeft + Xoffs: .Top = SeTop + Yoffs
              End With
            End With
          End If
        End If
NextTikI:
      Next i
    End If
    If NcTix > 0 Then ' Otherwise error
      On Error Resume Next
      With Ach.SeriesCollection(Nser).DataLabels
        .Font.Name = Opt.AgeTikFont: .Font.Size = fs
        If UseriesPlot Then
          .Interior.ColorIndex = ClrIndx(IIf(ColorPlot, Opt.PlotboxClr, vbWhite))
        Else
          If .Interior.ColorIndex <> xlNone Then .Interior.ColorIndex = xlNone
        End If
       .Font.Background = xlTransparent
      End With
      On Error GoTo 0
    End If ' (NxTix>0)
  End If
End If  ' Shape age-ticks?
End Sub

Sub AddResultsBox()
Dim i%, s$, s1$, H!, W!, xx!, yy!, Mem&
StatBar "adding results-box to chart"
If Not ConcAge Then
  If ConcPlot Then
    AgeRes$ = "Intercept"
    If Not PbAnchor And Not (Dim3 And Linear3D) Then AgeRes$ = AgeRes$ & "s"
    AgeRes$ = AgeRes$ & " at " & vbLf & Lir$
    If Not PbAnchor And Not (Dim3 And Linear3D) Then
      If Lir$ <> "" Then AgeRes$ = AgeRes$ & " & "
      AgeRes$ = AgeRes$ & Uir$
    Else
      AgeRes$ = AgeRes$ & " Ma"
    End If
  ElseIf Not ConcAge And Not UseriesPlot Then
    AgeRes$ = "Age = " & Uir$
    If Not PbPlot And Not UseriesPlot And (ArgonPlot Or Not Inverse) Then
      ' Add initial-ratios for inverse isochrons at some point!
      AgeRes$ = AgeRes$ & vbLf & "Initial "
      If ArgonPlot Then
        If Normal Then
          s$ = VandE(Crs(8), Crs(10), 2)
        Else
          s$ = VandE(1 / Crs(3), Crs(4) / SQ(Crs(3)), 2)
        End If
        s1$ = Menus("AxYn").Cells(Isotype)
      Else
        s$ = VandE(Crs(3), Crs(4), 2): s1$ = AxY$
      End If
      AgeRes$ = AgeRes$ & s1$ & " =" & s$
    End If
  End If
  If Not UseriesPlot And Not Robust Then AgeRes$ = AgeRes$ & vbLf & "MSWD = " & Msw$
  If ArChron And ArIso Then
    With IsoChrt
      H = .PlotArea.Height: W = .ChartArea.Width
      AddArRejSymbNote 0
    End With
    If ArSpect And ArInset Then
      Mem = App.MemoryFree
      On Error GoTo CantPaste
      ArChrt.Shapes("ArAge").Cut: IsoChrt.Paste
      On Error GoTo 0
    Else
      AddArTextbox PlCap$, 10
      Last(IsoChrt.TextBoxes).Select
    End If
    With Selection.ShapeRange
      .Left = IIf(Normal, W - .Width - 10, .Left = 10)
      .Top = H - .Height / 2
    End With
    If ArSpect And ArInset Then
      With ArChrt
        .Select
        With .ChartArea
          With .Interior
            .ColorIndex = ClrIndx(IIf(ColorPlot, Menus("cLtGreen"), Menus("cGray75")))
          End With
         .Border.Weight = xlHairline
        End With
        If Not ColorPlot Then .PlotArea.Interior.ColorIndex = ClrIndx(vbWhite)
        .CopyPicture Appearance:=xlPrinter, Format:=xlPicture
      End With
CantPaste: On Error GoTo NoCanPaste
      IsoChrt.Select: IsoChrt.Paste
      On Error GoTo 0
      With Selection.ShapeRange
        .ScaleWidth 0.45, False: .ScaleHeight 0.45, False
        .Left = IIf(Normal, 3, W - .Width - 3)
        .Top = 3: .Name = "ArStep"
        With .Fill: .ForeColor.SchemeColor = 8: .Visible = msoTrue: .Solid: End With
        With .Shadow
          .Type = msoShadow6
          If Not ColorPlot Then .ForeColor.SchemeColor = 63
        End With
        With .PictureFormat
          .CropTop = 10: .CropLeft = 10: .CropRight = 30: .CropBottom = 4
        End With
      End With
    End If
    Exit Sub
  End If
End If
With IsoChrt
  xx = 0 ' Stagger repeated age-result boxes
  For i = 1 To .TextBoxes.Count  ' "For Each" not reliable here for some reason
    s$ = .TextBoxes(i).Text
    If InStr(s$, "Intercepts at") Or InStr(s$, "Age =") Then xx = 10 + xx
  Next i
End With
GetScale
s = "ChartResBox"
Do
  s1 = ""
  For i = 1 To Ach.Shapes.Count
    If Ach.Shapes(i).Name = s Then
      s = s & "@"
      s1 = "x"
      Exit For
    End If
  Next i
Loop Until s1 = ""
IsoChrt.TextBoxes.Add(0, 0, 0, 0).Select
With Selection
  .Name = s
  .Characters.Text = AgeRes$
  .Font.Name = Opt.IsochResFont: .Font.Size = Opt.IsochResFontSize
  .Interior.ColorIndex = xlNone: .HorizontalAlignment = xlCenter
  If .VerticalAlignment <> xlTop Then .VerticalAlignment = xlTop
  If .Orientation <> xlHorizontal Then .Orientation = xlHorizontal
  With .Border
    .LineStyle = xlContinuous: .Color = vbBlack
    .Weight = xlThin
  End With
  .Interior.Color = vbWhite: .Placement = xlMoveAndSize
  If Not .PrintObject Then .PrintObject = True
  .AutoSize = True
  i = IIf(ConcPlot, Opt.ConcResboxRnd, Opt.IsochResboxRnd)
  If .RoundedCorners <> i Then .RoundedCorners = i
  IncreaseLineSpace Selection, 1.2  ' Changes font size of main characters!
  With .ShapeRange.Shadow
    If Opt.IsochResboxShadw Then
      .Type = msoShadow6
      .ForeColor.RGB = RGB(120, 120, 120)
    Else
      .Visible = False
    End If
  End With
End With
ConvertSymbols Selection
Superscript Phrase:=Selection
With Selection
  .Left = PlotBoxRight - .Width - 15 - xx
  .Top = IIf(ConcPlot And Inverse, PlotBoxTop + 15 + xx, PlotBoxBottom - .Height - 15 - xx)
End With
StatBar
Exit Sub
NoCanPaste: On Error GoTo 0
MsgBox "Sorry -- Can't complete the cut/paste operation" & vbLf & "(maybe out of memory)"
End Sub
