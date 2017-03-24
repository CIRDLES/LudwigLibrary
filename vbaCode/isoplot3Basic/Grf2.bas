Attribute VB_Name = "Grf2"
' Isoplot Module Grf2
Option Explicit: Option Base 1: Option Private Module

Sub PbGrowthCurveData(Cv As Curves)
Dim i%, j%, k%
Dim AlphaMin#, c#, XageMin#, YageMin#, XageMax#, YageMax#
Dim temp#, r1#, r2#, CurveXmax#, f#, Tstart#
Dim X#, y#, T#, LastT#, Match As Boolean, ff#
If pbStartAge <= 0 Then PbGrowth = False: Ncurves = 0: Exit Sub
AlphaMin = IIf(pbAlpha0 < MinX, MinX, pbAlpha0)
MaxAge = PbT(Max(AlphaMin, PbR0(0)), 0)
MinAge = 0
If MaxAge = 0 Then Exit Sub
XageMin = 0: YageMin = 0: XageMax = 0: YageMax = 0
temp = PbExp(0) - (MaxX - PbR0(0)) / MuIsh(0)
If temp > MINLOG And temp < MAXLOG Then
  XageMin = Log(temp) / PbLambda(0)
End If
temp = PbExp(0) - (MinX - PbR0(0)) / MuIsh(0)
If temp > MINLOG And temp < MAXLOG Then
  XageMax = Log(temp) / PbLambda(0)
End If
temp = PbExp(PbType) - (MaxY - PbR0(PbType)) / MuIsh(PbType)
If temp > MINLOG And temp < MAXLOG Then
  YageMin = Log(temp) / PbLambda(PbType)
End If
temp = PbExp(PbType) - (MinY - PbR0(PbType)) / MuIsh(PbType)
If temp > MINLOG And temp < MAXLOG Then
  YageMax = Log(temp) / PbLambda(PbType)
End If
MinAge = Max(XageMin, YageMin)
If MinAge >= XageMax Or MinAge >= YageMax Then GoTo BadLims
MaxAge = Min(XageMax, YageMax)
If MaxAge <= XageMin Or MaxAge <= YageMin Then GoTo BadLims
Tick MaxAge - MinAge, CurvTikInter
If CurvTikInter = 0 Then GoTo BadLims
f = (MaxAge - MinAge) / (XageMax - XageMin)
ff = (1 - (f < 0.5) - 2 * (f < 0.25) - 4 * (f < 0.125))
If PbType = 2 And ff > 1 Then ff = ff / 2
CurvTikInter = CurvTikInter * ff
temp = MinAge: MinAge = -CurvTikInter
Do
  MinAge = MinAge + CurvTikInter
Loop Until CSng(MinAge) >= CSng(temp)
temp = MaxAge: MaxAge = MinAge
Do
  MaxAge = MaxAge + CurvTikInter
Loop Until CSng(MaxAge) >= CSng(temp)
Tstart = Min(MaxAge + 2 * CurvTikInter, pbStartAge)
FirstCurvTik = MinAge
MaxAge = Min(PbT(MinX, 0), PbT(MinY, PbType))
MinAge = Max(PbT(MaxX, 0), PbT(MaxY, PbType))
StoreCurveData 1, Cv
Exit Sub

BadLims:
MinAge = 0: MaxAge = 0
End Sub

Sub LineInBox(ByVal Slope, ByVal Intercept, xStart, yStart, xEnd, yEnd, LineWithinBox As Boolean)
' Specify starting & ending pts of a line within plot-box boundaries.
Dim y1#, y2#
LineWithinBox = False
y1 = Slope * MinX + Intercept
If y1 > MaxY Then
  If Slope >= 0 Then Exit Sub
  xStart = (MaxY - Intercept) / Slope
  yStart = MaxY
ElseIf y1 < MinY Then
  If Slope <= 0 Then Exit Sub
  xStart = (MinY - Intercept) / Slope
  yStart = MinY
Else
  xStart = MinX
  yStart = Slope * MinX + Intercept
End If
y2 = Slope * MaxX + Intercept
If y2 > MaxY Then
  If Slope <= 0 Then Exit Sub
  xEnd = (MaxY - Intercept) / Slope
  yEnd = MaxY
ElseIf y2 < MinY Then
  If Slope >= 0 Then Exit Sub
  xEnd = (MinY - Intercept) / Slope
  yEnd = MinY
Else
  xEnd = MaxX
  yEnd = Slope * MaxX + Intercept
End If
LineWithinBox = True
End Sub

Sub PointsInBox(x1, y1, x2, y2, Optional WithinBox, Optional Changed)
' Return segment of line (x1,y1)-(x2,y2) that lies within the plotbox
'  boundaries. Ignores the case when both pts are outside, but their
'  connecting-line intersects the plotbox.
Dim Slope#, Inter#, LeftY#, RightY#
Dim TopX#, BottomX#, InX#, OutX#, OutY#
Dim FirstIn As Boolean, SecondIn As Boolean, DeltaX#, DeltaY#
FirstIn = InBox2(x1, y1):   SecondIn = InBox2(x2, y2)
If NIM(WithinBox) Then WithinBox = True
If FirstIn And SecondIn Then  ' Both pts in plotbox
  If NIM(Changed) Then Changed = False
  Exit Sub
ElseIf Not FirstIn And Not SecondIn Then
  If NIM(WithinBox) Then WithinBox = False
  If NIM(Changed) Then Changed = False
  Exit Sub
End If
' One of the pts in plotbox, one out
If NIM(Changed) Then Changed = True
If FirstIn Then
  InX = x1: OutX = x2: OutY = y2
Else
  InX = x2: OutX = x1: OutY = y1
End If
DeltaX = x2 - x1: DeltaY = y2 - y1
If CSng(DeltaX) = 0 Then
  OutX = InX
  If OutY >= MaxY Then
    OutY = MaxY
  ElseIf OutY <= MinY Then
    OutY = MinY
  End If
ElseIf CSng(DeltaY) = 0 Then
  OutY = y1
  If OutX <= MinX Then
    OutX = MinX
  ElseIf OutX >= MaxX Then
    OutX = MaxX
  End If
Else
  Slope = DeltaY / DeltaX
  Inter = y2 - Slope * x2
  LeftY = Slope * MinX + Inter                   ' Y at Min-X
  RightY = Slope * MaxX + Inter                  ' Y at Max-X
  BottomX = (MinY - Inter) / Slope               ' X at Min-Y
  TopX = (MaxY - Inter) / Slope                  ' X at Max-y
  If OutX > InX Then                             ' Line-segment intersects plotbox at:
    If RightY > MaxY Then
      OutX = (MaxY - Inter) / Slope: OutY = MaxY ' top
    ElseIf RightY < MinY Then
      OutX = (MinY - Inter) / Slope: OutY = MinY ' bottom
    Else
      OutX = MaxX: OutY = Slope * MaxX + Inter   ' right
    End If
  ElseIf LeftY > MaxY Then
    OutX = (MaxY - Inter) / Slope: OutY = MaxY   ' top
  ElseIf LeftY < MinY Then
    OutX = (MinY - Inter) / Slope: OutY = MinY   ' bottom
  Else
    OutX = MinX: OutY = Slope * MinX + Inter     ' left
  End If
End If
If FirstIn Then
  x2 = OutX: y2 = OutY
Else
  x1 = OutX: y1 = OutY
End If
End Sub

Sub Superscript(Phrase As Object, Optional DidSuper = False, Optional CanStack = False, Optional AllNukes = False)
' Superscript 2 numeric packets in a phrase, stack ratios if specified.
Dim s$, i%, j%, StartNum%, nCt%
Dim P$, u%, nU$(), Nukes$

ViM DidSuper, False
ViM CanStack, False
ViM AllNukes, False
If NoSuper Then Exit Sub

' Initialize the q s-string if necessary
On Error Resume Next
u = IIf(AllNukes, UBound(AllNuke$, 1), UBound(RadNuke$, 1))
On Error GoTo 0
If u < 2 Then GetNuclides u, AllNukes
s$ = Phrase.Text

If InStr(s$, "/") > 0 And CanStack And StackIso Then
  i = 0

  Do
    j = InStr(Mid$(s$, 1 + i), "/")
  If j = 0 Then Exit Do
    s$ = Left$(s$, i + j) & vbLf & Mid$(s$, i + j + 1)
    i = 1 + i + j
  Loop

  Phrase.Text = s$
End If

On Error GoTo NoPhrase
With Phrase

  For i = 1 To u
    If AllNukes Then P = AllNuke$(i) Else P = RadNuke$(i)
    StartNum = InStr(s$, P$)

    If InStr(P$, "159") Then
      i = i
    End If

    If StartNum > 0 Then

      For nCt = 1 To Len(P$) - 1
        If Not IsNum(P$, 1 + nCt) Then Exit For
      Next nCt

      .Characters(StartNum, nCt).Font.Superscript = True
      DidSuper = True
    End If

  Next i

End With
NoPhrase:
End Sub

Sub Ellipse(ByVal i%, Point#(), Oval#(), Nsegs%, _
  ByVal Emult#, Optional DidClip = False)
Dim Angle#, c1#, c2#, vx#, Vy#, test#, ConfLim#
Dim A#, b#, z#, RhoXY#, Denom#, SinAngle#, CosAngle#
Dim x0#, y0#, LastX0#, LastY0#, Xd#, Yd#, LastXd#
Dim LastYd#, X#, y#, Xerr#, Yerr#, Xpt#, Ypt#
Dim WithinBox As Boolean, LastWithinBox As Boolean, ns#, SegsMult#, ChiSqFact#
Dim j%, k%, r%, NumSegs%, Slope#, ThetaDeg#
Dim ba#, SqZ#, ThetaRad#, PhiRad#, DegreeStep#
Const Tiny = 0.000001, Niner = 0.999999, Bigg = 1E+32, Half = 0.5, Two = 2#
ReDim Oval(2, 1)

ViM DidClip, False
ConfLim = IIf(Emult = 2, 0.95, 0.6826) ' 95% conf ellipse if 2-sig input, 68.3% if 1-sig
Nsegs = 0
Xpt = Point(i, 1): Ypt = Point(i, 3)
If Xpt = 0 Then Xpt = Xspred / Million
If Ypt = 0 Then Ypt = Yspred / Million
'Xerr = Point(i, 2) * Emult: Yerr = Point(i, 4) * Emult
Xerr = Point(i, 2): Yerr = Point(i, 4)
If Abs(Xerr / Xpt) < Tiny Or Abs(Yerr / Ypt) < Tiny Then GoTo NoEllipse

ChiSqFact = App.ChiInv(1 - ConfLim, 2)
k = UBound(Point, 2)                       ' Rhoxy is the 5th column for XY data,
RhoXY = Point(i, 5 - 2 * (k = 9))          '  but the 7th for xyz data.
If Abs(RhoXY) >= 1 Then RhoXY = Sgn(RhoXY) * Niner
If Xerr = Yerr Then Yerr = 1.00001 * Xerr  ' To avoid various div-by-zero problems
vx = Xerr * Xerr: Vy = Yerr * Yerr
Denom = vx - Vy
Angle = Half * Atn(2 * RhoXY * Xerr * Yerr / Denom)
c1 = Two * (1 - SQ(RhoXY)) * ChiSqFact '/ 4
If Angle = 0 Then Angle = 1 / Bigg
c2 = 1 / Cos(2 * Angle)
test = ((1 + c2) / vx + (1 - c2) / Vy)

If test Then
  test = c1 / test
  If test <= 0 Then GoTo NoEllipse
Else
  test = Bigg
End If

A = Sqr(test)   ' Major axis length
test = ((1 - c2) / vx + (1 + c2) / Vy)

If test Then
  test = c1 / test
  If test <= 0 Then Exit Sub
Else
  test = Bigg
End If

b = Sqr(test)   ' Minor axis length

If DoShape Then

  Select Case Opt.CurveRes
    Case 1: ns = 22
    Case 2: ns = 32
    Case 3: ns = 48
  End Select

  ns = ns - 10 * Opt.ClipEllipse ' Compensate for noncurved lines if
  NumSegs = ns                   '  a clipped ellipse.
Else
  ns = 45 ' Default #line-segments in ellipse

  Select Case Opt.CurveRes
    Case 1: SegsMult = 0.6
    Case 2: SegsMult = 1
    Case 3: SegsMult = 2
  End Select

  Select Case Min(A / Xspred, b / Yspred)   ' More/less if larger/smaller
    Case Is < 0.02:  SegsMult = SegsMult / 2
    Case Is < 0.05:  SegsMult = SegsMult / 1.5
    Case Is < 0.1:   SegsMult = SegsMult / 1.2
    Case Is < 0.2
    Case Is < 0.5:   SegsMult = SegsMult * 1.5
    Case Is > 1:     SegsMult = 0
  End Select

  NumSegs = MinMax(8, 200 + 120 * DoShape, ns * SegsMult)
End If

DegreeStep = 360# / NumSegs
SinAngle = Sin(Angle): CosAngle = Cos(Angle)
r = 0: ba = b / A
ThetaDeg = -DegreeStep

For j = 1 To NumSegs + 1       ' Construct ellipse in polar coordinates
  ThetaDeg = ThetaDeg + DegreeStep
  ThetaRad = App.Radians(ThetaDeg)
  PhiRad = Atn(ba * iTan(ThetaRad))   ' Convert to accomodate (unrotated) ellipse shape.
  Slope = iTan(PhiRad)                        ' Of the radius vector.
  X = b / Sqr(Slope * Slope + ba * ba)
  If ThetaDeg > 90 And ThetaDeg < 270 Then X = -X

  If Abs(X) >= A Then
    y = 0
  Else
    y = b * Sqr(1 - SQ(X / A))
    If ThetaDeg > 180 Then y = -y
  End If

  x0 = Xpt + X * CosAngle - y * SinAngle
  y0 = Ypt + X * SinAngle + y * CosAngle
  Xd = x0: Yd = y0

  If Opt.ClipEllipse Then      ' Ignore any segments of ellipse outside of plotbox
    WithinBox = InBox2(x0, y0) '  but extend ellipse to boundaries of plotbox if larger.
    If Not WithinBox Then DidClip = True

    If j > 1 Then

      If WithinBox Then
        GoSub AddSeg

        If LastWithinBox Then  ' Both pts in plotbox
          Oval(1, r) = LastX0: Oval(2, r) = LastY0
        Else                   ' New point in plotbox, old point outside
          LastXd = LastX0: LastYd = LastY0
          PointsInBox x0, y0, LastXd, LastYd
          Oval(1, r) = LastXd: Oval(2, r) = LastYd
        End If

      ElseIf LastWithinBox Then  ' New point outside, last point inside
        PointsInBox Xd, Yd, LastX0, LastY0
        GoSub AddSeg
        Oval(1, r) = LastX0: Oval(2, r) = LastY0

      ElseIf InBox2(LastXd, LastYd) Then         ' New point outside, last point outside, but
        GoSub AddSeg
        Oval(1, r) = LastXd: Oval(2, r) = LastYd '  last-interpolated point inside

      End If

    End If

    LastX0 = x0: LastY0 = y0
    LastXd = Xd: LastYd = Yd
    LastWithinBox = WithinBox
  Else  ' Unclipped ellipse

    GoSub AddSeg
    Oval(1, r) = Xd: Oval(2, r) = Yd
  End If

Next j

If Opt.ClipEllipse Then

  If InBox2(Xd, Yd) Then   ' Close the ellipse back to its starting point
    GoSub AddSeg
    Oval(1, r) = Xd: Oval(2, r) = Yd
  End If

End If

If r > 0 Then Nsegs = r: Exit Sub

NoEllipse: Nsegs = 0
ReDim Oval(1, 2)
Exit Sub

AddSeg: r = r + 1
ReDim Preserve Oval(2, r)
Return
End Sub

Private Function InBox2(ByVal X#, ByVal y#) As Boolean ' Is an x-y point within the plot-box?
Dim Rx#, Ry#
InBox2 = False
Rx = Drnd(X, 7)
If Rx < Drnd(MinX, 7) Then Exit Function
If Rx > Drnd(MaxX, 7) Then Exit Function
Ry = Drnd(y, 7)
If Ry < Drnd(MinY, 7) Then Exit Function
If Ry > Drnd(MaxY, 7) Then Exit Function
InBox2 = True
End Function

Sub Tick(ByVal TickRange#, TikInterval#)
Dim b#

If TickRange = 0 Then
    MsgBox "Error in Tick: TickRange=0", vbCritical, Iso: ExitIsoplot
End If

TikInterval = 10 ^ zz(TickRange) / 8

While Abs(TickRange / TikInterval) > 12
  TikInterval = 2 * TikInterval
Wend

b = Abs(TikInterval) / 10 ^ zz(TikInterval)
If b <> Int(b) Then TikInterval = Int(b) * 10 ^ zz(TikInterval)
TikInterval = Drnd(TikInterval, 8)
End Sub

Sub IncreaseLineSpace(tB As Textbox, ByVal ExpFact)
Dim s$, i%, A%, sZ%
With tB
  s$ = .Text: A = 1                    ' Increase line-spacing in a textbox
  sZ = .Characters.Font.Size * ExpFact '  by increasing font-size of 1st
  Do                                   '  space in each line of the text-
    i = InStr(A, s$, " ")                '  string by a factor of ExpFact.
    If i Then .Characters(i, 1).Font.Size = sZ
    A = InStr(A + 1, s$, vbLf)
  Loop Until A = 0
End With
End Sub

Sub RemoveHdrFtr(Sht As Object)  ' Remove all headers/footers to the printed chart
With Sht
  On Error GoTo AfterHdrs ' In case no printer selected
  With .PageSetup
    'Error 999
    If Len(.LeftHeader) Then .LeftHeader = ""
    If Len(.CenterHeader) Then .CenterHeader = ""
    If Len(.RightHeader) Then .RightHeader = ""
    If Len(.LeftFooter) Then .LeftFooter = ""
    If Len(.CenterFooter) Then .CenterFooter = ""
    If Len(.RightFooter) Then .RightFooter = ""
    .ChartSize = xlScreenSize
    .Zoom = 100
    If .Orientation <> xlLandscape Then .Orientation = xlLandscape
    If ColorPlot And WideMargins Then
      ' Maximizes size of colored background to plotbox,
      '  resulting in larger plotbox and slower plot.
      StatBar "adjusting margins"
      If .LeftMargin Then .LeftMargin = 0
      If .RightMargin Then .RightMargin = 0
      If .TopMargin Then .TopMargin = 0
      If .BottomMargin Then .BottomMargin = 0
    End If
  End With
  ActiveWindow.Zoom = 100 ' To force no-bug in next zoom
  .ChartArea.Select
  With ActiveWindow
    .Zoom = True
    'If .Zoom > 90 Then .Zoom = 85
    ' Kluge to fix intermittent zooming to 400 (VBA bug)
  End With
  .Deselect
  On Error Resume Next ' in case not exist
  If .HasTitle Then .ChartTitle.Delete
End With
Exit Sub
AfterHdrs:
DelSheet
HandleNoPrinter
End Sub
Sub HandleNoPrinter()
MsgBox "You must select a printer to use with Excel before" _
  & vbLf & "you can make plots with Isoplot.", , Iso
ExitIsoplot
End Sub
Sub CenterPlotArea() ' Center the plot area in the physical sheet
Dim caL#, caR#, caT#, paL#, paR#
Dim caW#, caH#, paW#, paH#, paT#
Dim yaL#, xaH#
With Ach
  With .ChartArea
    caL = .Left: caW = .Width: caH = .Height: caT = .Top
  End With
  With .PlotArea
    paL = .Left: paW = .Width: paH = .Height: paT = .Top
  End With
  caR = caL + caW: paR = paL + paW
  With .Axes(xlCategory)
    If .HasTitle Then xaH = .AxisTitle.Font.Size
    xaH = xaH * 2 ' to account for uper/lower font space?
  End With
  With .Axes(xlValue)
    If .HasTitle Then yaL = .AxisTitle.Left
  End With
  .PlotArea.Left = paL + (caR - paR - yaL + caL) / 2
  '.PlotArea.Top = caT + (caH - paH - xaH) / 2 ' don't enable for now
End With
End Sub

Sub PlotErrorEnvelope(ByVal Slope, ByVal SlopeErr, ByVal Inter, ByVal InterErr, ByVal Xbar)
' Y = I + SX +- SQRT[SigmaI^2 + SigmaS^2 * X * (X - 2Xbar)]
Dim i%, M%, j%, k%, InBox As Boolean, X#, y#
Dim Changed As Boolean, ST#, Rad#, rCr As Range, cn%(2), cXY#()
Dim NextX#, NextY#, NextRad#, done(2) As Boolean, xy#()
Dim test1#, test2#, Bad As Boolean, Ie2#, Se2#
M = 60 * (1 - (2 - Opt.CurveRes) / 3)
ReDim xy(2, M + 4, 2)
ST = Xspred / M: SymbRow = Max(1, SymbRow)
For X = MinX To MaxX + 2 * ST Step ST
  NextX = X + ST
  Ie2 = SQ(InterErr): Se2 = SQ(SlopeErr)
  test1 = Ie2 + Se2 * X * (X - 2 * Xbar)
  test2 = Ie2 + Se2 * NextX * (NextX - 2 * Xbar)
  If test1 < 0 Or test2 < 0 Then Bad = True: Exit For
  Rad = Sqr(test1): NextRad = Sqr(test2)
  For j = 1 To 2
    If Not done(j) Then
      y = Inter + Slope * X + (3 - 2 * j) * Rad
      NextY = Inter + Slope * NextX + (3 - 2 * j) * NextRad
      PointsInBox X, y, NextX, NextY, InBox, Changed
      If InBox And cn(j) <= M Then
        cn(j) = 1 + cn(j)
        If cn(j) > 1 Then
          If Changed Or NextX >= MaxX Then
            xy(1, cn(j), j) = NextX
            xy(2, cn(j), j) = NextY
            done(j) = True
          Else
            xy(1, cn(j), j) = X: xy(2, cn(j), j) = y
          End If
        Else
          xy(1, cn(j), j) = X: xy(2, cn(j), j) = y
        End If
      End If
    End If
  Next j
Next X
If Bad Then
  MsgBox "Sorry -couldn't construct the error envelope.", , Iso
Else
  ChrtDat.Select
  For j = 1 To 2
    If cn(j) > 1 Then
      ReDim cXY(cn(j), 2)
      Set rCr = sR(SymbRow, SymbCol, cn(j) - 1 + SymbRow, 1 + SymbCol, ChrtDat)
      For i = 1 To cn(j)
        For k = 1 To 2
          cXY(i, k) = xy(k, i, j)
          rCr(i, k) = cXY(i, k)
      Next k, i
      AddSymbCol 2
      With IsoChrt
        .SeriesCollection.Add rCr, xlColumns, False, True, False
        .Select
        With Last(.SeriesCollection)
          .MarkerStyle = xlNone
          With .Border
            .LineStyle = xlContinuous: .Weight = xlHairline
            .Color = RGB(180, 0, 0)
          End With
        End With
      End With
    End If
  Next j
End If
End Sub
