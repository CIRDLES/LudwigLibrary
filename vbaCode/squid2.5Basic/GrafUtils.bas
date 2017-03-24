Attribute VB_Name = "GrafUtils"
'
'   ****************************************************************************
'
'   SQUID2 is a program for processing SHRIMP data
'
'   Program author: Dr Ken Ludwig (Berkeley Geochronology Center)
'
'   Supporters (members of the SQUID2 Development Group):
'       - Geoscience Australia
'       - United States Geological Survey
'       - Berkeley Geochronology Center
'       - All-Russian Institute of Geological Research
'       - Australian Scientific Instruments
'       - Geological Survey of Canada
'       - John de Laeter School of Mass Spectrometry (Curtin University)
'       - National Institute of Polar Research (Japan)
'       - Research School of Earth Sciences (Australian National University)
'       - Stanford University
'
'   ****************************************************************************
'
'   Copyright (C) 2009, the Commonwealth of Australia represented by Geoscience
'                 Australia, GPO box 378, Canberra ACT 2601, Australia
'   All rights reserved.
'   (http://www.ga.gov.au/minerals/research/methodology/geochron/index.jsp)
'
'   This file is part of SQUID2.
'
'   SQUID2 is free software. Permission to use, copy, modify, and distribute
'   this software for any purpose without fee is hereby granted under the terms
'   of the GNU General Public License as published by the Free Software
'   Foundation, either version 3 of the License, or (at your option) any later
'   version, provided that this notice is included in all copies of any
'   software which is, or includes, a copy or modification of this software and
'   in all copies of the supporting documentation for such software.
'
'   SQUID2 is distributed in the hope that it will be useful, but WITHOUT ANY
'   WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
'   FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
'   details.
'
'   You should have received a copy of the GNU General Public License along
'   with this program; if not see <http://www.gnu.org/licenses/gpl.html>.
'
'   ****************************************************************************
'
' 09/03/01 -- All lower array bounds explicit
' 09/05/03 -- In Sub AddSBMchart, prevent columns for SBM-chart data
'             from occupying a hidden column.
Option Explicit
Option Base 1

Sub ErrBars(ByVal SerName$, ByVal ErrRange$, ByVal Clr&, EndStyle%, ByVal ErrThick%)
' Add error bars to a chart's data-series

Dim p%, s$, d$, f$, SC As Series

On Error GoTo ErrBarsDone
s = Range(ErrRange).Address(, , xlR1C1)
f = "'" & ActiveSheet.Name & "'!"
d = f & s
Subst d, ",", "$"

Do
  p = InStr(d, "$")
If p = 0 Then Exit Do
  d = Left$(d, p - 1) & "," & f & Mid$(d, p + 1)
Loop

s = "=(" & d & ")"
Set SC = ActiveChart.SeriesCollection(SerName$)
FormatSeriesCol SC, , xlNone
FormatErrorBars SC, 2, s, Clr, ErrThick, IIf(EndStyle = xlCap, True, False)
ErrBarsDone: On Error GoTo 0
End Sub

Sub ChartShow(ByVal ChartNum%)
 ' Scroll the active window until SquidChart is at lower left
Dim Co%, dbl%, Rw&, rwH!, clW!, ChtH!, ChtL!

dbl = 0

Do
  dbl = 1 + dbl
  If dbl > peMaxCol Then GoTo NoSquidChart
Loop Until Cells(dbl, 1).Borders(xlBottom).LineStyle = xlDouble

On Error GoTo NoSquidChart
With ActiveSheet.ChartObjects("SquidChart" & fsS(ChartNum))
  ChtH = .Height + .Top: ChtL = .Left
End With
On Error GoTo 0
Rw = 1: Co = 1: clW = 0

With ActiveWindow
  .Activate
  .ScrollRow = 1: .ScrollColumn = 1
  rwH = .VisibleRange.Height

  Do
    Rw = 1 + Rw
    rwH = rwH + Cells(Rw - 1, 1).Height
    .SmallScroll 1
  Loop Until rwH > ChtH

  Do
    Co = 1 + Co
    clW = clW + Cells(1, Co).Width
    .ScrollColumn = Co
  Loop Until clW >= ChtL

' Remove pane-divider if not needed
'If .Panes(2).VisibleRange.Row <= .SplitRow Then .SplitRow = 0
End With
NoSquidChart:
End Sub

Sub AddSBMchart(ByVal HdrRow, SbmOffs#(), SbmOffsErr#())
' Add a chart showing the consistency of SBM readings for the different
'  peaks in the Run Table.

Dim NoErr As Boolean, tB As Boolean, yn$
Dim i%, j%, c%, wAcol%, ExclPos%, CommasPos%, ColNum%, PkNum%
Dim Arow&, RowNum&
Dim w!, h!, L!, t!, Xspred!, Xmin!, Xmax!, XaxFsize!, YaxFsize!
Dim Ymin#, Ymax#, Ystik#, Yltik#, Ym#
Dim ReDat As Range, ReO As Range, ReOerr As Range, reH As Range
Dim ReMa As Range, ReAll As Range, er As Range, SC As Series

If puTask.iNpeaks = 0 Then Exit Sub
c = 0

If pbUPb Then
  L = ActiveSheet.ChartObjects("squidchart1").Left

  Do
    c = c + 1
  Loop Until Cells(1, c + 1).Left > L

  If c > 8 Then
    wAcol = fvMax(1, c - 7): Arow = plaLastDatRw(1) + 4
  Else
    wAcol = 10 + fvMax(piaSacol(1) + 3, piaSacol(2))
    Arow = plaLastDatRw(1) + 7
  End If

Else
  Arow = 3 + plaLastDatRw(0)

  If puTask.iNeqns > 0 Then
    On Error GoTo 1
    wAcol = 3
  Else
1:  On Error GoTo 0
    wAcol = piLastCol + 1
  End If

  On Error GoTo 0
End If

Do
  Columns(wAcol).Select

  For ColNum = wAcol To wAcol + 3
    If Columns(ColNum).ColumnWidth < 3 Then wAcol = wAcol - 1: Exit For
  Next ColNum

Loop Until ColNum > (wAcol + 3) Or wAcol = 1

Do While Columns(wAcol - 2).Hidden
  wAcol = 1 + wAcol
Loop

If Not pbUPb Then
  Set ReAll = frSr(Arow, wAcol, Arow + puTask.iNpeaks - 1, wAcol + 5)
  With ReAll
    Set reH = Range(.Cells(0, 1), .Cells(0, 6))
    Set ReMa = Range(.Cells(1, 1), .Cells(puTask.iNpeaks, 1))
    Set ReO = Range(.Cells(1, 3), .Cells(puTask.iNpeaks, 3))
    Set ReOerr = Range(.Cells(1, 5), .Cells(puTask.iNpeaks, 5))
  End With
  BorderLine xlBottom, 1, Arow - 1, wAcol, , wAcol + 5
Else
  With puTask
    Set ReO = frSr(Arow, wAcol, .iNpeaks + Arow - 1) ' sbm %offs
    Set ReMa = frSr(Arow, wAcol - 2, .iNpeaks + Arow - 1) '  "    "    %err
    Set reH = frSr(Arow - 1, wAcol - 2, , wAcol + 3) ' Header
    Set ReOerr = frSr(Arow, wAcol + 2, Arow + .iNpeaks - 1) ' mass
    ReOerr(.iNpeaks, 1).Name = "ReOerr"
  End With
  Set ReAll = Union(reH, ReMa, ReO, ReOerr)
End If

Set ReDat = foAp.Union(ReMa, ReO)

With puTask
  frSr(Arow, wAcol - 2, Arow + .iNpeaks - 1, wAcol + 2).Name = "SBMChartDat"
  With ReO(0)
    .AddComment: .Comment.Text Text:="Mean %difference between the SBM beam " & _
       "at the indicated mass stations & the mass-averaged SBM beam"
    .Formula = "SBM %offs"
  End With
  ReOerr(0) = "2sigma err": ReMa(0) = "Mass"
  SigConv ReOerr(0)
  yn$ = fsQq("$+$0.00;$-$0.00;0")

  For RowNum = Arow - 1 To .iNpeaks + Arow
    For ColNum = ReMa.Column To ReOerr.Column Step 2
      frSr(RowNum, ColNum, , ColNum + 1).Merge
  Next ColNum, RowNum

  Fonts ReAll, , , , , , xlCenter, 11
  Fonts reH, , , , , True, xlCenter, 11

  For j = 1 To 3
    Box Arow, wAcol + 2 * (j - 1) + 2 * pbUPb, Arow + .iNpeaks - 1, wAcol + 2 * pbUPb + 2 * j - 1
  Next j

  Box reH.Row, reH.Column, reH.Row + .iNpeaks, ReOerr.Column + 1, RGB(0, 240, 240)
  Ymin = 1E+32: Ymax = -1E+32

  For PkNum = 1 To .iNpeaks
    ReO(PkNum) = SbmOffs(PkNum)
    ReOerr(PkNum) = SbmOffsErr(PkNum)
    ReMa(PkNum) = Drnd(pdaPkMass(PkNum), 4)
    Ymin = fvMin(Ymin, SbmOffs(PkNum) - SbmOffsErr(PkNum))
    Ymax = fvMax(Ymax, SbmOffs(PkNum) + SbmOffsErr(PkNum))
  Next PkNum

  ReO.NumberFormat = yn$: ReOerr.NumberFormat = pscZd2
  ReMa.NumberFormat = pscGen
  Ym = fvMax(Abs(Ymin), Abs(Ymax))

  Select Case Ym
    Case Is >= 2:   Ym = 1 + Int(Ym): Yltik = 1:   Ystik = 0.5
    Case Is >= 1:   Ym = 2:           Yltik = 1:   Ystik = 0.2
    Case Is >= 0.5: Ym = 1:           Yltik = 0.5: Ystik = 0.1
    Case Is >= 0.2: Ym = 0.4:         Yltik = 0.4: Ystik = 0.1
    Case Else:      Ym = 0.25:        Yltik = 0.25: Ystik = 0.05
  End Select

  If pbUPb Then
    L = Cells(1, ReMa.Column).Left + 30 * pbUPb
    t = Cells(ReMa.Row + .iNpeaks, 1).Top + 5
  Else
    L = 10 + ReAll.Left: t = 10 + fnBottom(ReAll)
  End If

End With

L = fvMax(L, 0)
w = Cells(1, ReOerr.Column + 2).Left - L + 52
h = 0.7 * w
ActiveSheet.ChartObjects.Add(L, t, w, h).Select
With Selection
  .Name = "SBM": .Placement = xlFreeFloating
End With

With ActiveChart
  .ChartType = xlXYScatter: .HasLegend = False
  .SetSourceData Source:=ReDat, PlotBy:=xlColumns

  With .ChartArea
    .Interior.Color = 13434879: .AutoscaleFont = False
    With .Border: .LineStyle = xlContinuous: .Weight = xlHairline: End With
  End With

  With .PlotArea
    .Interior.Color = 13434828 ' light green
    .Top = 0: .Left = 0: .Width = w: .Height = h
  End With

  With .Axes(xlValue)

    If Ym >= 4 Then
      .MinimumScaleIsAuto = True: .MaximumScaleIsAuto = True
    Else
      .MinimumScale = -Ym: .MaximumScale = Ym
      .MajorUnit = Yltik:  .MinorUnit = Ystik
    End If

    .MajorTickMark = xlCross:   .MinorTickMark = xlInside
    .HasMajorGridlines = False: .HasMinorGridlines = False
    .CrossesAt = .MinimumScale

    With .TickLabels:
      .AutoscaleFont = False
      .Font.Size = 11
      .Font.Name = psStdFont

      If Ym > 0.25 Then
        .NumberFormat = fsQq("$+$0.0;$-$0.0;0")
      End If

    End With

    .HasTitle = True
    With .AxisTitle
      With .Font
      .Size = 11 - pbUPb
      .Bold = False: .Name = psStdFont: End With
      .AutoscaleFont = False: .Caption = "SBM  %offset"
    End With

  End With

  With .Axes(xlCategory)
    .HasTitle = True
    Xmin = fvMin(ReMa): Xmax = fvMax(ReMa): Xspred = Xmax - Xmin
    .MinimumScale = fvMax(0, Drnd(Xmin - Xspred / 10, 3))
    .MaximumScale = Drnd(Xmax + Xspred / 10, 3)
    Xmin = fvMax(0, .MinimumScale): Xmax = .MaximumScale
    Xspred = Xmax - Xmin
    .MajorUnitIsAuto = True: .MinorUnitIsAuto = True

    With .TickLabels:
      .AutoscaleFont = False
      .Font.Size = 11
      .Font.Name = psStdFont
     End With

    .CrossesAt = .MinimumScale
    .MajorTickMark = xlCross:   .MinorTickMark = xlInside
    .HasMajorGridlines = False: .HasMinorGridlines = False

    With .AxisTitle
      With .Font: .Size = 11 - pbUPb
      .Bold = False: .Name = psStdFont: End With
      .AutoscaleFont = False: .Caption = "Mass Station"
    End With

  End With

  .PlotArea.Height = .ChartArea.Height - .PlotArea.Top - .Axes(1).AxisTitle.Font.Size
  .PlotArea.Left = .ChartArea.Left + .Axes(2).AxisTitle.Font.Size
  Set SC = .SeriesCollection(1)
  tB = (NoErr Or foAp.Average(ReOerr) < 0.05)

  For i = 1 To ReOerr.Count
    If ReOerr(i) <= 0 Then NoErr = True
  Next i

  FormatSeriesCol SC, , xlNone, , , , IIf(tB, xlCircle, xlNone), 6, vbRed, vbWhite

  If Not tB Then
    FormatErrorBars SC, 2, ReOerr, vbRed, xlThin
  End If

  XaxFsize = .Axes(1).AxisTitle.Font.Size
  YaxFsize = .Axes(2).AxisTitle.Font.Size
  .PlotArea.Left = YaxFsize + 6
  With ReMa: Set er = frSr(.Row + 1 + .Rows.Count, .Column, 2 + .Row + .Rows.Count, 1 + .Column): End With
  er(1, 1) = .Axes(1).MinimumScale: er(1, 2) = 0
  er(2, 1) = .Axes(1).MaximumScale: er(2, 2) = 0
  Range(er(1, 1), er(2, 2)).Font.Color = vbWhite
  ActiveChart.SeriesCollection.Add er, xlColumns

  With foLastOb(.SeriesCollection)
    .Border.Color = 0: .MarkerStyle = xlNone
  End With

End With

If pbSbmNorm Then
  If (fvMax(ReO) - fvMin(ReO)) > 1 Then

    If pbUPb Then
      L = [ReAll].Left
      t = [ReAll].Top
    Else
      L = ActiveSheet.Columns(wAcol).Left
      t = fnBottom(foLastOb(ActiveSheet.ChartObjects))
    End If

    With ActiveSheet.TextBoxes.Add(L, t, 1, 1)
      .AutoSize = True: .Interior.Color = vbRed
      .Text = fsVertToLF("SBM IS SCATTERED OR IMPRECISE|")
      .Text = .Text & "Suggest recalculating without SBM normalization"
      .Characters(Start:=InStr(.Text, "without"), Length:=7).Font.Underline = True

       With .Font
        .Name = psStdFont: .Size = 12
        .Bold = True
        .Color = vbWhite
      End With

      If pbUPb Then
        .Height = .Height + 5: .Width = .Width + 5
        .Top = [ReAll].Top - .Height * 1.1
        .Left = L - .Width / 2 + [ReAll].Width / 2
      End If

      .VerticalAlignment = xlTop: .HorizontalAlignment = xlCenter
      .Name = "Warning"

    End With

    With ActiveSheet.Shapes("warning").Shadow
      .Visible = msoTrue: .Type = msoShadow6: .ForeColor.SchemeColor = 23
    End With

  End If

End If

Cells(1, 1).Select
If Val(psExcelVersion$) >= 10 Then foAp.Calculate
End Sub

Sub SmallChart(Optional DataRange, Optional DataSheet, Optional ByVal PlaceSheet, _
 Optional Caption, Optional ByVal Xname, Optional ByVal Yname, Optional XaxisScale% = peAutoScale, _
 Optional ByVal XmajorTickMark% = xlCross, Optional ByVal YaxisScale% = peAutoScale, _
 Optional ByVal YmajorTickMark% = xlCross, Optional ByVal PlaceRow, Optional ByVal PlaceCol, _
 Optional ByVal ChtWidth! = 300, Optional ByVal ChtHeight! = 250, _
 Optional ByVal ChartBoxClr& = peStraw, Optional ByVal ChartBoxBrdr As Boolean = True, _
 Optional ByVal PlotBoxClr& = peLightGray, Optional ByVal XmajorGridlinesClr = -1, _
 Optional ByVal XminorGridlinesClr = -1, Optional ByVal YmajorGridlinesClr = -1, _
 Optional ByVal YminorGridlinesClr = -1, Optional ByVal Symbol% = xlCircle, _
 Optional ByVal SymbolSize! = 6, Optional ByVal SymbLineClr& = vbBlack, _
 Optional ByVal SymbInteriorClr& = vbWhite, Optional ByVal Transp% = 0, _
 Optional DataLineClr& = xlNone, Optional ByVal XerrBars As Boolean = False, _
 Optional XerrRange As Range, Optional ByVal YerrBars As Boolean = False, _
 Optional YerrRange As Range, Optional ByVal ErrBarsClr& = vbBlue, _
 Optional ByVal ErrBarsThick% = xlThin, Optional ByVal ErrBarsCap As Boolean = False, _
 Optional ByVal FontAutoScale As Boolean = True, Optional ByVal TikLabelSize% = 9, _
 Optional ByVal AxisNameSize = 11, Optional ByVal XerrCol% = 0, Optional ByVal YerrCol% = 0, _
 Optional BadPlot As Boolean = False, Optional PercentErrs = False)

' A general routine for constructing a small, customized chart-inset
' Assumes a contiguous source-data range.

Dim HasXerr As Boolean, HasYerr As Boolean
Dim LogX As Boolean, LogY As Boolean, ZeroX As Boolean, ZeroY As Boolean
Dim yErrAddr$
Dim Ncols%, Xcol%, Ycol%, DatCol%, NmajorTix%, LineStyle%
Dim i&, j&, rw2&, DatRow&, NumRows&, RowNum&, LineClr&
Dim Left!, Top!, t!, h!, hH!
Dim MinY#, MaxY#, MinX#, MaxX#
Dim xr As Range, yr As Range, Xerr As Range, Yerr As Range
Dim PlaceCell As Range, xyR As Range
Dim Cht As Chart, Xax As Object, Yax As Object, ChtA As Object
Dim SC As Series, PlotA As Object

NoUpdate
VIM DataSheet, ActiveSheet
DataSheet.Activate
VIM PlaceSheet, ActiveSheet
On Error GoTo 1
DataSheet.Activate

With DataRange
  VIM PlaceRow, .Row: VIM PlaceCol, .Column

  If .Areas.Count = 1 Then
    DatRow = .Row: DatCol = .Column: Xcol = DatCol
    NumRows = .Rows.Count: Ncols = .Columns.Count
  ElseIf .Areas.Count = 2 Then
    DatRow = .Row: DatCol = .Column
    Xcol = .Areas(1).Column: Ycol = .Areas(2).Column
    NumRows = .Rows.Count: Ncols = 2
  Else
    MsgBox "Bad data-range for x-y plot": BadPlot = True: Exit Sub
  End If

End With

rw2 = DatRow + NumRows - 1
Set xr = frSr(DatRow, Xcol, rw2, , DataSheet)

If DataRange.Areas.Count = 1 Then
  'Col2 = DatCol + Ncols - 1
  Ycol = IIf(XerrBars, Xcol + 2, Xcol + 1)
Else
  Ncols = 2
End If

Set yr = frSr(DatRow, Ycol, rw2)

If XerrBars Then
  HasXerr = True: If XerrCol = 0 Then XerrCol = 1 + Xcol
  Set Xerr = frSr(DatRow, XerrCol, rw2)
End If

If YerrBars Then
  HasYerr = True:  If YerrCol = 0 Then YerrCol = 1 + Ycol
  Set Yerr = frSr(DatRow, YerrCol, rw2, , DataSheet)

  If PercentErrs Then
    yErrAddr = Drnd(Yerr(1) / 100 * yr(1), 2)

    For i = 2 To Yerr.Count
      yErrAddr = yErrAddr & "," & Drnd(Yerr(i) / 100 * yr(i), 2)
    Next i

  Else
    yErrAddr = "=" & ActiveSheet.Name & "!" & Yerr.Address(ReferenceStyle:=xlR1C1)
  End If

End If

If DataRange.Areas.Count = 1 Then
  Set xyR = Union(xr, yr)
Else
  Set xyR = Range(xr.Address & "," & yr.Address)
End If

Set PlaceCell = PlaceSheet.Cells(PlaceRow, PlaceCol)

For RowNum = 1 To NumRows
  If Not IsNumeric(xr(RowNum)) Then xr(RowNum) = ""

  If HasXerr Then
    If Not IsNumeric(Xerr(RowNum)) Then Xerr(RowNum) = ""
  End If

  If HasYerr Then
    If Not IsNumeric(yr(RowNum)) Then Yerr(RowNum) = ""
  End If

  If Not IsNumeric(yr(RowNum)) Then yr(RowNum) = ""
Next RowNum


MinY = fvMin(yr): MaxY = fvMax(yr)
MinX = fvMin(xr): MaxX = fvMax(xr)
LogX = (XaxisScale = peLogScale)
LogY = (YaxisScale = peLogScale)
ZeroX = (XaxisScale = peZeroMinScale)
ZeroY = (YaxisScale = peZeroMinScale)
If MinY <= 0 Then LogY = False: ZeroY = False
If MinX <= 0 Then LogX = False: ZeroX = False
PlaceSheet.Activate

With Charts.Add
  .ChartType = xlXYScatter
  .SetSourceData Source:=xyR, PlotBy:=xlColumns
  .Location Where:=xlLocationAsObject, Name:=PlaceSheet.Name
  GoTo 2
End With

1: On Error GoTo 0
MsgBox "Unable to plot chart", , pscSq
BadPlot = True: Exit Sub
2: On Error GoTo 0
Set Cht = ActiveChart

With foLastOb(ActiveSheet.ChartObjects)
  .Left = PlaceCell.Left: .Top = PlaceCell.Top
  .Width = ChtWidth: .Height = ChtHeight
  .Interior.Color = ChartBoxClr
  If Not ChartBoxBrdr Then .Border.LineStyle = xlNone
End With

With Cht
  Set ChtA = .ChartArea
  Set PlotA = .PlotArea
  .HasLegend = False
  .HasTitle = False
  LineClr = IIf(DataLineClr = xlNone, xlNone, vbRed)
  LineStyle = IIf(LineClr = xlNone, xlNone, xlContinuous)
  Set SC = .SeriesCollection(1)
  FormatSeriesCol SC, , LineStyle, LineClr, xlThin, False, Symbol, _
                  SymbolSize, SymbLineClr, SymbInteriorClr

  If XerrBars Then
    FormatErrorBars SC, 1, Xerr, ErrBarsClr, ErrBarsThick, ErrBarsCap
  End If
  If YerrBars Then
    FormatErrorBars SC, 2, Yerr, ErrBarsClr, ErrBarsThick, ErrBarsCap
  End If

  On Error GoTo 0
  Set Xax = .Axes(1): Set Yax = .Axes(2)

  With Yax
    .HasTitle = (Yname <> "")

    If .HasTitle Then
      .AxisTitle.Characters.Text = Yname
      With .AxisTitle
        .AutoscaleFont = FontAutoScale
        With .Font
          .Name = "Arial": .FontStyle = "Regular": .Size = AxisNameSize
        End With
      End With
    End If

    .ScaleType = IIf(LogY, xlScaleLogarithmic, xlScaleLinear)
    .MaximumScaleIsAuto = True
    .MinimumScaleIsAuto = True
    If MinY < 0 Then ZeroY = False: .CrossesAt = .MinimumScale
    If MinY >= 0 And .MinimumScale < 0 Then .MinimumScale = 0: MinY = 0

    If LogY Then
      MinY = .MinimumScale: MaxY = .MaximumScale
      j = Int(fdLog10(MaxY))
      MaxY = 10 ^ (1 + j): MinY = 10 ^ j
      .MinorUnit = 10: .MajorUnit = 10
      .MinorTickMark = xlCross

    ElseIf ZeroY Then
      .MinimumScale = 0: .CrossesAt = 0: .MinorTickMark = xlTickMarkInside

    ElseIf MinY >= 0 Then

      If .MinimumScale > 0 Then
        .MinimumScaleIsAuto = True
      Else

        If HasYerr Then
          AxisScale yr, False, Yerr
        Else
          AxisScale yr, False
        End If

      End If

      .MinorTickMark = xlTickMarkInside
      .CrossesAt = .MinimumScale
    End If

    .HasMajorGridlines = (YmajorGridlinesClr >= 0)
    If .HasMajorGridlines Then .MajorGridlines.Border.Color = YmajorGridlinesClr
    .HasMinorGridlines = (YminorGridlinesClr >= 0)
    If .HasMinorGridlines Then .MinorGridlines.Border.Color = YminorGridlinesClr
    .MajorTickMark = YmajorTickMark

    If .MajorTickMark = xlTickMarkNone Then
      .TickLabelPosition = xlNone
       .MajorUnit = xlAutomatic
       .MinorUnit = xlAutomatic

    ElseIf Not LogY Then
      NmajorTix = (.MaximumScale - .MinimumScale) / .MajorUnit

      If NmajorTix > 7 Then
        .MajorUnit = 2 * .MajorUnit
        .MinorUnit = .MajorUnit / 2
        .CrossesAt = .MinimumScale
      End If

    End If
    .MajorTickMark = xlCross
    .MinorTickMark = xlInside

    With .TickLabels
      .AutoscaleFont = FontAutoScale
      With .Font
        .Name = "Arial": .FontStyle = "Regular": .Size = TikLabelSize ' pass as variable?
      End With
      .NumberFormat = pscGen

    End With
    If .MinimumScale = 0 And Not LogY Then
      .MaximumScale = .MaximumScale * 1.1 ' for headroom
      .MaximumScaleIsAuto = True
    End If

  End With ' Yname axis

  With Xax
    .HasTitle = (Xname <> "")

    If .HasTitle Then
      .AxisTitle.Characters.Text = Xname
      With .AxisTitle
        .AutoscaleFont = FontAutoScale
        With .Font
          .Name = "Arial": .FontStyle = "Regular": .Size = AxisNameSize ' pass as variable?
        End With
      End With
    End If

    If MinX >= 0 And .MinimumScale < 0 Then .MinimumScale = 0: MinX = 0
    If MinX < 0 Then ZeroX = False: .CrossesAt = .MinimumScale
    .ScaleType = IIf(LogX, xlScaleLogarithmic, xlScaleLinear)

    If LogX Then
      MinX = .MinimumScale: MaxX = .MaximumScale
      .MinorUnit = 10: .MajorUnit = 10
    ElseIf ZeroX Then
      .MinimumScale = 0: .CrossesAt = 0
    ElseIf HasXerr Then
      AxisScale xr, True, XerrRange
    Else
      AxisScale xr, True
    End If

    .HasMajorGridlines = (XmajorGridlinesClr >= 0)
    If .HasMajorGridlines Then .MajorGridlines.Border.Color = XmajorGridlinesClr
    .HasMinorGridlines = (XminorGridlinesClr >= 0)
    If .HasMinorGridlines Then .MinorGridlines.Border.Color = XminorGridlinesClr
    .MajorTickMark = XmajorTickMark

    If .MajorTickMark = xlTickMarkNone Then
      .TickLabelPosition = xlNone
      If Not LogX Then .MajorUnit = .MaximumScale - .MinimumScale
    Else
      If Not LogX Then
        NmajorTix = (.MaximumScale - .MinimumScale) / .MajorUnit

        If NmajorTix > 7 Then
          .MajorUnit = 2 * .MajorUnit
          .MinorUnit = .MajorUnit / 2
          .CrossesAt = .MinimumScale
        End If

      End If
    End If

    .MajorTickMark = xlCross
    .MinorTickMark = xlInside

    With .TickLabels
      .AutoscaleFont = FontAutoScale
      With .Font
        .Name = "Arial": .FontStyle = "Regular"
        .Size = TikLabelSize ' pass as variable?
      End With
      .NumberFormat = pscGen
    End With
  End With ' Xname-axis

  With Yax
    If .HasTitle Then .AxisTitle.Font.Size = AxisNameSize
  End With

  With Xax
    If .HasTitle Then .AxisTitle.Top = .AxisTitle.Top + 10
  End With

  NumFormatTicks Xax
  NumFormatTicks Yax

  With .PlotArea
    .Interior.Color = PlotBoxClr
    .Top = 0: t = 0: .Height = .Height + 15: h = .Height
    .Left = .Left - 5: .Width = .Width + 10
  End With

  If fbNIM(Caption) Then
    If Caption <> "" Then
      .HasTitle = True

      With .ChartTitle
        .AutoscaleFont = False
        .Caption = Caption
        With .Font
          .Size = 11: .Bold = False
        End With
        .HorizontalAlignment = xlRight
        .Left = Cht.ChartArea.Width: .Top = 0
      End With

    End If

    With Cht
      .PlotArea.Top = 12
      .PlotArea.Height = .ChartArea.Height
    End With
  End If

  If Xax.HasTitle Then
    With .PlotArea
      hH = Xax.AxisTitle.Font.Size
      h = Xax.TickLabels.Font.Size
      .Top = 0
      .Height = ChtA.Height - hH - h
    End With
  End If

  ActiveSheet.Cells(DatRow, DatCol).Activate
End With

End Sub

Sub CreateAutoCharts(Optional Bad As Boolean = False)
' Construct all AutoCharts specified by the current Task and place on the output sheet.
Dim SeparateSht As Boolean, Xerrs As Boolean, Yerrs As Boolean
Dim tB1 As Boolean, tB2 As Boolean, Ypers As Boolean
Dim s$, Ms$, ChtShtNa$
Dim SerColNum%, i%, k%, c%, Nchts%, Xcol%, Ycol%, LeftOffs%, TopOffs%
Dim Nplots%, YerCol%, Xscale%, Yscale%, ChtNum%, LineWt%, YxCt%
Dim FirstDatRow&, LastDatRow&, tRw&, p1&, RowNum&
Dim ChtLeft!, ChtTop!
Dim Slp#, Inter#, LwrSlp#, UpperSlp#, Udelt#, Ldelt#
Dim Yav#, YavPer#, Xmin#, Xmax#, Ymin#, Ymax#
Dim L As Range, xL As Range, yL As Range, Xrange As Range, DatRangeIn As Range
Dim Yrange As Range, Sra As Range, AbsYers As Range, TempRange As Range, DatRangeUsed As Range
Dim SourceSht As Worksheet, DestSht As Worksheet, Sht As Worksheet
Dim Cht As Chart, SC As Series
Dim PA As Autocharts, LastCht As ChartObject, ChtOb As ChartObject, ChtObs As ChartObjects

If pbUPb Then pbStd = True
YxCt = 0
Set SourceSht = IIf(pbStd, phStdSht, phSamSht)
SourceSht.Activate
SeparateSht = foUser("SeparateAutochtSht")
Set DestSht = SourceSht
Ms = "Can only plot positive values on Log axes"
ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
Nplots = puTask.iNumAutoCharts
FirstDatRow = 1 + flHeaderRow(pbStd): LastDatRow = plaLastDatRw(-pbStd)
Nchts = ActiveSheet.ChartObjects.Count

On Error GoTo 0

For ChtNum = 1 To Nplots
  On Error GoTo 0
  StatBar "Creating Auto-Chart" & StR(ChtNum)
  PA = puTask.uaAutographs(ChtNum)
  SourceSht.Activate
  With PA
    FindStr Phrase:=.sXname, ColFound:=Xcol, RowLook1:=flHeaderRow(pbStd), WholeWord:=True
    FindStr Phrase:=.sYname, ColFound:=Ycol, RowLook1:=flHeaderRow(pbStd), WholeWord:=True
    Bad = False

    If Xcol > 0 And Ycol > 0 Then
      Set Xrange = frSr(FirstDatRow, Xcol, LastDatRow)
      Set Yrange = frSr(FirstDatRow, Ycol, LastDatRow)
      On Error GoTo 2
      If foAp.Count(Xrange) = 0 Or foAp.Count(Yrange) = 0 Then GoTo 2
      If foAp.Sum(Xrange) = 0 Or foAp.Sum(Yrange) = 0 Then GoTo 2
      On Error GoTo 0

      If .bLogX Then
        For RowNum = 1 To Xrange.Rows.Count
          If Xrange(RowNum, 1) <= 0 Then
            MsgBox Ms, , pscSq: Bad = True: GoTo 2
          End If
        Next RowNum
      End If

      If .bLogY Then
        For RowNum = 1 To Yrange.Rows.Count
          If Yrange(RowNum, 1) <= 0 Then
            MsgBox Ms, , pscSq: Bad = True: GoTo 2
          End If
        Next RowNum
      End If

      Set DatRangeIn = Range(Xrange.Address & "," & Yrange.Address)
      N = Xrange.Rows.Count

      If .bAutoscaleX Then
        Xscale = peAutoScale
      ElseIf .bLogX Then
        Xscale = peLogScale
      ElseIf .bZeroXmin Then
        Xscale = peZeroMinScale
      End If

      If .bAutoscaleY Then
        Yscale = peAutoScale
      ElseIf .bLogY Then
        Yscale = peLogScale
      ElseIf .bZeroYmin Then
        Yscale = peZeroMinScale
      End If

      s = LCase(Yrange(0, 2))
      tB1 = InStr(s, "err")
      tB2 = InStr(s, "%")
      Yerrs = (.bAverage And tB1) ' average plot with y-errors
      Ypers = tB1 And tB2         ' percent errors
      With Yrange
        If Yerrs Then
          YerCol = 0

          For RowNum = plaFirstDatRw(1) To plaLastDatRw(1)
            YerCol = fvMax(YerCol, 1 + fiEndCol(RowNum))
          Next RowNum

          Set AbsYers = frSr(.Row, YerCol, .Row + .Rows.Count - 1)

          For RowNum = 1 To N
            s = "=" & Yrange(RowNum, 2).Address(0, 0)
            If Ypers Then s = s & "*" & Yrange(RowNum, 1).Address(0, 0) & " / 100"
            AbsYers(RowNum) = s
          Next RowNum

        Else
          YerCol = Ycol + 1
        End If

      End With

      If ChtNum = 1 And SeparateSht Then

        Do
          tB1 = True

          For Each Sht In ActiveWorkbook.Worksheets
            s = LCase(Left$(Sht.Name, 10))

            If s = "autocharts" Then
              tB1 = False: Sht.Delete
              Exit For
            End If

          Next Sht

        Loop Until tB1

        Sheets.Add
        ChtShtNa = "Autocharts"
        Sheetname ChtShtNa
        ActiveSheet.Name = ChtShtNa
        NoGridlines
        ActiveWindow.Zoom = 75
        Set DestSht = ActiveSheet
      End If

      Set DatRangeUsed = DatRangeIn

      With DatRangeIn

        If Ycol < Xcol Then ' Excel charts require that the x-column be higher-numbered
          YxCt = 1 + YxCt   '  than the y-column, so must make a separate data range.
          ' Put the chart x-y data in cols 2 & 3, starting 60 rows below the last data-line.
          tRw = fvMax(flEndRow(3), flEndRow(2)) + IIf(YxCt = 1, 60, 3)
          Set DatRangeUsed = frSr(tRw, 2, tRw + N - 1, 3)
          With DatRangeUsed ' Make font barely visible to avoid distraction
            .NumberFormat = "General"
            .Font.Color = Hues.peLightGray
            .Font.Size = 8
          End With

          For i = 1 To N
            DatRangeUsed(i, 1) = "=" & Xrange(i, 1).Address(0, 0)
            DatRangeUsed(i, 2) = "=" & Yrange(i, 1).Address(0, 0)
          Next i

        End If

      End With

      SmallChart DataRange:=DatRangeUsed, DataSheet:=SourceSht, PlaceSheet:=DestSht, Caption:="", _
       Xname:=.sXname, Yname:=.sYname, XaxisScale:=Xscale, XmajorTickMark:=xlInside, YaxisScale:=Yscale, _
       YmajorTickMark:=xlInside, PlaceRow:=1, PlaceCol:=4 + ChtNum, ChtWidth:=280, ChtHeight:=210, _
       ChartBoxClr:=RGB(192, 128, 255), ChartBoxBrdr:=True, PlotBoxClr:=RGB(255, 255, 128), _
       XmajorGridlinesClr:=-1, XminorGridlinesClr:=-1, YmajorGridlinesClr:=-1, YminorGridlinesClr:=-1, _
       Symbol:=xlCircle, SymbolSize:=6, SymbLineClr:=vbRed, SymbInteriorClr:=vbWhite, _
       Transp:=0, DataLineClr:=xlNone, XerrBars:=Xerrs, YerrBars:=Yerrs, YerrRange:=AbsYers, _
       ErrBarsClr:=vbRed, ErrBarsThick:=xlThin, ErrBarsCap:=(Yerrs And N < 15), FontAutoScale:=False, _
       TikLabelSize:=11, AxisNameSize:=12, YerrCol:=YerCol, BadPlot:=Bad

      If Bad Then GoTo 2

      If Yerrs Then ' Sam or Std sheet???
        Fonts AbsYers, , , , vbWhite, , , 8, , , , General
        phStdSht.Columns(YerCol).ColumnWidth = 0.5
      End If

      k = ActiveSheet.ChartObjects.Count

      If k > Nchts Or SeparateSht Then
        Set ChtOb = foLastOb(ActiveSheet.ChartObjects)
        Set Cht = ChtOb.Chart
        Nchts = 1 + Nchts

        If (Yscale <> peLogScale And Xscale <> peLogScale) _
            And (.bRegress Or .bAverage) Then
          ReDim xy(1 To N, 1 To 2)

          For RowNum = 1 To N
            xy(RowNum, 1) = Xrange(RowNum, 1): xy(RowNum, 2) = Yrange(RowNum)
          Next RowNum

          If .bRegress And N > 2 Then
            Isoplot3.RobustReg2 xy, Slp, LwrSlp, UpperSlp, Inter
            Udelt = Drnd(UpperSlp - Slp, 2): Ldelt = Drnd(Slp - LwrSlp, 2)
            If Abs(Udelt / Slp) < 0.00001 Then Udelt = 0
            If Abs(Ldelt / Slp) < 0.00001 Then Ldelt = 0
          Else

            If Yerrs Then
              ReDim bw(1 To 7, 1 To 1)
              phStdSht.Activate
            End If

            If .bAverage Then
              ReDim bw(1 To 3, 1 To 1)
              bw = Isoplot3.biweight(Yrange, 6)
              Yav = bw(1, 1): YavPer = 100 * bw(3, 1) / Yav
            End If

          End If

          DestSht.Select
          With Cht
            Xmin = .Axes(1).MinimumScale:   Xmax = .Axes(1).MaximumScale
            Ymin = .Axes(2).MinimumScale:   Ymax = .Axes(2).MaximumScale
            c = 2 + 2 * (ChtNum - 1) - 97 * SeparateSht

            If SeparateSht Then
              p1 = flEndRow(99 + c)
            Else
              p1 = fvMax(flEndRow(c), flEndRow(c + 1), flEndRow(c + 2)) + 2
            End If

            With frSr(p1, c, p1 + 1, c + 1)
              .NumberFormat = "General"
              .Font.Color = Hues.peLightGray
            End With
            Cells(p1, c) = Xmin
            Cells(p1 + 1, c) = Xmax

            If puTask.uaAutographs(ChtNum).bRegress Then
              Cells(p1, c + 1) = Inter + Xmin * Slp
              Cells(p1 + 1, c + 1) = Inter + Xmax * Slp
            Else
              Cells(p1, c + 1) = bw(1, 1)
              Cells(p1 + 1, c + 1) = bw(1, 1)
            End If

            Set L = frSr(p1, c, p1 + 1, c + 1)
            Set xL = frSr(p1, c, p1 + 1)
            Set yL = frSr(p1, c + 1, p1 + 1)
            s = DestSht.Name
            .SeriesCollection.Add Source:=DestSht.Range(L.Address), Rowcol:=xlColumns, _
              SeriesLabels:=False, CategoryLabels:=True, Replace:=False
            Set SC = foLastOb(.SeriesCollection)

            LineWt = IIf(puTask.uaAutographs(ChtNum).bRegress, xlThin, xlMedium)
            FormatSeriesCol SC, , xlContinuous, vbBlack, LineWt, , xlNone
            .Axes(1).MinimumScale = Xmin
            .Axes(1).MaximumScale = Xmax
            .Axes(1).CrossesAt = Xmin
            .Axes(2).MinimumScale = Ymin
            .Axes(2).MaximumScale = Ymax
            .Axes(2).CrossesAt = Ymin
            With .TextBoxes.Add(.Axes(1).Left + 5, .Axes(2).Top, 20, 15)
              .AutoSize = True: .Font.Size = 11

              If puTask.uaAutographs(ChtNum).bRegress Then
                s = "Slope=" & StR(Drnd(Slp, 3)) & "  +" & _
                    fsS(Drnd(Udelt, 2)) & "  " & StR(-Drnd(Ldelt, 2)) & _
                    "    Inter= " & StR(Drnd(Inter, 3))
              Else
                s = "Mean " & puTask.uaAutographs(ChtNum).sYname & _
                    " = " & fsS(Drnd(Yav, 4)) & _
                    " " & pscPm & fsS(Drnd(YavPer, 2)) & "% (95%conf)"
                'If YavProb >= 0.05 Then s = s & "  MSWD=" & fsS(Drnd(YavMSWD, 2))
              End If

              .Text = s: .HorizontalAlignment = xlCenter
              .Top = 0: .Left = 0: .Width = Cht.ChartArea.Width
              .Interior.Color = RGB(192, 128, 255)
            End With 'textbox
          End With 'cht

        End If

        With Cht
          .Axes(1).CrossesAt = .Axes(1).MinimumScale
          .Axes(2).CrossesAt = .Axes(2).MinimumScale
        End With

        ChtOb.Name = "Autochart" & fsS(ChtNum)
        Set ChtObs = ActiveSheet.ChartObjects

        If SeparateSht Then
          LeftOffs = (ChtNum - 1) Mod 4
          ChtLeft = 10 + LeftOffs * (ChtOb.Width + 10)
          TopOffs = (ChtNum - 1) \ 4
          ChtTop = TopOffs * (ChtOb.Height + 10)

        ElseIf pbUPb Then

          If ChtNum = 1 Then
            If fbChartExist("sbm") Then
              ChtTop = 8 + fnBottom(ChtObs("sbm"))
              ChtLeft = ChtObs("sbm").Left
            ElseIf fbRangeNameExists("sbmchartdat") Then
              Set Sra = Range("SBMchartdat")
              ChtTop = 8 + fnBottom(Sra)
              ChtLeft = Sra.Left + Sra.Width / 2 - ChtOb.Width / 2
            ElseIf fbChartExist("squidchart1") Then
              ChtTop = fnBottom(Rows(2 + plaLastDatRw(1)))
              ChtLeft = ChtObs("squidchart1").Left - ChtOb.Width - 3
            Else
              ChtTop = fnBottom(Rows(2 + plaLastDatRw(1)))
              ChtLeft = Range("ExtPerrA1").Left - ChtOb.Width
            End If

          Else
            ChtTop = fnBottom(ActiveSheet.ChartObjects(Nchts - 1)) + 5
          End If

        Else
          ChtTop = Rows(2 + flEndRow(1)).Top
          On Error Resume Next

          If ChtNum = 1 Then

            If pbSbmNorm Then
              ChtLeft = fnRight(ActiveSheet.Shapes("sbm")) + 20
            Else
              ChtLeft = Columns(2).Left
            End If

          Else
            ChtLeft = fnRight(LastCht) + 7
          End If

          On Error GoTo 0

          If .bRegress Or .bAverage Then

            For c = 1 To peMaxCol
              If Columns(c).Left > ChtLeft Then Exit For
            Next c

            frSr(p1, 2, p1 + 1, 3).Cut Cells(flEndRow(1) + 4, c)
          End If

        End If

        With ChtOb
          .Left = ChtLeft
          .Top = ChtTop
        End With
        Set LastCht = ChtOb
        On Error GoTo 0
      End If ' xcol

2:  On Error GoTo 0

    End If 'plots(ChtNum)
  End With

Next ChtNum
SourceSht.Activate
StatBar
End Sub

Sub SBMdata(SbmOffs#(), SbmPk#(), SbmOffsErr#(), SbmDeltaPcnt#(), CanChart As Boolean)
' Examine SBM counts to see if any of the mass-stations have a
'  consistent SBM offset from the mean SBM signal.
Dim PkNum%, SpotNumber%
CanChart = True

For PkNum = 1 To puTask.iNpeaks

  For SpotNumber = 1 To piaSpotCt(-pbStd)
    SbmPk(SpotNumber) = SbmDeltaPcnt(PkNum, SpotNumber)
  Next SpotNumber

  SbmOffs(PkNum) = 0
  On Error Resume Next  ' Just in case
  SbmOffs(PkNum) = foAp.Average(SbmPk)

  If fbNoNum(SbmOffs(PkNum)) Then
    SbmOffs(PkNum) = 0
  End If

  On Error GoTo 0
  If SbmOffs(PkNum) = 0 Then
    Erase SbmOffs: CanChart = False: Exit For
  End If

  On Error Resume Next
  SbmOffsErr(PkNum) = StudentsT(piaSpotCt(-pbStd) - 1) * foAp.StDev(SbmPk) / sqR(piaSpotCt(-pbStd))
  On Error GoTo 0

Next PkNum
End Sub

Function fbChartExist(ByVal ChtName$) As Boolean
Dim a! ' Does the specified chart exist in the active sheet?
On Error GoTo 1
a = ActiveSheet.ChartObjects(ChtName).Left
fbChartExist = True: Exit Function
1: fbChartExist = False
End Function

Sub NumFormatTicks(Ax As Axis)
' Format the axis-tick labels of the active chart, minimizing the number
'  of decimal places used.
Dim Tik$, PosChar$, Num1$, Num2$
Dim DecPos%, Ndec%, MaxDec%
Dim MinAx#, MaxAx#, TikVal#

With Ax
  MinAx = .MinimumScale
  MaxAx = .MaximumScale
  MaxDec = 0

  For TikVal = MinAx To MaxAx Step .MajorUnit
    Tik = fsS(Drnd(TikVal, 6))
    DecPos = InStr(Tik, ".")

    If DecPos > 0 Then
      Ndec = Len(Mid$(Tik, 1 + DecPos))
    Else
      Ndec = 0
    End If

    If Ndec < 9 And Abs(TikVal) > 0.0000000001 Then
      MaxDec = fvMax(Ndec, MaxDec)
    End If

  Next TikVal

  If MaxDec > 0 Then
    PosChar = IIf(MinAx < 0 And MaxAx > 0, "+", "")
    Num1 = "0." & String(MaxDec, "0")
    Num2 = PosChar & Num1 & ";-" & Num1 & ";0"
    .TickLabels.NumberFormat = Num2
  Else
    .TickLabels.NumberFormat = "General"
  End If

End With
End Sub

Sub FormatSeriesCol(SerCol As Series, Optional Point, Optional ByVal LineStyle, _
  Optional ByVal LineClr = 0, Optional ByVal LineThick = xlThin, _
  Optional ByVal LineSmooth As Boolean = True, _
  Optional ByVal MarkerStyle = xlNone, _
  Optional ByVal MarkerSize = 6, Optional ByVal MarkerForeClr& = vbRed, _
  Optional ByVal MarkerBackClr = vbWhite)
' Format a chart's SeriesCollection as specified by the input parameters.
Dim SCP As Object

If IsMissing(Point) Then
  Set SCP = SerCol
Else
  Set SCP = SerCol.Points(Point)
End If

With SCP

  If Not IsMissing(LineStyle) Then
    With .Border
      .LineStyle = LineStyle

      If LineStyle <> xlNone Then
        .Color = LineClr
        .Weight = LineThick
      End If

    End With
    .Smooth = LineSmooth
  End If

  If Not IsMissing(MarkerStyle) Then .MarkerStyle = MarkerStyle

  If MarkerStyle <> xlNone Then
    If Not IsMissing(MarkerSize) Then .MarkerSize = MarkerSize
    If Not IsMissing(MarkerForeClr) Then .MarkerForegroundColor = MarkerForeClr
    If Not IsMissing(MarkerBackClr) Then .MarkerBackgroundColor = MarkerBackClr
  End If

End With
End Sub

Sub AddTextbox(Phrase$, FontSize, Optional FontBold, Optional Left, Optional Top, _
  Optional BkrdClr, Optional HorizAlign, Optional VertAlign, Optional AutoscaleFont)
' Add a TextBox to a chart.
Dim Hal%, Val%, Bclr&, L!, t!, Ch As Chart

VIM Left, 0
VIM Top, 0
VIM BkrdClr, -1
VIM HorizAlign, xlLeft
VIM VertAlign, xlBottom
VIM FontBold, False
VIM AutoscaleFont, False
Set Ch = ActiveChart
Ch.TextBoxes.Add Left, Top, 1, 1

With foLastOb(Ch.TextBoxes)
  .AutoSize = True
  .Text = Phrase
  .Font.Size = FontSize
  .Font.Bold = FontBold
  .AutoscaleFont = AutoscaleFont

  If BkrdClr < 0 Then
    BkrdClr = ActiveChart.PlotArea.Interior.Color
  ElseIf BkrdClr = 0 Then
    .ShapeRange.Fill.Visible = False
  End If

  If BkrdClr <> 0 Then .Interior.Color = BkrdClr
  .HorizontalAlignment = HorizAlign
  .VerticalAlignment = VertAlign

  If Left < 0 Then
    .Left = Ch.Axes(1).Left + Ch.Axes(1).Width / 2 - .Width / 2
  End If

  If Top < 0 Then
    .Top = Ch.Axes(2).Top + Ch.Axes(2).Height / 2 - .Height / 2
  End If

End With
End Sub

Sub TrimMassStuff(ByVal Row&)
' Construct chart-insets for the secular variation of the Trim Masses of
'  the loaded raw-data file.
Dim Bad As Boolean, PkTest() As Boolean, s$
Dim k%, c%, m%, PkNum%, GrafNum%, ShapeNum%, GrafCt%, TrimNum%, nSh%, ColCtr%, ChartObjectsN%, Nchts%, Tt%
Dim RowCtr&, r&
Dim L!, w!, t0!, t!, WindR!, h!, LL!, wW!, hH!
Dim MinX#, MaxX#, MinY#, MaxY#, m1#, m2#, GrafPk#()
Dim Cht As Object, ChtOb As Object, Xax As Object, Yax As Object, Axis As Object
Dim tm As Worksheet, PlotDestSht As Worksheet
Dim ChO As ChartObject, Dr As Range

With puTask
  ReDim tma#(1 To piTrimCt, 1 To .iNpeaks), tta#(1 To piTrimCt, 1 To .iNpeaks)
  ReDim PkTest(1 To .iNpeaks)
  s = "Trim Masses"

  If pbUPb Then
    Set PlotDestSht = phStdSht
  Else
    Set PlotDestSht = phSamSht
  End If

  Sheetname s
  Sheets.Add
  ActiveSheet.Name = s
  Set tm = ActiveSheet
  NoGridlines
  ShowStatusBar
  m = 3 + piTrimCt

  For TrimNum = 1 To piTrimCt
    k = 0

    For PkNum = 1 To .iNpeaks
      PkTest(PkNum) = (pbCenteredPk(PkNum) And (Not pbUPb Or (PkNum <> pi204PkOrder _
                  And PkNum <> piBkrdPkOrder)))

      If PkTest(PkNum) Then
        k = 1 + k
        ReDim Preserve GrafPk(1 To k)
        GrafPk(k) = .daNominal(PkNum)
        tma(TrimNum, k) = pdaTrimMass(PkNum, TrimNum)
        tta(TrimNum, k) = pdaTrimTime(PkNum, TrimNum) - pdaTrimTime(1, 1)
      End If

    Next PkNum

  Next TrimNum

  If k = 0 Then Erase tma, tta: Exit Sub
End With

Cells.Font.Size = 8: GrafCt = k
ReDim Preserve tma(1 To piTrimCt, 1 To 1 + GrafCt), tta(1 To piTrimCt, 1 To 1 + GrafCt)
Erase pdaTrimMass, pdaTrimTime

For GrafNum = 1 To 2 * GrafCt Step 2
  Cells(3, GrafNum) = "time (hrs)"
  Cells(3, GrafNum + 1) = GrafPk((GrafNum + 1) / 2)
Next GrafNum

With Cells.Font: .Name = "Lucida Console": .Size = 8: End With

For GrafNum = 1 To GrafCt
  k = 2 * GrafNum - 1
  Fonts 1, k, m, , 92, , xlRight, 9, , , , "0.00"
  Fonts 1, k + 1, m, , RGB(0, 0, 92), , xlLeft, 8, , , , "0.0000"
  Fonts 1, 3, , 2 * GrafCt, 92, True, xlRight, 10
  Fonts rw1:=3, Col1:=k, Clr:=0, Bold:=True, Formul:=fsVertToLF("Time|(hrs)"), _
         NumFormat:="@", FontName:="Arial"
  Fonts rw1:=3, Col1:=k + 1, Clr:=0, Bold:=True, Formul:=GrafPk(GrafNum), _
         NumFormat:="@", FontName:="Arial"
Next GrafNum

For TrimNum = 1 To piTrimCt
  For GrafNum = 1 To GrafCt
    k = 2 * GrafNum - 1
    Cells(TrimNum + 3, k) = tta(TrimNum, GrafNum)
    Cells(TrimNum + 3, k + 1) = tma(TrimNum, GrafNum)
Next GrafNum, TrimNum

ColWidth picAuto, 1, 2 * GrafCt

For GrafNum = 1 To GrafCt
  frSr(4, 2 * GrafNum - 1, UBound(tta, 1) + 3, 2 * GrafNum).Sort _
    Cells(4, 2 * GrafNum - 1), Order1:=xlAscending
Next GrafNum

Erase tma, tta
Fonts 1, 1, 3, 2 * GrafCt, , True, , 10
BorderLine xlBottom, 2, 3, 1, , 2 * GrafCt
Columns.ColumnWidth = 11
Fonts 1, 1, , , 0, -1, xlLeft, 12, , , "Trim Mass Record", , "Arial"
w = ActiveSheet.Columns(1).Width
h = ActiveSheet.Rows(5).Height
c = 0: ColCtr = 0: RowCtr = 0

For GrafNum = 1 To GrafCt
  Nchts = ActiveSheet.ChartObjects.Count
  k = 2 * GrafNum - 1
  Set Dr = Range(Cells(4, k), Cells(m, k + 1))
  r = 6 + ColCtr * 200 / h
  c = 1 + RowCtr * 250 / w

  SmallChart DataRange:=Dr, DataSheet:=ActiveSheet, PlaceSheet:=ActiveSheet, _
    Caption:=fsS(GrafPk(GrafNum)), Xname:="Time (hrs)", Yname:="Mass", _
    XaxisScale:=peAutoScale, XmajorTickMark:=xlCross, YaxisScale:=peAutoScale, _
    YmajorTickMark:=xlCross, PlaceRow:=r, PlaceCol:=c, ChtWidth:=250, ChtHeight:=200, _
    ChartBoxClr:=RGB(255, 255, 192), ChartBoxBrdr:=True, PlotBoxClr:=peLightGray, _
    XmajorGridlinesClr:=-1, YmajorGridlinesClr:=vbWhite, XminorGridlinesClr:=-1, _
    YminorGridlinesClr:=-1, Symbol:=xlNone, DataLineClr:=0, FontAutoScale:=False, _
    TikLabelSize:=10, AxisNameSize:=14, BadPlot:=Bad

  If Bad Then GoTo 1
  k = ActiveSheet.ChartObjects.Count

  If k > Nchts Then
    Nchts = k

    If GrafNum Mod 3 = 0 Then
      ColCtr = ColCtr + 1: RowCtr = 0
    Else
      RowCtr = 1 + RowCtr
    End If

    Set ChtOb = foLastOb(ActiveSheet.ChartObjects)
    ChtOb.Activate
    Set Cht = ActiveChart

    With Cht
      With .ChartTitle
        With .Font
          .Size = 12: .italic = 0
          .Color = RGB(0, 0, 208)
          .Name = "comic sans ms"
        End With
        .Left = 0: .Top = Cht.ChartArea.Height
      End With
      With .Axes(2)
        MinY = .MinimumScale: MaxY = .MaximumScale
        .MinimumScale = MinY - 15 * .MajorUnit
        .CrossesAt = .MinimumScale
        .MajorUnit = .MajorUnit * 2
      End With
      MaxX = Dr(Dr.Rows.Count, 1)
      With .Axes(1)
        .MinimumScale = 0
        m2 = .MaximumScale

        Do
          m1 = m2
          m2 = m1 - .MinorUnit
        Loop Until m2 <= MaxX

        .MaximumScale = m1
      End With
    End With

  End If
  ActiveSheet.Cells(1, 1).Select

1: Next GrafNum
End Sub

Sub RatioGraphics() ' Construct one or more chart-insets of scan-by-scan isotope ratios
Dim Smpl As Boolean, Bad As Boolean, Prefd As Boolean, Std As Boolean, GrpSht As Boolean
Dim ChartOK As Boolean
Dim Sh As Worksheet, Na$, RatSht$, Rat$, SamName$, Ysca%, YaxTitleSize%, SymbSize%
Dim cc%, AreaNum%, SerColNum%, NumAreas%, NumCells%, Nshp%, Col%, LastCol%
Dim HdrRow&, LastRow&, r&, rw2&, rr&, CellNum&
Dim Ms$, ms1$, ms0$, SpotNa$
Dim Nu#, de#, SelR As Range, Rw&(), Co%(), ct%
Dim ThisCht As Shape, LastCht As Shape
Dim PlotSht As Worksheet, ShtIn As Worksheet, Ob As Object, TD As Variant, yr As Range

GetInfo
NoUpdate
Prefd = foUser.[ratiodat]
StatBar

If Not Prefd Then
  MsgBox "To obtain isotope-ratio plot-insets, the " & fsInQ("within-scan isotope ratio/eqn sheet") _
    & vbLf & "box in the Preferences panel must be checked before data reduction.", , pscSq
  Exit Sub
End If

Na = "Within-Spot Ratios"
ms0$ = "Ratio-graphics column must be a ratio or NU-switched user Equation."
ms1 = fsVertToLF("To create ratio-graphics you must select either the Sample Data|" & _
  "or StandardData1 worksheet, then select the cell whose row is|for the spot of interest " & _
  "and whose column is for the ratio of interest.")
Ms = ms0
If Workbooks.Count = 0 Then CreateNewWorkbook
Set ShtIn = ActiveSheet
If ShtIn.Type <> xlWorksheet Then Ms = ms1: GoTo 1

For Each Sh In ActiveWorkbook.Worksheets
  If Sh.Name = Na Then RatSht = Na: Exit For
Next Sh

Std = (ActiveSheet.Name = pscStdShtNa): Smpl = (ActiveSheet.Name = pscSamShtNa)
HdrRow = flHeaderRow(Std)
GrpSht = (Cells(1, 1) = "Errors are 1s unless otherwise specified")
plHdrRw = flHeaderRow(IIf(Std, True, False), , True)
LastRow = plaLastDatRw(-Std)
Set SelR = Selection
ct = 0

With SelR
  NumAreas = .Areas.Count
  NumCells = .Cells.Count
  ReDim Rw(1 To NumCells), Co(1 To NumCells)

  For AreaNum = 1 To NumAreas

    For CellNum = 1 To .Areas(AreaNum).Cells.Count
      ct = 1 + ct
      Rw(ct) = .Areas(AreaNum)(CellNum).Row
      Co(ct) = .Areas(AreaNum)(CellNum).Column
      r = Rw(AreaNum): Col = Co(AreaNum)

      If plHdrRw = 0 Or r <= HdrRow Or r > LastRow Then
        Ms = ms1
      ElseIf RatSht = "" Then
        Ms = fsVertToLF("To create ratio-graphics you must start from a workbook|" _
          & "containing a SQUID-created " & fsQq("$Within-Spot Ratios$") & " worksheet.")
      ElseIf (Not Std And Not Smpl And Not GrpSht) Or r <= HdrRow Or r > LastRow Then
        Ms = ms1
      Else

        For Col = 1 To 50
          Na = Cells(HdrRow, Col)

          If InStr(Na, "/") Then
            NumDenom Na, Nu, de
            If Nu > 1 And Nu < 300 And de > 1 And de < 300 Then Exit For
          End If

        Next Col

        If Col < Col Or Col > 50 Then
          Ms = "Column does not contain raw isotope-ratios or user-defined Equations."
        End If

      End If

      If Ms <> ms0 Then GoTo 1
  Next CellNum, AreaNum

End With
foAp.Calculate
Set PlotSht = ShtIn
ct = 0
On Error GoTo 1

For CellNum = 1 To NumCells
  ChartOK = False
  r = Rw(CellNum): Col = Co(CellNum)
  Rat = Cells(HdrRow, Col)

  If StrReverse(Left$(StrReverse(Rat), 3)) = "err" Then

    If NumCells = 1 Then
      Col = Col - 1: Rat = Cells(HdrRow, Col)
    Else
      GoTo Nexti
    End If

  End If

  ct = 1 + ct
  StatBar "Ratio plot", ct
  TD = Trim(Cells(r, 2)): Na = Cells(r, 1)
  Subst Rat, vbLf, " "
  Set phRatSht = ActiveWorkbook.Sheets("Within-spot ratios")
  phRatSht.Activate
  rr = flFindRow(1, False, TD, , , , , False, True) - 1
  FindStr Rat, , cc, rr
  If cc = 0 Then GoTo 1
  rw2 = flFindRow(rr, False, , cc, , , , False)
  Set yr = frSr(rr + 1, cc, rw2 - 1)
  frSr(rr + 1, cc - 1, rw2 - 1, cc + 1).Activate

  If fvMin(yr) >= 0 And foUser("zeroymin") Then
    Ysca = Plots.peZeroMinScale
  Else
    Ysca = Plots.peAutoScale
  End If

  SymbSize = IIf(Ysca = peAutoScale, 9, 6)
  YaxTitleSize = IIf(Len(Rat) > 15, 13, 16)
  SpotNa = "Spot " & Cells(rr, 1)

  SmallChart DataRange:=Selection, DataSheet:=ActiveSheet, PlaceSheet:=PlotSht, _
    Caption:=SpotNa, Xname:="", Yname:=Rat, XaxisScale:=peAutoScale, _
    XmajorTickMark:=xlNone, YaxisScale:=Ysca, YmajorTickMark:=xlCross, PlaceRow:=r, _
    PlaceCol:=Col, ChtWidth:=280, ChtHeight:=210, ChartBoxClr:=vbYellow, _
    ChartBoxBrdr:=True, PlotBoxClr:=peLightGray, XmajorGridlinesClr:=-1, _
    XminorGridlinesClr:=-1, YmajorGridlinesClr:=vbWhite, YminorGridlinesClr:=-1, _
    Symbol:=xlCircle, SymbolSize:=SymbSize, SymbLineClr:=vbRed, SymbInteriorClr:=vbWhite, _
    Transp:=0, DataLineClr:=vbBlue, XerrBars:=False, YerrBars:=True, _
    ErrBarsClr:=vbBlue, ErrBarsThick:=xlThick, ErrBarsCap:=xlCap, _
    FontAutoScale:=False, TikLabelSize:=11, AxisNameSize:=16, BadPlot:=Bad

  On Error Resume Next
  With ActiveSheet
    Set ThisCht = foLastOb(.Shapes)
    ThisCht.Name = Cells(rr, 1) & "_" & Rat

    If CellNum > 1 And Col = LastCol Then
      Set LastCht = .Item(Nshp - 1)
      ThisCht.Top = fnBottom(LastCht)
    End If

    foLastOb(.ChartObjects).Select
    Set Ob = ActiveChart
    Ob.Axes(2).AxisTitle.Font.Size = YaxTitleSize

    If yr.Cells.Count = 1 Then
      With Ob.SeriesCollection

        For SerColNum = 1 To .Count
          FormatSeriesCol .Item(SerColNum), , xlNone
        Next SerColNum

      End With
      With Ob.Axes(2)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
      End With
      AddTextbox "Only one " & Rat & " ratio", 13, True, _
                  -1, -1, , xlCenter, xlCenter
    End If

  End With
  LastCol = Col
  On Error GoTo 0
  PlotSht.Cells(r, Col).Activate
  ChartOK = True
Nexti:
Next CellNum

StatBar
Exit Sub

1: If ChartOK Then
  ActiveChart.Deselect
End If
ShtIn.Activate
MsgBox Ms
End Sub

Sub Tick(ByVal TickRange#, TikInterval#) ' Determine best tick interval for an axis
Dim b#

If TickRange <= 0 Then
  MsgBox "SQUID error in Sub Tick -- TickRange passed as zero.", , pscSq
  CrashEnd
End If

TikInterval = 10 ^ fiIntLogAbs(TickRange) / 8

Do While Abs(TickRange / TikInterval) > 8
  TikInterval = 2 * TikInterval
Loop

b = Abs(TikInterval) / 10 ^ fiIntLogAbs(TikInterval)
If b <> Int(b) Then TikInterval = Int(b) * 10 ^ fiIntLogAbs(TikInterval)
TikInterval = Drnd(TikInterval, 8)
End Sub

Sub AxisScale(DataRange As Range, ByVal Xaxis As Boolean, Optional ErrorRange, _
  Optional ByVal MinVal, Optional ByVal MaxVal)
' Determine the best-appearing Min-Max values for a chart Axis.
' assumes N-row by 2-adjacent col single-area range

Dim WithErrs As Boolean, i%, Ax%, j%, N%, tv#, v#

Ax = IIf(Xaxis, xlCategory, xlValue)
j = 2 + Xaxis
WithErrs = Not IsMissing(ErrorRange)

N = DataRange.Rows.Count

With ActiveChart.Axes(Ax)
  If IsMissing(MinVal) Then
    tv = .MaximumScale

    For i = 1 To N

      If DataRange(i) <> "" Then
        v = DataRange(i)
        If WithErrs Then v = v - ErrorRange(i)
        If v < tv Then tv = v
      End If

    Next i
  Else
    tv = MinVal
  End If

  v = Int(tv / (.MajorUnit / 2))
  .MinimumScale = v * .MajorUnit / 2

  If IsMissing(MaxVal) Then
    tv = .MinimumScale

    For i = 1 To N

      If DataRange(i) <> "" Then
        v = DataRange(i)
        If WithErrs Then v = v + ErrorRange(i)
        If v > tv Then tv = v
      End If

    Next i
  Else
    tv = MaxVal
  End If

  v = foAp.RoundUp(tv / .MajorUnit, 0)
  .MaximumScale = v * .MajorUnit

  If Not Xaxis Then .CrossesAt = .MinimumScale
End With
End Sub

Sub FormatErrorBars(SeriesColl As Series, ByVal XoneYtwo%, ByVal ErrAmt As Variant, _
    ByVal Clr&, ByVal LineThick%, Optional ByVal EndCap As Boolean = False)
With SeriesColl
  .ErrorBar Direction:=Choose(XoneYtwo, xlX, xlY), Include:=xlBoth, Type:=xlCustom, _
     Amount:=ErrAmt, MinusValues:=ErrAmt
  With .ErrorBars
    .Border.Color = Clr
    .Border.Weight = LineThick
    .EndStyle = IIf(EndCap, xlCap, xlNoCap)
  End With
End With
End Sub

Sub TestDispersion(TestSht As Worksheet, TestRange As Range, _
                   StdRange As Range, Similar As Boolean)
Dim b1 As Boolean, b2 As Boolean
Dim i%, nt%, ns%
Dim MedTest#, MedStd#, MadTest#, MadStd#, StdSpred#, TestSpred#, test1#, test2#
Dim Tr#(), Sr#()
' Do the times of analysis of a suite of grouped spots overlap substantally
'  with the times of the Standard-spot analyses?

ReDim Tr(TestRange.Count), Sr(StdRange.Count)
TestSht.Activate

With TestRange
  nt = .Count
  ReDim Tr(nt)

  For i = 1 To nt
    Tr(i) = .Cells(i, 1)
  Next

End With

phStdSht.Activate
With StdRange
  ns = .Count
  ReDim Sr(ns)

  For i = 1 To ns
    Sr(i) = .Cells(i, 1)
  Next

End With

MedTest = foAp.Median(Tr)
MedStd = foAp.Median(Sr)
GetMAD Tr, nt, MedTest, MadTest, 0
GetMAD Sr, ns, MedStd, MadStd, 0
TestSpred = 6 * MadTest
StdSpred = 6 * MadStd
test1 = TestSpred / StdSpred
b1 = (test1 < 1.3 And test1 > 0.75)
test2 = MedTest / MedStd
b2 = (test2 < 1.3 And test2 > 0.75)
Similar = b1 And b2
End Sub

Sub ExtractSerCollRange(SC As Series, SCaddress$)
' Extract the range address of a SeriesCollection
Dim Raddr$, addr$, ScF$, p%, q%
ScF = SC.Formula
p = InStr(ScF, ",")
Raddr = StrReverse(Mid(ScF, p + 1))
p = InStr(Raddr, ",")
Raddr = Mid(Raddr, p + 1)
addr = StrReverse(Raddr)
SCaddress = addr
End Sub

Sub AddGroupsToDriftCorr()
' Check all grouped-spot sheets for suitable time-%offset from media
'  data, & put with the Standard-sheet secular drift plot if OK.
Dim OK As Boolean, HasLegend As Boolean, CanAdd As Boolean
Dim SerNa$, GrpNames$(), Co%, GrpCt%, ClrInd%, i%, Le%
Dim DriftPlot As Chart, DriftObj As ChartObject, Sht As Worksheet, Clrs As Variant

On Error GoTo Endd
Set phStdSht = ActiveWorkbook.Sheets(pscStdShtNa)
phStdSht.Activate
ActiveSheet.[WtdMeanA1].Activate
FindStr "Corrected for secular drift", , Co, 1, 1, flHeaderRow(1)
If Co = 0 Then Exit Sub
Set DriftObj = phStdSht.ChartObjects("SquidChart1")
Set DriftPlot = DriftObj.Chart

Clrs = Array(vbCyan, vbYellow, vbGreen, vbBlue, RGB(127, 255, 128), _
  vbWhite, vbMagenta, peMedGray, 0, pePaleGreen, peStraw, PeDarkRed, _
  peDarkGray)
NoUpdate

For Each Sht In ActiveWorkbook.Worksheets
  Sht.Activate
  FindStr "SQUID grouped-sample", , Co, 1, , 1
  CanAdd = True

  If Co > 0 Then
    ClrInd = 1 + GrpCt Mod 13
    With DriftPlot
      If .HasLegend Then
        With .SeriesCollection
          For i = 1 To .Count
            SerNa = .Item(i).Name
            Le = Len(SerNa)
            If Left(Sht.Name, Le) = SerNa Then CanAdd = False: Exit For
          Next i
         End With
      End If

      If CanAdd Then
        AddGroupDriftCorr Clrs(ClrInd), OK

        If OK Then
          foLastOb(DriftPlot.SeriesCollection).Name = Sht.Name
          GrpCt = 1 + GrpCt
        End If

      End If

    End With
  End If

Next Sht

phStdSht.Activate

If GrpCt > 0 Then
  With DriftPlot
    .SeriesCollection(1).Name = "Std"
    .SeriesCollection(2).Name = "Std" & vbLf & "Drift"

    For i = 1 To .SeriesCollection.Count
      With .SeriesCollection(i)
        If .Points.Count < 3 Then .Delete: Exit For
      End With
    Next i

    DriftObj.Width = 368
    .HasLegend = True
    With .Legend
      .Font.Size = 9
      .Position = xlRight
    End With
    .PlotArea.Width = 277
  End With
End If

Endd:
End Sub

Sub AddGroupDriftCorr(ByVal MarkerClr&, OK As Boolean)
Dim SCaddr$, AgeHdr$, tmp$, RadHdr$
Dim i%, RadCo%, DriftCo%, Ntot%, Nused%, Co%, HoursCo%, AgeCo%
Dim FirstRw&, LastRw&, Hr&, Rw&, MinY#, MaxY#
Dim DriftHrs As Range, DriftVals As Range, Drift As Range
Dim SCrange As Range, StdHrs As Range, DriftAges As Range
Dim SC As SeriesCollection, GrpSht As Worksheet
Dim DriftChtObj As ChartObject, DriftCht As Chart
' To the secular-drift Standard chart, add points from a Grouped-spots sheet,
'  after checking for suitable overlap of their Hours with that of the std.

Set GrpSht = ActiveSheet
Set phStdSht = ActiveWorkbook.Sheets(pscStdShtNa)
phStdSht.Activate
Hr = flHeaderRow(1)
FindStr "Hours", , Co, Hr
OK = False
If Co = 0 Then Exit Sub
Set StdHrs = frSr(plaFirstDatRw(1), Co, plaLastDatRw(1))

GrpSht.Activate
Hr = flHeaderRow(0)
LastRw = flEndRow(1)
FirstRw = 1 + Hr
FindStr "Mean age of coherent group", , AgeCo, LastRw + 1, , LastRw + 9
If AgeCo = 0 Then Exit Sub
AgeCo = AgeCo + 2
AgeHdr = Cells(Hr, AgeCo).Formula
Subst AgeHdr, vbLf

Select Case AgeHdr
  Case "204corr206Pb/238UAge":  RadHdr = "4corr206*/238"
  Case "207corr206Pb/238UAge":  RadHdr = "7corr206*/238"
  Case "208corr206Pb/238UAge":  RadHdr = "8corr206*/238"
  Case "204corr208Pb/232ThAge": RadHdr = "4corr208*/232"
  Case "207corr208Pb/232ThAge": RadHdr = "7corr208*/232"
  Case "204corr207Pb/206PbAge": RadHdr = "4corr207*/206*"
  Case "208corr207Pb/206PbAge": RadHdr = "8corr207*/206*"
End Select

FindStr RadHdr, , RadCo, Hr, 1 + AgeCo, Hr
If RadCo = 0 Then Exit Sub
FindStr "Hours", , HoursCo, Hr
If HoursCo = 0 Then Exit Sub
DriftCo = 1 + fiEndCol(Hr)
Set DriftHrs = frSr(FirstRw, HoursCo, LastRw)
Set DriftAges = frSr(FirstRw, AgeCo, LastRw)
Ntot = LastRw - FirstRw + 1
CleanDataRange DriftAges, Nused
If (Nused / Ntot) < 0.85 Then OK = False: Exit Sub
TestDispersion GrpSht, DriftHrs, StdHrs, OK
If Not OK Then Exit Sub

GrpSht.Activate
Set DriftVals = frSr(FirstRw, DriftCo, LastRw)
Set Drift = Union(DriftHrs, DriftVals)

Set DriftChtObj = phStdSht.ChartObjects("SquidChart1")
DriftChtObj.Activate
Set DriftCht = DriftChtObj.Chart
Set SC = DriftCht.SeriesCollection
ExtractSerCollRange SC(1), SCaddr
Set SCrange = Range(SCaddr)

With SC(1)
  On Error Resume Next
  If .HasErrorBars Then .HasErrorBars = False

  For i = 1 To .Points.Count
    With .Points(i)
      If .MarkerStyle = xlX Then .MarkerStyle = xlNone
    End With
  Next i

  If .MarkerForegroundColor > 0 Then .MarkerForegroundColor = 0
  If .MarkerBackgroundColor <> vbRed Then .MarkerBackgroundColor = vbRed
  On Error GoTo 0
End With

GrpSht.Activate

Cells(1, DriftCo).Formula = "=MEDIAN(" & frSr(FirstRw, RadCo, LastRw).Address(0, 0) & ")"
Fonts rw1:=Hr, Col1:=DriftCo, Bold:=False, HorizAlign:=xlRight, _
      italic:=True, Phrase:="%drift"

tmp = "=100*(" & Cells(FirstRw, RadCo).Address(0, 0)
tmp = tmp & "/" & Cells(1, DriftCo).Address & "-1)"
PlaceFormulae tmp, FirstRw, DriftCo, LastRw
frSr(FirstRw, DriftCo, LastRw).NumberFormat = pscZd1
Drift.Copy
phStdSht.Activate

With DriftCht
  .SeriesCollection.Paste Rowcol:=xlColumns, SeriesLabels:=False, _
      CategoryLabels:=True, Replace:=False, NewSeries:=True
  With foLastOb(.SeriesCollection)
    .MarkerBackgroundColor = MarkerClr
    .MarkerForegroundColor = 0
    .MarkerStyle = xlCircle
    .MarkerSize = 6
  End With
End With

GrpSht.DisplayAutomaticPageBreaks = False
[WtdMeanA1].Select
OK = True
End Sub

Sub CleanDataRange(Vals As Range, Nclean%, Optional Errs, _
 Optional CleanRows, Optional CleanValandErrRange, Optional CleanVrange, _
 Optional CleanErrRange, Optional StrikeThruOK = False, _
 Optional NonzeroOnly As Boolean = False, Optional PositiveOnly As Boolean = False)

 ' Remove struck-through and non-numeric points, and those with abnormally large errors
Dim HasErrs As Boolean, OK As Boolean
Dim i%, Nin%, Nc%, r&, Vcol%, Ecol%, FirstRow&, LastRow&
Dim MedErr#, MadErr#, NmadE#, erV#, OkVals#(), OkErs#()
Dim CelVal#, ErVal#, Resid#, ResidToler#
Dim OkVrange As Range, OKerRange As Range, Vcel As Range, Ecel As Range

With Vals
  Nin = .Rows.Count
  Vcol = .Column
  FirstRow = .Row
  LastRow = FirstRow + Nin - 1
End With
HasErrs = True

If fbNIM(Errs) Then
  Ecol = Errs.Column
ElseIf Vals.Columns.Count = 2 Then
  Set Errs = frSr(FirstRow, Vcol + 1, LastRow)
  Ecol = 1 + Vcol
Else
  HasErrs = False
  Ecol = Vcol
End If

ReDim OkVals(Nin), CleanRows(Nin)
If HasErrs Then ReDim OkErs(Nin)
Nclean = 0: Nc = 0


For r = FirstRow To LastRow
  Set Vcel = Cells(r, Vcol)
  If HasErrs Then Set Ecel = Cells(r, Ecol)
  CheckDataCel Vcel, HasErrs, Ecel, OK, CelVal, ErVal, i, False, , True

  If OK Then
    Nc = 1 + Nc

    If Nc = 1 Then
      Set CleanVrange = Vcel

      If HasErrs Then
        Set CleanErrRange = Ecel
      End If

    Else
      Set CleanVrange = Union(CleanVrange, Vcel)
      If HasErrs Then Set CleanErrRange = Union(CleanErrRange, Ecel)
    End If

    OkVals(Nc) = CelVal
    If HasErrs Then OkErs(Nc) = ErVal
  End If

Next r

ReDim Preserve OkVals(Nc), CleanRows(Nc)

If HasErrs Then
  ReDim Preserve OkErs(Nc)
  Set OkVrange = CleanVrange
  Set OKerRange = CleanErrRange
  MedErr = foAp.Median(OkErs)
  GetMAD OkErs, Nc, MedErr, MadErr, 0
  NmadE = fdNmad(OkErs)
  ResidToler = 4 * NmadE

  For i = 1 To Nc
    r = CleanVrange(i, 1).Row
    Set Vcel = Cells(r, Vcol)
    Set Ecel = Cells(r, Ecol)
    erV = Val(Ecel)
    Resid = erV - MedErr

    If Resid < ResidToler Then
      Nclean = 1 + Nclean
      If fbNIM(CleanRows) Then CleanRows(Nclean) = r

      If Nclean = 1 Then
        Set CleanVrange = Vcel
        Set CleanErrRange = Ecel
      Else
        Set CleanVrange = Union(CleanVrange, Vcel)
        Set CleanErrRange = Union(CleanErrRange, Ecel)
      End If

    End If

  Next i
Else
  Nclean = Nc
End If
End Sub

Sub CheckDataCel(ValCel As Range, HasErCel, Optional ErCel, _
  Optional OK, Optional CellVal, Optional ErVal, Optional CounterToIncrement, _
  Optional StrikeThruOK = False, Optional NonzeroOnly As Boolean = False, _
  Optional PositiveOnly As Boolean = False)
' Does the ValCel cell contain a [nonzero][positive][non-struckthrough] number?

Dim Cv#, Ev#, tB As Boolean
If fbNIM(OK) Then OK = False
HasErCel = fbNIM(ErCel)
If HasErCel Then HasErCel = Not ErCel Is Nothing
tB = IsNumeric(ValCel)
If tB And HasErCel Then tB = IsNumeric(ErCel)

If tB Then
  Cv = Val(ValCel)
  If HasErCel Then Ev = Val(ErCel)

  If NonzeroOnly Then
    tB = (Cv <> 0)
    If HasErCel Then tB = tB And (Ev <> 0)
  End If

  If tB And PositiveOnly Then
    tB = (Cv > 0)
    If HasErCel Then tB = tB And (Ev > 0)
  End If

  If tB And Not StrikeThruOK Then tB = Not ValCel.Font.Strikethrough

  If tB Then
    OK = True
    If fbNIM(CounterToIncrement) Then CounterToIncrement = 1 + CounterToIncrement
    If fbNIM(CellVal) Then CellVal = Cv
    If fbNIM(ErVal) Then ErVal = Ev
  End If

End If

End Sub
