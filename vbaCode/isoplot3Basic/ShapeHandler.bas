Attribute VB_Name = "ShapeHandler"
'  Isoplot module Shapes
Option Private Module
Option Explicit: Option Base 1

Sub AddShape(ShpType$, xy As Range, ByVal FillColor&, ByVal LineColor&, _
  ByRef Curved As Boolean, ByVal SortType%, Optional Sname, Optional Pattern = -2, _
  Optional Transp, Optional LineThick, Optional AgeEllLim, Optional FewerPts As Boolean = False)
  Dim FreeForm As Object
' Add a non-rescaleable Excel shape whose X-Y coords are in range xy.
' SortType 0 =>leave at back; 1 => keep to front; 2 =
' If Pattern=-999, don't format -- use .PickUp - .Apply from outside sub
' Isoplot3.36
Dim ns%, j%, Etype&, SegType&, Shp As Object, AltShapeMethod As Boolean
Dim k#, tmp$, Sh As FreeformBuilder, Left!(), Top!()
Dim L#, T#, l1#, t1#, Sht As Object, ts$, Pts, PtsL
Dim r1%, rn%, c%, Nstep#, xx!, yy!
Dim xyT#(), xyN#(), i%, AC As Object, tLa&, tLb&
Static FailedFilled As Boolean

ViM Pattern, -2
ViM FewerPts, False
On Error GoTo 0
Set AC = Ach
Etype = msoEditingAuto: SymbRow = Max(1, SymbRow)
SegType = IIf(Curved, msoSegmentCurve, msoSegmentLine)

If True And xy.Rows.Count > 48 Then ' otherwise crashes!
  'Or (Curved And FewerPts And xy.Rows.Count > 50) Then
  ' Reduce #nodes for speed
  ns = 48: Nstep = xy.Rows.Count / ns
  ReDim xyN#(2, ns)
  k = 0: j = 0

  Do  ' Select every kth point from xy (k must be single)
    k = k + Nstep - 1 + SymbRow: j = j + 1
  If j > ns Then Exit Do
    xyN(1, j) = xy(k, 1): xyN(2, j) = xy(k, 2)
  Loop

  If xyN(1, 1) <> xyN(1, ns) Or xyN(2, 1) <> xyN(2, ns) Then
    ns = 1 + ns  ' Ending point must be same as starting point
    ReDim Preserve xyN(2, ns)
    xyN(1, ns) = xyN(1, 1): xyN(2, ns) = xyN(2, 1)
  End If

  ReDim xyT(ns, 2) ' Must do to satisfy PRESERVE requirements

  For i = 1 To 2: For j = 1 To ns: xyT(j, i) = xyN(i, j): Next j, i

  r1 = xy.Row: c = xy.Column ' Now replace the original xy range
  rn = r1 + xy.Rows.Count - 1
  ChrtDat.Visible = True: ChrtDat.Select
  xy.Clear
  Set xy = sR(r1, c, r1 + ns - 1, 1 + c, ChrtDat)
  xy.Value = xyT
  ChrtDat.Visible = False:  AC.Select
End If

ns = xy.Rows.Count

If xy(ns - 1 + SymbRow, 1) = xy(ns - 2 + SymbRow, 1) _
 And xy(ns - 1 + SymbRow, 2) = xy(ns - 2 + SymbRow, 2) Then
  xy(ns - 1 + SymbRow, 1) = "": xy(ns - 1 + SymbRow, 2) = ""
  ns = ns - 1
  Set xy = Range(ChrtDat.Cells(xy.Row, xy.Column), _
    ChrtDat.Cells(xy.Row + xy.Rows.Count - 2, xy.Column + 1))
End If

ReDim Left(ns), Top(ns)

For j = 1 To ns ' Convert logical X-Y to physical Left-Top
  xx = xy(j, 1): yy = xy(j, 2)
  LeftTop_XY_Convert Left(j), Top(j), xx, yy, True
Next j

AltShapeMethod = Mac
Start:

If Not (Windows Or Mac) Then GetOpSys ' Just in case

With AC
  On Error GoTo ShapeError

  If Not AltShapeMethod Then  ' To avoid .ConvertToShape crashes in Win/Excel97
    MenuSht.Shapes("miShape").Copy ' A rectangle with 5 nodes
    .Paste
    Set Shp = .Shapes("miShape")

    With Shp
      On Error GoTo ShapeError
      ' Now convert the dummy shape to the real shape
      With .Nodes                 ' Relocate the 1st 5 nodes of the rectangle, then

        For j = 1 To ns           '  add nodes (eg for Ellipse) as required (un-
          L = Left(j): T = Top(j) '  less last node @ same position as 1st node).
          If j = 1 Then l1 = L: t1 = T

          If j <= 5 Then
            .SetPosition j, L, T
          ElseIf j < ns Or l1 <> L Or t1 <> T Then
            'If j > 1 And L <> Left(j - 1) And T <> Top(j - 1) Then
              .Insert .Count, msoSegmentLine, Etype, L, T
            'End If
          End If

        Next j

        If SegType <> msoSegmentLine Then
          j = 0 ' Now convert to a curved object (eg ellipse), if necessary.

          Do    ' Setting to msoSegmentCurve adds nodes, so FOR loop no good.
            j = j + 1
            .SetSegmentType j, SegType
          Loop Until j = .Count

        End If

      End With

    End With

  Else ' Macintosh Excel 8/Office98
    Set FreeForm = .Shapes.BuildFreeform(msoEditingAuto, Left(1), Top(1))

    With FreeForm

      For j = 2 To ns  ' Build complete nodes for the shape
        .AddNodes SegType, Etype, Left(j), Top(j)
      Next j

      Set Shp = .ConvertToShape ' Voila
      On Error GoTo 0
    End With

  End If

End With

With Shp

  If NIM(AgeEllLim) Then
    ' Find leftmost, topmost, & bottommost point of age-tick ellipse
    AgeEllLim(1) = 1E+32: AgeEllLim(2) = 1E+32: AgeEllLim(3) = -1E+32

    For j = 1 To Shp.Nodes.Count
      Pts = .Nodes(j).Points
      AgeEllLim(1) = Min(AgeEllLim(1), Pts(1, 1)) ' Min Left
      AgeEllLim(2) = Min(AgeEllLim(2), Pts(1, 2)) ' Min Top
      AgeEllLim(3) = Max(AgeEllLim(3), Pts(1, 2)) ' Max Top
    Next j

  End If

  If IM(Sname) Then
    tmp$ = ChrtDat.Name & Und & tSt(xy.Column)
    tmp$ = tmp$ & "|" & tSt(xy.Row)
    tmp$ = tmp$ & "~" & tSt(xy.Row + xy.Rows.Count - 1)
    ' Shape name is name of sheet for range + X-column# + 1st row + last row, + sort-type
    '  e.g. "PlotDat4_11|1~33_2@2"
    tmp$ = tmp$ & Und + tSt(SortType)
    If Not Curved Then tmp$ = tmp$ & "X"  ' Add X if a polygon rather than an ellipse
    .Name = tmp$
  Else ' Specified by calling context
    .Name = Sname
  End If

  If Pattern <> -999 Then ' If =-999, .PickUp/.Apply to be used outside this sub.

    With .Fill

      If Pattern <> -2 Then

        If False And Pattern <> xlPatternSolid Then  ' ie if not solid
          On Error Resume Next  ' otherwise can get out-of-memory errors
          .Fill.Patterned Pattern
          On Error GoTo 0
        End If

      End If

      With .ForeColor
        'If FillColor < 65 Then
        '  tLa = Excel98Clr(FillColor)
        '  If .SchemeColor <> tLa Then .SchemeColor = tLa
        'Else
        '  If .RGB <> FillColor Then .RGB = FillColor
        'End If

        If FillColor = xlAutomatic Or FillColor = xlNone Then
          .ColorIndex = FillColor
        ElseIf .RGB <> FillColor Then
          .RGB = FillColor
        End If

      End With

      .Visible = True ' Not necessary with Mac
       On Error GoTo ShapeError2

      If NIM(Transp) Then
        If Transp > 0 Then .Transparency = Transp
      End If

      On Error GoTo 0
    End With

    With .Line

      With .ForeColor
        'If LineColor < 65 Then
        '  tLb = Excel98Clr(LineColor)
        '  If .SchemeColor <> tLb Then .SchemeColor = tLb
        'Else
        '  If .RGB <> LineColor Then .RGB = LineColor
        'End If

        If FillColor = xlAutomatic Or FillColor = xlNone Then
          .ColorIndex = LineColor
        ElseIf .RGB <> LineColor Then
          .RGB = LineColor
        End If

      End With

      If NIM(LineThick) Then

        If LineThick < 0 Then
          .Visible = False
        ElseIf .Weight <> LineThick Then
          .Weight = LineThick
        End If

      End If

    End With

  End If

  On Error Resume Next
  .Shadow.Visible = msoFalse
  On Error GoTo 0
End With

With xy
  rn = .Row + .Rows.Count
  c = .Column
  ts$ = xy(rn, 1).Text
  If ts$ <> ShpType$ Then LineInd xy, ShpType$
End With

HasShapes = True: FailedFilled = False
Exit Sub

ShapeError: On Error GoTo 0 ' Crash in ConvertToShape

If (True Or Mac) And Not FailedFilled Then
  AltShapeMethod = Not AltShapeMethod
  FailedFilled = True
  On Error Resume Next
  Shp.Delete
  On Error GoTo 0
  GoTo Start
End If

ShapeError2: DoShape = False

With StPc("cShapes"):
  .Value = xlOff: .Enabled = False: .Visible = False
End With

StatBar
tmp$ = "Error in creating filled symbols"
MsgBox tmp$, , Iso
NoAlerts
On Error Resume Next

With Ash

  If .Type = xlWorksheet Then
    .Select
    i = .DrawingObjects.Count
    If i > 0 Then .DrawingObjects(i).Delete
  Else
    AC.Delete
  End If

End With

ExitIsoplot
End Sub

Sub MakeFreeform(ByVal NumNodes%, Left!(), Top!(), ByVal Curved As Boolean, _
  ByVal EditingType&, Shp As Shape, Bad As Boolean)
Dim i%, SegType&
SegType = IIf(Curved, msoSegmentCurve, msoSegmentLine)
With ActiveChart.Shapes.BuildFreeform(msoEditingCorner, Left(1), Top(1))
  For i = 2 To NumNodes ' Build complete nodes for the shape
    .AddNodes msoSegmentCurve, msoEditingAuto, Left(i), Top(i)
  Next i
  On Error GoTo ShapeError
  Set Shp = .ConvertToShape ' Voila
  On Error GoTo 0
End With
Bad = False: Exit Sub
ShapeError: On Error GoTo 0
Bad = True
End Sub

Sub xxAddShape(ShpType$, xy As Range, ByVal FillColor&, ByVal LineColor&, _
  ByRef Curved As Boolean, ByVal SortType%, Optional Sname, Optional Pattern = -2, _
  Optional Transp, Optional LineThick, Optional AgeEllLim, Optional FewerPts As Boolean = False)
Attribute xxAddShape.VB_ProcData.VB_Invoke_Func = " \n14"
' Add a non-rescaleable Excel shape whose X-Y coords are in range xy.
' SortType 0 =>leave at back; 1 => keep to front; 2 =
' If Pattern=-999, don't format -- use .PickUp - .Apply from outside sub
Dim ns%, j%, Etype&, SegType&, Shp As Object
Dim k#, tmp$, Sh As FreeformBuilder, Left!(), Top!()
Dim L#, T#, l1#, t1#, Sht As Object, ts$, Pts, PtsL
Dim r1%, rn%, c%, Nstep#, xx!, yy!
Dim xyT#(), xyN#(), i%, AC As Object, tLa&, tLb&
Static FailedFilled As Boolean
ViM Pattern, -2
ViM FewerPts, False
On Error GoTo 0
Set AC = Ach
Etype = msoEditingAuto: SymbRow = Max(1, SymbRow)
SegType = IIf(Curved, msoSegmentCurve, msoSegmentLine)
If Curved And FewerPts And xy.Rows.Count > 50 Then
  ' Reduce #nodes for speed
  ns = 36: Nstep = xy.Rows.Count / ns
  ReDim xyN#(2, ns)
  k = 0: j = 0
  Do  ' Select every kth point from xy (k must be single)
    k = k + Nstep - 1 + SymbRow: j = j + 1
  If j > ns Then Exit Do
    xyN(1, j) = xy(k, 1): xyN(k + 1, j) = xy(k, 2)
  Loop
  If xyN(1, 1) <> xyN(1, ns) Or xyN(2, 1) <> xyN(2, ns) Then
    ns = 1 + ns  ' Ending point must be same as starting point
    ReDim Preserve xyN(2, ns)
    xyN(1, ns) = xyN(1, 1): xyN(2, ns) = xyN(2, 1)
  End If
  ReDim xyT(ns, 2) ' Must do to satisfy PRESERVE requirements
  For i = 1 To 2: For j = 1 To ns: xyT(j, i) = xyN(i, j): Next j, i
  r1 = xy.Row: c = xy.Column ' Now replace the original xy range
  rn = r1 + xy.Rows.Count - 1
  ChrtDat.Visible = True: ChrtDat.Select
  xy.Clear
  Set xy = sR(r1, c, r1 + ns - 1, 1 + c, ChrtDat)
  xy.Value = xyT
  ChrtDat.Visible = False:  AC.Select
End If
ns = xy.Rows.Count
If xy(ns - 1 + SymbRow, 1) = xy(ns - 2 + SymbRow, 1) _
 And xy(ns - 1 + SymbRow, 2) = xy(ns - 2 + SymbRow, 2) Then
  xy(ns - 1 + SymbRow, 1) = "": xy(ns - 1 + SymbRow, 2) = ""
  ns = ns - 1
  Set xy = Range(ChrtDat.Cells(xy.Row, xy.Column), _
    ChrtDat.Cells(xy.Row + xy.Rows.Count - 2, xy.Column + 1))
End If
ReDim Left(ns), Top(ns)
For j = 1 To ns ' Convert logical X-Y to physical Left-Top
  xx = xy(j, 1): yy = xy(j, 2)
  LeftTop_XY_Convert Left(j), Top(j), xx, yy, True
Next j

Start:
If Not (Windows Or Mac) Then GetOpSys ' Just in case
With AC
  If Windows Or FailedFilled Then  ' To avoid .ConvertToShape crashes in Win/Excel97
    With Ash.Shapes.BuildFreeform(msoEditingCorner, 0, 0)
      .AddNodes msoSegmentLine, msoEditingAuto, 10, 0
      .AddNodes msoSegmentLine, msoEditingAuto, 10, 10
      .AddNodes msoSegmentLine, msoEditingAuto, 0, 10
      .AddNodes msoSegmentLine, msoEditingAuto, 0, 0
      .ConvertToShape
    End With
    With Last(Ash.Shapes)         ' Now convert the dummy shape to the real shape
    .Name = "miShape"
      With .Nodes                 ' Relocate the 1st 5 nodes of the rectangle, then
        For j = 1 To ns           '  add nodes (eg for Ellipse) as required (un-
          L = Left(j): T = Top(j) '  less last node @ same position as 1st node).
          If j = 1 Then l1 = L: t1 = T
          If j < 5 Then ' Used to be <=5, cause crash???
            .SetPosition j, L, T
          ElseIf j < ns Or l1 <> L Or t1 <> T Then
            'If j > 1 And L <> Left(j - 1) And T <> Top(j - 1) Then
              .Insert .Count, msoSegmentLine, Etype, L, T
            'End If
          End If
        Next j
        If SegType <> msoSegmentLine Then
          j = 0 ' Now convert to a curved object (eg ellipse), if necessary.
          Do    ' Setting to msoSegmentCurve adds nodes, so FOR loop no good.
            j = j + 1
            .SetSegmentType j, SegType
          Loop Until j = .Count
        End If
      End With
      Set Shp = Ash.Shapes("mishape")
    End With
  Else ' Macintosh Excel 8/Office98
    With .Shapes.BuildFreeform(msoEditingAuto, Left(1), Top(1))
      For j = 2 To ns  ' Build complete nodes for the shape
        .AddNodes SegType, Etype, Left(j), Top(j)
      Next j
      On Error GoTo ShapeError
      Set Shp = .ConvertToShape ' Voila
      On Error GoTo 0
    End With
  End If
End With
With Shp
  If NIM(AgeEllLim) Then
    ' Find leftmost, topmost, & bottommost point of age-tick ellipse
    AgeEllLim(1) = 1E+32: AgeEllLim(2) = 1E+32: AgeEllLim(3) = -1E+32
    For j = 1 To Shp.Nodes.Count
      Pts = .Nodes(j).Points
      AgeEllLim(1) = Min(AgeEllLim(1), Pts(1, 1)) ' Min Left
      AgeEllLim(2) = Min(AgeEllLim(2), Pts(1, 2)) ' Min Top
      AgeEllLim(3) = Max(AgeEllLim(3), Pts(1, 2)) ' Max Top
    Next j
  End If
  If IM(Sname) Then
    tmp$ = ChrtDat.Name & Und & tSt(xy.Column)
    tmp$ = tmp$ & "|" & tSt(xy.Row)
    tmp$ = tmp$ & "~" & tSt(xy.Row + xy.Rows.Count - 1)
    ' Shape name is name of sheet for range + X-column# + 1st row + last row, + sort-type
    '  e.g. "PlotDat4_11|1~33_2@2"
    tmp$ = tmp$ & Und + tSt(SortType)
    If Not Curved Then tmp$ = tmp$ & "X"  ' Add X if a polygon rather than an ellipse
    .Name = tmp$
  Else ' Specified by calling context
    .Name = Sname
  End If
  If Pattern <> -999 Then ' If =-999, .PickUp/.Apply to be used outside this sub.
    With .Fill
      If Pattern <> -2 Then
        If False And Pattern <> xlPatternSolid Then  ' ie if not solid
          On Error Resume Next  ' otherwise can get out-of-memory errors
          .Fill.Patterned Pattern
          On Error GoTo 0
        End If
      End If
      With .ForeColor
        'If FillColor < 65 Then
        '  tLa = Excel98Clr(FillColor)
        '  If .SchemeColor <> tLa Then .SchemeColor = tLa
        'Else
        '  If .RGB <> FillColor Then .RGB = FillColor
        'End If
        If FillColor = xlAutomatic Or FillColor = xlNone Then
          .ColorIndex = FillColor
        ElseIf .RGB <> FillColor Then
          .RGB = FillColor
        End If
      End With
      .Visible = True ' Not necessary with Mac
       On Error GoTo ShapeError2
      If NIM(Transp) Then
        If Transp > 0 Then .Transparency = Transp
      End If
      On Error GoTo 0
    End With
    With .Line
      With .ForeColor
        'If LineColor < 65 Then
        '  tLb = Excel98Clr(LineColor)
        '  If .SchemeColor <> tLb Then .SchemeColor = tLb
        'Else
        '  If .RGB <> LineColor Then .RGB = LineColor
        'End If
        If FillColor = xlAutomatic Or FillColor = xlNone Then
          .ColorIndex = LineColor
        ElseIf .RGB <> LineColor Then
          .RGB = LineColor
        End If
      End With
      If NIM(LineThick) Then
        If LineThick < 0 Then
          .Visible = False
        ElseIf .Weight <> LineThick Then
          .Weight = LineThick
        End If
      End If
    End With
  End If
  On Error Resume Next
  .Shadow.Visible = msoFalse
  On Error GoTo 0
End With
With xy
  rn = .Row + .Rows.Count
  c = .Column
  ts$ = xy(rn, 1).Text
  If ts$ <> ShpType$ Then LineInd xy, ShpType$
End With
HasShapes = True
Exit Sub

ShapeError: On Error GoTo 0 ' Crash in ConvertToShape
If Mac And Not FailedFilled Then FailedFilled = True: GoTo Start
ShapeError2: DoShape = False
With StPc("cShapes"):
  .Value = xlOff ': .Enabled = False: .Visible = False
End With
StatBar
tmp$ = "Error in creating filled symbols"
MsgBox tmp$, , Iso
NoAlerts
On Error Resume Next
With Ash
  If .Type = xlWorksheet Then
    .Select
    i = .DrawingObjects.Count
    If i > 0 Then .DrawingObjects(i).Delete
  Else
    AC.Delete
  End If
End With
ExitIsoplot
End Sub

Private Function ColorRGB(ByVal ColorName$) As Long
ColorRGB = MenuSht.Shapes(ColorName$).Fill.ForeColor.RGB
End Function

Sub RescaleAndOrderShapes(Optional ChartSelected = False)
Attribute RescaleAndOrderShapes.VB_ProcData.VB_Invoke_Func = " \n14"
StatBar "rescaling"
ViM ChartSelected, False
GetOpSys
PutShapesBack True, ChartSelected
End Sub

Sub OrderOnly(Optional ExternalInvoked = True, Optional ChartSelected = False)
Attribute OrderOnly.VB_ProcData.VB_Invoke_Func = " \n14"
ViM ExternalInvoked, True
ViM ChartSelected, False
GetOpSys
PutShapesBack True, ChartSelected, True
End Sub

Sub PutShapesBack(ByVal OrderBySize As Boolean, Optional ChartSelected = False, _
  Optional SortOnly = False, Optional HiddenSheetName)
' Delete and replace each Isoplot shape with a new one, rescaled to the current
'  plotbox dimensions, ordered by diagonal-length (unless a no-sort shape, in which
'  case always put at back in original order).
' Isoplot3.36
Dim s As Object, i%, j%, L, T, X, y, u%, HasHid As Boolean
Dim DiagSq#(), Pts, ShpInd&(), Sh As Object, sn$(), sType$, Dsht As Object
Dim c%, r%, W As Object, Npts%, P%, q%
Dim xy As Object, Nshapes&, tmp$, AC As Object, Na$, xp#(), yp#()
Dim NonSorted%, LeaveBack As Boolean, KeepFront As Boolean, nNd%
Dim Lr%, SortType%, SortPass%, Sstep%, AxGrp%
Dim FirstShp%, LastShp%, Lft!, Tp!, Curved As Boolean
Dim PlotHidSheet As Object, Ob As Object, f$, aW As Object, Hi$, Cnm$, HshtName$
Dim zL!, zT!, zW!, zH!
ViM ChartSelected, False
ViM SortOnly, False
If IM(HiddenSheetName) Then HshtName$ = "PlotDat" Else HshtName = HiddenSheetName
If ChartSelected Then ' Chart on a worksheet already selected
  Set AC = Ach: GoTo Begin
Else
  On Error GoTo WorksheetSelected ' Invoked from an Isoplot chart-sheet?
  Set AC = Ach
  On Error GoTo SelectWorksheet   ' Is the chart-sheet an XY-scatter chart?
  If AC.Type = xlXYScatter Then GoTo Begin
  GoTo ChartNotSelected
End If

SelectWorksheet:
Set AC = Ash ' Invoked from a Worksheet
WorksheetSelected: On Error GoTo ChartNotSelected
With AC
  If .Type <> xlWorksheet Then GoTo ChartNotSelected
  If .ChartObjects.Count <> 1 Then GoTo ChartNotSelected
  ' Only 1 chart on the worksheet, so select it.
  .ChartObjects(1).Activate: Set AC = Ach
  If AC.Type <> xlXYScatter Then GoTo ChartNotSelected
End With

Begin: Set Dsht = ActiveSheet
i = 0 ' See if is a probability plot
On Error Resume Next
i = AC.SeriesCollection.Count
On Error GoTo 0
If i = 0 Then GoTo Begin2
Na$ = AC.SeriesCollection(1).Formula
j = InStr(Na$, HshtName$)
If j = 0 Then GoTo Begin2
GoSub CheckHiddenSheets
i = InStr(Na$, "!"): Na$ = Mid$(Na$, j, i - j)
On Error GoTo Begin2
If Left$(Sheets(Na$).Cells(2, 2), 8) <> "ProbPlot" Then GoTo Begin2 Else j = 0
With Sheets(Na$)
  .Visible = True
  For i = 2 To 30
    If .Cells(i, 7) = "MinY" Then j = 1: Exit For
  Next i
  If j = 0 Then GoTo Begin2
  With AC.Axes(xlValue): X = .MinimumScale: y = .MaximumScale: End With
  For j = 1 To i - 1
    .Cells(j, 7) = X: .Cells(j, 9) = y
  Next j
  .Visible = False
End With
App.Calculate
Exit Sub

Begin2: ' Not a probability plot
On Error GoTo 0
NoUp
DoShape = True
Set s = AC.Shapes: Nshapes = s.Count
On Error GoTo 0
If Nshapes = 0 Then StatBar: Exit Sub
ReDim sn$(Nshapes), DiagSq(Nshapes)
GetScale
If OrderBySize Then    ' Find relative size of shapes, using
  For i = 1 To Nshapes '  diagonal-length as an index.
    With s(i)
      Na$ = .Name
      If InStr(Na$, HshtName$) Then
        If i = 1 Then GoSub CheckHiddenSheets
        If (Right$(Na$, 1) = "T") Then
          DiagSq(i) = 1
        Else
          StatBar "sorting" & Str(i)
          nNd = .Nodes.Count
          ReDim xp(nNd), yp(nNd)
          For j = 1 To nNd
            Pts = .Nodes(j).Points
            xp(j) = Pts(1, 1): yp(j) = Pts(1, 2)
          Next j
          With App
            DiagSq(i) = SQ(.Max(xp) - .Min(xp)) + SQ(.Max(yp) - .Min(yp))
          End With
        End If
      End If
    End With
  Next i
  InitIndex ShpInd(), Nshapes
  QuickIndxSort DiagSq, ShpInd()
End If
For i = 1 To Nshapes: sn$(i) = s(i).Name: Next i
With AC.PlotArea: zL = .Left: zT = .Top: zW = .Width: zH = .Height: End With
For SortPass = 0 To -2 * OrderBySize
  If SortPass = 1 Then        ' pass 0: rescale leave-to-back shapes
    FirstShp = Nshapes: LastShp = 1 ' pass 1: rescale&sort size-sortWorksheetSelectedle shapes
  Else                        ' pass 2: rescale keep-to-front shapes
    FirstShp = 1: LastShp = Nshapes
  End If
  Sstep = Sgn(LastShp - FirstShp)
  If Sstep = 0 Then Sstep = 1
  For i = FirstShp To LastShp Step Sstep
    If SortPass = 1 Then j = ShpInd(i) Else j = i
    Na$ = sn$(j)
    If InStr(Na$, HshtName$) > 0 Then
      If i = 1 Then GoSub CheckHiddenSheets
      Set W = s(Na$)
      SortType = 0
      If OrderBySize Then
        For u = Len(Na$) To 1 Step -1
          If Mid$(Na$, u, 1) = Und Then Exit For
        Next u
        SortType = Val(Mid$(Na$, 1 + u))
      End If
      If SortType = SortPass Then
        ParseShapeName Na$, Sh, c, r, Lr, Curved, SortType, AxGrp
        If c = -1 Then GoTo BadSeries
        StatBar "rescaling" & Str(j)
        Set xy = Range(Sh.Cells(r, c), Sh.Cells(Lr, c + 1))
        If SortOnly Then ' Just re-order by size
          W.ZOrder msoBringToFront
        ElseIf (Right$(Na$, 1) = "T") Then   ' Relocate concordia-age tick-labels
          ' From the stored X-Y coords of the textbox Left/Top, calculate
          '  the Left/Top values for the present Plotbox scaling.
          With W
            LeftTop_XY_Convert Lft, Tp, xy(1, 1), xy(1, 2), True
            .Left = Lft: .Top = Tp
            .ZOrder msoBringToFront ' Put in front of all other shapes
          End With
        Else ' Rescale (ie reconstruct) the shape
          W.PickUp: W.Delete
          Set ChrtDat = Sh  ' Needed by AddShape
          Cnm$ = ChrtDat.Name
          sType$ = Sh.Cells(1 + Lr, c).Text
          If sType = "ShapeLine" Then
            AddLine xy, Na$
          Else
            AddShape sType$, xy, 0, 0, Curved, SortType, Na$, -999, -0.75 * FromSquid
          End If
          If Windows And ExcelVersion = 9 Then
            With ChrtDat
              If Not .Visible Then .Visible = True
              .Activate
            End With
          End If
          s(Na$).Apply
          With AC.PlotArea
            If .Left <> zL Then .Left = zL
            If .Top <> zT Then .Top = zT
            If .Width <> zW Then .Width = zW
            If .Height <> zH Then .Height = zH
          End With
          If Windows And ExcelVersion = 9 Then
            With Dsht
              .Select
              If .ChartObjects.Count > 0 Then Last(.ChartObjects).Activate
            End With
            ActiveChart.ChartArea.Select
          End If
        End If
      End If
    End If
  Next i
Next SortPass
If Cnm$ <> "" Then Sheets(Cnm$).Visible = False
StatBar
On Error Resume Next
AC.Deselect
Exit Sub
ChartNotSelected: On Error GoTo 0
StatBar
MsgBox "You must FirstShp select the Isoplot Chart to be rescaled", , Iso
KwikEnd
CheckHiddenSheets:                 ' Make sure each series doesn't refer to
For Each Ob In AC.SeriesCollection '  a PlotDat in another workbook (open or not).
  On Error GoTo NotIsoSht          ' Format would be "=Series(,[Workbook11]PlotData5!....)
  If InStr(Ob.Formula, "]PlotDat") Then
    MsgBox "Can't rescale a chart that has been moved from its source-data sheet", , Iso
    KwikEnd
  End If
NotIsoSht: On Error GoTo 0
Next
Return
BadSeries: MsgBox "Hidden PlotDat sheet for this chart is missing or corrupt -- can't rescale"
KwikEnd
End Sub

Sub GetScale(Optional PlotName, Optional PhysicalOnly As Boolean = False, Optional AxisGrp% = 1)
Attribute GetScale.VB_ProcData.VB_Invoke_Func = " \n14"
Dim A As Object, nPO As Boolean ' Get physical & logical plotbox scale for the active chart
ViM PhysicalOnly, False
ViM AxisGrp, 1
nPO = ((MinX = 0 And MaxX = 0) Or (MinY = 0 And MaxY = 0)) Or Not PhysicalOnly
If IM(PlotName) Then
  Set A = Ach
Else
  Set A = Sheets(PlotName)
End If
With Axxis(1, A)
  PlotBoxLeft = .Left: PlotBoxWidth = .Width
  If nPO Then MinX = .MinimumScale: MaxX = .MaximumScale
End With
With Axxis(2, A, AxisGrp)
  PlotBoxTop = .Top: PlotBoxHeight = .Height
  If nPO Then MinY = .MinimumScale: MaxY = .MaximumScale
End With
PlotBoxRight = PlotBoxLeft + PlotBoxWidth
PlotBoxBottom = PlotBoxTop + PlotBoxHeight
Xspred = MaxX - MinX: Yspred = MaxY - MinY
End Sub

Private Sub ShapesClick() ' Handle "Filled symbols" box-click
DoShape = IsOn(StPc("cShapes"))
End Sub

Sub CornerPoints(x1(), y1(), x2(), y2(), LftCorner(), RtCorner(), Corners() As Boolean)
Attribute CornerPoints.VB_ProcData.VB_Invoke_Func = " \n14"
'x1(),y1() contain the starting and ending x-y of curve 1; x2(),y2() of curve 2.
' RtCorner(),LftCorner() are the corner-points (if needed) for right & left curve ends;
' Corners() indicates whether such corners are needed.
' x1(1) is leftmost x of curve 1; x1(2) is rightmost
Dim i%
Corners(1) = False: Corners(2) = False
For i = 1 To 2
  If (x1(i) = MinX And y2(i) = MinY) Or (x2(i) = MinX And y1(i) = MinY) Then
    LftCorner(1) = MinX: LftCorner(2) = MinY: Corners(1) = True
  ElseIf (x1(i) = MinX And y2(i) = MaxY) Or (x2(i) = MinX And y1(i) = MaxY) Then
    LftCorner(1) = MinX: LftCorner(2) = MaxY: Corners(1) = True
  ElseIf (x1(i) = MaxX And y2(i) = MaxY) Or (x2(i) = MaxX And y1(i) = MaxY) Then
    RtCorner(1) = MaxX:  RtCorner(2) = MaxY:  Corners(2) = True
  ElseIf (x1(i) = MaxX And y2(i) = MinY) Or (x2(i) = MaxX And y1(i) = MinY) Then
    RtCorner(1) = MaxX:  RtCorner(2) = MinY:  Corners(2) = True
  End If
Next i
End Sub

Sub EllCorner(r As Range) ' Adds corner point(s) to clipped ellipses as required
Attribute EllCorner.VB_ProcData.VB_Invoke_Func = " \n14"
' R contains the ellipse x-y pts;
Dim i%, j%, k%, Frow%, Got As Boolean, L%
Dim x1#, x2#, y1#, y2#, CornerX#
Dim CornerY#, Col%, Lrow%, Cornered As Boolean
With r
  Frow = .Row: Col = .Column
  Lrow = Frow + .Rows.Count - 1
End With
Cornered = False: i = 0
Do
  i = i + 1
  k = IIf(i < Lrow, i + 1, 1) ' In case corner pt needs to between 1st & last pt
  x1 = r(i, 1): x2 = r(k, 1)  ' See if any pair of adjacent points brackets a
  y1 = r(i, 2): y2 = r(k, 2)  '  plotbox corner.  If so, insert a corner point.
  Got = False
  If x1 = MinX Or x1 = MaxX Then
    CornerX = x1
    If (y2 = MinY Or y2 = MaxY) And x1 <> x2 Then
      CornerY = y2: Got = True
    End If
  ElseIf y1 = MinY Or y1 = MaxY Then
    CornerY = y1
    If (x2 = MinX Or x2 = MaxX) And y1 <> y2 Then
      CornerX = x2: Got = True
    End If
  End If
  If Got Then ' Insert the corner point
    L = -(i = Lrow)
    For j = Lrow + 5 + L To i + 2 + L Step -1
      For k = 1 To 2
        r(j, k) = r(j - 1 - L, k)
    Next k, j
    i = i + 1
    r(i, 1) = CornerX: r(i, 2) = CornerY
    If L > 0 Then
      i = i + 1
      r(i, 1) = r(1, 1): r(i, 2) = r(1, 2)
    End If
    Lrow = 1 + Lrow + L: Cornered = True
  End If
Loop Until i = Lrow
If Cornered Then Set r = Range(ChrtDat.Cells(Frow, Col), ChrtDat.Cells(Lrow, Col + 1))
End Sub

Sub LeftTop_XY_Convert(Left!, Top!, X!, y!, _
  XYtoLeftTop As Boolean)
Attribute LeftTop_XY_Convert.VB_ProcData.VB_Invoke_Func = " \n14"
' Convert logical (Plotbox) X-Y coordinates to Physical Left/Top, or vice-versa
If XYtoLeftTop Then
  Left = PlotBoxLeft + (X - MinX) / Xspred * PlotBoxWidth
  Top = PlotBoxTop + (MaxY - y) / Yspred * PlotBoxHeight
Else ' LeftTop to XY
  X = (Left - PlotBoxLeft) * Xspred / PlotBoxWidth + MinX
  y = MaxY - (Top - PlotBoxTop) * Yspred / PlotBoxHeight
End If
End Sub

Sub ZoomToPlotbox()
Attribute ZoomToPlotbox.VB_ProcData.VB_Invoke_Func = " \n14"
Dim tbx As Object
If Ash.Type = xlWorksheet Then
  MsgBox "Can't zoom in a worksheet-embedded chart", , Iso
Else
  NoUp
  With ActiveWindow: .Zoom = 100: .Zoom = 400: End With ' To fix VBA bug
  Ach.PlotArea.Select
  ActiveWindow.Zoom = True
  With Ach
    For Each tbx In .TextBoxes
      tbx.AutoSize = True
    Next
    .Deselect
  End With
End If
End Sub

Sub ParseSeriesRange(Formu$, SeriesRange As Range, FirstRow%, Optional LastRow, _
  Optional FirstColumn, Optional LastColumn)
Attribute ParseSeriesRange.VB_ProcData.VB_Invoke_Func = " \n14"
' Extract pertinent details about a chart data-series source-range
Dim P%, q%, Na$
P = InStr(Formu$, ",PlotDat")
If P = 0 Then GoTo Crap
' Parse out the range --assume parallel columns with X-Y values
' Range formula is something like:
' =SERIES(,PlotDat5!$C$2:$C$9,PlotDat5!$D$2:$D$9,11)
q = InStr(Formu$, "!")
On Error GoTo Crap
Na$ = Mid$(Formu$, 1 + P, q - P - 1) ' Hidden plotdat-sheet name
Set ChrtDat = Sheets(Na$)
Formu$ = Mid$(Formu$, 1 + q)      ' so $C$2:$C$9,PlotDat5!$D$2:$D$9,11)
Formu$ = Strip(Formu$, Na$ & "!") ' so $C$2:$C$9,$D$2:$D$9,11)
P = InStr(Formu$, ":"): q = InStr(Formu$, ",")
Formu$ = Left$(Formu$, P - 1) & Mid$(Formu$, q) ' so $C$2,$D$2:$D$9,11)
P = InStr(Formu$, ":"): q = InStr(Formu$, ",")
Formu$ = Left$(Formu$, q - 1) & Mid$(Formu$, P) ' so $C$2:$D$9,11)
P = InStr(Formu$, ",")
Formu$ = Left$(Formu$, P - 1)                   ' so $C$2:$D$9
Set SeriesRange = ChrtDat.Range(Formu$)
With SeriesRange
  FirstRow = .Row
  If NIM(LastRow) Then
    LastRow = FirstRow + .Rows.Count - 1
    FirstColumn = .Column
    LastColumn = FirstColumn + .Columns.Count - 1
  End If
End With
Exit Sub
Crap: FirstRow = 0
End Sub

Sub FreeSpace(Tbox As Object)
Attribute FreeSpace.VB_ProcData.VB_Invoke_Func = " \n14"
' If any shapes are interfering with a textbox, move it to the other top/bottom of
'  the plotbox.  Tbox must be the last shape.
Dim tL#, tr#, TT#, tB#, Interf As Boolean
Dim Nshapes%, Sh As Object, InLR As Boolean, InTB As Boolean, i%
Dim TopMarg#, BottomMarg#, pltT#, pltB#, Na$, Ind%
Dim SL#(), Srt#(), ST#(), SB#()
Nshapes = Ach.Shapes.Count
If Nshapes < 2 Then Exit Sub
ReDim SL(Nshapes), Srt(Nshapes), ST(Nshapes), SB(Nshapes)
With Tbox
  tL = .Left: tr = tL + .Width
  TT = .Top: tB = TT + .Height
End With
With Ach
  Nshapes = .Shapes.Count - 1
  For i = 1 To Nshapes
    Set Sh = .Shapes(i)
    With Sh
      SL(i) = .Left: Srt(i) = SL(i) + .Width
      ST(i) = .Top: SB(i) = ST(i) + .Height
    End With
  Next i
  For i = 1 To Nshapes
    InLR = ((SL(i) > tL And SL(i) < tr) Or (Srt(i) > tL And Srt(i) < tr))
    InTB = ((ST(i) > TT And ST(i) < tB) Or (SB(i) > TT And SB(i) < tB))
    If InLR And InTB Then Exit For
  Next i
  If InLR And InTB Then
    With .PlotArea
      pltT = .Top: pltB = .Top + .Height
      TopMarg = TT - pltT
      BottomMarg = pltB - tB
      With Tbox
        .Top = IIf(TopMarg < BottomMarg, pltB - TopMarg - .Height, pltT + BottomMarg)
        ' Move textbox from top to bottom or bottom to top of plotarea
      End With
    End With
  End If
End With
End Sub

Sub LabelSeriesColor(SeriesRange As Range, ByVal Clr&)
Attribute LabelSeriesColor.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As Object
' Find name of the ellipse or box color, indicate below data-series
For Each s In MenuSht.Shapes
  With s
    If .Fill.ForeColor.RGB = Clr Then
      LineInd SeriesRange, .Name
      Exit For
    End If
  End With
Next
End Sub

Private Sub MatchShape() ' Copy format of specified shape to other filled symbols
Dim s As Object, Trans#, Ptrn&, FillClr&, LineClr&
Dim TextStyle&, Visi As Boolean, Sht As Object, ShpInd As String, sn$
Dim c%, Lr%
On Error GoTo NoCanDo
With Ach
  .Shapes(Selection.Name).PickUp
  On Error GoTo 0
  For Each s In .Shapes
    sn$ = s.Name
    If InStr(sn$, "PlotDat") Then
      ParseShapeName sn$, Sht, c, 0, Lr, 0, 0, 0
      If c = -1 Then MsgBox "Corrupt shape-name for Isoplot chart", , Iso: KwikEnd
      ShpInd = Sht.Cells(1 + Lr, c).Text
      If ShpInd = "ErrEll" Or ShpInd = "ErrBox" Then s.Apply
    End If
  Next
End With
Exit Sub
NoCanDo: MsgBox "Select a shape to copy format from", , Iso
End Sub

Sub ShapeTrans() ' Specify & implement transparency of filled symbols
Attribute ShapeTrans.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Sht As Object, c%, Lr%, ShpInd As String, sn$
Dim Sh As Object, Ev%, s As Object
On Error GoTo NoCanDo
Set Sh = Ach.Shapes
If Sh.Count = 0 Then GoTo NoCanDo
On Error GoTo 0
For Each s In Sh
  sn$ = s.Name
  If InStr(sn$, "PlotDat") Then
    ParseShapeName sn$, Sht, c, 0, Lr, 0, 0, 0
    If c = -1 Then MsgBox "Hidden PlotDatsheet not present in this workbook," & vbLf & _
      " or shape-name in Isoplot chart is corrupt", , Iso: KwikEnd
    ShpInd = Sht.Cells(1 + Lr, c).Text
    If ShpInd = "ErrEll" Or ShpInd = "ErrBox" Then
      Ev = Hun * s.Fill.Transparency
      Exit For
    End If
  End If
Next
LoadUserForm Transp
Transp.TranspBox.Value = Ev
Transp.Show
Exit Sub
NoCanDo: MsgBox "You must select a chart with shapes first", , Iso
End Sub

Sub ApplyTrans(ByVal Ev%) ' Change transparency of all filled symbols
Attribute ApplyTrans.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As Object, Sh As Object, Sht As Object, ShpInd As String
Dim c%, r%, Lr%, sn$, CanDo As Boolean
Set Sh = Ach.Shapes
For Each s In Sh
  With s
    sn$ = .Name
    CanDo = False
    If InStr(sn$, "AllSymbs") Then
      CanDo = True
    ElseIf InStr(sn$, "PlotDat") Then
      ParseShapeName sn$, Sht, c, r, Lr, 0, 0, 0
      If c = -1 Then MsgBox "Corrupt shape-name in Isoplot chart", , Iso: KwikEnd
      ShpInd = Sht.Cells(1 + Lr, c).Text
      If ShpInd = "ErrEll" Or ShpInd = "ErrBox" Then CanDo = True
    End If
    If CanDo Then
      If Not .Visible Then .Visible = True
      .Fill.Transparency = Ev / Hun
    End If
  End With
Next
End Sub

Sub ParseShapeName(ByVal ShpName$, Sht As Object, Col, FirstRow%, _
  LastRow%, Curved As Boolean, SortType%, AxisGrp%)
Attribute ParseShapeName.VB_ProcData.VB_Invoke_Func = " \n14"
' Get details of data=range defining a shape
Dim u%, Straight As Boolean, rc As String * 1
On Error GoTo BadName
u = InStr(ShpName$, Und)
Set Sht = Sheets(Left$(ShpName$, u - 1))
Col = Val(Mid$(ShpName$, u + 1))  ' X-column
FirstRow = Val(Mid$(ShpName$, 1 + InStr(ShpName$, "|")))  ' 1st row
LastRow = Val(Mid$(ShpName$, 1 + InStr(ShpName$, "~"))) ' last "
rc = Right$(ShpName$, 1)
Straight = (rc = "X" Or rc = "%")
Curved = Not Straight
For u = Len(ShpName$) To 1 Step -1
  If Mid$(ShpName$, u, 1) = Und Then Exit For
Next u
SortType = Val(Mid$(ShpName$, 1 + u))
u = InStr(ShpName$, "@")
If u = 0 Then
  AxisGrp = 1
Else
  AxisGrp = Val(Mid$(ShpName$, 1 + u))
End If
Exit Sub
BadName: Col = -1
End Sub

Sub SelectFilled() ' Group all error-boxes & -ellipses & select
Attribute SelectFilled.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Sht As Object, Shp As Object, s As Object, ST$
Dim c%, Lr%, sn$, Sct%, Sg$, cb, cbb As Object
On Error GoTo NoCanDo
Set Shp = Ach.Shapes
On Error GoTo 0
If Shp.Count = 0 Then Exit Sub
For Each s In Shp
  With s
    sn$ = .Name
    If InStr(sn$, "PlotDat") And Right$(sn$, 1) <> "T" Then
      ParseShapeName sn$, Sht, c, 0, Lr, 0, False, 0
      If c = -1 Then MsgBox "Hidden PlotDat sheet missing from Workbook," & vbLf & _
        " or corrupt shape-name in Isoplot chart", , Iso: KwikEnd
      ST$ = Sht.Cells(1 + Lr, c).Text
      If ST$ = "ErrEll" Or ST$ = "ErrBox" Then
        Sct = 1 + Sct
        If Sct = 1 Then
          Sg$ = sn$
        Else
          With Shp.Range(Array(Sg$, sn$)).Group
            .Name = "AllSymbs"
            Sg$ = .Name
            .Select
          End With
        End If
      End If
    End If
  End With
Next
If Len(Sg$) = 0 Then Exit Sub
NoUp False
On Error Resume Next
cb = App.CommandBars(2).Controls(5).Controls(1).Execute
' "Chart Menu Bar" and "Format"
'App.CommandBars("Formatting").Controls("Fill Color").execute
UnSelectFilled
Exit Sub
NoCanDo:
End Sub

Private Sub UnSelectFilled() ' Ungroup grouped error-symbols
Dim Shp As Object
On Error GoTo done
Set Shp = Ach.Shapes
If Shp.Count = 0 Then Exit Sub
Do
  Shp("AllSymbs").Ungroup
Loop
done:
End Sub

Private Function FalseTrans(ByVal f) ' Tranform colors to a more "linear" visual scale
Dim i%, j%, A, b

b = Array(0, 0.2, 0.32, 0.66, 0.84, 1)
A = Array(0, 0.18, 0.32, 0.58, 0.88, 1)

Do
  i = i + 1
Loop Until A(i + 1) >= f

Do
  j = j + 1
Loop Until b(j + 1) >= f

FalseTrans = (f - A(i)) / (A(i + 1) - A(i)) * (b(i + 1) - b(i)) + b(i)
End Function

Private Function FalseColor(ByVal X, ByVal Xmin, ByVal Xmax) As Long
Dim r#, G#, b#, f#, M%

M = 255
f = (X - Xmin) / (Xmax - Xmin)
f = MinMax(0, 1, f)
f = FalseTrans(f)

Select Case f
  Case 0 To 0.25
    r = 0: G = 4 * f: b = 1
  Case 0.25 To 0.5
    r = 0: G = 1: b = 2 - 4 * f
  Case 0.5 To 0.75
    r = 4 * f - 2: G = 1: b = 0
  Case 0.75 To 1
    r = 1: G = 4 * (1 - f): b = 0
End Select

FalseColor = RGB(r * M, G * M, b * M)
End Function

Sub UseFalseColors()
Attribute UseFalseColors.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, Dlb As Object, Col%, FirstRow%, FirstRow0%, LastRow%, tB(3) As Object
Dim Npts%, s$, Msg$, ShpType$, Sht As Object, ShpNa$(), OK As Boolean, AC As Object, Shp As Object
Dim Nshp%, db As Object, r As Range, Got As Boolean, MinCol%, NsV%
Dim Xmin#, Xmax#, v#, ShpCol%(), ShpFirstRow&(), ShpLastRow&(), sv#(), cs$, c&, k%
Dim A As Range, RtMargin#, tmp#, cc As Boolean
Dim tbH#(3), tbT#(3), ChtR#, vx#, vXmin#, vXmax#
Dim Xmn$, Xmx$, MaxTbW#, CharLeft#, ClrBar As Object, Delta#
Static Legend$, Lower#, Upper#, AutSc As Boolean, AsLog As Boolean, Srange$

GetOpSys
k = Ash.Type
If k = xlWorksheet Then MsgBox "Can't color-scale embedded charts", , Iso: KwikEnd
If k <> xlXYScatter Then MsgBox "Select an Isoplot chart first", , Iso: KwikEnd
On Error GoTo NoCanDo
Set AC = Ach
On Error GoTo 0
GetPlotInfo OK
PlotIdentify
Msg = "Can only color-scale filled error-ellipses or error-boxes." ' 10/10/12 -- added
ShpType = Choose(SymbType, "ErrEll", "", "ErrBox")                 '    ditto

If CumGauss Then
  MsgBox "Can't color-code this type of plot", , Iso
  KwikEnd
End If

On Error GoTo 1
Set Shp = AC.Shapes
Nshp = Shp.Count
' 10/10/12 -- added the ShpFirstRow and ShpLastRow variables
ReDim ShpNa$(Nshp), ShpCol(Nshp), ShpFirstRow(Nshp), ShpLastRow(Nshp)

1 On Error GoTo 0

If Nshp < 2 Or (SymbType <> 1 And SymbType <> 3) Or Not DoShape Then ' 10/10/12 -- added
  MsgBox Msg, , Iso
  KwikEnd
End If

Nshp = 0
FirstRow0 = 1

For i = 1 To Shp.Count

  If InStr(Shp(i).Name, "PlotDat") Then
    '  10/10/12 -- pass the FirstRow, LastRow params
    ParseShapeName Shp(i).Name, Sht, Col, FirstRow, LastRow, 0, 0, 0

    If Col < 1 Then ' 10/10/12 -- changed from Col = -1
      MsgBox "Corrupt shape-name in Isoplot chart", , Iso
      KwikEnd
    End If

    If InStr(Sht.Name, "PlotDat") Then
      s$ = Sht.Cells(1 + LastRow, Col).Text

      If s$ = ShpType Then           '  /
        Nshp = 1 + Nshp              ' |
        ShpNa$(Nshp) = Shp(i).Name   ' |  10/10/12 -- modified
        ShpCol(Nshp) = Col           ' |
        ShpFirstRow(Nshp) = FirstRow ' |
        ShpLastRow(Nshp) = LastRow   '  \
      End If

    End If

  End If

Next i

If Nshp < 2 Then MsgBox Msg, , Iso ' 10/10/12 -- added

ReDim Preserve ShpNa$(Nshp), ShpCol(Nshp), ShpFirstRow(Nshp), ShpLastRow(Nshp)
NoUp False
On Error GoTo NoCanDo
DatSht.Activate
On Error GoTo 0
LoadUserForm FalseClr
On Error Resume Next

If Selection.Cells.Count >= Nshp And Selection.Columns.Count < 2 Then
  FalseClr.eRange = Selection.Address
ElseIf Srange$ <> "" Then
  FalseClr.eRange = Srange$
End If

On Error GoTo 0
' Problem: How to make 1st RefEdit control range-input only

Begin:
DatSht.Activate
'LoadUserForm FalseClr      ' 10/10/12 -- commented out

With FalseClr
  .eName = Legend$
  cc = Not .cAutoscale
  With .fLimits: .Enabled = True: .Visible = True: End With
  .cAutoscale = AutSc: .cLog = AsLog
  .eUpper.Text = "": .eLower.Text = ""
  On Error Resume Next
  .eUpper.Text = IIf(AutSc Or Upper <= Lower, "", Upper)
  .eLower.Text = IIf(AutSc Or Upper <= Lower, "", Lower)
  On Error GoTo 0
  .eLower.Enabled = cc: .lLower.Enabled = cc
  .eUpper.Enabled = cc: .lUpper.Enabled = cc
  .Show
  If Canceled Then GoTo done
  AsLog = .cLog: AutSc = .cAutoscale: Legend$ = .eName

  If Not AutSc Then
    If Not IsNumeric(.eUpper) Or Not IsNumeric(.eLower) Then _
      MsgBox "Invalid range limits", , Iso: GoTo Begin
    Upper = .eUpper: Lower = .eLower
  End If

  On Error GoTo BadRange
  Set r = Range(.eRange)
  On Error GoTo 0
  Srange$ = .eRange
  If r.Columns.Count > 1 Then _
    MsgBox "Color-scaling values must occupy only 1 column", , Iso: GoTo Begin
  NsV = 0

  For i = 1 To r.Areas.Count
    Set A = r.Areas(i)

    For j = 1 To A.Rows.Count

      If A(j) <> "" And IsNumeric(A(j)) And Not A(j).Font.Strikethrough Then
        NsV = 1 + NsV
        ReDim Preserve sv(NsV)
        sv(NsV) = A(j)
      End If

  Next j, i

  If NsV < 2 Then
    MsgBox "No values in the scale-range", , Iso
    GoTo Begin
  End If

  If Not AutSc Then

    If Not IsNumeric(Upper) Or Not IsNumeric(Lower) Then
      AutSc = True
    ElseIf Upper = 0 And Lower = 0 Then
      AutSc = True
    End If

  End If

  If AutSc Then
    Xmin = App.Min(sv): Xmax = App.Max(sv)
  Else
    Xmin = Val(.eLower): Xmax = Val(.eUpper)
    Lower = Xmin: Upper = Xmax
  End If

End With

If Xmin >= Xmax Then
  MsgBox "Upper bound of scale must exceed lower bound", , Iso
  GoTo Begin
End If

If Nshp <> NsV Then
  MsgBox "Number of color-scale values must match number of plotted points", , Iso
  GoTo Begin
End If

If AsLog Then

  If Xmin <= 0 Or Xmax <= 0 Then

    MsgBox "Color-scale-values must be positive for Log-scaled colors", , Iso
    GoTo Begin
  Else

    For i = 1 To NsV

      If sv(i) <= 0 Then
        MsgBox "All index=values must be positive for Log-scaled colors", , Iso
        GoTo Begin
      End If

    Next i

  End If

End If

ScaleMinMax Xmin, Xmax, vXmin, vXmax, AsLog

With App

  For i = 1 To Nshp      ' Now false-color the shapes

    'MinCol = .Min(ShpCol)                  ' 10/10/12 -- commented out

    'For j = 1 To Nshp                      '    "
    '  If ShpCol(j) = mincol Then Exit For  '    "
    'Next j                                 '    "

    If ChrtDat.Cells(1 + ShpLastRow(i), ShpCol(i)) = ShpType Then  ' 10/10/12 -- added

      If AsLog Then vx = Log10(sv(i)) Else vx = sv(i)

      If ColorPlot Then
        c = FalseColor(vx, vXmin, vXmax)
      Else
        k = (vx - vXmin) * 255 / (vXmax - vXmin)
        c = RGB(k, k, k)
      End If

      Shp(ShpNa$(i)).Fill.ForeColor.RGB = c      ' 10/10/12 -- change j to i
      'ShpCol(i) = 32767                         ' 10/10/12 -- commented out
    End If

  Next i

  .ScreenUpdating = False
End With

AC.Select
GetScale

If AsLog Then
  Xmn$ = tSt(10 ^ vXmin): Xmx$ = tSt(10 ^ vXmax)
Else
  Xmn$ = tSt(vXmin): Xmx$ = tSt(vXmax)
End If

cs$ = IIf(ColorPlot, "ColorScale", "BWscale")

With AC ' Put in labelled color legend-bar
  On Error Resume Next ' Delete any old legend
  .Shapes(cs$).Delete:     .TextBoxes("IsoLegend").Delete
  .TextBoxes("IsoColorMin").Delete: .TextBoxes("IsoColorMax").Delete
  On Error GoTo 0
  ChtR = .ChartArea.Left + .ChartArea.Width

  For i = 1 To 3
    Set tB(i) = .TextBoxes.Add(0, 0, 1, 1)

    With tB(i)

      If i = 1 Then ' Legend-bar label
        .Text = Legend$: .Name = "IsoLegend"
      ElseIf i = 2 Then ' lower scale-limit
        .Text = Xmn$: .Name = "IsoColorMin"
      Else ' upper scale-limit
        .Text = Xmx$: .Name = "IsoColorMax"
      End If

      .AutoSize = True: .Font.Name = "Arial"
      .Font.Size = 12 + (i > 1)
      tbH(i) = .Height
      MaxTbW = Max(MaxTbW, .Width)
    End With

  Next i

  CharLeft = ChtR - MaxTbW
  MenuSht.Shapes(cs$).Copy
  .Select: .Paste ' get color-legend from Menus
  Set ClrBar = .Shapes(cs$)

  With ClrBar ' Provisionally, put at far right, mid-height of the chart
    .Left = ChtR - MaxTbW - 3 - .Width
    .Top = AC.PlotArea.Top + (AC.PlotArea.Height - .Height) / 2
  End With

  With ClrBar
    tB(1).Top = .Top + (.Height - tbH(1)) / 2
    tB(2).Top = .Top + .Height - tbH(2) / 2
    tB(3).Top = .Top - tbH(3) / 2
  End With

  For i = 1 To 3: tB(i).Left = CharLeft: Next i
  RtMargin = ClrBar.Left - PlotBoxRight
  ' Make sure legend is just to the right of the plotbox

  If RtMargin > 10 Then ' Plotbox way to left -- move legend left
    ClrBar.Left = PlotBoxRight + 6

    For i = 1 To 3
      tB(i).Left = ClrBar.Left + ClrBar.Width + 2
    Next i

  ElseIf RtMargin < 6 Then ' No room at right -- shrink plotbox towards left
    Delta = .PlotArea.Left + .PlotArea.Width - PlotBoxRight  ' approximate, but can
    .PlotArea.Width = .PlotArea.Width - Delta + RtMargin + 2 '  do no better.
    RescaleOnlyShapes False
  End If

End With

done: AC.Activate
If Ash.Type = xlXYScatter Then Ach.Deselect
Exit Sub

NoCanDo:
MsgBox "Not an Isoplot chart, or" & vbLf & "source-data is missing/corrupt, or" & vbLf & _
 "source data-sheet has been renamed.", , Iso
Rcalc
Exit Sub

BadRange: On Error GoTo 0
MsgBox "Invalid range", , Iso
GoTo Begin
End Sub

Sub CreateColorScale() ' False-color legend-bar to put in Menus sheet
Attribute CreateColorScale.VB_ProcData.VB_Invoke_Func = " \n14"
' START WITH BLANK SHEET, ***NOT*** THE MENUS SHEET
Dim L, T, W, H, bL, Bt, bw, bh, i%, f, Nsteps%, BaW As Boolean
Dim c&, j%, N$, cs As Variant
bL = 0: Bt = 0: bw = 20: bh = 100: Nsteps = 60
L = bL: T = Bt + bh: W = bw: H = bh / Nsteps
cs = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, _
  21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, _
  41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61)
BaW = True
With Ash
  If .Name = "Menus" Then KwikEnd
  With .Shapes
    .AddShape(msoShapeRectangle, bL, Bt, bw, bh).Line.Weight = 2
    For i = 1 To Nsteps
      f = (i - 1) / (Nsteps - 1)
      T = T - H
      With .AddShape(msoShapeRectangle, L, T, W, H)
        If BaW Then
          j = (i - 1) * 255 / Nsteps
          c = RGB(j, j, j)
          N$ = "BWscale"
        Else
          c = FalseColor(f, 0, 1)
          N$ = "ColorScale"
        End If
        .Fill.ForeColor.RGB = c
        .Line.Visible = False
      End With
    Next i
    .Range(cs).Group
  End With
  Last(.Shapes).Name = N$
End With
End Sub

Sub ScaleMinMax(ByVal LowerLim#, ByVal UpperLim#, _
  MinScale#, MaxScale#, ByVal AsLog As Boolean)
' Returns simple scaling Min/Max without too many sigfigs.
Dim TickInt#, v#, LogLow#, LogHi#
If AsLog Then
  MinScale = Int(Log10(LowerLim))
  MaxScale = -Int(Log10(UpperLim))
  If MinScale = MaxScale Then MaxScale = 1 + MaxScale
Else
  Tick UpperLim - LowerLim, TickInt
  While Drnd(v, 8) < Abs(Drnd(LowerLim, 8))
    v = v + TickInt
  Wend
  If LowerLim > 0 Then v = v - TickInt
  MinScale = Drnd(v * Sgn(LowerLim), 7)
  v = MinScale
  While Drnd(v, 8) < Drnd(UpperLim, 8)
    v = v + TickInt
  Wend
  MaxScale = Drnd(v, 7)
End If
End Sub

Sub LogicalSlopeToPhysicalAngle(ByVal Slope#, ByRef Angle#)
Attribute LogicalSlopeToPhysicalAngle.VB_ProcData.VB_Invoke_Func = " \n14"
Dim v# ' Converts a logical slope to a physical (on chart) angle, in degrees
v = (Yspred / Xspred) / (PlotBoxHeight / PlotBoxWidth)
Angle = Atn(Slope / v) * 180# / pi
End Sub

Sub RotLabel(ByVal Label$, ByVal Font$, ByVal FontSize#, ByVal Italic As Boolean, X!, y!, _
  ByVal Slope#, Optional Topsy = False, Optional AsAngle As Boolean = False)
Attribute RotLabel.VB_ProcData.VB_Invoke_Func = " \n14"
' Places a text-art label on the active chart, with specified font, size,
'  logical X-Y, & logical slope, grouped with a rounded rectangle behind,
'  whose color matches plotbox interior or exterior clr.
' If Topsy, then rotate 180 degrees.
Dim bLeft#, bTop#, Bwidth#, bHeight#, BkClr&
Dim lName$, bName$, lLeft!, lTop!, Angle#, v#, InBox As Boolean
ViM Topsy, False
ViM AsAngle, False
'If PlotBoxHeight = 0 Then GetScale
GetScale
With Ach ' Select bkrd color to match plot bkgrd.
  InBox = (X < MaxX And X > MinX And y < MaxY And y > MinY)
  If InBox Then
    BkClr = .PlotArea.Interior.Color
  Else
    If ConcPlot Then Exit Sub
    BkClr = .ChartArea.Interior.Color
  End If
End With
LeftTop_XY_Convert lLeft, lTop, (X), (y), True
If AsAngle Then Angle = Slope Else LogicalSlopeToPhysicalAngle Slope, Angle
If Topsy Then
  Angle = Angle - 180
ElseIf ConcPlot Then
  Angle = Angle - 90
End If
Ach.Shapes.AddTextEffect(msoTextEffect1, Label$, Font$, FontSize, _
     False, Italic, lLeft, lTop).Select
With Selection.ShapeRange
  lName$ = .Name: .Line.Visible = False
  .Fill.ForeColor.RGB = vbBlack ' font color
  .Top = .Top - .Height / 2
  .Left = .Left - .Width / 2      ' Center on X-Y
  If Not ConcPlot Then
    Bwidth = .Width * 1.12:    bLeft = .Left - .Width * 0.06 ' Slightly larger
    bHeight = .Height * 1.2:   bTop = .Top - .Height * 0.1   '  rounded rectangle.
  End If
End With
If Not ConcPlot Then
  ' Add the rounded rectangle background
  Ach.Shapes.AddShape(msoShapeRoundedRectangle, _
    bLeft, bTop, Bwidth, bHeight).Select
  With Selection.ShapeRange ' Blend with local bkrd color
    bName$ = .Name: .Line.Visible = False
    .Fill.ForeColor.RGB = BkClr
    .Shadow.Visible = False
    .ZOrder msoSendToBack
  End With
  Ach.Shapes.Range(Array(lName$, bName$)).Group.Select
End If
With Selection.ShapeRange
  .Rotation = -Angle ' Rotates 45 degrees anti-clockwise
  .ScaleWidth 1.4, False, msoScaleFromMiddle  ' Stretch slightly
  If ConcPlot Then
    InBox = (.Left > PlotBoxLeft And .Top > PlotBoxTop)
    InBox = InBox And (.Left + .Width) < PlotBoxRight And (.Top + .Height) < PlotBoxBottom
    If Not InBox Then .Delete
  End If
End With
End Sub

Sub AddLine(xy As Range, Optional Sname, Optional LineThick)
Attribute AddLine.VB_ProcData.VB_Invoke_Func = " \n14"
Dim l1!, t1!, l2!, t2!
Dim x1!, y1!, x2!, y2!
GetScale
x1 = xy(1, 1): x2 = xy(2, 1)
y1 = xy(1, 2): y2 = xy(2, 2)
LeftTop_XY_Convert l1, t1, x1, y1, True
LeftTop_XY_Convert l2, t2, x2, y2, True
Ach.Shapes.AddLine(l1, t1, l2, t2).Select
If NIM(Sname) Then Selection.Name = Sname
End Sub
