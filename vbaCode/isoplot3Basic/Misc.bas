Attribute VB_Name = "Misc"
' Isoplot module Misc
Option Private Module
Option Base 1: Option Explicit
Const IsoCap = "Iso&plot"
Dim Cap$(2, 11), Macro$(2, 11), Ni%(2), cb$(2), bg(2, 11) As Boolean

Sub GetMenus() ' Get Isoplot-menu defs from file IsoMenu
Attribute GetMenus.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, k%, iMr As Range
cb$(1) = "Worksheet Menu Bar": cb$(2) = "Chart Menu Bar"
Set iMr = Menus("IsoMenu")
For i = 1 To 2: iMr(i) = Ni(i): Next i
Set iMr = iMr(3): k = 0
For i = 1 To 2: Ni(i) = iMr(i): Next i
Set iMr = iMr(3)
For i = 1 To 2
  For j = 1 To Ni(i)
    k = 1 + k: Cap$(i, j) = iMr(k)
    k = 1 + k: Macro$(i, j) = iMr(k)
    k = 1 + k: bg(i, j) = iMr(k)
Next j, i
End Sub

Sub InsertWtdResids(rBox As Object, Optional Hdr) ' Insert a column containing the wtd Y-resids of the regression.
Attribute InsertWtdResids.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i&, b As Boolean, N&, r As Range, NewCol As Boolean, r1&, r2&
ViM Hdr, "Wtd" & vbLf & "Resids"
'If Windows And ExcelVersion >= 10 Then ' Doesn't always seeem to work
'  If Not Ash.AllowInsertingColumns Then
'    MsgBox "Can't insert wtd resids column in this protected sheet." & viv$ & _
'      "(to unprotect, select Tools/Protection from the Excel menu-bar)",,Iso
'    ExitIsoplot
'  End If
'End If
If IsOn(rBox.CheckBoxes("cAddWtdResids")) Then
  InvokeLinearizedProb True
  N = UBound(yf.WtdResid)
  r1 = RangeIn(1).Row: r2 = r1 + N - 1
  If Xcolumn = 1 Then
    Cells(1, Xcolumn).EntireColumn.Insert
    Xcolumn = 1 + Xcolumn
  End If
  Set r = sR(r1, Xcolumn - 1, r2)
  If App.CountBlank(r) <> N Then
    Cells(1, Xcolumn).EntireColumn.Insert
    NewCol = True
    Set r = sR(r1, Xcolumn, r2)
    Xcolumn = 1 + Xcolumn
  End If
  Cells(r1, Xcolumn).Select
  Xcolumn = Xcolumn - 1
  For i = 1 To UBound(ValidRow)
    Cells(ValidRow(i), Xcolumn) = yf.WtdResid(i) 'Sqr(yf.WtdResid(i))
  Next i
  With RangeIn(1): b = IsNumeric(Cells(.Row, .Column)): End With
  If HeaderRow And Not b Then Cells(RangeIn(1).Row, Xcolumn) = Hdr
  With Columns(Xcolumn)
    .HorizontalAlignment = xlRight
    With .Font
      .Bold = True: .Italic = False: .Color = vbRed
      .Underline = xlUnderlineStyleNone
    End With
    .NumberFormat = "0.00"
  End With
  rBox.CheckBoxes("cAddWtdResids") = xlOff
End If
End Sub

Sub AutoCalc()
Attribute AutoCalc.VB_ProcData.VB_Invoke_Func = " \n14"
Dim b As Controls, c As Boolean, i%, N$(2)
N$(1) = "Formatting": N$(2) = "Standard"
On Error Resume Next
With App
  For i = 1 To 2
    For Each b In .CommandBars(N$(i)).Controls
      If b.Caption = "Toggle AutoCalc" Then
        c = (.Calculation = xlCalculationAutomatic)
        b.State = Not c
        .Calculation = IIf(c, xlCalculationManual, xlCalculationAutomatic)
        Exit Sub
      End If
    Next
  Next i
End With
End Sub

Sub ViM(Optional Var, Optional DefaultVal)
Attribute ViM.VB_ProcData.VB_Invoke_Func = " \n14"
If IM(Var) Then Var = DefaultVal
End Sub

Sub HA(Obj As Object, ByVal Align&)
Attribute HA.VB_ProcData.VB_Invoke_Func = " \n14"
Obj.HorizontalAlignment = Align
End Sub

Sub MakeSpline(r As Range, Clr&()) ' Fit spline curve to data and plot
Attribute MakeSpline.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i&, FirstSlope#, LastSlope#, SecondDerivs#()
Dim ST#, sSt#, X#(), y#(), N&, j&, DupX As Boolean, xx#()
Dim IndX&(), nR As Range, xy#(), c&, Tx#(), M&
N = r.Rows.Count: SymbRow = Max(1, SymbRow)
If Aspline Then
  Dim cf#(), xa#(), ya#()
  ReDim cf(N, 3), X(N), y(N)
  For i = 1 To N
    X(i) = r(i, 1): y(i) = r(i, 2)
  Next i
  AkimaSpline X(), y(), N, cf()
  SplinePoints X(), y(), N, cf(), xa(), ya()
  M = UBound(xa)
  ReDim xy(M, 2)
  For i = 1 To M
    xy(i, 1) = xa(i): xy(i, 2) = ya(i)
  Next i
Else ' Nspline
  ReDim X(N), y(N), SecondDerivs(N)
  For i = 1 To N: X(i) = r(i, 1): Next i
  InitIndex IndX(), N
  QuickIndxSort X(), IndX()
  For i = 1 To N: y(i) = r(IndX(i), 2): Next i
  Erase IndX
  FirstSlope = (y(2) - y(1)) / NotSmall(X(2) - X(1))
  LastSlope = (y(N) - y(N - 1)) / NotSmall(X(N) - X(N - 1))
  Spline X(), y(), N, FirstSlope, LastSlope, SecondDerivs()
  ST = (X(N) - X(1)) / Hun
  sSt = ST / 10000#
  M = 101
  ReDim xx#(M + N)
  For i = 1 To 101
    xx(i) = X(1) + ST * (i - 1)
  Next i
  For j = 1 To N
    DupX = False
    For i = 1 To 101
      If Abs(X(j) - xx(i)) < sSt Then DupX = True: Exit For
    Next i
    If Not DupX Then
      M = 1 + M
      xx(M) = X(j)
    End If
  Next j
  ReDim xy(M, 2)
  ReDim Preserve xx(M)
  QuickSort xx()
  For i = 1 To M: xy(i, 1) = xx(i): Next i
  For i = 1 To M
    SPLINT X(), y(), SecondDerivs(), N, xy(i, 1), xy(i, 2)
  Next i
End If
'ChrtDat.Select
Set nR = sR(SymbRow, SymbCol, M - 1 + SymbRow, 1 + SymbCol, ChrtDat)
nR = xy
'Sheets(PlotName$).Select
c = Clr(1, 1)
With Ach.SeriesCollection(Nser)
  .Border.LineStyle = xlNone
  If Aspline Then
    .MarkerStyle = xlNone
  Else ' Nspline
    .MarkerStyle = xlCircle: .MarkerSize = 4
    .MarkerForegroundColor = c: .MarkerBackgroundColor = vbWhite
  End If
End With
AddSymbCol 2 ' Increment column-counter
Ach.SeriesCollection.Add nR, xlColumns, False, True, False
Nser = Ach.SeriesCollection.Count
LineInd nR, "IsoLine"
With Ach.SeriesCollection(Nser)
  .Smooth = True
  With .Border
    .LineStyle = xlContinuous: .Weight = xlThin
    .Color = c
  End With
End With
End Sub

Sub Spline(X#(), y#(), ByVal N&, ByVal Yp1#, ByVal YpN#, y2#())
Attribute Spline.VB_ProcData.VB_Invoke_Func = " \n14"
' Given x() and y() of length N, sorted by X, and given values Yp1 and YpN
'  for the 1st deriv of the interpolating Fn @ pts 1 and N, this routine returns y2() containing the
'  2nd derivs of the interpolating Fn at the x(i).  If Yp1 and/or YpN >=1E30, the routine sets the
'  corresponding boundary condition for a natural spline, with zero 2nd deriv on that boundary.
' From Press et al, 1986, Numerical recipes, p. 88
Dim i&, P#, sig#, k&, u#()
Dim qn#, UN#, d1#, d2#, d3#
ReDim u(N - 1)
If Yp1 > IsoLarge Then  ' Set the lower boundary condition to either "natural"
  y2(1) = 0: u(1) = 0   '  or else to have a specified 1st derv.
Else
  y2(1) = -0.5
  d1 = NotSmall(X(2) - X(1))
  u(1) = 3# / d1 * ((y(2) - y(1)) / d1 - Yp1)
End If
For i = 2 To N - 1                   ' The decomposition loop of the tridiagonal
  d1 = NotSmall(X(i + 1) - X(i - 1)) '  algorithm.  y2 and U are used for temporary
  sig = (X(i) - X(i - 1)) / d1       '  storage of the decomposed factors.
  P = sig * y2(i - 1) + 2#
  y2(i) = (sig - 1#) / P
  d1 = NotSmall(X(i + 1) - X(i)): d2 = NotSmall(X(i) - X(i - 1))
  d3 = NotSmall(X(i + 1) - X(i - 1))
  u(i) = (6# * ((y(i + 1) - y(i)) / d1 - (y(i) - y(i - 1)) _
         / d2) / d3 - sig * u(i - 1)) / P
Next i
If YpN > IsoLarge Then  '  Upper boundary cond'n set either to be "natural"
  qn = 0: UN = 0        '  or else to have a specified 1st derv.
Else
  qn = 0.5
  d1 = NotSmall(X(N) - X(N - 1)): d2 = NotSmall(X(N) - X(N - 1))
  UN = 3# / d1 * (YpN - (y(N) - y(N - 1)) / d2)
End If
y2(N) = (UN - qn * u(N - 1)) / (qn * y2(N - 1) + 1#)
For k = N - 1 To 1 Step -1 ' The backsubstitution loop of the tridiagonal algorithm.
  y2(k) = y2(k) * y2(k + 1) + u(k)
Next k
End Sub

Sub SPLINT(xa#(), ya#(), Y2a#(), ByVal N&, ByVal X#, ByRef y#)
Attribute SPLINT.VB_ProcData.VB_Invoke_Func = " \n14"
' Given Xa() and Ya() with the Xa(i)'s in order, and given Y2a(), which is the output from
'  SPLINE above, and given a value of X, this routine returns a cubic-spline interpolated value Y.
' From Press et al, 1986, Numerical recipes, p. 89
Dim k&, Klo&, Khi&, H#, A#, b#
Klo = 1: Khi = N         ' Find the right place in the table by means of bisection.
Do While (Khi - Klo) > 1 ' Optimal if sequential calls to this routine are at random
  k = (Khi + Klo) / 2    '   values of X.  If sequential calls are in order & closely spaced,
  If xa(k) > X Then Khi = k Else Klo = k '   one would do better to store previous values of Klo
Loop                                     '   & Khi and test if they remain appropriate on the
H = xa(Khi) - xa(Klo)                    '   next call.  Klo & Khi now bracket the input value of x.
If H = 0 Then MsgBox "Error in Spline routine": KwikEnd ' The XA's must be distinct.
A = (xa(Khi) - X) / H ' Cubic spline polynomial is now evaluated.
b = (X - xa(Klo)) / H
y = A * ya(Klo) + b * ya(Khi) + ((A ^ 3 - A) * Y2a(Klo) + (b ^ 3 - b) * Y2a(Khi)) * (H * H) / 6#
End Sub

Sub AkimaSpline(X#(), y#(), ByVal N&, cf#())
Attribute AkimaSpline.VB_ProcData.VB_Invoke_Func = " \n14"
' Determine the coefficients of a spline curve through the N X,Y points
'  using the method of Akima.
' The Cf(i,j) array contains the coefficients of the spline segment
'  between points i and i+1 such that:
'  Yc-Yc(i)=Cf(i,1)*[Xc-Xc(i)] + Cf(i,2)*[Xc-Xc(i)]^2 + Cf(i,3)*[Xc-Xc(i)]^3
Dim xc#(), Yc#(), M#(), Xdel#(), Xt#(), Yt#()
ReDim xc(5), Yc(5), M(4), Xdel(-1 To 2), Xt(N + 4), Yt(N + 4)
Dim i&, j&, k&, Xcc#, Ycc#
Dim Numer#, Denom#, temp#, Slope1#, Slope2#
For i = 3 To N + 2   ' Set up array of local points
  Xt(i) = X(i - 2): Yt(i) = y(i - 2)
Next i
' Add 2 extra estimated points at either end.
For i = 2 To 1 Step -1
  Xt(i) = Xt(i + 1) + Xt(i + 2) - Xt(i + 3)
  Xt(N - i + 5) = Xt(N - i + 4) + Xt(N - i + 3) - Xt(N - i + 2)
Next i
For i = 2 To 1 Step -1
  j = i + 1
  For k = 0 To 2
    Xdel(k) = Xt(j + k) - Xt(j - 1 + k)
  Next k
  If Xdel(0) <> 0 And Xdel(1) <> 0 And Xdel(2) <> 0 Then
    temp = (Yt(j + 2) - Yt(j + 1)) / Xdel(2) - 2 * (Yt(j + 1) - Yt(j))
    Yt(j - 1) = Yt(j) - Xdel(0) * temp / Xdel(1)
  Else
    Yt(j - 1) = Yt(j)
  End If
  j = N - i + 4
  For k = -1 To 1
    Xdel(k) = Xt(j + k) - Xt(j - 1 + k)
  Next k
  If Xdel(-1) <> 0 And Xdel(0) <> 0 And Xdel(1) <> 0 Then
    temp = 2 * (Yt(j) - Yt(j - 1)) / Xdel(0) - (Yt(j - 1) - Yt(j - 2))
    Yt(j + 1) = Yt(j) - Xdel(1) * temp / Xdel(-1)
  Else
    Yt(j + 1) = Yt(j)
  End If
Next i
If Xt(2) <> Xt(1) And Xt(3) <> Xt(2) Then
  temp = (Yt(2) - Yt(1)) / (Xt(2) - Xt(1))
  Slope1 = (temp + (Yt(3) - Yt(2)) / (Xt(3) - Xt(2))) / 2
Else
  Slope1 = 0
End If
For i = 0 To N - 1
  For j = 1 To 5
    xc(j) = Xt(i + j): Yc(j) = Yt(i + j)
  Next j
  For j = 1 To 4
    If xc(j + 1) <> xc(j) Then
      M(j) = (Yc(j + 1) - Yc(j)) / (xc(j + 1) - xc(j))
    Else
      M(j) = Hun * (X(N) - X(1)) / (y(N) - y(1))
    End If
  Next j
  If M(1) = M(2) Then
    Slope2 = (M(2) + M(3)) / 2
  Else
    Numer = Abs(M(4) - M(3)) * M(2) + Abs(M(2) - M(1)) * M(3)
    Denom = Abs(M(4) - M(3)) + Abs(M(2) - M(1))
    Slope2 = Numer / Denom
  End If
  If i Then ' Transform coords so origin is at 1st point of the pair
    Xcc = xc(3) - xc(2)
    If Xcc = 0 Then Xcc = 1E-30
    Ycc = Yc(3) - Yc(2)
    cf(i, 1) = Slope1
    cf(i, 2) = (3 * Ycc / Xcc - 2 * Slope1 - Slope2) / Xcc
    cf(i, 3) = (Slope2 - Slope1 - 2 * cf(i, 2) * Xcc) / (3 * Xcc * Xcc)
  End If
  Slope1 = Slope2
Next i
End Sub

Sub SplinePoints(Xin#(), Yin#(), ByVal N&, _
  cf#(), xOut#(), yOut#())
Attribute SplinePoints.VB_ProcData.VB_Invoke_Func = " \n14"
' Construct x,y array for spline-curves through the N X,Y points where the curve is a cubic
'  polynomial with transformed coords so that the origin is at X(i),Y(i),
'  where i is the first of the pair of pts through which the curve is drawn.
' The Cf(i,j) array contains the coefs of the spline segment
'  between pts i and i+1 such that:
'  Yc = Cf(i,1)*Xc + Cf(i,2)*Xc^2 + Cf(i,3)*Xc^3
'  where Xc = Xc-Xc(i) and Yc = Yc-Yc(i)
Dim i&, k&, Xstep#, MinStep&, ct&, IndX&()
Dim xv#, xc#, Yc#, ST#, MinSt#, MaxSt#, X#(), Xt#(), y#()
ReDim X#(N), Xt#(N), y#(N)
For i = 1 To N: Xt(i) = Xin(i): Next i
InitIndex IndX(), N
QuickIndxSort Xt(), IndX()
For i = 1 To N
  X(i) = Xin(IndX(i)): y(i) = Yin(IndX(i))
Next i
Erase IndX
MinStep = 4   ' No fewer than 4 steps between points
MaxSt = Xspred / Hun: MinSt = Xspred / 5000
For i = 1 To N - 1
  ST = Max(MinSt, Min(MaxSt, (X(i + 1) - X(i)) / MinStep))
  For xv = X(i) To X(i + 1) Step ST
    ct = 1 + ct
    ReDim Preserve xOut(ct), yOut(ct)
    xOut(ct) = xv: xc = xv - X(i)
    yOut(ct) = y(i) + cf(i, 1) * xc + cf(i, 2) * xc * xc + cf(i, 3) * xc * xc * xc
  Next xv
Next i
End Sub

Sub UserCurve(FormulaX$, FormulaY$, ByVal FirstPar, ByVal LastPar, ByVal FirstLabel, ByVal LabelStep)
Attribute UserCurve.VB_ProcData.VB_Invoke_Func = " \n14"
' Plot the user-specified curve
Dim r%, Vchar$, Parametric As Boolean, SerC As Object, y1, y2, xx As Object, yy As Object
Dim Xaddr$, Yaddr$, Paddr$, x1, x2, Sh As Object, X, y, fx$, Fy$, pv, p1, p2, Pspred, Uno, vv&
Dim Xspred, Xstep, Nsteps, Xr As Range, Yr As Range, pr As Range, f$, Ch As Object, Ap As Object
Dim Npts&, Nok&, nR&, P%, q%, M%, v As Variant
Dim Op As Object
Nsteps = 100: Vchar$ = "@"
AssignD "Curve", , , , Op
Parametric = IsOn(Op("oParam"))
NoUp
Set Ch = Ach: Set Ap = App: Set SerC = Ch.SeriesCollection
With Ch
  With .Axes(xlCategory): x1 = .MinimumScale: x2 = .MaximumScale: End With
  With .Axes(xlValue):    y1 = .MinimumScale: y2 = .MaximumScale: End With
End With
f$ = SerC(1).Formula ' Determine source-sheet.
P = InStr(f$, "!"): q = InStr(f$, ","): M = 1
If P = 0 Or q = 0 Then ExitIsoplot
On Error GoTo Ng
Set Sh = Sheets(Mid$(f$, q + 1, P - q - 1))
Sh.Visible = True: Sh.Activate
SymbCol = Cells(4, 2).Value: SymbRow = Max(1, Cells(15, 2).Value)
On Error GoTo done
If Parametric Then
  Cells(SymbRow, SymbCol) = "=" & ApSub(FormulaX$, Vchar$, Str(FirstPar))
  Cells(1 + SymbRow, SymbCol) = "=" & ApSub(FormulaX$, Vchar$, Str(LastPar))
  p1 = Cells(SymbRow, SymbCol).Value: p2 = Cells(1 + SymbRow, SymbCol).Value
  If Not IsNumeric(p1) Or Not IsNumeric(p2) Or p1 = p2 Then
    MsgBox "Invalid parameter-limits", , Iso
    On Error Resume Next
    If Left$(Ash.Name, 7) = "PlotDat" Then Ash.Visible = False
    On Error GoTo 0
    ExitIsoplot
  End If
  If p1 > p2 Then X = p1: p1 = p2: p2 = X: M = -1
End If
Xspred = x2 - x1
If Parametric Then
  Pspred = LastPar - FirstPar: Xstep = Pspred / Nsteps: Uno = FirstPar
Else
  Xstep = Xspred / Nsteps: Uno = x1
End If
With Sh
  For r = 1 To Nsteps + 1
    Set Xr = .Cells(r - 1 + SymbRow, SymbCol - Parametric)
    Set Yr = .Cells(r - 1 + SymbRow, 1 + SymbCol - Parametric)
    If Parametric Then
      Set pr = .Cells(r - 1 + SymbRow, SymbCol)
      Paddr$ = pr.Address(False, False)
      pr = Uno + (r - 1) * Xstep
      Xr = "=" & ApSub(FormulaX$, Vchar$, Paddr$)
      Yr = "=" & ApSub(FormulaY$, Vchar$, Paddr$)
      If IsNumeric(Xr) And IsNumeric(Yr) Then
        Nok = 1 + Nok
      Else
        Xr = "": Yr = "": Xr(1, 0) = ""
      End If
    Else
      Xaddr$ = Xr.Address(False, False)
      Xr = x1 + (r - 1) * Xstep
      Yr = "=" & ApSub(FormulaY$, Vchar$, Xaddr$)
      If IsNumeric(Yr) Then
        Nok = 1 + Nok
      Else
        Xr = "": Yr = ""
      End If
    End If
  Next r
End With
If Nok < 2 Then
  MsgBox "The formula you entered is invalid or cannot be parsed", , Iso
  GoTo done
End If
Set v = sR(SymbRow, SymbCol - Parametric, SymbRow + Nsteps, SymbCol + 1 - Parametric, Sh)
Ch.Activate ' Add the smooth curve
Ch.SeriesCollection.Add v, xlColumns, False, True, False
'Sh.Activate
AddSymbCol 2 - Parametric
'Ch.Activate
With Last(Ch.SeriesCollection)
  .MarkerStyle = xlNone: .Smooth = True
  With .Border
    .LineStyle = xlContinuous: .Weight = xlThin: .Color = vbBlue
  End With
End With
If LabelStep = 0 Or Not Parametric Then GoTo done
r = 0 ' Add the labelled curve-ticks
Sh.Activate
With Sh
  Do
    r = r + 1
    Set pr = .Cells(r - 1 + SymbRow, SymbCol): Paddr$ = pr.Address(False, False)
    Set Xr = .Cells(r - 1 + SymbRow, 1 + SymbCol): Set Yr = .Cells(r - 1 + SymbRow, 2 + SymbCol)
    pr = FirstLabel + (r - 1) * LabelStep
    Xr = "=" & ApSub(FormulaX$, Vchar$, Paddr$)
  If Not IsNumeric(Xr) Then Xr = "": Exit Do
  If pr > LastPar Or r > Thou Then Exit Do
    Yr = "=" & ApSub(FormulaY$, Vchar$, Paddr$)
    If Not IsNumeric(Yr) Then Yr = ""
  Loop
End With
sR(r - 1 + SymbRow, SymbCol, r - 1 + SymbRow, 1 + SymbCol) = ""
nR = r - 1
Set v = sR(1 - 1 + SymbRow, SymbCol + 1, r - 1 + SymbRow, SymbCol + 2)
Ch.SeriesCollection.Add v, xlColumns, False, 1, False
With Last(Ch.SeriesCollection)
  ' Format curve-ticks
  .Border.LineStyle = xlNone: .MarkerStyle = xlCircle: .MarkerSize = 5
  .MarkerBackgroundColor = vbWhite: .MarkerForegroundColor = vbBlue
  .ApplyDataLabels Type:=xlDataLabelsShowLabel ' Adds the Y-values as labels
  With .DataLabels
    '.H'alAlignment = xlCenter: .V'lAlignment = xlCenter: .Orientation = xlH'al
    If IsOn(Op("oLeft")) Then
      vv = xlLabelPositionLeft
    ElseIf IsOn(Op("oRight")) Then
      vv = xlLabelPositionRight
    ElseIf IsOn(Op("oTop")) Then
      vv = xlLabelPositionAbove
    ElseIf IsOn(Op("oBottom")) Then
      vv = xlLabelPositionBelow
    End If
    .Position = vv
  End With
  On Error Resume Next
  For r = 1 To nR ' Change to the parametric values
    Set xx = Sh.Cells(r - 1 + SymbRow, SymbCol + 1): Set yy = Sh.Cells(r - 1 + SymbRow, SymbCol + 2)
    f$ = ""
    If IsNumeric(xx) And IsNumeric(yy) Then
      If xx > x1 And xx < x2 And yy > y1 And yy < y2 Then f$ = Sh.Cells(r - 1 + SymbRow, SymbCol)
    End If
    .Points(r).DataLabel.Text = f$
  Next r
End With
AddSymbCol 3
done: On Error GoTo 0
Sh.Cells(4, 2) = SymbCol: Sh.Cells(15, 2) = SymbRow
Ch.Activate: Sh.Visible = False
Exit Sub
Ng: ExitIsoplot
End Sub

Sub CurveClick()
Attribute CurveClick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim e As EditBoxes, o As Object, L As Labels, G As GroupBoxes, b As Boolean, e1$, e3$
Dim s As String * 5, Shp As Object, Gry&, c1&, c2&
qq = Chr(34)
s = " " & qq & "@" & qq & " "
Gry = Menus("cGray75")
e1$ = "Formula for Y-values, using" & s & "as the ": e3$ = "parametric variable"
AssignD "Curve", , e, , o, L, G, , , , , , Shp
b = IsOn(o("oParam"))
L("lEq2").Visible = b:        e("eEq2").Visible = b
G("gRange").Enabled = b:      G("gCurveTix").Enabled = b
L("lstartparam").Visible = b: e("estartparam").Visible = b
L("lstarttick").Visible = b:  e("estarttick").Visible = b
L("lendparam").Visible = b:   e("eendparam").Visible = b
L("lInterval ").Visible = b:  e("eTickInterval").Visible = b
L("lExmpl2").Visible = b:     G("glabelPos").Enabled = b
If b Then c1 = vbBlack: c2 = vbWhite Else c1 = Gry: c2 = Gry
With Shp("sLabelPos")
  .Line.ForeColor.RGB = c1: .Fill.ForeColor.RGB = c2
End With
o("oLeft").Enabled = b:       o("oRight").Enabled = b
o("oTop").Enabled = b:        o("oBottom").Enabled = b
If b Then
  L("lEq1").Text = Left$(e1$, 12) & "X" & Mid$(e1$, 14) & e3$
  L("lEq2").Text = e1$ & e3$
  L("lExmpl1").Text = "e.g.    Exp(0.984E-3*@)-1"
  L("lExmpl2").Text = "e.g.    Exp(0.155E-3*@)-1"
Else
  L("lEq1").Text = e1$ & "X-variable"
  L("lExmpl1").Text = "e.g.   0.12 + 3.4*@ - 5.6*@^2"
End If
End Sub

Sub AddCurve() ' User specifies a curve to plot
Attribute AddCurve.VB_ProcData.VB_Invoke_Func = " \n14"
Dim e As Object, o As Object, L As Object, G As Object, b As Boolean, OK As Boolean, OkV As Boolean
Dim FormulaX$, FormulaY$, FirstLabel, LabelStep, FirstPar, LastPar, Parametric As Boolean, s$
If Ash.Type <> xlXYScatter Then
  MsgBox "You must start from an Isoplot-Created Chart Sheet for this function", _
  vbOKOnly, Iso: Exit Sub
End If
CurveClick
AssignD "Curve", , e, , o, L, G
Do
  OK = True: OkV = True
  If Not DialogShow("Curve") Then Exit Sub
  Parametric = IsOn(o("oParam"))
  If Parametric Then
    FormulaX$ = e("eEq1").Text: FormulaY$ = e("eEq2").Text
    FirstPar = EdBoxVal(e("eStartParam")): LastPar = EdBoxVal(e("eEndparam"))
    FirstLabel = EdBoxVal(e("eStartTick")): LabelStep = EdBoxVal(e("eTickInterval"))
    If FirstPar = LastPar Or LabelStep <= 0 Or Len(Trim(FormulaX$)) = 0 Then OK = False
  Else
    FormulaY$ = e("eEq1").Text
  End If
  If Len(Trim(FormulaY$)) = 0 Then OK = False
  If InStr(FormulaY$, "@") = 0 Or (Parametric And InStr(FormulaX$, "@") = 0) Then OkV = False
If OK And OkV Then Exit Do
  If Not OkV Then MsgBox "Please use " & qq & "@" & qq & _
    " to indicate the independent variable", , Iso
  If Not OK Then MsgBox "Invalid or missing entry", , Iso
Loop
If Parametric Then ' Trim spaces & leading "y=" or "="
  FormulaX$ = LCase(ApSub(FormulaX$, " ", ""))
  s$ = Left$(FormulaX$, 2)
  If Left$(s$, 1) = "=" Then FormulaX$ = Mid$(FormulaX$, 2)
  If s$ = "x=" Or s$ = "y=" Then FormulaX$ = Mid$(FormulaX$, 3)
End If
FormulaY$ = LCase(ApSub(FormulaY$, " ", ""))
s$ = Left$(FormulaY$, 2)
If Left$(s$, 1) = "=" Then FormulaY$ = Mid$(FormulaY$, 2)
If s$ = "x=" Or s$ = "y=" Then FormulaY$ = Mid$(FormulaY$, 3)
UserCurve FormulaX$, FormulaY$, FirstPar, LastPar, FirstLabel, LabelStep
End Sub

' Currently not used in Isoplot!
Sub PolyRegress(X#(), y#(), Coef As Variant, ByVal N&, ByVal Order%)
Attribute PolyRegress.VB_ProcData.VB_Invoke_Func = " \n14"
' Simple polynomial regression of arbitrary order (eg Order 3 is a cubic).
Dim i&, j%, k%, InvA As Variant, Nterms%
Dim A#(), yy#()
Nterms = 1 + Order
ReDim A(Nterms, Nterms), yy(Nterms, 1)
For i = 1 To N
  For j = 1 To Nterms
    For k = 1 To Nterms
      A(j, k) = A(j, k) + X(i) ^ (j + k - 2)
    Next k
    yy(j, 1) = yy(j, 1) + y(i) * X(i) ^ (j - 1)
Next j, i
With App
  InvA = .MInverse(A)
  If IsError(InvA) Then MsgBox "Error in matrix inversion, sub PolyRegress": ExitIsoplot
  Coef = .MMult(InvA, yy)
End With
End Sub

Sub NoAlerts(Optional Yes As Boolean = True)
Attribute NoAlerts.VB_ProcData.VB_Invoke_Func = " \n14"
App.DisplayAlerts = Not Yes
End Sub

Sub ConvertToPicture(Optional Irange)
Attribute ConvertToPicture.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Pic As Object, Sh0 As Shape, sh1 As Shape
' Convert an Excel wksht-embedded chart into a picture object
Ach.CopyPicture Appearance:=xlPrinter, _
  Size:=xlScreen, Format:=xlPicture
If NIM(Irange) Then
  Irange(1, 3).Select ' Put at right of input data
End If
Ash.Pictures.Paste '.Select
With Ash
  Set sh1 = Last(.Shapes)
  Set Sh0 = .Shapes(.Shapes.Count - 1)
End With
If IM(Irange) Then
  sh1.Left = Sh0.Left: sh1.Top = Sh0.Top
End If
Set Pic = Last(Ash.Pictures)
Sh0.Delete ' Delete the original chart
With Pic.ShapeRange
  With .Fill: .Visible = True: .ForeColor.RGB = IIf(ColorPlot, RGB(200, 255, 255), vbWhite): End With
  .Line.Visible = True
End With
End Sub

Sub CopyPicture(Optional ExternalInvoked = True, Optional Left, Optional Right, Optional Name)
Attribute CopyPicture.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, r As Range, P, W!, H!, c As Object, d As Object, e As Object, cn$
ViM ExternalInvoked, True
NoUp
GetPlotDat P
ViM ExternalInvoked, True
If Not IsObject(P) Then GoTo NoGo
s$ = P.Cells(1, 2).Text
Set d = Sheets(s$): Set c = Ach
cn$ = "ChartToData" & String(-Mac, "2")
GetOpSys
On Error Resume Next
If Windows Or True Then
  For Each e In c.Shapes
    If e.Name = cn$ Then e.Cut
  Next e
  On Error GoTo NoGo
  ' Make sure that the wholoe chart is copied -- not just the last-selected element
  c.PlotArea.Select
  c.CopyPicture ' Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
Else
  'ActiveWindow.SelectedSheets.PrintOut PrintToFile:=True, PrToFileName:="C:\Temp\temp.eps"
  Kill "temp.eps": On Error GoTo NoGo
  c.PrintOut PrintToFile:=True, PrToFileName:="temp.eps" ' Can't specify the name with Macs!
End If
d.Activate ' Switch to source-data sheet
Set r = Range(P.Cells(12, 2))
r(1, r.Columns.Count + 1).Select ' Put picture to right of input data
With Ash.Pictures
  If Windows Or True Then
    .Paste.Select
  Else
    .Insert("temp.eps").Select
    On Error Resume Next: Kill "temp.eps": On Error GoTo NoGo
 End If
End With
With Selection                   ' Scale down size
  W = .Width: H = .Height
  .Width = 300: .Height = H / W * .Width
  If NIM(Left) Then .Left = Left
  If NIM(Right) Then .Right = Right
  If NIM(Name) Then Name = .Name
End With
c.Select
AddCopyButton
d.Select
r(1, r.Columns.Count).Select
Exit Sub
NoGo: On Error GoTo 0
MsgBox "Can only copy a Chart from a Chart sheet"
End Sub

Sub GetPlotDat(P) ' Locate PlotDat sheet
Attribute GetPlotDat.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$
On Error GoTo NoGo
s = Ach.SeriesCollection(1).Formula
Do
  s = Mid$(s, 1 + InStr(s, ","))
Loop Until Left(s, 1) <> ","
s$ = Left(s, InStr(s, "!") - 1)
Set P = Sheets(s)
NoGo:
End Sub

Sub DeleteIsoChart()
Attribute DeleteIsoChart.VB_ProcData.VB_Invoke_Func = " \n14"
Dim P
GetPlotDat P
NoAlerts
On Error GoTo NoGo
Ach.Delete
If IsObject(P) Then P.Delete
NoGo:
End Sub

Sub AssignD(Optional Name, Optional dSheet, Optional EditBoxes, Optional CheckBoxes, _
  Optional OptionButtons, Optional Labels, Optional GroupBoxes, Optional TextBoxes, _
  Optional Buttons, Optional Spinners, Optional DropDowns, Optional ScrollBars, _
  Optional Shapes, Optional Lines, Optional Dframe)
Attribute AssignD.VB_ProcData.VB_Invoke_Func = " \n14"
' Assign variables to dialog-sheet objects.  If the name of the dialog sheet is passed and a dialog-sheet
'  variable, assign the dialog sheet to that named sheet; if no name is passed, must pass the sheet-object itself.
Dim Dlg As Object, i%
If IM(Name) Then
  Set Dlg = dSheet
Else
  Set Dlg = DlgSht(Name)
  If NIM(dSheet) Then Set dSheet = Dlg
End If
With Dlg
  If NIM(EditBoxes) Then Set EditBoxes = .EditBoxes
  If NIM(CheckBoxes) Then Set CheckBoxes = .CheckBoxes
  If NIM(OptionButtons) Then Set OptionButtons = .OptionButtons
  If NIM(Labels) Then Set Labels = .Labels
  If NIM(GroupBoxes) Then Set GroupBoxes = .GroupBoxes
  If NIM(TextBoxes) Then Set TextBoxes = .TextBoxes
  If NIM(Buttons) Then Set Buttons = .Buttons
  If NIM(Spinners) Then Set Spinners = .Spinners
  If NIM(DropDowns) Then Set DropDowns = .DropDowns
  If NIM(ScrollBars) Then Set ScrollBars = .ScrollBars
  If NIM(Shapes) Then Set Shapes = .Shapes
  If NIM(Lines) Then Set Lines = .Lines
  If NIM(Dframe) Then Set Dframe = .DialogFrame
End With
End Sub

Sub CheckInpRange(FromMenu As Boolean, Idat#(), Optional r As Range)
Attribute CheckInpRange.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i&, Nd&, nR&, OK As Boolean, cr As Range, b As Boolean
If FromMenu Then
  nR = N
Else
  GetOpSys
  On Error GoTo BadRange
  Set cr = r.CurrentRegion
  On Error GoTo 0
  If r.Count = 1 And cr.Columns.Count = 2 And cr.Rows.Count > 1 Then
    Set r = cr
    Do
      b = True: nR = r.Rows.Count
      If Not IsNumber(r(1, 1)) Then Set r = Range(r(2, 1), r(nR, 2)): b = False
    Loop Until b Or r.Rows.Count = 1
    Do
      b = True: nR = r.Rows.Count
      If Not IsNumber(r(nR, 1)) Then Set r = Range(r(1, 1), r(nR - 1, 2)): b = False
    Loop Until b Or r.Rows.Count = 1
    r.Select
  End If
  nR = Min(LastOccupiedRow(r.Column), r.Rows.Count)
  N = nR
  ReDim InpDat(nR, 3)
End If
On Error Resume Next
ReDim Idat#(N, 2)
On Error GoTo 0
Nd = 0
For i = 1 To nR
  OK = True
  If Not FromMenu Then
    If Not (IsNumeric(r(i, 1)) And IsNumeric(r(i, 2))) Then
        OK = False
    ElseIf r(i, 1).Font.Strikethrough Or r(i, 2).Font.Strikethrough Then
        OK = False
    ElseIf r(i, 2) <= 0 Or r(i, 1) = "" Or r(i, 2) = "" Then
        OK = False
    End If
  End If
  If OK Then
    Nd = 1 + Nd
    If FromMenu Then
      Idat(Nd, 1) = InpDat(i, 1): Idat(Nd, 2) = InpDat(i, 3)
    Else
      Idat(Nd, 1) = r(i, 1).Value: Idat(Nd, 2) = r(i, 2).Value
      InpDat(Nd, 1) = Idat(Nd, 1): InpDat(Nd, 3) = Idat(Nd, 2)
    End If
  End If
Next i
N = Nd
Exit Sub
BadRange: MsgBox "Invalid input range for isoplot", , Iso
ExitIsoplot
End Sub
Function LastOccupiedRow(ByVal Col%) As Long
LastOccupiedRow = Cells(EndRow, Col).End(xlUp).Row
End Function
Function Last(q As Object) As Object
Attribute Last.VB_ProcData.VB_Invoke_Func = " \n14"
Set Last = q(q.Count)
End Function
Function SQ(ByVal X) As Double ' Kluge to evade bug in Excel 2001
Attribute SQ.VB_ProcData.VB_Invoke_Func = " \n14"
SQ = X * X
End Function
Sub Landscape()
Attribute Landscape.VB_ProcData.VB_Invoke_Func = " \n14"
On Error GoTo A
Ach.PageSetup.Orientation = xlLandscape
Exit Sub
A: DelSheet
HandleNoPrinter
End Sub
Function ApSub(InThis$, ReplaceThis$, WithThis$)
Attribute ApSub.VB_ProcData.VB_Invoke_Func = " \n14"
ApSub = App.Substitute(InThis$, ReplaceThis$, WithThis$)
End Function
Function App() As Excel.Application
Attribute App.VB_ProcData.VB_Invoke_Func = " \n14"
Set App = Application
End Function
Sub Xcalc()
Attribute Xcalc.VB_ProcData.VB_Invoke_Func = " \n14"
ExcelCalc = Qcalc
On Error Resume Next
Sbar = App.DisplayStatusBar
End Sub
Sub Rcalc()
If IsoCalc Then SetCalc ExcelCalc
On Error Resume Next
App.DisplayStatusBar = Sbar
End Sub
Sub AddArRejSymbNote(ByVal Top!)
Attribute AddArRejSymbNote.VB_ProcData.VB_Invoke_Func = " \n14"
' Put at bottom-left of chart if Top=0, above upper-left of plotbox if not
With IsoChrt
  With .TextBoxes.Add(IIf(Top = 0, 10, .Axes(xlCategory).Left), 1, 1, 1)
    .AutoSize = True: .Font.Name = IIf(Mac, "Geneva", "Arial")
    .Font.Size = 9 + Mac
    .Characters.Text = IIf(ColorPlot, "Plateau steps are magenta, rejected steps are cyan", _
      "Plateau steps are filled, rejected steps are open")
      .VerticalAlignment = xlBottom
      .Top = IIf(Top = 0, IsoChrt.ChartArea.Height - 15, Axxis(2).Top - .Height)
  End With
End With
End Sub
Sub AddErrSymbSizeNote(ByVal EsymbPlotted As Boolean, Optional FontSize)
Attribute AddErrSymbSizeNote.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, T$, SL%, PBrt!, PBtop!
s = IIf(FromSquid, "sigma", "s")
If Not CumGauss Then
  If ((EsymbPlotted Or (ProbPlot And pBars) Or WtdAvPlot) And Not AddToPlot) Or ArgonStep Then
    SL = IIf(Opt.AlwaysPlot2sigma Or SigLev = 2, 2, 1)
    If ArgonStep Or ((WtdAvPlot Or AgeExtract) And DoShape) Then
      s$ = "box heights are " & tSt(SL) & s
    Else
      T$ = tSt(SL) & s
      If Eellipse Then
        s$ = "ellipses"
        If SL = 1 Then T$ = "68.3% conf."
      ElseIf Ebox Then
        s$ = "boxes"
      ElseIf eCross Then
        s$ = "crosses"
      Else
        s$ = "symbols"
        If ProbPlot Then T$ = "1s"
      End If
      s$ = "data-point error " & s$ & " are " & T$
    End If
    With IsoChrt
      PBtop = .Axes(xlValue).Top - 3 + 3 * ProbPlot
      PBrt = Right_(.Axes(xlCategory))
      With .TextBoxes.Add(PBrt, PBtop - 1, 1, 1)
        .AutoSize = True:  .Font.Name = "Arial"
        .Characters.Text = s$
        If Not FromSquid Then .Characters(Len(s$), 1).Font.Name = "Symbol"
        .VerticalAlignment = xlBottom: .HorizontalAlignment = xlRight
        If FromSquid Then
          On Error Resume Next
          .Font.Size = 24
          On Error GoTo 0
        ElseIf NIM(FontSize) Then
          .Font.Size = FontSize
        Else
          .Font.Size = 10 + WasPlat
        End If
        If WasPlat And Inverse Then
          .Left = Axxis(1).Left: .Top = 12
        Else
          .Left = PBrt - .Width + 5 * AgeExtract
        End If
        If AgeExtract Then
          .Top = PBtop + 5
        ElseIf FromSquid Then
          .Top = 0
        Else
          .Top = PBtop - .Height + .Font.Size / 2 - 1
        End If
        .Name = "ErrorSize"
        If ArInset And WasPlat Then .ShapeRange.ZOrder msoSendToBack
      End With
    End With
  End If
End If
End Sub

Function PairCompareProb(ByVal First&, Pts#()) As Double
Attribute PairCompareProb.VB_ProcData.VB_Invoke_Func = " \n14"
' returns probability that two x-y pts are the same
Dim Sums#, Bad As Boolean, P#, d#(2, 5)
Dim i&, j%, k%
For i = First To 1 + First
  j = 1 + j
  For k = 1 To 5
    d(j, k) = Pts(i, k)
Next k, i
WtdXYmean d(), 2, 0, 0, Sums, 0, 0, 0, Bad
If Not Bad Then P = ChiSquare(Sums / 2, 2)
PairCompareProb = P
End Function

Function PointDispersion(ByVal Npts&)
Attribute PointDispersion.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Nsim&, i&, s$, Chk%, j%, k&, IndX&()
Dim X#(), Pts#()
ReDim X(Npts), Pts(Npts, 5)
For i = 1 To Npts: X(i) = InpDat(i, 1): Next i
InitIndex IndX(), Npts
QuickIndxSort X(), IndX()
For i = 1 To Npts
  k = IndX(i)
  For j = 1 To 5: Pts(i, j) = InpDat(k, j): Next j
Next i
Erase IndX, X
For i = 1 To Npts - 1
  Nsim = Nsim - (PairCompareProb(i, Pts()) > 0.1)
Next i
If Nsim / Npts > 0.3 Or (Npts = 3 And Nsim = 1) Or (Npts = 4 And Nsim = 2) Then
  s$ = IIf(Npts = 3, "Warning", "Note") & ": " & viv$
  Chk = vbOKOnly
  s$ = s$ & Str(1 + Nsim) & " or more of the data-point pairs from " _
       & "this data-set are almost equivalent (within their assigned errors)," & _
      " so these data do not really constitute a" & Str(Npts) & "-point line."
  If Npts = 3 Then s$ = s$ & viv & "In any case, 3-point isochrons are of dubious reliability."
ElseIf Npts < 4 Then
  s$ = "Warning:" & viv$ & Str(Npts) & "-point isochrons are of doubtful reliability."
  Chk = vbOKOnly
End If
If s$ <> "" Then
  If MsgBox(s$, Chk, Iso) = vbCancel Then ExitIsoplot
End If
End Function
Function OpSys() As String
Attribute OpSys.VB_ProcData.VB_Invoke_Func = " \n14"
OpSys = App.OperatingSystem
End Function
Function Version(Optional Numeric As Boolean = False)
Attribute Version.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s
ViM Numeric, False
s = App.Version
If Numeric Then s = Val(s)
Version = s
End Function
Function Ash() As Object
Attribute Ash.VB_ProcData.VB_Invoke_Func = " \n14"
Set Ash = ActiveSheet
End Function
Function Awb() As Workbook
Attribute Awb.VB_ProcData.VB_Invoke_Func = " \n14"
Set Awb = ActiveWorkbook
End Function
Function Ach() As Chart
Attribute Ach.VB_ProcData.VB_Invoke_Func = " \n14"
Set Ach = ActiveChart
End Function
Function Axxis(ByVal XorY%, Optional ChartObj, Optional AxisGrp% = 1) As Object
Attribute Axxis.VB_ProcData.VB_Invoke_Func = " \n14"
ViM AxisGrp, 1
If IM(ChartObj) Then Set ChartObj = IsoChrt
Set Axxis = ChartObj.Axes(XorY, AxisGrp)
End Function
