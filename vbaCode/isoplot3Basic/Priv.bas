Attribute VB_Name = "Priv"
'Isoplot module Priv
Option Private Module
Option Explicit: Option Base 1

Function TickFor(ByVal tMin, ByVal tMax, ByVal TickInterval) As String
Attribute TickFor.VB_ProcData.VB_Invoke_Func = " \n14"
' Best number-format for axis tick-labels
Dim v#, Ndeci%, MaxDeci%, Tspred#, ti#
ti = 2 * TickInterval: v = Drnd(tMin - ti, 7)
Tspred = tMax - tMin

While v <= tMax
  v = Drnd(v + ti, 7)
  If Abs(v / Tspred) < 0.000001 Then
    v = 0
  ElseIf v >= 1E+15 Then
    TickFor = General: Exit Function
  End If
  Ndeci = NumDeci(v)
  MaxDeci = Max(Ndeci, MaxDeci)
Wend

If MaxDeci Then
  TickFor = "0." & String(MaxDeci, "0")
Else
  TickFor = "0"
End If

End Function

Function NumDeci(ByVal Num) As Integer
Attribute NumDeci.VB_ProcData.VB_Invoke_Func = " \n14"
' Returns # of digits beyond decimal point for Num.
Dim Sval$, P%
Sval$ = Sd$(Num, 7)
P = DecPos(Sval$)

If P = 0 Then
  NumDeci = 0
Else
  Sval$ = Mid$(Sval$, P + 1)
  NumDeci = Len(Sval$)
End If

End Function

Function Min(ByVal Num1, ByVal Num2)
Attribute Min.VB_ProcData.VB_Invoke_Func = " \n14"
Min = IIf(Num1 < Num2, Num1, Num2)
End Function

Function Max(ByVal Num1, ByVal Num2)
Attribute Max.VB_ProcData.VB_Invoke_Func = " \n14"
Max = IIf(Num1 > Num2, Num1, Num2)
End Function

Function MinMax(ByVal MinVal, ByVal MaxVal, ByVal Num)
Attribute MinMax.VB_ProcData.VB_Invoke_Func = " \n14"
MinMax = Min(MaxVal, Max(Num, MinVal))
End Function

Function NumChars(ByVal Number) As Integer  ' Return # of characters in a number             
Attribute NumChars.VB_ProcData.VB_Invoke_Func = " \n14"
NumChars = Len(Sd$(Number, 7)) ' + (Number >= 0)
End Function

Function Prnd(ByVal Number, ByVal Power%)  ' Return a number rounded to
Attribute Prnd.VB_ProcData.VB_Invoke_Func = " \n14"
Prnd = App.Round(Number, -Power) '  specified power-of-10
End Function

Function sn$(ByVal Num As Variant, Optional Signed = False, Optional Zeroed = False)
Attribute sn.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, v  ' String representing number Num, with specified formatting
ViM Signed, False
ViM Zeroed, False
v = Abs(Num): s$ = tSt(v)

If Zeroed And Left$(s$, 1) = Dsep Then s$ = "0" & s$

If Num < 0 Then
  s$ = "-" & s$
ElseIf Num > 0 And Signed Then
  s$ = "+" & s$
ElseIf Num = 0 Then
  s$ = "0"
End If

sn$ = s$
End Function

Function Sd$(ByVal v, ByVal Sigfigs%, Optional Signed = False, Optional Zeroed = False)
Attribute Sd.VB_ProcData.VB_Invoke_Func = " \n14"
' Convert # to string with specified #significant figures, signed or zeroed
ViM Signed, False
ViM Zeroed, False
Sd$ = sn$(Drnd(v, Sigfigs), Signed, Zeroed)
End Function

Function zz(ByVal Num) ' Utility function for various numeric-formatting needs
Attribute zz.VB_ProcData.VB_Invoke_Func = " \n14"
  zz = Int(Log10(Abs(Num)))
End Function

Sub GetMAD(X#(), ByVal N&, MedianVal#, Madd#, Err95#)
Attribute GetMAD.VB_ProcData.VB_Invoke_Func = " \n14"
' Determine the Median Absolute Deviation (MAD) from the median for the first
'   N values in vector X() with median MedianVal.
Dim i&, Tstar#, AbsDev#()
ReDim AbsDev(N)

For i = 1 To N
  AbsDev(i) = Abs(X(i) - MedianVal)
Next i

Madd = App.Median(AbsDev())

Select Case N ' KRL-derived numerical approx., valid for normal distr. w. Tuning=9
  Case Is < 2: Tstar = 0
  Case 2: Tstar = 12.7
  Case 3: Tstar = 15.3
  Case Else
     Tstar = 3.54 / Sqr(N) - 3.92 / N + 70.9 / (N * N) - 60.6 / N ^ 3
End Select

Err95 = Tstar * Madd
End Sub

Function IsNumber(ByVal v, Optional Nonzero As Boolean = False) As Boolean
Attribute IsNumber.VB_ProcData.VB_Invoke_Func = " \n14"
ViM Nonzero, False
IsNumber = (IsNumeric(v) And Not IsEmpty(v) And Not IsNull(v))

If Nonzero And IsNumber Then
  IsNumber = IsNumber And (v <> 0)
End If

End Function

Sub ReferenceChord(t1#, t2#, x1#, y1#, x2#, y2#)
Attribute ReferenceChord.VB_ProcData.VB_Invoke_Func = " \n14"
 ' Calculate endpoints of a concordia-plot reference-chord, forcing the endpoints to fall within the plotbox.
Dim Slope#, Inter#, Yleft#, Yright#
' t1 is lower-X; t2 is higher-X

If (Inverse And t1 < t2) Or (Not Inverse And t2 < t1) Then Swap t1, t2
x1 = ConcX(t1): y1 = ConcY(t1)

If Inverse And t2 = 0 Then  ' Avoid infinite X
  x2 = MaxX: y2 = y1
Else
  x2 = ConcX(t2): y2 = ConcY(t2)
End If

If x1 >= MinX And x2 >= MinX And x1 <= MaxX And x2 <= MaxX And _
   y1 >= MinY And y2 >= MinY And y1 <= MaxY And y2 <= MaxX Then Exit Sub

Slope = (y2 - y1) / (x2 - x1): Inter = y2 - Slope * x2
Yleft = Slope * x1 + Inter: Yright = Slope * x2 + Inter

If Inverse Then
  If Yleft > MaxY Then x1 = (MaxY - Inter) / Slope: y1 = MaxY
  If Yright < MinY Then x2 = (MinY - Inter) / Slope: y2 = MinY
  If x1 < MinX Then x1 = MinX: y1 = Slope * x1 + Inter
  If x2 > MaxX Then x2 = MaxX: y2 = Slope * x2 + Inter
Else
  If Yleft < MinY Then x1 = (MinY - Inter) / Slope: y1 = MinY
  If Yright > MaxY Then x2 = (MaxY - Inter) / Slope: y2 = MaxY
  If x1 < MinX Then x1 = MinX: y1 = Slope * x1 + Inter
  If x2 > MaxX Then x2 = MaxX: y2 = Slope * x2 + Inter
End If

End Sub

Sub GetPlotInfo(OK As Boolean)
Attribute GetPlotInfo.VB_ProcData.VB_Invoke_Func = " \n14"
' Get Plotbox limits, Chart-sheet name, data-sheet name, #data-series, &
'  first empty column in data-sheet
Dim aX As Object, SerC As Object, s$, P%
Dim i%, j%, Exist As Boolean, tB As Boolean
OK = False
On Error GoTo NoGet
With Ash
  Set aX = .Axes: Set SerC = .SeriesCollection(1)
End With
MinX = aX(xlCategory).MinimumScale: MaxX = aX(xlCategory).MaximumScale
MinY = aX(xlValue).MinimumScale:    MaxY = aX(xlValue).MaximumScale
Xspred = MaxX - MinX: Yspred = MaxY - MinY
PlotName$ = Ash.Name
i = InStr(SerC.Formula, ","): j = InStr(SerC.Formula, "!")
PlotDat$ = Mid$(SerC.Formula, i + 1, j - i - 1)
Set ChrtDat = Sheets(PlotDat$)
NoUp
With ChrtDat
  .Visible = True
  DatSheet$ = .Cells(1, 2).Text
  Isotype = .Cells(3, 2):   SymbCol = .Cells(4, 2): SymbRow = Max(1, .Cells(15, 2))
  SigLev = .Cells(5, 2):    AbsErrs = .Cells(6, 2)
  SymbType = .Cells(7, 2):  Inverse = .Cells(8, 2): Normal = Not Inverse
  ColorPlot = .Cells(9, 2): Dim3 = .Cells(10, 2)
  Linear3D = .Cells(11, 2): DoShape = .Cells(13, 2)
  ConcAge = .Cells(15, 2):  InvertPlotType = .Cells(15, 2)
  If Dim3 And Not Linear3D Then Planar3D = True
  .Visible = False
End With
Set DatSht = Sheets(DatSheet$)
DatSht.Select
OK = True
NoGet:
End Sub

Sub PutPlotInfo()
Attribute PutPlotInfo.VB_ProcData.VB_Invoke_Func = " \n14"
' Put basic plot information in left-2 columns of hidden PlotDat
'  sheet so can add to data-sheet & plot later
Dim i%, p_N As Boolean, p_D As Boolean, s$

For i = 1 To Sheets.Count
  s$ = ""
  On Error Resume Next
  s$ = Sheets(i).Name
  On Error GoTo 0
  If PlotDat$ = s$ Then p_D = True
  If PlotName$ = s$ Then p_N = True
Next i

If Not p_D Then GoTo ppiDone
If AddToPlot Then Set ChrtDat = Sheets(PlotDat$)
On Error Resume Next
ChrtDat.Visible = False ' Hide the actual plotting-data sheet
On Error GoTo ppiDone
With ChrtDat
  HA .Columns(1), xlRight
  HA .Columns(2), xlLeft
End With
PltD 1, "Source sheet", DatSheet$, True
PltD 2, "Plot name", PlotName$, True
PltD 3, "Plot Type", Isotype
PltD 4, "1st free col", SymbCol
PltD 5, "Sigma Level", SigLev
PltD 6, "Absolute Errs", AbsErrs
PltD 7, "Symbol Type", SymbType
PltD 8, "Inverse Plot", Inverse
PltD 9, "Color Plot", ColorPlot
PltD 10, "3D plot", Dim3
PltD 11, "Linear", Linear3D
PltD 12, "Data Range", StP.EditBoxes("eRange").Text
PltD 13, "Filled Symbols", DoShape
PltD 14, "ConcAge", (ConcAge Or ConcAgePlot)
PltD 15, "ConcSwap", InvertPlotType
PltD 16, "1st Symbol-row", Max(1, SymbRow)
ChrtDat.Columns("A:B").AutoFit

ppiDone:
On Error Resume Next
If p_N And PlotName$ <> "" Then Sheets(PlotName$).Select
ppi2:
End Sub

Sub PltD(ByVal r%, What$, ByVal Prop, Optional AsText As Boolean = False)
Attribute PltD.VB_ProcData.VB_Invoke_Func = " \n14"
' Fill out 1st two cols of PlotDat sheet.  AsText=true to retain any leading spaces with numbers
With ChrtDat
  .Cells(r, 1) = What$
  .Cells(r, 2).NumberFormat = "@"

  If AsText Then
    .Cells(r, 2) = Prop
  Else
    .Cells(r, 2).Formula = Prop
  End If

End With
End Sub

Sub InitializePlotTypes()
Attribute InitializePlotTypes.VB_ProcData.VB_Invoke_Func = " \n14"
ConcPlot = False:  ArgonPlot = False: PbPlot = False:      UseriesPlot = False
WtdAvPlot = False: CumGauss = False:  ConcAgePlot = False
OtherXY = False:   OtherIndx = 14:    uEvoCurve = False:   uLabelTiks = False
PbTicks = False:   ConcAge = False:   ConcAgePlot = False: PbTickLabels = False
UThPbIso = False:  KCaIso = False:    SmNdIso = False:     ClassicalIso = False
ProbPlot = False:  Stacked = False:   DoMix = False:       AgeExtract = False
ArPlat = False:    ArChron = False:   ArgonStep = False:   StackedUseries = False
YoungestDetrital = False
End Sub

Sub PlotIdentify()  ' Assign descriptive variables to the plot-type from IsoType
Attribute PlotIdentify.VB_ProcData.VB_Invoke_Func = " \n14"
InitializePlotTypes

Select Case Isotype
  Case 1:        ConcPlot = True: ConcAge = False: ConcAgePlot = False
                 If IsOn(StPc("cConcAge")) Then _
                   ConcAge = True: ConcAgePlot = True
  Case 2:        ArgonPlot = True
  Case 3 To 7:   ClassicalIso = True
                 If Isotype = 7 Then KCaIso = True
                 If Isotype = 4 Then SmNdIso = True
  Case 8, 9:     PbPlot = True: PbType = Isotype - 7
  Case 10 To 12: UThPbIso = True
  Case 13:       UseriesPlot = True
  Case 14:       OtherXY = True
  Case 15:       WtdAvPlot = True
  Case 16:       CumGauss = True
  Case 17:       ProbPlot = True
  Case 18:       ArPlat = True:  ArgonStep = True
  Case 19:       ArChron = True: ArgonStep = True
  Case 21:       Stacked = True
  Case 22:       StackedUseries = True
  Case 23:       DoMix = True
  Case 24:       AgeExtract = True
  Case 25:       YoungestDetrital = True
End Select

End Sub

Sub SearchHeader(ByVal RowIn&, ByVal ColIn&, IsAbsErrs As Boolean, _
   IsPercentErrs As Boolean, SL%, hR&, _
  NumIso%, DenomIso%, Optional Gas As Boolean = True)
Attribute SearchHeader.VB_ProcData.VB_Invoke_Func = " \n14"
' Search upwards from cells RowIn,ColIn for indication of an error-header
Dim r&, IsAbs As Boolean, IsPercent As Boolean, SL_%, GotInfo As Boolean, s$
Dim c As Range, pmq As String * 3
ViM Gas, True
r = RowIn: SL_ = 0: IsAbs = False: IsPercent = False
hR = 0: NumIso = 0: DenomIso = 0: pmq = qq & pm & qq

For r = RowIn To 1 Step -1
  Set c = Cells(r, ColIn)
  s$ = Trim(LCase(c.Text))
  ' Strip leading +- if just part of number format for cell

  If IsNumeric(c) And Left(c.NumberFormat, 3) = pmq Then
    If InStr(s$, pm) Then s$ = Mid(s$, 2)
  End If

  HeaderRec s$, IsAbs, IsPercent, SL_, NumIso, DenomIso

  If IsAbs Or IsPercent Or SL_ > 0 Or NumIso > 0 Or DenomIso > 0 Then
    GotInfo = True
    If Gas Then Exit For
  End If

  If Not Gas Then
    Gas = (InStr(s$, "gas") > 0) Or InStr(s$, "moles") > 0 Or _
     (InStr(s$, "39") > 0 And InStr(s$, "ar") > 0)
    If Gas Then hR = r: Exit Sub
  End If

Next r

Gas = False

If GotInfo Then
  IsAbsErrs = IsAbs: IsPercentErrs = IsPercent
  If SL_ > 0 Then SL = SL_
  hR = r
End If

End Sub

Sub HeaderRec(Sin$, IsAbsErrs As Boolean, IsPercentErrs As Boolean, _
  SL%, NM%, dm%)
Attribute HeaderRec.VB_ProcData.VB_Invoke_Func = " \n14"
' For a cell that is likely a header-cell, find out if it is an error-header (specifying
'   possibly the sigma-level & percent/abolute), or if it is an isotope-ratio header
'   (in which case return numerator & denominator isotopes)
Dim P As Boolean, e As Boolean, d%, eP As Boolean, u$, Numer$, Denom$
Dim s$, ss$, i%, L%, q%, si$, sig1, sig2
sig1 = Array("1s", "1-s", "1 sigma", "1sigma", "1-sigma", "68%")
sig2 = Array("2s", "2-s", "2 sigma", "2sigma", "2-sigma", "95%")
Const N$ = "123456789"
's$ = LCase(StripAster(Sin$))
With App
  s = .Substitute(LCase(Sin), "*", "")
  s = .Substitute(s, " ", "")
End With

Do ' strip any initial linefeeds
  si$ = s$
  If Left(s$, 1) = vbLf Then s$ = Mid(s$, 2)
Loop Until si$ = s$

ss$ = Strip(s$, vbLf)
e = (InStr(s$, "error") > 0 Or InStr(s$, "abs") > 0 Or InStr(s$, pm) > 0) _
     Or InStr(s, "+-") > 0 Or InStr(s, "+/-") > 0
If Not e And Right(LCase(Sin$), 4) = " err" Then e = True
P = (InStr(s$, "%") > 0 Or (InStr(s$, "perc") And e))
d = InStr(s$, "/")
If d = 0 Then d = InStr(s$, vbLf) ' Maybe linefeed instead of div-sign
SL = 0

For i = 1 To 6

  If InStr(s$, sig1(i)) Then
    SL = 1: Exit For
  ElseIf InStr(s$, sig2(i)) Then
    SL = 2: Exit For
  ElseIf InStr(s$, "95") And InStr(s$, "conf") Then
    SL = 2: Exit For
  End If

Next i

IsPercentErrs = False: IsAbsErrs = False

If Len(s$) > 0 Then

  If P Then
    IsPercentErrs = True
  ElseIf e Or ss$ = "err" Then
    IsAbsErrs = True
  ElseIf s$ = "x" Then
    NM = -1: dm = -1: Exit Sub
  ElseIf s$ = "y" Then
    NM = -2: dm = -2: Exit Sub
  ElseIf d > 1 And d < Len(s$) Then
    Numer$ = Left$(s$, d - 1): Denom$ = Mid$(s$, d + 1)

    For i = 1 To Len(Numer$)
      u$ = Mid$(Numer$, i)
      If InStr(N$, Left$(u$, 1)) Then
        NM = Val(u$): Exit For
      End If
    Next i

    If NM = 0 Then Exit Sub

    For i = 1 To Len(Denom$)
      u$ = Mid$(Denom$, i)

      If InStr(N$, Left$(u$, 1)) Then
        dm = Val(u$): Exit For
      End If

    Next i

    If dm = 0 Then NM = 0
  End If

End If
End Sub


Sub WhichIsotype(ByVal Xn, ByVal Xd, ByVal Yn, ByVal Yd)
Attribute WhichIsotype.VB_ProcData.VB_Invoke_Func = " \n14"
' Given numerator & denominator isotopes of X & Y isotope-ratios, determine what type of
'  isochron they indicate, & whether normal or inverse.
Dim Isotype0%, InverseIn As Boolean, tB As Boolean, q As Range, test As Boolean
Dim i%, j%, k%, M%, xy%(4), v
Isotype0 = Isotype: Isotype = 0: InverseIn = Inverse: Inverse = False
If Xn = -1 And Xd = -1 And Yn = -2 And Yd = -2 Then Isotype = 14: Exit Sub
Set q = Menus("IsoTypeIsotopes")
xy(1) = Xn: xy(2) = Xd: xy(3) = Yn: xy(4) = Yd

For i = 1 To q.Rows.Count

  For j = 1 To 2
    test = True
    M = (j - 1) * 5

    For k = 2 + M To 5 + M
      v = xy(k - M - 1)
      If v = 0 Or v <> q(i, k) Then test = False
    Next k

    If test Then
      Isotype = q(i, 1): Normal = (j = 1): Inverse = Not Normal
      GoTo Got
    End If

  Next j

Next i

Got:
If Isotype = 13 Then UsType = q(i, 6)
If Isotype = 0 Then Isotype = Isotype0: Inverse = InverseIn
Normal = Not Inverse
PlotIdentify
End Sub

Function StripAster(ByVal s$) As String
Attribute StripAster.VB_ProcData.VB_Invoke_Func = " \n14"
' Strip spaces, asterisks & dashes from s$ & convert to lowercase.
StripAster = LCase(Strip(Strip(Strip(s$, " "), "*"), "-"))
End Function

Sub WhatKindOfData(ByVal Nareas%, ByVal Nrows&, Optional zOptB, Optional zDropD)
Attribute WhatKindOfData.VB_ProcData.VB_Invoke_Func = " \n14"
' Is there a header-row?  If so, try to recognize what type of plot, error-type,
'  & sigma-level from the header-row.
Dim i&, j&, k&, s$, s1$, s2$, s3$, Pass&, GotInfo As Boolean
Dim Aerr1 As Boolean, Perr1 As Boolean, SigLev1%, nAc%, Gas As Boolean
Dim Aerr2 As Boolean, Perr2 As Boolean, SigLev2%, ci As Object, db As Object
Dim Xnm%, Ynm%, Xdm%, Ydm%, r&
Dim p1%, p2%, q1%, q2%, r1%, r2%
Dim cim$(), cima&()

Set ci = Selection:  HeaderRow = False
nAc = 0
With ci

  If ColWise Then
    For i = 1 To Nareas: nAc = nAc + .Areas(i).Columns.Count: Next i
  Else
    nAc = .Areas(1).Columns.Count
  End If

  ReDim cim$(nAc), cima(nAc, 2)
  k = 0: AxX$ = ""

  For i = 1 To Nareas  ' Assemble string-vector of the possible column-names
    If i > 1 And RowWise Then Exit For
    With .Areas(i)

      For j = 1 To .Columns.Count
        k = 1 + k
        cim$(k) = LCase(.Cells(1, j).Text)
        cima(k, 1) = .Cells(1, j).Row
        cima(k, 2) = .Cells(1, j).Column
      Next j

    End With
  Next i

End With

If nAc > 1 And Nrows > 1 Then
  HeaderRec cim$(1), 0, 0, 0, Xnm, Xdm ' X-axis isotopes

  If nAc = 2 Then
    HeaderRec cim$(2), 0, 0, 0, Ynm, Ydm ' Y-axis isotopes

    If Ynm = 0 Or Ydm = 0 Then  ' X-Y data w/o errors? Maybe Wtd Average data?
      ' Look upwards for header row
      SearchHeader cima(2, 1), cima(2, 2), Aerr1, Perr1, SigLev1, r, Ynm, Ydm

      If Aerr1 Or Perr1 Or Ynm Or Ydm Then
        ' Found some indication of header- see if contains isotopes
        If Xnm = 0 And r > 0 Then HeaderRec Cells(r, cima(1, 2)).Text, 0, 0, 0, Xnm, Xdm
        If Not CumGauss And Isotype <> 23 And Isotype <> 22 And Isotype <> 21 _
          And Isotype <> 24 And Isotype <> 25 Then Isotype = 15
        Aerr2 = Aerr1: Perr2 = Perr1: SigLev2 = SigLev1  '  case leave alone.
        If r = cima(2, 1) Then HeaderRow = True ' Header row is indeed the 1st row of data

        If r > 0 Then
          s$ = Trim(Cells(r, cima(1, 2)).Text)

          For i = 1 To Len(s$)   ' Strip vblf's from column-header, replace
            s1$ = Mid$(s$, i, 1)  '  with spaces unless adjacent to "/" or "-",

            If s1$ = vbLf Then   '  in which case just remove.

              If i < Len(s$) And i > 1 Then
                s2$ = Mid$(s$, i + 1, 1): s3$ = Mid$(s$, i - 1, 1)

                If s2$ <> "/" And s2$ <> "-" And s3$ <> "/" And s3$ <> "-" Then
                   AxX$ = AxX$ & " "
                End If

              End If

            Else
              AxX$ = AxX$ & s1$
            End If

          Next i

          If IsNumeric(AxX$) Then AxX$ = ""
        End If

      End If

    End If

  ElseIf nAc = 3 Or nAc = 6 Or nAc = 5 Then   ' Maybe Ar-Ar plateau or PlateauChron data?
      'If Not ArgonStep Then
      's$ = cim$(1)
      Isotype = 0: Gas = False
      SearchHeader cima(1, 1), cima(1, 2), 0, 0, 0, r, 0, 0, Gas
      'Gas = (InStr(s$, "gas") > 0) Or InStr(s$, "moles") > 0 Or _
            (InStr(s$, "39") > 0 And InStr(s$, "ar") > 0)

      If Gas Then

        If nAc = 3 Then ' Ar plateau
          SearchHeader r, cima(3, 2), Aerr1, Perr1, SigLev1, 0, 0, 0

          If SigLev1 Then
            SigLev = SigLev1

            If Aerr1 Or (Not Aerr1 And Not Perr1) Then
              AbsErrs = True
            ElseIf Perr1 Then
              AbsErrs = False
            End If

          ElseIf Aerr1 Then
            AbsErrs = True

          ElseIf Perr1 Then
            AbsErrs = False
          End If

        ElseIf nAc = 5 Or nAc = 6 Then ' Ar PlateauChron
          For i = 1 To nAc: cim$(i) = LCase(Cells(r, cima(i, 2)).Text): Next i
          p1 = InStr(cim$(2), "40"): q1 = InStr(cim$(2), "36"): r1 = InStr(cim$(2), "39")
          p2 = InStr(cim$(4), "40"): q2 = InStr(cim$(4), "36"): r2 = InStr(cim$(4), "39")

        If r1 > 0 And q1 > r1 And p2 > 0 And q2 > p2 Then
          ArChron = True: Isotype = 19: ArgonStep = True
          Inverse = False: Normal = True
        ElseIf r1 > 0 And p1 > r1 And q2 > 0 And p2 > q2 Then
          ArChron = True: Isotype = 19: ArgonStep = True
          Inverse = True: Normal = False
        End If

      ElseIf (InStr(cim$(2), "age") > 0 Or cim$(2) = "t") And (InStr(cim$(3), "er") > 0 _
        Or InStr(cim$(3), pm) > 0) Then
        ArPlat = True: Isotype = 18: ArgonStep = True
      End If

      If Isotype > 0 Then
        zDropD("dIsotype") = Isotype
        ' Look upwards for error-headers
        SearchHeader cima(3, 1), cima(3, 2), Aerr1, Perr1, SigLev1, r, 0, 0
        Perr2 = Perr1: Aerr2 = Aerr1: SigLev2 = 0
        If r = cima(3, 1) Then HeaderRow = True ' Header is 1st data-row
      End If

    End If

  End If

  If nAc > 3 And (Isotype <> 19 And nAc <> 6) Then ' X-Y data with errors?
    HeaderRec cim$(3), 0, 0, 0, Ynm, Ydm ' Y-axis isotopes
    ' Look updard for error-headers
    SearchHeader cima(2, 1), cima(2, 2), Aerr1, Perr1, SigLev1, r, 0, 0
    SearchHeader cima(4, 1), cima(4, 2), Aerr2, Perr2, SigLev2, 0, 0, 0
    If ((Aerr1 Or Perr1) And cima(2, 1) = r) And (Aerr2 Or Perr2) Then HeaderRow = True
    ' If found error headers up-row of data, see if x-y axis isotopes are there too
    If Xnm = 0 And r > 0 Then HeaderRec Cells(r, cima(1, 2)).Text, 0, 0, 0, Xnm, Xdm
    If Ynm = 0 And r > 0 Then HeaderRec Cells(r, cima(3, 2)).Text, 0, 0, 0, Ynm, Ydm
  End If

  WhichIsotype Xnm, Xdm, Ynm, Ydm

  If Isotype > 0 Then
    If NIM(zOptB) Then zOptB("oInverse") = Inverse: zOptB("oNormal") = Not Inverse
    If NIM(zDropD) Then zDropD("dIsoType") = Isotype
    Dim3 = ((nAc = 3 Or nAc > 5) And Not ArgonStep)  '~!
    If NIM(zOptB) Then zOptB(4) = Dim3: zOptB(3) = Not Dim3
    'If IsoType = 13 And Not Dim3 And Not HeaderRow Then UsType = -1  &&
  End If

  If NIM(zOptB) And ((Aerr1 Or Perr1) And (Aerr2 Or Perr2) And Aerr1 = Aerr2 And Perr1 = Perr2) Then
    If Aerr1 Then zOptB("oAbsolute") = xlOn: AbsErrs = True
    If Perr1 Then zOptB("oPercent") = xlOn:  AbsErrs = False
  End If

  j = Max(SigLev1, SigLev2)

  If j > 0 And j < 3 Then
    SigLev = j

    If NIM(zOptB) Then
      If j = 1 Then zOptB("o1sigma") = xlOn Else zOptB("o2sigma") = xlOn
    End If

  End If

End If

If Aerr1 And Aerr2 Then
  AbsErrs = True
ElseIf Perr1 And Perr2 Then
  AbsErrs = False
End If

If SigLev = 0 Then SigLev = 1
HeaderRow = (HeaderRow And Xnm > 0 And Xdm > 0 And Ynm > 0 And Ydm > 0)
End Sub

Sub AddResBox(tbx$, Optional AddRow = 0, Optional AddCol = 0, Optional Clr = Straw, _
  Optional BoxLeft, Optional ByVal WLE, Optional FixedFont = False, Optional NoShadow = False, _
  Optional NoSuper = False, Optional BoxTop, Optional FontSize = 11, Optional Name, _
  Optional Bold = 0, Optional Italics = 0, Optional OnChart = False)
Attribute AddResBox.VB_ProcData.VB_Invoke_Func = " \n14"
' Insert a text-box with the calculation results to the right of the source-data.
Dim s$, L%, dce%, tN$, Bt!, cc%
Dim Sht As Object, r As Object
ViM AddRow, 0
ViM AddCol, 0
ViM Clr, Straw
ViM FixedFont, False
ViM NoShadow, False
ViM NoSuper, False
ViM FontSize, 11 + 2 * Mac
ViM Bold, 0
ViM Italics, 0
ViM Name, ""
StatBar "adding text-box to " & IIf(OnChart, "chart", "data-sheet")

If NIM(WLE) Then
  If Not WLE Then dce = InStr(tbx$, "dce")
End If

L = Len(tbx$)
If L <= 255 Then s$ = tbx$ Else s$ = Left$(tbx$, 255)
NoUp

If OnChart Then
  Set Sht = ActiveSheet
Else
  If Len(DatSheet$) = 0 Then Set DatSht = Ash: DatSheet$ = DatSht.Name
  Set Sht = DatSht
End If

With Sht

  If OnChart Then
    Set r = ActiveChart.PlotArea
  Else
    Set r = Sht.Cells(Max(1, TopRow + AddRow), Max(1, RightCol + 1 + AddCol))
  End If

  With .TextBoxes.Add(r.Left, r.Top, 0, 0)
    .AutoSize = True: .Characters.Text = s$
    .Font.Name = IIf(FixedFont, IIf(Mac, "Monaco", "Courier New"), IIf(Mac, "Geneva", "Arial"))
    .Font.Size = FontSize: .Border.Weight = xlHairline
    .RoundedCorners = True
    If Not NoShadow Then .ShapeRange.Shadow.Type = msoShadow6
    '.Shadow = Not NoShadow
    .ShapeRange.Fill.ForeColor.RGB = Clr
    '.Interior.Color = Clr

    Do

      If L > 255 Then
        tbx$ = Mid$(tbx$, 256)
        L = Len(tbx$)
        If L <= 255 Then s$ = tbx$ Else s$ = Left$(tbx$, 255)
        ' Texboxes.Add method evidently can only tolerate <256-char strings
        cc = 256 + cc
        .Characters(cc).Insert String:=s$
      End If

    Loop Until Len(s$) < 255

    If dce Then ' Indicate use decay-const errs with "dce"
      If Not WLE Then .Characters(dce, 3).Font.Strikethrough = True
    End If  ' Indicate opposite with "dce" in strikethru.

    If NIM(BoxLeft) Then .Left = BoxLeft
    If NIM(BoxTop) Then .Top = BoxTop

    If Name <> "" Then
      .Name = Name: tN$ = Name
    Else
      tN$ = .Name
      Name = tN$
    End If

    If Bold Then .Characters(1, Bold).Font.Bold = True
    If Italics Then .Characters(1, Italics).Font.Italic = True
  End With
  If Not NoSuper Then ConvertSymbols .TextBoxes(tN$) '.Caption
End With
If Not OnChart Then Range(Irange$).Select
StatBar
End Sub

Sub TooLongCheck(StartTime, ByVal MaxTime)                ' If a plot or calculation is slow or
Attribute TooLongCheck.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Elapsed, Msg$, Lpn As Boolean, Lpd As Boolean   '  stalled, ask user if should quit.
Elapsed = Timer() - StartTime

If Elapsed > MaxTime Then
  Msg$ = "This is taking a rather long time -" & vbLf & "Do you want to continue?"

  If MsgBox(Msg$, vbYesNo, Iso) = vbNo Then
    Lpd = (PlotDat$ <> ""): Lpn = (PlotName$ <> "")

    If Lpd Or Lpn Then
      NoAlerts
      If Lpd Then ChrtDat.Delete
      If Lpn Then Sheets(PlotName$).Delete
    End If

    ExitIsoplot
  End If

  StartTime = Timer() + 60 * 60
End If

End Sub

Sub SymbolChar(T As Object, ByVal StartChar%, ByVal NumChars%)
Attribute SymbolChar.VB_ProcData.VB_Invoke_Func = " \n14"
T.Characters(StartChar, NumChars).Font.Name = "Symbol"
End Sub

Sub BadExp()
Attribute BadExp.VB_ProcData.VB_Invoke_Func = " \n14"
MsgBox "Exponential calculation out of range", , Iso: KwikEnd
End Sub

Sub BadLog()
Attribute BadLog.VB_ProcData.VB_Invoke_Func = " \n14"
MsgBox "Log calculation out of range", , Iso: KwikEnd
End Sub

Sub MTrans(Arr#(), TransArr#())
Attribute MTrans.VB_ProcData.VB_Invoke_Func = " \n14"
' Swap elements of an array so can define ellipse coords as a 2-column range
Dim i%, j%, a2%, b2%
a2 = UBound(Arr, 1): b2 = UBound(Arr, 2)
ReDim TransArr(b2, a2)

For i = 1 To a2
   For j = 1 To b2
     TransArr(j, i) = Arr(i, j)
Next j, i

Erase Arr
End Sub

Function SIGN(ByVal A, ByVal b) ' Emulate the FORTRAN function
Attribute SIGN.VB_ProcData.VB_Invoke_Func = " \n14"
SIGN = Abs(A) * Sgn(b)
End Function

Sub TestSqrt(ByVal Number#, SqrNumber#, Bad As Boolean)
Attribute TestSqrt.VB_ProcData.VB_Invoke_Func = " \n14"

If Number >= 0 Then
  SqrNumber = Sqr(Number): Bad = False
Else
  SqrNumber = 0: Bad = True
End If

End Sub

Function IsNum(ByVal s$, ByVal P%) As Boolean ' Is char p in string s$ a number (0-9)?
Attribute IsNum.VB_ProcData.VB_Invoke_Func = " \n14"
Dim A%

If Len(s$) Then
  A = Asc(Mid$(s$, P, 1))
  IsNum = (A > 47 And A < 58)
Else
  IsNum = False
End If

End Function

Function iTan(ByVal Angle#) ' Error-protected (for Excel98) Tan function
Attribute iTan.VB_ProcData.VB_Invoke_Func = " \n14"
iTan = 1E+16
On Error Resume Next
iTan = Tan(Angle)
End Function

Sub NumAndErr(ByVal X#, ByVal XerrIn#, ByVal ErrSigFigs%, xs$, eR$, _
  Optional ByVal Signed = False, Optional ByVal AddPm = False, Optional ByVal Percent = False)
Attribute NumAndErr.VB_ProcData.VB_Invoke_Func = " \n14"
' return formatted value & error as strings with specified sigfigs
Dim Rxerr#, Rx, Nn%, Ndif%, Ndec%, L%, P%, xf$, Erat%, XerrP#, Xerr#, Lx%

ViM Signed, False
ViM AddPm, False
ViM Percent, False
Xerr = Abs(XerrIn)
L = 0

If Xerr = 0 Then
  eR = "0": xs = tSt(X)
  If Percent Then xs = xs & "%"
  If AddPm Then eR = pm & eR
  Exit Sub
End If

If Percent Then
  If X = 0 Then xs = "0": eR = "--": Exit Sub
  XerrP = Drnd(Hun * Xerr / X, ErrSigFigs)
  eR = tSt(XerrP): L = zz(XerrP)
Else
  Rxerr = Drnd(Xerr, ErrSigFigs)
  eR = tSt(Rxerr): L = zz(Rxerr)
  If X <> 0 Then Lx = zz(X)
End If

If Percent Or (InStr(eR, "E") = 0 And Lx > -5 And Lx < 7) Then
  GoSub AddZers
Else
  Erat = Log10(Abs(X / Xerr))

  If Mac Then

    If Erat > -1 Then
      xf = FloatingKluge(X, 1 + Erat)
    Else
      xf = FloatingKluge(X, 3)
    End If

     xs = FloatingKluge(X, 3)
     eR = FloatingKluge(Xerr, ErrSigFigs)

  Else

    If Erat > -1 Then
      xf = "0." & String(Erat + 2, "0") & "E+00"
    Else
      xf = "0.00E+00"
    End If

    xs = Format(X, xf)
    eR = Format(Xerr, "0." & String(Max(0, ErrSigFigs - 1), "0") & "E+00")
  End If

End If

If Signed And X > 0 And Left(xs, 1) <> "+" Then
  xs = "+" & xs
End If

If Percent Then eR = eR & "%"
If AddPm Then eR = pm & eR
Exit Sub

AddZers:
If L >= 0 Then
  Nn = Len(eR) + (InStr(eR, Dsep) > 0)
Else
  Nn = Len(eR) + L
End If

Ndif = ErrSigFigs - Nn

If Ndif > 0 Then
  If DecPos(eR) = 0 Then eR = eR & Dsep
  eR = eR & String(Ndif, "0")
End If

If Left(eR, 1) = Dsep Then eR = "0" & eR
P = DecPos(eR)

If P > 0 Then
  Ndec = Len(eR) - P
  xf = "0." & String(Ndec, "0") & ";;0"
Else
  Ndec = 0:  xf = "0;-0;0"
End If

xs = Format(X, xf)
Return

End Sub

Function DecPos(s$)
DecPos = InStr(s$, ".") ' 09/12/09 was instr(s,Dsep)
End Function

Function VandE(ByVal X#, ByVal XerrIn#, ByVal ErrSigFigs%, _
  Optional Signed = False, Optional Percent = False, Optional Both = False) As String
Attribute VandE.VB_ProcData.VB_Invoke_Func = " \n14"
' Return formatted "xxx +- yyy"
Dim s$, e$, P%
ViM Signed, False
ViM Percent, False
ViM Both, False

If Both Then
  NumAndErr X, XerrIn, ErrSigFigs, s$, e$, Signed
  VandE = s$ & pm & e$ & "  [" & ErFo(X, XerrIn, ErrSigFigs, , True) & "]"
Else
  NumAndErr X, XerrIn, ErrSigFigs, s$, e$, Signed, , Percent
  VandE = s$ & pm & e$
End If

End Function

Function RhoRnd(ByVal Rho#) As String ' Return error correlation as a formatted string
Attribute RhoRnd.VB_ProcData.VB_Invoke_Func = " \n14"
Dim r#, Nd%, L%, Ro$, Ndif%
r = Abs(Rho)
If r > 1 Then
  RhoRnd = "error"
Else
  Nd = 2 - (r > 0.5) - (r > 0.95) - (r > 0.999)
  Ro$ = "0." & String(Nd, "0")
  RhoRnd = Format(Rho, "+" & Ro$ & ";-" & Ro$ & ";" & Ro$)
End If
End Function

Function ProbRnd(ByVal Prob#) As String ' Return probability as formatted string
Attribute ProbRnd.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, q%

If Prob > 1 Then
  ProbRnd = "error"
Else
  q = 2 - (Prob < 0.1 Or Prob > 0.99)
  ProbRnd = Format(Prob, "0." & String(q, "0"))
End If

End Function

Function Mrnd(ByVal M#) As String ' M is an MSWD; return as a formatted string
Attribute Mrnd.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, q%, mm%, zzM%

Select Case M
  Case Is < 0:       s$ = "error"
  Case Is < 0.0001:  s$ = Format(M, "0.000")
  Case Is > 100000#: s$ = FloatingKluge(M, 2)
  Case Is > 100:     s$ = Format(M, "0")
  Case Else
    zzM = zz(M)                  ' Add an extra sf if M =
    mm = Int(M / 10 ^ (zzM - 1)) '  1.0x10^x or 1.1x10^x
    q = MwF - (mm = 11 Or mm = 10) - zzM - 1

    If q > 0 Then
      s$ = Format(M, "0." & String(q, "0"))
    Else
      s$ = Format(M, "0")
    End If

End Select

Mrnd = s$
End Function

Function ErFo(ByVal X#, ByVal Xerr#, ByVal ErrSigFigs%, _
  Optional ByVal AddPm = False, Optional ByVal Percent = False)
Dim e$
ViM AddPm, False
ViM Percent, False
NumAndErr X, Xerr, ErrSigFigs, "", e$, , AddPm, Percent
ErFo = e$
End Function

Function NumDataSeries(ChartSheet As Object) As Integer
Attribute NumDataSeries.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, NdSer%, Pts As Object, Sco As Object, N$, ns%
ns = 0
On Error GoTo NoSeries
Set Sco = ChartSheet.SeriesCollection
On Error GoTo 0

For i = 1 To Sco.Count
  N$ = Sco(i).Name
  If Left$(N$, 6) = "IsoDat" Then ns = 1 + ns
Next i

NoSeries: NumDataSeries = ns
End Function

Function Epsilon(ByVal Age, ByVal Inter)
Attribute Epsilon.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Chur0, ChurSmNd, CHUR, test, MAP As Object
Set MAP = Menus("ModelAgeParams")
ChurSmNd = Val(MAP(2, 2)) ' SourcePD(1)
Chur0 = Val(MAP(2, 3))    ' SourceR(1)
test = Exp(Lambda147 * Age)

If Abs(test) < MAXEXP Then
  CHUR = Chur0 - ChurSmNd * (test - 1)
  Epsilon = 10000 * (Inter / CHUR - 1)
Else
  Epsilon = 0
End If

End Function

Function Strip(ByVal Phrase$, ByVal Del$) As String ' Remove all occurences in Phrase$ of Del$
Attribute Strip.VB_ProcData.VB_Invoke_Func = " \n14"
Strip = ApSub(Phrase$, Del$, "")
End Function

Function Log10(ByVal v)
Attribute Log10.VB_ProcData.VB_Invoke_Func = " \n14"
Log10 = Log(v) / Log_10
End Function

Function Sum(ByVal v As Variant) As Double
Attribute Sum.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, s#

For i = 1 To UBound(v)
  s = s + v(i)
Next i

Sum = s
End Function

Function FloatingKluge(ByVal v#, ByVal Nsig%) As String
Attribute FloatingKluge.VB_ProcData.VB_Invoke_Func = " \n14"
' Return string representation of V rounded to Nsig in scientific notation;
' If Nsig<0, then don't round.
Dim z%, T#, N#, s As String

If Nsig < 0 Then
  FloatingKluge = tSt(v)
Else
  z = zz(v): T = 10 ^ z
  N = Drnd(v, Min(Nsig, 9))
  If Abs(v) > 1 Then s = "+"
  FloatingKluge = tSt(N / T) & "E" & s & tSt(z)
End If

End Function

Function SumProduct(X#(), y#(), Optional z) As Double
Attribute SumProduct.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i&, Sp#, e&
e = UBound(X)

If IM(z) Then

  For i = 1 To e
    Sp = Sp + X(i) * y(i)
  Next i

Else

  For i = 1 To e
    Sp = Sp + X(i) * y(i) * z(i)
  Next i

End If

SumProduct = Sp
End Function

Function SelectFromWholeCols(Rin As Range) As Range
Attribute SelectFromWholeCols.VB_ProcData.VB_Invoke_Func = " \n14"
' From a whole-column range (Rin), return the range from within this range
' consisting of the non-empty cells.
Dim Fr&, Lr&, nc%
If Rin.Rows.Count <> EndRow Then SelectFromWholeCols = Rin: Exit Function
With Rin
  nc = .Columns.Count
  Fr = Cells(1, .Column).End(xlDown).Row
  Lr = Cells(EndRow, .Column).End(xlUp).Row

  If Fr = Lr Then
    Fr = 1
  ElseIf Fr > Lr Then
    Fr = 1: Lr = EndRow
  End If

  Set SelectFromWholeCols = sR(Fr, .Column, Lr, .Column + nc - 1)
End With
End Function
Function Pb7U5(AgeMa) ' Return radiogenic 207Pb/235U (secular equilbrium)
GetConsts
Pb7U5 = Exp(Lambda235 * AgeMa) - 1
End Function
Function Pb6U8(AgeMa) ' Return radiogenic 206Pb/238U (secular equilbrium)
GetConsts
Pb6U8 = Exp(Lambda238 * AgeMa) - 1
End Function
Function Pb8Th2(AgeMa) ' Return radiogenic 208Pb/232Th
GetConsts
Pb8Th2 = Exp(Lambda232 * AgeMa) - 1
End Function

Function Istat$()
Istat$ = "isostat.xls"
End Function

Function PathSep() As String
PathSep$ = App.PathSeparator
End Function

Function IsoPath$()
IsoPath$ = TW.Path
End Function

Function TW() As Workbook
Set TW = ThisWorkbook
End Function

Function DlgSht() As Object
Set DlgSht = TW.DialogSheets
End Function

Function MenuSht() As Worksheet
Set MenuSht = TW.Sheets("Menus")
End Function

Function Menus(r$) As Range
On Error GoTo 1
Set Menus = MenuSht.Range(r$)
Exit Function
On Error GoTo 0
1: Set Menus = MenuSht.Range("NullErr")
End Function

Function StP() As Object
Set StP = DlgSht("IsoSetup")
End Function

Function StPc() As Object
Set StPc = StP.CheckBoxes
End Function

Function NotSmall(v#) ' Return either V or a near-zero value

Select Case Abs(v)             '  (used to avoid div-by-zero errors)
  Case Is > IsoSmall: NotSmall = v
  Case 0: NotSmall = IsoSmall
  Case Else: NotSmall = Sgn(v) * IsoSmall
End Select

End Function

Function NotLarge(v#) ' Return either V or a large-limiting value

If Abs(v) < IsoLarge Then      '  (used to avoid div-by-zero errors)
  NotLarge = v
Else
  NotLarge = Sgn(v) * IsoLarge
End If

End Function

Function LastCol(Optional RowIn) ' Returns last occupied column in current or specified row
If Ash.Type <> xlWorksheet Then LastCol = 0: Exit Function
If IM(RowIn) Then RowIn = ActiveCell.Row
LastCol = Cells(RowIn, 256).End(xlToLeft).Column
End Function

Function LastRow(Optional ColIn) ' Returns last occupied row in current or specified column
If Ash.Type <> xlWorksheet Then LastCol = 0: Exit Function
If IM(ColIn) Then ColIn = ActiveCell.Column
LastRow = Cells(EndRow, ColIn).End(xlUp).Row
End Function
