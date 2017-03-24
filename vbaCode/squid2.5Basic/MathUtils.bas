Attribute VB_Name = "MathUtils"
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
' MathUtils module -- contains mathematical/statistical procedures
Option Explicit
Option Base 1

Sub Swap(a, b)
Dim c As Variant
c = a: a = b: b = c
End Sub

Public Function fiIntLogAbs%(ByVal Num#) ' Utility function
fiIntLogAbs = Int(fdLog10(Abs(Num)))
End Function

Function fdLog10#(ByVal v#)
If v <= 0 Then SqE 23
fdLog10 = Log(v) / pdcLog10
End Function

Function fvMax(ByVal a, Optional b, Optional c, _
  Optional d, Optional e, Optional f, Optional g)
' returns the largest value passed
If fbIM(b) Then
  fvMax = foAp.Max(a)
ElseIf fbIM(c) Then
  If a > b Then fvMax = a Else fvMax = b
Else
  fvMax = Application.Max(a, b, c, d, e, f, g)
End If

End Function

Function fvMin(ByVal a, Optional b, Optional c, _
  Optional d, Optional e, Optional f, Optional g)
' returns the smallest value passed
If fbIM(b) Then
  fvMin = foAp.Min(a)
ElseIf fbIM(c) Then
  If a < b Then fvMin = a Else fvMin = b
Else
  fvMin = foAp.Min(a, b, c, d, e, f, g)
End If

End Function

Public Function fvMinMax(Number As Variant, LowerLimit As Variant, _
                UpperLimit As Variant)
Select Case Number
  Case Is < LowerLimit: fvMinMax = LowerLimit
  Case Is > UpperLimit: fvMinMax = UpperLimit
  Case Else:            fvMinMax = Number
End Select
End Function

Sub BubbleSort(v, Optional Indx, Optional Descending As Boolean = False, _
  Optional AsString As Boolean = False)
' Quick and dirty string-sorter (also numeric, but slow).
Dim NotSwapped As Boolean, DoInd As Boolean, tB As Boolean
Dim i%, iC%, jC%, Lwr%, Uppr%
Dim ct&

DoInd = fbNIM(Indx)
Lwr = LBound(v): Uppr = UBound(v)

If DoInd Then
  ReDim Indx(Lwr To Uppr)
  For i = Lwr To Uppr: Indx(i) = i: Next
End If

ct = 0
Do
  NotSwapped = True
  ct = 1 + ct

  For i = Lwr To Uppr - 1
    iC = i - Descending
    jC = i + 1 + Descending

    If AsString Then
      tB = (StrComp(v(iC), v(jC), 1) = 1)
    Else
      tB = (v(iC) > v(jC))
    End If

    If tB Then
      Swap v(iC), v(jC)
      If DoInd Then Swap Indx(iC), Indx(jC)
      NotSwapped = False
    End If

  Next i

Loop Until NotSwapped

End Sub

Function fiNumDigitsPastDecimal(ByVal Num) As Integer
' Returns # of digits beyond decimal point for Num.
Dim Sval$, p%
Sval$ = fsNumWithSigFigs$(Num, 7)
p = InStr((Sval$), ".")

If p = 0 Then
  fiNumDigitsPastDecimal = 0
Else
  Sval$ = Mid$(Sval$, p + 1)
  fiNumDigitsPastDecimal = Len(Sval$)
End If

End Function

Function fsFormattedNum$(ByVal Num As Variant, _
          Optional Signed = False, Optional Zeroed = False)
Dim s$, v As Variant ' Returns a string representing the number
                     ' Num, with basic formatting.
VIM Signed, False
VIM Zeroed, False
v = Abs(Num): s$ = fsS(v)
If Zeroed And Left$(s$, 1) = Dsep Then s$ = "0" & s$

If Num < 0 Then
  s$ = "-" & s$
ElseIf Num > 0 And Signed Then
  s$ = "+" & s$
ElseIf Num = 0 Then
  s$ = "0"
End If

fsFormattedNum$ = s$
End Function

Function fsNumWithSigFigs$(ByVal v, ByVal Sigfigs%, _
         Optional Signed = False, Optional Zeroed = False)
' Convert # to string with specified #significant figures, signed or zeroed
VIM Signed, False
VIM Zeroed, False
fsNumWithSigFigs$ = fsFormattedNum$(Drnd(v, Sigfigs), Signed, Zeroed)
End Function

Public Function ThUfromPb86#(ByVal AgeMa#, ByVal Pb86#, _
  Optional Alpha, Optional Alpha0, Optional Gamma0)
' Returns 232Th/238U from the measured 208/206 (raw or radiogenic),
'  assuming that the 206/238-208/232 ages are concordant.
Dim Rad86#, t#, Numer#, Denom#, tmp As Variant
tmp = "#NUM!"
On Error GoTo Done
t = AgeMa

If fbIM(Alpha) Then
  Rad86 = Pb86
Else
  Rad86 = (Alpha - Alpha0) / (Alpha * Pb86 - Gamma0)
End If

Numer = Exp(lambda238 * t) - 1
Denom = Exp(lambda232 * t) - 1
tmp = Rad86 * Numer / Denom
Done: On Error GoTo 0
ThUfromPb86 = tmp
End Function

Sub FindCoherentGroup(ByVal AllN%, CoherentN%, BadCt%, BadOnes%(), _
        AllX#(), AllSigmaX#(), CoherentX#(), CoherentSigmaX#(), _
        Mean#, MeanErr#, MSWD#, Probfit#, _
        ByVal MinCoherentProb#, ByVal MinCoherentFract#, _
        OkVal As Boolean)

' Using the median, stepwise reject points with greatest weighted residuals until
' the probability is greater than MinCoherentProb.
' AllN is #points passed to sub; CoherentN is #points in coherent group;
' AllX() and AllSigmaX() are the values & 1-sigma errors of the passed points;
' CoherentX() and CoherentSigmaX() those for the coherent group
' Mean, MeanErr (1-sigma), Mswd, ProbFit refer to the coherent group;
' MinCoherentFract is the minimum# of points required for a coherent group;
' OKval indicates success or lack of finding a coherent group.

Dim MaxResIndx%, SpotCt%, NumCoherent%, CoherentCt%
Dim MinProb#, WtdResid#, Median#, MaxRes#, SCX#(), ScXe#()

CoherentN = 0: OkVal = False: BadCt = 0

ReDim BadOnes(1 To AllN)

For SpotCt = 1 To AllN ' Eliminate spots whose age-errors are greater than the ages

  If AllX(SpotCt) <> 0 Then

    If Abs(AllSigmaX(SpotCt) / AllX(SpotCt)) < 1 Then
      CoherentN = 1 + CoherentN
      ReDim Preserve CoherentX(1 To CoherentN)
      ReDim Preserve CoherentSigmaX(1 To CoherentN)
      CoherentX(CoherentN) = AllX(SpotCt)
      CoherentSigmaX(CoherentN) = AllSigmaX(SpotCt)
    End If

  End If

Next SpotCt

If CoherentN < 2 Then Exit Sub
If (CoherentN / AllN) < MinCoherentFract Then Exit Sub

SimpleWtdAv CoherentN, CoherentX(), CoherentSigmaX(), _
            Mean, MeanErr, , MSWD, Probfit
If Probfit > MinCoherentProb Then OkVal = True: Exit Sub

Do
  If ((CoherentN - 1) / AllN) < MinCoherentFract Or CoherentN < 2 Then
    Exit Sub
  End If
  MaxRes = 0
  Median = foAp.Median(CoherentX())

  For CoherentCt = 1 To CoherentN
    WtdResid = Abs(CoherentX(CoherentCt) - Median) / CoherentSigmaX(CoherentCt)

    If WtdResid > MaxRes Then
      MaxRes = WtdResid
      MaxResIndx = CoherentCt
    End If

  Next CoherentCt

  BadCt = 1 + BadCt
  ReDim SCX(1 To CoherentN), ScXe(1 To CoherentN)

  For CoherentCt = 1 To CoherentN
    SCX(CoherentCt) = CoherentX(CoherentCt)
    ScXe(CoherentCt) = CoherentSigmaX(CoherentCt)
  Next CoherentCt

  CoherentN = CoherentN - 1
  NumCoherent = 0
  ReDim CoherentX(1 To CoherentN), CoherentSigmaX(1 To CoherentN)

  For CoherentCt = 1 To CoherentN + 1

    If CoherentCt <> MaxResIndx Then
      NumCoherent = 1 + NumCoherent
      CoherentX(NumCoherent) = SCX(CoherentCt)
      CoherentSigmaX(NumCoherent) = ScXe(CoherentCt)
    End If

  Next CoherentCt

  SimpleWtdAv CoherentN, CoherentX(), CoherentSigmaX(), _
              Mean, MeanErr, , MSWD, Probfit
Loop Until Probfit > MinCoherentProb

If (CoherentN - 1) / AllN > MinCoherentFract Then OkVal = True
BadCt = 0

For SpotCt = 1 To AllN

  For CoherentCt = 1 To CoherentN
    If CoherentX(CoherentCt) = AllX(SpotCt) Then Exit For
  Next CoherentCt

  If CoherentCt > CoherentN Then
    BadCt = 1 + BadCt
    BadOnes(BadCt) = CoherentCt
  End If

Next SpotCt

If CoherentN > 1 Then OkVal = True
If BadCt > 0 Then ReDim Preserve BadOnes(1 To BadCt)
End Sub

Sub ExtractGroup(ByVal Std As Boolean, Mprob#, t As Range, _
 Optional RedoOnly = False, Optional StdCalc As Boolean = False, _
 Optional AgeResult, Optional TypeCol, Optional DpNum% = 1, _
 Optional Xrange As Range)
' 09/03/25 -- Add "Application.Calculate" before doing any calculations.

' Extract a statistically-coherent age group, construct its range addresses,
'  place results on the active sheet, construct a wtd-average chart inset.

Dim tB As Boolean, OkVal As Boolean, DoAll As Boolean, Rad76age As Boolean
Dim StartNewGrp As Boolean

Dim tW$, Dp$, tmp$, ts1$, ts2$, p$, RejSpots$
Dim OkAgeAddr$(), BadAgeAddr$(), BadAgeErAddr$(), OkAgeErAddr$()

Dim i%, j%, k%, m%, N%, Nn%, Co%, Rw%, Nrej%, NokGrps%, NbadGrps%, LenAddr%
Dim arc%, cc%, nGood%, Nbad%, BadCt%, Nok%, Indx%, OkCt%, rc%
Dim Tct%, OkCt0%, BadCt0%, Er2SigCol%, HdrRowGrp%, HdrRowStd%

Dim OkBreak%(), BadBreak%(), OKrwIndx%(), BadRwIndx%()
Dim TmpBadRwIndx%(), OK%()

Dim ar1&, ArL&, ArN&, Clr&, FirstGrpDatRw&, LastGrpDatRw&, rr&

Dim Mean#, MeanErr#, MSWD#, Prob#, WtdAvg#, WtdAvgErr#
Dim ExtPtSigma#

Dim X#(), Xerr#(), rX#(), RxE#()
Dim OkVals#(), OkErrVals#()
Dim OkAgeVals#(), OkAgeErrVals#(), BadAgeVals#(), BadAgeErrVals#()

Dim ColOne As Range, Col2 As Range, GrpMean As Range, Tt As Range
Dim OKspots As Range, OkErrs As Range
Dim t2 As Range, OutpR As Range, InpR As Range

Dim wW As Variant, vOK() As Variant

Dim OkAge As Range, OkAgeErrs As Range, BadAge As Range, BadAgeErrs As Range
Dim Ra3 As Range, Ra4 As Range, Ra5 As Range, Ra6 As Range, Ra7 As Range
Dim ur1 As Range, Ra1 As Range, Ra2 As Range, Hours As Range

Dim Shp As Object, ShtIn As Worksheet

Const MaxGrpAddrChars = 120 ' Can't have too many chars in seriescollection formula

If fbNIM(AgeResult) Then AgeResult = 0
DoAll = Not (RedoOnly)
Dp = IIf(Std, fsS(DpNum), "")

With t
  ar1 = .Row:           ArN = .Rows.Count
  ArL = ar1 + ArN - 1:  arc = .Column
End With

Set t2 = frSr(ar1, 1 + arc, ArL)
Rad76age = (InStr(Cells(ar1 - 1, arc), fsVertToLF("207Pb|/206Pb")) > 0)
k = 0
ReDim X(1 To ArN), vOK(1 To ArN), rX(1 To ArN), RxE(1 To ArN), Xerr(1 To ArN)

HdrRowGrp = flHeaderRow(-Std)
plHdrRw = HdrRowGrp
FirstGrpDatRw = plaFirstDatRw(-Std)
LastGrpDatRw = plaLastDatRw(-Std)

FindStr "Hours", , piHoursCol, HdrRowGrp
Set Hours = frSr(FirstGrpDatRw, piHoursCol, LastGrpDatRw)

foAp.Calculate

If pbUPb And Not StdCalc Then
  pbSecularTrend = fbDataIsDriftCorr
  pbCanDriftCorr = pbSecularTrend
  plHdrRw = HdrRowGrp
End If

N = ArN
If RedoOnly Or StdCalc Then
  SetArrayVal True, vOK()
Else

  For i = 1 To ArN
    tB = True: vOK(i) = False

    For j = 1 To 2
      tB = tB And (Not IsEmpty(t(i, j))) And IsNumeric(t(i, j))
    Next j

    If tB Then
      k = 1 + k

      If StdCalc Then
        rX(k) = t(i, 1): RxE(k) = t(i, 2)
      Else
        X(k) = t(i, 1):  Xerr(k) = t(i, 2)
      End If

      vOK(i) = True
    End If

  Next i

  Nn = k
  BadCt = N - Nn
  OkCt = Nn
End If

If N < 2 Then Exit Sub


If RedoOnly Or StdCalc Then
  Nn = 0: j = 0: k = 0
  ReDim rX(1 To ArN)

  For i = 1 To ArN
    Set Tt = t(i, 1)

    If Len(Tt.Text) > 0 And IsNumeric(Tt) Then

      If Val(Tt) > 0 Then

        If Not Tt.Font.Strikethrough And fbIsNumber(Tt.Text) Then
          Nn = 1 + Nn
          rX(Nn) = Tt
          RxE(Nn) = t(i, 2) / 100 * Tt
        End If

      End If

    End If

  Next i

  OkCt = Nn: BadCt = ArN - OkCt

ElseIf Not StdCalc Then
  FindCoherentGroup N, Nn, BadCt, BadRwIndx, X(), Xerr(), rX(), _
         RxE(), Mean, MeanErr, MSWD, Prob, Mprob, pdMinFract, OkVal

  If Not OkVal Then
    CFs LastGrpDatRw + 2, arc, "No coherent age group"
    Fonts 2 + LastGrpDatRw, arc, , , vbRed, True, xlCenter, 12, , , , , , , True
    Exit Sub
  End If

Else
  OkVal = (Nn > 2)
End If

If RedoOnly Then
  Rw = Cells.Find(IIf(-StdCalc, "Wtd Mean of", "Mean age of"), _
    Cells(ArL, arc), xlFormulas, xlPart, xlByColumns).Row
Else
  Rw = ArL + 3
End If

If Not OkVal And Not StdCalc Then
  With frSr(FirstGrpDatRw, arc, LastGrpDatRw, 1 + arc)
  End With
  With ActiveWindow
    .ScrollColumn = arc - 4: .ScrollRow = 1

    i = 0
    Do
      i = 1 + i
      .ScrollRow = i
    Loop Until (.VisibleRange.Row + .VisibleRange.Rows.Count) > Rw

  End With
  t.Font.italic = False
End If

If DoAll Then

  If Not StdCalc Then
    Columns(arc).Insert
    Fonts FirstGrpDatRw - 1, arc, , , vbWhite, , , 1, , , "SqidNum"
    ColWidth 0.1, arc
    For i = FirstGrpDatRw To LastGrpDatRw
      Cells(i, arc) = i - FirstGrpDatRw + 1
    Next

    Fonts Columns(arc), , , , vbWhite, , , , , , , pscGen
    arc = arc + 1: TypeCol = TypeCol + 1
  ElseIf Not RedoOnly And DpNum = 1 Then

    For i = 1 To ArN

      If Hours(i) = 0 Then
        t(i, 0) = 0.15 ' for visibility of error bar
      Else
        t(i, 0) = Hours(i)
      End If

    Next i 'sample# col

    Fonts t(0, 0), , , , vbWhite, , , 1, , , "SqidNum"
    Range(t(1, 0), t(ArL, 0)).NumberFormat = pscGen
  ' NOTE: when running this macro Isoplot can't read
  '       attributes of cell fonts - only their contents).
  End If

End If

If StdCalc Then FindStr pscPm & "1s", , i, HdrRowGrp, t.Column
Er2SigCol = 2 - 2 * StdCalc + t.Column

If Not RedoOnly And Not StdCalc And Cells(HdrRowGrp, Er2SigCol) <> "" Then
  Columns(Er2SigCol).Insert ' 2-sig errors

  For i = FirstGrpDatRw To LastGrpDatRw
    Cells(i, Er2SigCol) = "=2*" & Cells(i, Er2SigCol - 1).Address(0, 0)
  Next i

  Cells(HdrRowGrp, Er2SigCol).Formula = "SqidEr"
  ColWidth 0.1, Er2SigCol
  Fonts HdrRowGrp, Er2SigCol, LastGrpDatRw, , vbWhite, , , 8
End If

If Nn = 0 Then
  MsgBox "Can't reduce this data.", , pscSq
  End
End If

ReDim OK(1 To Nn), OKrwIndx(1 To Nn), OkErrVals(1 To Nn)
ReDim OkVals(1 To Nn, 1 To 3), OkVals(1 To Nn, 1 To 2)
ReDim OkErrVals(1 To Nn), OkAgeVals(1 To Nn, 1 To 2), OkAgeErrVals(1 To Nn)
ReDim BadVals(1 To N, 1 To 2), BadErrVals(1 To N)
ReDim BadAgeVals(1 To N, 1 To 2), BadAgeErrVals(1 To N), BadRwIndx(1 To N)

OkCt = 0: BadCt0 = BadCt: BadCt = 0: Tct = t.Rows.Count
If BadCt0 > 0 Then ReDim TmpBadRwIndx(1 To BadCt0)
NokGrps = 0: NbadGrps = 0: StartNewGrp = False

For j = 1 To Tct
  OkCt0 = OkCt

  For i = 1 To UBound(rX)
    Set Tt = t(j, 1)

    If fbIsNumber(Tt.Text) And Not Tt.Font.Strikethrough Then

      If Tt <> 0 And Tt = rX(i) Then
        m = 2 * StdCalc
        Set Ra1 = Tt(1, 0):     Set Ra2 = Tt(1, 1 - m)
        Set Ra3 = Tt(1, 3 - m):  Set Ra4 = Tt(1, 1): Set Ra5 = Tt(1, 2)
        Set Ra6 = Range(Ra4, Ra5): Set Ra7 = Range(Ra1, Ra4)
        Set ur1 = Union(Ra1, Ra2)
        With Ra6
          .Font.Strikethrough = False
          .Interior.Color = IIf(StdCalc, 13434828, vbWhite)
        End With
        OkCt = 1 + OkCt:  OK(OkCt) = rX(i)
        OKrwIndx(OkCt) = j

        If StdCalc Then
          OkVals(OkCt, 1) = Ra1:  OkVals(OkCt, 2) = Ra4
          OkErrVals(OkCt) = Ra5 ' 1sig absolute errs!
        End If

        OkAgeVals(OkCt, 1) = Ra1
        OkAgeVals(OkCt, 2) = Ra2
        OkAgeErrVals(OkCt) = Ra3 ' absolute 2sigma errs!

        If OkCt = 1 Then
          NokGrps = 1
          ReDim OkAgeAddr(1 To 1), OkAgeErAddr(1 To 1), OkBreak%(1 To 1)
          Set OKspots = Ra7: Set OkErrs = Ra5
          Set OkAge = ur1: Set OkAgeErrs = Ra3
        ElseIf StartNewGrp Then
          NokGrps = 1 + NokGrps
          ReDim Preserve OkAgeAddr(1 To NokGrps)
          ReDim Preserve OkAgeErAddr(1 To NokGrps), OkBreak%(1 To NokGrps)
          StartNewGrp = False
          Set OkAge = ur1: Set OkAgeErrs = Ra3
        Else
          Set OkAge = Union(OkAge, ur1)
          Set OkAgeErrs = Union(OkAgeErrs, Ra3)
        End If

        Set OKspots = Union(OKspots, Ra7)
        Set OkErrs = Union(OkErrs, Ra5)
        OkAgeAddr(NokGrps) = OkAge.Address
        OkAgeErAddr(NokGrps) = OkAgeErrs.Address
        OkBreak(NokGrps) = j
        LenAddr = fvMax(Len(OkAgeAddr(NokGrps)), Len(OkAgeErAddr(NokGrps)))

        If LenAddr > MaxGrpAddrChars Then
          StartNewGrp = True
        End If

        Exit For
      End If

    End If

  Next i

  If OkCt = OkCt0 Then
    BadCt = 1 + BadCt
    BadRwIndx(BadCt) = j
  End If

Next j

If BadCt > 0 Then
  k = fvMax(BadCt, BadCt0)
  ReDim Preserve BadRwIndx(1 To k), TmpBadRwIndx(1 To k)
  ReDim TmpBadRwIndx(1 To k)
  NbadGrps = 0: StartNewGrp = False
  For i = 1 To BadCt: TmpBadRwIndx(i) = BadRwIndx(i): Next
  k = 0

  For i = 1 To BadCt
    m = TmpBadRwIndx(i)
    With Range(t(m, 1), t(m, 2))
      .Font.Strikethrough = True
      .Interior.Color = vbYellow
    End With
    Set Tt = t(m, 1)

    If fbIsNumber(Tt.Text) Then

      If Val(Tt) <> 0 Then
        k = 1 + k
        Set Ra1 = Tt(1, 0)
        Set Ra2 = Tt(1, 1 - 2 * StdCalc)
        Set Ra3 = Tt(1, 3 - 2 * StdCalc)
        Set Ra7 = Union(Ra1, Ra2)
        BadRwIndx(k) = m
        BadAgeVals(k, 1) = Ra1
        BadAgeVals(k, 2) = Ra2
        BadAgeErrVals(k) = Ra3

        If k = 1 Then
          NbadGrps = 1
          ReDim BadAgeAddr(1 To 1), BadAgeErAddr(1 To 1), BadBreak%(1 To 1)
          Set BadAge = Ra7
          Set BadAgeErrs = Ra3
        ElseIf StartNewGrp Then
          NbadGrps = 1 + NbadGrps
          ReDim Preserve BadAgeAddr(1 To NbadGrps), BadAgeErAddr(1 To NbadGrps)
          ReDim Preserve BadBreak%(1 To NbadGrps)
          StartNewGrp = False
          Set BadAge = Ra7: Set BadAgeErrs = Ra3
        Else
          Set BadAge = Union(BadAge, Ra7)
          Set BadAgeErrs = Union(BadAgeErrs, Ra3)
        End If

        BadAgeAddr(NbadGrps) = BadAge.Address
        BadAgeErAddr(NbadGrps) = BadAgeErrs.Address
        LenAddr = fvMax(Len(BadAgeAddr(NbadGrps)), Len(BadAgeErAddr(NbadGrps)))
        BadBreak(NbadGrps) = m
        If LenAddr > MaxGrpAddrChars Then StartNewGrp = True
      End If

    End If

  Next i

  BadCt = k
  If NbadGrps > 0 Then
    If BadBreak(NbadGrps) = 0 Then BadBreak(NbadGrps) = m
  End If
  If BadCt > 0 Then ReDim Preserve BadRwIndx(1 To k)
End If

FindStr "SqidNum", , piSqidNumCol, HdrRowGrp

If StdCalc Then
  HdrRowGrp = flHeaderRow(StdCalc)
  FindStr "calibr.const", , piaSacol(DpNum), HdrRowGrp, Choose(DpNum, 1, 1 + piaSacol(1))
  FindStr "Age(Ma)", , piaSageCol(DpNum), HdrRowGrp, 1 + piaSacol(DpNum)
  piaSageEcol(DpNum) = 1 + piaSageCol(DpNum)
  Er2SigCol = 1 + piaSageEcol(DpNum)
  If DpNum = 1 Then piSqidNumCol = piaSacol(DpNum) - 1
End If

For i = 1 To BadCt
  Indx = BadRwIndx(i)
  Set Tt = Range(t(Indx, 1), t(Indx, j))
Next i

For i = 1 To OkCt
  Indx = OKrwIndx(i)
  Set Tt = Range(t(Indx, 1), t(Indx, 1 + j))
Next i

If StdCalc Then
  tB = False

  For Each Shp In ActiveSheet.Shapes
    If Shp.Name = "Redo" Then tB = True: Exit For
  Next

  If Not tB Then AddRedoButton 4 + ArL, 1.8 + arc
ElseIf Not RedoOnly Then
  cc = arc ' arc is the age column

  ' 10/04/19 -- was "Do While".  Corrects mislocated results-captions
  Do Until ActiveSheet.Columns(cc).ColumnWidth < 5 And cc > 1
    ' find the index# column
    cc = cc - 1
  Loop

  'cc = cc + 1 ' 10/04/19 -- deleted  09/12/18 -- added

  ts1$ = "age error (95% conf."
  If Rad76age Then tmp$ = ts1$ & ")" Else ts1$ = ts1$ & ", with"
  Cells(Rw, cc) = "Mean age of coherent group"
  If Not Rad76age Then tmp$ = ts1$ & "out error in Std)"
  Cells(Rw + 1, cc) = tmp$: Cells(Rw + 2, cc) = "MSWD"
  Cells(Rw + 3, cc) = "Probability"
  AddRedoButton ArL + 1.4, cc + 3.3

    Clr = Cells(ar1, arc).Font.Color  ' 10/04/19 -- relocated from below to prevent
    Fonts Rw, cc, Rw + 5, , Clr       '             invisible Rad76 results.
    Fonts Rw, arc, Rw + 5, , Clr

  If Not Rad76age Then
    Cells(Rw + 4, cc) = ts1$ & " error in Std)"
'    Clr = Cells(ar1, arc).Font.Color
'    Fonts Rw, cc, Rw + 5, , Clr
'    Fonts Rw, arc, Rw + 5, , Clr
    On Error Resume Next

    For i = 1 To 5 Step 4
      Cells(Rw + i, cc).Characters(23, 7 + 3 * (i > 1)).Font.Underline = True
    Next i

    On Error GoTo 0
  End If

End If

For i = 1 To t.Rows.Count
  Set ColOne = t(i, 1): Set Col2 = t(i, 2)
  tB = IsNumeric(ColOne) And IsNumeric(Col2)
  If tB Then tB = (ColOne <> 0 And Col2 > 0 And t(i, 1) <> "" And t(i, 2) <> "")
  If tB Then tB = tB And (t(i, 1) <> pdcErrVal And t(i, 2) <> pdcErrVal)
  If tB Then tB = tB And (t(i, 1) <> CSng(pdcErrVal) And t(i, 2) <> CSng(pdcErrVal))

  If tB Then
    Nok = 1 + Nok
  Else
    t(i, 1) = ""
    t(i, 2 - StdCalc) = ""
  End If

Next i

If Std And Nok < 2 Then
  tW$ = "Fewer than 2 Standard spots with useable data -- Must quit."
  If piStdCorrType = 2 Then tW$ = tW$ & pscLF2 & _
    "(try selecting 204- or 207-correction)"
  MsgBox tW$, , pscSq
  Alerts False
  If Workbooks.Count > 0 Then
    With ActiveWorkbook
      If .Name <> ThisWorkbook.Name Then .Close
    End With
  End If
  End
End If

tW$ = Range(t(1, 1), t(ArN, 2)).Address(False, False)
Set InpR = Range(tW)
Set OutpR = frSr(Rw, arc, Rw + 6)

If StdCalc And RedoOnly Then
  ' Percent in and out, canNOT reject   'if a Redo

  ' ------------------------------------------------------
  wW = Isoplot3.wtdav(InpR, True, True, 1, False, True, 1)
  ' ------------------------------------------------------

  k = (wW(3, 2) = "MSWD")
  WtdAvg = wW(1, 1)
  WtdAvgErr = wW(2, 1)
  ExtPtSigma = IIf(k, 0, wW(3, 1))
  MSWD = wW(4 + k, 1)
  Prob = wW(6 + k, 1)
  RejSpots = "": Nrej = 0

  For i = 1 To t.Count

    If t.Cells(i, 1).Font.Strikethrough Then
      Nrej = 1 + Nrej
      RejSpots = RejSpots & IIf(Nrej > 1, ",", "") & fsS(i)
    End If

  Next i

  If Nrej = 0 Then RejSpots = "none"
ElseIf Not StdCalc Then
  ' Abs errs in and out, can't reject, do NOT assume an external error!
  wW = Isoplot3.wtdav(InpR, False, False, 1, False, False, 2)
  WtdAvg = wW(1, 1):        WtdAvgErr = wW(2, 1)
  MSWD = wW(3, 1)
  Prob = wW(5, 1)
  OutpR(1) = WtdAvg: OutpR(2) = WtdAvgErr

  If MSWD >= 100 Then
    MSWD = Drnd(MSWD, 3)
    OutpR(3).NumberFormat = "general"
  End If

  OutpR(3) = MSWD
  OutpR(4) = Prob
End If

If RedoOnly Then
  OutpR(7).Formula = ""
ElseIf Not StdCalc Or DpNum = 1 Then
  Fonts OutpR(7), , , , vbRed, True, xlRight, 12, , , pscFluff, , "arial narrow"
End If

If Not StdCalc And fbNIM(AgeResult) Then AgeResult = OutpR(1)
Set GrpMean = frSr(Rw - 2, arc - 2, Rw - 1, arc - 1)
'frSr(rw - 2, arc - 2, rw - 1, arc - 1)
GrpMean.Font.Color = vbWhite

If StdCalc Then
  GrpMean(1, 1) = 0
  GrpMean(2, 1) = 1 + Int(Hours(Hours.Rows.Count))
  tmp = "=StdAge" & IIf((pbU And DpNum = 1) Or (pbTh And DpNum = 2), "UPb", "ThPb")
  GrpMean(1, 2) = tmp
  GrpMean(2, 2) = tmp
  With Range("WtdMeanA" & Dp)
    rr = .Row + 8
    rc = .Column
  End With

  If RedoOnly Then
    OutpR(1) = WtdAvg
    OutpR(2) = WtdAvgErr
    OutpR(3) = ExtPtSigma

    If MSWD >= 100 Then
      MSWD = Drnd(MSWD, 3)
      OutpR(4).NumberFormat = "general"
    End If

    OutpR(4) = MSWD
    OutpR(5) = Prob
    OutpR(6) = RejSpots
  End If
  ' ????
Else
  'tmp$ = OutpR(1, -1).Text
  FindStr "Mean age of coherent group", Rw, Co, , , pemaxrow
  If Co = 0 Or Rw = 0 Then
    Co = GrpMean(1, 1).Column - 3
    Rw = GrpMean(1, 1).Row
  End If
  tmp = Cells(Rw, Co).Text

  If InStr(tmp, "(N=") = 0 Then
    tmp = tmp & " (N="
  End If

'  Do
'    i = InStr(tmp$, " (N=")
'    If i = 0 Then tmp$ = tmp$ & " (N="
'  Loop Until i > 0

  i = InStr(tmp, "(N=")  ' 10/4/28 -- added
  tmp = Left(tmp, i + 2) '    "          "
  tmp$ = tmp & fsS(Nn) & ")" 'Left$(tmp$, i + 3) & fsS(Nn) & ")"
  Cells(Rw, Co) = tmp 'OutpR(1, -1) = tmp$
  GrpMean(1, 1) = 0.5: GrpMean(2, 1) = 1 + ArN
  GrpMean(1, 2) = "= " & Cells(Rw, arc).Address
  GrpMean(2, 2).Formula = GrpMean(1, 2).Formula
  RangeNumFor pscZd1, Rw, arc, Rw + 1
  RangeNumFor pscZd2, Rw + 2, arc, Rw + 4
  Fonts Rw, arc - 3, Rw + 6, arc, , True, xlRight ' was arc - 2
  rr = Rw + 7: rc = arc + 5

  If Not Rad76age Then
    With OutpR(5)
      ts1$ = "= SQRT(" & Cells(Rw + 1, arc).Address & "^2 + "
      ts2$ = "(" & pscStdShtNot & "WtdMeanAperr1/100*2*" & _
              Cells(Rw, arc).Address & ")^2)"
      .Formula = ts1$ & ts2$
      .NumberFormat = Cells(Rw + 1, arc).NumberFormat
    End With
  End If

End If

If BadCt0 > 0 And StdCalc Then
  i = GrpMean.Row: j = GrpMean.Column
  FindStr pscRejectedSpotNums, p, , Rw, arc - 2, Rw + 10, arc + 2

  If p > 0 Then
    Set Tt = Cells(p, arc)
    Tt = "": tmp = ""

    For i = 1 To BadCt0
      j = TmpBadRwIndx(i)
      tmp = tmp & fsS(j)
      If i < BadCt0 Then tmp = tmp & ","
    Next i

    Tt = tmp
    HA Choose(DpNum, xlLeft, xlRight), Tt
  End If

End If
If BadCt > 0 Then
  If foAp.Sum(BadAgeErrs) = 0 Then BadCt = 0
End If
foAp.Calculate

' ******************************************************************
AddGroupAgeChart rr, rc, GrpMean, DoAll, (StdCalc), NokGrps, NbadGrps, _
  OkCt, BadCt, OkAgeAddr, BadAgeAddr, OkAgeErAddr, BadAgeErAddr, _
  OkAgeVals, OkAgeErrVals, BadAgeVals, BadAgeErrVals, 1, Hours, DpNum
' ******************************************************************

ChartShow ActiveSheet.ChartObjects.Count

If Not StdCalc Then
  With ActiveWindow
    .ScrollColumn = fvMax(1, .ScrollColumn - 13)
  End With
  Cells(rr - 1, rc - 3).Activate
End If

End Sub

Sub SimpleWtdAv(ByVal N%, X#(), Sigma#(), Mean#, Optional SigmaMean, _
  Optional Err95, Optional MSWD, Optional Probfit)
' Calculate a simple, inverse-variance weighted average.
Dim i%, Nn%, df%, w#, s#, sw#, Sx#, m#

For i = 1 To N
  If IsNumeric(X(i)) And X(i) <> 0 Then
    Nn = 1 + Nn
    w = 1 / Sigma(i) ^ 2
    Sx = Sx + w * X(i): sw = sw + w
  End If
Next i

If Nn = 0 Then
  Mean = 0: Exit Sub
ElseIf Nn = 1 Then
  Mean = Sx / w: m = 0
  If fbNIM(SigmaMean) Then SigmaMean = sqR(1 / sw)
  If fbNIM(Err95) Then Err95 = 2 * SigmaMean
  If fbNIM(Probfit) Then Probfit = 1
Else
  df = Nn - 1: Mean = Sx / sw
  SigmaMean = sqR(1 / sw)
  If fbNIM(Err95) Then Err95 = 2 * SigmaMean
  s = 0

  For i = 1 To N
    If IsNumeric(X(i)) And X(i) <> 0 Then
      s = s + ((X(i) - Mean) / Sigma(i)) ^ 2
    End If
  Next i

  m = s / df
End If

If fbNIM(MSWD) Then MSWD = m

If fbNIM(Err95) And df > 0 Then
  If ChiSquare((m), df) < 0.1 Then Err95 = Err95 / 1 * sqR(m) * StudentsT(df)
End If

If fbNIM(Probfit) And df > 0 Then
  Probfit = ChiSquare(m, df)
End If

End Sub

Sub SortUpDown(UpDown%)
' In response to user-click of up or down arrow on grouped-sample worksheet.
Dim Hr&, c%, i%, SortCol%, Nr%, EndCol%, q%()
Dim rr&, rw2&, rw1&

Hr = flHeaderRow(False)
If Hr = 0 Then Exit Sub
SortCol = Selection.Column
EndCol = fiEndCol(Hr)
If EndCol = 255 Or EndCol < 2 Then Exit Sub

c = 0
Do
  c = c + 1
Loop Until Cells(Hr, c) = "SqidNum" Or c = 255

rw1 = 1 + Hr
rw2 = flEndRow
Nr = rw2 - rw1 + 1
If Nr = 0 Then Exit Sub

If c = 0 Or c = 255 Then
  frSr(rw1, 1, rw2, EndCol).Sort Key1:=Cells(rw1, SortCol), _
       Order1:=UpDown, Header:=xlNo, _
       OrderCustom:=1, MatchCase:=False, _
       Orientation:=xlTopToBottom, _
       DataOption1:=xlSortNormal
Else
  rr = Hr
  ReDim q(1 To 999)

  Do
    rr = rr + 1
    i = rr - Hr
    q(i) = Cells(rr, c)
  Loop Until q(i) = 0

  rw2 = rr - 1
  ReDim Preserve q(1 To i - 1)
  On Error GoTo 1
  frSr(rw1, 1, rw2, EndCol).Sort Key1:=Cells(rw1, SortCol), _
     Order1:=UpDown, Header:=xlNo, _
     OrderCustom:=1, MatchCase:=False, _
     Orientation:=xlTopToBottom

  If rr > 0 Then
    For rr = 1 + Hr To rw2
      Cells(rr, c) = q(rr - Hr)
    Next rr
  End If

End If
2: Redo
Exit Sub

1
If Err.Number = 1004 Then
      MsgBox "Sorry, SQUID can't sort a grouped sheet containing Array Functions", , pscSq
      On Error GoTo 0
      GoTo 2
    End If
End Sub

'Sub xSortUp_Click()
'SortUp
'End Sub
'Sub xSortDown_Click()
'SortDown
'End Sub
Sub SortUp()
SortUpDown 1
End Sub
Sub SortDown()
SortUpDown 2
End Sub

Sub Xintercept(Xinter#, XinterErr#, ByVal Yinter#, _
   ByVal YinterErr#, ByVal Slope#, ByVal SlopeErr#, _
   ByVal Xbar#)
' Return the x-intercept & error of a regressed line.
Dim a#, b#, c#, q#, root1#, root2#, discr#

Xinter = -Yinter / Slope
a = Slope ^ 2 - SlopeErr ^ 2
b = 2 * (Slope * Yinter + SlopeErr ^ 2 * Xbar)
c = Yinter ^ 2 - YinterErr ^ 2
discr = b * b - 4 * a * c
XinterErr = 0

If discr >= 0 Then
  q = -(b + Sgn(b) * sqR(discr)) / 2

  If a <> 0 Then
    root1 = q / a

    If q <> 0 Then
      root2 = c / q
      XinterErr = Abs(root2 - root1) / 2
    End If

  End If

End If
End Sub

Function sqWtdAv(Values As Range, Errs As Range, _
  Optional PercentErrsIn As Boolean = False, _
  Optional PercentErrsOut As Boolean = False, _
  Optional CanReject As Boolean = False) As Variant
' Calculate an error-weighted average using Isoplot's procedure,
'  demanding calculation of an external error if prob-fit is low;
'  return an array with the various wtd-avg statistics & headers.
Dim b As Boolean, k%, ss$, e$, w As Variant
Dim s(1 To 6, 1 To 2) As Variant, ValuesErrs As Range

Set ValuesErrs = Union(Values, Errs)
w = Isoplot3.wtdav(ValuesErrs, PercentErrsIn, PercentErrsOut, _
                   1, CanReject, True, 1)
' W(1,1)= wtd mean internal or external
' W(2,1)= 1-sigma or 68% conf.error
' W(3,1)= required external 1-sigma       (or MSWD)
' W(4,1)= MSWD                            (or #rejected)
' W(5,1)= #rejected                       (or Prob fit)
' W(6,1)= ProbFit
' W(7,1)= rej. item #s

If IsNull(w) Then
  sqWtdAv = "=#ERR": Exit Function
Else

  If IsNumeric(w(3, 1)) Then
    b = (w(3, 1) <= 0.0001)
  Else
    b = True
  End If

  e = IIf(PercentErrsOut, "%", "")
  k = (w(3, 2) = "MSWD") ' ie no external error
  s(1, 1) = w(1, 1): s(1, 2) = "wtd mean"
  s(2, 1) = w(2, 1): s(2, 2) = pscPm & e & "(" & _
           IIf(b, "1sigma", "68% conf") & ")"
  s(3, 1) = IIf(k, 0, w(3, 1))
  s(3, 2) = "ext. err" & e
  s(4, 1) = foAp.Fixed(w(6 + k, 1), 3)
  s(4, 2) = "prob. fit"

  If CanReject Then
    s(5, 1) = fsS(w(5, 1))
    s(5, 2) = "#rejected"

    If s(5, 1) > 0 Then
      ss = w(7, 1)
      Subst ss, " ", ", "
      s(6, 1) = ss
      s(6, 2) = "rejected item #s"
    Else
      s(6, 1) = "": s(6, 2) = ""
    End If

  End If

End If
sqWtdAv = s
End Function

Public Sub SolveThis(ToMinimize, ParamsToVary, Max1Min2Val3)
' Invoke Excel's SOLVER for specdified cells.
' Best initial value of ParamsToVary must already be set.
' MaxMinVal=1 for max, 2 for min, 3 for set to value
SolverOK SetCell:=ToMinimize, MaxMinVal:=Max1Min2Val3, _
         ByChange:=ParamsToVary
SolverSolve True
End Sub

'Public Function sqSolve(tominimize As Range, params As Range, MaxMinVal%)
'SolveThis tominimize, params, MaxMinVal
'End Function

Sub WtdAvCorr(Values#(), VarCov#(), ByVal N&, MeanVal#, _
  SigmaMeanVal#, MSWD#, Prob#, SigRho As Boolean, Bad As Boolean)
' Weighted average of a single variable (Values) whose errors of
'  index-adjacent values are correlated.
' If SigRho is True, then VarCov() contains sigma's and rho's
'  instead of variances & covariances.
Dim i&, j&
Dim Numer#, Denom#, Sums#, OMij#, SumWtdResids As Variant, SumOmega#, OmI#
Dim OmegaInv#(), UnwtdResids#()
Dim TransUnwtdResids As Variant, TempMat As Variant, Omega As Variant
ReDim OmegaInv(1 To N, 1 To N), UnwtdResids(1 To N, 1 To 1)

For i = 1 To N ' Construct variance-covariance matrix
  For j = 1 To N

    If Not SigRho Then
      OmI = VarCov(i, j)
    ElseIf i = j Then ' convert sigma's to variances
      OmI = VarCov(i, i) ^ 2
    Else              ' convert rho's to covariances
      OmI = VarCov(i, j) * VarCov(j, j) * VarCov(i, i)
    End If

    OmegaInv(i, j) = OmI
Next j, i

Bad = True
On Error GoTo BadMat

Omega = foAp.MInverse(OmegaInv)

If IsError(Omega) Then Exit Sub

For i = 1 To N
  For j = 1 To N
    If N > 1 Then                     ' 10/10/05 -- because if N=1 the matrix inverse
      OMij = Omega(i, j)              '             has only one dimension.
    Else
      OMij = Omega(1)
    End If
    Numer = Numer + (Values(i) + Values(j)) * OMij
    Denom = Denom + OMij
Next j, i
If Denom <= 0 Then Exit Sub

MeanVal = Numer / Denom / 2
SigmaMeanVal = sqR(1 / Denom)

For i = 1 To N
  UnwtdResids(i, 1) = Values(i) - MeanVal
Next i

With foAp
  TransUnwtdResids = .Transpose(UnwtdResids)
  TempMat = .MMult(TransUnwtdResids, Omega)
  SumWtdResids = .MMult(TempMat, UnwtdResids)
End With

If N = 1 Then                        ' 10/10/05 -- added the "If N=1" part to avoid div-by-zero
  MSWD = 0: Prob = 0
Else
  MSWD = SumWtdResids(1) / (N - 1)
  Prob = ChiSquare(MSWD, N - 1)
End If
Bad = False

BadMat: On Error GoTo 0
End Sub


Public Sub PoissonLimits(ByVal Counts&, RejectIfLessThan&, _
                         RejectIfGreaterThan&)
' Using 95%-conf. limits from a Monte Carlo simulation, determine rejection-
'  cutoff values for various #counts.
Dim Offs&, ArrL As Variant, ArrU As Variant

If Counts < 0 Then
  Offs = 2 ^ 31 - 1
  RejectIfLessThan = -Offs
  RejectIfGreaterThan = Offs
ElseIf Counts = 0 Then
  RejectIfLessThan = 0
  RejectIfGreaterThan = 6 ' arbitrary, very conservativ e(?)
ElseIf Counts <= 100 Then

  ArrL = Array(0, 0, 0, 1, 1, 2, 2, 3, 4, 4, 5, 6, 6, 7, 8, 9, 9, 10, _
    11, 12, 13, 13, 14, 15, 16, 17, 17, 18, 19, 20, 21, 21, 22, 23, _
    24, 25, 26, 26, 27, 28, 29, 30, 31, 31, 32, 33, 34, 35, 36, 37, _
    38, 38, 39, 40, 41, 42, 43, 44, 44, 45, 46, 47, 48, 49, 50, 51, _
    51, 52, 53, 54, 55, 56, 57, 58, 59, 59, 60, 61, 62, 63, 64, 65, _
    66, 67, 67, 68, 69, 70, 71, 72, 73, 74, 75, 75, 76, 77, 78, 79, 80, 81)

  ArrU = Array(2, 4, 6, 7, 9, 10, 12, 13, 14, 16, 17, 18, 20, 21, 22, 23, _
    25, 26, 27, 28, 29, 31, 32, 33, 34, 35, 37, 38, 39, 40, 41, 43, 44, _
    45, 46, 47, 48, 50, 51, 52, 53, 54, 55, 56, 58, 59, 60, 61, 62, 63, _
    64, 66, 67, 68, 69, 70, 71, 72, 74, 75, 76, 77, 78, 79, 80, 81, 82, _
    84, 85, 86, 87, 88, 89, 90, 91, 93, 94, 95, 96, 97, 98, 99, 100, 101, _
    103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 114, 115, 116, 117, _
    118, 119)

  RejectIfLessThan = ArrL(Counts)
  RejectIfGreaterThan = ArrU(Counts)
Else
  Offs = 2 * sqR(Counts)
  RejectIfLessThan = Counts - Offs
  RejectIfGreaterThan = Counts + Offs
End If
End Sub

Function SumSquares(N%, Resids#(), Omega) As Double
Dim i%, d%, Tr#(), residsT As Variant, Mm As Variant
' Return sum of squared residuals.  Doesn't matter if Resids is (N) or (N,1)
d = 0
On Error Resume Next
d = UBound(Resids, 2)
On Error GoTo 0

If d < 1 Then
  ReDim Tr(1 To N, 1 To 1)
  For i = 1 To N: Tr(i, 1) = Resids(i): Next
End If

With foAp

  If d > 0 Then
    residsT = .Transpose(Resids)
  Else
    residsT = .Transpose(Tr)
  End If

  If d > 0 Then
    Mm = .MMult(.MMult(residsT, Omega), Resids)
  Else
    Mm = .MMult(.MMult(residsT, Omega), Tr)
  End If

End With
SumSquares = Mm(1)
End Function

Sub CopyDblArray(SourceVect#(), CopiedVect#(), Optional Ndim% = 0)
' Copy a double-precision array
Dim i%, j%, Ub1%, lb1%, Ub2%, lb2%

If Ndim = 0 Then
  Ub1 = UBound(SourceVect)
  lb1 = LBound(SourceVect)
  ReDim CopiedVect(lb1 To Ub1)

  For i = lb1 To Ub1
    CopiedVect(i) = SourceVect(i)
  Next i

Else
  Ub1 = UBound(SourceVect, 1)
  lb1 = LBound(SourceVect, 1)
  Ub2 = UBound(SourceVect, 2)
  lb2 = LBound(SourceVect, 2)
  ReDim CopiedVect(lb1 To Ub1, lb2 To Ub2)

  For i = lb1 To Ub1
    For j = lb2 To Ub2
     CopiedVect(i, j) = SourceVect(i, j)
  Next j, i

End If
End Sub

Sub CopyVarArray(SourceVect(), CopiedVect(), Optional Ndim% = 0)
' Copy a Variant array
Dim i%, j%, Ub1%, lb1%, Ub2%, lb2%

If Ndim = 0 Then
  Ub1 = UBound(SourceVect)
  lb1 = LBound(SourceVect)
  ReDim CopiedVect(lb1 To Ub1)

  For i = lb1 To Ub1
    CopiedVect(i) = SourceVect(i)
  Next i

Else
  Ub1 = UBound(SourceVect, 1)
  lb1 = LBound(SourceVect, 1)
  Ub2 = UBound(SourceVect, 2)
  lb2 = LBound(SourceVect, 2)
  ReDim CopiedVect(lb1 To Ub1, lb2 To Ub2)

  For i = lb1 To Ub1
    For j = lb2 To Ub2
     CopiedVect(i, j) = SourceVect(i, j)
  Next j, i

End If
End Sub

'Sub test()
'Dim r, ra As Range, bmean6, bmean9, avg, Nf%
'Dim scts#(10), Pav#, ct#(10), i, d#(10), nff%
'Dim Nrej, Lwr&, Upr&
'
'For r = 3 To 120
'  Set ra = frSr(r, 2, , 11)
'  For i = 1 To 10: ct(i) = ra(1, i): Next i
'  PoissonLimits foAp.Median(ra), Lwr, Upr
'  bmean6 = isoplot3.biweight(ra, 6)
'  bmean9 = isoplot3.biweight(ra, 9)
'  PoissonOutliers 10, ct, d, Pav, 0, 0, 0, nff, 1
'  Nrej = 10 - nff
'  Cells(r, 12) = Lwr
'  Cells(r, 13) = Upr
'  Cells(r, 14) = Pav / 10
'  Cells(r, 15) = IIf(Nrej = 0, "", Nrej)
'  Cells(r, 16) = bmean6(1, 1)
'  Cells(r, 17) = bmean9(1, 1)
'  Cells(r, 18) = foAp.Average(ra)
'Next r
'End Sub

Sub PoissonOutliers(Nfields%, sub10PkCts#(), sub10SbmCts#(), _
    CorrPkTotCts#, SigmaMeanPkTotCts#, SbmTotCts#, _
    SigmaMeanSBMtotCts#, Nf%, IntegrTime#)
' If median counts for the 10 integrations is >0, use
'  TukeysBiweight to get mean, SigmaMeanPkCts, & total cts.
' If <=100 cts each, look for outliers based on the asymmetric
'  95% conf-limits from a Poisson distribution.
' 09/06/18 -- Modified so that only 1 rejection is permitted,
'             and that one must correspond to the dataum with
'             the largest residual.
Dim sumX#, sumX2#, med#, IntegrN%, MaxDeltIndx%, PkI#, PoissonSig#, PkMeanCts#
Dim LowerL&, UpperL&, Pk#(), Cps#, scts#, SigmaPkCts#, MaxAbsDelt#, AbsDelt#
Dim SigmaMeanPkMeanCts#, SigmaSbmCts#

sumX = 0: sumX2 = 0: Nf = Nfields
ReDim Pk(1 To Nf)

For IntegrN = 1 To Nfields ' 10 integrations in Int(PkNum) seconds
  Pk(IntegrN) = sub10PkCts(IntegrN)
  sumX = sumX + Pk(IntegrN)
  sumX2 = sumX2 + Pk(IntegrN) ^ 2
Next IntegrN

Nf = Nfields
med = foAp.Median(sub10PkCts)

If med <= 100 Then
  PoissonLimits med, LowerL, UpperL
  MaxAbsDelt = 0: MaxDeltIndx = 0

  For IntegrN = 1 To Nfields
    PkI = sub10PkCts(IntegrN)

    If PkI < LowerL Or PkI > UpperL Then
      AbsDelt = Abs(PkI - med)
      If AbsDelt > MaxAbsDelt Then
        MaxAbsDelt = AbsDelt
        MaxDeltIndx = IntegrN
      End If
    End If

  Next IntegrN

  If MaxDeltIndx > 0 Then
    PkI = sub10PkCts(MaxDeltIndx)
    sumX = sumX - PkI
    sumX2 = sumX2 - PkI ^ 2
    Nf = Nf - 1
  End If

  PkMeanCts = sumX / Nf ' Mean of the Nf integrations
  SigmaPkCts = sqR((sumX2 - sumX ^ 2 / Nf) / (Nf - 1))
  PoissonSig = sqR(PkMeanCts)
  SigmaPkCts = fvMax(SigmaPkCts, PoissonSig)
Else
  TukeysBiweight sub10PkCts(), Nfields, PkMeanCts, 9, SigmaPkCts
  scts = sqR(PkMeanCts)
  SigmaPkCts = fvMax(SigmaPkCts, scts)
End If

SigmaMeanPkMeanCts = SigmaPkCts / sqR(Nf)
Cps = PkMeanCts * Nfields / IntegrTime
CorrPkTotCts = IntegrTime * fdDeadTimeCorrCPS(Cps) ' Total cts, dead-time corr

If PkMeanCts = 0 Then
  SigmaMeanPkTotCts = 0
Else
  SigmaMeanPkTotCts = SigmaMeanPkMeanCts * CorrPkTotCts / PkMeanCts
End If

TukeysBiweight sub10SbmCts(), Nfields, scts, 6, SigmaSbmCts
SbmTotCts = Nfields * scts
SigmaMeanSBMtotCts = sqR(SbmTotCts)
End Sub

Sub TukeysBiweight(X#(), ByVal N&, Mean#, Optional ByVal Tuning = 6, _
  Optional Sigma, Optional Err95)
' Calculates Tukey's biweight estimator of location & scale.
' Mean is a very robust estimator of "mean", Sigma is the robust estimator of
'   "sigma".  These estimators converge to the true mean & true sigma for
'   Gaussian distributions, but are very resistant to outliers.
' The lower the "Tuning" constant is, the more the tails of the distribution
'   are effectively "trimmed" (& the more robust the estimators are against
'   outliers), with the price that more "good" data is disregarded.  pts
'   that deviate from the "mean" greater that "Tuning" times the "standard
'   deviation" are assigned a weight of zero ('rejected').
' Err95 is the 95% conf-limit on Mean.  "Gaussian" is returned as -1 if
'   the distribution of x() appears to be normal (at 95%-conf. limit), 0 if  not.
' Adapted & inferred from Hoaglin, Mosteller, & Tukey, 1983, Understanding
'   Robust & Exploratory Data Analysis: John Wiley & Sons, pp. 341, 367,
'   376-378, 385-387, 423,& 425-427.

Dim TbiMatch As Boolean, SbiMatch As Boolean
Dim j%, Iter%
Dim w#, t#, Snsum#, Sdsum#, Tnsum#, Delta#, u#, Tuner#, MedianVal#
Dim U1#, U2#, U5#, U12#, Madd#, LastTbi#, LastSbi#, TbiDelt#, SbiDelt#

Const MaxIter = 100, Small = 1E-30
Const ZerTest = 0.0000000001, NonzerTest = 0.0000000001

MedianVal = iMedian(X()) ' Initial estimator of location is Median.
Mean = MedianVal
GetMAD X(), N, Mean, Madd, 0   ' Initial estimator of scale is MAD.
Sigma = fvMax(Madd, Small)

Do
  Iter = Iter + 1
  Tuner = Tuning * Sigma
  Snsum = 0: Sdsum = 0: Tnsum = 0

  For j = 1 To N
    Delta = X(j) - Mean

    If Abs(Delta) < Tuner Then
      u = Delta / Tuner
      U2 = u * u      ' U^2
      U1 = 1 - U2     ' 1-U^2
      U12 = U1 * U1   '(1-U^2)^2
      U5 = 1 - 5 * U2 ' 1-5U^2
      Snsum = Snsum + (Delta * U12) ^ 2
      Sdsum = Sdsum + U1 * U5
      Tnsum = Tnsum + u * U12
    End If

  Next j

  LastTbi = Mean: LastSbi = Sigma
  Sigma = sqR(N * Snsum) / Abs(Sdsum)
  If Sigma < Small Then Sigma = Small
  Mean = LastTbi + Tuner * Tnsum / Sdsum ' Newton-Raphson method
  TbiDelt = Abs(Mean - LastTbi)
  SbiDelt = Abs(Sigma - LastSbi)

  If Mean = 0 Then
    TbiMatch = (TbiDelt < ZerTest)
  Else
    TbiMatch = ((TbiDelt / Mean) < NonzerTest)
  End If

  If Sigma = 0 Then
    SbiMatch = (SbiDelt < ZerTest)
  Else
    SbiMatch = ((SbiDelt / Sigma) < NonzerTest)
  End If

Loop Until (TbiMatch And SbiMatch) Or Iter > MaxIter

If Sigma <= Small Then Sigma = 0
' t-approx. for near-Gaussian distr's; from Monte Carlo
'  simulations followed by Simplex fit (valid for Tuning=9).

Select Case N
  Case 2, 3: t = 47.2  ' really only for N=3
  Case 4:    t = 4.736
  Case Is >= 5
    w = N - 4.358
    t = 1.96 + 0.401 / sqR(w) + 1.17 / w + 0.0185 / w ^ 2
End Select

Err95 = t * Sigma / sqR(N)
End Sub

Public Function MSWD(ValuesAndErrors, Optional PercentErrs = False, _
                     Optional SigmaLevel = 1) As Variant
Dim tmp As Variant
tmp = "#NUM!"
On Error GoTo Done
tmp = wtdav(ValuesAndErrors, , PercentErrs, SigmaLevel)(3, 1)
Done: On Error GoTo 0
MSWD = tmp
End Function

Function fdMSWD#(ValuesAndErrors, Optional PercentErrs = False, _
                 Optional SigmaLevel = 1)
fdMSWD = MSWD(ValuesAndErrors, PercentErrs, SigmaLevel)
End Function

Public Function Probfit(ValuesAndErrors, Optional PercentErrs = False, _
                        Optional SigmaLevel = 1) As Variant
Dim tmp As Variant
tmp = "#NUM!"
On Error GoTo Done
tmp = wtdav(ValuesAndErrors, , PercentErrs, SigmaLevel)(5, 1)
Done: On Error GoTo 0
Probfit = tmp
End Function

Function fdProbfit#(ValuesAndErrors, Optional PercentErrs = False, _
                    Optional SigmaLevel = 1)
fdProbfit = Probfit(ValuesAndErrors, PercentErrs, SigmaLevel)
End Function

Public Function ExternalSigma(ValuesAndErrors, _
   Optional PercentErrs = False, Optional SigmaLevel = 1) As Variant
Dim tmp As Variant
tmp = "#NUM!"
On Error GoTo Done
tmp = wtdav(ValuesAndErrors, False, PercentErrs, SigmaLevel, False, True, 1)(3, 1)
Done: On Error GoTo 0
ExternalSigma = tmp
End Function

Function fdExternalSigma#(ValuesAndErrors, Optional PercentErrs = False, _
                          Optional SigmaLevel = 1)
fdExternalSigma = ExternalSigma(ValuesAndErrors, PercentErrs, SigmaLevel)
End Function
