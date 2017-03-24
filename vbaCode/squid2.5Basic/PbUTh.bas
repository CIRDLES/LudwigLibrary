Attribute VB_Name = "PbUTh"
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
Option Explicit
Option Base 1
' Module PnUTh - procedures for U-Th-Pb isotope systematics

Sub sqConcAge(Optional Automatic As Boolean = False)
' Calculate Concordia Age for either selected range, or for all
'  points with rejection until probability or MSWD is acceptable.
' Place the formatted results on the Grouped Samples worksheet.
Dim Bad As Boolean, b As Boolean, NoXY As Boolean, Manual As Boolean
Dim s$, tmp$
Dim i%, N%, n0%, Ccol%, df%, MaxRind%, Col, Area%, PtNum%
Dim Rw&
Dim Xbar#, Ybar#, ErrX#, ErrY#, SumsXY#, RhoXY#
Dim MSWD#, Prob#, SigmaXp#, SigmaYp#, Age#, cProb#, cMSWD#, AgeSigma#
Dim SigmaAge#, Emult#, MaxRsd#, MinCprob#
Dim X#(), Y#(), Xerr#(), Yerr#(), Pts#()
Dim tObj As Range, ciN As Range

If ActiveSheet.Type <> xlWorksheet Then Exit Sub
MinCprob = 0.1 ' fvMin prob for auto xy-equiv and concordance
plHdrRw = flHeaderRow(0)
If plHdrRw = 1 Then Exit Sub
NoUpdate
Manual = Not Automatic
If Automatic Then Set ciN = ActiveCell
FindStr "238/206*", , Ccol, plHdrRw, 1
If Ccol = 0 Then Exit Sub
N = 0

If Automatic Then
  If pdMinProb = 0 Then pdMinProb = 0.15
  If pdMinFract = 0 Then pdMinFract = 0.7
  Rw = plHdrRw

  Do
    Rw = Rw + 1
    If Rw > 999 Then Exit Sub
  Loop Until Cells(Rw, Ccol).Borders(xlEdgeBottom).LineStyle = xlContinuous

  frSr(1 + plHdrRw, Ccol, Rw, 3 + Ccol).Select
  Selection.Font.Strikethrough = False

Else
  With Selection ' Can select ANY cells in the radiogenic T-W range.

    For Area = 1 To .Areas.Count

      For Col = .Areas(Area).Column To _
                .Areas(Area).Column + .Areas(Area).Columns.Count - 1

        If Col < Ccol Or Col > (Ccol + 3) Then
          s$ = "You must select one or more rows from the "
          s$ = s$ & "4|radiogenic Tera-Wasserburg concordia columns"
          MsgBox fsVertToLF(s$), , pscSq
          Exit Sub
        End If

      Next Col

    Next Area

  End With
End If

foAp.Calculate
Do
  N = 0
  With Selection

    For Area = 1 To .Areas.Count

      For Rw = .Areas(Area).Row To _
               .Areas(Area).Row + .Areas(Area).Rows.Count - 1
        b = True

        For Col = Ccol To 3 + Ccol
          Set tObj = Cells(Rw, Col)
          If fbNoNum(tObj) Or tObj.Font.Strikethrough Then b = False: Exit For
          If tObj.Value = 0 Then b = False: Exit For
        Next Col

        If b Then
          N = 1 + N
          ReDim Preserve X(1 To N), Y(1 To N), Xerr(1 To N), Yerr(1 To N)
          X(N) = Cells(Rw, Ccol): Y(N) = Cells(Rw, 2 + Ccol)
          Xerr(N) = Cells(Rw, 1 + Ccol) / 100 * X(N)
          Yerr(N) = Cells(Rw, 3 + Ccol) / 100 * Y(N)
        End If
      Next Rw

    Next Area

  End With
  If N < 1 Then Exit Sub

  If n0 = 0 Then n0 = N
  ReDim Pts(1 To N, 1 To 5)

  For PtNum = 1 To N
    Pts(PtNum, 1) = X(PtNum): Pts(PtNum, 2) = Xerr(PtNum)
    Pts(PtNum, 3) = Y(PtNum): Pts(PtNum, 4) = Yerr(PtNum)
    Pts(PtNum, 5) = 0
  Next PtNum

  WtdXYmean Pts, (N), Xbar, Ybar, SumsXY, ErrX, ErrY, RhoXY, Bad
  df = 2 * N - 2

  If df <= 0 Then
    MSWD = 0: Prob = 1
  Else
    MSWD = SumsXY / df: Prob = ChiSquare(MSWD, (df))
  End If

If Manual Or (Prob >= MinCprob Or MSWD <= 1.4) Then Exit Do
  If N = 1 Then ciN.Select: Exit Sub

  If ((N - 1) / n0) < pdMinFract Or Bad Then
    Selection.Font.Strikethrough = False
    ciN.Select: NoXY = True: Exit Do
  End If

  MaxRsd = 0 ' Trim point with highest wtd-resid

  For PtNum = 1 To N

    If Yf.WtdResid(PtNum) = 0 Or Yf.WtdResid(PtNum) > MaxRsd Then
      With Selection ' Find which row this is
        Col = .Column

        For Rw = .Row To .Row + .Rows.Count - 1

          If IsNumeric(Cells(Rw, Col)) And IsNumeric(Cells(Rw, 2 + Col)) Then

            If Cells(Rw, Col) = X(PtNum) And Cells(Rw, 2 + Col) = Y(PtNum) Then

              ' Don't reject twice
              If Not frSr(Rw, Col, Rw, 3 + Col).Font.Strikethrough Then
                MaxRsd = Yf.WtdResid(PtNum): MaxRind = PtNum
              End If

              Exit For
            Else
              MaxRsd = Yf.WtdResid(PtNum): MaxRind = PtNum
              Exit For
            End If

          End If

        Next Rw

      End With
    End If

  Next PtNum

  b = False ' Find the new rejected point
  With Selection
    Col = .Column

    For Rw = .Row To .Row + .Rows.Count - 1

      If IsNumeric(Cells(Rw, Col)) And IsNumeric(Cells(Rw, 2 + Col)) Then

        If Cells(Rw, Col) = X(MaxRind) And Cells(Rw, 2 + Col) = Y(MaxRind) Then

          If Not frSr(Rw, Col, Rw, 3 + Col).Font.Strikethrough Then
            frSr(Rw, Col, Rw, 3 + Col).Font.Strikethrough = True: b = True
            Exit For
          End If

        End If

      End If

    Next Rw

  End With
  If Not b Then Exit Do
Loop

Cells(flEndRow(Ccol), Ccol).Select
With ActiveCell
  If InStr(Cells(.Row + 2, .Column - 1).Text, "Concordia Age") Then
    Rw = flVisRow(.Row + 2)
  Else
    Rw = flVisRow(.Row + 2)
  End If
End With

Rw = Rw + 2
Fonts Rw, Ccol, , , vbBlue, True, , 11
If Bad And Manual Then MsgBox "Could not calculate the X-Y wtd mean", , pscSq: Exit Sub

If Prob < MinCprob And MSWD > 1.4 And Not NoXY Then
  If Automatic Then Exit Sub
  If Prob < 0.00001 Then tmp$ = "~zero)" Else tmp$ = fsS(Drnd(Prob, 2)) & ")"
  MsgBox "Probability of X-Y equivalence is too low (" & tmp$, , pscSq
Else
  If Prob < 0.15 And Not NoXY Then
    Emult = sqR(MSWD) * foAp.TInv(1 - 0.683, df) ' Students-T @ 68.3% conf lim
    ErrX = ErrX * Emult: ErrY = ErrY * Emult
  End If
  SigmaXp = 100 * ErrX / Xbar: SigmaYp = 100 * ErrY / Ybar

  If N > 1 Then

    If NoXY Then
      s$ = "No coherent Concordia group"
    Else
      Fonts Rw, Ccol, , , , , xlLeft, 11, , , _
        "X-Y wtd mean (68%-conf. errs), incl. error from Standard:"
      Rw = Rw + 1
      Cells(Rw, Ccol) = Xbar
      Cells(Rw, 1 + Ccol) = "=SQRT(" & StR(Drnd(SigmaXp, 3)) & _
                  "^2 +(" & pscStdShtNot & "WtdMeanAPerr1)^2)"
      Cells(Rw, 2 + Ccol) = Ybar
      Cells(Rw, 3 + Ccol) = Drnd(SigmaYp, 3)
      Fonts Rw, Ccol, Rw, 3 + Ccol, vbBlack, False, xlRight, 11
      IntClr peLightGray, Rw, Ccol, Rw, 3 + Ccol
      RangeNumFor pscZd2 & String$(-(Xbar < 10) - _
        (Xbar < 1) - (Xbar < 0.1), pscZq), Rw, Ccol
      RangeNumFor "0.0000" & String$(-(Ybar < 0.1), pscZq), Rw, Ccol + 2
      RangeNumFor pscZd2, Rw, Ccol + 1: RangeNumFor "0.000", Rw, Ccol + 3
      If Automatic And N = n0 Then s$ = "on all" Else s = "on"
      s$ = s$ & StR(N) & " points"
      If Automatic And N < n0 Then s$ = s$ & " of" & StR(n0)
      If Automatic And N < n0 Then s$ = s$ & " (arbitrary rejections!)"
      s = s & "."
    End If

    Fonts Rw + 1, Ccol - 1 - 1 * Not NoXY, , , vbBlue, True, xlLeft, , , , s

    If Not NoXY Then
      ReDim Pts(1 To 1, 1 To 5)
      Pts(1, 1) = Xbar: Pts(1, 3) = Ybar
      Pts(1, 2) = Cells(Rw, 1 + Ccol).Value / 100 * Xbar
      Pts(1, 4) = ErrY: Pts(1, 5) = 0
      ReDim inpdat(1 To 1, 1 To 5)

      For i = 1 To 5: inpdat(1, i) = Pts(1, i): Next i

      For i = Ccol To Ccol + 3
        Cells(Rw, i).NumberFormat = pscGen
        Box Rw, i
      Next i

      Rw = flVisRow(2 + Rw)
      s = "Probability of equivalence = 0" & fsS(Drnd(Prob, 2)) & _
          "  (mswd = " & fsS(Drnd(MSWD, 2)) & ")"
      Fonts Rw, Ccol, , , vbBlue, True, xlLeft, , , , s
      Rw = flVisRow(1 + Rw)
      GetConsts ' Set & reset ISOPLOT variables
      Inverse = True
      Isoplot3.ConcordiaAges Pts(), 1, Bad, Age, SigmaAge, cMSWD, cProb

      If Bad Or cProb < 0.05 Then
        s = "DISCORDANT  (probability of concordance =" & StR(Drnd(cProb, 2)) & ")"
        Fonts Rw, Ccol, , , vbRed, True, xlLeft, , , , s
      Else
        With foAp
          tmp$ = "Concordia Age = " & fsS(.Fixed(Age, 1)) & _
            " " & Chr(177) & StR(.Fixed(1.96 * SigmaAge, 1)) & "   (95%-conf."
          End With
        If N = 1 Then tmp$ = tmp$ & ", incl. error from Standard"
        Fonts Rw, Ccol, , , vbRed, True, xlLeft, , , , tmp
        Rw = flVisRow(1 + Rw)
        tmp = "Probability of concordance = 0" & fsS(foAp.Fixed(cProb, 3))
        Fonts Rw, Ccol, , , vbRed, True, xlLeft, , , , tmp, pscGen
      End If

    End If
  End If
End If

With ActiveWindow ' Put last row of the new info at bottom of screen
  .ScrollRow = 1 + .SplitRow

  Do
    .SmallScroll 1
  Loop Until frSr(.ScrollRow, , Rw + 4).Height < .Height

End With
ClearObj ciN, tObj
End Sub

Sub CreateStdConstBoxes()
' Copy & populate the three Std Const boxes from foUser,
'  paste well out of the way from the data columns (for now).
Dim Pb76rad$, s1$
Dim Col1%, Col2%, i%, rw1%, CPbSpecType%
Dim t#, rc#(1 To 3)
Dim cP As Range

rw1 = plHdrRw + 1
Col1 = piSlastCol + 6
CopySbox "StdCommPb", rw1, Col1, Col2
Set cP = [stdcommpb]
Box cP, , , , RGB(212, 212, 212), , , , , True ' 09/10/08  change from dbl to single thick line
Fonts rw1:=cP, Size:=11, FontName:="Arial Narrow"
HA xlRight, frSr(cP.Row + 1, cP.Column + 1, cP.Row + 3)
HA xlRight, frSr(cP.Row + 1, cP.Column + 4, cP.Row + 3)
Fonts cP.Row + 1, cP.Column + 2, cP.Row + 3, , , , xlLeft, 9
Fonts cP.Row + 1, cP.Column + 5, cP.Row + 3, , , , xlLeft, 9
CPbSpecType = foUser("CPbSpecType")

If foUser("cPb64") = "" Then foUser("cPb64") = 18.3
If foUser("cPb76") = "" Then foUser("cPb76") = 0.854
If foUser("cPb86") = "" Then foUser("cPb86") = 2.09

pdComm64 = foUser("cpb64")
pdComm76 = foUser("cpb76")
pdComm86 = foUser("cpb86")
NameRange psaC64(1), -1, cP(2, 3), , pdComm64, pscZd2, 0
NameRange psaC76(1), -1, cP(3, 3), , pdComm76, pscZd3, 0
NameRange psaC86(1), -1, cP(4, 3), , pdComm86, pscZd2, 0
NameRange psaC74(1), -1, cP(2, 6), , pdComm64 * pdComm76, pscZd2, 0
NameRange psaC84(1), -1, cP(3, 6), , pdComm64 * pdComm86, pscZd2, 0

If CPbSpecType = 1 Or CPbSpecType = 2 Then '"DefSKageType") > 0 Then ' 09/12/09 -- mod
  t = Choose(CPbSpecType, pdAgeStdAge, foUser("CPbSKage"))           ' 09/12/09 -- mod
  If t < -2000 Or t > 6000 Then t = 0
  For i = 1 To 3: rc(i) = Drnd(SingleStagePbR(t, i - 1), 4): Next
  Range(psaC64(1)) = Drnd(rc(1), 4)
  Range(psaC74(1)) = Drnd(rc(2), 4)
  Range(psaC84(1)) = Drnd(rc(3), 4)
  Range(psaC76(1)) = Drnd(rc(2) / rc(1), 4)
  Range(psaC86(1)) = Drnd(rc(3) / rc(1), 4)
End If

rw1 = rw1 + 5
If (pbU And pbHasUconc) Or (pbTh And pbHasThConc) Then
  CopySbox "StdConc", rw1, Col1, Col2: Set cP = [StdConc]
  NameRange "ConcStdPpm", -1, cP(1, 4), , pdConcStdPpm, 0, 0
  NameRange "ConcStdConst", -1, cP(2, 4), , Drnd(pdMeanParentEleA, 5), pscGen, 0
  rw1 = rw1 + 3
End If

CopySbox "UPbStdAgesRatios", rw1, Col1, Col2
Set cP = [upbstdagesratios]
Box cP, , , , RGB(212, 212, 212), , , , , True ' 09/10/08  change from dbl to single thick line
Fonts cP, , , , , , , 11
HA xlRight, frSr(cP.Row, cP.Column + 2, cP.Row + 2)
HA xlLeft, frSr(cP.Row, cP.Column + 3, cP.Row + 2)
NameRange "StdAgeUPb", -1, cP(1, 4), , pdAgeStdAge, pscGen, 0, _
  "Age in Ma of the U-Pb age-Standard"

NameRange "StdUPbRatio", -1, cP(2, 4), , "=" & pscEx8 & "StdAgeUPb" & ")-1", _
  pscDd4, 0, psaPDeleRat(piU1Th2) & " of the U/Pb age standard"

If pdStdPbPbAge > 0 Then
  Pb76rad = fsS(pdStdPbPbAge)
Else
  Pb76rad = "AgeStdAge"
End If

NameRange "Std_76", -1, cP(3, 4), , "=Pb76(" & Pb76rad & ")", _
  pscDd4, 0, "207Pb*/206Pb* of the age standard"
rw1 = rw1 + 4

CopySbox "ThPbStdAgesRatios", rw1, Col1, Col2
Set cP = [ThPbStdAgesRatios]
Box cP, , , , RGB(212, 212, 212), , , , , True ' 09/10/08  change from dbl to single thick line
Fonts cP, , , , , , , 11
HA xlRight, frSr(cP.Row, cP.Column + 2, cP.Row + 3)
HA xlLeft, frSr(cP.Row, cP.Column + 3, cP.Row + 3)
NameRange "StdAgeThPb", -1, cP(1, 4), , pdStdAgeThPb, pscGen, 0, _
    "Age(ma) of the Th-Pb age-Standard"

If pbU Then
  [stdageupb].Name = "AgeStdAge"
Else
  [stdagethpb].Name = "AgeStdAge"
End If

s1 = IIf(pbU, "=AgeStdAge", pdAgeStdAge)
pdStdAgeThPb = pdAgeStdAge
NameRange "StdAgeThPb", -1, cP(1, 4), , s1, pscGen, 0, _
  "Age (ma) of the Th-Pb age-Standard"
NameRange "StdThPbRatio", -1, cP(2, 4), , "=" & pscEx2 & "StdAgeThPb" & ")-1", _
 pscDd4, 0, "Th/Pb of the U-Pb age standard"


If pbTh And piNumDauPar = 1 Then
  NameRange "Std_76", -1, cP(3, 4), , "=Pb76(" & "StdAgeThPb" & ")", _
    pscDd4, 0, "207Pb*/206Pb* of the Th/Pb age standard"
  NameRange "StdRad86fact", -1, cP(4, 4), , _
    "=(" & pscEx2 & "StdAgeThPb)-1)/(" & pscEx8 & "StdAgeThPb)-1)", _
    pscDd4, 0, "Std 208Pb*/206Pb* x pbStd 238U/232Th" ' 09/06/24 was "stdageth"
Else
  NameRange "", -1, cP(3, 4), , "=Std_76", pscDd4, 0, _
            "207Pb*/206Pb* of the U/Pb age standard"
  NameRange "StdRad86fact", -1, cP(4, 4), , _
    "=(" & pscEx2 & "StdAgeThPb)-1)/(" & pscEx8 & "StdAgeUPb)-1)", _
    pscDd4, 0, "Std 208Pb*/206Pb* x Std 238U/232Th"
End If

ColWidth picAuto, Col1 + 3
Set cP = Nothing
End Sub

Sub ThUfromFormula(ByVal Std As Boolean, ByVal Rw&)
' Calculate the spot's 232Th/238U using the 232Th/238U Task Equation from
'  the "U-Pb Special" panel (U-Pb geochron only).
Dim Equa$, p%, c%, MeanV#, MeanVferr#

p = -Std:  c = piaTh2U8col(p)
 piSpotOutputCol = piaEqCol(Std, -3)
With puTask
  Equa = .saEqns(-3)
  EqnInterp Equa, -3, MeanV, MeanVferr, 1, 1
  CFs Rw, c, fsS(MeanV)

  If (Not Std Or piaTh2U8ecol(1) > 0) And MeanV <> pdcErrVal _
      And MeanVferr <> pdcErrVal Then
    CFs Rw, c + 1, fsS(100 * MeanVferr)
  End If

End With
End Sub

Sub ThUfromA1A2(ByVal Std As Boolean, ByVal Rw&, _
  Optional ByVal Only1 As Boolean = False)
' Caclulate 232/238 by ratioing the 206/238 calibration-constant from
'  the 208/232 calibration constant (for Tasks with U-Pb Special equations
'  specified as "calculate 208Pb/232Th directly, and place the results
'  on the Std or Sample output-data worksheet (U-Pb geochron only).
Dim t1$, t2$, t3$, Th2U8col%, Th2U8ecol%, p%, ParentIs238%
Dim Exp238$, Exp232$

' If the two wtdMeanA formulae & values have not yet been calculated,
'   must pass first time & put dummy valiues of 1.  Put in real formulae later
p = -Std
Th2U8col = piaTh2U8col(p)
Th2U8ecol = piaTh2U8ecol(p)
t2 = ""

If Only1 Then
  t1 = "1"
ElseIf Th2U8col > 0 Then

  If Std And piStdRadPb86col > 0 Then
    ParentIs238 = (puTask.iParentIso = 238)
    Exp238 = "(exp(" & pscLm8 & "* sage(" & fsS(2 + ParentIs238) & ") )-1)"
    Exp232 = "(exp(" & pscLm2 & "* sage(" & fsS(1 - ParentIs238) & ") )-1)"
    t1 = "= StdRadPb86 *" & Exp238 & "/" & Exp232
  ElseIf Not Std Then
    t1 = "= Pb86 * Pb6U8_tot / Pb8Th2_tot "
  End If

  ' 09/07/09  correct eqns (ie with Pb86%err) put in place.
  If Th2U8ecol > 0 Then

    If Std Then
      t3 = " StdRadPb86e "

      If fiColNum(t3) > 0 Then
        t2 = "=sqrt(" & t3 & " ^2+ saecol(1) ^2+ saecol(2) ^2)"
      End If
    ElseIf piStdCorrType < 2 Then
      t2 = "=sqrt( Pb86e ^2+ Pb6U8_tote ^2+ Pb8Th2_tote ^2)"
    End If

  End If

End If

PlaceFormulae t1, Rw, Th2U8col
If t2 <> "" Then PlaceFormulae t2, Rw, Th2U8ecol
End Sub

Sub Tot68_82_fromA(ByVal Rw&) ' 232/238 & tot 206/238-208/232 from A(6/8)-A(8/2)
' SAMPLE SPOTS ONLY.
' For Tasks with 206/238 as the main parent/daughter, produces tot206/238,
'  total 238/206 and tot208/232; for 208/232 Tasjs, produces total 208/232,
'  total 206/238.  Put results on output-data worksheet.
' U-Pb geochron Tasks only.
Dim t1$, t2$, t3$, j$, m%, DpNum%, w%

For DpNum = 1 To piNumDauPar
  If (pbU And DpNum = 1) Or (pbTh And DpNum = 2) Then w = 1 Else w = 2
  m = Choose(w, piPb6U8_totCol, piPb8Th2_totCol)
  j = fsS(w)
  t1 = pscStdShtNot & "WtdMeanA" & j
  t2 = "=" & pscStdShtNot & "WtdMeanAperr" & j
  t3 = Choose(w, "StdUPbRatio", "StdThPbRatio")
  t3 = pscStdShtNot & t3
  ' 206t/238 or 208t/232
  PlaceFormulae "=" & fsCP("A", w) & "/" & t1 & "*" & t3, Rw, m
  PlaceFormulae "=SQRT(" & fsCP("Ae", w) & "^2+(" & t2 & "/2)^2)", Rw, m + 1
Next DpNum
End Sub

Sub SecondaryParentPpmFromThU(ByVal Std As Boolean, ByVal Rw&)
' Calculate secondary parent (Th or U) concentration from the primary
'  parent (U or Th) plus 232Th/238U.  U-Pb geochron Tasks only.
' (232/238 and ppm U or Th must already be calculated)

Dim tB As Boolean, t1$, pp$, Up$, tP$, tu$, p%, c%

p = -Std:  pp = "(" & fsS(p) & ")"
c = IIf(pbUconcStd, piaPpmThcol(p), piaPpmUcol(p))

With puTask
  If piaTh2U8col(p) > 0 And piaPpmUcol(p) > 0 And piaPpmThcol(p) > 0 Then
    Up = "ppmu" & pp
    tP = "ppmth" & pp
    tu = "th2u8" & pp

    If pbUconcStd Then
      t1 = Up & " * " & tu & " /"
    ElseIf pbThConcStd Then ' oops, should calc th232/u238 first
      t1 = tP & " / " & tu & " *"
    End If

    PlaceFormulae "= " & t1 & "1.033", Rw, c
  End If
End With
End Sub

Sub StdElePpm(ByVal Std As Boolean, ByVal SpotRow&)
' Calculate primary parent (U or Th) concentration
' U/Pb geochron Tasks only.
Dim p%, c%, v#

p = -Std: c = IIf(pbUconcStd, piaPpmUcol(p), piaPpmThcol(p))
If c = 0 Then Exit Sub
If pdMeanParentEleA > 0 And pdMeanParentEleA <> pdcErrVal Then
  piSpotOutputCol = piaEqCol(Std, piLwrIndx)
  EqnInterp puTask.saEqns(piLwrIndx), piLwrIndx, v, 0, 1, 0
  If v = pdcErrVal Then Exit Sub
  v = v / pdMeanParentEleA * pdConcStdPpm
End If
If v > 0 Then CFs SpotRow, c, fsS(v)
End Sub

Sub PlaceRawRatios(ByVal DatRw&, rr#(), rrFerr#())
' Fill the raw isotope-ratio columns (for the output-data sheet)
'  and XML-specific columns. U-Pb geochron Tasks only.
Dim PkNum%, Co%, DpNum%, Rw&

If piBkrdPkOrder Then CF DatRw, piBkrdCtsCol, pdBkrdCPS
If pi204PkOrder Then CF DatRw, piPb204ctsCol, pdTotCps204
If pi206PkOrder Then CF DatRw, piPb206ctsCol, pdTotCps206
On Error GoTo 0

With puTask
  If pbXMLfile Then

    Rw = plaSpotNameRowsCond(piSpotNum) + 4
    Co = picDatCol + 5 * .iNpeaks + 2
    With phCondensedSht
      CF DatRw, piStageXcol, .Cells(Rw, Co)
      CF DatRw, piStageYcol, .Cells(Rw, Co + 1)
      CF DatRw, piStageZcol, .Cells(Rw, Co + 2)
      CF DatRw, piQt1yCol, .Cells(Rw, Co + 3)
      CF DatRw, piQt1Zcol, .Cells(Rw, Co + 4)
      CF DatRw, piPrimaryBeamCol, .Cells(Rw, Co + 5)
    End With
  End If

  For PkNum = 1 To .iNpeaks
    If .baCPScol(PkNum) And piaCPScol(PkNum) > 0 Then
      CF DatRw, piaCPScol(PkNum), pdaTotCps(PkNum)
    End If
  Next PkNum

End With

If piPb46col Then
  CF DatRw, piPb46col, IIf(rr(pi46ratOrder) = pdcErrVal, _
     pdcTiny, rr(pi46ratOrder))
  CF DatRw, piPb46eCol, IIf(rrFerr(pi46ratOrder) = pdcErrVal, _
     1000000000#, rrFerr(pi46ratOrder)), -1
End If

If piPb76col Then
  CF DatRw, piPb76col, rr(pi76ratOrder)
  CF DatRw, piPb76eCol, rrFerr(pi76ratOrder), -1
End If

If piPb86col Then
  CF DatRw, piPb86col, rr(pi86ratOrder)
  CF DatRw, piPb86eCol, rrFerr(pi86ratOrder), -1
End If

Co = fvMax(piPb46col, piPb76col, piPb86col)

For PkNum = 1 To puTask.iNrats

  If piaIsoRatOrder(PkNum) > 0 And piaIsoRatCol(PkNum) > Co Then
    CF DatRw, piaIsoRatCol(PkNum), rr(piaIsoRatOrder(PkNum))
    CF DatRw, piaIsoRatEcol(PkNum), rrFerr(piaIsoRatOrder(PkNum)), -1
   End If

Next PkNum
End Sub

Function fsCellAddr$(ByVal Row&, ByVal Col%)
' Return the A1-style address of a cell.
fsCellAddr = Cells(Row, Col).Address(0, 1)
End Function

Sub OverCountColumns(ByVal Row&) ' Create the Standard-Worksheet 204-overcount columns
' Alpha(true) = [Beta0-Phi(rad, Std,true)]/[Phi(meas, bkrd-corr) - Phi(rad, Std,true)]
' Where phi = either 207/206Pb or 208/206Pb (use Gamma0 instead of Beta0 if 208/206)
' 204overcts = [206cps(bkrd uncorr, meas) - BkrdCps]/Alpha(true) + BkrdCps - 204cps(bkrd uncorr, meas)

Dim t1$, t2$, s$, k%, m%

If Not pbHasU Or pi204PkOrder = 0 Or piBkrdPkOrder = 0 Or piPb46col = 0 _
  Or Not foUser("ShowOverCtCols") Then Exit Sub

For k = 1 To 2
  m = k + 6: s = fsS(m) & ") "

  If piaOverCts4Col(m) > 0 Then

    If k = 1 And piPb76col > 0 Then
      t1 = "=( Pb76 -Std_76)/(sComm1_74-Std_76*sComm1_64)"
      t2 = "=abs( Pb76ecol * Pb76 /( Pb76 -Std_76))"
    ElseIf piPb86col > 0 Then
      t1 = "=( Pb86 -StdRad86fact* th2u8(1) )/" & _
           "(sComm1_84-StdRad86fact* th2u8col(1) *sComm1_64)"
    End If

    If piaOverCts46Col(m) > 0 Then
      PlaceFormulae t1, Row, piaOverCts46Col(m)
      If k = 1 Then PlaceFormulae t2, Row, piaOverCts46eCol(7)
      t1 = "= Pb204cts - BkrdCts - overcts46(" & s & "*( Pb206Cts - BkrdCts )"
      PlaceFormulae t1, Row, piaOverCts4Col(m)
    End If
  End If

  If piacorrAdeltCol(m) Then
    t1 = "(1-scomm1_64* Pb46 )"
    t2 = "(1-scomm1_64* overcts46(" & s & " )"
    PlaceFormulae "=100*(" & t1 & "/" & t2 & "-1)", Row, piacorrAdeltCol(m)
  End If

Next k

foAp.Calculate
End Sub

Sub OverCtMeans(ByVal Lrow&)

Dim t0$, t1$, t2$, t3$, t4$, Alpha$, DeltaAlpha$, Beta$, DeltaBeta$
Dim k%, c%, MinC%
Dim bw As Range

If pbU Then

  For k = 0 To 1
    c = IIf(k, piaOverCts4Col(8), piaOverCts4Col(7))
    If c Then
      If MinC = 0 Then MinC = c
    End If
  Next k

End If

If MinC = 0 Then MinC = 2 + fvMax(piaSacol(1), piaSacol(2))

If piaAgePb76_4Col(1) > 0 Then
  t0 = "=( Pb76 / Pb46 -" & psaC74(1) & ")/(1/ Pb46 -" & psaC64(1) & ")"
  PlaceFormulae t0, plaFirstDatRw(1), piStdPb76_4Col, plaLastDatRw(1)
  Alpha = "1/ Pb46col "
  Beta = " Pb76col / Pb46col "
  DeltaAlpha = "(" & Alpha & "-" & psaC64(1) & ")" ' Alpha-Alpha0
  DeltaBeta = "(" & Beta & "-" & psaC74(1) & ")"   ' Beta-Beta0
  t1 = "(( Pb76 * Pb76e )^2+( Pb46 *( StdPb76_4 *" & psaC64(1) & _
       "-" & psaC74(1) & ")* Pb46e )^2)"
  t2 = "( Pb76 - Pb46 *" & psaC74(1) & ")^2"
  t3 = "=sqrt(" & t1 & "/" & t2 & ")"
  PlaceFormulae t3, plaFirstDatRw(1), piStdPb76_4eCol, plaLastDatRw(1)
  t4 = "=AgePb76( StdPb76_4 )"
  PlaceFormulae t4, plaFirstDatRw(1), piaAgePb76_4Col(1), plaLastDatRw(1)
  t3 = "=AgeerPb76( StdPb76_4 , StdPb76_4e /100* StdPb76_4 )"
  PlaceFormulae t3, plaFirstDatRw(1), piaAgePb76_4eCol(1), plaLastDatRw(1)
  RangeNumFor fsOptNumFor(plaFirstDatRw(1), Lrow, piaAgePb76_4Col(1)), _
              plaFirstDatRw(1), piaAgePb76_4Col(1), Lrow, piaAgePb76_4eCol(1)
  RangeNumFor fsOptNumFor(plaFirstDatRw(1), Lrow, piStdPb76_4Col), _
              plaFirstDatRw(1), piStdPb76_4Col, Lrow, piStdPb76_4Col + 1
  ColWidth picAuto, piaAgePb76_4Col(1), piStdPb76_4Col + 1
End If

If foUser("ShowOverCtCols") Or piaAgePb76_4Col(1) > 0 Then
  If piaOverCts4Col(7) > 0 Or piaOverCts4Col(8) > 0 Or piaAgePb76_4Col(1) > 0 Then
    Fonts 1 + Lrow, MinC, 3 + Lrow, piSlastCol + 1, RGB(0, 0, 192), True, xlRight
    HA xlLeft, 1 + Lrow, piSlastCol + 1, 3 + Lrow

    For k = 1 To 3 Step 2
      With Cells(k + Lrow, piSlastCol + 1)
        If k = 1 Then .Formula = "Biweight Mean" Else .Formula = "95% conf uncertainty"
      End With
    Next k

    frSr(1 + Lrow, piSlastCol + 1, 3 + Lrow).Name = "RbAv"

    For k = IIf(foUser("ShowOverCtCols"), 1, 5) To 5

      Select Case k
        Case 1: c = piaOverCts4Col(7)
        Case 2: c = piaOverCts4Col(8)
        Case 3: c = piacorrAdeltCol(7)
        Case 4: c = piacorrAdeltCol(8)
        Case 5: c = piaAgePb76_4Col(1)
      End Select

      If c Then
        Set bw = frSr(1 + Lrow, c, 3 + Lrow)
        bw.FormulaArray = "=Biweight(" & frSr(plaFirstDatRw(1), c, Lrow).Address & ",9)"

        If k < 3 Then
          t1 = pscStdShtNa & "!" & "Pb204OverCts" & fsS(6 + k) & "corr"
          AddName t1, True, 1 + Lrow, c
          AddName t1 & "Er", True, 3 + Lrow, c
          t1 = "Robust avg 204 overcts assuming  " & IIf(k = 1, pscR6875, pscR6882)
        ElseIf k = 3 Or k = 4 Then
          t1 = "OverCtsDeltaP" & fsS(4 + k) & "corr"
          AddName t1, True, 1 + Lrow, c
          AddName t1 & "Er", True, 2 + Lrow, c
          t1 = "Robust avg of diff. between 20" & IIf(k = 3, "7", "8") & _
            "-corr. and 204-corr. calibr. const."
        Else
          t1 = "Robust average of 204-corrected 207/206 age"
        End If

        On Error Resume Next
        Note 1 + Lrow, c, t1
        t1 = "95%-conf. error in above"
        If k = 3 Or k = 4 Then t1 = t1 & " difference"
        Note 3 + Lrow, c, t1
        frSr(1 + Lrow, c).NumberFormat = Cells(Lrow, c).NumberFormat
        Cells(3 + Lrow, c).NumberFormat = fsInQ(Chr(177)) & IIf(k = 5, pscZq, pscZd2)
      End If

    Next k

    RangeNumFor pscZq, 2 + Lrow, MinC, , piSlastCol
  End If

  MinC = fvMax(piaSageEcol(piNumDauPar) + 3, MinC)

  ColWidth picAuto, MinC, piSlastCol
  ClearObj bw
End If

End Sub

Sub GetRatios(Ratios#(), RatioFractErrs#(), ConcOnly As Boolean, BadSbm%())

Dim tB As Boolean
Dim Rct%, j%, Rnum%(), Rdenom%(), SpottCt%, iC%

ReDim RatioFractErrs(1 To 99), Ratios(1 To 99), Rnum(1 To 99), Rdenom(1 To 99)

On Error GoTo 0
Rct = 0
SpottCt = IIf(pbStd, piaSpotCt(1), piaSpotCt(0))

If Not ConcOnly Then

  If pi206PkOrder > 0 Then

    If pi204PkOrder > 0 Then
      Rct = 1 + Rct: pi46ratOrder = Rct
      Rnum(Rct) = pi204PkOrder
      Rdenom(Rct) = pi206PkOrder
      InterpRat pi204PkOrder, pi206PkOrder, Ratios(Rct), _
                RatioFractErrs(Rct), BadSbm(), pbStd, tB
      If pbStd Then
        pbStdRej(SpottCt, Rct) = tB
      Else
        pbSamRej(SpottCt, Rct) = tB
      End If
      pdNetCps204 = pdNetCps206 * IIf(Ratios(Rct) = pdcErrVal, 0, Ratios(Rct))
      pdTotCps204 = pdNetCps204 + pdBkrdCPS
    End If

    If pi207PkOrder > 0 Then
      Rct = 1 + Rct: pi76ratOrder = Rct
      Rnum(Rct) = pi207PkOrder
      Rdenom(Rct) = pi206PkOrder
      InterpRat pi207PkOrder, pi206PkOrder, Ratios(Rct), _
                RatioFractErrs(Rct), BadSbm(), pbStd, tB  ' 207/206
      If pbStd Then pbStdRej(SpottCt, Rct) = tB Else pbSamRej(SpottCt, Rct) = tB
    End If

    If pi208PkOrder > 0 Then
      Rct = 1 + Rct: pi86ratOrder = Rct
      Rnum(Rct) = pi208PkOrder: Rdenom(Rct) = pi206PkOrder
      InterpRat pi208PkOrder, pi206PkOrder, Ratios(Rct), _
                RatioFractErrs(Rct), BadSbm(), pbStd, tB  ' 208/206
      If pbStd Then pbStdRej(SpottCt, Rct) = tB Else pbSamRej(SpottCt, Rct) = tB
    End If

  End If

  For j = 1 To puTask.iNrats
    iC = piaIsoRatCol(j)

    If iC > 0 And iC <> piPb46col And iC <> piPb76col _
        And iC <> piPb86col And Rct < puTask.iNrats Then
      Rct = 1 + Rct
      piaIsoRatOrder(j) = Rct
      Rnum(Rct) = piaIsoRatsPkOrd(j, 1)
      Rdenom(Rct) = piaIsoRatsPkOrd(j, 2)
      InterpRat Rnum(Rct), Rdenom(Rct), Ratios(Rct), _
                RatioFractErrs(Rct), BadSbm(), pbStd, tB
      If pbStd Then
        pbStdRej(SpottCt, Rct) = tB
      Else
        pbSamRej(SpottCt, Rct) = tB
      End If
    End If

  Next j

End If
ReDim Preserve Rnum(1 To Rct), Rdenom(1 To Rct), Ratios(1 To Rct)
ReDim Preserve RatioFractErrs(1 To Rct)
End Sub

Sub WtdMeanAcalc(BadSbm%(), Adrift#(), AdriftErr#())
' 09/06/10 -- Eliminate the UseLowess variable & replace with pbCanDriftCorr,
'             since the latter now includes the min #spots constraint.
Dim Da As Boolean, NoReject As Boolean
Dim ErNm$, s$, Ele$, t1$, t2$, t3$, t4$, t5$, SSC$, ssCe$, ssCa$, psMswd$
Dim ErN%, Npts%, LargeErRegN%, Nrej%, NtotRej%, Col1%, Acol%, Aecol%
Dim DpNum%, Ch%, i%, j%, k%, c%, Ndp%, LastExtboxR%, NLowess%
Dim LargeErRej%()

Dim rw1&, rwn&, RwS&, RejClr&, r&, Rw&, tmp&, h&, fClr&, wHA1&, wHA2&, Bclr&
Dim L!, t!, Rt!
Dim Prob#, Lambda#, MeanAdrift#, maErr95#, WtdMeanErr#, MedianEr#
Dim Nmadd#, WtdMean#, IntSigmaMean#, ExtSigmaMean#, ExtPtSigma#, MSWD#
Dim ErrVals#(), ad#(), adE#()

Dim AllDeltaP As Range, Sig2AllDeltaP As Range
Dim extBox As Range, UPbConst As Range
Dim Arange As Range, AerRange As Range, TmpR As Range
Dim Hours As Range, AllHrsDeltaP As Range, LowessRange As Range
Dim AvgLine As Range, LowessHours As Range
Dim w(2) As Range, w0(2) As Range, Rr1 As Range, Rr2 As Range
Dim Sig2LowessDeltaP As Range, ExtPerr As Range, LowessDeltaP As Range
Dim SerCol As Series, Ara As Range, LowSchart As Chart, Up As Shape, Down As Shape
Dim StdW As Window, ChtObj As ChartObject, LowS As Lowess
Dim Rejected() As Variant, SC As Series, wW As Variant


HideColumns True

rw1 = plaFirstDatRw(1)
rwn = plaLastDatRw(1)
With frSr(2 + rwn, 1, , piSlastCol)
  .Font.Color = vbWhite: .RowHeight = 1
End With
Set StdW = ActiveWindow
NoReject = (foUser("NoUPbConstAutoreject") And Not pbCanDriftCorr)
Da = puTask.bDirectAltPD
Ndp = piNumDauPar
If Ndp = 2 And Da And puTask.saEqns(-2) = "" Then Ndp = 1
foAp.Calculate

For DpNum = 1 To Ndp
  h = rw1: Acol = piaSacol(DpNum)
  Aecol = piaSaEcol(DpNum)
  SSC = " sAcol(" & fsS(DpNum) & ") "
  ssCe = " SaeCol(" & fsS(DpNum) & ") "
  ssCa = " sAgecol(" & fsS(DpNum) & ") "
  s = fsS(DpNum): Ch = 2 + ((pbU And DpNum = 1) Or (pbTh And DpNum = 2))
  Lambda = Choose(Ch, pscLm8, pscLm2):   Ele = Choose(Ch, "U", "Th")
  fClr = Choose(Ch, RGB(0, 0, 100), 100)
  Bclr = Choose(Ch, RGB(200, 255, 200), RGB(200, 200, 255))
  AddName "Arr_" & s, True, h, Acol, rwn, Acol
  ErNm = "Aer_" & s
  AddName ErNm, True, h, Aecol, rwn, Aecol
  AddName "Adat" & s, True, h, Acol, rwn, Aecol
  Set Arange = Range("Adat" & s)
  Set Ara = Range("Arr_" & s)
  Set UPbConst = Range("Arr_" & s)
  Set AerRange = Range(ErNm)
  foAp.Calculate
  Npts = Arange.Count / 2
  ErN = frSr(ErNm).Count
  ReDim ErrVals(1 To ErN)

  For i = 1 To ErN

    If fbIsNum(Arange(i, 2)) And Not IsEmpty(Arange(i, 2)) Then
      ErrVals(i) = Arange(i, 2)
    Else
      ErrVals(i) = 0: Arange(i, 2) = ""
    End If

  Next

  MedianEr = 0
  On Error Resume Next
  MedianEr = foAp.Median(Range(ErNm))
  On Error GoTo 0
  GetMAD ErrVals, ErN, MedianEr, 0, 0
  Nmadd = fdNmad(ErrVals): LargeErRegN = 0

  If Not NoReject Then
    ReDim LargeErRej(1 To 99)

    For i = 1 To ErN

      If Abs(ErrVals(i) - MedianEr) > 10 * Nmadd Or ErrVals(i) = 0 Then

        For j = 0 To 1
          Range(ErNm).Cells(i, j).Font.Strikethrough = True
        Next j

        LargeErRegN = 1 + LargeErRegN: LargeErRej(LargeErRegN) = i
      End If

    Next i

    If LargeErRegN > 0 Then ReDim Preserve LargeErRej(1 To LargeErRegN)
  End If

  h = picWtdMeanAColOffs + rwn
  AddName "WtdMeanA" & s, True, h, Acol
  AddName "WtdMeanAPerr" & s, True, 1 + h, Acol
  ColWidth picAuto, Acol
  k = Acol
  c = Acol + Choose(DpNum, -1, 1)  ' label column
  wHA1 = Choose(DpNum, xlLeft, xlRight)
  wHA2 = Choose(DpNum, xlRight, xlLeft)

  ' *********************************************************************
  If pbCanDriftCorr Then

    frSr(flHeaderRow(True), piLowessHrsCol, , _
         1 + piLowessMeasCol).WrapText = False
    FindStr "Hours", , piHoursCol, plHdrRw
    Set Hours = frSr(plaFirstDatRw(-pbStd), piHoursCol, rwn)

    With LowS
      .iWindow = fvMinMax(foUser("SmoothingWindow"), 4, 99)
      Set .rX = Hours
      Set .rY = UPbConst
      Set .rYsig = AerRange
      .bPercentErrs = True

      SecularTrend LowS, foUser("AutoWindow")

      j = DpNum
      NLowess = UBound(.daX)
      RwS = rw1 + NLowess - 1
      Set LowessHours = frSr(rw1, piLowessHrsCol, rwn)
      LowessHours.Name = "HoursDrift"
      Set LowessDeltaP = frSr(rw1, piLowessDeltaPcol, rwn)
      LowessDeltaP.Name = "ConstDrift"
      Set Sig2LowessDeltaP = frSr(rw1, piLowessDeltaPcol + 1, rwn)
      Sig2LowessDeltaP.Name = "SigmaConstDrift"
      Set LowessRange = Union(LowessHours, LowessDeltaP)
      Set AllDeltaP = frSr(rw1, piLowessMeasCol, rwn)
      AllDeltaP.Name = "AllConstDrift"
      Set Sig2AllDeltaP = frSr(rw1, piLowessMeasCol + 1, rwn)
      Sig2AllDeltaP.Name = "SigmaAllConstDrift"

      For i = 1 To Npts
        AllDeltaP(i) = 100 * (UPbConst(i) / .dMean - 1)
        Sig2AllDeltaP(i) = 2 * AerRange(i)

        If i <= NLowess Then
          LowessHours(i) = .daX(i)
          LowessDeltaP(i) = 100 * (.daY(i) / .dMean - 1)
          Sig2LowessDeltaP(i) = 2 * 100 * .daYsig(i) / .daY(i)
        End If

      Next i

      ExtPtSigma = 100 * .dExtSigma / .dMean
      WtdMeanErr = 100 * .dSigmaMean / .dMean
    End With

    Set AllHrsDeltaP = Union(Hours, AllDeltaP)

  Else ' *********************************************************************
    On Error GoTo 11
    wW = Isoplot3.wtdav(Arange, True, True, 1, Not NoReject, True, 1)
    GoTo 22
11:   On Error GoTo 0
    MsgBox "Wtd average calculation failed -- " & _
           "please check Task definition.", , pscSq
    End
22:   On Error GoTo 0
    k = (wW(3, 2) = "MSWD")
    WtdMean = wW(1, 1):     WtdMeanErr = wW(2, 1)
    ExtPtSigma = IIf(k, 0, wW(3, 1))
    MSWD = wW(4 + k, 1)
    Prob = wW(6 + k, 1)
    ParseLine Trim(wW(7, 1)), Rejected(), Nrej, " "

    If LargeErRegN > 0 Then

      For i = 1 To Nrej
        For j = 1 To LargeErRegN
          If LargeErRej(j) < Rejected(i) Then Rejected(i) = 1 + Rejected(i)
        Next j
      Next i

    End If
    NtotRej = Nrej + LargeErRegN

    If NtotRej > 0 Then
      ReDim Preserve Rejected(1 To Nrej + LargeErRegN)

      For i = Nrej + 1 To LargeErRegN + Nrej
        Rejected(i) = LargeErRej(i - Nrej)
      Next i

      BubbleSort Rejected
      Nrej = NtotRej
      t1 = Rejected(1)  'JOIN?

      For i = 2 To Nrej
        t1 = t1 & ", " & Rejected(i)
      Next i

      wW(7, 1) = t1
    Else
      t1 = "none"
    End If
    Cells(h, Acol) = WtdMean
    Cells(h + 5, Acol) = wW(7, 1)
    Cells(h + 3, Acol) = MSWD
    Cells(h + 4, Acol) = Prob
    t2 = "Wtd Mean of Std Pb/" & Ele & " calibr."
    Cells(h + 1, Acol) = WtdMeanErr
  End If ' **********************************************************

  ColWidth 0.1, 1 + piaSageEcol(1)
  Set ExtPerr = Cells(h + 2, Acol)
  ExtPerr.Name = "ExtPerr" & s
  ExtPerr = ExtPtSigma

  If pbCanDriftCorr Then
    Set w0(DpNum) = Cells(h + 2, Acol)
    Set w(DpNum) = Cells(h + 2, Acol)
    Cells(h, Acol) = LowS.dMean
    Fonts h, Acol - 1, , , vbBlue, True, IIf(DpNum, xlRight, xlLeft), _
           , , , , , , , , "Outlier-resistant mean"
  End If

  If MSWD >= 100 Then
    MSWD = Drnd(MSWD, 3)
    Cells(h + 3, Acol).NumberFormat = "general"
  End If

  Fonts h, Acol, h + 7, c, fClr, True, xlRight

  If pbCanDriftCorr Then ' *********************************************
    With frSr(plHdrRw, piLowessHrsCol, rwn, 1 + piLowessMeasCol)
      With .Font
      .Bold = False: .Size = 9: .Color = vbWhite
      End With
      .Columns.ColumnWidth = 0.1
    End With
  Else

    For i = 0 To 1

      If DpNum = 1 Or i > 0 Then
        k = IIf(i = 0, Acol - 1, piaSageEcol(DpNum) + 1)
        With Columns(k): .ColumnWidth = 0.1: .Font.Color = vbWhite: End With
      End If

    Next i

    StdResFmt h, c, Acol, t2, , , wHA2
  End If ' ****************************************************************

  Fonts h, c, h + 5, , vbBlue, , wHA2
  Range("WtdMeanA" & s).NumberFormat = psaCalibConstNumFor(DpNum)

  If Not pbCanDriftCorr Then
    StdResFmt h + 1, c, Acol, "1sigma error of mean", , fsQq(pscPcnt), wHA1
    StdResFmt h, c, Acol, t2, , , wHA1
    StdResFmt h + 5, c, Acol, pscRejectedSpotNums, t1, "@", wHA1
    StdResFmt h + 3, c, Acol, "MSWD", , pscZd2, wHA1
    StdResFmt h + 4, c, Acol, "Prob. of fit", , pscZd2, wHA1
  End If

  StdResFmt h + 2, c, Acol, "1sigma external spot-to-spot error", , _
            fsQq(pscPcnt), wHA1
  Note h + 2, Acol, "Additional, non-analytical error required to " & _
       "explain observed scatter (will be included in calculation " & _
       "of sample-spot errors)"
  For j = 1 + h To 2 + h: SigConv j, c: Next j
  Box plHdrRw, Acol, rwn, Acol + 1, Bclr
  Box plHdrRw, piaSageCol(DpNum), rwn, piaSageEcol(DpNum), _
      RGB(220 + 30 * (DpNum - 1), 220, 220)

  If pbCanDriftCorr Then
    StdResFmt h + 3, c, Acol, "Width of smoothing window (spots)", , "0", wHA1
    With Cells(h + 3, Acol)
      .Value = LowS.iActualWindow
      .Name = "Window"
    End With
    Fonts h + 4, Acol, , , Hues.peDarkGray, , xlLeft, 10, , , _
          Drnd(LowS.dInitialMSWD, 3), "General"
    Cells(h + 4, Acol).Name = "PseudoMSWD"
    psMswd = "pseudo-mswd on" & StR(LowS.iActualWindow - 3) & " d.f."
    Fonts h + 4, Acol - 1, , , Hues.peDarkGray, True, xlRight, , , , psMswd, "@"

    For Rw = rw1 To rwn
      If Cells(Rw, Acol).Font.Strikethrough Then
        IntClr vbYellow, Rw, Acol, , 3 + Acol
      End If
    Next Rw

  ElseIf Cells(h + 5, Acol).Value > 0 Then

    ' Show rejected A-values with strikethrough
    For i = 1 To Nrej
      j = Arange(Rejected(i), 1).Row
      Fonts j, Acol, , 3 + Acol, StrkThru:=True
      IntClr vbYellow, j, Acol, , 3 + Acol
    Next i

  End If

  Fonts plHdrRw, Acol, rwn, Acol + 1, fClr, True
  Rows(plHdrRw).AutoFit
  AddName "Aadat" & s, True, rw1, 1 + Aecol, rwn, 1 + Aecol
  AddName "Aaerdat1_" & s, True, rw1, 2 + Aecol, rwn, 2 + Aecol
  AddName "Aaerdat2_" & s, True, rw1, 3 + Aecol, rwn, 3 + Aecol
  tmp = IIf(Da, fClr, vbBlack)
  Fonts plHdrRw, piaSageCol(DpNum), rwn, _
        piaSageEcol(DpNum), tmp, True  ' Std age cols

  RangeNumFor pscZq, "Aadat" & s: RangeNumFor pscZq, "Aaerdat1_" & s

  t1 = "=Ln(1+" & SSC & "/WtdMeanA" & s
  t2 = "*" & IIf((pbU And DpNum = 1) Or (pbTh And DpNum = 2), _
       "StdUPbRatio", "StdThPbRatio")
  t3 = ")/" & Lambda

  For i = rw1 To rwn

    If Not IsEmpty(Cells(i, piaSacol(1))) Then

      If pbCanDriftCorr Then
        t1 = "=Ln(1+(1-" & Cells(i, piLowessDeltaPcol).Address & _
             "/100)*" & SSC & "/WtdMeanA" & s
      Else
        t1 = "=Ln(1+" & SSC & "/WtdMeanA" & s
      End If

      PlaceFormulae t1 & t2 & t3, i, piaSageCol(DpNum)
      t4 = "Exp(" & Lambda & "*" & ssCa & ")"
      t5 = ssCe & "/100*(" & t4 & "-1)/" & Lambda & "/" & t4
      PlaceFormulae "=" & t5, i, piaSageEcol(DpNum)
      PlaceFormulae "=2* sAgeE(" & fsS(DpNum) & ") ", i, 1 + piaSageEcol(DpNum)
    End If

  Next i

  HA xlRight, "aaerdat1_" & s
  If DpNum = 1 Then ColWidth picAuto, piaSageCol(DpNum), 1 + piaSageEcol(DpNum)
  If Not pbCanDriftCorr Then
    With Range("WtdMeanA" & s): r = .Row: c = .Column: End With
  End If

  If DpNum = 2 Then
    With frSr(rw1, Acol - 1, rwn)
      .Font.Color = vbWhite:  .ColumnWidth = 0.05
      For i = 1 To .Rows.Count
        .Cells(i, 1) = .Cells(i, -5)
      Next i
    End With
  End If

  If pbCanDriftCorr Then
    Set AvgLine = frSr(rwn + 1, Acol, rwn + 2, Acol + 1)
    tmp = foAp.RoundUp(Hours(Hours.Rows.Count), 0)
    AvgLine.Font.Color = vbWhite
    AvgLine(1, 1) = 0:   AvgLine(1, 2) = 0
    AvgLine(2, 1) = tmp: AvgLine(2, 2) = 0

    SmallChart DataRange:=AllHrsDeltaP, _
               Xname:="Hours", _
               Yname:="Calibr. const. drift (%)", _
               PlaceRow:=[pseudomswd].Row + 1, _
               PlaceCol:=Acol - 12.5, _
               YerrBars:=True, _
               YerrCol:=Sig2AllDeltaP.Column, _
               PercentErrs:=False, _
               SymbLineClr:=vbRed, _
               FontAutoScale:=False, _
               ChartBoxClr:=peLightGray, _
               PlotBoxClr:=peStraw

    Set ChtObj = foLastOb(ActiveSheet.ChartObjects)
    With ChtObj
      .Activate
      .Name = "SquidChart" & fsS(DpNum)
      psaWtdMeanAChartName(DpNum) = .Name
      .Top = .Top + 6
    End With
    Set LowSchart = ActiveChart

    With LowSchart
      .PlotArea.Height = .PlotArea.Height - 10
      .PlotArea.Top = 10
      RejClr = .PlotArea.Interior.Color
      Set SC = ActiveChart.SeriesCollection(1)

      For i = 1 To AllHrsDeltaP.Rows.Count
        If Arange(i, 1).Font.Strikethrough Then
          FormatSeriesCol SC, i, , , , , xlMarkerStyleX, 8, vbBlue, RejClr
        End If
      Next i

      .Axes(1).MaximumScale = foAp.RoundUp(fvMax(Hours), 0)
      .SeriesCollection.Add Source:=LowessRange, Rowcol:=xlColumns, _
           SeriesLabels:=False, CategoryLabels:=True, Replace:=False

      FormatSeriesCol foLastOb(ActiveChart.SeriesCollection), , _
                      xlContinuous, vbRed, xlMedium, MSWD > 0, xlNone

      .SeriesCollection.Add Source:=AvgLine, Rowcol:=xlColumns, _
         SeriesLabels:=False, CategoryLabels:=True, Replace:=False
      Set SerCol = foLastOb(ActiveChart.SeriesCollection)
      FormatSeriesCol SerCol, , xlContinuous, vbBlue, xlHairline, False, xlNone

      TwoSigText
      LowSchart.Deselect
      StdW.Activate
      Cells(rwn, c + 1).Activate
      'frSr(rw1, Acol, rwn, 1 + Acol).Font.Strikethrough = False
      fhSquidSht.Shapes("spinup").Copy
      phStdSht.Paste
      fhSquidSht.Shapes("spindown").Copy
      Cells(Rw + 1, c + 1).Activate
      phStdSht.Paste
      With ActiveSheet.Shapes
        Set Up = ActiveSheet.Shapes("spinup")
        Set Down = .Item("spindown")
      End With
     End With
     StdW.Activate
  Else ' not pbCanDriftCorr
    ExtractGroup True, 0, UPbConst, False, True, 0, , DpNum
    Set ChtObj = foLastOb(ActiveSheet.ChartObjects)
    On Error Resume Next
    psaWtdMeanAChartName(DpNum) = ChtObj.Name
    On Error GoTo 0
    Set w(DpNum) = Range("WtdMeanA" & s)
    Set w0(DpNum) = w(DpNum)
    piaSacol(DpNum) = w(DpNum).Column
    piaSaEcol(DpNum) = 1 + w(DpNum).Column

    If pbCanDriftCorr And DpNum = 1 Then ' ie must be a NumDP=2 type of task.
      r = fnLeftTopRowCol(1, ChtObj.Top) - 1 'fnBottom(ChtObj))
      Fonts r, Acol, , , vbBlue, False, xlRight, 9, , , _
        "(drift correction not supported for this type of Task)"
      Cells(r, Acol).VerticalAlignment = xlTop
    End If

  End If

  With ActiveSheet.ChartObjects(DpNum)
    .Activate
    .Interior.ColorIndex = xlNone
    c = 1 + w0(DpNum).Column
    r = w0(DpNum).Row
    If DpNum = 1 Then
      L = ActiveSheet.Columns(c).Left - .Width
    Else
      L = ActiveSheet.Columns(Acol).Left - 35
    End If
    .Left = L
    Rt = .Left + .Width
    t = .Top
    .Interior.ColorIndex = xlNone

    If Not pbCanDriftCorr Then
      .Top = 12 + ActiveSheet.Rows(r + 6).Top
    End If

    ActiveSheet.Cells(r, c).Select
  End With

  If pbCanDriftCorr Then
    ActiveSheet.Cells(1, 1).Select
    Up.Top = Cells(rwn + 4, c + 1).Top
    Down.Top = fnBottom(Up)
    Down.Left = Rt + 5
    Up.Left = Rt + 5
    Cells(rwn + 5, c + 1) = "Click to change smoothing" ' /
    Cells(rwn + 6, c + 1) = "window or to recalculate"  '| -- added
    frSr(rwn + 5, c + 1, rwn + 6).Font.italic = True    ' \

  End If

  If piaNumSpots(1) > 1 Then
    Bclr = Choose(DpNum, RGB(200, 255, 200), RGB(200, 200, 255))

    If Not pbCanDriftCorr Then
      t = 0
      On Error Resume Next
      With ActiveSheet.ChartObjects(psaWtdMeanAChartName(DpNum))
        .Top = .Top + 12
        L = .Left: t = .Top + .Height + 25
      End With
      On Error GoTo 0
      If t = 0 Then t = phStdSht.Rows(rwn + 15).Top
      t1 = psaPDeleRat(DpNum)
      ActiveSheet.Cells(r, Acol).Select
      r = rwn + 5
      Do
         r = r + 1
      Loop Until fnBottom(Rows(r)) > t
    End If

    c = w0(DpNum).Column + Choose(DpNum, -8, 0)
    t2 = " Assigned " & t1 & " external spot-to-spot error (1s) "


    If DpNum = 1 Then
      With ActiveSheet.ChartObjects("squidchart1")
        L = .Left + .Width - 150: t = .Top + .Height
      End With
      r = fnLeftTopRowCol(1, t)
      c = piaSacol(DpNum) - 2
      Col1 = fnLeftTopRowCol(2, L + 30)
      LastExtboxR = r
    Else
      If LastExtboxR > 0 Then r = LastExtboxR
      Col1 = c + 3
    End If

    StdW.Activate

    Set extBox = frSr(r, c, r + 2, c + 2)
    Box extBox, , , , Bclr, True
    With extBox
      .Merge: .HorizontalAlignment = xlCenter
      .Font.Size = 12: .Font.Bold = True: .Font.Color = 0
      .Name = "ExtPerrA" & s: t3 = .Name
      .Formula = "=Max(" & fsS(foUser("MinUPbErr")) & _
                 ", " & ExtPerr.Address & ")"
      .VerticalAlignment = xlCenter
      .Select
    End With
    t1 = "Assigned Pb/" & Ele & " external err (1s)"
    Rw = extBox.Row + 3: Col1 = extBox.Column - 1
    Fonts Rw, Col1, , , , , xlLeft, , , , t1
    Cells(Rw, Col1).Characters(Len(t1) - 1, 1).Font.Name = "symbol"
    ExtractRowDblDbl Adrift(), ad(), DpNum, True
    ExtractRowDblDbl AdriftErr(), adE(), DpNum, True

    If Not pbCanDriftCorr And foAp.Sum(ad) > 0 Then
      SimpleWtdAv piaSpotCt(1), ad(), adE(), MeanAdrift, , maErr95
      Set TmpR = Range("ExtPerrA" & fsS(DpNum))
      L = Choose(DpNum, fnRight(TmpR), TmpR.Left)
      ActiveSheet.Shapes.AddShape(5, L, fnBottom(TmpR) + 20, 1, 1).Select
      With Selection
        With .ShapeRange
          .Fill.ForeColor.RGB = RGB(255, 128, 255) 'Bclr
          .Line.ForeColor.RGB = vbBlack
        End With
        t1 = "Mean within-spot " & psaPDeleRat(DpNum) & _
          " calibr. drift per min. (%) = " & _
           Format(MeanAdrift, pscZd2) & vbLf & _
           pscPm & "95% conf. = " & Format(maErr95, " 0.00 ")
        .Characters.Text = t1: .Font.Size = 11
        .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
        .AutoSize = True
        .Left = Choose(DpNum, fnRight(TmpR) - .Width, TmpR.Left)
      End With
    End If
    extBox(1, 1).Name = "ExtPerrA" & s
  End If ' Nscans>3

Next DpNum ' For DpNum = 1 To Ndp

BorderLine xlBottom, 2, plHdrRw, 1, , piaAgePb76_4eCol(1)
ClearObj extBox, Arange, AerRange
End Sub

Sub IncrWindow()
UpdateLowess 1
End Sub

Sub DecrWindow()
UpdateLowess -1
End Sub

Sub UpdateLowess(ByVal WindowIncr%)
Dim tB As Boolean, tmp$
Dim i%, Wind%, TotN%, N%, UnstrukClr&
Dim AllConstDrift As Range, HoursDrift As Range
Dim SigmaAllConstDrift As Range, ConstDrift As Range
Dim SigmaConstDrift As Range, Hours As Range, StdConst As Range
Dim w As Range, RawConstRange As Range, DriftCorrected As Range, SqCht As ChartObject, LowS As Lowess

foAp.Calculation = xlManual
Set w = ActiveSheet.Range("Window")
Wind = WindowIncr + w
If Wind < 4 Then Exit Sub
NoUpdate True

With Range("Arr_1")
  TotN = .Count
  UnstrukClr = .Item(0).Interior.Color
End With

With LowS
  .bPercentErrs = True
  Set .rX = Range("StdHrs")
  Set .rY = Range("Arr_1")
  Set .rYsig = Range("Aer_1")
  .iWindow = fvMinMax(Wind, foUser("MinWIndow"), TotN)
  N = 0
  ReDim .daX(TotN), .daY(TotN), .daYsig(TotN)

  plHdrRw = flHeaderRow(True)
  FindStr "Hours", , piHoursCol, plHdrRw, , plHdrRw, , True, , , , , , , , , True
  Set Hours = frSr(plaFirstDatRw(1), piHoursCol, plaLastDatRw(1))
  Set HoursDrift = Range("HoursDrift")
  Set ConstDrift = Range("ConstDrift")
  Set SigmaConstDrift = Range("SigmaConstDrift")
  Set AllConstDrift = Range("AllConstDrift")
  Set SigmaAllConstDrift = Range("SigmaAllConstDrift")
  Set StdConst = Range("arr_1")

  For i = 1 To TotN
    tB = False
    With Range("arr_1")(i)
      If Not IsNumeric(.Value) Then .Formula = ""

      If .Formula = "" Or .Value = 0 Or .Font.Strikethrough Then
        tB = True
        .Font.Strikethrough = True
        .Interior.Color = vbYellow
      Else
        N = 1 + N
        LowS.daX(N) = Hours(i)
        LowS.daY(N) = AllConstDrift(i)
        LowS.daYsig(N) = SigmaAllConstDrift(i)
      End If

      With frSr(.Row, .Column, , .Column + 3)
        .Font.Strikethrough = tB
        .Interior.Color = IIf(tB, vbYellow, UnstrukClr)
      End With
    End With
  Next i
  If N < 4 Then Exit Sub

  SecularTrend LowS

  FindDriftCorrRanges RawConstRange, DriftCorrected   ' /
  DriftCorr RawConstRange, DriftCorrected             '|  09/07/17 -- added
  phStdSht.Activate                                   ' \

  Range("WtdMeanA1") = .dMean
  Range("ExtPerr1") = 100 * .dExtSigma / .dMean
  Fonts rw1:=[Window], Clr:=vbWhite, Bold:=True, HorizAlign:=xlCenter, _
         Size:=12, Formul:=.iActualWindow, InteriorColor:=Hues.PeDarkRed
  Fonts [pseudomswd], , , , Hues.PeDarkRed, , xlCenter, 10, _
         , , Drnd(.dInitialMSWD, 3), "General"
  tmp = "pseudo-mswd on" & StR(LowS.iActualWindow - 3) & " d.f."
  Fonts rw1:=[pseudomswd].Cells(1, 0), Clr:=Hues.PeDarkRed, Phrase:=tmp
  HoursDrift.Clear
  ConstDrift.Clear
  SigmaConstDrift.Clear
  AllConstDrift.Clear
  SigmaAllConstDrift.Clear
  N = UBound(.daX)

  For i = 1 To TotN
    If i <= N Then
      HoursDrift(i) = .daX(i)
      ConstDrift(i) = 100 * (.daY(i) / .dMean - 1)
      SigmaConstDrift(i) = 2 * 100 * .daYsig(i) / .daY(i)
    End If
    If Not StdConst(i).Font.Strikethrough Then
      AllConstDrift(i) = 100 * (Range("Arr_1")(i) / .dMean - 1)
      SigmaAllConstDrift(i) = 2 * Range("Aer_1")(i)
    End If
  Next i

  .iWindow = .iActualWindow
  Set SqCht = ActiveSheet.ChartObjects("SquidChart1")
  SqCht.Activate

  AxisScale AllConstDrift, False, SigmaAllConstDrift
  [WtdMeanA1].Cells(5, 1).Select
  NoUpdate False
End With
foAp.Calculate
End Sub

Sub DriftCorr(RawConstRange As Range, DriftCorrected As Range)

Dim t1$, t2$, t3$, t4$, Std$, f$, SamConst$, SamHrs$
Dim StdDrift$(3), StdHrs$(3)
Dim i%, j%, k%, HoursCol%, Ndrift%, Nsam%, Indx%(3), Hrow&
Dim Std_Drift As Range, Hours As Range, Std_Hrs As Range

Hrow = flHeaderRow(False)
Set phStdSht = ActiveWorkbook.Sheets(pscStdShtNa)  ' /
Std = pscStdShtNa & "!"                            '| 09/07/17 -- added
phStdSht.Activate                                  ' \

With phStdSht.[ConstDrift]
  Ndrift = .Rows.Count

  Do While .Cells(Ndrift, 1) = ""
    Ndrift = Ndrift - 1
  Loop

  Set Std_Drift = frSr(.Row, .Column, .Row + Ndrift - 1)
End With

With phStdSht.[HoursDrift]
  Set Std_Hrs = frSr(.Row, .Column, .Row + Ndrift - 1)
End With
phSamSht.Activate
FindStr "Hours", , HoursCol, Hrow, 1, Hrow, , True
Set Hours = frSr(plaFirstDatRw(0), HoursCol, plaLastDatRw(0))
Nsam = RawConstRange.Count

For i = 1 To Nsam
  SamHrs = Hours(i).Address(0, 0)
  SamConst = RawConstRange(i).Address(0, 0)

  If Hours(i) < Std_Hrs(1) Then
    Indx(1) = 1: Indx(2) = 2: Indx(3) = 1
    'f = [ StdHrs(1) - Hours(i) ] / [ StdHrs(2) - StdHrs(1) ]
    'Drift = StdDrift(1) - f * [ StdDrift(2) - StdDrift(1) ]
  ElseIf Hours(i) > Std_Hrs(Ndrift) Then
    Indx(1) = Ndrift: Indx(2) = Ndrift: Indx(3) = Ndrift - 1
    'f = [ Hours(i) - StdHrs(Ndrift) ] / [ StdHrs(Ndrift) - StdHrs(Ndrift - 1) ]
    'Drift = StdDrift(Ndrift) + f * [ StdDrift(Ndrift) - StdDrift(Ndrift-1) ]
  Else

    For j = 1 To Ndrift - 1
      k = j + 1
      If Hours(i) > Std_Hrs(j) And Hours(i) < Std_Hrs(k) Then
        Indx(1) = j: Indx(2) = k: Indx(3) = j
        'f = [ Hours(i) - StdHrs(j) ] / [ StdHrs(j+1) - StdHrs(j) ]
        'Drift = StdDrift(j) + f * [ StdDrift(j+1) - StdDrift(j) ]
        Exit For
      End If
    Next j
  End If

  For j = 1 To 3
    StdHrs(j) = Std & Std_Hrs(Indx(j)).Address
    StdDrift(j) = Std & Std_Drift(Indx(j)).Address
  Next j

  'Interpolated std bias% = StdDrift1+{SamHrs-StdHrs1)/
  '                        (StdHrs2-StdHrs3)*(StdDrft2-StdDrift3)

  t1 = "(" & StdDrift(1) & "+(" & SamHrs & "-" & StdHrs(1) & ")/("
  t2 = StdHrs(2) & "-" & StdHrs(3) & ")*(" & StdDrift(2) & _
       "-" & StdDrift(3) & "))"
  t3 = t1 & t2
  t4 = "=(1-" & t3 & "/100)*" & SamConst
  DriftCorrected(i) = t4

Next i
End Sub
