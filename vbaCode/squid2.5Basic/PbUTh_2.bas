Attribute VB_Name = "PbUTh_2"
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
' 09/03/02 -- All lower array bounds explicit

Option Explicit
Option Base 1

Sub ConcordiaClr(arc%)
' Color-code & outline the columns for concordia plots,
'  & put rejected rows in yellow+strikethrough.
Dim Struk As Boolean
Dim d As String * 17, j%, Col1%, Col2%, c3%
Dim ChtObj As Object

If InStr(ActiveSheet.Name, pscStdShtNa) Then Exit Sub
d = String(17, "-")
FindStr "total238/206", 0, Col1, plHdrRw
If Col1 = 0 Then Exit Sub
Box plHdrRw, Col1, plaLastDatRw(0), 3 + Col1, RGB(255, 150, 255)
Box plHdrRw, 4 + Col1, plaLastDatRw(0), 7 + Col1, RGB(255, 128, 128)
Box plHdrRw, 8 + Col1, plaLastDatRw(0), 12 + Col1, RGB(200, 255, 200)
Alerts False
Cells(plHdrRw - 1, 4 + Col1) = d & " 204 corrected " & d
Fonts plHdrRw - 1, 4 + Col1, , 12 + Col1, , True, xlCenter, , , , , , , True
FindStr "8corr238/206*", 0, Col2, plHdrRw

If Col2 > 0 Then
  Cells(plHdrRw - 1, Col2) = d & " 208 corrected " & d
  Fonts plHdrRw - 1, Col2, , 8 + Col2, , True, xlCenter, , , , , , , True
  Box plHdrRw, Col2, plaLastDatRw(0), Col2 + 3, RGB(255, 255, 200)
  Box plHdrRw, Col2 + 4, plaLastDatRw(0), Col2 + 8, RGB(200, 200, 255)
  Col2 = Col2 + 8
Else
  FindStr "err|corr", 0, Col2, plHdrRw
  FindStr "err|corr", 0, c3, plHdrRw, 1 + Col2 ' Find 2nd "err corr" in case 8corr conc cols

  If c3 > 0 Then
    Box plHdrRw, 1 + Col2, plaLastDatRw(0), 4 + Col2, RGB(255, 255, 200)
    Box plHdrRw, c3 - 4, plaLastDatRw(0), c3, RGB(200, 200, 255)
    Col2 = c3
  End If

End If

Alerts True
If ActiveSheet.ChartObjects.Count > 0 Then

  For Each ChtObj In ActiveSheet.ChartObjects
    If LCase(Left$(ChtObj.Name, 11)) = "squidchart" Then

      For j = plaFirstDatRw(0) To plaLastDatRw(0)
        Struk = Cells(j, arc).Font.Strikethrough
        If Struk Then IntClr vbYellow, j, Col1, , Col2
        Cells(j, 1 + arc).Font.Strikethrough = Struk
      Next j

      Exit For

    End If
  Next ChtObj

End If
Set ChtObj = Nothing
End Sub

Function fbHasRatio(NumerIso#, DenomIso#) As Boolean
' Are any of the Task isotope-ratios NumerIso/DenomIso ?
Dim RatNum%, Numer#, Denom#

NumerIso = Drnd(NumerIso, 4)
DenomIso = Drnd(DenomIso, 4)

With puTask

  For RatNum = 1 To .iNrats
    NumDenom .saIsoRats(RatNum), Numer, Denom
    Numer = Drnd(Numer, 4)
    Denom = Drnd(Denom, 4)
    If NumerIso = Numer And DenomIso = Denom Then Exit For
  Next RatNum

  fbHasRatio = (RatNum <= .iNrats)
End With
End Function

Sub AssignUPbColNumbers()
' Assign column numbers for a U/Pb geochron Task & create related
'  column-index values.
' 09/07/02 -- Major changes to accomodate columns 7cor46, 8cor46,
'             4corcom6, 7corcom6, 8corcom6, 4corcom8, 7corcom8,4cor86,
'             7cor86, 4corppm6, 7corppm6, 8corppm6, 4corppm8, 7corppm8

Dim HasThU As Boolean, HasPbU As Boolean, HasPbTh As Boolean, tB As Boolean, Da As Boolean
Dim CanPbU As Boolean, CanPbTh As Boolean
Dim s$, t$, Nom$, Hdr$
Dim i%, p%, q%, m%, DpNum%, Col%, sCol%, tmpCol%, testCol%
Dim Nomi#, t1#, t2#

ThisWorkbook.Sheets("ColIndex").Activate
Cells.Delete
Columns(1).ColumnWidth = 18

With puTask
  Da = .bDirectAltPD
  piPrimeDP = 1 - (.iParentIso = 232)
  pbHasU = (pi238PkOrder > 0 Or pi254PkOrder > 0 Or pi270PkOrder > 0)
  pbHasTh = (pi232PkOrder > 0 Or pi248PkOrder > 0 Or pi264PkOrder > 0)
  HasPbU = (piPrimeDP = 1 Or piNumDauPar = 2)
  HasPbTh = (piPrimeDP = 2 Or piNumDauPar = 2) Or _
            (piStdCorrType = 2 And pbHasTh = True)
  ' 03/16/10 = second part of above line added to force addition of 8-corr 206* column.
  tB = (.saEqns(-2) <> "" Or .saEqns(-3) <> "")
  CanPbU = (.bIsUPb Or (Not .bIsUPb And tB))
  CanPbTh = (Not .bIsUPb Or (.bIsUPb And tB))
  tB = (pbHasU And pbHasTh)
  HasThU = ((tB And Not Da) Or (Da And piNumDauPar = 2))
  pbHasUconc = pdConcStdPpm > 0 And (pbHasU Or (Da And pbHasTh))
  pbHasThConc = pdConcStdPpm > 0 And (pbHasTh Or (Da And pbHasU))
  Col = 0
  ColInc Col, "Spot|Name", "Name", 1, piNameCol
  ColInc Col, "Date/Time", "DateTime", 1, piDateTimeCol
  ColInc Col, "Hours", "Hours", 1, piHoursCol

  If pbXMLfile Then
    ColInc Col, "stage|X", "StageX", 1, piStageXcol
    ColInc Col, "stage|Y", "StageY", 1, piStageYcol
    ColInc Col, "stage|Z", "StageZ", 1, piStageZcol
    ColInc Col, "Qt1y", "Qt1y", 1, piQt1yCol
    ColInc Col, "Qt1z", "Qt1y", 1, piQt1Zcol
    ColInc Col, "Primary|beam|(na)", "PrimaryBeam", 1, piPrimaryBeamCol
  End If

  ColInc Col, "Bkrd|cts|/sec", "BkrdCts", 1, piBkrdCtsCol, , piBkrdPkOrder > 0
  If pi204PkOrder > 0 Then .baCPScol(pi204PkOrder) = True
  If pi206PkOrder > 0 Then .baCPScol(pi206PkOrder) = True

  For i = 1 To .iNpeaks

    If .baCPScol(i) And i <> piBkrdPkOrder Then
      Nomi = .daNominal(i): Nom = fsS(Nomi)
      psaCPScolHdr(i) = "total|" & Nom & "|cts|/sec"

      If Nomi = 204 Or Nomi = 206 Then
        Hdr = "Pb"
      Else
        Hdr = "Iso"
      End If

      ColInc Col, "total|" & Nom & "|cts|/sec", Hdr & Nom & "cts", 1, piaCPScol(i)
      piFirstRatCol = piaCPScol(i) + 1

      If Nomi = 206 Then
        piPb206ctsCol = piaCPScol(i)
      ElseIf Nomi = 204 Then
        piPb204ctsCol = piaCPScol(i)
      End If

    End If
  Next i


  If fbHasRatio(204, 206) Then
    ColInc Col, "204|/206", "Pb46", 2, piPb46col, , pi204PkOrder > 0, piPb46eCol, pscPpe
  End If

  piFirstRatCol = piPb46col
  ColInc Col, "207|/206", "Pb76", 2, piPb76col, , pi207PkOrder > 0, piPb76eCol, pscPpe
  If piFirstRatCol = 0 Then piFirstRatCol = piPb76col
  ColInc Col, "208|/206", "Pb86", 2, piPb86col, , pi208PkOrder > 0, piPb86eCol, pscPpe
  If piFirstRatCol = 0 Then piFirstRatCol = piPb86col

  For i = 1 To .iNrats
    t1 = .daNmDmIso(1, i): t2 = .daNmDmIso(2, i): piaIsoRatCol(i) = 0

    If t2 = 206 Then
      If t1 = 204 Then piaIsoRatCol(i) = piPb46col: piaIsoRatEcol(i) = piPb46eCol
      If t1 = 207 Then piaIsoRatCol(i) = piPb76col: piaIsoRatEcol(i) = piPb76eCol
      If t1 = 208 Then piaIsoRatCol(i) = piPb86col: piaIsoRatEcol(i) = piPb86eCol
    End If

    If piaIsoRatCol(i) = 0 Then
      s = fsS(t1) & "|/" & fsS(t2)
      ColInc Col, s, s, 2, piaIsoRatCol(i), , True, piaIsoRatEcol(i), pscPpe
    End If
  Next i

  sCol = Col
  AddEqnCols False, Col, Col, False
  AddEqnCols True, sCol, sCol, False

  If pbCanDriftCorr Then

      For i = 1 To 5
        psaLowessColHdrs(i) = Array("Lowess|cconst|hours", "Lowess|cconst|delta%", pscPm, _
                                   "meas.|cconst|delta%", pscPm)(i)
      Next i

      ColInc sCol, psaLowessColHdrs(1), "LowessHrs", 1, piLowessHrsCol
      ColInc sCol, psaLowessColHdrs(2), "LowessDeltaP", 2, piLowessDeltaPcol, , , , pscPm
      ColInc sCol, psaLowessColHdrs(4), "LowessMeas", 2, piLowessMeasCol, , , , pscPm

    ColInc Col, "Raw|calib|const", "RawCalibConst", 1, piUnDriftCorrConstCol
  End If

' --------------------------------------------------------------------
  For DpNum = 1 To piNumDauPar ' Add sample-sheet calibr-const columns
    s = IIf((pbU And DpNum = 1) Or (pbTh And DpNum = 2), "206Pb|/238U", "208Pb|/232Th")
    s = s & "|calibr.|const"
    ColInc Col, s, "A", 2, piaAcol(DpNum), DpNum, , piaAeCol(DpNum), pscPm & "1s"
    piaEqCol(0, -DpNum) = piaAcol(DpNum)
    piaEqEcol(0, -DpNum) = piaAeCol(DpNum)
    .saEqnNames(-DpNum) = fsLegalName(s)
  Next DpNum

  ' 09/07/02 -- added ----------------------
  ColInc Col, "7-corr|204Pb|/206Pb", "Pb46_7", 1, piPb46_7col, , _
               piPb76col > 0 And CanPbU
  ColInc Col, "8-corr|204Pb|/206Pb", "Pb46_8", 1, piPb46_8col, , _
               piPb86col > 0 And pbHasU And CanPbTh
  ColInc Col, "4-corr|%com|206", "Com6_4", 1, piCom6_4col, , _
               piPb46col > 0
  ColInc Col, "7-corr|%com|206", "Com6_7", 1, piCom6_7col, , _
              piPb76col > 0 And CanPbU
  ColInc Col, "8-corr|%com|206", "Com6_8", 1, piCom6_8col, , _
               piPb86col > 0 And CanPbU And CanPbTh
  ColInc Col, "4-corr|%com|208", "Com8_4", 1, piCom8_4col, , _
               piPb46col > 0 And piPb86col > 0
  ColInc Col, "7-corr|%com|208", "Com8_7", 1, piCom8_7col, , _
              piPb76col > 0 And CanPbU And CanPbTh
  ColInc Col, "4-corr|208Pb*|/206Pb*", "Pb86_4", 2, piPb86_4col, , _
        piPb46col > 0 And piPb86col > 0, piPb86_4ecol, pscPm & "1s"
  ColInc Col, "7-corr|208Pb*|/206Pb*", "Pb86_7", 2, piPb86_7col, , _
        piPb76col > 0 And piPb86col > 0 And CanPbU And CanPbTh, _
        piPb86_7ecol, pscPm & "1s"
  ' ---------------------------------------------

  tB = IIf(pbUconcStd, pbHasUconc, pbHasThConc)
  piaPpmUcol(0) = 0: piaPpmThcol(0) = 0
  piaEqCol(0, -4) = 0

  If .saEqns(-4) <> "" And psConcStdNa <> "" Then ' Add ppmU and ppmTh cols to sample sheet
    If pbUconcStd And pbHasUconc Then  ' 09/07/21 -- below subst. " pbHasUconc" for "tB"
      ColInc Col, "ppm|U", "ppmU", 1, piaPpmUcol(0), 0, pbHasUconc And foUser("siConcStdPpm") > 0
      piaEqCol(0, -4) = piaPpmUcol(0)
      .saEqnNames(-4) = "ppmU"
      ColInc Col, "ppm|Th", "ppmTh", 1, piaPpmThcol(0), 0, pbHasThConc And foUser("siConcStdPpm") > 0
    ElseIf pbThConcStd And pbHasThConc Then
      ColInc Col, "ppm|Th", "ppmTh", 1, piaPpmThcol(0), 0, tB
      piaEqCol(0, -4) = piaPpmThcol(0)
      .saEqnNames(-4) = "ppmTh"
      ColInc Col, "ppm|U", "ppmU", 1, piaPpmUcol(0), 0, pbHasUconc
    End If
  End If

  ' 09/07/02 -- added -----------------------------------------------
  tB = (piPb46col > 0 And ((pbUconcStd And piaPpmUcol(0) > 0) Or _
       (pbThConcStd And piaPpmThcol(0) > 0)))
  s = "-corr|ppm|20"
  t = "RadDau_"

  ' for  ppm daughter-isotope (sample only)
  ColInc Col, "4" & s & "6*", t & "4", 1, piaRadDauCol_4(6), 6, _
               CanPbU And piPb46col > 0
  ColInc Col, "7" & s & "6*", t & "7", 1, piaRadDauCol_7(6), 6, _
              CanPbU And piPb76col > 0
  ColInc Col, "8" & s & "6*", t & "8", 1, piaRadDauCol_8, , _
              CanPbU And piPb86col > 0
  ColInc Col, "4" & s & "8*", t & "4", 1, piaRadDauCol_4(8), 8, _
              CanPbTh And piPb46col > 0
  ColInc Col, "7" & s & "8*", t & "7", 1, piaRadDauCol_7(8), 8, _
              CanPbTh And piPb76col > 0
  ' ------------------------------------------------------------------

  If HasThU And piPb86col > 0 Then
    ColInc Col, "232Th|/238U", "Th2U8", 1, piaTh2U8col(0), 0
    piaEqCol(0, -3 - Da) = Col


    ColInc Col, pscPpe, "Th2U8e", 1, piaTh2U8ecol(0), 0
    If Not Da Then
      piaEqEcol(0, -3) = Col
    End If

  End If

  ColInc Col, "Total|206Pb|/238U", "Pb6U8_tot", 2, piPb6U8_totCol, , _
    pbHasU And pi206PkOrder > 0, piPb6U8_totEcol, pscPpe
  ColInc Col, "Total|208Pb|/232Th", "Pb8Th2_tot", 2, piPb8Th2_totCol, , _
     pbHasTh And pi208PkOrder > 0, piPb8Th2_totEcol, pscPpe
  IncrementColNums Col, 1

  ' Create columns for Standard sheet.
  piSqidNumCol = sCol - 1
  ColIndx piSqidNumCol, "SqidNum", "SqidNum"
  sCol = 1 + sCol + pbCanDriftCorr

  For DpNum = 1 To piNumDauPar        'add std calibr-const & age cols
    If DpNum = 2 Then sCol = 1 + sCol ' for 2nd time scol
    s = IIf((pbU And DpNum = 1) Or (pbTh And DpNum = 2), "206Pb|/238U", "208Pb|/232Th")
    s = s & "|calibr.|const"
    ColInc sCol, s, "sA", 2, piaSacol(DpNum), DpNum, , piaSaEcol(DpNum), pscPpe
    piaEqCol(-1, -DpNum) = piaSacol(DpNum)
    piaEqEcol(-1, -DpNum) = piaSaEcol(DpNum)
    ColInc sCol, Choose(DpNum, "", "alt|") & "Age|(Ma)", "sAge", 2, _
      piaSageCol(DpNum), DpNum, , piaSageEcol(DpNum), psPm1sig
    sCol = sCol + 1 ' to leave room for 2sigma errors
    piSlastCol = sCol
  Next DpNum

  ' 09/07/02 -- added ------------------------------------------------
  s = "-corr|%com|20" 'ppm|20"  10/11/18 - mod to the correct text
  t = "StdCom"

  If piStdCorrType = 0 Then  ' 10/11/18 - was "6*", "8*"
    ColInc sCol, "4" & s & "6", t & "6_4", 1, piStdCom6_4col, , HasPbU
    ColInc sCol, "4" & s & "8", t & "8_4", 1, piStdCom8_4col, , HasPbTh
  ElseIf piStdCorrType = 1 Then
    ColInc sCol, "7" & s & "6", t & "6_7", 1, piStdCom6_7col, , HasPbU
    ColInc sCol, "7" & s & "8", t & "8_7", 1, piStdCom8_7col, , HasPbU And HasPbTh
  ElseIf piStdCorrType = 2 Then
    ColInc sCol, "4" & s & "8", t & "8_4", 1, piStdCom8_4col, , HasPbTh
    ColInc sCol, "8" & s & "6", t & "6_8", 1, piStdCom6_8col, , HasPbU And HasPbTh
  End If

  If piStdCorrType < 2 Then
    ColInc sCol, IIf(piStdCorrType = 0, "4", "7") & "-corr|208Pb*|/206Pb*", _
           "StdRadPb86", 2, piStdRadPb86col, , piPb86col > 0, piStdRadPb86ecol, _
           "StdRadPb86e"
  End If
  ' --------------------------------------------------------------------

  piaPpmUcol(1) = 0: piaPpmThcol(1) = 0:  piaEqCol(-1, -4) = 0

  If pbUconcStd Then
    ColInc sCol, "ppm|U", "ppmU", 1, piaPpmUcol(1), 1, piaPpmUcol(0) > 0
    piaEqCol(-1, -4) = piaPpmUcol(1)
    ColInc sCol, "ppm|Th", "ppmTh", 1, piaPpmThcol(1), 1, piaPpmThcol(0) > 0
  ElseIf pbThConcStd Then
    ColInc sCol, "ppm|Th", "ppmTh", 1, piaPpmThcol(1), 1, piaPpmThcol(0) > 0
    piaEqCol(-1, -4) = piaPpmThcol(1)
    ColInc sCol, "ppm|U", "ppmU", 1, piaPpmUcol(1), 1, piaPpmUcol(0) > 0
  End If

  ColInc sCol, "232Th|/238U", "Th2U8", 2, piaTh2U8col(1), 1, HasThU And piPb86col > 0, _
    piaTh2U8ecol(1), pscPpe
  piaEqCol(-1, -3) = piaTh2U8col(1): .saEqnNames(-3) = "232Th238U"


  If foUser("showoverctcols") And pbHasU Then
    s = "204|overcts|/sec|(fr. 20"
    ColInc sCol, s & "7)", "OverCts4", 1, piaOverCts4Col(7), 7, _
      piPb46col > 0 And piPb76col > 0
    ColInc sCol, s & "8)", "OverCts4", 1, piaOverCts4Col(8), 8, _
      piPb46col > 0 And piPb86col > 0 And pbHasTh

    ColInc sCol, "204|/206|(fr. 207)", "OverCts46", 2, _
         piaOverCts46Col(7), 7, piaOverCts4Col(7) > 0 And piPb46col > 0, _
         piaOverCts46eCol(7), pscPpe
    ColInc sCol, "204|/206|(fr. 208)", "OverCts46", 1, piaOverCts46Col(8), _
      8, piaOverCts4Col(8) > 0
    s = "-corr|206Pb|/238U|const.|delta%"
    ColInc sCol, "7" & s, "corrAdelt", 1, piacorrAdeltCol(7), 7, piaOverCts4Col(7) > 0
    ColInc sCol, "8" & s, "corrAdelt", 1, piacorrAdeltCol(8), 8, piaOverCts4Col(8) > 0
  End If

  If piStdCorrType < 2 And piNumDauPar = 1 Then ' 09/06/16 -- added
    s = fsS(Choose(1 + piStdCorrType, 4, 7)) & "-corr|"
  Else
    s = ""
  End If

  For DpNum = 1 To piNumDauPar
    s = "Uncorr|" & psaPDeleRat(DpNum) & "|const"
    ColInc sCol, s, "sUncorrA", 2, piaStdUnCorrAcol(DpNum), DpNum, True, _
      piaStdUnCorrAerCol(DpNum), pscPpe
  Next DpNum

  s = IIf(piPb46col > 0, "4-corr", "uncorr") & "|207Pb|/206Pb"
  t = s & "|age"
  ColInc sCol, t, "AgePb76_4", 2, piaAgePb76_4Col(1), 1, piPb76col > 0 And piPb46col > 0, _
    piaAgePb76_4eCol(1), psPm1sig
  ColInc sCol, s, "StdPb76_4", 2, piStdPb76_4Col, , piPb76col > 0 And piPb46col > 0, _
    piStdPb76_4eCol, psPm1sig

  If foUser("StdConcPlots") Then
    ColInc sCol, "207*Pb|/235U", "StdPb7U5_4", 2, piStdPb7U5_4col, , piPb76col > 0 And piPb46col > 0, _
      piStdPb7U5_4eCol, psPm1sig
    ColInc sCol, "206*Pb|/238U", "StdPb6U8_4", 2, piStdPb6U8_4col, , piPb76col > 0 And piPb46col > 0, _
      piStdPb6U8_4eCol, psPm1sig
    ColInc sCol, "err|corr", "StdPb7U5Pb6U8_4rho", 1, piStdPb7U5Pb6U8_4rhoCol, , piPb76col > 0 And piPb46col > 0

  End If

  AddEqnCols True, sCol, sCol, True
  piSlastCol = sCol

End With

Col = piLastCol
AddEqnCols False, Col, Col, True

End Sub

Sub AddEqnCols(StdCalc As Boolean, ColIn%, ColOut%, LA_ONLY As Boolean)
' Assign column numbers for Task Equations of a U/Pb geochron Task
'  & create related column-index values.

Dim STA As Boolean, Seq$, Seq2$, i%, j%, Col%

Col = ColIn
With puTask

  For i = 1 To .iNeqns
    Seq = fsStrip(.saEqnNames(i))

    If InStr(Seq, "<<solve>>") = 0 Then
      With .uaSwitches(i)
        STA = (StdCalc And .SA) Or (Not StdCalc And .ST)

        If Not STA And ((.LA And LA_ONLY) Or (Not .LA And Not LA_ONLY)) Then

            If Not .Ar Then
              ColInc Col, Seq, Seq, 2 + (.SC Or .FO), piaEqCol(StdCalc, i), , , piaEqEcol(StdCalc, i), pscPpe
              If piaEqCol(StdCalc, i) > 0 And (.SC Or .FO) Then piaEqEcol(StdCalc, i) = 0
            Else

              For j = 1 To .ArrNcols

                If j = 1 Then
                  ColInc Col, Seq, Seq, piaEqCol(StdCalc, i), piaEqCol(StdCalc, i)
                Else
                  Seq2 = "Eqn" & fsS(i) & "ArrCol" & fsS(j)
                  ColInc Col, "", Seq2, 1, piaEqCol(StdCalc, i) + j - 1
                End If

              Next j

            End If  ' not AR

          End If

      End With
    End If

  Next i

End With
ColOut = Col
End Sub

Sub IncrementColNums(ByRef FromCol%, NumIncr%)
' Increment all col numbers >=AgePb6U8_4col by NumIncr%.  U/Pb Tasks only.
Dim Col%, i
Col = FromCol

For i = 1 To NumIncr
  ColInc Col, "204corr|206Pb|/238U|Age", "AgePb6U8_4", 2, piAgePb6U8_4col, , _
    pbHasU, piAgePb6U8_4ecol, psPm1sig
  ColInc Col, "207corr|206Pb|/238U|Age", "AgePb6U8_7", 2, piAgePb6U8_7col, , _
    piPb46col > 0 And pi207PkOrder > 0 And pbHasU, piAgePb6U8_7ecol, psPm1sig
  ColInc Col, "208corr|206Pb|/238U|Age", "AgePb6U8_8", 2, piAgePb6U8_8col, , _
     piPb46col > 0 And pi208PkOrder > 0 And pbHasU And pbHasTh, piAgePb6U8_8ecol, psPm1sig
  ColInc Col, "204corr|207Pb|/206Pb|Age", "AgePb76_4", 2, piaAgePb76_4Col(0), 0, _
     piPb76col > 0, piaAgePb76_4eCol(0), psPm1sig
  ColInc Col, "204corr|208Pb|/232Th|Age", "AgePb8Th2_4", 2, piAgePb8Th2_4col, , _
     pi208PkOrder > 0 And pbHasTh, piAgePb8Th2_4ecol, psPm1sig

  If foUser("Calc7corrPbThages") Or (pbTh And piStdCorrType = 1) Then  ' 09/06/22 -- modified
    ColInc Col, "207corr|208Pb|/232Th|Age", "AgePb8Th2_7", 2, piAgePb8Th2_7col, , _
       pi208PkOrder > 0 And pbHasTh And pi207PkOrder > 0, piAgePb8Th2_7ecol, psPm1sig
  End If

  If piAgePb6U8_8col > 0 And pbCalc8corrConcPlotRats Then
    ColInc Col, "208corr|207Pb|/206Pb|Age", "AgePb76_8", 2, piAgePb76_8col, , True, _
                piAgePb76_8ecol, pscPpe
  End If

  ColInc Col, "%|Dis-|cor-|dant", "Discord", 1, piDiscordCol, , pbHasU And piPb76col > 0
  ColInc Col, "7corr|206*|/238", "Pb6U8_7", 2, piPb6U8_7col, , _
       piAgePb6U8_7col > 0, piPb6U8_7ecol, pscPpe    ' 09/06/22 added
  ColInc Col, "4corr|208*|/232", "Pb8Th2_4", 2, _
     piPb8Th2_4col, , piPb46col > 0 And piPb86col > 0 And pbHasTh, piPb8Th2_4eCol, pscPpe
  ColInc Col, "7corr|208*|/232", "Pb8Th2_7", 2, piPb8Th2_7col, , _
       piAgePb8Th2_7col > 0, piPb8Th2_7ecol, pscPpe  ' 09/06/22 added

  piLastVisibleCol = Col

  If piPb6U8_totCol > 0 Then
    ColInc Col, "Total|238|/206", "U8Pb6_tot", 2, piU8Pb6_totCol, , piPb76col > 0, piU8Pb6_TotEcol, pscPpe

    If piPb76col > 0 Then
      ColInc Col, "Total|207|/206", "Pb76_tot", 2, piPb76_totCol, , piPb76col > 0, piPb76_totEcol, pscPpe
    End If

    If piPb46col > 0 Then

      ColInc Col, "4corr|238|/206*", "U8Pb6_4", 2, piU8Pb6_4col, , True, piU8Pb6_4ecol, pscPpe

      If piPb76col > 0 Then
        ColInc Col, "4corr|207*|/206*", "Pb76_4", 2, piPb76_4col, , piPb76col > 0, piPb76_4eCol, pscPpe
        ColInc Col, "4corr|207*|/235", "Pb7U5_4", 2, piPb7U5_4col, , piPb76col > 0, piPb7U5_4ecol, pscPpe
      End If

      ColInc Col, "4corr|206*|/238", "Pb6U8_4", 2, piPb6U8_4col, , True, piPb6U8_4ecol, pscPpe
      ColInc Col, "err|corr", "Pb7U5Pb6U8_4rho", 1, piPb7U5Pb6U8_4rhoCol, , piPb7U5_4col > 0 And piPb6U8_4col > 0
      piLastVisibleCol = Col
    End If

    ' 09/06/14 -- added
    If piAgePb6U8_8col > 0 And pbCalc8corrConcPlotRats Then
      ColInc Col, "8corr|238|/206*", "U8Pb6_8", 2, piU8Pb6_8col, , True, piU8Pb6_8ecol, pscPpe
      If piPb76col > 0 Then
        ColInc Col, "8corr|207*|/206*", "Pb76_8", 2, piPb76_8col, , True, piPb76_8ecol, pscPpe
        ColInc Col, "8corr|207*|/235", "Pb7U5_8", 2, piPb7U5_8col, , True, piPb7U5_8ecol, pscPpe
        ColInc Col, "8corr|206*|/238", "Pb6U8_8", 2, piPb6U8_8col, , True, piPb6U8_8ecol, pscPpe
        ColInc Col, "err.|corr.", "Pb7U5Pb6U8_8rho", 1, piPb7U5Pb6U8_8rhoCol
      End If

    End If

  End If

  piLastCol = Col  ' Rightmost visible column
Next i

FromCol = Col
End Sub

Sub PlaceUPbHeaders(HdrRw&, Std As Boolean)
' Place U/Pb Task column headers on the Standard or Sample sheet.
' 09/07/09 -- Put in code for columns 7cor46, 8cor46, 4corcom6, 7corcom6,
'             8corcom6, 4corcom8, 7corcom8,4cor86, 7cor86, 4corppm6,
'             7corppm6, 8corppm6, 4corppm8, 7corppm8

Dim L As Boolean
Dim t$, pu$, pe$, sig1$, t1$, t2$, t3$, t4$, t5$, t6$, t7$
Dim i%, j%, k%, m%, g%, ArrHdrCt%, DpNum%, tmpCol%
Dim Rnom#
Dim ArrHdrs As Variant

sig1 = fsVertToLF("1sigma|err"): pe = fsVertToLF(pscPpe): L = True
pu = psaPDeleRat(piU1Th2)
CFs HdrRw, piNameCol, "Spot Name": CFs HdrRw, piDateTimeCol, "Date/Time"
CFs HdrRw, piHoursCol, "Hours", , "Elapsed since first spot of this session"

If pbXMLfile Then
  CFs HdrRw, piStageXcol, "stage|X", L
  CFs HdrRw, piStageYcol, "stage|Y", L
  CFs HdrRw, piStageZcol, "stage|Z", L
  CFs HdrRw, piQt1yCol, "Qt1y"
  CFs HdrRw, piQt1Zcol, "Qt1z"
  CFs HdrRw, piPrimaryBeamCol, "Primary|beam|(na)", L
End If

PlaceHdr HdrRw, piBkrdCtsCol, "Bkrdcts"
PlaceHdr HdrRw, piPb204ctsCol, "Pb204cts"
PlaceHdr HdrRw, piPb206ctsCol, "Pb206cts"

With puTask

  For i = 1 To .iNpeaks
    Rnom = .daNominal(i)
    If Rnom <> 204 And Rnom <> 206 Then
      CFs HdrRw, piaCPScol(i), psaCPScolHdr(i), L
    End If
  Next i

  PlaceHdr HdrRw, piPb46col, "pb46"
  PlaceHdr HdrRw, piPb76col, "pb76"
  CFs HdrRw, piPb46eCol, pe, L
  CFs HdrRw, piPb76eCol, pe, L
  PlaceHdr HdrRw, piPb86col, "pb86"
  CFs HdrRw, piPb86eCol, pe, L
  k = fvMax(piPb46col, piPb76col, piPb86col)

  For i = 1 To .iNrats
    If piaIsoRatCol(i) > (1 + k) Then
      CFs HdrRw, piaIsoRatCol(i), fsRatioHdrStr(.daNmDmIso(1, i), .daNmDmIso(2, i)), L
      CFs HdrRw, piaIsoRatEcol(i), pe, L
    End If
  Next i

  For i = 1 To .iNeqns

    If piaEqCol(Std, i) > 0 Then
      m = 1
      With .uaSwitches(i)
        ParseLine puTask.saEqnNames(i), ArrHdrs, ArrHdrCt, "||"

        If Not (.SC And .LA) And Not ((.ST And Not Std) Or (.SA And Std)) Then
          RangeNumFor "@", HdrRw, piaEqCol(Std, i)
          CFs HdrRw, piaEqCol(Std, i), ArrHdrs(1), True
          CFs HdrRw, piaEqEcol(Std, i), pe
          m = 2
        ElseIf Not .LA And piaEqCol(Std, i) > 0 And ArrHdrCt = 0 Then
          frSr(1, piaEqCol(Std, i), , piaEqCol(Std, i)).ColumnWidth = 0.05
        End If

        For j = m To fvMin(ArrHdrCt, .ArrNcols)
          CFs HdrRw, j - 1 + piaEqCol(Std, i), ArrHdrs(j), True
        Next j

      End With
    End If

  Next i

End With


For i = 1 To piNumDauPar

  If pbCanDriftCorr Then

    If Std Then
      CFs HdrRw, piLowessHrsCol, psaLowessColHdrs(1), True
      CFs HdrRw, piLowessDeltaPcol, psaLowessColHdrs(2), True
      CFs HdrRw, piLowessDeltaPcol + 1, psaLowessColHdrs(3), True
      CFs HdrRw, piLowessMeasCol, psaLowessColHdrs(4), True
      CFs HdrRw, piLowessMeasCol + 1, psaLowessColHdrs(5), True
    Else
      CFs HdrRw, piUnDriftCorrConstCol, "Raw|calib|const", True
    End If

  End If

  If Std And piStdCorrType >= 0 And piStdCorrType <= 2 And piPb46col > 0 Then
    t = fsS(4 - 3 * (piStdCorrType = 1) - 4 * (piStdCorrType = 2)) & "-corr|"
  Else
    t = ""
  End If

  t = t & IIf((pbU And i = 1) Or (pbTh And i = 2), "206Pb|/238U", "208Pb|/232Th")
  t = t & "|calibr.|const"
  psaUThPbConstColNames(Std, i) = t
  CFs HdrRw, IIf(Std, piaSacol(i), piaAcol(i)), t, L
  CFs HdrRw, IIf(Std, piaSaEcol(i), piaAeCol(i)), pe, L
Next i

  ' 09/06/16 -- added
If Not Std Then
  PlaceHdr HdrRw, piPb46_7col, "Pb46_7"
  PlaceHdr HdrRw, piPb46_8col, "Pb46_8"
End If

t3 = "Radiogenic 208Pb/206Pb corrected "

If piNumDauPar = 1 And piStdCorrType <> 2 Then
  t1 = fsS(Choose(1 + piStdCorrType, 4, 7)) & "-corr|"
  t2 = t3 & "using " & Choose(piStdCorrType + 1, "204", "207") & "Pb"
ElseIf piNumDauPar = 2 Then
  t1 = ""
  t2 = t3 & "from assigned age and measured 232Th/238U"
Else
  t1 = "": t2 = ""
End If

If Std Then ' 09/07/02 -- added
  PlaceHdr HdrRw, piStdCom6_4col, "StdCom6_4"
  PlaceHdr HdrRw, piStdCom6_7col, "StdCom6_7"
  PlaceHdr HdrRw, piStdCom6_8col, "StdCom6_8"
  PlaceHdr HdrRw, piStdCom8_4col, "StdCom8_4"
  PlaceHdr HdrRw, piStdCom8_7col, "StdCom8_7"
  PlaceHdr HdrRw, piStdRadPb86col, "StdRadPb86", t2
  CFs HdrRw, piStdRadPb86ecol, pe, L
End If

t = "(" & fsS(-Std) & ")"
PlaceHdr HdrRw, piaPpmUcol(-Std), "PpmU" & t
PlaceHdr HdrRw, piaPpmThcol(-Std), "PpmTh" & t
PlaceHdr HdrRw, piaTh2U8col(-Std), "Th2U8" & t
CFs HdrRw, piaTh2U8ecol(-Std), pe, L

If Std Then

  For DpNum = 1 To piNumDauPar
    CFs HdrRw, piaSageCol(DpNum), fsVertToLF(Choose(DpNum, "", "alt|") & "Age|(Ma)"), L
    CFs HdrRw, piaSageEcol(DpNum), psPm1sig
  Next DpNum

  t2 = "204Pb|/206Pb|(fr. 20"
  t3 = "Calculated assuming "
  t4 = "-corr|206Pb|/238U|const.|delta%"

  For j = 7 To 8
    k = piaOverCts4Col(j): t = "(" & fsS(j) & ")"

    If k > 0 Then
      t5 = IIf(j = 7, pscR6875, pscR6882)
      m = piaOverCts46Col(j): g = piacorrAdeltCol(j)

      PlaceHdr HdrRw, k, "OverCts4" & t, t3 & t5
      PlaceHdr HdrRw, m, "OverCts46" & t, t3 & t5
      PlaceHdr HdrRw, g, "corrAdelt" & t, "Difference between 20" & fsS(j) & "-corr. and 204-corr. constant"

      CFs HdrRw, piaOverCts46eCol(j), pscPpe, True
    End If

  Next j

  For DpNum = 1 To piNumDauPar
    j = IIf((pbU And DpNum = 1) Or (pbTh And DpNum = 2), 1, 2)

    CFs HdrRw, piaStdUnCorrAcol(DpNum), fsVertToLF("Uncorr|" & _
         psaPDeleRat(j) & "|const"), L, _
      "Uncorrected for common Pb"
    CFs HdrRw, piaStdUnCorrAerCol(DpNum), fsVertToLF(pscPpe), L
  Next DpNum

  ' 09/07/02 -- added
  If piPb76col > 0 Then

    If piPb46col > 0 Then
      PlaceHdr HdrRw, piStdPb76_4Col, "StdPb76_4"
      CFs HdrRw, piStdPb76_4eCol, pe
    End If

    PlaceHdr HdrRw, piaAgePb76_4Col(1), "AgePb76_4(1)"


    CFs HdrRw, piaAgePb76_4Col(1), IIf(piPb46col > 0, "4-corr", "uncorr") _
        & "|207Pb|/206Pb|age", L
    CFs HdrRw, piaAgePb76_4eCol(1), psPm1sig
  End If

Else ' 09/07/02 -- added
  PlaceHdr HdrRw, piCom6_4col, "Com6_4"  ' col-hdrs for %common daughter-isot
  PlaceHdr HdrRw, piCom6_7col, "Com6_7"  '   (sample sheet)
  PlaceHdr HdrRw, piCom6_8col, "Com6_8"
  PlaceHdr HdrRw, piCom8_4col, "Com8_4"
  PlaceHdr HdrRw, piCom8_7col, "Com8_7"
  PlaceHdr HdrRw, piPb86_4col, "Pb86_4"
  CFs HdrRw, piPb86_4ecol, pe, L
  PlaceHdr HdrRw, piPb86_7col, "Pb86_7"
  CFs HdrRw, piPb86_7ecol, pe, L
  PlaceHdr HdrRw, piaRadDauCol_4(6), "raddau_4(6)" ' col-hdrs for ppm rad daughter-isot
  PlaceHdr HdrRw, piaRadDauCol_7(6), "raddau_7(6)"
  PlaceHdr HdrRw, piaRadDauCol_8, "raddau_8"
  PlaceHdr HdrRw, piaRadDauCol_4(8), "raddau_4(8)"
  PlaceHdr HdrRw, piaRadDauCol_7(8), "raddau_7(8)"
  PlaceHdr HdrRw, piPb6U8_totCol, "Pb6U8_tot"
  CFs HdrRw, piPb6U8_totEcol, pe, L

  If piPb86col > 0 And pbHasTh Then
    PlaceHdr HdrRw, piPb8Th2_totCol, "Pb8Th2_tot"
    CFs HdrRw, piPb8Th2_totEcol, pe, L
  End If

  PlaceHdr HdrRw, piAgePb6U8_4col, "AgePb6U8_4"
  CFs HdrRw, piAgePb6U8_4ecol, sig1, L
  PlaceHdr HdrRw, piAgePb6U8_7col, "AgePb6U8_7"
  CFs HdrRw, piAgePb6U8_7ecol, sig1, L
  PlaceHdr HdrRw, piAgePb6U8_8col, "AgePb6U8_8"
  CFs HdrRw, piAgePb6U8_8ecol, sig1, L
  PlaceHdr HdrRw, piaAgePb76_4Col(0), "AgePb76_4(0)"
  CFs HdrRw, piaAgePb76_4eCol(0), sig1, L
  PlaceHdr HdrRw, piAgePb8Th2_4col, "AgePb8Th2_4"
  CFs HdrRw, piAgePb8Th2_4ecol, sig1, L
  PlaceHdr HdrRw, piAgePb8Th2_7col, "AgePb8Th2_7"
  CFs HdrRw, piAgePb8Th2_7ecol, sig1, L
  PlaceHdr HdrRw, piPb6U8_7col, "Pb6U8_7"           ' /
  CFs HdrRw, piPb6U8_7ecol, pe, L                   '| ' 09/06/22 added
  PlaceHdr HdrRw, piPb8Th2_7col, "Pb8Th2_7"         '|
  CFs HdrRw, piPb8Th2_7ecol, pe, L                  ' \
  PlaceHdr HdrRw, piAgePb76_8col, "AgePb76_8"
  CFs HdrRw, piAgePb76_8ecol, sig1, L
  PlaceHdr HdrRw, piDiscordCol, "Discord"
  PlaceHdr HdrRw, piPb8Th2_4col, "Pb8Th2_4"
  CFs HdrRw, piPb8Th2_4eCol, pe, L
  PlaceHdr HdrRw, piU8Pb6_totCol, "U8Pb6_tot"
  CFs HdrRw, piU8Pb6_TotEcol, pe, L
  PlaceHdr HdrRw, piPb76_totCol, "Pb76_tot"
  CFs HdrRw, piPb76_totEcol, pe, L
  PlaceHdr HdrRw, piU8Pb6_4col, "U8Pb6_4"
  CFs HdrRw, piU8Pb6_4ecol, pe, L
  PlaceHdr HdrRw, piPb76_4col, "Pb76_4"
  CFs HdrRw, piPb76_4eCol, pe, L
  PlaceHdr HdrRw, piPb7U5_4col, "Pb7U5_4"
  CFs HdrRw, piPb7U5_4ecol, pe, L
  PlaceHdr HdrRw, piPb6U8_4col, "Pb6U8_4"
  CFs HdrRw, piPb6U8_4ecol, pe, L
  PlaceHdr HdrRw, piPb7U5Pb6U8_4rhoCol, "Pb7U5Pb6U8_4rho"
  PlaceHdr HdrRw, piU8Pb6_8col, "U8Pb6_8"
  CFs HdrRw, piU8Pb6_8ecol, pe, L
  PlaceHdr HdrRw, piPb76_8col, "Pb76_8"
  CFs HdrRw, piPb76_8ecol, pe, L
  PlaceHdr HdrRw, piPb7U5_8col, "Pb7U5_8"
  CFs HdrRw, piPb7U5_8ecol, pe, L
  PlaceHdr HdrRw, piPb6U8_8col, "Pb6U8_8"
  CFs HdrRw, piPb6U8_8ecol, pe, L
  PlaceHdr HdrRw, piPb7U5Pb6U8_8rhoCol, "Pb7U5Pb6U8_8rho"
End If

For i = 1 To fiEndCol(HdrRw)
  If InStr(Cells(HdrRw, i), "sigma") Then SigConv HdrRw, i
Next i

With frSr(HdrRw, 1, , IIf(Std, piSlastCol, piLastCol))
  .Borders(xlEdgeBottom).LineStyle = xlDouble
End With
plaFirstDatRw(-Std) = 1 + HdrRw
End Sub

Sub MakeTable() ' Make a publication-ready table from a Grouped Sample worksheet;
                '  U/Pb Geochron only.
Dim Na$, CF$, f$, g$, FontName$
Dim i%, j%, k%, c%, Col%, EndColm%
Dim HdrRw&, LastRw&, v#
Dim StdE As Range
Dim ShtIn As Object, Shp As Shape, cc As Range
Dim Ce As Range, Sht As Worksheet

g$ = "You must start from a Squid Grouped-sample sheet"
If ActiveSheet.Type <> xlWorksheet Then Exit Sub
Set Sht = ActiveSheet
Set ShtIn = Sht
On Error Resume Next
Na$ = Sht.Name
On Error GoTo 0
If Na$ = "" Then End
If Na = pscSamShtNa Or Na$ = pscStdShtNa Or Sht.Type <> xlWorksheet Then MsgBox g$, , pscSq: End
c = 0
If InStr(Na, "Table") Then MsgBox "Sheet is already a table.": End
HdrRw = flHeaderRow(False)

If HdrRw = 0 Or LCase(Cells(1, 1)) <> "errors are 1s unless otherwise specified" Then
  MsgBox g, , pscSq
  End
End If

NoUpdate
StatBar "Starting table": FontName$ = "Lucida Console"
Freeze False

Sht.Copy After:=Sht
Na$ = Na$ & " Table"
Sheetname Na$
Set Sht = ActiveSheet: Sht.Name = Na$
On Error GoTo 0
FreezeValues
ActiveWindow.SplitRow = 0
ActiveWindow.SplitColumn = 0
plHdrRw = flHeaderRow(0)
StatBar "Deleting unused columns"

For i = 1 To 2
  FindDelCol Choose(i, "SqidNum", "SqidEr"), 0, plHdrRw, 2, , 50
Next i

For Each Shp In Sht.Shapes: Shp.Delete: Next
FindStr "%com206", , c, plHdrRw

If c > 0 Then
  CF$ = fsQq("[<0]$-- $;0.00")
  RangeNumFor CF$, 1 + plHdrRw, c, plaLastDatRw(0)
  FindDelCol "%com208", 0, plHdrRw, 2, , 50
End If

For i = 1 To 5
  j = Choose(i, 196, 248, 254, 264, 270)
  FindDelCol "/" & fsS(j), 1, plHdrRw, 2, , 50
  FindDelCol fsS(j) & "/", 1, plHdrRw, 2, , 50
Next i

For i = 1 To 2
  FindDelCol "calibr.const", 1, plHdrRw, 2, , 50
Next i

FindDelCol "total206Pb/238U", 1, plHdrRw, 2, , 50
FindDelCol "total208Pb/232Th", 1, plHdrRw, 2, , 50
FindDelCol "total238/206", 1, plHdrRw, 2, , 50
FindDelCol "total207/206", 1, plHdrRw, 2, , 50
FindDelCol "S-KcomPb", 0, plHdrRw, 2, , 50

For i = 1 To 3
  FindDelCol "C-Pb", 0, plHdrRw, 2, , 50
Next i

FindDelCol "208Pb*/206Pb*", 1, plHdrRw, 2, , 50
FindDelCol "flags", 0, plHdrRw, 2, , 50
FindDelCol "4corr208*/232", 1, plHdrRw, 2, , 50
FindDelCol "7corr206*/238", 1, plHdrRw, 2, , 50
FindDelCol "8corr206*/238", 1, plHdrRw, 2, , 50
FindDelCol "date/time", 0, plHdrRw, 1, , 9
FindDelCol "hours", 0, plHdrRw, 1, , 9
FindDelCol "#rej", 0, plHdrRw, 1, , 19
FindDelCol "SqidNum", 0, plHdrRw, 1, , 99
FindDelCol "SqidEr", 0, plHdrRw, 1, , 99
FindDelCol "SqidEr", 0, plHdrRw, 1, , 99
FindDelCol "StageX", 0, plHdrRw, 1, , 99
FindDelCol "StageY", 0, plHdrRw, 1, , 99
FindDelCol "StageZ", 0, plHdrRw, 1, , 99
FindDelCol "Qt1y", 0, plHdrRw, 1, , 99
FindDelCol "Qt1z", 0, plHdrRw, 1, , 99
FindDelCol "Primary beam (na)", 0, plHdrRw, 1, , 99

For i = 1 To 3
  FindDelCol "cts/sec", 0, plHdrRw, 1, , 19
Next i

StatBar "Formatting fonts & colors"
Fonts Cells, , , , 0, 0, , 11, 0, , , , FontName, , 0
With Cells
  .Interior.ColorIndex = xlNone
  For k = 1 To 7: .Borders(k).LineStyle = xlNone: Next k
End With
plHdrRw = 2
EndColm = fiEndCol(2)
StatBar "Making headers"
Cells(plHdrRw, 1) = " Spot"

For i = 3 To EndColm
  If Right$(Cells(plHdrRw, i - 1), 3) = "Age" And _
   Right$(Cells(plHdrRw, i), 3) = "err" Then
    RangeNumFor pscPm & Cells(1 + plHdrRw, i).NumberFormat, _
                1 + plHdrRw, i, plaLastDatRw(0)
    Cells(plHdrRw, i) = ""
    HA xlLeft, , i
    HA xlRight, , i - 1
    With frSr(plHdrRw, i, , i - 1)
      .Merge: .HorizontalAlignment = xlCenter
    End With
    i = i + 1
  End If
Next i

Rows(plHdrRw).Insert: Rows(1).Delete

For j = 1 To EndColm
  Set cc = Cells(2, j)
  Subst cc, "204corr", "(1)", "4corr", "(1)", "207corr", "(2)"
  Subst cc, "208corr", "(3)", "8corr", "(3)"
Next j

With frSr(plHdrRw, 1, , EndColm)
  .Borders(xlBottom).LineStyle = xlDouble
End With

For i = 2 To EndColm
  f = Cells(plHdrRw, i).Text
  Subst f, fsVertToLF(pscPpe), pscPm & "%", fsVertToLF("1s|"), , "206*", "206Pb*"
  Subst f, "207*", "207Pb*", "208*", "208Pb*", "204", "204Pb"
  Subst f, "206", "206Pb", "207", "207Pb", "208", "208Pb"
  Subst f, "232", "232Th", "235", "235U", "238", "238U"
  Subst f, "PbPb", "Pb", "ThTh", "Th", "UU", "U"
  Cells(plHdrRw, i) = f$
  j = InStr(f$, "*")
  With Cells(plHdrRw, i)

    If j > 0 Then
      .Characters(j, 1).Font.Superscript = True
      k = InStr(Mid$(f$, j + 1), "*")

      If k > 0 Then
        .Characters(j + k, 1).Font.Superscript = True
      End If

    End If

  End With
Next i

For i = 206 To 208 Step 2
  FindStr "%com" & fsS(i), , c, plHdrRw, , 20

  If c Then
    j = InStr(Cells(plHdrRw, c), vbLf)
    Cells(plHdrRw, c) = fsVertToLF("%|" & fsS(i) & "Pbc")
    Cells(plHdrRw, c).Characters(Len(Cells(plHdrRw, c)), 1).Font.Subscript = True
  End If

  FindStr "ppmRad" & fsS(i), 0, c, plHdrRw, , 20

  If c Then
    Cells(plHdrRw, c) = fsVertToLF("ppm|" & fsS(i) & "*")
    RangeNumFor "@", 1 + plHdrRw, c, plaLastDatRw(0)

    For j = 1 + plHdrRw To plaLastDatRw(0)
      Set Ce = Cells(j, c)

      If Not IsEmpty(Ce) And IsNumeric(Ce) Then
        v = Drnd(Cells(j, c), 3)
        ErrFor v, 0.1 * v
        Cells(j, c) = v
      End If

    Next j

    Fonts 1 + plHdrRw, c, plaLastDatRw(0), , , , xlRight
  End If

Next i

StatBar "Adjusting cell height/width"
For k = 1 To plHdrRw - 2: Rows(k).Delete: Next k
frSr(1, , 2).Insert Shift:=xlDown
plHdrRw = flHeaderRow(0)
EndColm = fiEndCol(plHdrRw)
Range(Rows(plaLastDatRw(0) + 1), Rows(plaLastDatRw(0) + 40)).Delete
' Pad spot names with a space
For i = 1 + plHdrRw To plaLastDatRw(0): Cells(i, 1) = Cells(i, 1) & " ": Next i
With frSr(1 + plHdrRw, , plaLastDatRw(0)).Font
  .Name = "Lucida Console": .Size = 10
End With
ActiveWindow.Zoom = 100
With Cells
  .ColumnWidth = 20: .Columns.AutoFit
  .RowHeight = 50:   .Rows.AutoFit
End With
HA xlCenter, plHdrRw, 1
plHdrRw = flHeaderRow(0)
EndColm = fiEndCol(plHdrRw)
Rows(plHdrRw - 1).Font.Size = 10 ' Footnote indicators
Rows(plHdrRw).Select
' SuperIso  From Isoplot
Rows(plHdrRw).RowHeight = Rows(plHdrRw).RowHeight + 5
With Rows(1 + plHdrRw): .AutoFit: .RowHeight = 4 + .RowHeight: End With
Cells.Columns.AutoFit
StatBar "Adding footnotes"
f$ = Space(3)
g$ = "Errors are 1-sigma; Pbc and Pb* indicate the common and radiogenic portions, respectively."

With Cells(1 + plaLastDatRw(0), 1)
  .Formula = f$ & g$
  .Characters(26, 1).Font.Subscript = True
  .Characters(34, 1).Font.Superscript = True
End With

Set phStdSht = Sheets(pscStdShtNa)
Set StdE = phStdSht.[wtdmeanaperr1]: j = 0
On Error Resume Next

If Not IsEmpty(StdE) And IsNumeric(StdE) Then

  If StdE <> 0 Then
    j = 1
    Cells(1 + j + plaLastDatRw(0), 1) = f$ & "Error in Standard calibration was " & _
    foAp.Fixed(StdE, 2) & "%" & " (not included in above errors" _
    & " but required when comparing data from different mounts)."
  End If

End If

On Error GoTo 0
g$ = ") Common Pb corrected "
Cells(2 + j + plaLastDatRw(0), 1) = f$ & "(1" & g$ & "using measured 204Pb."
Cells(3 + j + plaLastDatRw(0), 1) = f$ & "(2" & g$ & "by assuming " & pscR6875
Cells(4 + j + plaLastDatRw(0), 1) = f$ & "(3" & g$ & "by assuming " & pscR6882
Fonts 1 + plaLastDatRw(0), 1, 4 + j + plaLastDatRw(0), EndColm, , , , 10
BorderLine xlBottom, 1, plaLastDatRw(0), 1, , EndColm      ' Bottom of data rows
BorderLine xlBottom, 1, 4 + j + plaLastDatRw(0), 1, , EndColm ' Bottom of footnotes
BorderLine xlTop, 2, plHdrRw - 1, 1, , EndColm         ' Top of table
BorderLine xlBottom, xlNone, plHdrRw, 1, , EndColm     ' Erase double-underlines below headers
BorderLine xlBottom, 1, plHdrRw, 1, , EndColm          ' Replace with single underline.
BorderLine xlLeft, 1, plHdrRw - 1, 1, 4 + j + plaLastDatRw(0)

For i = 1 To EndColm
  If Cells(plHdrRw, i + 1) = "err" Then Cells(plHdrRw, i + 1) = pscPm

  If InStr(Cells(plHdrRw, i + 1).Text, pscPm) = 0 And InStr(Cells(plHdrRw, i).Text, "Age") = 0 _
      And InStr(Cells(1 + plHdrRw, i + 1).Text, pscPm) = 0 Then
    BorderLine xlRight, 1, plHdrRw - 1, i, plaLastDatRw(0) - (4 + j) * (i = EndColm)
  End If

Next i

HA xlCenter, , EndColm
With Rows(1 + plaLastDatRw(0)): .RowHeight = .RowHeight + 3: End With ' Clearance at start of footnotess
frSr(2 + plaLastDatRw(0), 1, 4 + j + plaLastDatRw(0)).Select
' SuperIso
Selection.RowHeight = Cells(2 + plaLastDatRw(0), 1).RowHeight + 2 ' For superscript room
Cells(1, 1).Select
With ActiveWindow
  .ScrollRow = 1: .ScrollColumn = 1
End With
ShtIn.Select:  Zoom 75
Sht.Activate:  Zoom 85
WidenCols
GoTo mtDone

SqEr: On Error GoTo 0
If ActiveSheet.Name = Na$ Then DelSheet
MsgBox "Can't create table -- is this an intact SQUID Grouped-sample sheet?", , pscSq

mtDone: On Error GoTo 0
StatBar
End Sub

Sub StdRadiogenicAndAgeCols(FirstRow&, LastRow&)
' Put formulas for Pb/Pb, Pb/U, and Pb/Th ratios & ages in the Standard sheet.
'  (Standard-sheet U/Pb Tasks & output columns only)
Dim SA$, sae$, s$, p$, q$, t$
Dim Rad75$, Rad75e$, Rad68$, Rad68e$, Rad86$, Rad86e$, Rad76$, Rad76e$
Dim PeCol%, i%, j%, Acol%, f&, L&
Dim CwidthHdr!, CwidthNum!, Cwidth!
Dim rw1 As Range, rw2 As Range, r3 As Range, r4 As Range

phStdSht.Activate
If pbTh And piNumDauPar = 1 Then j = 1 Else j = piU1Th2
SA = " sA(" & fsS(j) & ") "
f = FirstRow: L = LastRow
piLastCol = fiEndCol(flHeaderRow(True))
PeCol = piLastCol
foAp.Calculate

If piStdPb7U5_4col > 0 Then
  If pi204PkOrder > 0 And pi207PkOrder > 0 And pbHasU Then
    CFs plHdrRw, piStdPb7U5_4col, "207*|/235", -1
    CFs plHdrRw, piStdPb7U5_4eCol, pscPpe, -1
    CFs plHdrRw, piStdPb6U8_4col, "206*|/238", -1
    CFs plHdrRw, piStdPb6U8_4eCol, pscPpe, -1
    Rad75 = fsCellAddr(f, piStdPb7U5_4col): Rad75e = fsCellAddr(f, piStdPb7U5_4eCol)
    Rad68 = fsCellAddr(f, piStdPb6U8_4col): Rad68e = fsCellAddr(f, piStdPb6U8_4eCol)
    CFs plHdrRw, piStdPb7U5Pb6U8_4rhoCol, "err|corr", -1
  End If
End If

If piPb76col And pi204PkOrder > 0 Then
  Rad76 = fsCellAddr(f, piStdPb76_4Col)
  Rad76e = fsCellAddr(f, piStdPb76_4eCol)
End If

sae = fsCellAddr(f, piaSaEcol(1))

If piStdPb6U8_4col > 0 Then
  If pbTh And piNumDauPar = 1 Then q = "StdThPbRatio" Else q = "StdUPbRatio"
  s = "=" & SA & "/" & pscStdShtNot & "WtdMeanA1*" & pscStdShtNot & q      ' 206*/238
  p = "=1/" & Rad68        ' 238/206*
  PlaceFormulae s, f, piStdPb6U8_4col, L
  PlaceFormulae "= sAe(1) ", f, 1 + piStdPb6U8_4col, L
End If

If piStdPb7U5_4col > 0 Then

  If piStdPb76_4Col > 0 Then
    t = "= StdPb76_4 *" & pscUra & "* StdPb6U8_4 "
    PlaceFormulae t, f, piStdPb7U5_4col, L
  End If

  q = "=sqrt( StdPb6U8_4e ^2+ StdPb76_4e ^2)"
  PlaceFormulae q, f, piStdPb7U5_4eCol, L    ' 7*/5 %err"
  PlaceFormulae "= StdPb6U8_4e / StdPb7U5_4e ", f, 2 + piStdPb6U8_4col, L
End If

If piStdPb6U8_4col > 0 And piStdPb7U5_4col > 0 Then

  For i = 1 + plHdrRw To LastRow  '  move to later!!!!!!!!!!!
    Set rw1 = Cells(i, piStdPb7U5_4col): Set r3 = Cells(i, 2 + piStdPb7U5_4col)
    Set rw2 = Cells(i, piStdPb6U8_4col): Set r4 = Cells(i, 2 + piStdPb6U8_4col)
    With frSr(i, piStdPb7U5_4col, , 2 + piStdPb6U8_4col).Font

      If IsError(rw1) Or IsError(rw2) Then
        .Strikethrough = True
      ElseIf IsNumeric(rw1) And IsNumeric(rw2) And _
             IsNumeric(r3) And IsNumeric(r4) Then
        If rw2 < 0.0001 Or rw1 < 0.00001 Or r4 >= 1 Or r4 <= 0 Then
          .Strikethrough = True
        End If
      End If

    End With
  Next i

End If

For i = 0 To 1
  If piStdPb7U5_4col > 0 Then Nformat piStdPb7U5_4col + i, i, True

  If piStdPb6U8_4col > 0 Then
    Nformat piStdPb6U8_4col + i, i, True
  End If

  If piStdPb7U5_4col > 0 And piStdPb6U8_4col > 0 And i = 1 Then Nformat 2 + piStdPb6U8_4col, 0, True
Next i

With puTask

  For i = 1 To .iNeqns
    With .uaSwitches(i)

      If .Ar And .ArrNcols > 1 Then

        For j = 2 To .ArrNcols
          Acol = piaEqCol(True, i) + j - 1
          Columns(Acol).AutoFit
          Cells(1, 256) = Cells(HdrRow, Acol)
          Columns(256).AutoFit
          CwidthHdr = Columns(256).ColumnWidth
          Cells(1, 256) = ""
          CwidthNum = Columns(Acol).ColumnWidth
          Cwidth = CwidthHdr
          If CwidthNum > CwidthHdr Then Cwidth = fvMin(9, CwidthNum)
          ColWidth Cwidth, Acol
        Next j

      End If

    End With
  Next i

End With
BorderLine xlBottom, 2, plHdrRw, 1, , fiEndCol(plHdrRw)
End Sub

Sub SamRadiogenicAndAgeCols(FirstRow&, LastRow&)
' Put formulas for Pb/Pb, Pb/U, and Pb/Th ratios & ages in the Sample sheet.
'  (Sample-sheet U/Pb Tasks & output columns only)
Dim rw1$, rw2$, r3$, r4$, HdrRw$, tmp$, FinalTerm$
Dim radd6$, radd8$, Alpha$, Beta$, Gamma$, NetAlpha$, NetBeta$, NetGamma$
Dim t1$, t2$, t3$, t4$, t5$, t6$, t7$, t8$, t9$, t10$
Dim Col1$, d1$, d3$, d4$, d5$
Dim i%, j%, k%, m%, f%, L%, p%, q%, MinC%, MaxC%, Hrow&, Pb64#

Alpha = "1/ Pb46 ": Beta = " Pb76 / Pb46 ": Gamma = " Pb86 / Pb46 "
radd6 = "(1-" & psaC64(0) & "* Pb46 )"         ' (1 - alpha0/alpha)
radd8 = "(1-" & psaC84(0) & "* Pb46 / Pb86 )"  ' (1 - gamma0/gamma)
NetAlpha = "(" & Alpha & "-" & psaC64(0) & ")" ' (alpha-alpha0)
NetBeta = "(" & Beta & "-" & psaC74(0) & ")"   ' (beta-beta0)
NetGamma = "(" & Gamma & "-" & psaC84(0) & ")" ' (gamma-gamma0)
f = FirstRow: L = LastRow
Hrow = flHeaderRow(False)
StatBar "Finishing (isotope ratios)"
phSamSht.Activate
foAp.Calculate
' Calculate 206t/238, 208t/232, 238/206t, ppm206*, ppm208*, 206*/238,
'   238/206*, Age4corr (or uncorr), 207/206t, 207*/206*, Pb76age (tot or rad),
'   207*/235, 206*/238, discordance, age7corr, 206*/232, 208/232age,
'   Age8corr, 232/238err (from formula),

For i = 1 To piNumDauPar

  If (i = 1 And pbU) Or (i = 2 And pbTh) Then
    m = piPb6U8_totCol
  ElseIf (i = 1 And pbTh) Or (i = 2 And pbU) Then
    m = piPb8Th2_totCol
  Else
    m = 0
  End If

  If m > 0 Then
    ' 206t/238 or 208t/232
    Col1 = fsS(i)
    rw2 = "WtdMeanA" & Col1
    rw1 = phStdSht.Range(rw2).Address
    r3 = "Std" & IIf(m = piPb6U8_totCol, "U", "Th") & "PbRatio"
    r4 = "Extperra" & Col1
    PlaceFormulae "= A(" & Col1 & ") /" & pscStdShtNot & rw2 & "*" _
                  & pscStdShtNot & r3, f, m, L
    PlaceFormulae "=SQRT( Ae(" & Col1 & ") ^2+" & pscStdShtNot & r4 & _
                  "^2)", f, m + 1, L
  End If

Next i

' 238/206t
If piU8Pb6_totCol > 0 Then
  PlaceFormulae "=1/ Pb6U8_tot ", f, piU8Pb6_totCol, L  ' 238U/206Pbtot
  PlaceFormulae "= Pb6U8_tote ", f, piU8Pb6_TotEcol, L  ' %err
End If

If pbU Then

  If piPb8Th2_totCol > 0 And Not puTask.bDirectAltPD Then
  ' 208t/232
    PlaceFormulae "= Pb6U8_tot * Pb86 / Th2U8(0) ", f, _
                   piPb8Th2_totCol, L
    PlaceFormulae "=sqrt( Pb86e ^2+ Pb6U8_totE ^2+ Th2U8e(0) ^2)", f, _
                   piPb8Th2_totEcol, L
  End If

ElseIf pbTh And piPb6U8_totCol > 9 And Not puTask.bDirectAltPD Then
' t206/238
  PlaceFormulae "= Pb8Th2_tot / Pb86 * th2u8(0) ", f, piPb6U8_totCol, L
  PlaceFormulae "=SQRT( Pb8Th2_tote ^2+ Pb86eCol ^2+ Th2U8e(0) ^2)", f, _
                 piPb6U8_totEcol, L
End If


If piaPpmUcol(0) > 0 Then  ' 09/07/21 -- added
  t1 = "= Pb6U8_tot * ppmU(0) *0.859*(1- Pb"
  t2 = "46"
  t3 = "46_7"
  t4 = "46_8"
  t5 = " )"
  t6 = "_7)"
  t7 = " *" & psaC64(-pbStd) & ")"

  For p = 1 To 5 ' Put ppm rad daughter isotope values

    Select Case p
      Case 1
        q = piaRadDauCol_4(6)
        FinalTerm = t1 & t2 & t7
      Case 2
        q = piaRadDauCol_7(6)
        FinalTerm = t1 & t3 & t7
      Case 3
        q = piaRadDauCol_8
         FinalTerm = t1 & t4 & t7
      Case 4
        q = piaRadDauCol_4(8)
        FinalTerm = "= RadDau_4(6) * Pb86_4 *208/206"
      Case 5
        q = piaRadDauCol_7(8)
        FinalTerm = "= RadDau_7(6) * Pb86_7 *208/206"
    End Select

    If q > 0 Then
      PlaceFormulae FinalTerm, f, q, L
    End If

  Next p

Else ' 09/07/21 -- added
  HideMarkCol 1, piaRadDauCol_4(6)
  HideMarkCol 1, piaRadDauCol_7(6)
  HideMarkCol 1, piaRadDauCol_8
  HideMarkCol 1, piaRadDauCol_4(8)
  HideMarkCol 1, piaRadDauCol_7(8)
End If

'ppm206tot=(206tot/238)*ppmU/AtWtU*(1379/1389)*AtWt206
'         =(206tot/238)*ppmU*0.859
'         =(206tot/238)*ppmTh*AtWt206/AtWt232/(232/238)
'         =(206tot/238)*ppmTh*/(232/238)*0.8876
'ppm208tot=(208tot/232)*ppmTh*AtWt208/AtWt232
'         =(208tot/232)*ppmTh*0.896
'         =(208tot/232)*ppmU/AtWtU*(1379/1389)*(232/238)*AtWtTh
'         =(208tot/232)*ppmU*(232/238)*0.968

If piPb6U8_4col > 0 Then
  t1 = "= Pb6U8_tot *" & radd6
  ' 206Pb*/238U
  PlaceFormulae t1, f, piPb6U8_4col, L
  ' Var(Rad6/8)=Var(Tot6/8) + (Alpha0/(Alpha-Alpha0))^2 * Var(Alpha)
  t1 = "=SQRT( Pb6U8_tote ^2+(" & psaC64(0) & "* Pb46eCol /(1/ Pb46 -" _
       & psaC64(0) & "))^2)"
  PlaceFormulae t1, f, piPb6U8_4ecol, L

  If piU8Pb6_4col > 0 Then
    ' 238U/206Pb*
    PlaceFormulae "=1/ Pb6U8_4 ", f, piU8Pb6_4col, L
    PlaceFormulae "= Pb6U8_4e ", f, piU8Pb6_4ecol, L
  End If

  ' 204-corr 206*238 age
  PlaceFormulae "=LN(1+ Pb6U8_4 )/" & pscLm8, f, piAgePb6U8_4col, L
' 4-corr206/238 age-err
  t1 = fsRa(f, piAgePb6U8_4col)
  d1 = "(" & NetAlpha & "*.01* Pb6U8_totE )^2"
  d3 = "(.01* Pb46eCol *" & psaC64(0) & ")^2"
  d4 = "( Pb6U8_tot * Pb46 )^2"
  d5 = "(1/" & pscLm8 & "/Exp(" & pscLm8 & "* AgePb6U8_4 ))^2"
  tmp = d5 & "*" & d4 & "*(" & d1 & "+" & d3 & ")"
  tmp = "=SQRT(" & tmp & ")"
  PlaceFormulae tmp, f, piAgePb6U8_4ecol, L

ElseIf piPb46col = 0 And piAgePb6U8_4col > 0 Then
' age from ucorr 206/238
  PlaceFormulae "=LN(1+ Pb6U8_tot )/" & pscLm8, f, piAgePb6U8_4col, L
  CFs plHdrRw, piAgePb6U8_4col, "total|206Pb|/238U|age", True
  PlaceFormulae "= Pb6U8_tot /" & pscLm8 & "/(1+ Pb6U8_tot )* Pb6U8_totE /100", _
                f, piAgePb6U8_4ecol, L
End If ' If piPb6U8_4col > 0

If piPb76col > 0 Then

  If piU8Pb6_totCol > 0 Then
' tot 207/206 (T-W)
    PlaceFormulae "= Pb76col ", f, piPb76_totCol, L
    PlaceFormulae "= Pb76eCol ", f, piPb76_totEcol, L
  End If

  If piPb46col > 0 And piU8Pb6_4col > 0 Then
    tmp = "=abs(" & NetBeta & "/" & NetAlpha & ")"
' 207*/206* (T-W)
    PlaceFormulae tmp, f, piPb76_4col, L
' 207*/206* (T-W) %ERR
    t1 = "(( Pb76 -  Pb76_4 )* Pb46e /100/ Pb46 )^2"
    t3 = "( Pb76e / Pb46 /100* Pb76 )^2"
    t5 = NetAlpha
    tmp = "=ABS(SQRT(" & t1 & "+" & t3 & ")/" _
           + t5 & "*100/  Pb76_4 )"
    PlaceFormulae tmp, f, piPb76_4eCol, L

    If piaAgePb76_4Col(0) Then
      r3 = IIf(piPb46col > 0, " Pb76_4", psTwbName & "Pb76")
      r4 = IIf(piPb46col > 0, " Pb76_4e", "Pb76e")
      On Error Resume Next
' Pb-7/6 age - total or rad
      PlaceFormulae "=AgePb76( " & r3 & " )", f, piaAgePb76_4Col(0), L
      PlaceFormulae "=AgeErPb76( " & r3 & " , " & r4 & " ,,,,true)", _
                    f, piaAgePb76_4eCol(0), L   ' Pb-7/6 age-err
      If piPb46col = 0 Then CFs plHdrRw, piaAgePb76_4Col(0), _
                            "total|207Pb|/206Pb|age", True
      On Error GoTo 0
    End If

    If piPb7U5_4col > 0 And piPb6U8_4col > 0 Then
' 207*/235
      tmp = "=  Pb76_4 * Pb6U8_4 *" & fsS(pdcUrat)
      PlaceFormulae tmp, f, piPb7U5_4col, L
      rw1 = fsRa(f, piU8Pb6_4ecol): rw2 = fsRa(f, piU8Pb6_4ecol)
' 207*/235 %err
      PlaceFormulae "=SQRT( Pb6U8_4e ^2+ Pb76_4e ^2)", f, piPb7U5_4ecol, L
' rho
      PlaceFormulae "= Pb6U8_4e / Pb7U5_4e ", f, piPb7U5Pb6U8_4rhoCol, L
    End If

  End If

  If piDiscordCol > 0 And piPb46col > 0 Then
   ' better - difference between obs. 6/8 ratio and 6/8 ratio for the 7/6 age
     PlaceFormulae "=100*(1- Pb6U8_4 /(" & pscEx8 & " AgePb76_4(0) )-1))", _
                    f, piDiscordCol, L
  End If

  If piAgePb6U8_7col > 0 Then
' 207-corr 206/238 age
    On Error Resume Next
    PlaceFormulae "=Age7corr( Pb6U8_tot , Pb76 ," & psaC76(0) & ")", _
                  f, piAgePb6U8_7col, L
    On Error GoTo 0
    rw1 = fsRa(f, piAgePb6U8_7col)
    t1 = "=AgeEr7corr( AgePb6U8_7 , Pb6U8_tot , Pb6U8_tote /100* Pb6U8_tot "
    t2 = ", Pb76 , Pb76e /100* Pb76 ," & psaC76(0) & ",0)"
    PlaceFormulae t1 & t2, f, piAgePb6U8_7ecol, L    ' 207-corrected age-err
  End If

End If ' If piPb76col > 0

If piPb8Th2_totCol > 0 And piPb86col > 0 Then

  If piPb46col > 0 Then
    tmp = "= Pb8Th2_tot *" & radd8
    PlaceFormulae tmp, f, piPb8Th2_4col, L   ' 208Pb*/232Th
    rw1 = Gamma
    tmp = "=sqrt( Pb8Th2_tote ^2+(" & psaC84(0) & "/" & NetGamma & ")^2* Pb46e ^2)"
    ' neglecting the 208/206 error
    PlaceFormulae tmp, f, piPb8Th2_4eCol, L  ' rad 208Pb/232Th %err
  End If

' 4-corr 208*/232 age
  rw1 = IIf(piPb46col > 0, "Pb8Th2_4", "Pb8Th2_tot")
  rw2 = IIf(piPb46col > 0, "Pb8Th2_4e", "Pb8Th2_totE")
  PlaceFormulae "=LN(1+ " & rw1 & " )/" & fsS(pscLm2), f, piAgePb8Th2_4col, L
  PlaceFormulae "= " & rw1 & " /" & fsS(pscLm2) & "/(1+ " & rw1 & " )* " _
    & rw2 & " /100", f, piAgePb8Th2_4ecol, L
  If piPb46col = 0 Then CFs plHdrRw, piAgePb8Th2_4col, _
                        "total|208Pb|/232Th|age", True

' 4-corr 208*/232 %err
  If pbTh And piPb46col > 0 Then
    PlaceFormulae "=SQRT( Pb8Th2_tote ^2+(" & psaC84(0) & "/abs" & NetGamma & ")^2" & _
       "* Pb46e ^2)", f, piPb8Th2_4eCol, L
  End If

  If piAgePb6U8_7col > 0 Then
    t1 = "=" & pscEx8 & " AgePb6U8_7 )-1"                   ' /
    PlaceFormulae t1, f, piPb6U8_7col, L                    '| ' 09/06/22 added
    t1 = "=" & pscLm2 & "*" & pscEx8 & " AgePb6U8_7 " & ")" '| 7-corr 206*/238 & err
    t2 = "* AgePb6U8_7e / Pb6u8_7 *100"                     '|
    t3 = t1 & t2                                            '|
    PlaceFormulae t3, f, piPb6U8_7ecol, L                   ' \
  End If

  If piAgePb8Th2_7col > 0 Then
    t1 = "=Age7corrPb8Th2( Pb6U8_tot , Pb8Th2_tot , Pb86 , Pb76 ,sComm0_64," & _
         "sComm0_76,sComm0_86)"
    PlaceFormulae t1, f, piAgePb8Th2_7col, L
    t1 = "=AgeErr7corrPb8Th2( Pb6U8_tot ,  Pb6U8_tote , Pb8Th2_tot , Pb8Th2_tote "
    t2 = ", Pb76 , Pb76e , Pb86 , Pb86e ,sComm0_64,sComm0_76,sComm0_86)"
    PlaceFormulae t1 & t2, f, piAgePb8Th2_7ecol, L
    t1 = "=" & pscEx2 & " AgePb8Th2_7 )-1"                   ' /
    PlaceFormulae t1, f, piPb8Th2_7col, L                    '| ' 09/06/22 added
    t1 = "=" & pscLm2 & "*" & pscEx2 & " AgePb8Th2_7 " & ")" '| 7-corr 208*/232 & err
    t2 = "* AgePb8Th2_7e / Pb8Th2_7 *100"                    '|
    t3 = t1 & t2                                             '|
    PlaceFormulae t3, f, piPb8Th2_7ecol, L                   ' \
  End If

  If piPb46col > 0 Then  ' 208-corr 206*/238 age
    rw1 = fsRa(f, piaTh2U8col(0))
    t1 = " ,1/" & psaC86(0)
    tmp = "=Age8Corr( Pb6U8_tot , Pb8Th2_tot , th2u8(0)" & t1 & ")"
    On Error Resume Next
    PlaceFormulae tmp, f, piAgePb6U8_8col, L
    tmp = "=AgeEr8Corr( AgePb6U8_8 , Pb6U8_tot , Pb6U8_totE /100* Pb6U8_tot ,"
    tmp = tmp & " Pb8Th2_tot , Pb8Th2_tote /100* Pb8Th2_tot , th2u8(0) ,0" & t1 & ",0)"
' 208-corr 206*/238 age %err
     PlaceFormulae tmp, f, piAgePb6U8_8ecol, L
    On Error GoTo 0
  End If

  If pbCalc8corrConcPlotRats And piAgePb6U8_8col > 0 Then

    If piU8Pb6_8col > 0 Then
      PlaceFormulae "=Pb206U238rad( AgePb6U8_8 )", f, piPb6U8_8col, L
      PlaceFormulae "=" & pscLm8 & "*(1+ Pb6U8_8 )* AgePb6U8_8e *100/ Pb6U8_8 ", _
                     f, piPb6U8_8ecol, L
      PlaceFormulae "=1/ Pb6U8_8 ", f, piU8Pb6_8col, L
      PlaceFormulae "= Pb6U8_8e ", f, piU8Pb6_8ecol, L
    End If

    If piPb7U5_8col > 0 Then
      t1 = "=Rad8corPb7U5( AgePb6U8_8 , Pb6U8_tot , Pb76 ,sComm0_76)"
      PlaceFormulae t1, f, piPb7U5_8col, L

      t1 = " Pb6U8_tot , Pb6U8_totE , Pb6U8_8 , Pb6U8_tot * Pb76 /137.88, Th2U8(0) , Th2U8e(0) "
      t2 = ", Pb76 , Pb76e , Pb86 , Pb86e ,scomm0_76,scomm0_86"
      t3 = "=Rad8corPb7U5perr(" & t1 & t2 & ")"
      PlaceFormulae t3, f, piPb7U5_8ecol, L

      t1 = "( Pb6U8_tot , Pb6U8_totE , Pb6U8_8 , Th2U8(0) , Th2U8e(0) , Pb76 ," & _
           " Pb76e , Pb86 , Pb86e ,scomm0_76,scomm0_86)"
      t2 = "=Rad8corConcRho" & t1

      PlaceFormulae t2, f, piPb7U5Pb6U8_8rhoCol, L
    End If

    If piPb76_8col > 0 Then
      PlaceFormulae "= Pb7U5_8 / Pb6U8_8 /137.88", f, piPb76_8col, L
      tmp = "=sqrt( Pb7U5_8e ^2+ Pb6U8_8e ^2-2* Pb7U5_8e * Pb6U8_8e * Pb7U5Pb6U8_8rho )"
      PlaceFormulae tmp, f, piPb76_8ecol, L
    End If

    If piAgePb76_8col > 0 Then
      PlaceFormulae "=AgePb76( Pb76_8 )", f, piAgePb76_8col, L
      PlaceFormulae "=AgeErPb76( Pb76_8 , Pb76_8e ,,,,true)", f, piAgePb76_8ecol, L
    End If

  End If

End If ' If piPb8Th2_totCol > 0 And piPb86col > 0

StatBar "Final formatting"
HA xlRight, plaFirstDatRw(0), fvMax(4, piPb206ctsCol), plaLastDatRw(0), fiEndCol(plaFirstDatRw(0))
' Assume that all ages (4-corr, 7-corr, 8-corr, Th/Pb, & Pb/Pb) are adjacent
' Put in bold, different font-colors
FormatAge piAgePb6U8_4col, vbBlue
FormatAge piAgePb6U8_7col, 128
FormatAge piAgePb6U8_8col, RGB(0, 0, 128)
FormatAge piAgePb8Th2_4col, 8421376
FormatAge piaAgePb76_4Col(0), vbRed
FormatAge piAgePb8Th2_7col, RGB(0, 128, 0)
FormatAge piAgePb76_8col, RGB(0, 160, 0)

For i = peMaxCol To 1 Step -1
  If Not (IsEmpty(Cells(plHdrRw, i)) Or Columns(i).Hidden _
    Or Columns(i).ColumnWidth <= 1) Then Exit For
Next i

piLastVisibleCol = i

For i = f To plHdrRw + piaSpotCt(0)  ' Put U-Std rows in lite-yellow bkrd

  If piaPpmUcol(0) > 0 And psConcStdNa <> "" And _
    InStr(LCase(Cells(i, 1)), LCase(psConcStdNa)) Then ' 09/07/21 -- mod
    IntClr RGB(255, 255, 192), i, 1, , piLastVisibleCol
    Cells(i, 1).Font.Color = vbBlue
  End If

Next i

BorderLine xlBottom, 2, plHdrRw, 1, , piLastVisibleCol

' 09/07/21 -- commented out
'For i = piLastVisibleCol + 1 To 1 Step -1
'  k = 0
'
'  For j = 1 To 3
'    If Cells(j, i) <> "" Then k = 1: Exit For
'  Next j
'
'  If k Then Exit For
'Next i

For i = 1 To fiEndCol(plHdrRw) ' 09/07/21 -- added
  If Not Columns(i).Hidden Then
    ColWidth picAuto, i
  End If
Next i

'ColWidth picAuto, i + 1, piLastVisibleCol ' 09/07/21 -- commented out
StatBar
End Sub

Sub CreateSampleCommPbBox()
' Create the Standard-sheet box containing common-Pb isotope ratios
'  (U/Pb Tasks only)
Dim c%, Fcol%, Lcol%, r&, Lrow&, FRow&
Dim b As Range

phStdSht.Activate
Set b = phStdSht.[stdcommpb]
phSamSht.Activate
FRow = 2
Fcol = fiEndCol(2) + 5
Lrow = FRow + b.Rows.Count - 1
Lcol = Fcol + b.Columns.Count - 1
b.Copy Cells(FRow, Fcol)
Set b = frSr(FRow, Fcol, Lrow, Lcol)
b.Name = "SamCommPb"
r = 1 + b.Row: c = 2 + b.Column
Cells(r, c).Name = psaC64(0)
Cells(r, c).Name = psaC64(0)
Cells(r + 1, c).Name = psaC76(0)
Cells(r + 2, c).Name = psaC86(0)
Cells(r, c + 3).Name = psaC74(0)
Cells(r + 1, c + 3).Name = psaC84(0)
End Sub

Sub PlaceGroupShtBoxes(ToSortBox As Object, CalibrConstBox As Object, _
                        Ar As Range, BoxCol%, CPbBox)
' For a U/Pb Task Sample sheet, place the formatted weighted-average results for the
'  currently specified or calculated "Coherent age group) spots.
Dim So$, Nf$
Dim i%, j%, k%, rw1&, rw2, tW!, Cw!

ToSortBox.Top = Rows(3 + plaLastDatRw(0)).Top
tW = ToSortBox.Width
ToSortBox.Left = fvMax(10, Columns(BoxCol).Left - tW - 200)
With CalibrConstBox
  .Left = ToSortBox.Left: .Top = fnBottom(ToSortBox) + 25
End With

j = plaLastDatRw(0)
Do
  j = j + 1
  rw1 = Rows(j).Top
Loop Until rw1 > CalibrConstBox.Top

k = 1
Do
  k = k + 1
  rw2 = Columns(k).Left
Loop Until rw2 > fnRight(CalibrConstBox)

Ar.Cut Cells(j, k)
Cw = Columns(k).ColumnWidth
Ar.Font.Size = 11: Ar.IndentLevel = 1
Nf = Range(pscStdShtNot & "WtdMeanA1").NumberFormat
If Left$(Nf, 2) = "0." Then Nf = Mid$(Nf, 2)
Ar(1).NumberFormat = Nf
Ar(2).NumberFormat = pscZd1 & fsQq("$%$")
Ar(3).NumberFormat = pscZd2 & fsQq("$%$")
Box Ar.Row, k, Ar(3).Row, k + 1, peLightGray
CalibrConstBox.Top = Ar.Top - (CalibrConstBox.Height - Ar.Height) / 2
CalibrConstBox.Left = Ar.Left - CalibrConstBox.Width
Columns(k).ColumnWidth = Cw

For j = 1 To 3
  Range(Ar(j, 1), Ar(j, 2)).Merge
Next j

For i = 0 To 1
  So$ = IIf(i, "Down", "Up")
  fhSquidSht.Shapes("xSort" & So$).Copy
  ActiveSheet.Paste
  With ActiveSheet.Shapes("xSort" & So$)
    .Width = 25: .Height = 28
    .Left = 7 + fnRight(ToSortBox)
    .Top = ToSortBox.Top + i * (.Height + 2)
    .OnAction = ThisWorkbook.Name & "!Sort" & So$
  End With
Next i

If CPbBox.Count > 1 Then
  j = k - CPbBox.Columns.Count + 2
  CPbBox.Cut Cells(Ar(5).Row, j)

  For k = j + 2 To j + 5 Step 3
    With Cells(1, k)
      .ColumnWidth = fvMax(5, .ColumnWidth)
    End With
  Next k

End If

Fonts 1 + Ar(5).Row, 2 + j, 3 + Ar(5).Row, , , , , 10
Fonts 1 + Ar(5).Row, 5 + j, 3 + Ar(5).Row, , , , , 10
End Sub

Sub SquidInvokedConcPlot(OnSheet As Worksheet, NotDone As Boolean)
' Construct a conventional Concordia plot-inset for either the Age Std or
'  Grouped Samples on the relevant worksheet.
Dim BadPt As Boolean, NoObj As Boolean
Dim s$
Dim i%, j%, ct%, k%, m%, Col1%, Col2%, N%, Xcol%, Ycol%, RhoCol%
Dim rw1&, rw2&
Dim Top!, Left!, wW!, t2!
Dim DatInp#()
Dim TmpR(1 To 5) As Variant
Dim Cr As Range, LastChart As ChartObject

NotDone = True
FindStr fsVertToLF("207*|/235"), , Xcol, plHdrRw  ' 09/06/10 -- added "4corr|
' 10/04/02 -- line below was finding just "206*|/238", and so was using
'             7-corr 206/238 for concordia ratios.
FindStr fsVertToLF("4corr|206*|/238"), , Ycol, plHdrRw  '     "          "
FindStr fsVertToLF("err|corr"), , RhoCol, plHdrRw
If Xcol = 0 Or Ycol = 0 Or RhoCol = 0 Then Exit Sub

plHdrRw = flHeaderRow(pbStd)
rw1 = plaFirstDatRw(-pbStd)
rw2 = plaLastDatRw(-pbStd)
Col1 = fvMin(Xcol, Ycol, RhoCol)
Col2 = fvMax(Xcol, Xcol + 1, Ycol, Ycol + 1, RhoCol)
N = rw2 - rw1 + 1: ct = 0: i = 0: m = 0
ReDim DatInp(1 To 5, 1 To N)
Set Cr = frSr(rw1, Xcol, rw2)

Do
  BadPt = False
  m = 1 + m

  If fbNoNum(Cr(m, 1)) Or fbNoNum(Cr(m, 2)) Or fbNoNum(Cr(m, 3)) Or fbNoNum(Cr(m, 4)) Then
    Exit Do
  ElseIf Cr(m, 1) <= 0 Or Cr(m, 2) <= 0 Or Cr(m, 3) <= 0 Or Cr(m, 4) <= 0 Then
    Exit Do
  End If

  TmpR(1) = Cr(m) ' 1st spot conv concplot x-xer-y-yer-rho
  TmpR(2) = Cr(m, 2) / 100 * TmpR(1)
  TmpR(3) = frSr(rw1, Ycol, rw2)(m)
  TmpR(4) = frSr(rw1, 1 + Ycol, rw2)(m) / 100 * TmpR(3)
  TmpR(5) = frSr(rw1, RhoCol, rw2)(m)

  For j = 1 To 4
    If Cr(m).Font.Strikethrough Then BadPt = True: Exit For
  Next

  If Not BadPt Then
    ct = 1 + ct

    For k = 1 To 5
      DatInp(k, ct) = TmpR(k)
    Next k

  End If

Loop Until m = N

If ct = 0 Then Exit Sub
N = ct
ReDim inpdat(1 To N, 1 To 5)

For i = 1 To N
  For j = 1 To 5
    inpdat(i, j) = DatInp(j, i)
Next j, i

With Cells(2 + 5, Col1)
  Top = .Top + 5: Left = .Left
End With

Isoplot3.N = N
SetLastChart LastChart, NoObj
StartFromSquid
OnSheet.DisplayAutomaticPageBreaks = False
FromInit = False: FromIso = True: FromSquid = False: HardRej = False
ActiveChart.CopyPicture ' Appearance:=xlScreen, Format:=xlPicture
If ActiveSheet.Name <> pscStdShtNa Then DelSheet
DelSheet Sheets(PlotDat)
OnSheet.PasteSpecial
On Error GoTo 0

With Selection.ShapeRange

  i = 0
  Do
    i = i + 1
    s = "sqCncrd" & fsS(i)
  Loop Until Not fbRangeNameExists(s) Or i = 99

  If i < 99 Then .Name = s
  .Height = 210: .Width = 280

  If NoObj Then
    .Top = fnBottom(Rows(plaLastDatRw(0) + 2))
    .Left = fvMax(2, fnRight(Columns(Col1 - 22))) - 10
  ElseIf ActiveSheet.Name = pscStdShtNa Then
    k = IIf(pbUPb, piNumDauPar, 1)
    j = Range(pscStdShtNot & "wtdmeana" & fsS(k)).Column
    .Left = fnRight(Columns(j)) + 120 - 120 * (pbUPb And piNumDauPar = 2)
    .Top = fnBottom(Rows(plaLastDatRw(1) + 4))
  Else
    Set Cr = Cells(piConcordPlotRow, piConcordPlotCol)
    .Top = fnBottom(Cr(4, 2)) + 4
    .Left = fnRight(Cr) - .Width - 5
  End If

  wW = .Left + .Width
End With

With foLastOb(ActiveSheet.Shapes).Line
  .Weight = 1: .Visible = True
End With

FindStr "calibr", , Col2, plHdrRw

With ActiveWindow
  .ScrollRow = fvMax(1, rw2 - 16)
  Do
    Col2 = 1 + Col2
    .ScrollColumn = Col2
    With .VisibleRange: t2 = 5 + .Left + .Width: End With
  Loop Until t2 > wW
End With

NotDone = False
Cells(rw1, Col1).Activate
End Sub

Sub GetRelocColnum(EqNum, EqnRatToReloc, ByVal EqnStr$, ModStr$, ByVal Std As Boolean)
' If a Task equation contains the Swap Column indicator ("<=>"), determine which isotope
'  ratio to swap with.
Dim ErCol As Boolean
Dim tse$, s$, Msg$, Eq0$
Dim p%, c%, N%, Indx%, IndxType%

' If an error-column, return the negative# of the corresponding value column
Eq0 = EqnStr
EqnStr = Eq0

With puTask
  If EqNum > 0 Then
    If (.uaSwitches(EqNum).SA And Std) Or (.uaSwitches(EqNum).ST And Not Std) Then
      Exit Sub
    End If
  End If
End With

EqnStr = Trim(Mid$(EqnStr, 1 + InStr(EqnStr, "}")))
p = InStr(EqnStr, "<=>"): c = 0

If p > 0 Then
  With puTask
    ModStr = Left$(EqnStr, p - 1)
    tse = LCase(Trim(Mid$(EqnStr, p + 3)))
    ErCol = (Mid$(tse, 2, 1) = pscPm)
    ExtractEqnRef tse, s, Indx, IndxType, .saIsoRats, .saEqnNames

    If IndxType = peRatio Then            ' Indx is an isotope-ratio index#
      c = FindHeaderCol(.saIsoRats(-Indx))
    ElseIf IndxType = peColumnHeader Then ' Indx is 1000 + col#
      c = Indx - 1000
    ElseIf Indx > -1000 Then              ' not a const
      c = piaEqCol(Std, Indx)
      If c = 0 Then
        c = FindHeaderCol(.saEqnNames(Indx))
      End If
    End If

  End With

  If c = 0 Then
    CrashNoise
    SqBrakQuExtract Mid$(EqnStr, p), Msg
    s = fsInQ(Msg)
    Msg = "No column header in the " & IIf(Std, "Standard", "Sample") & _
      " worksheet matches " & s & "." & pscLF2 & "Please re-check " & _
      "Task Equation" & StR(EqNum) & ". "
    MsgBox Msg, , pscSq
    End
  End If

  EqnRatToReloc = c - ErCol
Else
  EqnRatToReloc = 0
End If

End Sub

Public Function Age7corrPb8Th2(ByVal TotPb206U238#, ByVal TotPb208Th232#, _
          ByVal TotPb86#, ByVal TotPb76#, ByVal CommonPb64#, _
          ByVal CommonPb76#, ByVal CommonPb86#) As Variant
' 09/06/14 -- created.
' Returns the 208Pb/232Th age, assuming the true 206/204 is that required
'  to force 206/238-207/235 concordance.
Dim Alpha#, Gamma#, Beta0#, Gamma0#, RadPb6U8#, RadPb8Th2#, term#
Dim Radfract8#, Age7corPb6U8#, tmp As Variant

tmp = "#NUM!"
On Error GoTo Done

Beta0 = CommonPb64 * CommonPb76
Gamma0 = CommonPb64 * CommonPb86
On Error Resume Next
Age7corPb6U8 = Isoplot3.Age7corr(TotPb206U238, TotPb76, CommonPb76)
RadPb6U8 = Exp(pscLm8 * Age7corPb6U8) - 1
term = TotPb206U238 - RadPb6U8
If term = 0 Then term = 1E-16
Alpha = CommonPb64 * TotPb206U238 / term
Gamma = TotPb86 * Alpha
Radfract8 = (Gamma - Gamma0) / Gamma
RadPb8Th2 = TotPb208Th232 * Radfract8
On Error Resume Next
tmp = Log(1 + RadPb8Th2) / pscLm2

Done: On Error GoTo 0
Age7corrPb8Th2 = tmp
End Function

Public Function AgeErr7corrPb8Th2(ByVal TotPb206U238#, ByVal TotPb206U238percentErr#, _
   ByVal TotPb208Th232#, ByVal TotPb208Th232percentErr#, ByVal TotPb76#, _
   ByVal TotPb76percentErr#, ByVal TotPb86#, ByVal TotPb86percentErr#, _
   ByVal CommPb64#, ByVal CommPb76#, ByVal CommPb86#) As Variant
' 09/06/14 -- created.
' Returns the error of the 208Pb/232Th age, where the 208Pb/232Th age is calculated
'   assuming the true 206/204 is that required to force 206/238-207/235 concordance.
' The error is calculated numerically, by successive perturbation of the input errors.
Dim Age7corr68#, Gamma#, Age#, Pert%, Pphi#, Ptheta#, PtotPb6U8#, PtotPb8Th2#, tmp As Variant
Dim Delta#(0 To 4), PhiIn#, ThetaIn#, TotPb6U8In#, TotPb8Th2In#, AgeVariance#, DeltaT#

tmp = "#NUM!"
On Error GoTo Done

' Perturb each input variable by its assigned error
Pphi = TotPb76 * (1 + TotPb76percentErr / 100)
Ptheta = TotPb86 * (1 + TotPb86percentErr / 100)
PtotPb6U8 = (1 + TotPb206U238percentErr / 100) * TotPb206U238
PtotPb8Th2 = (1 + TotPb208Th232percentErr / 100) * TotPb208Th232

' archive the input variables
PhiIn = TotPb76
ThetaIn = TotPb86
TotPb6U8In = TotPb206U238
TotPb8Th2In = TotPb208Th232
AgeVariance = 0

' Find ages perturbing each input variable in succession

For Pert = 0 To 4

  Select Case Pert ' perturb the input variables
    Case 0
    Case 1: TotPb206U238 = PtotPb6U8
    Case 2: TotPb208Th232 = PtotPb8Th2
    Case 3: TotPb86 = Ptheta
    Case 4: TotPb76 = Pphi
  End Select

  Delta(Pert) = Age7corrPb8Th2(TotPb206U238, TotPb208Th232, TotPb86, TotPb76, CommPb64, CommPb76, CommPb86)

  If Pert > 0 Then
    ' restore the input variables
    TotPb76 = PhiIn
    TotPb86 = ThetaIn
    TotPb206U238 = TotPb6U8In
    TotPb208Th232 = TotPb8Th2In
    DeltaT = Delta(Pert) - Delta(0)
    AgeVariance = AgeVariance + DeltaT ^ 2
  End If

Next Pert

tmp = sqR(AgeVariance)

Done: On Error GoTo 0
AgeErr7corrPb8Th2 = tmp
End Function

Public Function Rad8corPb7U5(ByVal Age8corPb6U8#, ByVal TotPb6U8#, ByVal TotPb76#, _
   ByVal CommPb76#) As Variant
' 09/06/14 -- created.
' Returns the radiogenic 208-corrected 207Pb*/235U ratio.
Dim RadPb6U8#, CommFract6#, tmp As Variant
tmp = "#NUM!"
On Error GoTo Done

RadPb6U8 = Pb206U238rad(Age8corPb6U8)
CommFract6 = 1 - RadPb6U8 / TotPb6U8
tmp = (TotPb76 - CommFract6 * CommPb76) * pscUra * TotPb6U8

Done: On Error GoTo 0
Rad8corPb7U5 = tmp
End Function

Public Function Age8corPb46(ByVal Age8corPb6U8#, ByVal TotPb6U8#, ByVal CommPb64#) As Variant
' 09/06/14 -- created.
' Returns the 204Pb/206Pb implied by a 208-corrected 206Pb*/238U age.
Dim RadPb6U8#, tmp As Variant
tmp = "#NUM!"
On Error GoTo Done

RadPb6U8 = Exp(pscLm8 * Age8corPb6U8) - 1
tmp = (TotPb6U8 - RadPb6U8) / (CommPb64 * TotPb6U8)

Done: On Error GoTo 0
Age8corPb46 = tmp
End Function

Public Function Rad8corConcRho(ByVal TotPb6U8#, ByVal TotPb6U8per#, ByVal RadPb6U8#, _
  ByVal Th2U8#, ByVal Th2U8per#, ByVal TotPb76#, ByVal TotPb76per#, _
  ByVal TotPb86#, ByVal TotPb86per#, ByVal CommPb76#, ByVal CommPb86#) As Variant
' 09/06/14 -- created.
' Returns the error correlation for 208-corrected 206Pb*/238U-207Pb*/235U ratio-pairs.

Dim RadPb8Th2#, RadPb7U5#, CommFract6#, RadFract6#
Dim SigmaRadPb6U8#, SigmaRadPb7U5#, SigmaTotPb86#, SigmaTotPb76#
Dim CovTotPb86CommFract6#, CovRad68Rad75#
Dim SigmaCommfract6#, SigmaTotPb6U8#, SigmaK#
Dim u#, q#, m1#, m2#, h1#, h2#, k#
Dim Term1#, Term2#, Term3#, term4#, tmp As Variant

tmp = "#NUM!"
On Error GoTo Done

u = pscUra
k = 1 / Th2U8
SigmaK = Th2U8per / 100 * k
SigmaTotPb76 = TotPb76per / 100 * TotPb76
SigmaTotPb86 = TotPb86per / 100 * TotPb86
SigmaTotPb6U8 = TotPb6U8per / 100 * TotPb6U8
RadFract6 = RadPb6U8 / TotPb6U8
CommFract6 = 1 - RadFract6
q = TotPb86 - CommPb86 * CommFract6

RadPb8Th2 = (TotPb86 - CommFract6 * CommPb86) * TotPb6U8 / Th2U8
RadPb7U5 = (TotPb76 - CommFract6 * CommPb76) * u * TotPb6U8
m1 = pscLm8 * (1 + RadPb6U8)
m2 = pscLm2 * (1 + RadPb8Th2)
h1 = RadFract6 / m1 - q * k / m2
h2 = 1 / (TotPb6U8 / m1 - k * TotPb6U8 * CommPb86 / m2)

Term1 = (h1 * SigmaTotPb6U8) ^ 2
Term2 = (TotPb6U8 / m2) ^ 2
Term3 = (q * SigmaK) ^ 2
term4 = (k * SigmaTotPb86) ^ 2

SigmaCommfract6 = sqR(h2 ^ 2 * (Term1 + Term2 * (Term3 + term4)))

CovTotPb86CommFract6 = h1 * h2 * SigmaTotPb6U8 ^ 2
SigmaRadPb6U8 = sqR((RadFract6 * SigmaTotPb6U8) ^ 2 + TotPb6U8 ^ 2 _
  * SigmaCommfract6 ^ 2)

Term1 = (RadPb7U5 / TotPb6U8 * SigmaTotPb6U8) ^ 2
Term2 = (u * TotPb6U8) ^ 2 * (SigmaTotPb76 ^ 2 + (CommPb76 * SigmaCommfract6) ^ 2)
Term3 = -2 * RadPb7U5 * u * CommPb76 * CovTotPb86CommFract6

SigmaRadPb7U5 = sqR(Term1 + Term2 + Term3)

Term1 = RadFract6 * RadPb7U5 / TotPb6U8 * SigmaTotPb6U8 ^ 2
Term2 = u * TotPb6U8 ^ 2 * CommPb76 * SigmaCommfract6 ^ 2
Term3 = -CovTotPb86CommFract6 * (u * TotPb6U8 * RadFract6 * CommPb76 + RadPb7U5)

CovRad68Rad75 = Term1 + Term2 + Term3
tmp = CovRad68Rad75 / (SigmaRadPb6U8 * SigmaRadPb7U5)

Done: On Error GoTo 0
Rad8corConcRho = tmp
End Function

Public Function Rad8corPb7U5perr(ByVal TotPb6U8#, ByVal TotPb6U8per#, ByVal RadPb6U8#, ByVal TotPb7U5#, _
  ByVal Th2U8#, ByVal Th2U8per#, ByVal TotPb76#, ByVal TotPb76per#, ByVal TotPb86#, _
  ByVal TotPb86per#, ByVal CommPb76#, ByVal CommPb86#) As Variant
' 09/06/14 -- created
' Returns the %error of a 208-corrected 207Pb*/235U.

Dim RadPb8Th2#, RadPb7U5#, CommFract6#, RadFract6#
Dim SigmaRadPb7U5#, SigmaTotPb76#
Dim CovTotPb86CommFract6#, SigmaTotPb86#
Dim SigmaCommfract6#, SigmaTotPb6U8#, k#, SigmaK#
Dim u#, q#, m1#, m2#, h1#, h2#, Term1#
Dim Term2#, Term3#, term4#, tmp As Variant

tmp = "#NUM!"
On Error GoTo Done

k = 1 / Th2U8
SigmaTotPb6U8 = TotPb6U8per / 100 * TotPb6U8
SigmaTotPb76 = TotPb76per / 100 * TotPb76
SigmaTotPb86 = TotPb86per / 100 * TotPb86
SigmaK = Th2U8per / 100 * k
u = pscUra
RadFract6 = RadPb6U8 / TotPb6U8
CommFract6 = 1 - RadFract6
q = TotPb86 - CommPb86 * CommFract6

RadPb8Th2 = (TotPb86 - CommFract6 * CommPb86) * TotPb6U8 / Th2U8
RadPb7U5 = (TotPb76 - CommFract6 * CommPb76) * u * TotPb6U8
m1 = pscLm8 * (1 + RadPb6U8)
m2 = pscLm2 * (1 + RadPb8Th2)
h1 = RadFract6 / m1 - q * k / m2
h2 = 1 / (TotPb6U8 / m1 - k * TotPb6U8 * CommPb86 / m2)

Term1 = (h1 * SigmaTotPb6U8) ^ 2
Term2 = (TotPb6U8 / m2) ^ 2
Term3 = (q * SigmaK) ^ 2
term4 = (Th2U8 * SigmaTotPb86) ^ 2

SigmaCommfract6 = sqR(h2 ^ 2 * (Term1 + Term2 * (Term3 + term4)))
CovTotPb86CommFract6 = h1 * h2 * SigmaTotPb6U8 ^ 2

Term1 = (RadPb7U5 / TotPb6U8 * SigmaTotPb6U8) ^ 2
Term2 = (u * TotPb6U8) ^ 2 * (SigmaTotPb76 ^ 2 + (CommPb76 * SigmaCommfract6) ^ 2) '
Term3 = -2 * RadPb7U5 * u * CommPb76 * CovTotPb86CommFract6

SigmaRadPb7U5 = sqR(Term1 + Term2 + Term3)
tmp = SigmaRadPb7U5 / RadPb7U5 * 100

Done: On Error GoTo 0
Rad8corPb7U5perr = tmp
End Function

Public Function Pb206U238rad(ByVal AgeMa) As Variant
' 09/06/14 -- created
' Returns the radiogenic 206Pb/238U ratio for the specified age.
Dim tmp As Variant
tmp = "#NUM!"
On Error GoTo Done
tmp = Exp(pscLm8 * AgeMa) - 1

Done: On Error GoTo 0
Pb206U238rad = tmp
End Function

Public Function Pb207U235rad(ByVal AgeMa) As Variant
' 09/06/14 -- created
' Returns the radiogenic 207Pb/235U ratio for the specified age.
Dim tmp As Variant
tmp = "#NUM!"
On Error GoTo Done
tmp = Exp(pscLm5 * AgeMa) - 1

Done: On Error GoTo 0
Pb207U235rad = tmp
End Function

Public Function Pb207Pb206rad(ByVal AgeMa) As Variant
' 09/06/19-07/02 created
Dim temp#, tmp As Variant
' Returns the radiogenic 207Pb/206Pb ratio for the specified age.
tmp = "#NUM!"
On Error GoTo Done

If AgeMa <> 0 Then
  temp = Pb207U235rad(AgeMa) / Pb206U238rad(AgeMa)
Else
  temp = pscLm5 / pscLm8
End If

tmp = temp / pscUra

Done: On Error GoTo 0
Pb207Pb206rad = tmp
End Function

Public Function Pb86radCor4per(ByVal Pb86#, ByVal Pb86perr#, ByVal Pb46#, ByVal Pb46perr#, _
  ByVal PbRad86cor4#, CommPb64#, ByVal CommPb84#) As Variant
Dim Numer#, Denom#, SigmaPb86#, SigmaPb46#, tmp As Variant, Var#
' Returns the %error of the radiogenic (204-corrected) 208Pb/206Pb ratio.
' 09/06/14 -- created
tmp = "#NUM!"
On Error GoTo Done

SigmaPb86 = Pb86perr / 100 * Pb86
SigmaPb46 = Pb46perr / 100 * Pb46

Denom = (1 - Pb46 * CommPb64) ^ 2
Numer = SigmaPb86 ^ 2 + (PbRad86cor4 * CommPb64 - CommPb84) ^ 2 * SigmaPb46 ^ 2

Var = Numer / Denom
tmp = sqR(Var) * 100 / Abs(PbRad86cor4)

Done: On Error GoTo 0
Pb86radCor4per = tmp
End Function

Public Function Pb46cor7(ByVal Pb76tot#, ByVal Alpha0#, ByVal Beta0#, ByVal Age7corPb6U8#)
' 09/06/13 -- created
' Returns 204Pb/206Pb required to force 206Pb/238U-207Pb/206Pb ages to concordance
Dim Pb76true#, tmp
tmp = "#NUM!"
On Error GoTo Done
Pb76true = Pb76(Age7corPb6U8)
tmp = (Pb76tot - Pb76true) / (Beta0 - Pb76true * Alpha0)

Done: On Error GoTo 0
Pb46cor7 = tmp
End Function

Public Function Pb46cor8(ByVal Pb86tot#, ByVal Th2U8#, ByVal Alpha0#, _
                  ByVal Gamma0#, ByVal Age8corPb6U8#)
' 09/06/13 -- created
' Returns 204Pb/206Pb required to force 206Pb/238U-208Pb/232Th ages to concordance
Dim Pb86rad#, tmp, Age, Numer#, Denom#
tmp = "#NUM!"
On Error GoTo Done
Numer = Exp(Age8corPb6U8 * pscLm2) - 1
Denom = Exp(Age8corPb6U8 * pscLm8) - 1
Pb86rad = Numer / Denom * Th2U8
tmp = (Pb86tot - Pb86rad) / (Gamma0 - Pb86rad * Alpha0)

Done: On Error GoTo 0
Pb46cor8 = tmp
End Function

Public Function Pb86radCor7(ByVal Pb86tot#, Pb76tot#, ByVal Th2U8#, ByVal Alpha0#, _
                     ByVal Beta0#, ByVal Gamma0#, ByVal Age7corPb6U8#)
' 09/06/13 -- created
' Returns radiogenic 208Pb/206Pb where the common 204Pb/206Pb is that
'   required to force the 206Pb/238U-208Pb/232Th ages to concordance
' Alternate: RadFract208cor7/RadFract206cor7*Pb86tot
'            (EXP(0.000049475* Age7corPb8Th2 -1)/ Pb8Th2tot / Pb6U87cor * Pb6U8tot * Pb86
Dim tmp, Age, Numer#, Denom#, Pb46#, GammaTot#
tmp = "#NUM!"
On Error GoTo Done
Pb46 = Pb46cor7(Pb76tot, Alpha0, Beta0, Age7corPb6U8)
GammaTot = Pb86tot / Pb46
Numer = GammaTot - Gamma0
Denom = 1 / Pb46 - Alpha0
tmp = Numer / Denom

Done: On Error GoTo 0
Pb86radCor7 = tmp
End Function

Public Function Pb86radCor7per(ByVal Pb86tot#, ByVal Pb86totPer#, ByVal Pb76tot#, _
         ByVal Pb76totPer#, ByVal Pb6U8tot#, ByVal Pb6U8totPer#, ByVal Age7corPb6U8#, _
         ByVal Alpha0#, ByVal Beta0#, ByVal Gamma0#)
' 09/06/14 -- created
' Returns radiogenic 208Pb/206Pb %err where the common 204Pb/206Pb is that
'   required to force the 206Pb/238U-208Pb/232Th ages to concordance
Dim tmp, Age, Numer#, Denom#, Pb46#, GammaTot#
Dim k1#, k2#, k3#, k4#, k5#, k6#, k7#, d1#, d2#, d3#
Dim m1#, m2#, p#, j1#, j2#, r#, s#, rStar#, sStar#, Exp5#, Exp8#
Dim Phi0#, Phi#, PhiStar#, Theta#, ThetaStar7#, u#, AlphaPrime#, Var#
Dim SigmaTheta#, SigmaPhi#, SigmaR#, SigmaThetaStar7#, VarThetaStar7#
tmp = "#NUM!"
On Error GoTo Done

r = Pb6U8tot:        SigmaR = Pb6U8totPer / 100 * r
Phi = Pb76tot:       SigmaPhi = Pb76totPer / 100 * Phi
Theta = Pb86tot:     SigmaTheta = Pb86totPer / 100 * Theta
u = 1 / pscUra
Phi0 = Beta0 / Alpha0

Exp5 = Exp(Age7corPb6U8 * pscLm5)
Exp8 = Exp(Age7corPb6U8 * pscLm8)
rStar = Exp8 - 1
sStar = Exp5 - 1
PhiStar = sStar / rStar * u
p = rStar / r

AlphaPrime = (Phi - PhiStar) / (Beta0 - PhiStar * Alpha0)
ThetaStar7 = (Theta / AlphaPrime - Gamma0) / (1 / AlphaPrime - Alpha0)
m1 = pscLm8 * Exp8
m2 = pscLm5 * Exp5

j1 = p / m1 - sStar / r / m2
j2 = 1 / (r / m1 - u * r * Phi0 / m2)
d1 = 1 - Alpha0 * AlphaPrime
d2 = Beta0 - PhiStar * Alpha0

k1 = (p - r * j1 * j2) / m1
k2 = u * j2 * r ^ 2 / m1 / m2
k3 = u / rStar * (m2 - PhiStar * m1)
k4 = (Alpha0 * AlphaPrime - 1) / d2
k5 = AlphaPrime / d1
k7 = (Alpha0 * ThetaStar7 - Gamma0) / d1

VarThetaStar7 = (SigmaTheta / d1) ^ 2 + (k1 * k3 * k4 * k7 * SigmaR) ^ 2 + _
      (k7 * (1 / d2 + k2 * k3 * k4) * SigmaPhi) ^ 2
If VarThetaStar7 < 0 Then
  tmp = 999
Else
  SigmaThetaStar7 = sqR(VarThetaStar7)
  tmp = 100 * SigmaThetaStar7 / Abs(ThetaStar7)
End If

Done: On Error GoTo 0
Pb86radCor7per = tmp
End Function

Public Function StdPb86radCor7per(ByVal Pb86tot#, ByVal Pb86totPer#, ByVal Pb76tot#, _
         ByVal Pb76totPer#, ByVal RadPb86cor7#, ByVal Pb46cor7#, ByVal StdRadPb76#, _
         ByVal Alpha0#, ByVal Beta0#, ByVal Gamma0#)
' 09/06/14 -- created
' Returns radiogenic 208Pb/206Pb where the common 204Pb/206Pb is that
'   required to force the 206Pb/238U-208Pb/232Th ages to concordance
Dim tmp, k5#, k7#, d1#, d2#, Phi#, PhiStar#, Theta#, ThetaStar7#, AlphaPrime#
Dim SigmaTheta#, SigmaPhi#, SigmaThetaStar7#, VarThetaStar7#

tmp = "#NUM!"
On Error GoTo Done

AlphaPrime = Pb46cor7
Phi = Pb76tot:       SigmaPhi = Pb76totPer / 100 * Phi
Theta = Pb86tot:     SigmaTheta = Pb86totPer / 100 * Theta
ThetaStar7 = RadPb86cor7
PhiStar = StdRadPb76

d1 = 1 - AlphaPrime * Alpha0
d2 = Beta0 - Alpha0 * PhiStar
k7 = (Alpha0 * ThetaStar7 - Gamma0) / d1

VarThetaStar7 = (SigmaTheta / d1) ^ 2 + (k7 / d2 * SigmaPhi) ^ 2

If VarThetaStar7 < 0 Then
  tmp = 999
Else
  SigmaThetaStar7 = sqR(VarThetaStar7)
  tmp = 100 * SigmaThetaStar7 / Abs(ThetaStar7)
End If

Done: On Error GoTo 0
StdPb86radCor7per = tmp
End Function
