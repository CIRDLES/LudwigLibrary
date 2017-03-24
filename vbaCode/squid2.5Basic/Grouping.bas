Attribute VB_Name = "Grouping"
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
' 09/05/02 -- Add Sub CompareUOUwithStd (places chart-inset of
'              UOx/UOy of std & grouped spots for comparison)
' 09/05/03 -- In Sub GroupThis, make sure that Squid-hidden columns remain hidden
'              when copied to a grouped-sample sheet.
' 09/12/04 -- Minor mods to Sub GroupThis to yield correct reponse to choices of
'              SK versus Specified common-Pb ratios.
Option Explicit
Option Base 1

Sub GroupThis() ' Create separate sheets of reduced data for different samples
' 09/03/25 -- Rewrite 204-overcount correction of 204/206 ratio & err to obviate problems
'             reliance on tot204cps-Bkrdcps being nonzero.  Instead of correcting the measured
'             204/206 for overcounts, calculate the ratio entirely from the 204, 206, & bkrd cps
'             columns.  Also, instead of just keeping the uncorrected 4/6 error, calculate rigorously
'             from the counting stats used in the corrected 4/6.  Also, instead of using the "fsFcell"
'             function, which relies on the possibly out-dated ColIndex sheet, specifically find the
'             required col-headers required for the cell addresses in the cell formulae.
' 09/04/10 -- Replace the "99" in "frSr(Frw, 1, Lrw, 99).Sort" by "fiEndCol(flHeaderRow(0))"
'             so that can deal with sample sheets with >=100 columns.
' 09/07/09 -- Major code-rewrite.  From the Sample sheet, keep only the relevant cols (for the common-Pb
'             index isotope and the radiogenic isotope) of:  7cor46, 8cor46, 4corcom6, 7corcom6, 8corcom6,
'             4corcom8, 7corcom8,4cor86, 7cor86, 4corppm6, 7corppm6, 8corppm6, 4corppm8, 7corppm8
'             after copying to the Grouped sheet.  Use only the piGrpDateType variable as an indication
'             of what age-type to group.

Dim Bad As Boolean, TooFew As Boolean, No4 As Boolean, No7 As Boolean, No8 As Boolean
Dim Sx$, Sp$, Sg$, wn$, nSh$, SK_Age$, a2$, tmp$, Na$, Term1$
Dim Term2$, Term3$, term4$, term5$, term6$, term7$, FinalTerm$, Msg$
Dim i%, j%, k%, g%, m%, Sct%, NaCt%, MinAgeCol%, AgeCol%, Lcol%, Ln%, Col%, RawCol%, CorrCol%, Iter%
Dim SortCol%, BoxCol%, FirstAcol%, ErrNum%, GrpAgetypeCol%, SKageCol%
Dim GrpNum%, GrpAgeCol%, ComDauCol%, Rad86col%, ComColIndx%, Pb86radColIndx%
Dim Rw&, Frw&, Lrw&, LastRow&
Dim Top!, tL!
Dim AgeVal#
Dim Com64grp As Range, Com74grp As Range, Com84grp As Range, Com76grp As Range
Dim Com86grp As Range, ar2 As Range, CPbBox As Range, Ar As Range, Ce As Range
Dim SourceSht As Object, SsSht As Worksheet, Sht As Worksheet
Dim SourceWbk As Object, CalibrConstBox As Object, ns As Object, ToSortBox As Object
Dim TempVal As Variant, GrpAgeNa As Variant, temp As Variant, TabClrs As Variant

TabClrs = Array(8388608, 26367, 65280, 13209, 65535, 16764057, _
                8421376, 16751052, 52479, 16776960, 13056, 32896)

GrpAgeNa = Array("Total206Pb/238uage", "204corr206Pb/238uage", _
                 "207corr206Pb/238uage", "208corr206Pb/238uage", _
                 "204corr207Pb/206Pbage", "204corr208Pb/232Thage", _
                 "207corr208Pb/232Thage", "208corr207Pb/206Pbage")
If Not pbFromSetup Then GetInfo
CheckForUPbWorkbook ' 10/03/30 added
ManCalc
NoUpdate

Set ns = ActiveSheet
Set SourceWbk = ActiveWorkbook
On Error Resume Next
Na = ActiveSheet.Name
If Na = "" Then End
pbStd = (Na = pscStdShtNa)

For Each Sht In ActiveWorkbook.Worksheets
  Na = LCase(Sht.Name)

  If Na = LCase(pscSamShtNa) Then
    Set phSamSht = Sht            ' 09/07/23 -- essential to avoid jumping to another
    Exit For                      '   workbook later (in some cases).
  End If

Next Sht

If Na <> LCase(pscSamShtNa) Then
  MsgBox "No " & pscSamShtNa & " sheet exists in this workbook."
  End
End If

On Error GoTo 0
Msg = "This workbook is not an intact SQUID-created U/Pb workbook."

For Each SsSht In Worksheets
  With SsSht

    If .Visible Then

      If .Cells(1, 5) = "SquidSampleData" Then
        Set SourceSht = SsSht
        .Activate: Cells(1, 1).Activate
        plHdrRw = flHeaderRow(0, , , True)

        If plHdrRw = 0 Then
          MsgBox Msg, , pscSq
          Exit Sub
        End If

        Cells(plHdrRw, 1).Activate
        Exit For
      End If

    End If

  End With
Next SsSht

If plHdrRw <= 2 Then GoTo NoSampleData

Set SourceSht = SsSht
piaNumSpots(0) = plaFirstDatRw(0) - plHdrRw
RefreshSampleNames
SourceSht.Activate

ReDim psaSpotNames(1 To piaNumSpots(0))

Again:

SourceWbk.Activate
On Error Resume Next
phSamSht.Activate ' 09/06/24
On Error GoTo 0

Group.Show

NoUpdate
ShowStatusBar
StatBar "Recalculating formulae"
foAp.Calculate
Alerts False
StatBar
i = 0
' Loop thru each specified name-frag

For NaCt = 1 To -pbGrpAll - (Not pbGrpAll) * UBound(psaGrpNames)
  Na$ = fsStrip(psaGrpNames(NaCt), pbIgCase, pbIgSpaces, pbIgDashes, pbIgSlashes)
  k = 0
  Ln = Len(Na$)

  If Ln = 0 And Not pbGrpAll Then GoTo Next_NaCt

  nSh$ = ""
  i = i + 1 ' Sample Data data-row counter
  Lrw = flEndRow
  plHdrRw = flHeaderRow(0, , , True)
  piIzoom = 75
  FindDriftCorrRanges , , , CorrCol

  For j = 1 + plHdrRw To Lrw
    wn$ = Left$(fsStrip(SsSht.Cells(j, 1).Text, pbIgCase, _
                pbIgSpaces, pbIgDashes, pbIgSlashes), Ln)

    If pbGrpAll Or (Not pbGrpAll And Na$ = wn$) Then  ' Sample-name frag to match
      k = 1 + k ' Extracted-sample count              ' old: (pbGrpAll And OK)
      StatBar wn$ & StR(k)

       If k = 1 Then
        Sheets.Add

        With Cells.Font
          .Name = psStdFont
          .Size = 11
        End With

        If pbGrpAll Then
          nSh = "AllSamples"
        Else
          nSh$ = UCase(fsStrip(psaGrpNames(NaCt), False, True, -1, -1, -1))
        End If

        Sheetname nSh$ ' Name of new name-grouped sample sheet
        GrpNum = fvMax(1, fvMin(12, 1 + Val(Right$(nSh, 2))))
        Set ns = ActiveSheet
        ns.Name = LTrim(fsLSN(nSh$))
        ns.Tab.Color = TabClrs(GrpNum)
        Zoom piIzoom
        SsSht.Rows(plHdrRw).Copy Destination:=ns.Rows(2) ' Copy hdr-row to new sample sht
      End If


      ManCalc
      ' crashes under Windows unless xlCalculationManual
      SsSht.Rows(j).Copy Destination:=ns.Rows(k + 2)

      SsSht.Cells(j, CorrCol).Copy
      Cells(k + 2, CorrCol).PasteSpecial xlPasteValues
    End If

  Next j

  SsSht.Rows(1).Copy
  Rows(1).PasteSpecial xlPasteFormats

  For Col = 1 To fiEndCol(plHdrRw)                          ' /
    If fbIsSquidHid(1, Col) Then Columns(Col).Hidden = True '|  09/07/21 -- added
  Next Col                                                  ' \

  plHdrRw = flHeaderRow(0, , , True)
  If plHdrRw = 0 Then GoTo NoSampleData

  If nSh$ = "" Or (Na$ = "" And Not pbGrpAll) Then GoTo Next_NaCt

  ns.Activate
  foAp.CutCopyMode = False
  NoGridlines
  Rows(plHdrRw).Font.Bold = True

  For Col = 1 To fiEndCol(plHdrRw) ' 09/07/21 -- added
    If Not Columns(Col).Hidden Then Columns(Col).AutoFit
  Next Col

  'Cells.EntireColumn.AutoFit
  g = piGrpDateType
  No4 = Not (g = 1 Or g = 4 Or g = 5)
  No7 = Not (g = 2 Or g = 6)
  No8 = Not (g = 3 Or g = 7)

  ' Find & delete unnecessary Pb46, Com%6(8), Radppm6(8),Pb86rad cols
  FindStr "7-corr204Pb/206Pb", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No7 Then Columns(Col).Delete
  FindStr "8-corr204Pb/206Pb", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No8 Then Columns(Col).Delete
  FindStr "4-corr%com206", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No4 Then Columns(Col).Delete
  FindStr "7-corr%com206", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No7 Then Columns(Col).Delete
  FindStr "8-corr%com206", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No8 Then Columns(Col).Delete
  FindStr "4-corr%com208", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No4 Then Columns(Col).Delete
  FindStr "7-corr%com208", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No7 Then Columns(Col).Delete
  FindStr "4-corr208Pb*/206Pb*", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No4 Then Columns(Col).Delete: Columns(Col).Delete
  FindStr "7-corr208Pb*/206Pb*", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No7 Then Columns(Col).Delete: Columns(Col).Delete
  FindStr "4-corrppm206*", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No4 Then Columns(Col).Delete
  FindStr "7-corrppm206*", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No7 Then Columns(Col).Delete
  FindStr "8-corrppm206*", , Col, plHdrRw, WholeWord:=True
  If Col And No8 Then Columns(Col).Delete
  FindStr "4-corrppm208*", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No4 Then Columns(Col).Delete
  FindStr "7-corrppm208*", , Col, plHdrRw, WholeWord:=True
  If Col > 0 And No7 Then Columns(Col).Delete

  ' Locate relevant col#s
  FindStr "Date/Time", , piDateTimeCol, plHdrRw, , plHdrRw
  FindStr "Hours", , piHoursCol, plHdrRw, , plHdrRw
  FindStr "Bkrd|cts|/sec", , piBkrdCtsCol, plHdrRw, , plHdrRw
  FindStr "204|/206", , piPb46col, plHdrRw, , plHdrRw
  FindStr "204|cts|/sec", , piPb204ctsCol, plHdrRw, , plHdrRw
  FindStr "206|cts|/sec", , piPb206ctsCol, plHdrRw, , plHdrRw

  FindStr "204/206", , piPb46col, plHdrRw, , plHdrRw
  FindStr "207/206", , piPb76col, plHdrRw, , plHdrRw
  FindStr "208/206", , piPb86col, plHdrRw, , plHdrRw
  FindStr "calibr.const", , FirstAcol, plHdrRw, , plHdrRw
  FindStr "total206Pb/238U", , piPb6U8_totCol, plHdrRw, , plHdrRw
  FindStr "total208Pb/232Th", , piPb8Th2_totCol, plHdrRw, , plHdrRw

  FindStr "204corr206Pb/238UAge", , piAgePb6U8_4col, plHdrRw, , plHdrRw
  FindStr "207corr206Pb/238UAge", , piAgePb6U8_7col, plHdrRw, , plHdrRw
  FindStr "208corr206Pb/238UAge", , piAgePb6U8_8col, plHdrRw, , plHdrRw
  FindStr "204corr207Pb/206PbAge", , piaAgePb76_4Col(0), plHdrRw, , plHdrRw
  FindStr "204corr208Pb/232ThAge", , piAgePb8Th2_4col, plHdrRw, , plHdrRw
  FindStr "207corr208Pb/232ThAge", , piAgePb8Th2_7col, plHdrRw, , plHdrRw
  FindStr "208corr207Pb/206PbAge", , piAgePb76_8col, plHdrRw, , plHdrRw
  FindStr "total206Pb/238U", , piPb6U8_totCol, plHdrRw, , plHdrRw
  FindStr "total208Pb/232Th", , piPb8Th2_totCol, plHdrRw, , plHdrRw
  FindStr "%Dis-cor-dant", , piDiscordCol, plHdrRw, , plHdrRw

  Frw = plaFirstDatRw(0)
  Lrw = plaLastDatRw(0)

  If pbExtractAgeGroups And piOverCtCorrType > 0 Then
    ' OverctCorr 4/6 = (204cpsTot)-BkrdCps-OverctCPS)/(206cpsTot-BkrdCps)

    For m = Frw To Lrw
      GetSpotGroupingInfo Cells(m, 1)
      temp = Cells(m, piPb46col)
      FinalTerm = "=1E-32"
      Term1 = "(" & fsAddr(m, piPb204ctsCol) & "-" & fsAddr(m, piBkrdCtsCol)
      Term2 = pscStdShtNa & "!" & "Pb204OverCts" & fsS(6 + piOverCtCorrType) & "corr"
      Term3 = "=" & Term1 & "-" & Term2 & ")"
      term4 = "/(" & fsAddr(m, piPb206ctsCol) & "-" & fsAddr(m, piBkrdCtsCol) & ")"
      term5 = fsFcell(Term3 & term4, m)

      If Not IsError(Evaluate(term5)) Then
        FinalTerm = term5
      End If

      Range(fsAddr(m, piPb46col)).Formula = FinalTerm

      If piNscans > 0 Then
         ' var tot204cts
        Term1 = fsAddr(m, piPb204ctsCol) & "/" & fsS(piNscans * pdaIntT(pi204PkOrder))
        Term2 = pscStdShtNa & "!Pb204OverCts7corrEr^2" ' var xs 204cps
        Term3 = "(1-" & fsAddr(m, piPb46col) & ")^2*" & fsAddr(m, piBkrdCtsCol) & "/" & _
                fsS(piNscans * pdaIntT(piBkrdPkOrder)) ' var bkrdcps
        term4 = "(" & fsAddr(m, piPb206ctsCol) & "-" & fsAddr(m, piBkrdCtsCol) & ")"
        term5 = "100*Sqrt(" & Term1 & "+" & Term2 & "+" & Term3 & ")/" & _
                 term4 & "/abs(" & fsAddr(m, piPb46col) & ")"
        Set Ce = Range(fsAddr(m, 1 + piPb46col))
        Ce.Formula = "=" & fsFcell(term5, m)

        If IsNumeric(Ce) Then
          temp = fvMax(0.000000001, fvMin(Ce, 9999))
        End If

        Ce.NumberFormat = "0"
      End If

    Next m

    Application.Calculate

    With Cells(Frw - 1, piPb46col)
      .Formula = fsVertToLF("overct|corr.|") & .Text
      Note .Row, .Column, "Corrected for excess 204 counts inferred from average 20" _
         & fsS(6 + piOverCtCorrType) & "-corrected Standard spots"
    End With
  End If

  If pbSortThis Then
    FindStr psGrpAgeTypeColName, , SortCol, Frw - 1
    ManCalc

    If Lrw > Frw And SortCol > 0 Then
      On Error GoTo NoSort
      frSr(Frw, 1, Lrw, fiEndCol(plHdrRw)).Sort _
           Key1:=Columns(SortCol), Order1:=xlAscending
      On Error GoTo 0
      GoTo 1

NoSort:  SortingError Err

      On Error GoTo 0
    End If

  End If

1
  If plHdrRw = 0 Then GoTo NoSampleData
  Cells(1, 1).Activate
  FindStr "Age", , BoxCol, plHdrRw, 2
  sCopyPaste "toSort", ThisWorkbook, "Squid", 1 + Lrw, 3
  Set ToSortBox = Selection
  ' Put reference to calibr const & ext %err
  fhSquidSht.Shapes("CalibrConst").Copy
  ActiveSheet.Paste
  Set CalibrConstBox = ActiveSheet.Shapes("CalibrConst")
  Set Ar = frSr(Lrw + 10, 5, Lrw + 12)
  Ar(1) = "=" & pscStdShtNot & "WtdMeanA1": Ar(2) = "=" & pscStdShtNot & "ExtPerrA1"
  Ar(3) = "=" & pscStdShtNot & "WtdMeanAperr1"
  Ar.NumberFormat = pscGen
  tL = Columns(Ar.Column).ColumnWidth
  Columns(Ar.Column).ColumnWidth = fvMax(tL, 7)
  CalibrConstBox.Left = Ar.Left - CalibrConstBox.Width
  CalibrConstBox.Top = Ar.Top - (CalibrConstBox.Height - Ar.Height) / 2
  CalibrConstBox.Fill.BackColor.RGB = peLightGray
  HA xlLeft, Ar
  Box Ar, , , , peLightGray


  If pbExtractAgeGroups Then '----------------------------------------------

    If piPb46col = 0 And piPb76col = 0 And piPb86col = 0 Then piGrpDateType = 0
    MinAgeCol = 9999

    For i = 0 To 6 ' Find the first-appearing age-type in the column headers
      FindStr GrpAgeNa(i + 1), , AgeCol, plHdrRw

      If AgeCol > 0 And AgeCol < MinAgeCol Then MinAgeCol = AgeCol 'Then Exit For
    Next i

    AgeCol = MinAgeCol

    If piGrpDateType > 0 And Not pbGrpCommPbSpecific Then
      ' (ie if grouping by age, with S-K defined C-Pb)

      ' 09/12/04 -- "And Not pbGrpCommPbSpecific" added to line above
      m = AgeCol - (piPb46col > 0) - (piPb76col > 0) - (piPb86col > 0)
      frSr(, AgeCol, , m).Insert Shift:=xlToRight
      SKageCol = AgeCol ' SKageCol is the NEW column containing the age
      m = AgeCol        ' used to calculate the S-K ratios.

      If piPb46col > 0 Then
        m = m + 1
        piSK64col = m
      End If

      If piPb76col > 0 Then
        m = m + 1
        piSK76col = m
      End If

      If piPb86col > 0 Then
        m = m + 1
        piSK86col = m
      End If

      Cells(plHdrRw, SKageCol) = fsVertToLF("Age|S-K|com|Pb")
      Note plHdrRw, SKageCol, "Age to use for Stacey-Kramers single-stage common-Pb calculation"
      tmp$ = "C-Pb" & vbLf & "20"

      If piPb46col > 0 Or piGrpDateType = 0 Then
        Cells(plHdrRw, piSK64col) = tmp$ & "6" & vbLf & "/204"
      End If

      If piPb76col > 0 Then Cells(plHdrRw, piSK76col) = tmp$ & fsVertToLF("7|/206")
      If piPb86col > 0 Then Cells(plHdrRw, piSK86col) = tmp$ & fsVertToLF("8|/206")
      Note plHdrRw, SKageCol + 1, "Stacey-Kramers single-stage crustal Pb isotope ratios"
    End If 'piGrpDateType > 0 And Not pbGrpCommPbSpecific

    'If pbExtractAgeGroups Then
      FindStr psGrpAgeTypeColName, , GrpAgetypeCol, plHdrRw
    'End If

    Fonts plHdrRw, SKageCol, , piSK86col, , False

    If piAgePb6U8_4col > 0 Then
      Term1 = "Common-Pb correction using "

      If piOverCtCorrType > 0 And piaOverCts4Col(7) Or piaOverCts4Col(8) Then
        Term2 = "overcount-corrected"
      Else
        Term2 = "measured"
      End If

      Note plHdrRw, piAgePb6U8_4col, Term1 & Term2 & " 204Pb" ' ???

      If piAgePb6U8_7col > 0 Then
        Note plHdrRw, piAgePb6U8_7col, pscCpbC & pscR6875
        Fonts plHdrRw, piAgePb6U8_7col, , , , False
      End If

      If piAgePb6U8_8col > 0 Then
        Note plHdrRw, piAgePb6U8_8col, pscCpbC & pscR6882
      End If

      If piAgePb8Th2_7col > 0 Then
        Note plHdrRw, piAgePb8Th2_7col, pscCpbC & pscR6875
      End If

      If piAgePb76_8col > 0 Then
        Note plHdrRw, piAgePb76_8col, pscCpbC & pscR6882
      End If

      If piDiscordCol > 0 Then
          Note plHdrRw, piDiscordCol, "Amount of zero-age Pb loss implied by " _
                                    & "206/238-207/206 age-discordance"
      End If

    End If 'piAgePb6U8_4col > 0

    psGrpAgeTypeColName = GrpAgeNa(1 + piGrpDateType)

    Lcol = fiEndCol(plHdrRw)

    With frSr(Frw, FirstAcol, Lrw, Lcol)
      .Replace pscStdShtNa & "!WtdMeanA1", Ar(1).Address
      .Replace pscStdShtNa & "!WtdMeanAperr1", Ar(2).Address
      .Replace pscStdShtNa & "!ExtPerrA1", Ar(3).Address
    End With

    If GrpAgetypeCol > 0 And Not pbGrpCommPbSpecific Then ' 09/12/04 -- added Not pbGrpCommPbSpecific
      AgeVal = 0

      SetSKageForCPb Frw, Lrw, GrpAgetypeCol, SKageCol, 1 ' 09/12/10 - replaces commented-out lines below

      For i = Frw To Lrw

'        TempVal = Cells(i, GrpAgetypeCol)
'
'        If IsNumeric(TempVal) Then
'          AgeVal = Val(Cells(i, GrpAgetypeCol))
'
'          If AgeVal < 0 Or AgeVal > 3700 Then
'            SK_Age = fvMin(fvMax(AgeVal, 0), 3700)
'          Else
'           SK_Age = CInt(fvMin(3700, fvMax(Cells(i, GrpAgetypeCol), 0)))
'          End If
'
'        Else
'          SK_Age = "0"
'        End If

        ' SK_Age is the Stacey-Kramers age to use for the spot.

'        If piGrpDateType > 0 Then ' i.e. if the  common-Pb ratios are defined by Stacey-Kramers age
         ' Cells(i, SKageCol).Formula = SK_Age
          SK_Age$ = "=SingleStagePbR(" & Cells(i, SKageCol).Address(False, True) & ", "
          a2$ = "/" & Cells(i, SKageCol + 1).Address(False, True)
          If piSK64col > 0 Then Cells(i, piSK64col) = SK_Age$ & "0)"       ' 206/204
          If piSK76col > 0 Then Cells(i, piSK76col) = SK_Age$ & "1)" & a2$ ' 207/206
          If piSK86col > 0 Then Cells(i, piSK86col) = SK_Age$ & "2)" & a2$ ' 208/206
'        End If

      Next i

      For i = 0 To 3
        m = -(i = 0) * SKageCol - (i = 1) * piSK64col - (i = 2) _
            * piSK76col - (i = 3) * piSK86col

        If m > 0 Then

          With Columns(m)
            tmp = pscZq & IIf(i, "." & String(2 - (i > 1), pscZq), "")
            RangeNumFor tmp, Frw, m, Lrw
            .ColumnWidth = 10
             Cells(plHdrRw, m).Font.Bold = True
             If m = SKageCol Then Cells(1 + Lrw, m) = "-1234"
            .EntireColumn.AutoFit
            If m = SKageCol Then Cells(1 + Lrw, m) = ""
          End With

        End If

      Next i

      Set CPbBox = Cells(1, 1) 'dummy

      For i = Frw To Lrw
        If piSK64col > 0 Then Sx = Cells(i, piSK64col).Address(False, True)
        If piSK76col > 0 Then Sp = Cells(i, piSK76col).Address(False, True)
        If piSK86col > 0 Then Sg = Cells(i, piSK86col).Address(False, True)

        With frSr(i, FirstAcol, , Lcol)
          If piSK64col > 0 Then .Replace psaC64(0), Sx, xlPart

          If piSK76col > 0 Then
            .Replace psaC76(0), Sp, xlPart
            If piSK64col > 0 Then .Replace psaC74(0), "(" & Sx & "*" & Sp & ")"
          End If

          If piSK86col > 0 Then
            .Replace psaC86(0), Sg
            If piSK64col > 0 Then .Replace psaC84(0), "(" & Sx & "*" & Sg & ")", xlPart
          End If

        End With

      Next i

      foAp.Calculate                                       ' 09/12/10 - added
      SetSKageForCPb Frw, Lrw, GrpAgetypeCol, SKageCol, 2  ' 09/12/10 - added

    ElseIf piAgePb6U8_4col > 0 Or piAgePb6U8_7col > 0 Or piAgePb6U8_8col > 0 Then
      ' pbGrpCommPbSpecific is True, so use specific common-Pb ratios
      Cells(1, 1).Select ' in case CalibrConst box still selected
      Rw = 2 + Ar(3).Row
      Col = Ar(3).Column
      [samcommpb].Copy Ar(5, 1)
      Set CPbBox = frSr(Rw, Col, Rw + 3, Col + [samcommpb].Columns.Count - 1)
      Set Com64grp = CPbBox(2, 3): Set Com76grp = CPbBox(3, 3)
      Set Com86grp = CPbBox(4, 3): Set Com74grp = CPbBox(2, 6)
      Set Com84grp = CPbBox(3, 6)

      If pbGrpCommPbSpecific Then
        Com64grp = foUser("GrpCPb64")
        Com76grp = foUser("GrpCPb76")
        Com86grp = foUser("GrpCPb86")
        Com64grp = fvMinMax(Val(Com64grp), 8, 1000)
        Com76grp = fvMinMax(Val(Com76grp), 0.04, 1.5)
        Com86grp = fvMinMax(Val(Com86grp), 0.0001, 1000)
        Com64grp.NumberFormat = pscZd3
        Com76grp.NumberFormat = pscZd3
        Com86grp.NumberFormat = pscZd3
        Com74grp.NumberFormat = pscZd2
        Com84grp.NumberFormat = pscZd2
      'End If

      Com64grp.Name = "Com64grp": Com74grp.Name = "Com74grp"
      Com84grp.Name = "Com84grp": Com76grp.Name = "Com76grp"
      Com86grp.Name = "Com86grp"
      If Com64grp = 0 Or Trim(Com64grp) = "" Then Com64grp = 18.3
      If Com76grp = 0 Or Trim(Com76grp) = "" Then Com76grp = 0.8536
      If Com86grp = 0 Or Trim(Com86grp) = "" Then Com64grp = 2.093
      Com74grp = Com64grp * Com76grp
      Com84grp = Com64grp * Com86grp
      FindStr "calibr.const", , Col, plHdrRw, 4, , , , , True, , True
       End If
    End If ' piGrpAgetypeCol > 0

    StatBar wn$ & StR(k)
    FindStr "total238/206", , piU8Pb6_totCol, plHdrRw, 1

' 09/07/30 -- lines below commented out
'    If piU8Pb6_totCol > 0 Then
'
'      If piAgePb6U8_7col Then
'        piAgePb6U8_7ecol = 1 + piAgePb6U8_7col
'        frSr(, piU8Pb6_totCol, , piU8Pb6_totCol + 1).Insert Shift:=xlToRight
'        With frSr(plHdrRw, piAgePb6U8_7col, , piAgePb6U8_7ecol)
'          .Copy Destination:=Cells(plHdrRw, piU8Pb6_totCol)
'        End With
'        piPb6U8_7col = piU8Pb6_totCol
'        piPb6U8_7ecol = 1 + piU8Pb6_totCol
'        Cells(plHdrRw, piU8Pb6_totCol) = fsVertToLF("7corr|206*|/238")
'        Cells(plHdrRw, piPb6U8_7ecol) = fsVertToLF(pscPpe)
'        frSr(Frw, piAgePb6U8_7col, Lrw, piAgePb6U8_7ecol).Copy
'        Cells(Frw, piU8Pb6_totCol).PasteSpecial Paste:=xlFormats
'        Fonts plHdrRw, piU8Pb6_totCol, , piPb6U8_7ecol, , True
'        Fonts Frw, piU8Pb6_totCol, Lrw, piPb6U8_7ecol, , False
'        piU8Pb6_totCol = 2 + piU8Pb6_totCol
'        Term1 = pscLm8 & "*" & pscEx8 & fsRa(Frw, piAgePb6U8_7col) & ")*" _
'                & fsRa(Frw, piAgePb6U8_7ecol)
'        FinalTerm = "=" & pscEx8 & fsRa(Frw, piAgePb6U8_7col) & ")-1"
'        PlaceFormulae FinalTerm, Frw, piPb6U8_7col, Lrw
'        PlaceFormulae "=100*" & Term1 & "/" & fsRa(Frw, piPb6U8_7col), _
'                       Frw, piPb6U8_7ecol, Lrw
'        RangeNumFor pscDd4, , piPb6U8_7col
'        RangeNumFor pscZd1, , piPb6U8_7ecol
'        ColWidth 20, piPb6U8_7col, piPb6U8_7ecol
'        ColWidth picAuto, piPb6U8_7col, piPb6U8_7ecol
'      End If
'
'    End If 'piU8Pb6_totCol > 0

    If piAgePb6U8_8col > 0 Then
      piAgePb6U8_8ecol = 1 + piAgePb6U8_8col
      FindStr "total238206", , piU8Pb6_totCol, plHdrRw, 3, plHdrRw

      If piU8Pb6_totCol > 0 Then
        frSr(, piU8Pb6_totCol, , piU8Pb6_totCol + 1).Insert Shift:=xlToRight
        With frSr(plHdrRw, piAgePb6U8_8col, , piAgePb6U8_8ecol)
          .Copy Destination:=Cells(plHdrRw, piU8Pb6_totCol)
        End With
        piPb6U8_8col = piU8Pb6_totCol
        piPb6U8_8ecol = 1 + piPb6U8_8col
        Cells(plHdrRw, piPb6U8_8col) = fsVertToLF("8corr|206*|/238")
        Cells(plHdrRw, piPb6U8_8ecol) = fsVertToLF(pscPpe)
        frSr(Frw, piAgePb6U8_8col, Lrw, piAgePb6U8_8ecol).Copy
        Cells(Frw, piPb6U8_8col).PasteSpecial Paste:=xlFormats
        Fonts plHdrRw, piPb6U8_8col, , piPb6U8_8ecol, , True
        Fonts Frw, piPb6U8_8col, Lrw, piPb6U8_8ecol, , False

        Term1 = "=" & pscEx8 & Cells(Frw, piAgePb6U8_8col).Address(0, 0) & ")-1"
        PlaceFormulae Term1, Frw, piPb6U8_8col, Lrw
        Term2 = "=100*" & pscLm8 & "*" & pscEx8 & Cells(Frw, piAgePb6U8_8col).Address(0, 0)
        Term3 = ")/" & Cells(Frw, piPb6U8_8col).Address(0, 0) & "*" & _
                Cells(Frw, piAgePb6U8_8ecol).Address(0, 0)
        PlaceFormulae Term2 & Term3, Frw, piPb6U8_8ecol, Lrw
        piU8Pb6_totCol = 2 + piU8Pb6_totCol

        RangeNumFor ".0000", Frw, piPb6U8_8col, Lrw
        RangeNumFor pscZd1, Frw, piPb6U8_8ecol, Lrw
        ColWidth 20, piPb6U8_8col, piPb6U8_8ecol
        ColWidth picAuto, piPb6U8_8col, piPb6U8_8ecol
      End If

    End If 'piAgePb6U8_8col > 0

    If pbGrpCommPbSpecific Then
      With frSr(Frw, Col, Lrw, Lcol)
        .Replace "sComm0_64", [Com64grp].Address, xlPart
        .Replace "sComm0_76", [Com76grp].Address, xlPart
        .Replace "sComm0_86", [Com86grp].Address, xlPart
        .Replace "sComm0_74", [Com74grp].Address, xlPart
        .Replace "sComm0_84", [Com84grp].Address, xlPart
      End With
    End If


    plHdrRw = flHeaderRow(0, , , True)
    TooFew = (Lrw <= (3 + plHdrRw))

    With Sheets(pscSamShtNa)
      LastRow = .Cells(pemaxrow, 1).End(xlUp).Row + 1

      For Col = 8 To fiEndCol(plHdrRw)

        If .Columns(Col).Hidden Then

          If foAp.CountA(frSr(LastRow, Col, 99 + LastRow)) = 0 Then
'            HideMarkCol 1, Col  ' 09/07/21 -- added  09/11/06 commented out -- what was I trying to do?
'            With Cells(1, Col).Borders(xlEdgeTop)
'              .Weight = xlHairline
'              .Color = vbWhite
'            End With
'            ActiveSheet.Columns(Col).Hidden = True
          End If

        End If

      Next Col

    End With


    If Not TooFew And pbExtractAgeGroups Then 'And Not pbGrpallNoAge

      StatBar "Extracting coherent group for " & nSh$
      Set ar2 = frSr(Frw, GrpAgetypeCol, Lrw, 1 + GrpAgetypeCol)
      ' ***********************************************************************
      ExtractGroup False, pdMinProb, ar2, AgeResult:=AgeVal, TypeCol:=GrpAgetypeCol
      ' ***********************************************************************
      StatBar
    End If

  Else   ' if not pbExtractAgeGroups
    Set CPbBox = [samcommpb]
  End If ' if pbExtractAgeGroups

  For i = 1 To 2
    FindStr "%com" & psaPDdaMass(i), , Col, plHdrRw

    If Col Then
      Box plHdrRw, Col, Lrw, , 13434828
      ColWidth picAuto, Col
    End If

  Next i

  FindStr "ppmU", , Col, plHdrRw
  If Col > 0 Then Box plHdrRw, Col, Lrw, , RGB(0, 255, 255)
  FindStr "ppmTh", , Col, plHdrRw
  If Col > 0 Then Box plHdrRw, Col, Lrw, , peLightGray
  FindStr "232Th/238U", , Col, plHdrRw
  If Col > 0 Then Box plHdrRw, Col, Lrw, , RGB(255, 255, 128)
  If Not pbExtractAgeGroups Then Set ar2 = Cells(1, 1)

  If Lrw > (3 + plHdrRw) And piPb6U8_totCol > 0 And piPb76col > 0 Then
    Call ConcordiaClr(ar2.Column)
  End If

  For i = 4 To 4 - 4 * pbDo8corr Step -4 * pbDo8corr
    tmp$ = fsVertToLF(fsS(i) & "corr|")
    FindStr tmp$ & fsVertToLF("238/|206*"), , Col, plHdrRw
    If Col > 0 Then Subst Cells(plHdrRw, Col), tmp$, ""
    FindStr tmp$ & fsVertToLF("207*|/206*"), , Col, plHdrRw
    If Col > 0 Then Subst Cells(plHdrRw, Col), tmp$, ""
    FindStr tmp$ & fsVertToLF("207*|/235"), , Col, plHdrRw
    If Col > 0 Then Subst Cells(plHdrRw, Col), tmp$, ""
    FindStr tmp$ & fsVertToLF("206*|/238"), , Col, plHdrRw
    If Col > 0 Then Subst Cells(plHdrRw, Col), tmp$, ""
    If Not pbDo8corr Then Exit For
  Next i

  If piPb46col > 0 And piPb76col > 0 And piPb6U8_totCol > 0 Then
    FindStr "238/206*", , Col, plHdrRw

    If Col > 0 Then

      If AgeVal = 0 Or AgeVal > 300 Then
        ReDim inpdat(1 To 1, 1 To 5)    ' If not DIMmed then sqConcAge crashes
        sqConcAge True      '  because the WtdXYmean call becomes illegal.
      End If

      If InStr(Cells(3 + Lrw, Col), "No coherent") = 0 Then
        piConcordPlotRow = 4 + Lrw
        piConcordPlotCol = Col - 1
        Cells(piConcordPlotRow, Col - 1) = "To recalculate Concordia Age using different spots,"
        Cells(1 + piConcordPlotRow, Col - 1) = "select desired rows from the red columns"
        Cells(2 + piConcordPlotRow, Col - 1) = "then press button at right."
        HA xlRight, 4 + Lrw, Col - 1, 6 + Lrw
        frSr(piConcordPlotRow, Col - 1, 6 + Lrw).IndentLevel = 1
        AddButton piConcordPlotRow - 3 + 0.5, Col + 0.5, "ConcAge", "Concordia Age", "sqConcAge", , 12
      End If

    End If

    Fonts 1, 1, , , , True, , 12, , , "Errors are 1s unless otherwise specified"
    On Error Resume Next
    Cells(1, 1).Characters(13, 1).Font.Name = "Symbol"
    On Error GoTo 0

  End If

  Rows(plHdrRw).Font.Bold = True
  frSr(plHdrRw, 1, flEndRow(1), 10).Columns.AutoFit

  If piPb204ctsCol > 0 Then
    Note plHdrRw, piPb204ctsCol, "Measured counts -- no overcount correction"
    FindStr "204/206", , piPb46col, plHdrRw
  End If

  FindStr "206/238", , Col, plHdrRw
  If Col > 0 Then Note plHdrRw, Col, "Raw 206Pb/238U -- Uncorrected for Pb-U sputtering bias"
  FindStr "total206Pb/238U", , Col, plHdrRw
  If Col > 0 Then Note plHdrRw, Col, "Total 206Pb/238U corrected for Pb-U sputtering bias"
  Rows(1 + Lrw).EntireRow.Insert
  With Rows(1 + Lrw)
    .RowHeight = 1: .ClearFormats
  End With

  If ActiveSheet.Name = pscStdShtNa And ActiveSheet.ChartObjects.Count > 0 Then
    foLastOb(ActiveSheet.ChartObjects).Left = fnRight(Columns(1 + GrpAgetypeCol))
  End If

  foAp.Calculate ' Essential!

' ********************************************************
  If foUser("GrpConcPlots") And Not TooFew Then
    SquidInvokedConcPlot ns, Bad
  End If
' ********************************************************

  If pbExtractAgeGroups Then CompareUOUwithStd GrpAgetypeCol '+ 5

  Cells(1, 10) = "SQUID grouped-sample sheet"

  If pbExtractAgeGroups Then
    BoxCol = fvMax(3, BoxCol - 4) ' 09/12/04 -- mod, move all to left
    PlaceGroupShtBoxes ToSortBox, CalibrConstBox, Ar, BoxCol, CPbBox
  End If

  If pbGrpCommPbSpecific And (piAgePb6U8_4col > 0 Or _
     piAgePb6U8_7col > 0 Or piAgePb6U8_8col > 0) Then 'piGrpSKageType
    ' 09/12/04 the "pbGrpCommPbSpecific above was "piGrpDateType = 0"
    On Error Resume Next
    Com64grp.NumberFormat = "0.00"
    Com74grp.NumberFormat = "0.00"
    Com84grp.NumberFormat = "0.00"
    Com76grp.NumberFormat = "0.000"
    Com86grp.NumberFormat = "0.00"
    On Error GoTo 0
  End If

  ScrollW 1, 1: Cells(Frw, 2).Activate

  If foUser("FreezeHeaders") Then
    Cells(Frw, 2).Activate
    Freeze
  End If

  If Not TooFew Then
    ScrollW fvMax(Frw, Lrw - 19), ar2.Column - 10
  End If

  Sheets(SsSht.Name).Activate
  Sct = 1 + Sct
Next_NaCt:

Next NaCt

If Sct = 0 Then
  MsgBox "None of the specified sample names were found.", , pscSq
  GoTo Again
End If

foUser("LastSample") = psSpotName
Zoom piIzoom
ns.Activate
On Error GoTo 0
StatBar
foAp.Calculate
StatBar
ClearObj ns, CPbBox, SourceSht, SsSht, Ar, ToSortBox, SourceWbk
ClearObj Com64grp, Com74grp, Com84grp, Com76grp, Com86grp
Exit Sub

NoSampleData:
MsgBox "Please switch to an intact SQUID-created SampleData sheet", , pscSq
End Sub

Sub AddGroupAgeChart(ByVal Rw, ByVal Co, GrpMean As Range, DoAll As Boolean, _
  StdCalc As Boolean, OkGrpCt%, BadGrpCt%, OkPtsCt%, BadPtsCt%, OkAgeAddr$(), BadAgeAddr$(), _
  OkAgeErAddr$(), BadAgeErAddr$(), OkAgeVals#(), OkAgeErrVals#(), _
  BadAgeVals#(), BadAgeErrVals#(), SigLev%, Hours As Range, Optional DpNum% = 1)

' Add small chart showing ages & errors of the Grouped spots, with rejected in blue.

Dim ChtName$
Dim i%, j%, EndStyle%, ErrThick%, NumChts%
Dim Ytik#, MinYval#, MaxYval#, mV#, ErT#
Dim w!, h!, L!, t!
Dim v As Variant, u As Variant
Dim Ac As Worksheet, SC As SeriesCollection, Ach As ChartObjects

Set Ac = ActiveSheet
Set Ach = Ac.ChartObjects

'If Not StdCalc Then
'  With Ac.Columns(Co - 1)
'    .ColumnWidth = 0.05   ' Can't hide the cols & still have
'    .Font.Color = vbWhite  '  their data plotted on a chart.
'  End With
'  frSr(rw - 1, Co - 2).Font.Color = vbWhite
'End If

Select Case OkGrpCt + BadGrpCt
  Case Is < 26: EndStyle = xlCap: ErrThick = xlThick
  Case Is < 52: EndStyle = xlCap: ErrThick = xlMedium
  Case Is < 99: EndStyle = xlCap: ErrThick = xlThin
  Case Else:    EndStyle = xlNoCap: ErrThick = xlThin
End Select

If StdCalc Then v = Co Else v = Co - 6

If Not DoAll Then
  NumChts = ActiveSheet.ChartObjects.Count
  If NumChts = 0 Then Exit Sub
  ChtName = "squidchart" & IIf(StdCalc, fsS(DpNum), "1")
  For i = 1 To NumChts
    If LCase(Ach(i).Name) = ChtName Then Exit For
  Next i

  If i > NumChts Then Exit Sub
  Ach(ChtName).Activate

Else
  w = 300: h = 250
  L = Columns(Co).Left - StdCalc * Columns(Co).Width
  t = Rows(Rw).Top + 5

  If StdCalc Then
    L = L - w
    t = t - 8
  Else
    L = L - w
  End If

  Ac.ChartObjects.Add(L, t, w, h).Select

  With Selection
    .Left = L: .Top = t: .Width = w: .Height = h
    .Placement = xlFreeFloating
    With .Border: .LineStyle = 0: .ColorIndex = xlNone: End With
    .Name = "SquidChart" & fsS(Ac.ChartObjects.Count)
    psLastSquidChartName = .Name
  End With

  With ActiveChart
    .ChartType = xlXYScatter: .HasLegend = False
    PlaceDataSeries OkAgeAddr, OkGrpCt, "ok", True
    .PlotArea.Interior.Color = 13434879 ' Straw

    With .Axes(xlCategory)
      .HasMajorGridlines = False: .HasMinorGridlines = False: .MinimumScale = 0

      If pbStd Then
        .MaximumScale = 1 + Int(Hours(Hours.Rows.Count))
      Else
        .MaximumScale = OkGrpCt + BadGrpCt + 1
      End If

      .MajorTickMark = xlCross:   .MinorTickMark = xlInside
      .TickLabelPosition = xlTickLabelPositionNextToAxis

      If StdCalc Then
        .HasTitle = True
        With .AxisTitle
          .AutoscaleFont = False
          .Caption = "Hours": .Font.Size = 12: .Font.Bold = False
        End With
      End If

      With .TickLabels
        .AutoscaleFont = False
        .Font.Size = 11: .Font.Name = psStdFont
      End With

    End With

    With .Axes(xlValue)
      .HasMajorGridlines = False: .HasMinorGridlines = False
      .MajorTickMark = xlCross:   .MinorTickMark = xlInside
      .TickLabelPosition = xlTickLabelPositionNextToAxis

      If StdCalc Then
        .HasTitle = True

        With .AxisTitle
          .Caption = "Std Age (Ma)"
          .AutoscaleFont = False
          .Font.Size = 12: .Font.Bold = False
        End With

      End If

      With .TickLabels
        .AutoscaleFont = False
        .Font.Size = 11: .Font.Name = psStdFont
      End With

    End With

  End With
End If

Set SC = ActiveChart.SeriesCollection

If Not DoAll Then
  On Error Resume Next
  j = SC.Count

  For i = 1 To j
    SC(1).Delete
  Next i

  On Error GoTo 0
  PlaceDataSeries OkAgeAddr, OkGrpCt, "ok", True
  With foLastOb(ActiveChart.SeriesCollection)
    .MarkerStyle = xlNone
    .Border.LineStyle = xlNone
  End With
  On Error GoTo 0
End If

If BadGrpCt > 0 Then
   PlaceDataSeries BadAgeAddr, BadGrpCt, "bad", False
End If

ActiveChart.SeriesCollection.Add GrpMean, xlColumns, False, True, False
FormatSeriesCol foLastOb(SC), , xlContinuous, vbGreen, xlMedium
With foLastOb(SC)
  .Name = "AverageLine": .MarkerStyle = xlNone
End With

With ActiveChart.Axes(1)
  .MinimumScale = 0: .MaximumScale = GrpMean(2, 1)
End With

With ActiveChart

  For i = 1 To OkGrpCt
    ErrBars fsS(i) & "ok", OkAgeErAddr(i), vbRed, EndStyle, ErrThick
  Next i

  If BadGrpCt > 0 Then
    For i = 1 To BadGrpCt
      ErrBars fsS(i) & "bad", BadAgeErAddr(i), vbBlue, EndStyle, ErrThick
    Next i
  End If

  On Error Resume Next
  For i = 1 To SC.Count
    SC(i).MarkerStyle = xlNone
  Next i
  On Error GoTo 0

  MinYval = 9999: MaxYval = 0
End With

For j = 1 To OkPtsCt
   v = OkAgeVals(j, 2)
   u = OkAgeErrVals(j)

  If v > 0 And u > 0 Then
    ErT = Abs(u / v): v = fvMax(pdcTiny, v)
    If ErT > 1 And v > 1 Then u = v / 2 ' To deal with huge error bars
    If MinYval > 0 Then MinYval = fvMin(MinYval, v - u)
    If MaxYval <= 2 * GrpMean(1, 2) Then MaxYval = fvMax(MaxYval, v + u)
  End If

Next j

For j = 1 To BadPtsCt
  v = BadAgeVals(j, 2)
  u = BadAgeErrVals(j)

  If v > 0 And u > 0 Then
    ErT = Abs(u / v): v = fvMax(pdcTiny, v)
    If ErT > 1 And v > 1 Then u = v / 2 ' To deal with huge error bars
    If MinYval > 0 Then MinYval = fvMin(MinYval, v - u)
    If MaxYval <= 2 * GrpMean(1, 2) Then MaxYval = fvMax(MaxYval, v + u)
  End If

Next j

With ActiveChart.Axes(xlValue)
  .MinimumScaleIsAuto = True: .MaximumScaleIsAuto = True
  Tick MaxYval - MinYval, Ytik
  mV = 0

  Do
    mV = mV + Ytik
  Loop Until mV > MinYval

  MinYval = mV - Ytik

  Do
    mV = mV + Ytik
  Loop Until mV > MaxYval

  MaxYval = mV
  .MinimumScale = MinYval: .MaximumScale = MaxYval
  .TickLabels.Font.Size = 10 ' keep this last, otherwise will autoresize
  .TickLabels.AutoscaleFont = False
  .MajorUnitIsAuto = True: .MinorUnitIsAuto = True
   TickFor
  If DoAll Then Call TwoSigText(2)
  .TickLabels.Font.Size = 10 ' keep this last, otherwise will autoresize
End With

With ActiveChart ' Kluge to correct an intermittent Excel bug
  If .TextBoxes.Count = 0 Then Call TwoSigText(1)
End With
'psLastSquidChartName = ActiveChart.Name

1: On Error GoTo 0
Ac.Cells(Rw + 5, Co).Select
ClearObj Ac, SC
End Sub

Sub ReCalcNow()
foAp.Calculate
ActiveSheet.Shapes("ReCalcNow").Delete
End Sub

Sub Redo()
' Force recalculation of all wtd averages, and recognize any struck-through cells as rejected.

Dim StdCalc As Boolean, Struck As Boolean
Dim f$, v$, Cname$
Dim i%, j%, p%, ChrtYcol%, r0%, C0%, r&, c%, ciN%, Nu%, DpNum%, acC%, EffNumDauPar%
Dim Rin&, Bclr1&, Bclr2&, acR&
Dim wr As Range, qR As Range, Ob As Object

Static NoShow As Boolean

If ActiveSheet.Type <> xlWorksheet Then
  MsgBox "Only relevant for a SQUID-created worksheet", , pscSq
  End
End If

StatBar "Wait"
With ActiveWindow: Rin = .ScrollRow: ciN = .ScrollColumn: End With
With ActiveCell:   acR = .Row: acC = .Column:             End With
ManCalc
NoUpdate
' Auto calculation really screws things up -- Excel cannot be relied on to recalculate
'  worksheets from the Calculate command from within VBA.
GetInfo False
StdCalc = (InStr(ActiveSheet.Name, pscStdShtNa) > 0)
plHdrRw = flHeaderRow(StdCalc)
EffNumDauPar = 1 - (StdCalc And fbRangeNameExists("aadat2", , Sheets(pscStdShtNa)))

Do
  For DpNum = 1 To EffNumDauPar

    For Each Ob In ActiveSheet.ChartObjects
      Cname = Ob.Name

      If LCase(Left$(Cname, 10)) = "squidchart" Then

        If Not StdCalc Or Val(Right$(Cname, 1)) = DpNum Then
          v = Ob.Chart.SeriesCollection("1ok").Formula
          p = InStr(v, "!"): v = Mid$(v, 1 + p)
          Subst v, "'" & ActiveSheet.Name & "'", "$"
          j = InStr(v, ":"): p = InStr(v, ",")

          If j = 0 Then
            j = p
          ElseIf p = 0 Then
            p = j
          End If

          If p = 0 Then ChrtYcol = 0: Exit For
          p = fvMin(p, j)
          v = Left$(v, p - 1)
          If Left$(v, 1) = "R" Then v = foAp.ConvertFormula(v, xlR1C1, xlA1)
          ChrtYcol = 1 + Range(v).Column
          Exit For
        End If

      End If

    Next Ob

    v = IIf(StdCalc, "Wtd Mean of", "Mean age of")
    Cells(1, 1).Select

    If ActiveCell.Row = 1 And ActiveCell.Column = 1 Then
      On Error GoTo NoWtdAv ' Find first WtdAv range

      If StdCalc Then
        f = "WtdMeanA" & fsS(DpNum)
        r = Range(f).Row: c = Range(f).Column
        Cells(r, c).Select
      Else
        r = 1: c = 1
        Cells(r, c).Select
        Cells.Find("Mean age of", Cells(r, 1), xlFormulas, xlPart).Activate
        '10/03/30 - mod
        Cells(ActiveCell.Row, ActiveCell.Column + 1).Activate ' was + 2).Activate
      End If

      On Error GoTo 0
      r0 = ActiveCell.Row: C0 = ActiveCell.Column
      If r0 = 1 And C0 = 1 Then Exit Sub ' None found in sheet
       r = r0: c = C0

    Else ' Find next WtdAv range
      Cells(r, c).Select
      Cells.FindNext(After:=ActiveCell).Activate
      r = ActiveCell.Row
      c = ActiveCell.Column
      If r = r0 And c = C0 Then Exit Do ' Back to 1st one, so no more
    End If

    f = Cells(r, c).Formula

    If r <> 1 And Cells(r - 1, c).Formula = "" Then ' ie top row of the WtdAv range
      Set wr = frSr(plaFirstDatRw(-StdCalc), c, plaLastDatRw(-StdCalc)) ' the data range
      Bclr1 = Cells(plHdrRw, c).Interior.Color
      Bclr2 = Cells(plHdrRw, c + 2).Interior.Color

      For i = 1 To wr.Rows.Count ' Strikethrough , so don't use
        Struck = False

        For j = 1 To 2 - 2 * StdCalc
          If wr(i, j).Font.Strikethrough Then
            Struck = True: Exit For
          End If
        Next j

        If Struck Then
          Set qR = Range(wr(i, 1), wr(i, 2 - 2 * StdCalc))

          If fbIsNum(qR(1, 1), True) Then
            IntClr vbYellow, qR
            Fonts qR, , , , , , , , True
          End If

        Else
          With Range(wr(i, 1), wr(i, 2))
            .Interior.Color = Bclr1: .Font.Strikethrough = False
          End With

          If StdCalc Then
            With Range(wr(i, 3), wr(i, 4))
              .Interior.Color = Bclr2: .Font.Strikethrough = False
            End With
          End If

        End If

      Next i

      Nu = 0
      For i = 1 To wr.Count
        If Not wr(i).Font.Strikethrough Then Nu = 1 + Nu
      Next i

      If Nu < 2 Then
        MsgBox "Insufficient data to average", , pscSq: Exit Sub
        j = wr.Rows.Count

        With Range(wr(1, 1), wr(j, 3))
          With .Interior

            If StdCalc Then
              .Color = Cells(plHdrRw, c).Color
            Else
              .ColorIndex = xlNone
            End If

           End With
          .Font.Strikethrough = False
        End With

        End
      End If

      On Error GoTo 0 '****************

      ExtractGroup StdCalc, 0, wr, True, StdCalc, , , DpNum

      If Not StdCalc Then
        With wr
          plHdrRw = .Row - 1
          plaFirstDatRw(-StdCalc) = .Row
          plaLastDatRw(-StdCalc) = .Row + .Rows.Count - 1
          ConcordiaClr .Column
        End With
      End If

    End If

  Next DpNum

Loop Until True

NoWtdAv: On Error GoTo 0
StatBar
ActiveWindow.ScrollRow = Rin: ActiveWindow.ScrollColumn = ciN
Cells(acR, acC).Activate
foAp.Calculate
ClearObj wr, qR, Ob
End Sub

Sub ColIndx(ByVal ColNum%, ByVal ColHeader, ByVal ColVarName, Optional Index, _
  Optional ErrColNum As Boolean = False, Optional ErrColName)
Dim s$, r&, e$, t$
Dim i%, Csht As Worksheet

' Assigns an index-name for the column and the matching col# of the output sheet.
' Index & Ix2 are indexes for the U-Pb variables such as ThU(0)

Set Csht = ThisWorkbook.Sheets("ColIndex")

With Csht

  If .Cells(1, 1) = "" Then
    r = 1
  Else
    r = .Cells(pemaxrow, 1).End(xlUp).Row + 1
    If ColNum = 0 Then ColNum = 1 + Cells(r - 1, 2)
  End If

  For i = 1 To 1 - ErrColNum
    s = ColVarName
    If i = 2 Then s = s & "e"

    If fbNIM(Index) Then
      s = s & "(" & fsS(Index) & ")"
    End If

    .Cells(r, 1) = LCase(s)
    .Cells(r, 2) = ColNum
    .Cells(r, 3) = Choose(i, ColHeader, ErrColName)
    ColNum = 1 + ColNum
    r = 1 + r
  Next i

End With
End Sub

Sub RefreshColIndx(ByVal IndxName$, ByVal CurrentColNum%, Optional NotPresent As Boolean = False)
Dim Rw&, IndexSht As Worksheet, ShtIn As Worksheet
If IndxName = "" Then Exit Sub
Set IndexSht = ThisWorkbook.Sheets("ColIndex")
Set ShtIn = ActiveSheet
IndexSht.Activate
FindStr Phrase:=IndxName, RowFound:=Rw, RowLook1:=1, RowLook2:=999, ColLook2:=1

If Rw = 0 Then
  NotPresent = True
Else
  IndexSht.Cells(Rw, 2) = Rw
End If

ShtIn.Activate
End Sub

Sub PlaceDataSeries(SourceRangeAddr$(), ByVal NumGrps%, ByVal SeriesNameFrag$, _
    ByVal FirstSer As Boolean)

Dim i%, ScCt%

With ActiveChart

  If FirstSer Then
    .SetSourceData Range(SourceRangeAddr(1)), xlColumns
  End If

  With .SeriesCollection
    ScCt = .Count + FirstSer
    On Error Resume Next

    For i = 1 - FirstSer To NumGrps
      .Add Range(SourceRangeAddr(i)), xlColumns, False, True, False
    Next i

    For i = 1 To NumGrps
      .Item(i + ScCt).Name = fsS(i) & SeriesNameFrag
    Next i

  End With
End With
On Error GoTo 0
End Sub

Sub GetSpotGroupingInfo(SpotName$)
' Determine #scans, #peaks, 204 & 206 pk-order of a spot for Grouping
Dim Na$, s$, Pk%, pk1%, pk0%, Npks%, Col%, Rw&
Dim Delt#, Delt4#, DeltBkrd#, Delt6#, m#, Mass#()
Dim ShtIn As Worksheet, Ce As Range

Set ShtIn = ActiveSheet
phCondensedSht.Activate
Na = SpotName & ",  "
Set Ce = Columns(picNameDateCol).Find(Na, Cells(1, picNameDateCol), xlFormulas, xlPart)

If Ce Is Nothing Then
  piNscans = 0
Else
  Rw = Ce.Row
  s = Cells(Rw, picPksScansCol)
  Npks = Val(s)
  piNscans = Val(Mid$(s, 2 + InStr(s, ",")))

  If piNscans = 0 Then
    ShtIn.Activate
    Exit Sub
  End If

  ReDim Mass(Npks), pdaIntT(Npks)

  For Pk = 1 To Npks
    Col = picDatCol + 1 + 5 * (Pk - 1)
    Mass(Pk) = Val(Cells(Rw + 3, Col))
    pdaIntT(Pk) = Val(Cells(Rw + 2, Col - 1))
  Next Pk

  Delt4 = 999: Delt6 = 999
  pi204PkOrder = 0: piBkrdPkOrder = 0: pi206PkOrder = 0

  ' Find the Run-Table peak-order of 204 and 206 (if present)
  For Pk = 1 To Npks
    Delt = Abs(Mass(Pk) - 204)
    If Delt < Delt4 Then Delt4 = Delt: pk0 = Pk
    Delt = Abs(Mass(Pk) - 206)
    If Delt < Delt6 Then Delt6 = Delt: pi206PkOrder = Pk
  Next Pk

  pk1 = 1 + pk0
  Delt = Mass(pk1) - Mass(pk0)

  If Abs(Delt) > 0.3 Then
    pk1 = pk0
    pk0 = pk1 - 1
    Delt = Mass(pk0) - Mass(pk1)
  End If

  If Abs(Delt) <= 0.3 Then

    If Delt > 0 Then
      pi204PkOrder = pk1
      piBkrdPkOrder = pk0
    Else
      piBkrdPkOrder = pk1
      pi204PkOrder = pk0
    End If

  Else
    piNscans = 0
  End If

End If
ShtIn.Activate
End Sub

Sub CompareUOUwithStd(PlaceCol%)
' Add chart-inset showing std UO/U (or UO2/U or UO2/UO) versus grouped-spot ditto.
' 09/06/12 -- Increase #possible UOx/UOy combinations to 6 by including reciprocals.
Dim Bad As Boolean, NoCht As Boolean
Dim TopX$, TopY$, UOUaddr$(0 To 1), UOUperrAddr$(0 To 1), Hrs$(0 To 1), Gyerr$, Yerr$
Dim Typ%, ct%, CelCt%, Capt%, Sh%, Col%, UOtype%, UOUcol%(0 To 1), HrsCol%(0 To 1)
Dim Hrow&, PlaceRow&, Rw&, r&, fr&(0 To 1), Lr&(0 To 1), CaptLeft!, Bott!, Topp!
Dim UOU#, UOUperr#, MedYerr#, MinYval#, MaxYval#, MedYerr0#, MedYerr1#
Dim MedYval0#, MedYval1#, MinXval#, MaxXval#, MinX#, MaxX#, Xtik#
Dim UOiso As Variant, UOtitle As Variant, GrpYerr As Variant
Dim Cel1 As Range, Cel2 As Range, UOUra(0 To 1) As Range, UOUperrRa(0 To 1) As Range
Dim HrsRa(0 To 1) As Range, XYra(0 To 1) As Range
Dim Sht(0 To 1) As Worksheet, UOUobj As ChartObject, UOUcht As Chart, GrpCht As ChartObject
Dim UOUplot As Object, SC As Series, ChtObjs As ChartObjects, WtdMeanCht As ChartObject

Const AsPic = False, ErrBars = False, EndCap = False
' Must add code to create absolute UO/U error column if want errorbars.

If Not foUser("CompareGroupedUOUwithStd") Then Exit Sub
UOiso = Array("254/238", "270/238", "270/254", "238/254", "238/270", "254/270")
UOtitle = Array("UO/U", "UO2/U", "UO2/UO", "U/UO", "U/UO2", "UO/UO2")

NoUpdate
Set Sht(0) = ActiveSheet
Set Sht(1) = Sheets(pscStdShtNa)

For Sh = 0 To 1
  Sht(Sh).Activate
  Hrow = flHeaderRow(-Sh)

  For Typ = 1 To UBound(UOiso) ' Find 254/238, 270/238, or 270/254 column
    FindStr UOiso(Typ), , UOUcol(Sh), Hrow
    If UOUcol(Sh) > 0 Then Exit For
  Next Typ

  If UOUcol(Sh) = 0 Then Exit Sub ' none of the 3 UOx/UOy ratios present
  FindStr "Hours", , HrsCol(Sh), Hrow
  If HrsCol(Sh) = 0 Then Exit Sub
  fr(Sh) = plaFirstDatRw(Sh) ' 1st, last data-rows
  Lr(Sh) = plaLastDatRw(Sh)

  Set UOUra(Sh) = frSr(fr(Sh), UOUcol(Sh), Lr(Sh)) ' UOx/UOy range
  Set UOUperrRa(Sh) = frSr(fr(Sh), 1 + UOUcol(Sh), Lr(Sh)) ' %err range
  Set HrsRa(Sh) = frSr(fr(Sh), HrsCol(Sh), Lr(Sh))   ' hours range

  CelCt = 0
  For Rw = fr(Sh) To Lr(Sh)
    Set Cel1 = Cells(Rw, UOUcol(Sh))
    Set Cel2 = Cells(Rw, 1 + UOUcol(Sh))

    If Not (IsNumeric(Cel1) And IsNumeric(Cel2)) Then
      Cel1 = "": Cel2 = ""
    ElseIf Cel1 = 0 Or Cel2 = 0 Then
      Cel1 = "": Cel2 = ""
    Else
      CelCt = 1 + CelCt
    End If

  Next Rw

  If CelCt < 2 Then Exit Sub

  UOUaddr(Sh) = UOUra(Sh).Address
  UOUperrAddr(Sh) = UOUperrRa(Sh).Address
  TopX = HrsRa(Sh).Address(0, 0)
  TopY = UOUra(Sh).Address(0, 0)
  Set XYra(Sh) = Range(TopX & "," & TopY) ' the grouped x-y range
Next Sh

On Error GoTo Done
With foAp
  MedYval0 = .Median(UOUra(0))
  MedYval1 = .Median(UOUra(1))
  MedYerr0 = .Median(UOUperrRa(0)) / 100 * MedYval0
  MedYerr1 = .Median(UOUperrRa(1)) / 100 * MedYval1
  MedYerr = (MedYerr0 + MedYerr1) / 2
  MinYval = .Min(.Min(UOUra(0), .Min(UOUra(1))))
  MaxYval = .Max(.Max(UOUra(0), .Max(UOUra(1))))
  MinYval = MinYval - MedYerr
  MaxYval = MaxYval + MedYerr
End With

If PlaceCol = 0 Then PlaceCol = UOUcol(0) ' chart-left column


' Construct the chart inset and the std UOx/UOy error bars

SmallChart DataRange:=XYra(1), DataSheet:=Sht(1), PlaceSheet:=Sht(0), Xname:="Hours", _
   Yname:=UOtitle(Typ), Symbol:=xlCircle, YerrBars:=ErrBars, YerrRange:=UOUperrRa(1), _
   ErrBarsClr:=vbBlue, ErrBarsThick:=xlThin, ErrBarsCap:=EndCap, FontAutoScale:=False, _
   PlaceRow:=Lr(0), PlaceCol:=PlaceCol, AxisNameSize:=12, PercentErrs:=True, _
   TikLabelSize:=10, BadPlot:=Bad, SymbInteriorClr:=vbCyan, SymbolSize:=7
If Bad Then Exit Sub

On Error GoTo 0

' 09/12/04 -- Code below, specifying Left/Top of the UOU chart, rewritten
Set UOUobj = foLastOb(ActiveSheet.ChartObjects)
Set UOUcht = UOUobj.Chart
Set ChtObjs = ActiveSheet.ChartObjects

For ct = ChtObjs.Count To 1 Step -1  ' find the Group Wtd-Mean chart
  Set WtdMeanCht = ChtObjs(ct)
  If LCase(WtdMeanCht.Name) = "squidchart1" Then Exit For
Next ct

With UOUobj
  If ct <= 0 Then ' ie no wtd-mean chart
    PlaceRow = 9 + Lr(0)
    r = flEndRow(1)
    FindStr "coherent group", Rw, Col, r, , 10 + r
    If Rw = 0 Or Col = 0 Then FindStr "No coherent ", Rw, Col, r, , 10 + r
    If Col = 0 Then Col = 10 ' 09/12/18 -- added
    .Left = Columns(Col).Left
    Rw = flEndRow(Col) + 3
    .Top = Rows(Rw).Top
  Else ' Put the UOU chart just below the Wtd Mean chart.
    Bott = fnBottom(WtdMeanCht)
    Topp = Bott
    PlaceRow = fnLeftTopRowCol(1, Topp)
    .Top = Topp
    .Left = WtdMeanCht.Left
  End If
End With
' ------------------------------------------------------ end rewrite

Gyerr = ""
ct = Lr(0) - fr(0) + 1
If ct < 3 Then Exit Sub
On Error GoTo Done

For Rw = fr(0) To Lr(0) ' assemble UOx/UOy-error array for grouped spots
  UOU = 0: UOUperr = 0: Yerr = 0
  On Error Resume Next
  UOU = Cells(Rw, UOUcol(0))
  UOUperr = Cells(Rw, 1 + UOUcol(0))
  Yerr = fsS(Drnd(UOUperr / 100 * UOU, 2))
  On Error GoTo 0

  If Rw = fr(0) Then
    Gyerr = Yerr
  Else
    Gyerr = Gyerr & "," & Yerr
  End If

Next Rw

If ct > 30 Then ' Crashes if large# of specific err-values used
  GrpYerr = MedYerr
Else
  GrpYerr = Gyerr
End If

With UOUobj   ' Add the grouped UOx/UOy error bars
  .Name = "UOUchart_" & Sht(0).Name
  With .Chart
    .SeriesCollection.Add Source:=XYra(0), Rowcol:=xlColumns, _
        SeriesLabels:=False, CategoryLabels:=True, Replace:=False
    Set SC = foLastOb(.SeriesCollection)
    FormatSeriesCol SC, , , , , , xlCircle, 7, 0, vbRed ' RGB(256, 92, 55)

    If ErrBars Then
      FormatErrorBars SC, 2, GrpYerr, vbRed, xlThin, False
    End If

    UOUobj.Activate

    MinXval = foAp.Min(XYra(0))
    MaxXval = foAp.Max(XYra(0))

    With .Axes(1)

      Do While .MinimumScale > MinXval
        .MinimumScale = .MinimumScale - .MajorUnit
      Loop

      Do While .MaximumScale < MaxXval
        .MaximumScale = .MaximumScale + .MajorUnit
      Loop

      .CrossesAt = .MinimumScale
    End With

    AxisScale UOUra(0), False, , MinYval, MaxYval   ', Optional ErrorRange)

    With ActiveChart.Axes(2)
      .CrossesAt = ActiveChart.Axes(1).MinimumScale
      .TickLabels.NumberFormat = "General"
    End With

    With .PlotArea
      .Top = 12: .Height = .Height - 5
    End With

    With .Axes(2)
      .MajorUnit = 2 * .MajorUnit
      If .MinimumScale < 0 Then .MinimumScale = 0
      .CrossesAt = .MinimumScale
      .TickLabels.NumberFormat = "General"
      .HasMajorGridlines = True
      .MajorGridlines.Border.Color = Hues.peGray
    End With

    For Capt = 1 To 2
      CaptLeft = UOUcht.Axes(1).Left + 15 + Capt * 40
      With .TextBoxes.Add(CaptLeft, 0, 12, 22)
        .AutoSize = True
        .Font.Size = 13
        .Font.Color = Choose(Capt, vbBlue, vbRed)
        .Text = Choose(Capt, "Std", "Grouped")
      End With
    Next Capt

    If AsPic Then   ' convert the Excel chart to a picture
      .CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
    End If

  End With
  If AsPic Then .Delete    ' (the chart)
End With

ActiveSheet.Cells(PlaceRow, PlaceCol).Select
NoCht = True
On Error GoTo 1
Set GrpCht = ActiveSheet.ChartObjects(psLastSquidChartName)
NoCht = False
1 On Error GoTo 0

If AsPic Then
  Sht(0).Pictures.Paste
  Set UOUplot = foLastOb(ActiveSheet.Pictures)
Else
  Set UOUplot = UOUobj
End If

'With UOUplot ' 09/12/04 commented out
'  .Top = Rows(flEndRow + 4).Top
'  On Error Resume Next
'
'  If NoCht Then
'    .Left = Columns(33).Left
'  Else
'  .Left = GrpCht.Left - .Width 'fnRight(GrpCht)
'  .Top = GrpCht.Top + 15 ' 09/11/06 -- changed from +5 to +15
'  End If
'
'  On Error GoTo 0
'End With

Cells(PlaceRow, PlaceCol).Select
Done: On Error GoTo 0
End Sub

Sub CheckHeadersForCommPbCorrTypes(Ok4UPb As Boolean, OK4ThPb As Boolean, OK4PbPb As Boolean, _
  Ok7UPb As Boolean, Ok8UPb As Boolean, Ok7ThPb As Boolean)
Dim HasUPb As Boolean, HasThPb As Boolean
Dim Hdr$(), i%, Nh%, Hr&, Arr As Variant

Hr = flHeaderRow(False)
Arr = Array("204/206", "207/206", "208/206", "206/238", "206/254", "206/270", "208/232", "208/248", "208/264")
Nh = UBound(Arr)
ReDim Col(Nh)

For i = 1 To Nh
  FindStr Arr(i), , Col(i), Hr, , Hr
Next i

HasUPb = (Col(4) > 0 Or Col(5) > 0 Or Col(6) > 0)
HasThPb = (Col(7) > 0 Or Col(8) > 0 Or Col(9) > 0)

Ok4UPb = (Col(1) > 0)
Ok7UPb = (Col(2) > 0 And HasUPb)
Ok8UPb = (Col(3) > 0 And HasThPb And HasUPb)
OK4PbPb = (Col(1) > 0 And Col(2) > 0)
Ok7ThPb = (Col(2) > 0 And HasUPb And HasThPb)
End Sub

Sub SortingError(ErrorNumber)
On Error GoTo 0
If ErrorNumber = 1004 Then
  MsgBox "Can't sort sheets containing Array Equations", , pscSq
End If
End Sub

Sub SetSKageForCPb(ByVal FirstRw%, ByVal LastRw%, ByVal GrpAgetypeCol%, _
                   ByVal SKageCol%, ByVal NumIters%)  ' Added 09/12/10
Dim Iter%, i%, AgeVal#, TempVal As Variant, SK_Age#
' Iteratively polish the Stacey-Kramers common-Pb age for goruped samples

For Iter = 1 To NumIters

  For i = FirstRw To LastRw
    TempVal = Cells(i, GrpAgetypeCol)

    If IsNumeric(TempVal) Then
      AgeVal = Val(Cells(i, GrpAgetypeCol))

      If AgeVal < 0 Or AgeVal > 3700 Then
        SK_Age = fvMin(fvMax(AgeVal, 0), 3700)
      Else
       SK_Age = fvMin(3700, fvMax(Cells(i, GrpAgetypeCol), 0))
      End If

    Else
      SK_Age = 0
    End If

    Cells(i, SKageCol) = SK_Age
  Next i

  foAp.Calculate
Next Iter

End Sub

Sub CheckForUPbWorkbook() ' 10/03/30 added
Dim GotStd As Boolean, GotSam As Boolean, CellA1$, w As Worksheet

GotStd = False: GotSam = False

For Each w In ActiveWorkbook.Worksheets
  CellA1 = w.Cells(1, 1).Formula
  If w.Name = pscStdShtNa And CellA1 = "Isotope Ratios of Standards" Then
    GotStd = True
  ElseIf w.Name = pscSamShtNa And CellA1 = "Isotope Ratios of Samples" Then
    GotSam = True
  End If
Next w

pbUPb = (GotStd And GotSam)
End Sub
