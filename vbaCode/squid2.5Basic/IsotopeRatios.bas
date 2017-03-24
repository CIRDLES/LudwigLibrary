Attribute VB_Name = "IsotopeRatios"
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

Sub GenIsoRunStartup()
' General Isotope-ratio data-reduction master Sub.  Respond to user request to
'  reduce GenIso data; show the GenIsoSetup form; reduce the data.

' 09/03/14 - Pass the new "NotForGrouping" parameter to GetNameList as TRUE to
'            so that case is ignored in creating the trimmed sample-name list
' 09/05/11 - Modify for rewritten Sub Reprocess, addding the ExistingCondensedSht,
'            NewReducedWbk, and DeleteReducedShts Boolean variables.
Dim GotIsoplot As Boolean, BadFile As Boolean, Exists As Boolean
Dim CanChart As Boolean, HasRejPk As Boolean, DeleteReducedShts As Boolean
Dim ExistingCondensedSht As Boolean, NewReducedWbk As Boolean
Dim RawDatFolder$, OrigDir$, OrigDrv$, When$, s$, tmp1$, tmp2$, EqnFractErr$
Dim Rcolname$(), FileLines$()
Dim Col%, tmpCol%, NumDenom%, NameDelim%, PkNum%, EqNum%, RatNum%
Dim RatCt%, BadSbm%(1), RatPkOrd%()
Dim r&, Rw&, FirstSpotRow&, LastRow&, InitialCalcSetting&
Dim Seconds#, FirstSec#, EqnRes#, EqnResult$, EqnFerr#
Dim SbmOffs#(), SbmOffsErr#(), SbmPk#() ', EqRes#(), EqFerr#()
Dim BlankedZeroesRange As Range
Dim vTmp As Variant
Dim RawDat() As RawData

pbUPb = False: foUser("sqGeochron") = False
piLwrIndx = 1

With foAp
  On Error Resume Next
  pbSqdBars = .CommandBars(pscSq).Visible
  On Error GoTo 0
  .DisplayCommentIndicator = xlCommentIndicatorOnly
End With
GetIsoplot GotIsoplot  ' Make sure Isoplot3.xla is loaded

If Not GotIsoplot Then
  ComplainCrash "Can't find ISOPLOT3.xla - must terminate."
End If

'NewReducedWbk = foUser("SlowReprocess")

ReProcess NewReducedWbk, ExistingCondensedSht, DeleteReducedShts

Do
  NoUpdate
  ShowStatusBar
  Alerts False

  If Not ExistingCondensedSht Then
    OrigDrv = fsCurDrive
    OrigDir = CurDir
    RawDatFolder = Trim(foUser("sqPDfolder"))
    If RawDatFolder <> "" Then ChDirDrv RawDatFolder

    Do
      ' Not re-using an open workbook's data - open a new file.
      GetFile RawDat, FileLines, BadFile
    Loop Until Not BadFile

    ChDirDrv OrigDir
  End If

  piNsChars = foUser("NsChars")
  If piNsChars = 0 Then piNsChars = fvMax(1, foUser("NsChars"))
  RefreshSampleNames
  pbFromSetup = False 'True
  GetNameList , , , True
  InitialCalcSetting = foAp.Calculation ' so can restore when done
  ManCalc
  psStS$ = "'" & pscSamShtNa & "'!"
  psProgName = fhSquidSht.[ProgName]
  Cbars pscSq, False  ' Hide the SQUID toolbar
  If Not fbIsFresh Then BuildTaskCatalog

  foAp.DisplayCommentIndicator = xlCommentIndicatorOnly
  piTrimCt = 0: piaSpotCt(0) = 0: piaSpotCt(1) = 0

  If Workbooks.Count > 0 And ExistingCondensedSht And NewReducedWbk Then
    phCondensedSht.Copy
    Set phCondensedSht = ActiveSheet
    Set pwDatBk = ActiveWorkbook
  End If

  GenIsoSetup.Show

  If piFormRes = peLoadNewFile Then ExistingCondensedSht = False

Loop Until piFormRes <> peCancel And piFormRes <> peLoadNewFile

If piFormRes = peUndefined Then CrashEnd , "in Sub GenIsoRunStartup.  piFormRes = peUndefined."

CheckForSolver

If Not ExistingCondensedSht Then
  CondenseRawData piFileNpks, pdaFileMass(), False, RawDat, FileLines
End If

Erase RawDat, FileLines

If pbRatioDat Then CreateRatioDatSheet

FirstSpotRow = plaSpotNameRowsCond(1)

With puTask
  .iNpeaks = piFileNpks
  FindStr "dead time", , tmpCol, 2 + FirstSpotRow, picDatCol, , , , , True
  pdDeadTimeSecs = Cells(4 + FirstSpotRow, tmpCol) / pdcBillion

  If pdDeadTimeSecs > 0 Then
    FindStr "sbm zero", , tmpCol, 3 + FirstSpotRow, picDatCol, , , , , True
    plSBMzero = Cells(picDatRowOffs + FirstSpotRow, tmpCol)
  End If

  If pbPDfile And pdDeadTimeSecs = 0 Then
    ParseLine Cells(FirstSpotRow + 1, 1), vTmp, 0, ","
    pdDeadTimeSecs = Val(vTmp(4)) / pdcBillion
    plSBMzero = Val(Mid$(vTmp(5), 1 + Len("sbm zero ")))
  End If

  ReDim pdaPkMass(1 To .iNpeaks), pdaTotCps(1 To .iNpeaks)
  ReDim piaCPScol(1 To .iNpeaks), psaCPScolHdr(1 To .iNpeaks)
End With

For PkNum = 1 To puTask.iNpeaks
    pdaPkMass(PkNum) = pdaFileMass(PkNum)
Next PkNum

piaNumSpots(0) = piNumAllSpots

With puTask
  ReDim RatPkOrd(1 To .iNrats, 1 To 2), pdaSbmDeltaPcnt(1 To .iNpeaks, 1 To piNumAllSpots)
  ReDim pdaPkMass(1 To .iNpeaks), RatPkOrd(1 To .iNrats, 1 To 2), pdaTrimMass(1 To .iNpeaks, 1 To 2000)
  ReDim pdaTrimTime(1 To .iNpeaks, 2000), piaCPScol(1 To .iNpeaks)
  ReDim psaCPScolHdr(1 To .iNpeaks), pdaTotCps(1 To .iNpeaks)

  For RatNum = 1 To puTask.iNrats

    For NumDenom = 1 To 2

      For PkNum = 1 To puTask.iNpeaks

        If .daNmDmIso(NumDenom, RatNum) = .daNominal(PkNum) Then
          RatPkOrd(RatNum, NumDenom) = PkNum
          Exit For
        End If

      Next PkNum

    Next NumDenom

  Next RatNum

  pbDone = False
  ColWidth picAuto, 1
  ShowStatusBar
  StatBar

  For PkNum = 1 To puTask.iNpeaks
    pdaPkMass(PkNum) = pdaFileMass(PkNum)
  Next PkNum

  EqnDetails

  If pbRatioDat Then
    CreateRatioDatSheet
    phCondensedSht.Activate
  End If

  plHdrRw = 6
  ShowStatusBar
  Cells(1, 1).Select
  piaSpotCt(0) = 0
  ReDim rr#(1 To 99), rrFerr#(1 To 99), pdaTrimMass(1 To 3 + .iNpeaks, 1 To 2000)
End With

ReDim pbSamRej(1 To piNumAllSpots, 1 To 99)
NoUpdate

pbDone = False: piNscans = 0:     piTrimCt = 0
piSpotNum = 0:  piaSpotCt(0) = 0

ColWidth picAuto, 1
ShowStatusBar
NoUpdate
StatBar
piSpotNum = piaStartSpotIndx(0) - 1
Sheets.Add
ActiveSheet.Name = fsLSN(pscSamShtNa)
Set phSamSht = ActiveSheet
NoGridlines
With Cells.Font: .Name = psStdFont: .Size = 11: End With
Zoom piIzoom

AssignGenIsoTaskColumns Rcolname()
CollateUserConstants

Columns(1).NumberFormat = "@"

Do
  phCondensedSht.Activate
  r = ActiveCell.Row
  ParseRawData 1 + piSpotNum, True, False, When, True

  If r = 0 Or pbDone Then
    If piSpotNum > 0 Then Exit Do
    ComplainCrash "Can't interpret this data sheet."
  End If

  piSpotNum = 1 + piSpotNum: piaSpotCt(0) = 1 + piaSpotCt(0)
  Cells(r + 1, 1).Activate
  StatBar psSpotName
  ParseTimedate When, Seconds
  If FirstSec = 0 Then FirstSec = Seconds
  phSamSht.Activate
  If FirstSec = 0 Then FirstSec = Seconds
  Rw = piaSpotCt(0) + plHdrRw
  CFs Rw, 1, psSpotName$
  CFs Rw, piDateTimeCol, When
  CF Rw, piHoursCol, foAp.Fixed((Seconds - FirstSec) / 3600#, 3)

  With puTask

    For PkNum = 1 To .iNpeaks

      If .baCPScol(PkNum) And piaCPScol(PkNum) > 0 Then
        CF Rw, piaCPScol(PkNum), pdaTotCps(PkNum)
      End If

    Next PkNum

  End With

  If pbXMLfile Then
    r = plaSpotNameRowsCond(piSpotNum) + picDatRowOffs
    Col = picDatCol + 5 * puTask.iNpeaks + 2
    With phCondensedSht
      CF Rw, piStageXcol, .Cells(r, Col)
      CF Rw, piStageYcol, .Cells(r, Col + 1)
      CF Rw, piStageZcol, .Cells(r, Col + 2)
      CF Rw, piQt1yCol, .Cells(r, Col + 3)
      CF Rw, piQt1Zcol, .Cells(r, Col + 4)
      CF Rw, piPrimaryBeamCol, .Cells(r, Col + 5)
    End With
  End If

  RatCt = 0
  phSamSht.Activate
  With puTask

    For RatNum = 1 To .iNrats
      RatCt = 1 + RatCt
      InterpRat RatPkOrd(RatNum, 1), RatPkOrd(RatNum, 2), rr(RatCt), _
        rrFerr(RatCt), BadSbm(), 0, HasRejPk
      pbSamRej(piaSpotCt(0), RatCt) = HasRejPk
      Rw = piaSpotCt(0) + plHdrRw
      CF Rw, piaIsoRatCol(RatCt), rr(RatCt)
      rrFerr(RatCt) = fvMin(9.99, rrFerr(RatCt))
      CF Rw, piaIsoRatEcol(RatCt), rrFerr(RatCt), True
    Next RatNum

    For EqNum = 1 To .iNeqns

      If Not .baSolverCall(EqNum) Then

        If fbOkEqn(EqNum) Then
          With .uaSwitches(EqNum)
            If .SC Or .FO Or piaEqnRats(EqNum) = 0 Then ' Put in sequence with existing data columns

              If Not .SC Or (.FO And piSpotNum = 1) Then
                Formulae puTask.saEqns(EqNum), EqNum, False, Rw, piaEqCol(0, EqNum), Rw
                On Error Resume Next
                If Not .FO Then EqnResult = Cells(Rw, piaEqCol(0, EqNum))
                On Error GoTo 0
              End If

            Else
              piSpotOutputCol = piaEqCol(0, EqNum)
              EqnInterp puTask.saEqns(EqNum), EqNum, EqnRes, EqnFerr, 1, 0
              EqnResult = fsS(EqnRes)
              EqnFractErr = fsS(100 * EqnFerr)
              CFs Rw, piaEqCol(0, EqNum), EqnResult  ' results of eqn(eqnum) in eqcol(0,eqnum)
              CFs Rw, piaEqEcol(0, EqNum), EqnFractErr
            End If

          End With
        End If

      End If

    Next EqNum

  End With
  plaFirstDatRw(0) = 1 + plHdrRw: plaLastDatRw(0) = plHdrRw + piSpotNum
Loop Until piSpotNum = piaEndSpotIndx(0)

TaskSolverCall
StatBar ""

piLastCol = fiEndCol(plHdrRw)
LastRow = flEndRow

With puTask

  For EqNum = 1 To .iNeqns
    With .uaSwitches(EqNum)

      If .SC And .Nu Then
        Formulae puTask.saEqns(EqNum), EqNum, False, 1 + plHdrRw, piaEqCol(0, EqNum)
      End If


    End With
  Next EqNum

End With

'ColWidth picAuto, piNameCol, piLastCol

' Add titles & labels, do some formatting
Cells(1, 1) = "Isotope Ratios"
Cells(2, 1) = "(errors are 1s unless otherwise specified)"
Cells(2, 1).Characters(14, 1).Font.Name = "Symbol"
With Cells(1, 1).Font
  .Bold = True: .Size = 1.2 * .Size
End With
Nformat piHoursCol
If piBkrdCtsCol > 0 Then Nformat piBkrdCtsCol, , , , True

For PkNum = 1 To puTask.iNpeaks

  If piaCPScol(PkNum) > 0 And Not fbIsSquidHid(1, piaCPScol(PkNum)) Then ' 09/07/21 -- MOD
    Nformat piaCPScol(PkNum) ' 09/07/21 -- added
  End If

Next PkNum

For Col = piFirstRatCol To piLastCol
  If Not fbIsSquidHid(1, Col) Then Nformat Col ' 09/07/21 -- added
  If Not fbIsSquidHid(1, Col + 1) Then         ' 09/07/21 -- added
    s = fsStrip(Cells(plHdrRw, Col + 1))
    If s = "%err" Then
      Nformat Col + 1, True
      Col = Col + 1
    End If
  End If
Next Col

With puTask
  For EqNum = 1 To .iNeqns

    If Not .baSolverCall(EqNum) Then
      With .uaSwitches(EqNum)

        If .ZC Then

          If .SC Then
            Set BlankedZeroesRange = Cells(plaFirstDatRw(0), piaEqCol(0, EqNum))
          ElseIf .Ar And (.Ar And .ArrNcols > 1 Or .ArrNrows > 1) Then
            s = puTask.saEqnNames(EqNum)
            NameDelim = InStr(s, "||")
            If NameDelim > 0 Then s = Left$(s, NameDelim - 1)
            FindStr s, r, Col, flHeaderRow(0), 1, plaLastDatRw(0)

            If r > 0 And Col > 0 Then

              If .ArrNrows = 1 And .ArrNcols > 1 Then
                Set BlankedZeroesRange = frSr(plaFirstDatRw(0), Col, plaLastDatRw(0), Col + .ArrNcols - 1)
              ElseIf .ArrNrows > 1 And .ArrNcols > 1 Then
                Set BlankedZeroesRange = frSr(r, Col, r + .ArrNrows - 1, Col + .ArrNcols - 1)
              ElseIf .ArrNrows > 1 And .ArrNcols = 1 Then

              End If
            Else
              Set BlankedZeroesRange = Nothing
            End If

          Else
            Set BlankedZeroesRange = frSr(plaFirstDatRw(0), piaEqCol(0, EqNum), plaLastDatRw(0))
          End If

          BlankZeroCells BlankedZeroesRange
        End If

      End With
    End If

  Next EqNum

End With

HA xlRight, plHdrRw, 1, LastRow, piLastCol
Columns(2).AutoFit ' 09/11/12 -- added


If pbXMLfile Then
  If piStageXcol > 0 Then HA xlCenter, plHdrRw, piStageXcol
  If piStageYcol > 0 Then HA xlCenter, plHdrRw, piStageYcol
  If piStageZcol > 0 Then HA xlCenter, plHdrRw, piStageZcol
  If piQt1yCol > 0 Then HA xlCenter, plHdrRw, piQt1yCol
  If piQt1Zcol > 0 Then HA xlCenter, plHdrRw, piQt1Zcol
End If

'ColWidth picAuto, 2, piLastCol ' 09/07/21 -- commented out

For Col = 3 To piLastCol

  If Not fbIsSquidHid(1, Col) Then ' 09/07/21 -- added
    If Columns(Col).ColumnWidth > 15 Then
      RangeNumFor "0", frSr(1 + plHdrRw, Col, LastRow)
      ColWidth 22, Col
      ColWidth picAuto, Col, , 9
      RangeNumFor "general", frSr(1 + plHdrRw, Col, LastRow)
    Else
      Columns(Col).AutoFit ' 09/07/21 -- added
    End If

  End If

Next Col

foAp.Calculate
ActiveSheet.DisplayAutomaticPageBreaks = False

With puTask

  For EqNum = 1 To .iNeqns
    ColWidth picAuto, piaEqCol(0, EqNum)
    ColWidth picAuto, piaEqEcol(0, EqNum)
  Next EqNum

  ChangeCols 1, plHdrRw
  On Error GoTo 0
  tmp1 = "Ratios are " & IIf(pbSbmNorm, "", "NOT ") & "normalized to SBM ("
  If pbSbmNorm And Not foUser("interpsbmnorm") Then tmp1 = tmp1 & "un-"
  tmp1 = tmp1 & "interpolated)"
  tmp2 = "Spot values for Task eqns calculated " & _
    IIf(pbLinfitEqns, "at mid spot-time", "as spot average")
  tmp2 = tmp2 & ", for isotope ratios of the same element " & _
    IIf(pbLinfitRats, "at mid spot-time", "as spot average")
  tmp2 = tmp2 & ", for isotope ratios of different elements " & _
    IIf(pbLinfitRatsDiff, "at mid spot-time", "as spot average")   ' 09/06/18 -- added
  Fonts 1, 7, , , vbBlue, 0, xlLeft, 12, , , , , , , , tmp1
  Fonts 2, 7, , , 160, 0, xlLeft, 12, , , , , , , , tmp2
  Fonts rw1:=3, Col1:=7, Bold:=True, Clr:=RGB(0, 128, 0), Formul:="Task: " & .sName
  With fhSquidSht
    Fonts 4, 7, , , RGB(0, 0, 128), , xlLeft, 12, , , "SQUID " & .[Version] & ", rev. " & .[revdate].Text
  End With
  With frSr(plaFirstDatRw(0), 1, plaLastDatRw(0), piLastCol)
    .FormatConditions.Delete
    .FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(A1)"
    .FormatConditions(1).Font.Color = vbWhite
  End With
  BlankZeroCells BlankedZeroesRange
  PlaceEqnBox

  If pbSbmNorm Then
    ReDim SbmOffs#(1 To .iNpeaks), SbmOffsErr#(1 To .iNpeaks), SbmPk#(1 To piaSpotCt(0))
    SBMdata SbmOffs, SbmPk, SbmOffsErr, pdaSbmDeltaPcnt, CanChart
    If CanChart Then AddSBMchart plHdrRw, SbmOffs(), SbmOffsErr()
  End If

End With

StatBar
Cells(4, 1) = " from file:"
Fonts 4, 7, , , , True, xlLeft
Cells(4, 2) = phCondensedSht.Name
Fonts 4, 7, , , RGB(150, 0, 0), True, xlLeft
Cells(3, 1) = phCondensedSht.Cells(2, 1)
Fonts 3, 1, , , , , xlLeft, 10
Cbars pscSq, pbSqdBars

With ActiveWindow
  .DisplayWorkbookTabs = True
  .DisplayHorizontalScrollBar = True
  .TabRatio = 0.5
End With

Cells(peMaxCol, 1).Select: Cells(1, 1).Select
Set phSamSht = ActiveSheet
Set phStdSht = ActiveSheet
pbStd = False
phSamSht.Activate
LoadStorePrefs 2

If puTask.iNumAutoCharts > 0 Then CreateAutoCharts

pbFromSetup = False
phSamSht.Activate

If pbDoMagGrafix And piRefPkOrder > 0 Then
  TrimMassStuff LastRow + 5
  phSamSht.Activate
End If

If foUser("AttachTask") Then
  AttachWorksheetInFileToOpenWorkbook puTask.sFileName, 0, , , , True
End If

If foUser("DatRedParamsSeparate") Then MakeDatRedParamsSht

ActiveSheet.DisplayAutomaticPageBreaks = False
StatBar
foAp.Calculation = InitialCalcSetting
End Sub

Sub AssignGenIsoTaskColumns(Rcolname$())
' Determine headers, order, and indexes of the output-data columns'
Dim b As Boolean
Dim pe$, s$, Seq$, Nom$, Nom2$
Dim c%, k%, h%, i%, j%, ArrHdrCt%, EqNum%, RatNum%, PkNum%
Dim Nomi#
Dim ArrHdrs As Variant

Cells.Delete
'sig1 = fsVertToLF("1sigma|err")
pe = fsVertToLF(pscPpe)
h = plHdrRw: c = 0: j = 0

With puTask
  ReDim piaIsoRatCol(1 To .iNrats), piaIsoRatEcol(1 To .iNrats)

  If .iNeqns > 0 Then
    ReDim piaEqCol(-1 To 0, 1 To .iNeqns), piaEqEcol(-1 To 0, 1 To .iNeqns)
  End If

  ReDim Rcolname(1 To .iNrats)

  ColInc c, "SpotName", "Name", 1, piNameCol
  CFs h, piNameCol, "Spot Name"
  ColInc c, "Date/Time", "DateTime", 1, piDateTimeCol
  CFs h, piDateTimeCol, "Date/Time"
  ColInc c, "Hours", "Hours", 1, piHoursCol
  CFs h, piHoursCol, "Hours", , "Elapsed since first spot of this session"

  If pbXMLfile Then
    ColInc c, "stage|X", "StageX", 1, piStageXcol
    CFs h, piStageXcol, "stage|X", -1
    ColInc c, "stage|Y", "StageY", 1, piStageYcol
    CFs h, piStageYcol, "stage|Y", -1
    ColInc c, "stage|Z", "StageZ", 1, piStageZcol
    CFs h, piStageZcol, "stage|Z", -1
    ColInc c, "Qt1y", "Qt1y", 1, piQt1yCol
    CFs h, piQt1yCol, "Qt1y"
    ColInc c, "Qt1z", "Qt1z", 1, piQt1Zcol
    CFs h, piQt1Zcol, "Qt1z"
    ColInc c, "Primary|beam|(na)", "PrimaryBeam", 1, piPrimaryBeamCol
    CFs h, piPrimaryBeamCol, "Primary|beam|(na)", -1
  End If

  For PkNum = 1 To .iNpeaks

    If .baCPScol(PkNum) Then  ' 09/11/11 -- removed "And PkNum <> piBkrdPkOrder"
      Nomi = .daNominal(PkNum): Nom = fsS(Nomi): Nom2 = Nom

      If PkNum = piRefPkOrder Then
        Nom2 = "Ref"
      End If

      ColInc c, "total|" & Nom & "|cts|/sec", Nom2 & "cts", 1, piaCPScol(PkNum)
      CFs h, piaCPScol(PkNum), "total|" & Nom & "|cts|/sec", True
      'If PkNum = piRefPkOrder Then piGenRefCol = c
      piFirstRatCol = piaCPScol(PkNum) + 1
    End If

  Next PkNum

  piFirstRatCol = 1 + c
  HA xlRight, h, 1, , piHoursCol
  k = 0

  For RatNum = 1 To .iNrats
    s = fsRatioHdrStr(.daNmDmIso(1, RatNum), .daNmDmIso(2, RatNum))
    Rcolname(RatNum) = s
    ColInc c, Rcolname(RatNum), Rcolname(RatNum), 2, piaIsoRatCol(RatNum), , , piaIsoRatEcol(RatNum), pe
    CFs h, piaIsoRatCol(RatNum), Rcolname(RatNum), -1:  CFs h, piaIsoRatEcol(RatNum), pe
    RangeNumFor pscGen, , piaIsoRatCol(RatNum)
    RangeNumFor fsQq(pscZd2), , piaIsoRatEcol(RatNum)
  Next RatNum

  For RatNum = piaIsoRatCol(1) To piaIsoRatCol(.iNrats) Step 2
    HA xlRight, plHdrRw, RatNum
    HA xlCenter, plHdrRw, RatNum + 1
  Next RatNum

  For k = 1 To 3

    For EqNum = 1 To .iNeqns
      If InStr(LCase(.saEqnNames(EqNum)), "<<solve>>") = 0 Then
        Seq = .saEqnNames(EqNum)
        With .uaSwitches(EqNum)
          Select Case k
            Case 1: b = Not .SC And Not .Ar
            Case 2: b = .SC And Not .Ar
            Case 3: b = .Ar
          End Select

          If b Then
            If k < 3 Then

              ColInc c, Seq, Seq, 1 - (.Nu And piaEqnRats(EqNum) > 0), _
                  piaEqCol(0, EqNum), , , piaEqEcol(0, EqNum), pscPpe
              CFs h, piaEqCol(0, EqNum), puTask.saEqnNames(EqNum), -1
              CFs h, piaEqEcol(0, EqNum), pe
            Else
              ParseLine Seq, ArrHdrs, ArrHdrCt, "||"

              For j = 1 To fvMin(ArrHdrCt, .ArrNcols)

                If j = 1 Then
                  ColInc c, ArrHdrs(j), ArrHdrs(j), piaEqCol(0, EqNum), piaEqCol(0, EqNum)
                  CFs h, piaEqCol(0, EqNum), ArrHdrs(j), True
                Else
                  ColInc c, ArrHdrs(j), "", 1, piaEqCol(0, EqNum) + j - 1
                  CFs h, c, ArrHdrs(j), True
                End If

              Next j

            End If  ' k<3

          End If    ' tb

        End With
      End If

    Next EqNum

  Next k

End With

piLastCol = c
With frSr(plHdrRw, 1, , piLastCol)
  .Borders(xlEdgeBottom).LineStyle = xlDouble
  .Font.Bold = True
End With
End Sub

Sub EqnInterp(ByVal Formula$, ByVal EqNum%, MeanEq#, EqValFerr#, _
  SqrMinMSWD#, piWLrej%, Optional NoRatSht As Boolean = False)

' Calculate means of the scan-by-scan Task Equation results, using double-interpolation.

Const Numer = 1, Denom = 2
Dim IsPkOrder As Boolean, Bad As Boolean, Singlescan As Boolean
Dim i%, NumRatPks%, InterpEqCt%, NumPksInclDupes%, NumRej%, sIndx%
Dim ScanNum%, RatNum%, TotalNumRatios%, Indx%, Num1Denom2%, PkOrder%
Dim InterpTime#, InterpTimeSpan#, PkFractErr#, ScanPkCts#, TotRatTime#
Dim RedPk1Ht#, RedPk2Ht#, NetPkCps#, EqValTmp#, EqFerr#, UndupedPk%, MSWD#
Dim PkF1#, PkF2#, PkTdelt#, MidTime#, Intercept#, SigmaIntercept#
Dim CovSlopeInter#, MeanTime#, MeanEqSig#, Slope#, SigmaSlope#, Probfit#
Dim FractInterpTime#, FractLessInterpTime#
Dim ReducedPkHt#(), ReducedPkHtFerr#(), SigRho#(), EqVal#()
Dim EqTime#(), AbsErr#(), FractErr#(), PkInterp#(), PkInterpFerr#()

TotalNumRatios = piaEqnRats(EqNum)
If TotalNumRatios = 0 Or piNscans < 1 Then
  MeanEq = pdcErrVal: Exit Sub
End If

NumRatPks = 2 * piaEqnRats(EqNum)

Indx = piNscans - 1
Singlescan = (piNscans = 1)
sIndx = Indx - Singlescan

With puTask
  ReDim PkInterpFerr(1 To .iNpeaks, 1 To sIndx), PkInterp#(1 To .iNpeaks, 1 To sIndx)
  ReDim PkFSig#(1 To .iNpeaks), ReducedPkHt(1 To .iNpeaks, 1 To piNscans)
  ReDim ReducedPkHtFerr(1 To .iNpeaks, 1 To piNscans)
End With

For ScanNum = 1 To piNscans

  For RatNum = 1 To TotalNumRatios   ' Convert SBM pkhts, calc errors & assign working

    For Num1Denom2 = Numer To Denom  '  pkht ("ReducedPkHt") to both sbm norm & un-norm.

      PkOrder = piaEqPkOrd(EqNum, RatNum, Num1Denom2)

      If PkOrder > 0 Then
        NetPkCps = pdaPkNetCps(PkOrder, ScanNum)
        If NetPkCps = pdcErrVal Then
          ReducedPkHt(PkOrder, ScanNum) = pdcErrVal
          Exit For
        End If

        ScanPkCts = NetPkCps * pdaIntT(PkOrder) ' counts for this scan

        If ScanPkCts <= 0 And ScanPkCts > 16 Then
          ReducedPkHt(PkOrder, ScanNum) = pdcErrVal
          Exit For
        End If

        If pdaSBMcps(PkOrder, ScanNum) <= 0 Then pbSbmNorm = False

        If pbSbmNorm Then
          ' Normalize to SBM counts for the on-peak time
          ScanPkCts = ScanPkCts / pdaSBMcps(PkOrder, ScanNum)
        End If

        PkFractErr = pdaPkFerr(PkOrder, ScanNum)

        If pbSbmNorm Then
          PkFractErr = sqR(PkFractErr ^ 2 + 1 / pdaSBMcps(PkOrder, ScanNum) _
                           / pdaIntT(PkOrder))
        End If

        ReducedPkHt(PkOrder, ScanNum) = ScanPkCts / pdaIntT(PkOrder)
        ReducedPkHtFerr(PkOrder, ScanNum) = PkFractErr
      End If

    Next Num1Denom2

  Next RatNum

Next ScanNum

For ScanNum = 1 To piNscans - 1 - Singlescan

  If Not Singlescan Then

    InterpTimeSpan = 0

    For RatNum = 1 To TotalNumRatios   ' Calc. mean time of/for interpolation

      For Num1Denom2 = Numer To Denom
        PkOrder = piaEqPkOrd(EqNum, RatNum, Num1Denom2)
        If PkOrder > 0 Then _
          InterpTimeSpan = InterpTimeSpan + pdaPkT(PkOrder, ScanNum) _
                         + pdaPkT(PkOrder, 1 + ScanNum)
      Next Num1Denom2

    Next RatNum

    InterpTime = InterpTimeSpan / NumRatPks / 2 ' time of interpolation
  End If

  ' Calculate the interpolated pk-heights
  For RatNum = 1 To TotalNumRatios

    For Num1Denom2 = Numer To Denom
      PkOrder = piaEqPkOrd(EqNum, RatNum, Num1Denom2)

      If PkOrder > 0 Then

        If Not Singlescan Then
          PkInterp(PkOrder, ScanNum) = pdcErrVal

          PkTdelt = pdaPkT(PkOrder, 1 + ScanNum) - pdaPkT(PkOrder, ScanNum)
          If PkTdelt <= 0 Then GoTo 1

          FractInterpTime = (InterpTime - pdaPkT(PkOrder, ScanNum)) / PkTdelt
          FractLessInterpTime = 1 - FractInterpTime
          RedPk2Ht = ReducedPkHt(PkOrder, 1 + ScanNum)
        End If

        RedPk1Ht = ReducedPkHt(PkOrder, ScanNum)
        If RedPk1Ht = pdcErrVal Or RedPk2Ht = pdcErrVal Then GoTo 1
        PkF1 = ReducedPkHtFerr(PkOrder, ScanNum)

        If Singlescan Then
           PkInterp(PkOrder, ScanNum) = RedPk1Ht
           PkInterpFerr(PkOrder, ScanNum) = PkF1
        Else
          PkInterp(PkOrder, ScanNum) = FractLessInterpTime * RedPk1Ht + _
                                       FractInterpTime * RedPk2Ht
          PkF2 = ReducedPkHtFerr(PkOrder, 1 + ScanNum)
          PkInterpFerr(PkOrder, ScanNum) = _
            sqR((FractLessInterpTime * PkF1) ^ 2 + _
                (FractInterpTime * PkF2) ^ 2)
        End If

      End If

    Next Num1Denom2
  Next RatNum
1:
Next ScanNum

InterpEqCt = 0

For ScanNum = 1 To piNscans - 1 - Singlescan
  '   sbm-norm ctg errs should be ok
  'If EqNum = -1 And Not pbStd Then i = 1 / 0
  FormulaEval Formula, EqNum, ScanNum, PkInterp(), PkInterpFerr(), EqValTmp, EqFerr

  If EqFerr > 0 And EqValTmp <> pdcErrVal Then
    InterpEqCt = InterpEqCt + 1
    ReDim Preserve EqVal(1 To InterpEqCt), FractErr(1 To InterpEqCt), AbsErr(1 To InterpEqCt)
    ReDim Preserve EqTime(1 To InterpEqCt)
    EqVal(InterpEqCt) = EqValTmp

    ' Kluge to approximate effect of interpolation combining peaks
    ' 10/04/02 commented out If piNscans > 2 And Not Singlescan Then EqFerr = EqFerr * 1.2

    AbsErr(InterpEqCt) = Abs(EqFerr * EqValTmp)
    FractErr(InterpEqCt) = EqFerr

    If AbsErr(InterpEqCt) = 0 Then
      AbsErr(InterpEqCt) = pdcTiny
      FractErr(InterpEqCt) = pdcTiny
    End If

    NumPksInclDupes = 0: TotRatTime = 0

    For UndupedPk = 1 To piNoDupePkN(EqNum)
      PkOrder = piaEqPkUndupeOrd(EqNum, UndupedPk)
      IsPkOrder = False

      For RatNum = 1 To TotalNumRatios

        For Num1Denom2 = Numer To Denom

          If piaEqPkOrd(EqNum, RatNum, Num1Denom2) = PkOrder Then
            TotRatTime = TotRatTime + pdaPkT(PkOrder, ScanNum)
            NumPksInclDupes = NumPksInclDupes + 1

            If Not Singlescan Then
              TotRatTime = TotRatTime + pdaPkT(PkOrder, ScanNum + 1)
              NumPksInclDupes = NumPksInclDupes + 1
            End If

            IsPkOrder = True
            Exit For
          End If

        Next Num1Denom2

        If IsPkOrder Then Exit For
      Next RatNum

    Next UndupedPk
    EqTime(InterpEqCt) = TotRatTime / NumPksInclDupes

  End If
3:
Next ScanNum
'If EqNum = -1 And piSpotNum > 0 Then
'  If psaSpotNames(piSpotNum) = "2B-6.1" Then Stop
'End If
If InterpEqCt > 0 Then
  ReDim Preserve EqTime(InterpEqCt), EqVal(1 To InterpEqCt), AbsErr(1 To InterpEqCt)
  ReDim SigRho(1 To InterpEqCt, 1 To InterpEqCt)

  For i = 1 To InterpEqCt
    SigRho(i, i) = AbsErr(i)

    If i > 1 Then
      SigRho(i, i - 1) = 0.25 ' Aprroximate kluge
      SigRho(i - 1, i) = 0.25
    End If

  Next i

  If InterpEqCt = 1 Then
    MeanEq = EqVal(1): EqValFerr = FractErr(1)

  ElseIf InterpEqCt > 3 And (pbLinfitEqns Or _
        (pbLinfitSpecial And EqNum < 0)) Then
    WtdLinCorr 2, InterpEqCt, EqVal, SigRho, MSWD, Probfit, NumRej, Intercept, _
           SigmaIntercept, Bad, Slope, SigmaSlope, CovSlopeInter, EqTime

    If Bad Then
       MeanEq = pdcErrVal: EqValFerr = pdcErrVal: Exit Sub
    End If

    MidTime = (pdaPkT(puTask.iNpeaks, piNscans) + pdaPkT(1, 1)) / 2
    MeanEq = Slope * MidTime + Intercept
    MeanEqSig = sqR((MidTime * SigmaSlope) ^ 2 + _
                SigmaIntercept ^ 2 + 2 * MidTime * CovSlopeInter)

    'ApChangePerMin = Slope / MeanEq * 60    ' %drift/minute of EqVal
    'APCPMerr = SigmaSlope / MeanEq * 60     ' error in above

  Else
    WtdLinCorr 1, InterpEqCt, EqVal, SigRho, MSWD, Probfit, NumRej, MeanEq, _
               MeanEqSig, Bad

    If Bad Then
      MeanEq = pdcErrVal: EqValFerr = pdcErrVal: Exit Sub
    End If

  End If

  If MeanEq = 0 Then
    EqValFerr = 1
  Else
    EqValFerr = Abs(MeanEqSig / MeanEq)
  End If

End If

If pbRatioDat And Not NoRatSht And InterpEqCt > 0 Then ' 09/11/11 -- added "And InterpEqCt > 0"
  PlaceRats psSpotName, piSpotNum, 2, EqNum, EqTime, EqVal, FractErr
End If
End Sub

Sub WtdLin(X#(), Y#(), SigmaY#(), ByVal N%, Slope As Double, _
  Slope_Sig#, Y_bar#, Ybar_Sig#, X_bar#, piWLrej%, Optional Inter_, _
  Optional SigmaIntercept, Optional CovSI)
' Linear regression (or weighted average if N<4) of scan-by-scan
'    isotope ratios or Task Equations, weighted only by their y errors.
' Find y-error at average X; Test effect on MSWD of rejecting each point
'  sequentially; if drops by specified factor, reject.
'  ASSUMES UNCORRELATED Y-ERRORS

Dim i%, j%, k%, m%, Pass%, MaxRej%, MinInd%
Dim MinProb#, f#, MswdRatToler#, MaxProb#, MswdRat#, MinMSWD#
Dim x2#(), y2#(), SigmaY_2#(), SlopeSig#()

ReDim Rej(1 To N) As Boolean
ReDim x2(1 To N - 1), y2(1 To N - 1), SigmaY_2(1 To N - 1)
ReDim mw#(0 To N), Prob#(0 To N), Xbar#(0 To N), Ybar#(0 To N), YbarSig#(0 To N)
ReDim Inter#(0 To N), Slp#(0 To N), Prb#(0 To N), x1#(1 To N), y1#(1 To N), SigmaY_1#(1 To N)
ReDim InterSig#(0 To N), CovInterSlp#(0 To N), SlopeSig(0 To N), InterSig(0 To N)

If N > 7 Then
  MswdRatToler = 0.3
Else
  MswdRatToler = Choose(N - 1, 0.01, 0.1, 0.15, 0.2, 0.2, 0.25)
End If
MinProb = 0.1: piWLrej = 0
Pass = 0:      MaxRej = 2 + (N < 9)

For i = 1 To N
  x1(i) = X(i): y1(i) = Y(i)
  SigmaY_1(i) = SigmaY(i)
Next i

f = fvMax(foAp.Median(SigmaY_1), 0.0000000001)

For i = 1 To N
  SigmaY_1(i) = fvMax(SigmaY_1(i), f)
Next i

If N = 1 Then
  Y_bar = y1(1): Ybar_Sig = SigmaY_1(1)
  Exit Sub
ElseIf N < 4 Then
  SimpleWtdAv N, y1(), SigmaY_1(), Y_bar, Ybar_Sig
Else

  Do
    WeightedLinear x1(), y1(), N, SigmaY(), Slp(0), Inter(0), mw(0), Prb(0), _
      SlopeSig(0), InterSig(0), CovInterSlp(0), , Xbar(0), Ybar(0), YbarSig(0)
    MaxProb = Prb(0): MinInd = 0:  MinMSWD = mw(0)
  If MaxProb > 0.1 Then Exit Do

    For i = 1 To N ' Make array missing 1 point
      j = 0

      For k = 1 To N

        If k <> i And Not Rej(k) Then
          j = 1 + j
          x2(j) = x1(k): y2(j) = y1(k)
          SigmaY_2(j) = SigmaY_1(k)
        End If

      Next k

      m = j
      WeightedLinear x2(), y2(), m, SigmaY_2(), Slp(i), Inter(i), mw(i), Prb(i), _
      SlopeSig(i), InterSig(i), CovInterSlp(i), , Xbar(i), Ybar(i), YbarSig(i)
    Next i

    For i = 1 To N
      MswdRat = mw(i) / fvMax(pdcTiny, mw(0))

      If MswdRat < MswdRatToler And mw(i) < MinMSWD And Prb(i) > MinProb Then
        Rej(i) = True: MinInd = i: MaxProb = Prb(i): MinMSWD = mw(i)
      End If

    Next i

    If MinInd > 0 Then Pass = 1 + Pass: piWLrej = 1 + piWLrej
  If MinInd = 0 Or Pass = MaxRej Or MaxProb > 0.1 Then Exit Do
    j = 0

    For i = 1 To N
      If Not Rej(i) Then
        j = 1 + j
        x2(j) = x1(i): y2(j) = y1(i):  SigmaY_2(j) = SigmaY_1(i)
      End If
    Next i

    N = N - 1
    ReDim x1(1 To N), y1(1 To N), SigmaY_1(1 To N), Prob(0 To N), Xbar(0 To N)

    For i = 1 To N
      x1(i) = x2(i): y1(i) = y2(i): SigmaY_1(i) = SigmaY_2(i)
    Next i

  Loop

  f = sqR(MinMSWD)
  Y_bar = Ybar(MinInd)
  X_bar = Xbar(MinInd)
  Ybar_Sig = YbarSig(MinInd)
  Slope = Slp(MinInd)
  Slope_Sig = SlopeSig(MinInd)
  Inter_ = Inter(MinInd)
  SigmaIntercept = InterSig(MinInd)

  If MaxProb < 0.1 Then
    Ybar_Sig = Ybar_Sig * f
    SigmaIntercept = SigmaIntercept * f
    Slope_Sig = Slope_Sig * f
  End If

  CovSI = CovInterSlp(MinInd)
End If
End Sub

Sub WtdLinCorr(Avg1LinRegr2%, ByVal N%, Y#(), SigRho#(), MSWD#, Probfit#, piWLrej%, _
    Intercept#, SigmaIntercept#, Bad As Boolean, Optional Slope#, _
    Optional SigmaSlope#, Optional CovSlopeInter#, Optional X)

' Same as Sub WtdLin, but does not assume that  y(i)-y(i+1) errors are uncorrelated.

Dim LinReg As Boolean
Dim i%, j%, Nw%, Pass%, MaxRej%, MinIndex%, TooSmall%
Dim MinProb#, f#, MswdRatToler#, MaxProb#, MswdRat#, MinMSWD#
Dim x2#(), y2#(), SlopeSigmaW#(), SigRho2#(), SigRho1#()
Dim SlopeW#(), Prob#(), InterW#(), ProbW#(), x1#(), y1#(), InterSigmaW#()
Dim CovSlopeInterW#(), SigmaY#(), MswdW#()

LinReg = (Avg1LinRegr2 = 2)
If LinReg Then
  ReDim x2(1 To N), SlopeW(0 To N), CovSlopeInterW(0 To N)
  ReDim SlopeSigmaW(0 To N), x1(1 To N)
End If

ReDim y2(1 To N), SigRho1(1 To N, 1 To N), SigRho2(1 To N, 1 To N), SigmaY(1 To N), y1(1 To N)
ReDim Rej(1 To N) As Boolean, SigRho1(1 To N, 1 To N), SigRho2(1 To N, 1 To N)
ReDim MswdW(0 To N), Prob(0 To N), ProbW(0 To N), InterW(0 To N), InterSigmaW(0 To N)

Bad = False
TooSmall = Choose(Avg1LinRegr2, 2, 4)

If N < TooSmall Then
  MsgBox "N=" & StR(N) & " passed to WtdLinCorr with Avg1LinRegr2=" _
         & fsS(Avg1LinRegr2), , pscSq: End
ElseIf N > 7 Then
  MswdRatToler = 0.3
Else
  MswdRatToler = Choose(N - Avg1LinRegr2, 0, 0.1, 0.15, 0.2, 0.2, 0.25)
End If

MaxRej = 1 + (N - Avg1LinRegr2) \ 8
MinProb = 0.1
piWLrej = 0
Pass = 0

For i = 1 To N

  If LinReg Then
    x1(i) = X(i): x2(i) = X(i)
  End If

  y1(i) = Y(i): y2(i) = Y(i)
  SigmaY(i) = SigRho(i, i)

  For j = 1 To N
    SigRho1(i, j) = SigRho(i, j)
    SigRho2(i, j) = SigRho(i, j)
  Next j

Next i

f = fvMax(foAp.Median(SigmaY), 0.0000000001)

For i = 1 To N
  SigRho1(i, i) = fvMax(SigRho1(i, i), f)
  SigRho2(i, i) = SigRho1(i, i)
Next i

MinIndex = -1

Do

  For i = 0 To N

    If i > 0 Then
      DeletePoint LinReg, N, y1, y2, SigRho1, SigRho2, i, x1, x2
      Nw = N - 1
    Else
      Nw = N
    End If

    If Nw = 1 And False Then
      ProbW(i) = 1
      MswdW(i) = 0
      InterSigmaW(i) = 1
      InterW(i) = 1
    ElseIf LinReg Then
      WeightedLinearCorr x2, y2, Nw, SigRho2, SlopeW(i), InterW(i), MswdW(i), ProbW(i), _
        SlopeSigmaW(i), InterSigmaW(i), CovSlopeInterW(i), Bad
    Else
      WtdAvCorr y2, SigRho2, Nw, InterW(i), InterSigmaW(i), MswdW(i), ProbW(i), True, Bad
    End If

    If i = 0 Then
      If ProbW(0) > 0.1 Then
        MinIndex = 0: MinMSWD = MswdW(0)
        Exit For
      End If
      MaxProb = Prob(0)
    End If

  Next i

If MinIndex = 0 Then Exit Do
  MinIndex = 0:  MinMSWD = MswdW(0)

  For i = 1 To N
    MswdRat = MswdW(i) / fvMax(pdcTiny, MswdW(0))

    If MswdRat < MswdRatToler And MswdW(i) < MinMSWD And ProbW(i) > MinProb Then
      Rej(i) = True
      piWLrej = 1 + piWLrej
      MinIndex = i
      MaxProb = ProbW(i)
      MinMSWD = MswdW(i)
    End If

  Next i

  Pass = 1 + Pass

If Pass > 0 And (MinIndex = 0 Or Pass = MaxRej Or MaxProb > 0.1) Then Exit Do

  DeletePoint LinReg, N, y1, y2, SigRho1, SigRho2, MinIndex, x1, x2

  N = N - 1
  ReDim x1(1 To N), y1(1 To N), Prob(0 To N), Xbar(0 To N), SigRho1(1 To N, 1 To N)

  For i = 1 To N
    If LinReg Then x1(i) = x2(i)
    y1(i) = y2(i)
    For j = 1 To N - 1
      SigRho1(i, j) = SigRho2(i, j)
  Next j, i

Loop

If LinReg And MinIndex > 0 Then                ' 10/10/05 -- added "And MinIndex > 0"
  Slope = SlopeW(MinIndex)
  SigmaSlope = SlopeSigmaW(MinIndex)
  CovSlopeInter = CovSlopeInterW(MinIndex)
End If

Intercept = InterW(MinIndex)
SigmaIntercept = InterSigmaW(MinIndex)
MSWD = MswdW(MinIndex)
Probfit = ProbW(MinIndex)

If Probfit < 0.05 Then
  f = sqR(MSWD)
  SigmaIntercept = SigmaIntercept * f
  If LinReg Then SigmaSlope = SigmaSlope * f
End If

End Sub

Sub DeletePoint(LinReg As Boolean, ByVal N%, y1#(), y2#(), SigRho1#(), _
                SigRho2#(), RejPoint%, Optional x1, Optional x2)
' Delete a data point from a vector of data-points, errors, & error correls.
Dim j%, m%, p%, Nn%

Nn = N - 1
If LinReg Then ReDim x2(1 To Nn)
ReDim SigRho2(1 To Nn, 1 To Nn)

For j = 1 To N
  m = j + 1:  p = j + 2

  If j < RejPoint Then
    SigRho2(j, j) = SigRho1(j, j)
    y2(j) = y1(j)
    If LinReg Then x2(j) = x1(j)
  ElseIf j < N Then
    SigRho2(j, j) = SigRho1(m, m)
    y2(j) = y1(m)
    If LinReg Then x2(j) = x1(m)
  End If

  If j < (RejPoint - 1) Then
    SigRho2(j, m) = SigRho1(j, m)
    SigRho2(m, j) = SigRho1(m, j)
  ElseIf j = (RejPoint - 1) And m < N Then
    SigRho2(j, m) = 0
    SigRho2(m, j) = 0
  ElseIf j < (N - 1) Then
    SigRho2(j, m) = SigRho1(m, p)
    SigRho2(m, j) = SigRho1(p, m)
  End If

Next j

End Sub

Sub WeightedLinear(X#(), Y#(), N%, Ysig#(), Slope#, Inter#, MSWD#, Prob#, _
  Optional SlopeSig# = 0, Optional InterSig# = 0, Optional SlopeInterCov# = 0, _
  Optional RhoInterCov = 0, Optional Xbar# = 0, Optional Ybar# = 0, Optional YbarSig# = 0)
' Y-error weighted linear regression, no y(i)-y(j) error correlations.
Dim i%
Dim Sx#, Sy#, Sxy#, Sx2#, sw#, Sums#, Ypred#, Yresid#, Denom#
Dim Fischer(1 To 2, 1 To 2) As Variant, FischerInv As Variant

ReDim w(1 To N)
sw = 0: Sx = 0: Sy = 0: Sxy = 0: Sx2 = 0: Sums = 0: Ybar = 0: Xbar = 0

For i = 1 To N
  w(i) = 1 / fvMax(pdcTiny, Ysig(i)) ^ 2
  Sx = Sx + w(i) * X(i)
  Sx2 = Sx2 + w(i) * X(i) ^ 2
  Sxy = Sxy + w(i) * X(i) * Y(i)
  Sy = Sy + w(i) * Y(i)
  sw = sw + w(i)
Next i

Denom = sw * Sx2 - Sx ^ 2
Sx2 = fvMax(pdcTiny, Sx2)
If Denom = 0 Then Denom = pdcTiny
Inter = (Sx2 * Sy - Sx * Sxy) / Denom
Slope = (Sxy - Inter * Sx) / Sx2
Fischer(1, 1) = Sx2
Fischer(2, 2) = sw
Fischer(2, 1) = Sx
Fischer(1, 2) = Sx
On Error Resume Next
FischerInv = Application.MInverse(Fischer)
SlopeSig = fvMax(pdcTiny, sqR(FischerInv(1, 1)))
InterSig = fvMax(pdcTiny, sqR(FischerInv(2, 2)))
SlopeInterCov = FischerInv(1, 2)
RhoInterCov = SlopeInterCov / SlopeSig / InterSig

For i = 1 To N
  Ypred = Slope * X(i) + Inter
  Yresid = Y(i) - Ypred
  Sums = Sums + Yresid ^ 2 * w(i)
Next i

MSWD = Sums / (N - 2)
Prob = ChiSquare(MSWD, N - 2)
Xbar = Sx / sw
Ybar = Slope * Xbar + Inter
YbarSig = sqR(1 / sw)
End Sub

Sub WeightedLinearCorr(X#(), Y#(), N%, SigmaRhoY#(), Slope#, Inter#, MSWD#, Prob#, _
  SlopeSig#, InterSig#, SlopeInterCov#, Bad As Boolean)
' Y-error weighted linear regression with nonzero y(i)-y(i+1) error correlations.
Dim i%, j%
Dim SumSqWtdResids#, Mx#, Px#, Py#, Pxy#, w#, InvOm# ', SlopeInterRho#
Dim Omega#(), Fischer#(1 To 2, 1 To 2), Resid#()
Dim FischerInv As Variant, InvOmega As Variant

ReDim Omega(1 To N, 1 To N), Resid(1 To N, 1 To 1)

For i = 1 To N

  For j = 1 To N

    If i = j Then
      Omega(i, i) = SigmaRhoY(i, j) ^ 2
    ElseIf Abs(i - j) > 1 Then
      Omega(i, j) = 0
    ElseIf i < N Then
      Omega(i, j) = SigmaRhoY(i, j) * SigmaRhoY(i, i) * SigmaRhoY(j, j)
      Omega(j, i) = Omega(i, j)
    End If

  Next j

Next i
InvOmega = foAp.MInverse(Omega)

Mx = 0: Px = 0: Py = 0: Pxy = 0: w = 0

For i = 1 To N

  For j = 1 To N
    InvOm = InvOmega(i, j)
    w = w + InvOm
    Px = Px + (X(i) + X(j)) * InvOm
    Py = Py + (Y(i) + Y(j)) * InvOm
    Pxy = Pxy + (X(i) * Y(j) + X(j) * Y(i)) * InvOm
    Mx = Mx + X(i) * X(j) * InvOm
  Next j

Next i

Slope = (2 * Pxy * w - Px * Py) / (4 * Mx * w - Px ^ 2)
Inter = (Py - Slope * Px) / (2 * w)

Bad = True
On Error GoTo BadMat

Fischer(1, 1) = Mx:     Fischer(2, 2) = w
Fischer(1, 2) = Px / 2: Fischer(2, 1) = Px / 2
FischerInv = foAp.MInverse(Fischer)
SlopeSig = sqR(FischerInv(1, 1))
InterSig = sqR(FischerInv(2, 2))
SlopeInterCov = FischerInv(1, 2)
'SlopeInterRho = SlopeInterCov / SlopeSig / InterSig

For i = 1 To N
  Resid(i, 1) = Y(i) - Slope * X(i) - Inter
Next i

SumSqWtdResids = SumSquares(N, Resid, InvOmega)
MSWD = SumSqWtdResids / (N - 2)
Prob = ChiSquare(MSWD, N - 2)
Bad = False

BadMat: On Error GoTo 0
End Sub

Sub EqnDetails() ' Get information about the current Task Equations
' Including whether Excel's Solver is required, if column-swapping is
'  required and details thereof, what the numerator & denominator peaks are
'  of the isotope ratios used by the Equations, disregarding duplicated isotope-
'  ratio peaks.
Dim tB As Boolean, CanSolve As Boolean, bSolver() As Boolean
Dim UnswapEqns$(), Swappo$()
Dim j%, m%, p%, q%, RatCt%, EqIndx%, NumDenom%, EqAsc%(1 To peMaxRats)
Dim NomMass#

With puTask
  If Not pbUPb And .iNeqns = 0 Then Exit Sub
  p = piLwrIndx
  ReDim piaEqnRats(p To .iNeqns), piaNeqnTerms(p To .iNeqns)
  ReDim UnswapEqns(p To .iNeqns), Swappo(p To .iNeqns)
  ReDim piaEqPkOrd%(p To .iNeqns, 40, 1 To 2), piaEqPkUndupeOrd(p To .iNeqns, 1 To 30)
  ReDim piaBrakType(1 To 99, p To .iNeqns), piNoDupePkN(p To .iNeqns)

  If .iNeqns > 0 Then
    ReDim EqnDest(1 To .iNeqns), RatSource(1 To .iNeqns, 1 To peMaxRats), bSolver(p To .iNeqns)
  End If


  For EqIndx = piLwrIndx To .iNeqns   ' Look for redirect indicator ("<=>")

    If EqIndx <> 0 Then
      CanSolve = False

      If EqIndx > 0 Then
        bSolver(EqIndx) = False
        If .saEqnNames(EqIndx) <> "" Then
          CanSolve = (InStr(.saEqnNames(EqIndx), "<<solver>>") > 0)
          bSolver(EqIndx) = CanSolve
        End If
      End If

      .saEqns(EqIndx) = Trim(.saEqns(EqIndx))
      UnswapEqns(EqIndx) = .saEqns(EqIndx)

      If Not CanSolve Then
        q = InStr(.saEqns(EqIndx), "<=>")

        If q > 0 Then
          Swappo(EqIndx) = Trim(Mid$(.saEqns(EqIndx), q))
          UnswapEqns(EqIndx) = Trim(Left$(.saEqns(EqIndx), q - 1))
        Else
          Swappo(EqIndx) = ""
        End If

      End If
    End If

  Next EqIndx

  For EqIndx = piLwrIndx To .iNeqns
    tB = (EqIndx <> 0 And .saEqns(EqIndx) <> "")

    If tB Then
      ' if U-Pb Special eqn for Th/U but not relevant because
      ' calculating Th/U from 206/238 and 208/232, or
      ' U-Pb special eqn for Pb/Th but not relevant because
      ' calculating from 232/238, or
      ' a General, not Solver-invoke eqn, then skip this eqn

      If (EqIndx = -3 And .bDirectAltPD) _
        Or (EqIndx = -2 And Not .bDirectAltPD) Then tB = False

      If tB And EqIndx > 0 Then
        If bSolver(EqIndx) Then tB = False
      End If

      If tB Then
        ' Determine # of square-bracketed isotope-ratio refs and
        '  other equation refs in the equation
        EqnParse EqIndx, UnswapEqns(EqIndx), piaEqnRats(EqIndx), _
                 .iNrats, EqAsc(), piaNeqnTerms(EqIndx)
        RatCt = 0

        For j = 1 To UBound(EqAsc)
        ' Determine the numerator & denominator nuclides of the ratio.

          If piaBrakType(j, EqIndx) = 1 Then
            ' if the bracketed-expression references an isotope ratio
            RatCt = 1 + RatCt

            For NumDenom = 1 To 2
              NomMass = .daNmDmIso(NumDenom, EqAsc(RatCt)) 'RatCt) ' EqAsc(RatCt))
              FindPkOrd NomMass, piaEqPkOrd(EqIndx, j, NumDenom)
              m = 2 * RatCt - 2 + NumDenom '2 * j - 2 + NumDenom
              piaEqPkUndupeOrd(EqIndx, m) = piaEqPkOrd(EqIndx, j, NumDenom)
            Next NumDenom

          End If

        Next j

        If piaEqnRats(EqIndx) > 0 Then
         ' Remove duplicate peaks from piaEqPkOrd
          m = 2 * (piaNeqnTerms(EqIndx) + piaEqnRats(EqIndx))
          DupeElim piaEqPkUndupeOrd, EqIndx, m, piNoDupePkN(EqIndx), p
        End If

      End If

      .saEqns(EqIndx) = UnswapEqns(EqIndx) & Swappo(EqIndx)
    End If

  Next EqIndx

End With
End Sub

Function fbIsRatioString(ByVal Phrase$, Optional Ratio$, _
  Optional Numer, Optional Denom, Optional RatioNum, Optional Isorats) As Boolean
' Is Phrase is an isotope-ratio string of the form 206/204 or ["206/204"] or h or [h]?
' (in the latter 2 cases, must pass the list of Task ratios as IsoRats$()
' Ratio$ is returned as the ratio-string ("206/204" or "h");
' Numer, Denom are returned as the (numeric) numerator- and denominator isotopes,
' RatioNum is returned as the index# of the ratio if Isorats is passed or if
'  Phrase$ is a  letter-index.

Dim tB As Boolean
Dim Nume$, Deno$
Dim p%, Le1%, Le2%, Nrats%

If fbNIM(Ratio) Then Ratio = ""
If fbNIM(RatioNum) Then RatioNum = 0
If fbNIM(Numer) Then Numer = 0
If fbNIM(Denom) Then Denom = 0
Subst Phrase, psBrQL, , psBrQR
Le1 = Len(Phrase)
Subst Phrase, "[", , "]", , " "
Le2 = Len(Phrase)
tB = False
p = InStr(Phrase, "/")

If p > 0 And p < 5 Then
  Nume = Left$(Phrase, p - 1)

  If fbIsAllNumChars(Nume) Then
    Deno = Mid$(Phrase, p + 1)

    If Len(Deno) > 0 And Len(Deno) < 4 Then

      If fbIsAllNumChars(Deno) Then
        tB = True

        If fbNIM(RatioNum) And fbNIM(Isorats) Then
          Nrats = UBound(Isorats)

          For RatioNum = 1 To Nrats
            If Phrase = Isorats(RatioNum) Then Exit For
          Next RatioNum

          If RatioNum > Nrats Then RatioNum = 0
        End If

      End If

    End If

  End If

ElseIf Le1 > 2 And (Le2 = 1 Or Le2 = 2) And fbNIM(RatioNum) Then

  If fbIsAllAlphaChars(Phrase) Then
    Phrase = UCase(Phrase): tB = True
    RatioNum = Asc(Right$(Phrase, 1)) - 64
    If Le2 = 2 Then RatioNum = RatioNum + 26 * Asc(Left$(Phrase, 1))

    If fbNIM(Isorats) Then
      Ratio = Isorats(RatioNum)

      If fbNIM(Numer) And fbNIM(Denom) Then
        p = InStr(Ratio, "/")
        Nume = Left$(Ratio, p - 1)
        Deno = Mid$(Ratio, p + 1)
      End If

    End If

  End If

End If

If tB Then
  If fbNIM(Ratio) Then Ratio = Phrase
  If fbNIM(Numer) Then Numer = Val(Nume)
  If fbNIM(Denom) Then Denom = Val(Deno)
End If
fbIsRatioString = tB
End Function

Sub EqnParse(ByVal EqnIndx%, EqnIn$, NeqnRats%, IsoRats_N%, EqnIndxCharAsc%(), NeqnTerms%)
' Checks parsing of the specified Task Equation, determines corresponding output-column,
'  # of SQ2 references (ie eqns, ratos, constgants, col-hdrs, range names) & the indexes
'  of each, #ratios...
Dim IsEq As Boolean, IsRat As Boolean
Dim s$, EqnOut$, EqBld$, d$, w$, lw$, LgNa$
Dim RatNum%, j%, BrakNum%, IndxCharAsc%, h%, BrakCt%, m%, b1%, b2%, EqNum%
Dim EqAscQ%, EqNumQ%, Bloc%(1 To 20, 1 To 2), EqNumArr()

ReDim EqNumArr(piLwrIndx To puTask.iNeqns, 99)

BrakNum = 0: s = EqnIn: EqnOut = EqnIn
BrakCt = 0: NeqnTerms = 0: NeqnRats = 0

Do ' Look for paired brackets
  b1 = fiInstanceLoc(s, BrakCt + 1, "[")
  b2 = fiInstanceLoc(s, BrakCt + 1, "]")

  If b1 > 0 And b2 > b1 Then ' paired criteria
    BrakCt = 1 + BrakCt
    Bloc(BrakCt, 1) = b1: Bloc(BrakCt, 2) = b2 ' Bloc(Y,1), Bloc(Y,2) are the
  End If                                       '  start, end positions of the Yth bracket-pair.

Loop Until b1 = 0
' average(["206/207"]

' ********** Use "fbIsRatioString()" *******************
If BrakCt > 0 Then ' the number of bracketed expressions
  With puTask
    EqBld = EqnOut

    For BrakNum = 1 To BrakCt

      s = LCase(Mid$(EqnOut, Bloc(BrakNum, 1) + 1))

      If Bloc(BrakNum, 2) - Bloc(BrakNum, 1) > 2 Then  ' ie aa, ba ...     ' /
        IndxCharAsc = (Asc(Left(s, 1)) - 96) * 26 + Asc(Mid(s, 2, 1)) - 96 '| 10/10/11 -- added so that ratios with index
      Else                                                                 ' \            numbers >26 are correctly parsed.
        IndxCharAsc = Asc(s) - 96
      End If

      m = Val(s)



      If m = 0 And Mid$(EqnOut, Bloc(BrakNum, 1) + 1, 1) = Chr(34) _
          And Mid$(EqnOut, Bloc(BrakNum, 2) + -1, 1) = Chr(34) Then
        w = Mid$(EqnOut, Bloc(BrakNum, 1) + 2, Bloc(BrakNum, 2) - Bloc(BrakNum, 1) - 3)
        lw = LCase(fsLegalName(w))

        For EqNum = 1 To .iNeqns
          LgNa = LCase(fsLegalName(.saEqnNames(EqNum)))
          If LgNa = lw Then
            m = EqNum
            Exit For
          End If
        Next EqNum

      End If

      EqAscQ = 0: EqNumQ = 0: IsRat = False: IsEq = False
      s = Mid$(EqnOut, 1 + Bloc(BrakNum, 1), Bloc(BrakNum, 2) - Bloc(BrakNum, 1) - 1)
      ' s is the qth bracketed expression
      h = fiInstanceCount(pscQ, s)  ' position of the sth quote (")

      ' 10/10/11 -- line below changed from the incorrect "IndxCharAsc < 26" to "IndxCharAsc <= peMaxRats"

      If IndxCharAsc > 0 And IndxCharAsc <= peMaxRats And Len(s) < 3 And _
         Right$(LCase(s), 4) <> ".xls" Then ' a ratio#-indicating letter?


        IsRat = True
        NeqnRats = NeqnRats + 1 ' the ratio number
        EqnIndxCharAsc(BrakNum) = IndxCharAsc
      ElseIf m > 0 Then 'And InStr(s, pscQ) = 0 Then  ' a numbered Equation?
        NeqnTerms = NeqnTerms + 1  ' the equation number
        IsEq = True
        EqNumArr(EqnIndx, BrakNum) = m
        psaEqHdr(EqnIndx, BrakNum) = fsLegalName(fsStrip(.saEqnNames(m)))
        'If Mid$(s, j + 2, 1) = "+" ThenErrColRef(BrakNum) = True
      ElseIf h > 1 And h Mod 2 = 0 And (Bloc(BrakNum, 2) - Bloc(BrakNum, 1) > 1) Then
      ' Is s an expression in quotes?
        s = Mid$(s, 2, Len(s) - 2) ' the quoted expression

        For j = 1 To IsoRats_N     ' Is the expression a defined isotope ratio?
          If .saIsoRats(j) = s Then
            NeqnRats = 1 + NeqnRats
            IsRat = True
            RatNum = j
            EqnIndxCharAsc(NeqnRats) = RatNum
            Exit For
          End If
        Next j

        If IsEq Or True Then       ' Is the expression a defined equation?

          For EqNum = 1 To .iNeqns
            d = fsLegalName(fsStrip(.saEqnNames(EqNum)))
            w = fsLegalName(fsStrip(s))

            If d = w Then
              EqNumArr(EqnIndx, BrakNum) = EqNum
              psaEqHdr(EqnIndx, BrakNum) = d
              IsEq = True
              Exit For
            End If

          Next EqNum

        End If
      Else ' Reference to another workbook?
      End If

      If IsEq Then
        piaBrakType(BrakNum, EqnIndx) = 2
        s = psaEqHdr(EqnIndx, BrakNum)

        w = fsS(EqNumArr(EqnIndx, BrakNum))
          Subst EqBld, "[" & w & "]", psBrQL & s & psBrQR
      ElseIf IsRat Then
        piaBrakType(BrakNum, EqnIndx) = 1
      End If

    Next BrakNum

    If EqBld <> "" And EqBld <> EqnOut Then
      EqnIn = EqBld
    Else
      EqnIn = EqnOut  ' No reason to change to EqBld ???
    End If

  End With
End If
End Sub

Sub DupeElim(ListIn, EqnInd%, ByVal Nin%, Nout%, ByVal Lind%)
' Eliminate any duplicate entries in the list of Task Equation peaks.
Dim i%, ct%, v%
Dim ListOut() As Variant, tmp() As Variant

ReDim tmp(1 To Nin), ListOut(Lind To Nin)

ct = 0
For i = 1 To Nin
  v = ListIn(EqnInd, i)
  If v > 0 Then
    ct = 1 + ct
    tmp(ct) = ListIn(EqnInd, i)
  End If
Next i

ReDim Preserve tmp(1 To ct)
BubbleSort tmp
Nout = 1: ListOut(1) = tmp(1)

For i = 2 To ct
  If tmp(i - 1) <> tmp(i) And tmp(i - 1) <> 0 Then
    Nout = Nout + 1
    ListOut(Nout) = tmp(i)
  End If
Next i

For i = 1 To Nin
  If i > Nin Then
    ListIn(EqnInd, i) = 0
  Else
    ListIn(EqnInd, i) = ListOut(i)
  End If
Next i
End Sub

Sub FindPkOrd(ByVal NominalMass#, PkOrder%)
Dim i% ' Determine the Run Table peak-order of a nuclide

For i = 1 To UBound(puTask.daNominal)
  If puTask.daNominal(i) = NominalMass Then
    PkOrder = i: Exit Sub
  End If
Next i

PkOrder = 0
End Sub

Function fsSquidUserFolder$() ' Return the path of the SquidUser folder
fsSquidUserFolder = ThisWorkbook.Path & fsPathSep & "SquidUser" & fsPathSep
End Function

Function fsSquidDrive$() ' Return the Drive of the SquidUser folder
Dim s$, p%
s = ThisWorkbook.Path
p = InStr(s, fsPathSep)

If p > 0 Then
  fsSquidDrive = Left$(s, p - 1)
Else
  fsSquidDrive = Left$(s, InStr(s, ":"))
End If
End Function

Function fsCurDrive$(Optional Path$ = "")
Dim s$, p% ' Return the current drive

If Path = "" Then
  s = CurDir
Else
  s = Path
End If

p = InStr(s, fsPathSep)

If p > 0 Then
  fsCurDrive = Left$(s, p - 1)
Else
  fsCurDrive = Left$(s, InStr(s, ":"))
End If
End Function

Function fbIsFresh() As Boolean
' Is the saved TaskCatalog the same as the one in this workbooks's Task catalog sheet?
' If not, either rebuild the TaskCatalog or add/delete entgries for missing/absent Tasks.
Dim FoundTask As Boolean, Bad As Boolean, TcatWbkCreated As Boolean
Dim ucTaskCatFilename$, ucSquidUserFilename$
Dim TaskFileNames$(), TaskNames$(), MissingTask$()
Dim i%, j%, TaskType%, p%, NumAllTaskFiles%, Nmissing%, TaskFileNum%, TaskVarNum%
Dim TaskCatWbk As Workbook

Const General = 0, UPb = 1

On Error GoTo 0

fhTempCat.Activate
CopyTempCatShtToTaskCatVar

Restart:
ChDirDrv fsSquidUserFolder
GetTaskFileNames TaskFileNames, TaskNames, NumAllTaskFiles
TcatWbkCreated = False

' Identify any Task in the TaskCat variable that doesn't have a corresponding
'  Task File in the SquidUser folder.
With puTaskCat

  Nmissing = 0
  For TaskType = General To UPb

    For TaskVarNum = 1 To .iaNumTasks(TaskType)
      ucTaskCatFilename = UCase(.saFileNames(TaskType, TaskVarNum))
      If ucTaskCatFilename = "" Then fbIsFresh = False: Exit Function

      For TaskFileNum = 1 To NumAllTaskFiles
        ucSquidUserFilename = UCase(TaskFileNames(TaskFileNum))
        If ucSquidUserFilename = ucTaskCatFilename Then Exit For
      Next TaskFileNum

      If TaskFileNum > NumAllTaskFiles Then
        Nmissing = 1 + Nmissing
        p = InStr(ucTaskCatFilename, "_")
        ucTaskCatFilename = Mid$(ucTaskCatFilename, p + 1)
        p = InStr(ucTaskCatFilename, ".")
        ucTaskCatFilename = Left$(ucTaskCatFilename, p - 1)
        ReDim Preserve MissingTask(1 To Nmissing)
        MissingTask(Nmissing) = ucTaskCatFilename
      End If

    Next TaskVarNum

  Next TaskType

  If Nmissing > 0 Then

    For i = 1 To Nmissing
      DeleteTaskFromTaskcatVar MissingTask(i)
    Next i

    CopyTaskCatVarToTempCatSht
    CopyTempCatShtToTaskCatWbk
  End If

  ' Identify Task Files not present in the Task Variable and add to
  '  the Task Catalog
  Do
    FoundTask = True

    For i = 1 To NumAllTaskFiles

      FoundTask = False
      ucSquidUserFilename = UCase(TaskFileNames(i))

      For TaskType = General To UPb

        For j = 1 To .iaNumTasks(TaskType)
          ucTaskCatFilename = UCase(.saFileNames(TaskType, j))

          If ucTaskCatFilename = ucSquidUserFilename Then
            FoundTask = True
            Exit For
          End If

        Next j

        If FoundTask Then Exit For
      Next TaskType

      If Not FoundTask Then
        StatBar "Adding file to Task Catalog   " & TaskFileNames(i)

        If Not TcatWbkCreated Then
          CreateEmptyTaskCatSheet TaskCatWbk
          TcatWbkCreated = True
          TaskCatWbk.Activate
          CopyTaskCatVarToTempCatSht True
        End If

        AddTaskToTempCatSht TaskFileNames(i), Bad, True
        CopyTempCatShtToTaskCatVar , True
        Exit For
      End If

    Next i

  Loop Until FoundTask
End With

If TcatWbkCreated Then
  Columns.AutoFit
  SortTempCat
  CopyTempCatShtToTaskCatWbk True
  CopyTaskCatWbkToTempCatSht
  Alerts False
  TaskCatWbk.Close
End If

StatBar
fbIsFresh = True
Alerts True
Exit Function

BadCatalog: On Error GoTo 0
BuildTaskCatalog Bad
If Bad Then
  MsgBox "Unable to build the Task Catalog.", , pscSq: End
Else
  GoTo Restart
End If
End Function

Function fsLegalRangeAndSheetName(ByVal Name$, Optional ByVal PeriodsOK = False)
' return a range or sheet name stripped of illegal characters.
' Used only be Sub AddGrafSht.
Dim s As String * 1, t As String * 1, NewNa$, i%, c%, t0%
' alphanumeric and underscore chars are OK
NewNa = ""

For i = 1 To Len(Name)
  t = Mid$(Name, i, 1)
  s = UCase(t)
  c = Asc(s)

  If c = 95 Or (c > 47 And c < 58) Or (c > 64 And c < 91) Or (PeriodsOK And c = 46) Then
    NewNa = NewNa & t
  End If

Next i

If Len(NewNa) = 0 Then
  t0 = 0
Else
  t0 = Asc(Left$(NewNa, 1))
End If

' If the name starts with a number, precede the # with "Sq_"
If t0 = 0 Or (t0 > 47 And t0 < 58) Then
  NewNa = "Sq_" & NewNa
End If

fsLegalRangeAndSheetName = NewNa
DupeNames NewNa, 1
End Function

Function fsLegalName$(ByVal Name$, Optional RangeName As Boolean = False, _
  Optional SpacesOK = False, Optional EqNaOnly As Boolean = True)
' Slightly different function for converting a Name to an acceptable range or
'  worksheet name.
Const NoNo = "*/:?[\].", NoNoEqNa = "|."

Dim c$, Ln$, s As String * 1
Dim p%, q%, ct%, LenNono%

Ln = "": ct = 0
LenNono = Len(NoNo)
If RangeName Then EqNaOnly = False

For p = 1 To Len(Name)
  c = Mid$(Name, p, 1)

  If RangeName Then
    If ct = 0 And fbIsNumChar(c, , True) Then
      Ln = "sq_"
      ct = 3
    End If

    If (fbIsAlphaChar(c) Or c = "," Or c = "_") Or _
       (fbIsNumChar(c, , True) And ct > 0) Then
      ct = 1 + ct: Ln = Ln & c
    End If

  ElseIf Not EqNaOnly Then

    For q = 1 To LenNono
      s = Mid$(NoNo, q)
      If s = c Then c = "": Exit For
    Next q

    ct = 1 + ct: Ln = Ln & c
  Else

    For q = 1 To Len(NoNoEqNa)
      s = Mid$(NoNoEqNa, q)
      If s = c Then c = "": Exit For
    Next q

    ct = 1 + ct: Ln = Ln & c
  End If

Next p

Ln = Trim(Ln)
If Not SpacesOK Then Subst Ln, " "
Subst Ln, "|"
Subst Ln, "/"
fsLegalName = Ln
End Function

Function fbIsLegalName(ByVal Name$, Optional RangeName As Boolean = False) As Boolean
' Is Name a legal Excel sheet- or range name?
' Worksheet names can't contain any of  *?/:[]\"
' Range names can't start with number or contain any non-alphanumeric chars but underscore.
'  !"#$%&'()*,-.  is 32 to 46
' 0123456789 is 48 to 57
' :;<=>?@    is  58 to 64
' A-Z        is 65 to 90
' [/]^_`     is 91 to 93
' a-z        is 97-122
' {|}~       is 123-126

Const LenNono = 7, NoNo = "*/:?[\]" ' 42 47 58 63 91 92 93

Dim Ln As Boolean, c As String * 1
Dim i%, j%, p%
Ln = True

For i = 1 To Len(Name$)
  c = Mid$(Name, i, 1)

  If RangeName Then

    If i = 1 Then
      If fbIsNumChar(c) Then Ln = False: Exit For
    ElseIf Not fbIsNumChar(c) And Not fbIsAlphaChar(c) And c <> "_" Then
      Ln = False: Exit For
    End If

  Else

    For j = 1 To LenNono
      If Mid$(NoNo, j, 1) = c Then Ln = False: Exit For
    Next j

    If j <= LenNono Then Exit For

  End If

Next i
fbIsLegalName = Ln
End Function
