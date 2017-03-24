Attribute VB_Name = "TaskCatalog"
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

' Module TaskCatalog
Option Explicit
Option Base 1

Sub CopyTaskWbkToTempTaskSht(Optional ThenCloseTask As Boolean = True)

Dim Bad As Boolean

With pwTaskBook
  .Activate
  frSr(1, 1, flEndRow(1), 99).Copy fhTempTask.Cells(1, 1)
  fhTempTask.Activate
  Alerts False
  ClearAllClipboard
  GetTaskRows Bad
  If ThenCloseTask Then .Close
End With

fhTempTask.Activate
frSr(1 + puTask.laAutoGrfRw(peMaxAutochts), 1, Cells(9999, 256).End(xlDown).Row, 256).Clear
End Sub

Sub CopyTaskVarsToSheet(ByVal ToTempTask1OrToTaskWbk2%)
' 09/04/10 -- Add code for hidden CPS columns
' 09/06/09 -- Force each isotope-ratio cell (eg "206/208" to be text format to avoid
'             ratios like 7/30 begin converted to text format, eg "30-Jul"'
Dim s$, Cv%, Cn%, Col%, PkIndx%, RatIndx%, EqnIndx%, ConstIndx%, PkNum%
Dim Col1%, MxEq%, MxPk%, MxRat%, MxAu%, ChtNum%
Dim Rw&, rw0&, rw1&

Select Case ToTempTask1OrToTaskWbk2
  Case 1
    fhTempTask.Activate
  Case 2
    OpenWorkbook psTaskFilename, , True
    Set pwTaskBook = ActiveWorkbook
    Cells.NumberFormat = "General"
End Select

Columns(1).ColumnWidth = 20
GetTaskRows
Cv = peTaskValCol
Cn = peTaskNvalsCol
Col1 = Cv - 1

With puTask
  Cells(.lTypeRw, Cv) = IIf(.bIsUPb, "Geochron", "General")
  Cells(.lNameRw, Cv) = .sName
  '.sBySquidVersion = Cells(.lsquidverrw, cV)
  '.sDateCreated = Cells(, dv)
  '.sLastAccessed = Cells(.lLastRevRw, cV)
  Cells(.lLastRevRw, Cv) = .sLastModified
  Cells(.lDescrRw, Cv) = .sDescr
  Cells(.lMineralRw, Cv) = .sMineral
  MxPk = Cv + peMaxNukes - 1
  MxRat = Cv + peMaxRats - 1
  MxEq = Cv + peMaxEqns - 1
  MxAu = Cv + picNumAutochtVars - 1
  frSr(.lNuclidesRw, Cv, , MxPk).ClearContents
  frSr(.lEqnsRw, Cv, , MxEq).ClearContents
  frSr(.lEqnNamesRw, Cv, , MxEq).ClearContents
  frSr(.lRatiosRw, Cv, , MxRat).ClearContents
  frSr(.lNominalMassRw, Cv, , MxPk).ClearContents
  If .lCPScolsRW > 0 Then frSr(.lCPScolsRW, Cv, , MxPk).ClearContents
  If .lHiddenMassRw > 0 Then frSr(.lHiddenMassRw, Cv, , MxPk).ClearContents
  frSr(.lTrueMassRw, Cv, , MxPk).ClearContents
  frSr(.lRatioNumRw, Cv, , MxRat).ClearContents
  frSr(.lRatioDenomRw, Cv, , MxRat).ClearContents
  frSr(.lEqnSubsNaFrRw, Cv, , MxPk).ClearContents
  frSr(.laAutoGrfRw(1), Cv, .laAutoGrfRw(peMaxAutochts), _
        picNumAutochtVars + Cv - 1).ClearContents
  frSr(.lEqnSwSTrw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwSArw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwSCrw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwLArw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwFOrw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwNUrw, Cv, , MxEq).ClearContents
  If .lEqnSwZCrw > 0 Then frSr(.lEqnSwZCrw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwARrw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwARrowsRw, Cv, , MxEq).ClearContents
  frSr(.lEqnSwARcolsRw, Cv, , MxEq).ClearContents
  Cells(.lNuclidesRw, Cn) = .iNpeaks
  Cells(.lNominalMassRw, Cn) = .iNpeaks
  Cells(.lTrueMassRw, Cn) = .iNpeaks
  If .lHiddenMassRw > 0 Then
    Cells(.lHiddenMassRw, Cn) = .iNpeaks
  End If
  If .lCPScolsRW > 0 Then Cells(.lCPScolsRW, Cn) = .iNpeaks
  Cells(.lRatiosRw, Cn) = .iNrats
  Cells(.lRatioNumRw, Cn) = .iNrats
  Cells(.lRatioDenomRw, Cn) = .iNrats
  Cells(.lBkrdmassRw, Cv) = .dBkrdMass
  Cells(.lRefmassRw, Cv) = .dRefTrimMass
  Cells(.lParentNuclideRw, Cv) = .iParentIso
  Cells(.lDirectAltPDrw, Cv) = .bDirectAltPD
  Cells(.lEqnsRw, Cn) = .iNeqns
  Cells(.lEqnNamesRw, Cn) = .iNeqns
  If .lEqnSwHIrw > 0 Then
    Cells(.lEqnSwHIrw, Cn) = .iNeqns
  End If
  Cells(.lEqnSwSTrw, Cn) = .iNeqns
  Cells(.lEqnSwSArw, Cn) = .iNeqns
  Cells(.lEqnSwSCrw, Cn) = .iNeqns
  Cells(.lEqnSwLArw, Cn) = .iNeqns
  Cells(.lEqnSwFOrw, Cn) = .iNeqns
  Cells(.lEqnSwNUrw, Cn) = .iNeqns
  If .lEqnSwZCrw > 0 Then Cells(.lEqnSwZCrw, Cn) = .iNeqns
  Cells(.lEqnSwARrw, Cn) = .iNeqns
  Cells(.lEqnSwARrowsRw, Cn) = .iNeqns
  Cells(.lEqnSwARcolsRw, Cn) = .iNeqns

  For PkIndx = 1 To .iNpeaks
    Col = Col1 + PkIndx
    Cells(.lNuclidesRw, Col) = .saNuclides(PkIndx)
    Cells(.lNominalMassRw, Col) = .daNominal(PkIndx)
    Cells(.lTrueMassRw, Col) = .daTrueMass(PkIndx)
    Cells(.lHiddenMassRw, Col) = IIf(.baHiddenMass(PkIndx), True, False)
    If .lCPScolsRW > 0 Then Cells(.lCPScolsRW, Col) = .baCPScol(PkIndx)
  Next PkIndx

  For RatIndx = 1 To .iNrats
    Col = Col1 + RatIndx
    With Cells(.lRatiosRw, Col)
      .NumberFormat = "@"                   ' essential to prevent ratios like
      .Formula = puTask.saIsoRats(RatIndx)  '   7/30 being converted to a date.
    End With
    Cells(.lRatioNumRw, Col) = .daNmDmIso(1, RatIndx)  ' MUST BE DONE CELL-BY-CELL!
    Cells(.lRatioDenomRw, Col) = .daNmDmIso(2, RatIndx)
  Next RatIndx

  If pbUPb Then

    For EqnIndx = -4 To -1
      s = .saEqns(EqnIndx)

      Select Case EqnIndx
        Case -1: Cells(.lPrimUThPbEqnRw, Cv) = s
        Case -2: Cells(.lSecUThPbEqnRw, Cv) = s
        Case -3: Cells(.lThUeqnRw, Cv) = s
        Case -4: Cells(.lPpmparentEqnRw, Cv) = s
      End Select

    Next EqnIndx

  End If

  For EqnIndx = 1 To .iNeqns
    Col = Col1 + EqnIndx
    Cells(.lEqnsRw, Col).NumberFormat = "@"
    Cells(.lEqnsRw, Col) = .saEqns(EqnIndx)
    Cells(.lEqnNamesRw, Col).NumberFormat = "@"
    Cells(.lEqnNamesRw, Col) = .saEqnNames(EqnIndx)

    If .lEqnSwHIrw > 0 Then
      Cells(.lEqnSwHIrw, Col) = .uaSwitches(EqnIndx).HI
    End If

    With .uaSwitches(EqnIndx)
      Cells(puTask.lEqnSwSTrw, Col) = .ST
      Cells(puTask.lEqnSwSArw, Col) = .SA
      Cells(puTask.lEqnSwSCrw, Col) = .SC
      Cells(puTask.lEqnSwLArw, Col) = .LA
      Cells(puTask.lEqnSwFOrw, Col) = .FO
      Cells(puTask.lEqnSwNUrw, Col) = .Nu
      Cells(puTask.lEqnSwHIrw, Col) = .HI
      If puTask.lEqnSwZCrw > 0 Then Cells(puTask.lEqnSwZCrw, Col) = .ZC
      Cells(puTask.lEqnSwARrw, Col) = .Ar
      Cells(puTask.lEqnSwARrowsRw, Col) = .ArrNrows
      Cells(puTask.lEqnSwARcolsRw, Col) = .ArrNcols
    End With
    Cells(.lEqnSubsNaFrRw, Col) = .saSubsetSpotNa(EqnIndx)
  Next EqnIndx

  If .lConstNamesRw > 0 Then
    frSr(.lConstNamesRw, Cn, , 256).ClearContents
    Cells(.lConstNamesRw, Cn) = .iNconsts
    For ConstIndx = 1 To .iNconsts
      Cells(.lConstNamesRw, Cv + ConstIndx - 1) = .saConstNames(ConstIndx)
    Next ConstIndx
  End If

  If .lConstValsRw > 0 Then
    frSr(.lConstValsRw, Cn, , 256).ClearContents
    Cells(.lConstValsRw, Cn) = .iNconsts

    For ConstIndx = 1 To .iNconsts
      Cells(.lConstValsRw, Cv + ConstIndx - 1) = .vaConstValues(ConstIndx)
    Next ConstIndx

  End If

  frSr(.laAutoGrfRw(1), peTaskIndxCol, .laAutoGrfRw(peMaxAutochts), 256).ClearContents
  Cells(.laAutoGrfRw(1), peTaskIndxCol) = .iNumAutoCharts
  rw1 = .laAutoGrfRw(1)
  frSr(rw1, 1, rw1 + 3, 256).Copy
  frSr(rw1 + 4, 1, rw1 + peMaxAutochts - 1, 256).PasteSpecial xlFormats
  frSr(rw1, 1, rw1 - 1 + peMaxAutochts, 256).ClearContents
  frSr(rw1 + peMaxAutochts, 1, Cells(9999, 256).End(xlDown).Row, 256).Clear
  rw0 = rw1

  For ChtNum = 1 To peMaxAutochts
    Rw = .laAutoGrfRw(ChtNum)
    If Rw = 0 Then Rw = rw0 + 1
    Cells(Rw, 1).Formula = "AutoGraf" & fsS(ChtNum)
    If ChtNum = 1 Then Cells(Rw, peTaskIndxCol) = .iNumAutoCharts

    If ChtNum <= .iNumAutoCharts Then
      With .uaAutographs(ChtNum)
        Cells(Rw, Cn) = picNumAutochtVars
        Cells(Rw, Cv) = .sXname
        Cells(Rw, Cv + 1) = .bAutoscaleX
        Cells(Rw, Cv + 2) = .bZeroXmin
        Cells(Rw, Cv + 3) = .bLogX
        Cells(Rw, Cv + 4) = .sYname
        Cells(Rw, Cv + 5) = .bAutoscaleY
        Cells(Rw, Cv + 6) = .bZeroYmin
        Cells(Rw, Cv + 7) = .bLogY
        Cells(Rw, Cv + 8) = .bRegress
        Cells(Rw, Cv + 9) = .bAverage
      End With
    End If

    rw0 = Rw
  Next ChtNum

  Cells.Interior.Color = vbWhite
  Cells.Font.Bold = False

  For Rw = puTask.lFileNameRw To flEndRow Step 2
    Rows(Rw).Interior.Color = pePaleGreen
  Next Rw

End With

ColWidth picAuto, 1
If ToTempTask1OrToTaskWbk2 = 2 Then
  Alerts False
  On Error Resume Next
  pwTaskBook.Save
  pwTaskBook.Close
  On Error GoTo 0
  Alerts True
End If
End Sub

Sub CopyTaskToTaskVars(ByVal FromTempTask1_OrTaskWbk2%, Bad As Boolean, Optional TaskCatWbk)
' 09/04/10 -- Add code for hidden CPS columns
Dim Cv%, Cn%, ConstIndx%, RatIndx%, Col%, CelCol%, NumChts%
Dim EqnIndx%, ChtIndx%, PkIndx%, ChtRw&
Dim rw1&, rw0&, Rw&
'Dim Ce1 As Range

Bad = True

Select Case FromTempTask1_OrTaskWbk2
  Case 1
    fhTempTask.Activate
  Case 2
    pwTaskBook.Activate
End Select

GetTaskRows
Cv = peTaskValCol
Cn = peTaskNvalsCol
CelCol = Cv - 1

With puTask
  .bIsUPb = (Cells(.lTypeRw, Cv) = "Geochron")
  .sName = Cells(.lNameRw, Cv)
  .sFileName = Cells(.lFileNameRw, Cv)
  .sLastModified = Cells(.lLastRevRw, Cv)
  .sDescr = Cells(.lDescrRw, Cv)
  .sMineral = Cells(.lMineralRw, Cv)
  .sCreator = Cells(.lDefByRw, Cv)
  .iNpeaks = Cells(.lNuclidesRw, Cn)

  If .iNpeaks < 2 Then
    If MsgBox("No entries in the Run Table of Task " & fsInQ(.sName) & "." _
            & pscLF2 & "Delete this Task file?", vbYesNo, pscSq) = vbYes Then
      Alerts False
      On Error Resume Next
      pwTaskBook.Close
      If fbNIM(TaskCatWbk) Then
        TaskCatWbk.Close
      End If
      Kill fsSquidUserFolder & .sFileName
      On Error GoTo 0
      Alerts True
    End If
    End
  End If

  .iNrats = Cells(.lRatiosRw, Cn)

  If .iNrats = 0 Then
    MsgBox "No isotope ratios have been specified in Task " & _
     .sName & pscLF2 & "The Task File has been renamed ~" & .sFileName, , pscSq
    On Error Resume Next
    Alerts False
    pwTaskBook.Close
    Name fsSquidUserFolder & puTask.sFileName As "~" & .sFileName
    On Error GoTo 0
    End
  End If

  .iNeqns = Cells(.lEqnsRw, Cn)
  .dBkrdMass = Cells(.lBkrdmassRw, Cv)
  .dRefTrimMass = Cells(.lRefmassRw, Cv)
  .iParentIso = Cells(.lParentNuclideRw, Cv)
  .bDirectAltPD = Cells(.lDirectAltPDrw, Cv)

  If .bIsUPb And .iParentIso <> 232 And .iParentIso <> 238 Then
    MsgBox "Invalid parent isotope (must be 232 or 238)", , pscSq
    Exit Sub
  End If

  ReDim .saNuclides(1 To .iNpeaks), .daNominal(1 To .iNpeaks), _
        .daTrueMass(1 To .iNpeaks), .baCPScol(1 To .iNpeaks), .baHiddenMass(1 To .iNpeaks)

  For PkIndx = 1 To .iNpeaks
    Col = CelCol + PkIndx
    .saNuclides(PkIndx) = Cells(.lNuclidesRw, Col)
    .daNominal(PkIndx) = Cells(.lNominalMassRw, Col)
    .daTrueMass(PkIndx) = Cells(.lTrueMassRw, Col)
    If .lCPScolsRW > 0 Then
      .baCPScol(PkIndx) = Cells(.lCPScolsRW, Col)
     Else
      .baCPScol(PkIndx) = False
    End If
    If .lHiddenMassRw > 0 Then
      .baHiddenMass(PkIndx) = Cells(.lHiddenMassRw, Col)
    Else
      .baHiddenMass(PkIndx) = False
    End If
  Next PkIndx

  ReDim .saIsoRats(1 To .iNrats), .daNmDmIso(1 To 2, 1 To .iNrats)

  For RatIndx = 1 To .iNrats
    Col = CelCol + RatIndx
    .saIsoRats(RatIndx) = Cells(.lRatiosRw, Col)
    .daNmDmIso(1, RatIndx) = Cells(.lRatioNumRw, Col)
    .daNmDmIso(2, RatIndx) = Cells(.lRatioDenomRw, Col)
  Next RatIndx

  ReDim .saEqns(-4 To .iNeqns), .saEqnNames(-4 To .iNeqns)

  If .iNeqns > 0 Then
    ReDim .uaSwitches(1 To .iNeqns), .saSubsetSpotNa(1 To .iNeqns)
  End If

  If .bIsUPb Then

    For EqnIndx = -4 To -1

      Select Case EqnIndx
        Case -1: .saEqns(EqnIndx) = Cells(.lPrimUThPbEqnRw, Cv)
        Case -2: .saEqns(EqnIndx) = Cells(.lSecUThPbEqnRw, Cv)
        Case -3: .saEqns(EqnIndx) = Cells(.lThUeqnRw, Cv)
        Case -4: .saEqns(EqnIndx) = Cells(.lPpmparentEqnRw, Cv)
      End Select

    Next EqnIndx

    If .saEqns(-1) = "" Then MsgBox "Undefined Pb/U (or Pb/Th) equation", , pscSq: Exit Sub
  End If

  For EqnIndx = 1 To .iNeqns
    Col = CelCol + EqnIndx
    .saEqns(EqnIndx) = Cells(.lEqnsRw, Col)
    .saEqnNames(EqnIndx) = Cells(.lEqnNamesRw, Col)
    With .uaSwitches(EqnIndx)
      .ST = Cells(puTask.lEqnSwSTrw, Col)
      .SA = Cells(puTask.lEqnSwSArw, Col)
      .SC = Cells(puTask.lEqnSwSCrw, Col)
      .LA = Cells(puTask.lEqnSwLArw, Col)
      .FO = Cells(puTask.lEqnSwFOrw, Col)
      .HI = Cells(puTask.lEqnSwHIrw, Col)
      .Nu = Cells(puTask.lEqnSwNUrw, Col)
      If puTask.lEqnSwZCrw > 0 Then .ZC = Cells(puTask.lEqnSwZCrw, Col) Else .ZC = False
      .Ar = Cells(puTask.lEqnSwARrw, Col)
      .ArrNrows = Cells(puTask.lEqnSwARrowsRw, Col)
      .ArrNcols = Cells(puTask.lEqnSwARcolsRw, Col)
    End With
  Next EqnIndx

  For EqnIndx = 1 To .iNeqns
    .saSubsetSpotNa(EqnIndx) = Cells(.lEqnSubsNaFrRw, EqnIndx + Cv - 1)
  Next EqnIndx

  If .lConstNamesRw > 0 And .lConstValsRw > 0 Then
    .iNconsts = Cells(.lConstValsRw, Cn)

    If .iNconsts > 0 Then
      ReDim .saConstNames(1 To .iNconsts), .vaConstValues(1 To .iNconsts)

      For ConstIndx = 1 To .iNconsts
        .saConstNames(ConstIndx) = Cells(.lConstNamesRw, Cv + ConstIndx - 1)
        .vaConstValues(ConstIndx) = Cells(.lConstValsRw, Cv + ConstIndx - 1)
      Next ConstIndx

    End If

  End If

  rw1 = .laAutoGrfRw(1)
  frSr(rw1, 1, rw1 + 3, 256).Copy
  frSr(rw1 + 4, 1, rw1 + peMaxAutochts - 1, 256).PasteSpecial xlFormats
  frSr(rw1 + peMaxAutochts, 1, Cells(9999, 256).End(xlDown).Row, 256).Clear
  NumChts = 0: rw0 = rw1

  For ChtIndx = 1 To peMaxAutochts
    Rw = .laAutoGrfRw(ChtIndx)
    If Rw = 0 And ChtIndx > 1 Then Rw = rw0 + 1
    Cells(Rw, 1).Formula = "AutoGraf" & fsS(ChtIndx)
    If Cells(Rw, Cv + 4) <> "" Then NumChts = NumChts + 1
    rw0 = Rw
  Next ChtIndx

  .iNumAutoCharts = NumChts

  If .iNumAutoCharts > 0 Then
    ReDim .uaAutographs(1 To .iNumAutoCharts)

    For ChtIndx = 1 To .iNumAutoCharts
      ChtRw = .laAutoGrfRw(ChtIndx)
      With .uaAutographs(ChtIndx)
        .sXname = Cells(ChtRw, Cv)
        .bAutoscaleX = Cells(ChtRw, Cv + 1)
        .bZeroXmin = Cells(ChtRw, Cv + 2)
        .bLogX = Cells(ChtRw, Cv + 3)
        .sYname = Cells(ChtRw, Cv + 4)
        .bAutoscaleY = Cells(ChtRw, Cv + 5)
        .bZeroYmin = Cells(ChtRw, Cv + 6)
        .bLogY = Cells(ChtRw, Cv + 7)
        .bRegress = Cells(ChtRw, Cv + 8)
        .bAverage = Cells(ChtRw, Cv + 9)
      End With
    Next ChtIndx

  End If
End With
Bad = False
End Sub

Sub CopyTempTaskShtToTaskWbk()
Dim Bad As Boolean, Na$
CreateNewWorkbook
Set pwTaskBook = ActiveWorkbook
StatBar "Saving " & pwTaskBook.Name

With pwTaskBook
  .Sheets(1).Name = "Task"
  fhTempTask.Activate
  frSr(1, 1, flEndRow, 256).Copy .Sheets(1).Cells(1, 1)
  ChDirDrv fsSquidUserFolder, Bad
  If Bad Then Exit Sub
  GetSaveTaskVal 1, Na, , "File name", , , peTaskValCol
  pwTaskBook.Sheets(1).Columns(1).ColumnWidth = 24
  frSr(1 + puTask.laAutoGrfRw(peMaxAutochts), 1, 9999, 256).Clear
  Alerts False
  .SaveAs Na
  On Error GoTo 0
  StatBar
  .Close
End With
fhTempTask.Activate
End Sub

Sub GetTaskCatVarNominal(ByVal TaskName$, Npks%, Optional NominalMasses)
Dim Got As Boolean, Typ%, TaskNum%, NukeNum%
With puTaskCat
  Got = False

  For Typ = 1 To 0 Step -1

    For TaskNum = 1 To .iaNumTasks(Typ)
      If .saNames(Typ, TaskNum) = TaskName Then Got = True: Exit For
    Next TaskNum

    If Got Then Exit For
  Next Typ

  If Not Got Then Error 9994
  Npks = .iaNpeaks(Typ, TaskNum)

  If fbNIM(NominalMasses) Then
    ReDim NominalMasses(1 To Npks)

    For NukeNum = 1 To Npks
      NominalMasses(NukeNum) = .daNominalMass(Typ, NukeNum, TaskNum%)
    Next NukeNum

  End If

End With
End Sub

Sub GetTaskCatVarSubsets(ByVal TaskName$, SubsetEqnNames, SubsetNameFrags, MaxSubsEqnNum%, _
  Optional AllEqnNames As Boolean = False, Optional Neqns)
' If namesonly, return all task equation names; if not, return only eqnnames 1 to MasSubsEqnNum

Dim Got As Boolean, Typ%, TaskNum%, EqNum%, Neq%

If TypeName(SubsetEqnNames) <> "Range" Then
  ReDim SubsetEqnNames(1 To peMaxEqns), SubsetNameFrags(1 To peMaxEqns)
End If

With puTaskCat
  Got = False

  For Typ = 1 To 0 Step -1

    For TaskNum = 1 To .iaNumTasks(Typ)
      If .saNames(Typ, TaskNum) = TaskName Then Got = True: Exit For
    Next TaskNum

    If Got Then Exit For
  Next Typ

  If Not Got Then Error 9995

  For EqNum = 1 To peMaxEqns
    TcatStringToArray .saEqnNaList(Typ, TaskNum), Typ, TaskNum, .saEqnNa, 0
    TcatStringToArray .saEqnSubsetNaList(Typ, TaskNum), Typ, TaskNum, .saEqnSubsetNa, 0
    SubsetEqnNames(EqNum) = .saEqnNa(Typ, EqNum, TaskNum)
    SubsetNameFrags(EqNum) = .saEqnSubsetNa(Typ, EqNum, TaskNum)
  Next EqNum

  MaxSubsEqnNum = 0

  For EqNum = peMaxEqns To 1 Step -1

    If SubsetNameFrags(EqNum) <> "" Then
      MaxSubsEqnNum = EqNum
      Exit For
    End If

  Next EqNum

  Neq = .iaNeqns(Typ, TaskNum)
  If fbNIM(Neqns) Then Neqns = Neq
End With
End Sub

Sub CopyTaskCatVarToTempCatSht(Optional NotTempCat As Boolean = False)

Dim Bad As Boolean, TaskTyp%, TaskIndx%, TotCt%, UPbCt%, GeneralCt%, Rw&

GetTaskCatRowsCols Bad, NotTempCat
If Not NotTempCat Then fhTempCat.Activate
NumForTaskcat True

With puTaskCat
  frSr(.lFirstRw, 1, 1 + flEndRow, 99).ClearContents
  UPbCt = .iaNumTasks(1): GeneralCt = .iaNumTasks(0)
  TotCt = UPbCt + GeneralCt
  Cells(.lUPbNtasksRw, .iUPbNtasksCol) = UPbCt
  Cells(.lGenNtasksRw, .iGenNtasksCol) = GeneralCt
  Cells(.lTotNtasksRw, .iTotNtasksCol) = TotCt
  TotCt = 0

  For TaskTyp = 1 To 0 Step -1

    For TaskIndx = 1 To .iaNumTasks(TaskTyp)
      TotCt = 1 + TotCt
      Rw = .lFirstRw + TotCt - 1
      Cells(Rw, .iNameCol) = .saNames(TaskTyp, TaskIndx)
      Cells(Rw, .iFileNameCol) = .saFileNames(TaskTyp, TaskIndx)
      Cells(Rw, .iCreatorCol) = .saCreators(TaskTyp, TaskIndx)
      Cells(Rw, .iMinCol) = .saMinerals(TaskTyp, TaskIndx)
      Cells(Rw, .iTypeCol) = Choose(1 + TaskTyp, "General", "UPb")
      Cells(Rw, .iDescrCol) = .saDescr(TaskTyp, TaskIndx)
      Cells(Rw, .iNpksCol) = .iaNpeaks(TaskTyp, TaskIndx)
      Cells(Rw, .iNominalCol) = .saNomiMassList(TaskTyp, TaskIndx)
      Cells(Rw, .iTrueCol) = .saTrueMassList(TaskTyp, TaskIndx)
      Cells(Rw, .iNeqnsCol) = .iaNeqns(TaskTyp, TaskIndx)
      Cells(Rw, .iEqnNaCol) = .saEqnNaList(TaskTyp, TaskIndx)
      Cells(Rw, .iSubsNaCol) = .saEqnSubsetNaList(TaskTyp, TaskIndx)
    Next TaskIndx

  Next TaskTyp

End With

End Sub

Sub CopyTempCatShtToTaskCatVar(Optional Bad As Boolean = False, _
                         Optional NotTempCat As Boolean = False)
Dim Typ$, TaskIndx%, TaskTyp%, TotNumTasks%, MaxCt%, ct%, UPbCt%, GeneralCt%, Rw&
GetInfo
If Not NotTempCat Then fhTempCat.Activate

GetTaskCatRowsCols Bad, NotTempCat
If Bad Then GoTo NoRowsCols

With puTaskCat
  ct = 0

  Do
    UPbCt = Cells(.lUPbNtasksRw, .iUPbNtasksCol)
    GeneralCt = Cells(.lGenNtasksRw, .iGenNtasksCol)
  If UPbCt > 0 Or GeneralCt > 0 Then Exit Do
    BuildTaskCatalog
    ct = 1 + ct
    If ct > 1 Then MsgBox "Unable to build Task Catalog": End
  Loop

  TotNumTasks = UPbCt + GeneralCt
  MaxCt = fvMax(UPbCt, GeneralCt)
  RedimTaskCat MaxCt, TotNumTasks, False, False
  ReDim .saEqnNa(0 To 1, 1 To peMaxEqns, 1 To MaxCt), .saEqnSubsetNa(0 To 1, 1 To peMaxEqns, 1 To MaxCt)
  ReDim .daNominalMass(0 To 1, 1 To peMaxNukes, 1 To MaxCt), .daTrueMass(0 To 1, 1 To peMaxNukes, 1 To MaxCt)
  UPbCt = 0: GeneralCt = 0: ct = 0

  For TaskIndx = 1 To TotNumTasks
    Rw = .lFirstRw + TaskIndx - 1
    Typ = LCase(Cells(Rw, .iTypeCol))
    .baTypes(TaskIndx) = (Typ = "upb")
    TaskTyp = -.baTypes(TaskIndx)

    If .baTypes(TaskIndx) Then
      UPbCt = UPbCt + 1
      ct = UPbCt
    Else
      GeneralCt = GeneralCt + 1
      ct = GeneralCt
    End If

    .saNames(TaskTyp, ct) = Trim(Cells(Rw, .iNameCol))
    .saFileNames(TaskTyp, ct) = Cells(Rw, .iFileNameCol)
    .saCreators(TaskTyp, ct) = Cells(Rw, .iCreatorCol)
    .saMinerals(TaskTyp, ct) = Cells(Rw, .iMinCol)
    .saDescr(TaskTyp, ct) = Cells(Rw, .iDescrCol)
    .iaNpeaks(TaskTyp, ct) = Cells(Rw, .iNpksCol)
    .saNomiMassList(TaskTyp, ct) = Cells(Rw, .iNominalCol)
    .saTrueMassList(TaskTyp, ct) = Cells(Rw, .iTrueCol)
    .iaNeqns(TaskTyp, ct) = Cells(Rw, .iNeqnsCol)
    .saEqnNaList(TaskTyp, ct) = Cells(Rw, .iEqnNaCol)
    .saEqnSubsetNaList(TaskTyp, ct) = Cells(Rw, .iSubsNaCol)
    TcatStringToArray .saEqnNaList(TaskTyp, ct), TaskTyp, ct, .saEqnNa, 0
    TcatStringToArray .saEqnSubsetNaList(TaskTyp, ct), TaskTyp, ct, .saEqnSubsetNa, 0
    TcatStringToArray .saNomiMassList(TaskTyp, ct), TaskTyp, ct, .daNominalMass, 1
    TcatStringToArray .saTrueMassList(TaskTyp, ct), TaskTyp, ct, .daTrueMass, 1
  Next TaskIndx

  .iaNumTasks(0) = GeneralCt
  .iaNumTasks(1) = UPbCt
  .iNumAllTasks = GeneralCt + UPbCt
End With

Exit Sub

NoRowsCols: Alerts True
On Error GoTo 0
MsgBox "The Task Catalog file appears to be corrupt." & pscLF2 & "Please " _
       & "request a rebuild from the Preferences panel.", , pscSq
End
End Sub

Sub CopyTempCatShtToTaskCatWbk(Optional NotTempCat As Boolean = False)
' Copy the SQUID-internal TaskCatalog sheet to an external workbook and save.

Dim iNu%, iNg%, iNa%, TaskTyp%
Dim StartRW&(0 To 1), EndRw&(0 To 1)
Dim TaskCatWbk As Workbook, ShtIn As Worksheet

Set ShtIn = ActiveSheet
If Not NotTempCat Then fhTempCat.Activate

With puTaskCat
  iNa = Cells(.lTotNtasksRw, .iTotNtasksCol)
  iNu = Cells(.lUPbNtasksRw, .iUPbNtasksCol)
  iNg = Cells(.lGenNtasksRw, .iGenNtasksCol)
  StartRW(1) = .lFirstRw
  StartRW(0) = StartRW(1) + iNu
  EndRw(1) = StartRW(1) + iNu - 1
  EndRw(0) = EndRw(1) + iNg
  foAp.CutCopyMode = False

  For TaskTyp = 0 To 1
    frSr(StartRW(TaskTyp), .iNameCol, EndRw(TaskTyp), 99).Sort _
      Key1:=Cells(StartRW(TaskTyp), .iNameCol), Order1:=xlAscending
  Next TaskTyp

End With

StatBar "Saving TaskCatalog.xls"
CreateNewWorkbook
NumForTaskcat
Set TaskCatWbk = ActiveWorkbook
NoGridlines

If NotTempCat Then
  ShtIn.Activate
Else
  fhTempCat.Activate
End If

On Error Resume Next
frSr(1, 1, EndRw(0), 26).Copy TaskCatWbk.Sheets(1).Cells(1, 1)
TaskCatWbk.Activate
Columns.AutoFit
ChDirDrv fsSquidUserFolder
Alerts False

With TaskCatWbk
  .SaveAs "TaskCatalog.xls"
  .Close
End With
Alerts True
StatBar
End Sub

Sub CopyTaskCatWbkToTempCatSht(Optional Bad As Boolean = False)
' Copy the external TaskCatalog file onto the SQUID-internal TaskCatalog sheet

Dim Exists As Boolean, TaskCatWbk As Workbook

OpenWorkbook fsSquidUserFolder & "TaskCatalog.xls", Exists

If Not Exists Then
  MsgBox "Unable to open the Task Catalog workbook.", , pscSq
  Bad = True: Exit Sub
End If

Set TaskCatWbk = ActiveWorkbook
NumForTaskcat True
Cells.Copy fhTempCat.Cells(1, 1)
Alerts False
TaskCatWbk.Close
Alerts True
Bad = False
End Sub

Sub GetTaskCatRowsCols(Optional BadCat As Boolean, Optional NotTempCat As Boolean = False)

Dim Msg$, iNa%, iNg%, iNu%, CodeLine%, HdrIndx%
Dim Rw&, Col%, HdrRw&
Dim CodeLoc As Variant, ColHdrs As Variant

If Not NotTempCat Then fhTempCat.Activate
BadCat = True
CodeLoc = Array("# Tasks", "# UPbTasks", "# GeneralTasks", "Task Name")
ColHdrs = Array("File Name", "Creator", "Mineral", "Type", "Description", _
             "# peaks", "Nominal Masses", "True masses", "# Equations", _
             "Equation names", "Equation Spot-Subset Names")
CodeLine = 0: HdrIndx = 0

With puTaskCat
  CodeLine = 1
  FindStr "# Tasks", .lTotNtasksRw, Col, 1, 1, 19, 19
  If .lTotNtasksRw = 0 Then FindStr "#Tasks", .lTotNtasksRw, Col, 1, 1, 19, 19
  If .lTotNtasksRw = 0 Then GoTo BadCat
  .iTotNtasksCol = 1 + Col
  iNa = Cells(.lTotNtasksRw, Col + 1)
  CodeLine = 2
  FindStr "# UPbTasks", .lUPbNtasksRw, Col, 1, 1, 19, 19
  If .lUPbNtasksRw = 0 Then GoTo BadCat
  .iUPbNtasksCol = 1 + Col
  iNu = Cells(.lUPbNtasksRw, Col + 1)
  CodeLine = 3
  FindStr "# GeneralTasks", .lGenNtasksRw, Col, 1, 1, 19, 19
  If .lGenNtasksRw = 0 Then GoTo BadCat
  .iGenNtasksCol = 1 + Col
  iNg = Cells(.lGenNtasksRw, Col + 1)
  CodeLine = 4
  FindStr "Task Name", HdrRw, .iNameCol, 1, 1, 19, 19
  If HdrRw = 0 Or .iNameCol = 0 Then GoTo BadCat
  .lFirstRw = HdrRw + 1

  For HdrIndx = 1 To UBound(ColHdrs)
    FindStr ColHdrs(HdrIndx), , Col, HdrRw, 1, HdrRw, 19
    If Col = 0 Then GoTo BadCat

    Select Case HdrIndx
      Case 1:  .iFileNameCol = Col
      Case 2:  .iCreatorCol = Col
      Case 3:  .iMinCol = Col
      Case 4:  .iTypeCol = Col
      Case 5:  .iDescrCol = Col
      Case 6:  .iNpksCol = Col
      Case 7:  .iNominalCol = Col
      Case 8:  .iTrueCol = Col
      Case 9:  .iNeqnsCol = Col
      Case 10: .iEqnNaCol = Col
      Case 11: .iSubsNaCol = Col
    End Select

  Next HdrIndx

End With
FindCatStartEndRows
BadCat = False
Exit Sub

BadCat: If HdrIndx = 0 Then Msg = CodeLoc(CodeLine) Else Msg = ColHdrs(HdrIndx)
MsgBox "Couldn't find  " & Msg & "  in Task Catalog."
End
End Sub

Sub FindCatStartEndRows()
Dim iNu%, iNg%
With puTaskCat
  iNu = fhTempCat.Cells(.lUPbNtasksRw, .iUPbNtasksCol)
  iNg = fhTempCat.Cells(.lGenNtasksRw, .iGenNtasksCol)
  .lFirstUPbRw = IIf(iNu > 0, .lFirstRw, 0)
  .lLastUPbRw = IIf(iNu > 0, .lFirstRw + iNu - 1, 0)

  If iNg > 0 Then
    .lFirstGenRW = IIf(iNu > 0, 1 + .lLastUPbRw, .lFirstRw)
    .lLastGenRW = .lFirstGenRW + iNg - 1
  Else
    .lFirstGenRW = 0
    .lLastGenRW = 0
  End If

End With
End Sub

Sub AddTaskToTempCatSht(ByVal TaskFileName$, Optional BadAdd As Boolean = False, _
  Optional LeaveCopycatOpen As Boolean = False)

' Extract information from the specified Task file and put the info
'  in the first empty row in the TempCat sheet.  Update the 3 "number of
'  Tasks" cells in TempCat, but do not sort the Task names.
' Assumes the active workbook contains the Task Catalog

Dim Exists As Boolean, Bad As Boolean
Dim p%, i%
Dim TcatRw&, tmp1$, tmp2$, Msg$, StoredFilename$, tmpStored$, tmpCat$
Dim t As SquidTask, Cel As Range, TcatWbk As Workbook

Alerts False
Set TcatWbk = ActiveWorkbook
OpenWorkbook fsSquidUserFolder & TaskFileName, Exists

If Not Exists Then
  MsgBox "Unable to find or open the   " & TaskFileName & _
          "   workbook.", , pscSq
  BadAdd = True
  Exit Sub
End If

GetTaskRows
StoredFilename = Cells(puTask.lFileNameRw, peTaskValCol)

' Look for conflict between actual Task-file name and
'  the filename stored with the actual Task workbook.
' say filename is SquidTask_Titanite.KRL.xls
tmpCat = Mid$(TaskFileName, 11)            ' Titanite.KRL.xls
p = InStr(tmpCat, ".")
If p < 2 Then p = 2
tmpCat = Left$(tmpCat, p - 1)              ' Titanite
tmpStored = Mid$(StoredFilename, 11)       ' Titanite.KRL.xls
p = InStr(tmpStored, ".")
tmpStored = Left$(tmpStored, p - 1)        ' Titanite

If LCase(tmpCat) <> LCase(tmpStored) Then
  On Error Resume Next
  Workbooks(TaskFileName).Close
  On Error GoTo 0
  Msg = "The Task filename stored in the Task workbook is " _
         & fsInQ(StoredFilename)
  Msg = "Task file is corrupt!" & pscLF2 & Msg & ", " & vbLf & _
        "but the workbook's actual name is " & fsInQ(TaskFileName) _
         & " ." & pscLF2
  tmp1 = "~" & TaskFileName

  Do

'    If Left$(tmp1, 6) = "~~~~~~" Then
'      Msg = Msg & "Please delete or repair the corrupt file."
'
'      Exit Do
'    End If

    tmp1 = "~" & tmp1
    On Error GoTo 11
    Name TaskFileName As tmp1
    Msg = Msg & "The file has been renamed " & tmp1 & " ."
    GoTo 12
11
  Loop

12
  On Error Resume Next
  Workbooks(TaskFileName).Close
  TcatWbk.Close
  On Error GoTo 0
  MsgBox Msg, vbOKOnly, pscSq
  MsgBox "Exiting SQUID.", , pscSq
  End
End If

Set pwTaskBook = ActiveWorkbook
If Not LeaveCopycatOpen Then GetTaskCatRowsCols
CopyTaskToTaskVars 2, Bad, TcatWbk
pwTaskBook.Close

If Bad Then
  MsgBox "The Task workbook  " & StoredFilename & "  is corrupt." & pscLF2 & _
          "Please delete or repair.", , pscSq
  BadAdd = True
  Exit Sub
End If


TcatRw = flEndRow(1) + 1
t = puTask

' Extract TaskCatalog info from the Task variable
With puTaskCat
  Cells(TcatRw, .iNameCol) = tmpCat
  Cells(TcatRw, .iFileNameCol) = TaskFileName
  Cells(TcatRw, .iCreatorCol) = t.sCreator
  Cells(TcatRw, .iMinCol) = t.sMineral
  Cells(TcatRw, .iTypeCol) = IIf(t.bIsUPb, "UPb", "General")
  Cells(TcatRw, .iDescrCol) = t.sDescr
  Cells(TcatRw, .iNpksCol) = t.iNpeaks
  Cells(TcatRw, .iNeqnsCol) = t.iNeqns
  tmp1 = "": tmp2 = ""

  For i = 1 To t.iNpeaks
    tmp1 = tmp1 & fsS(t.daNominal(i))
    tmp2 = tmp2 & fsS(t.daTrueMass(i))

    If i < t.iNpeaks Then
      tmp1 = tmp1 & ", " ' DO NOT change to ","
      tmp2 = tmp2 & ", " ' DO NOT change to ","
    End If

  Next i

  Cells(TcatRw, .iNominalCol).NumberFormat = "@"
  Cells(TcatRw, .iTrueCol).NumberFormat = "@"
  Cells(TcatRw, .iEqnNaCol).NumberFormat = "@"
  Cells(TcatRw, .iSubsNaCol).NumberFormat = "@"
  Cells(TcatRw, .iNominalCol) = tmp1
  Cells(TcatRw, .iTrueCol) = tmp2
  tmp1 = "": tmp2 = ""

  For i = 1 To t.iNeqns
    tmp1 = tmp1 & t.saEqnNames(i)
    tmp2 = tmp2 & t.saSubsetSpotNa(i)

    If i < t.iNeqns Then
      tmp1 = tmp1 & ", " ' DO NOT change to ","
      tmp2 = tmp2 & ", " ' DO NOT change to ","
    End If

  Next i

  Cells(TcatRw, .iEqnNaCol) = tmp1
  Cells(TcatRw, .iSubsNaCol) = tmp2

  If t.bIsUPb Then
    Set Cel = Cells(.lUPbNtasksRw, .iUPbNtasksCol)
  Else
    Set Cel = Cells(.lGenNtasksRw, .iGenNtasksCol)
  End If

  Cel = 1 + Cel
  Set Cel = Cells(.lTotNtasksRw, .iTotNtasksCol)
  Cel = 1 + Cel
End With

End Sub

Sub SortTempCat()
' Alphabetically sort the open Task Catalog worksheet
' Sort the UPb and General tasks separately.

Dim UPbCt%, GenCt%, Nupb%, Ngen%, Ntot%
Dim FirstUPbRw&, LastUPbRw&, FirstGenRw&, LastGenRw&, Rw&, rw1&, rw2&
Dim UPbRange As Range, GenRange As Range

GetTaskCatRowsCols , True
With puTaskCat
  Ntot = Cells(.lTotNtasksRw, .iTotNtasksCol)
  Ngen = Cells(.lGenNtasksRw, .iGenNtasksCol)
  Nupb = Cells(.lUPbNtasksRw, .iUPbNtasksCol)
  rw1 = .lFirstRw
  rw2 = flEndRow(1)
  FirstUPbRw = rw2 + 2
  LastUPbRw = FirstUPbRw + Nupb - 1
  FirstGenRw = LastUPbRw + 2
  LastGenRw = FirstGenRw + Ngen - 1
  UPbCt = 0: GenCt = 0

  ' Separate the UPb and General Tasks and place copies
  '  in a blank area below the originals
  Set UPbRange = frSr(FirstUPbRw, 1, LastUPbRw, 99)
  Set GenRange = frSr(FirstGenRw, 1, LastGenRw, 99)

  For Rw = rw1 To rw2

    If Cells(Rw, .iTypeCol) = "UPb" Then
      UPbCt = 1 + UPbCt
      frSr(Rw, 1, Rw, 99).Copy UPbRange(UPbCt, 1)
    Else
      GenCt = 1 + GenCt
      frSr(Rw, 1, Rw, 99).Copy GenRange(GenCt, 1)
    End If

  Next Rw

  ' Sort each Task type separately, put in proper rows
  frSr(rw1, 1, rw2, 99).Clear
  UPbRange.Sort Key1:=UPbRange(1, 1), Order1:=xlAscending
  GenRange.Sort Key1:=GenRange(1, 1), Order1:=xlAscending
  UPbRange.Cut (Cells(rw1, 1))
  GenRange.Cut (Cells(rw1 + UPbCt, 1))
End With
End Sub

Sub CreateEmptyTaskCatSheet(TaskCatWorkbook As Workbook)
' Construct an unpopulated Task Catalog sheet

Dim i%, ArN%, Arr As Variant

Arr = Array("TaskName", "File Name", "Creator", "Mineral", "Type", "Description", "# peaks", _
  "Nominal masses", "True masses", "# Equations", "Equation Names", "Equation Spot-Subset Names")

Workbooks.Add
Fonts 1, 1, 999, 99, , , xlLeft, 8, , , , , "Tahoma"
Fonts 2, 1, 6, 99, , , xlLeft, 9
Cells(2, 1) = "# Tasks": Cells(3, 1) = "# UPbTasks"
Cells(4, 1) = "# GeneralTasks"
Cells(2, 5) = "UPb or General"
ArN = UBound(Arr)

For i = 1 To ArN
  Cells(6, i) = Arr(i)
Next i

frSr(6, 1, 6, ArN).Borders(xlEdgeBottom).LineStyle = xlContinuous
ColWidth picAuto, 1, ArN
NoGridlines
Alerts False
Columns.HorizontalAlignment = xlLeft
GetTaskCatRowsCols , True

With puTaskCat
  HA xlCenter, , .iTypeCol
  HA xlCenter, , .iNpksCol
  HA xlCenter, , .iNeqnsCol
  Fonts .lFirstRw - 1, 1, , 99, , , , 8
End With

Cells.NumberFormat = "General"
Fonts 1, 1, , , , , , 15, , , "SQUID2 Task Catalog", , "Tahoma"
ActiveSheet.Name = "TempCat"
NumForTaskcat
Set TaskCatWorkbook = ActiveWorkbook
End Sub

Sub NumForTaskcat(Optional TempCatSht As Boolean = False)
Dim ShtIn As Worksheet, Sht As Worksheet

If TempCatSht Then
  Set Sht = fhTempCat
Else
  Set Sht = ActiveSheet
End If

With puTaskCat
  Sht.Columns(.iNominalCol).NumberFormat = "@"
  Sht.Columns(.iTrueCol).NumberFormat = "@"
  Sht.Columns(.iEqnNaCol).NumberFormat = "@"
  Sht.Columns(.iSubsNaCol).NumberFormat = "@"
End With
End Sub

Sub BuildTaskCatalog(Optional Bad As Boolean = False)
' Construct a completely new Task Catalog from scratch
Dim TaskFileNames$(), TaskNames$()
Dim i%, NumTasks%
Dim TaskCatWbk As Workbook, Arr As Variant

' Create a working Task catalog sheet
CreateEmptyTaskCatSheet TaskCatWbk
Cells.Copy fhTempCat.Cells(1, 1)
'fhTempCat.Cells.NumberFormat = "General"
NumForTaskcat True

NoUpdate
Alerts False
GetTaskFileNames TaskFileNames, TaskNames, NumTasks
puTaskCat.iNumAllTasks = 0

For i = 1 To NumTasks
  StatBar "Adding to Task Catalog:   " & fsS(i) & "          " & TaskFileNames(i)
  AddTaskToTempCatSht TaskFileNames(i), Bad, (i > 1)
Next i

Columns.AutoFit
SortTempCat ' Sort alphabetically for each task-type
Fonts 1, 1, , , , False, xlLeft, 15, , , , , , , , "SQUID2 Task Catalog"
' Copy the working catalog to the internal TempCat sheet
ActiveSheet.Cells.Copy Destination:=fhTempCat.Cells(1, 1)
NumForTaskcat
TaskCatWbk.Close
CopyTempCatShtToTaskCatVar ' Populate the TaskCat variable
CopyTempCatShtToTaskCatWbk ' Save a copy of the new Task Catalog
End Sub

Sub CopyTaskCatNamesToUserSheet()
Dim UPb%, i%, r1&, r2&
GetTaskCatRowsCols

For UPb = True To False
  With puTaskCat
    r1 = IIf(UPb, .lFirstUPbRw, .lFirstGenRW)
    r2 = IIf(UPb, .lLastUPbRw, .lLastGenRW)
  End With
  With frUserTaskList(UPb) ' Refresh Task list in User sheet
    .ClearContents         '  from the TempCat sheet.

    For i = r1 To r2

      If i > 0 Then
        .Cells(i + 1 - r1, 1) = fhTempCat.Cells(i, puTaskCat.iNameCol)
      End If

    Next i

  End With
Next UPb
End Sub

Sub TcatStringToArray(ByVal TcatListVar$, ByVal TcatType%, _
    ByVal TcatTaskNum%, TcatVar, String0Val1)
' 09/04/13 -- Add forcing of numbers as text to a double-prec. number, in case the TaskCat file
'             has this sort of error.
Dim Vtype$, Ttype$, Nlist%, MidLim%, MidNum%
Dim Outp As Variant, v As Variant, Default As Variant

MidLim = UBound(TcatVar, 2)
Default = Choose(1 + String0Val1, "", 0)
Outp = Split(TcatListVar, ",")
' SPLIT returns a zero-based, one-dimensional array containing a
' specified number of substrings.
Nlist = UBound(Outp) + 1

For MidNum = 1 To MidLim

  If MidNum > Nlist Then
    v = Default
  Else
    v = Outp(MidNum - 1)
  End If
  Ttype = TypeName(TcatVar(1, 1, 1))
  Vtype = TypeName(v)
  If String0Val1 = 0 Then v = Trim(v)
  If Vtype = "String" And (Ttype = "Double" Or Ttype = "Integer" _
              Or Ttype = "Long" Or Ttype = "Single") Then
    v = CDbl(Val(v))
  End If
  TcatVar(TcatType, MidNum, TcatTaskNum) = v
Next MidNum

End Sub

Sub DeleteTaskFromTaskcatVar(ByVal TaskNameToDelete$)

Dim Bad As Boolean, Got As Boolean, Exists As Boolean
Dim TaskNa$
Dim Typ%, Ntasks%, TaskNum%, i%, j%, k%, TotTasks%
Dim TaskRow&

TaskNameToDelete = LCase(TaskNameToDelete)
Got = False

With puTaskCat

  For Typ = 0 To 1

    Ntasks = .iaNumTasks(Typ)
    TotTasks = .iNumAllTasks

    For TaskNum = 1 To Ntasks
      TaskNa = LCase(.saNames(Typ, TaskNum))

      If TaskNa = LCase(TaskNameToDelete) Then
        GetFileInfo .saFileNames(Typ, TaskNum), Exists

        If Exists Then
          Alerts False
          ChDirDrv fsSquidUserFolder, Bad
          If Bad Then End
          On Error Resume Next
          Kill (.saFileNames(Typ, TaskNum))
          On Error GoTo 0
          Alerts True
        End If

        Got = True
        .iaNumTasks(Typ) = .iaNumTasks(Typ) - 1
        .iNumAllTasks = .iNumAllTasks - 1
        Ntasks = Ntasks - 1

        For i = TaskNum To Ntasks
          k = i + 1
          .saNames(Typ, i) = .saNames(Typ, k)
          .saFileNames(Typ, i) = .saFileNames(Typ, k)
          .saCreators(Typ, i) = .saCreators(Typ, k)
          .saDescr(Typ, i) = .saDescr(Typ, k)
          .saEqnNaList(Typ, i) = .saEqnNaList(Typ, k)
          .saEqnSubsetNaList(Typ, i) = .saEqnSubsetNaList(Typ, k)
          .saMinerals(Typ, i) = .saMinerals(Typ, k)
          .saNomiMassList(Typ, i) = .saNomiMassList(Typ, k)
          .saTrueMassList(Typ, i) = .saTrueMassList(Typ, k)
          .iaNeqns(Typ, i) = .iaNeqns(Typ, k)
          .iaNpeaks(Typ, i) = .iaNpeaks(Typ, k)

          For j = 1 To peMaxNukes - 1
            .daNominalMass(Typ, j, i) = .daNominalMass(Typ, j, k)
            .daTrueMass(Typ, j, i) = .daTrueMass(Typ, j, k)
          Next j

          .daNominalMass(Typ, peMaxNukes, i) = 0
          .daTrueMass(Typ, peMaxNukes, i) = 0

          For j = 1 To peMaxEqns - 1
            .saEqnNa(Typ, j, i) = .saEqnNa(Typ, j, k)
            .saEqnSubsetNa(Typ, j, i) = .saEqnSubsetNa(Typ, j, k)
          Next j

          .saEqnNa(Typ, peMaxEqns, i) = ""
          .saEqnSubsetNa(Typ, peMaxEqns, i) = ""
        Next i

        Exit For
      End If

    Next TaskNum

    For i = 1 To .iNumAllTasks
      .baTypes(i) = (i <= .iaNumTasks(1))
    Next i

    If Got Then Exit For

  Next Typ

End With

End Sub

Sub RedimTaskCat(LastTaskNum%, TotTaskNum%, AsPreserve As Boolean, Just2D As Boolean)
Dim N%
N = LastTaskNum
With puTaskCat

  If AsPreserve Then
    ReDim Preserve .saNames(0 To 1, 1 To N), .saDescr(0 To 1, 1 To N), .saFileNames(0 To 1, 1 To N)
    ReDim Preserve .baTypes(1 To TotTaskNum), .saMinerals(0 To 1, 1 To N), .saCreators(0 To 1, 1 To N)
    ReDim Preserve .iaNpeaks(0 To 1, 1 To N), .saNomiMassList(0 To 1, 1 To N), .saTrueMassList(0 To 1, 1 To N)
    ReDim Preserve .iaNeqns(0 To 1, 1 To N), .saEqnSubsetNaList(0 To 1, 1 To N), .saEqnNaList(0 To 1, 1 To N)

    If Not Just2D Then
      ReDim Preserve .saEqnNa(0 To 1, 1 To peMaxEqns, 1 To N), .saEqnSubsetNa(0 To 1, 1 To peMaxEqns, 1 To N)
      ReDim Preserve .daNominalMass(0 To 1, 1 To peMaxNukes, 1 To N), .daTrueMass(0 To 1, 1 To peMaxNukes, 1 To N)
    End If

  Else
    ReDim .saNames(0 To 1, 1 To N), .saDescr(0 To 1, 1 To N), .saFileNames(0 To 1, 1 To N)
    ReDim .baTypes(1 To TotTaskNum), .saMinerals(0 To 1, 1 To N), .saCreators(0 To 1, 1 To N)
    ReDim .iaNpeaks(0 To 1, 1 To N), .saNomiMassList(0 To 1, 1 To N), .saTrueMassList(0 To 1, 1 To N)
    ReDim .iaNeqns(0 To 1, 1 To N), .saEqnSubsetNaList(0 To 1, 1 To N), .saEqnNaList(0 To 1, 1 To N)

    If Not Just2D Then
      ReDim .saEqnNa(0 To 1, 1 To peMaxEqns, 1 To N), .saEqnSubsetNa(0 To 1, 1 To peMaxEqns, 1 To N)
      ReDim .daNominalMass(0 To 1, 1 To peMaxNukes, 1 To N), .daTrueMass(0 To 1, 1 To peMaxNukes, 1 To N)
    End If

  End If

End With
End Sub

Function fbLegalTaskFilename(ByVal Na) As Boolean
Dim RevNa$, p%, Le%

Na = LCase(Na)

If Left$(Na, 10) = "squidtask_" Then
  Na = Mid$(Na, 11)

  If Right$(Na, 4) = ".xls" Then
    Le = Len(Na)

    If Le > 5 Then
      RevNa = Mid$(StrReverse(Na), 5)
      p = InStr(RevNa, ".")

      If p > 0 Then
        Na = StrReverse(Mid$(RevNa, p + 1))

        If InStr(Na, ".") = 0 Then
          fbLegalTaskFilename = True
          Exit Function
        End If

      End If

    End If

  End If

End If
fbLegalTaskFilename = False
End Function
