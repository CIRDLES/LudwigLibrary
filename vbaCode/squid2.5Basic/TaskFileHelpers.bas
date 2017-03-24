Attribute VB_Name = "TaskFileHelpers"
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
' Module TaskFileHelpers
' 09/05/06 -- Modify Sub ConstsCheck for the new ConstantsCheck UserForm.
Option Explicit
Option Base 1

Sub GetTaskRows(Optional Bad As Boolean = False)
Dim Changed As Boolean, MissingConstRow As Boolean
Dim Nparams%, ParamNum%, ct%
Dim Rw&, StartRW&, EndRw&
Dim Arr As Variant, Cel As Range

Arr = Array("Created by SQUID version", "First Task row", "Last Task row", "File name", "Type (Geochron, General)", _
     "Task Name", "Description", "Defined by", "Last revision", "Target mineral", "Nuclides", "True masses", "CPS column", _
     "Nominal masses", "Hidden mass", "Ratios", "Ratio Numerator", "Ratio Denominator", "Refmass", "Bkrdmass", "Parent nuclide", _
     "Direct ALT parent-daughter", "Primary U/Th/Pb eqn", "Secondary U/Th/Pb eqn", "Th/U eqn", "Ppm parent eqn", _
     "Equations", "Eqn Names", "Eqn Switch ST", "Eqn Switch SA", "Eqn Switch SC", "Eqn Switch LA", "Eqn Switch NU", _
     "Eqn Switch FO", "Eqn Switch HI", "Eqn Switch ZC", "Eqn Switch AR", "Eqn Switch AR rows", "Eqn Switch AR cols", _
     "Eqn spot-subset name-frags", "Constant names", "Constant values", "AutoGraf1", "AutoGraf2", "AutoGraf3", "AutoGraf4")

Nparams = UBound(Arr)
FindStr "Eqn Switch HI", Rw, , 1, 1, 99, 1

If Rw = 0 Then
  FindStr "Eqn Switch FO", Rw, , 1, 1, 99, 1
  Rows(1 + Rw).Insert
  Cells(1 + Rw, 1) = "Eqn Switch HI"
  Cells(1 + Rw, peTaskNvalsCol) = Cells(Rw, peTaskNvalsCol)
End If

FindStr "Hidden mass", Rw, , 1, 1, 99, 1

If Rw = 0 Then
  FindStr "Nominal masses", Rw, , 1, 1, 99, 1
  Rows(1 + Rw).Insert
  Cells(1 + Rw, 1) = "Hidden mass"
  Cells(1 + Rw, peTaskNvalsCol) = Cells(Rw, peTaskNvalsCol)
End If

FindStr "File name", Rw, , 1, 1, 99, 99
Changed = False

With puTask
  .lFirstTaskRw = Rw
  .lLastTaskRw = flEndRow

  For ParamNum = 13 To 34 Step 21
     ' Is there a "CPS column" or "Eqn Switch ZC" row?
    FindStr Arr(ParamNum), Rw, , .lFirstTaskRw, , .lLastTaskRw, 1

    If Rw = 0 Then ' If not, then create one.
      FindStr Arr(ParamNum - 1), Rw, , .lFirstTaskRw, , .lLastTaskRw, 1
      Rows(1 + Rw).Insert Shift:=xlUp
      .lLastTaskRw = 1 + .lLastTaskRw
      Cells(1 + Rw, 1) = Arr(ParamNum)
      Changed = True
    End If

  Next ParamNum

  Rw = 1: Bad = True

  For ParamNum = 1 To Nparams
    StartRW = IIf(ParamNum > 3, .lFirstTaskRw, 1)
    EndRw = IIf(ParamNum > 3, .lLastTaskRw, 99)
    ct = 0
    FindStr Arr(ParamNum), Rw, , StartRW, , EndRw, 1

    If Rw = 0 And (ParamNum = 2 Or ParamNum = 3) And ct < 1 Then
      Subst Arr(ParamNum), "Task", "run"
      FindStr Arr(ParamNum), Rw, , StartRW, , EndRw, 1
      ct = 1 + ct
    End If

    If Rw > 0 Or ParamNum = 13 Or ParamNum = 36 Then ' the ZC switch and cps cols

      Select Case ParamNum
        Case 1:  .lCreatedByRw = Rw
        Case 2:  .lFirstRowRw = Rw
                 Cells(.lFirstRowRw, peTaskIndxCol) = .lFirstTaskRw
        Case 3:  .lLastRowRw = Rw
                 Cells(.lLastRowRw, peTaskIndxCol) = .lLastTaskRw
        Case 4:  .lFileNameRw = Rw
        Case 5:  .lTypeRw = Rw
        Case 6:  .lNameRw = Rw
        Case 7:  .lDescrRw = Rw
        Case 8:  .lDefByRw = Rw
        Case 9:  .lLastRevRw = Rw
        Case 10: .lMineralRw = Rw
        Case 11: .lNuclidesRw = Rw
        Case 12: .lTrueMassRw = Rw
        Case 13: .lCPScolsRW = Rw
        Case 14: .lNominalMassRw = Rw
        Case 15: .lHiddenMassRw = Rw
        Case 16: .lRatiosRw = Rw
        Case 17: .lRatioNumRw = Rw
        Case 18: .lRatioDenomRw = Rw
        Case 19: .lRefmassRw = Rw
        Case 20: .lBkrdmassRw = Rw
        Case 21: .lParentNuclideRw = Rw
        Case 22: .lDirectAltPDrw = Rw
        Case 23: .lPrimUThPbEqnRw = Rw
        Case 24: .lSecUThPbEqnRw = Rw
        Case 25: .lThUeqnRw = Rw
        Case 26: .lPpmparentEqnRw = Rw
        Case 27: .lEqnsRw = Rw
        Case 28: .lEqnNamesRw = Rw
        Case 29: .lEqnSwSTrw = Rw
        Case 30: .lEqnSwSArw = Rw
        Case 31: .lEqnSwSCrw = Rw
        Case 32: .lEqnSwLArw = Rw
        Case 33: .lEqnSwNUrw = Rw
        Case 34: .lEqnSwFOrw = Rw
        Case 35: .lEqnSwHIrw = Rw
        Case 36: .lEqnSwZCrw = Rw
        Case 37: .lEqnSwARrw = Rw
        Case 38: .lEqnSwARrowsRw = Rw
        Case 39: .lEqnSwARcolsRw = Rw
        Case 40: .lEqnSubsNaFrRw = Rw
        Case 41: .lConstNamesRw = Rw
        Case 42: .lConstValsRw = Rw
        Case 43: .laAutoGrfRw(1) = Rw
        Case 44: .laAutoGrfRw(2) = Rw
        Case 45: .laAutoGrfRw(3) = Rw
        Case 46: .laAutoGrfRw(4) = Rw
      End Select

    ElseIf Arr(ParamNum) <> "Constant names" And Arr(ParamNum) <> "Constant values" _
           And Arr(ParamNum) <> "Hidden eqn" And ParamNum <= Nparams Then
      If ParamNum = 2 Or ParamNum = 3 Then Subst Arr(ParamNum), "run", "Task"
      MsgBox "Can't find the  " & fsInQ(Arr(ParamNum)) & "  row in this Task-file.", , pscSq
      Bad = True: Exit Sub
    End If

  Next ParamNum

  For ParamNum = 5 To peMaxAutochts
    .laAutoGrfRw(ParamNum) = .laAutoGrfRw(4) + ParamNum - 4
    Set Cel = Cells(.laAutoGrfRw(ParamNum), 1)

    If Left$(Cel, 8) <> "AutoGraf" Then
      Cel = "AutoGraf" & fsS(ParamNum)
      Cel.Font.Color = 0
      Changed = True

      If ParamNum = peMaxAutochts Then
        frSr(1 + .laAutoGrfRw(ParamNum), , pemaxrow).Clear
      End If

    End If

  Next ParamNum

  MissingConstRow = (.lConstNamesRw = 0)
  Rw = IIf(MissingConstRow, 1, .lConstNamesRw)

  If MissingConstRow Or Cells(Rw, 1) <> "Constant names" Then
    Rw = .lEqnSubsNaFrRw + 1
    Rows(Rw).Insert
    Cells(Rw, 1) = "Constant names"
    .lConstNamesRw = Rw
    RwIncr
    Changed = True
  End If

  MissingConstRow = (.lConstValsRw = 0)
  Rw = IIf(MissingConstRow, 1, .lConstValsRw)

  If .lConstValsRw = 0 Or Cells(Rw, 1) <> "Constant values" Then
    Rw = .lConstNamesRw + 1
    Rows(Rw).Insert
    Cells(Rw, 1) = "Constant values"
    .lConstValsRw = Rw
    RwIncr
    Changed = True
  End If

End With
If Changed Then ActiveWorkbook.Save
Bad = False
End Sub

Sub RwIncr()
Dim i%
With puTask

  For i = 1 To peMaxAutochts
    .laAutoGrfRw(i) = 1 + .laAutoGrfRw(i)
  Next i

End With
End Sub

Sub LoadOneTask(TaskFileName$, Bad As Boolean, _
  Optional DontCopyToTempSht As Boolean = False, _
  Optional Canceled As Boolean = False, _
  Optional CopyToTaskbook As Boolean = False)
' Opens a task workbook, copes to TempTask worksheet and the puTask variable.
' DOES NOT CLOSE THE TASK WORKBOOK

Dim NoResp As Boolean, ContainsConst As Boolean, Exists As Boolean
Dim Conflict As Boolean, Updated As Boolean
Dim SquidPath$, SquidUserPath$, EqNum%, Dummy%(), Rw&
Dim FsObj As Variant, FileSystemObj As Variant

Alerts False
Bad = False
Canceled = False
piLwrIndx = IIf(pbUPb, -4, 1)

On Error GoTo BadPath
SquidPath = ThisWorkbook.Path
ChDirDrv SquidPath, Bad
If Bad Then Exit Sub
SquidUserPath = fsSquidUserFolder

ChDirDrv SquidUserPath
Set FsObj = CreateObject("Scripting.FileSystemObject")

If LCase(Right$(TaskFileName, 4)) <> ".xls" Then
  TaskFileName = TaskFileName & ".xls"
End If

On Error GoTo BadFile
Set FileSystemObj = FsObj.GetFile(TaskFileName)
On Error GoTo 0

With puTask
  .sDateCreated = FileSystemObj.DateCreated
  .sLastAccessed = FileSystemObj.DateLastAccessed
  .sLastModified = FileSystemObj.DateLastModified
  .sFileName = TaskFileName
End With

OpenWorkbook TaskFileName, Exists
If Not Exists Then GoTo BadFile
Set pwTaskBook = ActiveWorkbook
psTaskFilename = TaskFileName
Columns(1).ColumnWidth = 20
Cells.Interior.ColorIndex = xlNone
Cells.NumberFormat = "General"

For Rw = 1 To flEndRow Step 2
  frSr(Rw, 1, Rw, 3 + peMaxEqns).Interior.Color = 13434828
Next Rw

HA xlLeft, , TaskCol.peTaskValCol

CopyTaskToTaskVars FromTempTask1_OrTaskWbk2:=2, Bad:=Bad
If Bad Then Exit Sub

With puTask

  For EqNum = 1 To .iNeqns
    ConstNumToText .saEqns(EqNum), ContainsConst

    If ContainsConst Then
      ConstsCheck .saEqns(EqNum), EqNum, Conflict, NoResp, _
                   Dummy, , CopyToTaskbook, Updated, Canceled

      If Canceled Then
        If Workbooks.Count > 0 Then
          With ActiveWorkbook
            If .Name <> ThisWorkbook.Name Then .Close
          End With
        End If
        Canceled = True
        Exit Sub
      End If

    End If

  Next EqNum

End With
If Not Updated Then CopyTaskWbkToTempTaskSht False
Exit Sub

BadFile: On Error GoTo 0
MsgBox "Can't find or load the task file (" & TaskFileName & ")", , pscSq
Bad = True: Exit Sub

BadPath: On Error GoTo 0
MsgBox "Can't find the SquidUser path.", , pscSq
Bad = True
End Sub

Sub LoadCompleteTask(ByVal TaskFileName$, Optional Short As Boolean = False, _
  Optional TaskNum% = 0, Optional TaskName$, Optional Bad As Boolean = False, _
  Optional SubsetSpotNaFrags, Optional Nsubsets%, _
  Optional Canceled As Boolean = False)

Dim Na$, RatNum%, PkNum%, NumDenom%, EqNum%, Npks1%, nSubsetEqns%, NomVal#


'To open a workbook file,use the Open method.
'To mark a workbook as saved without writing it to a disk, set its Saved property to True.
'The first time you save a workbook, use the SaveAs method to specify a name for the file.
'If a workbook has never been saved, its Path property returns an empty string ("").
' Some File properties:  DateCreated, DateLastModified
' Task file names example:    SquidTask_ZirconFloatingSlope.KRL.xls

NoUpdate
Canceled = False
LoadOneTask TaskFileName, Bad, , Canceled, False
If Canceled Then Exit Sub

If Bad Then
  MsgBox "Unable to load Task " & fsInQ(TaskFileName) & ".", , pscSq: Exit Sub
End If

With puTask
  .sName = Cells(.lNameRw, peTaskValCol)
  .sFileName = Cells(.lFileNameRw, peTaskValCol)
  'CopyTaskWbkToTempTaskSht
  GetTaskRows
  pbUPb = (Cells(.lTypeRw, peTaskValCol) = "Geochron")
  fhTempTask.Activate

  If pbUPb Then
    ReDim .saEqns(-4 To 0), .saSubsetSpotNa(1 To 1)
    GetTaskUPbEqns .saEqns(-1), .saEqns(-2), .saEqns(-3), .saEqns(-4)
    'For EqNum = piLwrIndx To -1: .saEqns(EqNum) = .saEqns(EqNum): Next
    GetTaskMisc
    pbU = (.iParentIso = 238)
    pbTh = Not pbU: piU1Th2 = 1 - pbTh
    piNumDauPar = 1 - (.bDirectAltPD And .saEqns(-2) <> "")
  End If

  GetTaskNominal
  Npks1 = .iNpeaks
  GetTaskNuclides
  If .iNpeaks <> Npks1 Then GoTo 1
  GetTaskIsotopesAndCPScols
  If .iNpeaks <> Npks1 Then GoTo 1
  GetTaskRatios
  ReDim piaIsoRatsPkOrd(1 To .iNrats, 1 To 2)
  GetTaskNumerDenomIsos

  For RatNum = 1 To .iNrats
    For NumDenom = 1 To 2
      FindPkOrd .daNmDmIso(NumDenom, RatNum), piaIsoRatsPkOrd(RatNum, NumDenom)
    Next NumDenom
  Next RatNum

  GetTaskEqns
  GetTaskSwitches
  GetTaskAutoGrafs

  If pbUPb Then

    For PkNum = 1 To .iNpeaks
      NomVal = .daNominal(PkNum)

      Select Case NomVal
        Case 204: pi204PkOrder = PkNum
        Case 206: pi206PkOrder = PkNum
        Case 207: pi207PkOrder = PkNum
        Case 208: pi208PkOrder = PkNum
        Case 232: pi232PkOrder = PkNum
        Case 238: pi238PkOrder = PkNum
        Case 248: pi248PkOrder = PkNum
        Case 254: pi254PkOrder = PkNum
        Case 264: pi264PkOrder = PkNum
        Case 270: pi270PkOrder = PkNum
      End Select

    Next PkNum

  End If

  For PkNum = 1 To .iNpeaks
    If .daNominal(PkNum) = .dBkrdMass Then piBkrdPkOrder = PkNum
    If Drnd(.daNominal(PkNum), 3) = Drnd(.dRefTrimMass, 3) Then piRefPkOrder = PkNum
  Next PkNum

  If .iNeqns > 0 Or pbUPb Then
    ReDim piaEqCol(-1 To 0, piLwrIndx To .iNeqns), piaEqEcol(-1 To 0, piLwrIndx To .iNeqns)
    ReDim psaEqHdr(piLwrIndx To .iNeqns, 99)
  End If

  ReDim piaIsoRatOrder(1 To .iNrats)

  If fbIM(Nsubsets) Then
    GetTaskSubsets
    prSubsSpotNameFr.ClearContents
    prSubsetEqnNa.ClearContents
    nSubsetEqns = 0

    For EqNum = 1 To .iNeqns
      Na = .saSubsetSpotNa(EqNum)

      If Na <> "" Then
        nSubsetEqns = 1 + nSubsetEqns
        prSubsSpotNameFr(EqNum) = Na
        prSubsetEqnNa(EqNum) = .saEqnNames(EqNum)
      End If

    Next EqNum

  Else
    nSubsetEqns = Nsubsets
  End If

  If Bad Then Exit Sub
  CheckForTemp

  If nSubsetEqns > 0 Then
    If fbIM(TaskName) Then TaskName = Cells(.lNameRw, peTaskValCol)
    foUser("FragTaskType") = TaskName
  End If

  If pbUPb And .iNeqns > 0 Then
    ReDim Preserve piaSwapCols(-4 To .iNeqns)
  End If

End With

On Error Resume Next
pwTaskBook.Close
Exit Sub

CantMakeDirectory: On Error GoTo 0
BadPath: On Error GoTo 0
NoTaskFilesFound: On Error GoTo 0

1:
If fbNIM(Bad) Then
  Bad = True
Else
  MsgBox "Task " & fsInQ(puTask.sName) & "is corrupt.  Please repair or delete.", , pscSq
  End
End If
End Sub

Sub FindCatVarIndx(ByVal TaskName$, ByVal IsUPb As Boolean, TaskTypeIndx%)
Dim s$, t$, j%, ct%, Typ%, nt%, PassCt%
t = Trim(LCase(TaskName))
ct = 0: PassCt = 1: Typ = -IsUPb

Do
  With puTaskCat
    nt = .iaNumTasks(Typ)

    For TaskTypeIndx = 1 To nt
      ct = 1 + ct
      s = Trim(LCase(.saNames(Typ, TaskTypeIndx)))
      If s = t And Typ = -IsUPb Then Exit For
    Next TaskTypeIndx

  End With
If TaskTypeIndx <= nt Then Exit Do

  If PassCt = 1 Then
    BuildTaskCatalog
    PassCt = 1 + PassCt
  Else
    If MsgBox("Can't find task in Task Catalog.  " & _
      "Continue anyway?", vbYesNo, pscSq) = vbNo Then End
    TaskTypeIndx = 1
  End If

Loop

End Sub

Sub CheckForTemp()
If ActiveSheet.Name <> fhTempTask.Name Then fhTempTask.Activate
If puTask.lEqnsRw = 0 Then GetTaskRows
End Sub

Sub GetTaskEqns()
Dim i%, Cm%, Uppr%, EqRow&, NaRow&
CheckForTemp
Cm = peTaskValCol - 1

With puTask
  .iNeqns = Cells(.lEqnsRw, peTaskNvalsCol)

  If .iNeqns > 0 Then
    ReDim Preserve .saEqns(-4 To .iNeqns), .saEqnNames(-4 To .iNeqns)

    For i = 1 To .iNeqns
      .saEqns(i) = Cells(.lEqnsRw, Cm + i)
      .saEqnNames(i) = Cells(.lEqnNamesRw, Cm + i)
    Next i

  End If

End With

End Sub

Sub SaveTaskEqns()
Dim s$, t$, i%, Cm%, N%, EqRow&, NaRow&, tmp&, SwitchRow&(0 To 10)
CheckForTemp
Cm = peTaskValCol - 1

With puTask
  N = .iNeqns
  Cells(.lEqnsRw, peTaskNvalsCol) = N
  Cells(.lEqnNamesRw, peTaskNvalsCol) = N
  Cells(.lEqnSwHIrw, peTaskNvalsCol) = N
  Cells(.lEqnSwSTrw, peTaskNvalsCol) = N
  Cells(.lEqnSwSArw, peTaskNvalsCol) = N
  Cells(.lEqnSwSCrw, peTaskNvalsCol) = N
  Cells(.lEqnSwLArw, peTaskNvalsCol) = N
  Cells(.lEqnSwNUrw, peTaskNvalsCol) = N
  Cells(.lEqnSwFOrw, peTaskNvalsCol) = N
  Cells(.lEqnSwARrw, peTaskNvalsCol) = N
  Cells(.lEqnSwARrowsRw, peTaskNvalsCol) = N
  Cells(.lEqnSwARcolsRw, peTaskNvalsCol) = N
  frSr(.lEqnsRw, peTaskValCol, , Cm + peMaxEqns).ClearContents
  frSr(.lEqnNamesRw, peTaskValCol, , Cm + peMaxEqns).ClearContents
  frSr(.lEqnSwHIrw, peTaskValCol, , Cm + peMaxEqns).ClearContents

  If pbUPb Or N > 0 Then ReDim Preserve .saEqns(-4 To N), .saEqnNames(-4 To N)

  If N > 0 Then
    ReDim Preserve .saSubsetSpotNa(1 To N) '.baHiddenEqn(1 To N),

    For i = 1 To N
      s = IIf(Left$(.saEqns(i), 1) = "'", "'", "")
      Cells(.lEqnsRw, Cm + i).Formula = s & .saEqns(i)
      s = .saEqnNames(i)
      CurlyExtract s, t, , , , , True
      Cells(.lEqnNamesRw, Cm + i).Formula = s
      .saEqnNames(i) = s
      Cells(.lEqnSwHIrw, Cm + i) = .uaSwitches(i).HI
    Next i

  End If

End With
End Sub

Sub GetTaskSwitches()
Dim i%, c%
With puTask
  If .iNeqns = 0 Then Exit Sub
  CheckForTemp
  ReDim Preserve .uaSwitches(1 To .iNeqns)
  piLastN = 0

  For i = 1 To .iNeqns
    With .uaSwitches(i)
      c = peTaskValCol - 1 + i
      .ST = Cells(puTask.lEqnSwSTrw, c)
      .SA = Cells(puTask.lEqnSwSArw, c)
      .SC = Cells(puTask.lEqnSwSCrw, c)
      .LA = Cells(puTask.lEqnSwLArw, c)
      If .LA Then piLastN = 1 + piLastN
      .Nu = Cells(puTask.lEqnSwNUrw, c)
      .FO = Cells(puTask.lEqnSwFOrw, c)
      .HI = Cells(puTask.lEqnSwHIrw, c)
      .Ar = Cells(puTask.lEqnSwARrw, c)
      .ArrNrows = Cells(puTask.lEqnSwARrowsRw, c)
      .ArrNcols = Cells(puTask.lEqnSwARcolsRw, c)
    End With
  Next i

End With
End Sub

Sub SaveTaskEqnSwitches()
Dim i%, j%
With puTask
  If .iNeqns = 0 Then Exit Sub
  CheckForTemp
  frSr(.lEqnSwSTrw, peTaskNvalsCol, .lEqnSwARcolsRw, 99).ClearContents

  For i = 1 To .iNeqns
    j = peTaskValCol - 1 + i
    With .uaSwitches(i)
      Cells(puTask.lEqnSwSTrw, j) = .ST
      Cells(puTask.lEqnSwSArw, j) = .SA
      Cells(puTask.lEqnSwSCrw, j) = .SC
      Cells(puTask.lEqnSwLArw, j) = .LA
      Cells(puTask.lEqnSwNUrw, j) = .Nu
      Cells(puTask.lEqnSwFOrw, j) = .FO
      Cells(puTask.lEqnSwHIrw, j) = .HI
      If puTask.lEqnSwZCrw > 0 Then Cells(puTask.lEqnSwZCrw, j) = .ZC
      Cells(puTask.lEqnSwARrw, j) = .Ar
      Cells(puTask.lEqnSwARrowsRw, j) = .ArrNrows
      Cells(puTask.lEqnSwARcolsRw, j) = .ArrNcols
    End With
  Next i

End With
End Sub

Sub GetTaskAutoGrafs()
Dim Nparams%, i%, ct%, Rw&, tmp As Variant
CheckForTemp
ct = 0

With puTask
  ReDim .uaAutographs(1 To peMaxAutochts)
  If LCase(ActiveSheet.Name) <> "temptask" Then fhTempTask.Activate

  Do
    FindStr "AutoGraf" & fsS(1 + ct), Rw, , .lFirstTaskRw, , .lLastTaskRw, 1
    If Rw = 0 Then Exit Do
    Nparams = Cells(Rw, peTaskNvalsCol)
  If Nparams = 0 Then Exit Do
    ct = 1 + ct
    With .uaAutographs(ct)

      For i = 1 To Nparams
        tmp = Cells(Rw, i + peTaskValCol - 1)

        Select Case i
          Case 1:  .sXname = tmp
          Case 2:  .bAutoscaleX = tmp
          Case 3:  .bZeroXmin = tmp
          Case 4:  .bLogX = tmp
          Case 5:  .sYname = tmp
          Case 6:  .bAutoscaleY = tmp
          Case 7:  .bZeroYmin = tmp
          Case 8:  .bLogY = tmp
          Case 9:  .bRegress = tmp
          Case 10: .bAverage = tmp
        End Select

      Next i

    End With
  Loop

  .iNumAutoCharts = ct
   If ct > 0 Then ReDim Preserve .uaAutographs(1 To ct)
End With

End Sub

Sub SaveTaskAutoGrafs()
Dim i%, ct%, Co%, Rw&
Dim TaskCell As Range
CheckForTemp

With puTask
  frSr(.laAutoGrfRw(1), peTaskIndxCol, 3 + .laAutoGrfRw(peMaxAutochts), _
        3 + picNumAutochtVars).ClearContents

  For ct = 1 To .iNumAutoCharts
    Rw = .laAutoGrfRw(ct) ' + ct - 1
    Cells(Rw, 1) = "AutoGraf" & fsS(ct)
    Cells(Rw, peTaskNvalsCol) = picNumAutochtVars
    Cells(Rw, peTaskIndxCol) = ct

    With .uaAutographs(ct)

      For i = 1 To 10
        Co = peTaskValCol + i - 1
        Set TaskCell = Cells(Rw, Co)

        Select Case i
          Case 1:  TaskCell = .sXname
          Case 2:  TaskCell = .bAutoscaleX
          Case 3:  TaskCell = .bZeroXmin
          Case 4:  TaskCell = .bLogX
          Case 5:  TaskCell = .sYname
          Case 6:  TaskCell = .bAutoscaleY
          Case 7:  TaskCell = .bZeroYmin
          Case 8:  TaskCell = .bLogY
          Case 9:  TaskCell = .bRegress
          Case 10: TaskCell = .bAverage
        End Select

      Next i

    End With

  Next ct

End With

End Sub

Sub GetTaskNumerDenomIsos()
Dim i%, j%, NumRw&, DenomRw&
CheckForTemp
With puTask
  NumRw = .lRatioNumRw
  DenomRw = .lRatioDenomRw
  .iNrats = Cells(NumRw, 3)
  Cells(NumRw, 3) = .iNrats
  Cells(DenomRw, 3) = .iNrats
  ReDim .daNmDmIso(1 To 2, 1 To .iNrats)

  For i = 1 To .iNrats
    j = 3 + i
    .daNmDmIso(1, i) = Cells(NumRw, j)
    .daNmDmIso(2, i) = Cells(DenomRw, j)
  Next i

End With
End Sub

Sub SaveTaskNumerDenomIsos()
Dim i%, j%, NumRw&, DenomRw&
CheckForTemp
With puTask
  NumRw = .lRatioNumRw
  DenomRw = .lRatioDenomRw
  frSr(NumRw, peTaskValCol, , peTaskValCol - 1 + peMaxRats).ClearContents
  frSr(DenomRw, peTaskNvalsCol, , peTaskValCol - 1 + peMaxRats).ClearContents
  Cells(NumRw, peTaskNvalsCol) = .iNrats
  Cells(DenomRw, peTaskNvalsCol) = .iNrats

  For i = 1 To .iNrats
    j = peTaskValCol + i
    Cells(NumRw, j - 1) = .daNmDmIso(1, i)
    Cells(DenomRw, j - 1) = .daNmDmIso(2, i)
  Next i

End With
End Sub

Sub GetTaskNominal()
Dim i%
CheckForTemp
With puTask
  .iNpeaks = Cells(.lNominalMassRw, peTaskNvalsCol)
  ReDim .daNominal(1 To .iNpeaks)

  For i = 1 To .iNpeaks
    .daNominal(i) = Cells(.lNominalMassRw, peTaskValCol - 1 + i)
  Next i

End With
End Sub

Sub SaveTaskNominal()
Dim i%, Cm%
CheckForTemp
With puTask
  Cm = peTaskValCol - 1
  Cells(.lNominalMassRw, peTaskNvalsCol).Formula = .iNpeaks
  frSr(.lNominalMassRw, peTaskValCol, , Cm + peMaxNukes).ClearContents

  For i = 1 To .iNpeaks
    Cells(.lNominalMassRw, Cm + i).Formula = .daNominal(i)
  Next i

End With
End Sub

Sub GetTaskSubsets()
Dim i%, TmpN%
CheckForTemp
With puTask
  TmpN = Cells(.lEqnsRw, peTaskNvalsCol)
  If TmpN = 0 Then Exit Sub
  ReDim .saSubsetSpotNa(1 To TmpN)

  For i = 1 To TmpN
    .saSubsetSpotNa(i) = Cells(.lEqnSubsNaFrRw, peTaskValCol + i - 1)
  Next i

End With
End Sub

Sub SaveTaskSubsets()
Dim s$, i%, Cm%, ct%
With puTask
  If .iNeqns = 0 Then Exit Sub
  CheckForTemp
  Cm = peTaskValCol - 1: ct = 0
  With puTask
    Cells(.lEqnSubsNaFrRw, peTaskNvalsCol).Formula = .iNeqns
    frSr(.lEqnSubsNaFrRw, peTaskValCol, , Cm + peMaxEqns).ClearContents

    For i = 1 To .iNeqns
      s = Trim(.saSubsetSpotNa(i))

      If s <> "" Then
        ct = 1 + ct
        Cells(.lEqnSubsNaFrRw, Cm + i) = s
      End If

    Next i

    Cells(.lEqnSubsNaFrRw, peTaskNvalsCol) = ct
  End With
End With
End Sub

Sub GetTaskIsotopesAndCPScols()
' 09/04/10 -- Add code for hidden CPS columns
Dim i%, j%

CheckForTemp
With puTask
  .iNpeaks = Cells(.lTrueMassRw, 3)
  ReDim .daTrueMass(1 To .iNpeaks), .baCPScol(1 To .iNpeaks), .baHiddenMass(1 To .iNpeaks)

  For i = 1 To .iNpeaks
    j = 3 + i
    .daTrueMass(i) = Cells(.lTrueMassRw, j)
    If .lCPScolsRW > 0 Then .baCPScol(i) = Cells(.lCPScolsRW, j) Else .baCPScol(i) = False
    If .lHiddenMassRw > 0 Then .baHiddenMass(i) = Cells(.lHiddenMassRw, j) Else .baHiddenMass(i) = False
  Next i

End With
End Sub

Sub SaveTaskIsotopesAndCPScols()
' 09/04/10 -- Add code for hidden CPS columns
Dim i%, Cm%

CheckForTemp
With puTask
  Cm = peTaskValCol - 1
  Cells(.lTrueMassRw, peTaskNvalsCol).Formula = .iNpeaks
  frSr(.lTrueMassRw, peTaskValCol, , Cm + peMaxNukes).ClearContents

  For i = 1 To .iNpeaks
    Cells(.lTrueMassRw, Cm + i).Formula = .daTrueMass(i)

    If .lHiddenMassRw > 0 Then
      Cells(.lHiddenMassRw, Cm + i).Formula = .baHiddenMass(i)
    End If
    If .lCPScolsRW > 0 Then
      Cells(.lCPScolsRW, Cm + i).Formula = .baCPScol(i)
    End If

  Next i

End With
End Sub

Sub GetTaskRatios()
Dim i%, Rw&
CheckForTemp
With puTask
  Rw = .lRatiosRw
  .iNrats = Cells(Rw, peTaskNvalsCol)
  ReDim .saIsoRats(1 To .iNrats)

  For i = 1 To .iNrats
    .saIsoRats(i) = Cells(Rw, peTaskValCol - 1 + i)
  Next i

End With
End Sub

Sub SaveTaskRatios()
Dim i%, Cm%, Rw&
' 09/06/09 -- added .NumberFormat = "@"
CheckForTemp
With puTask
  Cm = peTaskValCol - 1
  Rw = .lRatiosRw
  Cells(Rw, peTaskNvalsCol).Formula = .iNrats
  frSr(Rw, peTaskValCol, , Cm + peMaxRats).ClearContents

  For i = 1 To .iNrats
    With Cells(Rw, Cm + i)
      .NumberFormat = "@"
      .Formula = puTask.saIsoRats(i)
    End With
  Next i

End With
End Sub

Sub GetTaskNuclides()
Dim i%, Rw&, tmp1%, tmp2$
CheckForTemp
With puTask
  Rw = .lNuclidesRw
  .iNpeaks = Cells(Rw, peTaskNvalsCol)
  ReDim .saNuclides(1 To .iNpeaks)

  For i = 1 To .iNpeaks
    .saNuclides(i) = Cells(Rw, peTaskValCol - 1 + i)
  Next i

End With
End Sub

Sub SaveTaskNuclides()
Dim i%, Cm%, Rw&
CheckForTemp
With puTask
  Cm = peTaskValCol - 1
  Rw = .lNuclidesRw
  Cells(Rw, peTaskNvalsCol).Formula = .iNpeaks
  frSr(Rw, peTaskValCol, , Cm + peMaxNukes).ClearContents

  For i = 1 To .iNpeaks
    Cells(Rw, Cm + i).Formula = .saNuclides(i)
  Next i

End With
End Sub

Sub GetTaskMisc(Optional RefMass, Optional BkrdMass, Optional ParentIso, _
  Optional DirectAltPD As Boolean)
Dim i%, Rw&, tmp As Variant
CheckForTemp

With puTask
  tmp = Cells(.lRefmassRw, peTaskValCol)
  If fbIM(RefMass) Then .dRefTrimMass = tmp Else RefMass = tmp
  tmp = Cells(.lBkrdmassRw, peTaskValCol)

  If tmp <> "" Then
    .dBkrdMass = tmp
    If fbIM(BkrdMass) Then .dBkrdMass = tmp Else BkrdMass = tmp
  End If

  tmp = Cells(.lParentNuclideRw, peTaskValCol)

  If tmp <> "" Then
    .iParentIso = tmp
    If fbIM(ParentIso) Then piParentIso = tmp Else ParentIso = tmp
  End If

  tmp = Cells(.lDirectAltPDrw, peTaskValCol)

  If tmp <> "" Then
    .bDirectAltPD = tmp
    If fbIM(DirectAltPD) Then .bDirectAltPD = tmp Else DirectAltPD = tmp
  End If

End With

End Sub

Sub SaveTaskMisc(Optional ByVal RefMass, Optional BkrdMass, Optional ParentIso, _
  Optional DirectAltPD)
CheckForTemp
With puTask
  VIM RefMass, .dRefTrimMass
  VIM BkrdMass, .dBkrdMass
  VIM ParentIso, .iParentIso
  VIM DirectAltPD, .bDirectAltPD
  Cells(.lRefmassRw, peTaskValCol).Formula = IIf(RefMass > 0, RefMass, "")
  Cells(.lBkrdmassRw, peTaskValCol).Formula = IIf(BkrdMass > 0, BkrdMass, "")
  Cells(.lParentNuclideRw, peTaskValCol).Formula = IIf(ParentIso > 0, ParentIso, "")
  Cells(.lDirectAltPDrw, peTaskValCol).Formula = DirectAltPD
End With
End Sub

Sub GetTaskUPbEqns(PrimaryUThPbEqn$, SecondaryUThPbEqn$, ThUeqn$, PpmParentEqn$)
Dim i%, c%, Rw&
CheckForTemp
With puTask
  c = peTaskValCol
  PrimaryUThPbEqn = Cells(.lPrimUThPbEqnRw, c)
  SecondaryUThPbEqn = Cells(.lSecUThPbEqnRw, c)
  ThUeqn = Cells(.lThUeqnRw, c)
  PpmParentEqn = Cells(.lPpmparentEqnRw, c)
End With
End Sub

Sub SaveTaskUPbEqns(PrimaryUThPbEqn$, SecondaryUThPbEqn$, ThUeqn$, PpmParentEqn$)
Dim i%, c%, Rw&
CheckForTemp
With puTask
  c = peTaskValCol
  Cells(.lPrimUThPbEqnRw, c).Formula = PrimaryUThPbEqn
  Cells(.lSecUThPbEqnRw, c).Formula = SecondaryUThPbEqn
  Cells(.lThUeqnRw, c).Formula = ThUeqn
  Cells(.lPpmparentEqnRw, c).Formula = PpmParentEqn
  On Error Resume Next
  ReDim Preserve .saEqns(-4 To .iNeqns)
  On Error GoTo 0
End With

End Sub

Sub CheckWholeTask(BadTask As Boolean)
Dim BadRats As Boolean, BadEqns As Boolean, Msg$, Spa$
BadTask = True
CheckRatios BadRats

If BadRats Then
  Msg = "Isotope ratios for this Task are not consistent with its Run Table."
Else
  CheckEqns BadEqns

  If BadEqns Then
    Msg = "Equations defined for this Task refer to undefined isotope ratios."
  Else
    BadTask = False
  End If

End If

If BadTask Then
  Spa = pscLF2 & String(Len(Msg) / 2 - 7, " ")
  MsgBox Msg & Spa & "Task not saved.", , pscSq
End If
End Sub

Sub CheckRatios(BadTask As Boolean)
Dim GotNumer As Boolean, NuclideExists As Boolean, GotDenom As Boolean
Dim Iso$, Rat$, NumDenom%, IsoNum%, RatNum%, p%, IsoVal#

BadTask = True
With puTask

  For RatNum = 1 To .iNrats
    Rat = .saIsoRats(RatNum)
    p = InStr(Rat, "/")
    If p = 0 Then Exit Sub

    For NumDenom = 1 To 2
      Iso = Choose(NumDenom, Left$(Rat, p - 1), Mid$(Rat, p + 1))
      If Not IsNumeric(Iso) Then Exit Sub
      IsoVal = Val(Iso)

      For IsoNum = 1 To .iNpeaks
        If IsoVal = .daNominal(IsoNum) Then Exit For
      Next IsoNum

      If IsoNum > .iNpeaks Then Exit Sub
    Next NumDenom

  Next RatNum

End With
BadTask = False
End Sub

Sub CheckEqns(BadTask As Boolean)
Dim OkRat As Boolean, Done As Boolean
Dim Eqn$, EqFrag$, EqNum%, SafeCount%
Dim Indx%, ApparentRefType%, RatNum%, IndxType%
BadTask = True: SafeCount = 0

With puTask

  For EqNum = 1 To .iNeqns
    Eqn = .saEqns(EqNum)

    If Eqn <> "" Then
      OkRat = True

      Do
        ExtractEqnRef Phrase:=Eqn, IndxStr:=EqFrag, IndxNum:=Indx, IndxType:=IndxType, _
                       RefType:=ApparentRefType

        If IndxType = peRatio And ApparentRefType = 2 Then

          For RatNum = 1 To .iNrats
            If .saIsoRats(RatNum) = EqFrag Then Exit For
          Next RatNum

          If RatNum > .iNrats Then OkRat = False
        End If

        Subst Eqn, EqFrag
        Subst EqFrag, Chr(34)
        Done = (EqFrag = "" And Indx = 0 And ApparentRefType = 0)
        SafeCount = 1 + SafeCount
      Loop Until Done Or SafeCount > 99

    End If

  Next EqNum

End With

BadTask = False
End Sub

Function fbGotSolver()
Dim SolverFound As Boolean, AddinName$, i%

For i = 1 To AddIns.Count
  AddinName = LCase(AddIns(i).Name)
  If AddinName = "solver.xla" Then SolverFound = True: Exit For
Next i

If SolverFound Then
  If Not AddIns(i).Installed Then AddIns(i).Installed = True
  fbGotSolver = True
Else
  fbGotSolver = False
End If

End Function

Sub CheckForSolver()
Dim i%, AlreadyChecked As Boolean
With puTask
  If .iNeqns > 0 Then ReDim .baSolverCall(1 To .iNeqns)
  AlreadyChecked = False

  For i = 1 To .iNeqns

    If InStr(LCase(.saEqnNames(i)), "<<solve>>") > 0 Then
      .baSolverCall(i) = True

      If Not AlreadyChecked Then

        If Not fbGotSolver Then
          MsgBox "Unable to find and install the SOLVER add-in" & _
            "required by the current Task." & pscLF2 & _
            "Please install SOLVER and try again.", , pscSq
          CrashEnd
        Else
          AlreadyChecked = True
        End If

      End If

    End If

  Next i

End With
End Sub

Sub GetTaskFileNames(TaskFileList$(), TaskNameList$(), NumTasks%)
' 09/03/25 -- Note drive in and Squid2.xla drive as well as folder in & SquidUserFolder.
'             Restore folder AND drive when done.
Dim Exists As Boolean, Bad As Boolean
Dim Msg$, FolderIn$, DriveIn$, SquidDrive$, SquidUserFolder$
Dim FileNa$, LcFileNa$, TaskNa$, Rna$, FileList$()
Dim p%, Nfiles%

SquidDrive = fsSquidDrive
SquidUserFolder = fsSquidUserFolder
FolderIn = CurDir
DriveIn = fsCurDrive
ChDrive SquidDrive
ChDir fsSquidUserFolder
ct = 0: NumTasks = 0

FileNamesInDir SquidUserFolder, FileList, Nfiles, 1
If Nfiles = 0 Then Msg = "No files present in directory " & CurDir: GoTo Bad

ReDim TaskFileList(1 To Nfiles), TaskNameList(1 To Nfiles)

For ct = 1 To Nfiles
  FileNa = FileList(ct)

  If fbLegalTaskFilename(FileNa) Then
1   TaskNa = Mid$(FileNa, 11)
    If TaskNa = "" Then Msg = "Null taskname in line 1, GetTaskFileNames": GoTo Bad
2   TaskNa = Left$(TaskNa, Len(TaskNa) - 4)
    If TaskNa = "" Then Msg = "Null taskname in line 2, GetTaskFileNames": GoTo Bad
3   Rna = StrReverse(TaskNa)
4   p = InStr(Rna, ".")
    If p = 0 Then Msg = "No period in string Rna, line 4, GetTaskFileNames": GoTo Bad
5   TaskNa = StrReverse(Mid$(Rna, p + 1))
6   NumTasks = 1 + NumTasks
7   TaskFileList(NumTasks) = FileNa
8   TaskNameList(NumTasks) = TaskNa
  End If

Next ct

If NumTasks = 0 Then Msg = "No task files present in directory " & CurDir: GoTo Bad

ReDim Preserve TaskFileList(1 To NumTasks), TaskNameList(1 To NumTasks)
'ChDrive DriveIn
ChDirDrv FolderIn
Exit Sub

Bad: If Msg <> "" Then MsgBox Msg, , pscSq
Msg = "Names in " & CurDir & " are:"

For ct = 1 To Nfiles
  Msg = Msg & vbLf & FileList(ct)
Next ct

MsgBox Msg, , pscSq
'ChDrive DriveIn
ChDirDrv FolderIn
End
End Sub

Sub ConstsCheck(ByVal Eqn$, ByVal EqNum%, Conflict As Boolean, _
  NoResponse As Boolean, Optional ConstsList, Optional NumEqnConsts%, _
  Optional UpdateTaskWbk As Boolean = False, _
  Optional UpdatedWbk As Boolean = False, _
  Optional Canceled As Boolean = False)
' 09/05/06 -- Modify for the new ConstantsConflict UserForm
' 09/06/10 -- Make Constlist an Optional Variant param
' 10/04/27 -- Correctly place the constant names in the ConstantsConflict user form.
Dim IsNum As Boolean
Dim Msg$, s$, s1$, s2$, s3$, Extr$, lcExtr$, TaskVal$, PrefsVal$, Query1$, Query2$
Dim TaskConstName$, FinalConstName$, PrefsConstName$, Eqn0$
Dim i%, j%, p%, Indx%, IndxType%, PrefsConstNum%, TaskConstNum%, LbrakPos%, Lct%, Nprefs%
Dim PrefsToTask1_TaskToPrefs2%
Dim Resp&, tmp#, v As Variant

On Error GoTo 0
Canceled = False: Lct = 0
If fbNIM(NumEqnConsts) Then NumEqnConsts = 0
If Trim(Eqn) = "" Then Exit Sub
Eqn0 = Eqn
UpdatedWbk = False
Eqn = Eqn0

For i = 1 To peMaxConsts
  If prConstNames(i) = "" Then Nprefs = i - 1: Exit For
Next i

Conflict = False: NoResponse = False
' (1) If name is in both Prefs and puTask, check values: if different, query user.
' (2) If name is not in Prefs list but is in Task list, add name & value to Prefs.
' (3) If name is in Prefs list but not in task list, add name and Prefs value to puTask.
' (4) If Const name not in Prefs list and not in Task list, ask user for value,
'     then add to both prefs list and puTask.
Subst Eqn, "<=>"

Do
  PrefsConstNum = 0: TaskConstNum = 0
  Lct = 1 + Lct
  ExtractEqnRef Eqn, Extr, Indx, IndxType

  If Extr = "" Or InStr(Eqn, "<") = 0 Then Exit Sub

  With puTask
    If IndxType = peUndefinedConstant Then
      Msg = "Do you wish to define a new Task Constant named " & fsInQ(Extr) & "?"

      Select Case MsgBox(Msg, vbYesNoCancel, pscSq)
        Case vbYes

          Do
            v = InputBox("Value of new constant?", pscSq)
            If v = "" Then Canceled = True: Exit Sub
          Loop Until fbIsAllNumChars(v)

          Nprefs = 1 + Nprefs
          prConstNames(Nprefs) = "_" & fsLegalName(Extr, True)
          prConstValues(Nprefs) = v
          prConstsRange.Sort prConstNames(1)
          pbChangedEquations = True
          Subst Eqn, "<" & Extr & ">"

        Case vbNo, vbCancel
          Canceled = True
          Exit Sub

      End Select

    ElseIf Indx < -1000 Then ' constant

      If IndxType = pePrefsConstant Or IndxType = peBothConstant Then
        PrefsConstNum = -IIf(Indx < -3000, 3000, 1000) - Indx
        PrefsConstName = prConstNames(PrefsConstNum)
        FinalConstName = PrefsConstName

        If IndxType = peBothConstant Then  ' determine the Task Constant number
          s1 = fs_(prConstNames(PrefsConstNum), True)

          For i = 1 To puTask.iNconsts
            s2 = fs_(puTask.saConstNames(i), True)
            If s1 = s2 Then TaskConstNum = i: Exit For
          Next i

        End If

      ElseIf IndxType = peTaskConstant Then
        TaskConstNum = -2000 - Indx
        TaskConstName = .saConstNames(TaskConstNum)
      End If

      If PrefsConstNum > 0 Then
        PrefsVal = prConstValues(PrefsConstNum)
        If TaskConstNum > 0 Then TaskVal = .vaConstValues(TaskConstNum)

        If TaskConstNum = 0 Then
        ' Const in Prefs but not in Task. Copy the prefs-const name & value to the Task
          ChangeTaskConstant 0, PrefsConstNum, 1, UpdateTaskWbk, UpdatedWbk
          Subst Eqn, "<" & Extr & ">"
          FinalConstName = PrefsConstName

        ElseIf TaskVal <> PrefsVal Then ' A conflict -- must resolve
          If Not (pbPrefsEqualsTask Or pbTaskEqualsPrefs Or pbDefineNewConst) Then
            Load ConstantsConflict
            Query1 = "The Task constant " & s1 & " (= " & TaskVal & ") in Equation" & StR(EqNum) & _
              " conflicts with" & vbLf & "a constant having the same name in the Preferences sheet (= " _
              & PrefsVal & ")" & "."
            With ConstantsConflict
              .lQuery1.Caption = Query1
              ' 10/04/27 -- next 3 lines added.
              .opPrefsEqualsTask.Caption = "Set the value of the Preferences " & s1 & " to that of the Task " & s1
              .opTaskEqualsPrefs.Caption = "Set the value of the Task " & s1 & " to that of the Preferences " & s1
              .opDefineNew.Caption = "Define a new value of " & s1 & " for both the Preferences and Task constants"
              .Show
            End With
            If FormRes.peCancel Then Canceled = True: Exit Sub
          End If

          Conflict = True

          If pbPrefsEqualsTask Or pbTaskEqualsPrefs Then
            PrefsToTask1_TaskToPrefs2 = IIf(pbPrefsEqualsTask, 2, 1)
            ChangeTaskConstant TaskConstNum, PrefsConstNum, PrefsToTask1_TaskToPrefs2, _
                               UpdateTaskWbk, UpdatedWbk
            FinalConstName = Choose(PrefsToTask1_TaskToPrefs2, TaskConstName, PrefsConstName)

          ElseIf pbDefineNewConst Then
            ' Change Prefs const value to something else
            s = Mid(PrefsConstName, 2)
            Msg = "Please rename and or redefine the value of " & s & " as necessary." & pscLF2 _
             & "The new value for " & s & " will be written to both the Task " & _
             "and Preferences workbooks."
            Resp = MsgBox(Msg, vbOKCancel, pscSq)
            If Resp = vbCancel Then Canceled = True: Exit Sub
            pbActiveEquation = EqNum
            pbCanAppendConstant = False
            Constants.Show vbModal
            ChangeTaskConstant TaskConstNum, PrefsConstNum, 1, UpdateTaskWbk, UpdatedWbk
            FinalConstName = PrefsConstName
          End If

        End If

        pbChangedEquations = True
        Subst Eqn, "<" & Extr & ">"

      ElseIf PrefsConstNum = 0 And TaskConstNum > 0 Then
        ' Const in Task but not in Prefs. Copy the Task-const name & value to Prefs.
        Nprefs = 1 + Nprefs
        prConstNames(Nprefs) = TaskConstName
        prConstValues(Nprefs) = .vaConstValues(TaskConstNum)

      ElseIf PrefsConstNum = 0 And TaskConstNum = 0 Then
        ' No such constant name exists.
        TaskVal = InputBox("What value do you wish to assign to the undefined constant " _
            & fsInQ(Extr) & "?", pscSq)
        If TaskVal = "" Then NoResponse = True: Conflict = True: Exit Sub
        .iNconsts = 1 + .iNconsts
        Nprefs = 1 + Nprefs
        .saConstNames(.iNconsts) = TaskConstName
        .vaConstValues(.iNconsts) = TaskVal
        prConstNames(Nprefs) = "_" & TaskConstName
        prConstValues(Nprefs) = TaskVal
      End If ' If PrefsConstNum > 0

      If fbNIM(ConstsList) Then
        NumEqnConsts = 1 + NumEqnConsts
        ReDim Preserve ConstsList(1 To NumEqnConsts)
        FinalConstName = fs_(FinalConstName, True)
        i = 1

        Do While fs_(prConstNames(i), True) <> FinalConstName
          i = i + 1
        Loop

        ConstsList(NumEqnConsts) = i
      End If

    Else
      Subst Eqn, psBrQL & Extr & psBrQR
      Subst Eqn, "[" & Extr & "]"
    End If ' If Indx < -1000

  End With

Loop Until Eqn = "" Or (Indx = 0 And Extr = "") Or _
     (InStr(Eqn, "[") = 0 And InStr(Eqn, "<") = 0) Or Lct = 99

If Lct = 99 Then
  MsgBox "Loop-locked in Sub ConstsCheck with equation" & vbLf & _
    Eqn0 & pscLF2 & "Please inform Ken Ludwig", , pscSq
  End
End If

End Sub

Sub ConstNumToText(Eqn$, Optional FoundConst As Boolean = False)
Dim Eq$, TxtStr$, BrkStr$, BrkCt%, BctL%, BctR%, BstrL%(), BstrR%()

Eq = Eqn
AllInstanceLoc "<", Eq, BstrL, BctL
AllInstanceLoc ">", Eq, BstrR, BctR

If BctL = BctR Then

  For BrkCt = 1 To BctL

    If BrkCt > 1 Then
      AllInstanceLoc "<", Eq, BstrL, BctL
      AllInstanceLoc ">", Eq, BstrR, BctR
    End If

    If BstrR(BrkCt) > BstrL(BrkCt) Then
      BrkStr = Mid$(Eq, BstrL(BrkCt) + 1, BstrR(BrkCt) - BstrL(BrkCt) - 1)
      FoundConst = True

      If fbIsAllNumChars(BrkStr) Then
        TxtStr = fs_(prConstNames(Val(BrkStr)))
        Eq = Left$(Eq, BstrL(BrkCt)) & TxtStr & Mid$(Eq, BstrR(BrkCt))
      End If

    End If

  Next BrkCt

End If

Eqn = Eq
End Sub

Sub ChangeTaskConstant(ByVal TaskConstNum, ByVal PrefsConstNum, _
  ByVal PrefsToTask1_TaskToPrefs2%, Optional UpdateTaskWbk As Boolean = False, _
  Optional UpdatedWbk As Boolean = False)
Dim i%, p%, t%

p = PrefsConstNum
t = TaskConstNum
With puTask

  Select Case PrefsToTask1_TaskToPrefs2
    Case 1

      If t = 0 Then
        .iNconsts = 1 + .iNconsts
        t = .iNconsts
        ReDim Preserve .saConstNames(1 To t), .vaConstValues(1 To t)
      End If

      .saConstNames(t) = prConstNames(p)
      .vaConstValues(t) = prConstValues(p)
    Case 2
      If p = 0 Then p = 1
      prConstValues(p) = .vaConstValues(t)
    Case Else
      MsgBox "Coding error: PrefsToTask1_TaskToPrefs2=" & _
              StR(PrefsToTask1_TaskToPrefs2) & " in ChangeTaskCOnstant.", _
              vbOKOnly, pscSq
      End
  End Select

End With

If UpdateTaskWbk Then
  CopyTaskVarsToSheet 2
  CopyTaskVarsToSheet 1
  UpdatedWbk = True
End If
End Sub

Sub AttachWorksheetInFileToOpenWorkbook(NameOfSourceWorkbook$, BadLoad As Boolean, _
  Optional NameOfSourceWorksheet = "", Optional DestinationWorkBookNa, _
  Optional DeleteShapes As Boolean = False, Optional TaskSht As Boolean = False)

Dim Bad As Boolean, Exists As Boolean, LcNa$, DestNa$, Col%, Rw&
Dim SourceSht As Worksheet, SourceWbk As Workbook
Dim DestSht As Worksheet, ShtIn As Worksheet, WbkIn As Workbook, Shp As Shapes

Set WbkIn = ActiveWorkbook
Set ShtIn = ActiveSheet
NoUpdate

If fbIM(DestinationWorkBookNa) Then
  DestNa = ActiveWorkbook.Name
Else
  DestNa = DestinationWorkBookNa
End If

BadLoad = False

OpenWorkbook NameOfSourceWorkbook, Exists
If Not Exists Then GoTo BadFile

Set SourceWbk = ActiveWorkbook
With ActiveWorkbook

  If NameOfSourceWorksheet = "" Then
    Set SourceSht = .Sheets(1)
    NameOfSourceWorksheet = SourceSht.Name
  Else
    Set SourceSht = .Sheets(NameOfSourceWorksheet)
  End If

End With

On Error Resume Next
Workbooks(DestNa).Sheets("task").Delete
On Error GoTo 0

If DeleteShapes Then
  On Error Resume Next

  For Each Shp In SourceSht.Shapes
    If Left$(Shp.Name, 7) <> "Comment" Then Shp.Delete
  Next Shp

  ' otherwise end up with replicate shapes
  On Error GoTo 0
End If

If TaskSht Then
  Cells.Interior.ColorIndex = xlNone
  Col = 3 + peMaxEqns

  For Rw = 1 To flEndRow Step 2
    frSr(Rw, 1, Rw, Col).Interior.Color = 13434828
  Next Rw

End If

SourceSht.Copy Before:=Workbooks(DestNa).Sheets(1)
SourceWbk.Close
WbkIn.Activate
ShtIn.Activate
Exit Sub

BadFile:
On Error GoTo 0
BadLoad = True
End Sub
