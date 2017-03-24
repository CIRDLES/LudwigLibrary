Attribute VB_Name = "Main"
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
' 09/03/02 -- Start specifically declaring lower bound of all arrays, rather than
'             relying on Option Base 1 (for portability)
' 09/02/25 -- Modify Sub ChDirDrv to always specifically change to the drive indicated by Path$
' 09/04/07 -- Finished delimiting all For-Next, If-Then & Do-Loop constructs with linefeeds in
'             all modules.
' 09/04/07 -- Insert 3rd char "a" into all Public and Module-level array variables.
' 09/05/09 -- Add the poAppObject As New AppClass variable, & InitWbkOpenHandler & WBopen subs
'             to handle the WorkbookOpen event.
' 09/05/11 -- Rewrite Sub Reprocess, modify UPbGeochronStartup, SquidGeochron, and
'             GenIsoRunStartup accordingly.  Add the FormRes.peLoadNewFile variable.
' 09/05/11 -- Add the explicit "Order1:=xlAscending" to all otherwise unspecified .SORT calls.
' 09/06/10 -- Added "peMinDriftcorrNumSpots = 8"
' 09/06/19 -- Removed the piDefSKageType variable from code, removed S-K age-type choices from
'             Group panel, from Sub Group,and from the common-Pb panel of Preferences.  The index
'             isotope for common-Pb correction is now forced to be the same as that in the Geochron
'             Setup panel, as is the S-K age-type to calculate the common Pb.
' 09/07/02 -- Rename most U/Pb* column# variables and associated column-index names for mutual consistency.

Option Explicit
Option Base 1

' SQUID Data reduction program for raw SHRIMP zircon analyses.  K.R. Ludwig, BGC.

' References isoplot3.xla WITH A PROJECT NAME OF ISOPLOT3!

' p    Public            First char
' m    Module
' s    Sub or Function

' i    Integer           Second Char
' l    Long
' n    Single
' d    Double
' e    Enumerated
' a    Date
' s    String
' v    Variant
' o    Object
' b    Boolean
' r    Range
' h    Worksheet
' t    Control
' z    Controls
' w    Workbook
' c    Chart
' u    User defined
' wi   Window

' a Array             Third char (optional)
' f Function
' c Const
' e Enum

' PUBLIC CONSTANTS

Public Const pscIsoplotAddinName$ = "Isoplot3nx.xla"
Public Const pscSquidToolbarName = "SQUID 2.5"

' Strings
Public Const pscStdShtNa = "StandardData", pscStdShtNot$ = "StandardData!" ' was "'StandardData'!"
Public Const pscSamShtNa = "SampleData"
Public Const pscUra = "137.88", pscPcnt = "0.00$%$", pscDd1 = ".0"
Public Const pscDd2 = ".00", pscDd3 = ".000", pscDd4 = ".0000", pscDd5 = ".00000"
Public Const pscZq = "0", pscZd1 = "0.0", pscZd2 = "0.00"
Public Const pscZd3 = "0.000", pscPpe = "%|err", pscErF = "[>=1000]0E+0;[>=10]0;0.00"
Public Const pscFluff = "For visualization & preliminary evaluation only"
Public Const pscRejectedSpotNums = "rejected spot #s"
Public Const pscCpbC = "Common-Pb correction by assuming "
Public Const pscGen = "General", pscNoCol = "~!@#$"

Public Const pscLF2 = vbLf & vbLf, pscQ = """", pscDot = "�", pscPermil = "�"
Public Const pscDblLeftArrow = "�", pscOneHalf = "�", pscMuchGreaterThan = "�"
Public Const pscMult = "�", pscDiv = "�", pscEndash = "�", pscEmdash = "�"
Public Const pscPm = "�", pscPmm = " � ", pscSUM = "�", pscSqrt = "�"
Public Const pscIdentity = "�", pscLeftArrow = "�", pscRightArrow = "�"
Public Const pscWingdingsLeftArrow = "�", pscWingdingsRightArrow = "�"
Public Const pscWingdings2Check = "P"

Public Const pscR6875 = "206Pb/238U-207Pb/235U age-concordance"
Public Const pscR6882 = "206Pb/238U-208Pb/232Th age-concordance"
Public Const pscLm8 = "0.000155125", pscLm2 = "0.000049475", pscLm5 = ".00098485"
Public Const pscEx8 = "Exp(0.000155125*", pscEx5 = "Exp(.00098485*"
Public Const pscEx2 = "Exp(.000049475*"
Public Const pscAgeFormat = "[>=100]0;[>0]0.0;0"
Public Const pscAgeErrFormat = "[>=1]0;[>0.1]0.0;0.00"
Public Const pscSq = "SQUID 2"

Public Const picDatRowOffs = 4, picNameDateCol = 8, picPksScansCol = 13
Public Const picAuto = -1, picNumAutochtVars = 10, picDatCol = 3
Public Const picSBMcol = 30, picWtdMeanAColOffs = 3, picRawfiletypeCol = 19
Public Const picRawfileFirstcolCol = 20


Public Const pdcErrVal = -9.87654321012346, pdcLog10 = 2.30258509299405
Public Const pdcTiny = 1E-32, pdcMillion = 1000000, pdcBillion = 1000000000
Public Const pdcUrat = 137.88, pdcSecsPerYear = 31557600, pdcSecsPerDay = 86400
Public Const pdcSecsPerHour = 3600

Public Enum IndexType
  peRatio = 1
  peEquation = 2
  peColumnHeader = 3
  pePrefsConstant = 4
  peTaskConstant = 5
  peBothConstant = 6
  peUndefinedConstant = 7
End Enum

Public Enum Plots
  peNoTicks
  peAutoScale
  peZeroMinScale
  peLogScale
  peAutoTicks
  peErrBox
End Enum

Public Enum MinOrMax
  peMaxAutochts = 8
  peMaxConsts = 300
  peMinNumInGroup = 2
  peMaxNukes = 50
  peMaxRats = 52
  peMaxEqns = 50
  peMinMaxTicks = 2
  pemaxrow = 65536
  peMaxCol = 256
  peMinDriftcorrNumSpots = 8
End Enum

Public Enum TaskCol
  peTaskNameCol = 1
  peReadyCol = 2
  peTaskValCol = 4
  peTaskIndxCol = 2
  peTaskNvalsCol = 3
End Enum

Public Enum FormRes
  peUndefined = -1
  peCancel
  peOk
  peHelp
  peBack
  peDelete
  peStored
  peNext
  peRatiosOnly
  peUPbEqnsOnly
  peRunTableOnly
  peReturn
  peGeneralEqnsOnly
  peAutoChartsOnly
  peYes
  peNo
  pePrefsEqualsTask
  peTaskEqualsPrefs
  peDefineNewConst
  peLoadNewFile
End Enum

Public Enum Hues
  peUformBclr = -2147483633
  peMedGray = 10526880
  peDarkGray = 7368816
  peStraw = 8454143
  peLightGray = 12632256
  peGray = 8421504
  peVeryDarkred = 64&
  PeDarkRed = 128&
  pePaleGreen = 13434828
End Enum

' Boolean
Public pbSbmNorm As Boolean, pbSortThis As Boolean, pbExtractAgeGroups As Boolean
Public pbGrpallNoAge As Boolean, pbIgSpaces As Boolean, pbIgCase As Boolean
Public pbIgDashes As Boolean, pbIgSlashes As Boolean
Public pbDo8corr As Boolean, pbSqdBars As Boolean, pbChangedName As Boolean
Public pbShortCondensed As Boolean, pbChangedRunTable As Boolean, pbChangedRatios As Boolean
Public pbChangedUPbSpecial As Boolean, pbChangedEquations As Boolean
Public pbChangedAutoGrafs As Boolean, pbAlwaysShort As Boolean
Public pbGrpCommPbSpecific As Boolean, pbExtractSpotNameGroups As Boolean ' 09/12/06 -- added

Public pbIgPeriods As Boolean, PbIgCommas As Boolean, pbIgColons As Boolean, pbIgSemicolons As Boolean
Public pbGrpAll As Boolean, pbDone As Boolean, pbStdsOnly As Boolean, pbFromSetup As Boolean
Public pbBothStd As Boolean, pbStd As Boolean, pbMonthDayYear As Boolean
Public pbUPb As Boolean, pbFoundStdName As Boolean, pbCanAppendConstant As Boolean
Public pbHasTh As Boolean, pbHasU As Boolean, pbRecycleCondensedSht As Boolean
Public pbTh As Boolean, pbU As Boolean, pbDoMagGrafix As Boolean
Public pbDefiningNew As Boolean, pbNameFragsMatched As Boolean, pbNotRename As Boolean
Public pbHasUconc As Boolean, pbHasThConc As Boolean, pbTaskChanged As Boolean, pbRedefineTaskConst As Boolean
Public pbUconcStd As Boolean, pbThConcStd As Boolean, pbEditingTask As Boolean, PbCopyTask As Boolean
Public pbLinfitSpecial As Boolean, pbLinfitEqns As Boolean, pbLinfitRats As Boolean, pbLinfitRatsDiff As Boolean
Public pbSecularTrend As Boolean, pbRatioDat As Boolean, pbPDfile As Boolean, pbXMLfile As Boolean
Public pbTaskEqualsPrefs  As Boolean, pbPrefsEqualsTask As Boolean, pbDefineNewConst As Boolean
Public pbCanDriftCorr As Boolean, pbEscapeSquid As Boolean, pbCalc8corrConcPlotRats As Boolean

' Array Boolean
Public pbCenteredPk() As Boolean, pbSamRej() As Boolean, pbStdRej() As Boolean
Public pbRatsPlaced() As Boolean, pbEqnsPlaced() As Boolean

' SIMPLE STRING
Public psAgeStdNa$, psConcStdNa$, psStN$, psStS$, psSpotName$, psExcelVersion$, psWname$
Public psStdFont$, psProgName$, psSelectedColHdr$, psBrQL$, psBrQR$, psGrpAgeTypeColName$
Public psRadify6$, psRadify8$, psSqrt$, psPm1sig$, psTwbName$
Public psTaskFilename$, psUstandard$, psNewTaskName$, psNewTaskFilename$
Public psRawFileName$, psXmlFileType$, psRawdatSoftwareVer$, psLastSquidChartName$

' ARRAY STRING
Public psaGrpNames$(1 To 6), psaPDdaNuke$(1 To 2), psaSpotNames$(), psaSpotDateTime$()
Public psaPDdaMass$(1 To 2), psaPDnumRat$(1 To 2), psaPDrat$(1 To 2), psaPDrat_$(1 To 2)
Public psaPDele$(1 To 2), psaCalibConstNumFor$(1 To 2), psaWtdMeanAChartName$(1 To 2)
Public psaStOrSa$(-1 To 0), psaPDeleRat$(1 To 2), psaPDpaNuke$(1 To 2), psaPDpaMass$(1 To 2)
Public psaPDradRat$(1 To 2), psaPDradRat_$(1 To 2), psaC64$(0 To 1), psaC76$(0 To 1)
Public psaC86$(0 To 1), psaC74$(0 To 1), psaC84$(0 To 1), psaUsrConstNa$()
Public psaEqHdr$(), psaTotCtsHdrs$(), psaEqShow$(-4 To peMaxEqns, 1 To 3)
Public psaCPScolHdr$(), psaUThPbConstColNames$(-1 To 0, 1 To 2), psaLowessColHdrs$(1 To 5)

' SIMPLE INTEGER
Public piPrimeDP%, piNtotctsHdrs%, pbActiveEquation%
Public piConcordPlotCol%, piConcordPlotRow%
Public piNconstsUsed%, piNumAllSpots%, piSpotOutputCol%
Public piLastCol%, piLastVisibleCol, piSlastCol%, piStartingSpot%, piEndingSpot%

Public piNameCol%, piDateTimeCol%, piDiscordCol%, piBkrdCtsCol%, piPb204ctsCol%, piPb206ctsCol%
Public piPb46col%, piPb76col%, piPb86col%, piPb46eCol%, piPb76eCol%, piPb86eCol%
Public piSK64col%, piSK76col%, piSK86col%, piPb68eCol%, piHoursCol%, piLowessHrsCol%
Public piStageXcol%, piStageYcol%, piStageZcol%, piQt1yCol%, piQt1Zcol%, piPrimaryBeamCol%
Public piPb7U5_4col%, piPb7U5_4ecol%, piPb6U8_4col%, piPb6U8_4ecol%, piPb7U5Pb6U8_4rhoCol%
Public piPb6U8_7col%, piPb6U8_7ecol%, piPb8Th2_7col%, piPb8Th2_7ecol%, piU8Pb6_8col%, piU8Pb6_8ecol%
Public piPb7U5_8col%, piPb7U5_8ecol%, piPb6U8_8col%, piPb6U8_8ecol%, piPb7U5Pb6U8_8rhoCol%
Public piPb76_8col%, piPb76_8ecol%, piPb8Th2_4col%, piPb8Th2_4eCol%
Public piU8Pb6_totCol%, piU8Pb6_TotEcol%, piPb76_totCol%, piPb76_totEcol%
Public piPb6U8_totCol%, piPb6U8_totEcol%, piPb8Th2_totCol%, piPb8Th2_totEcol%
Public piStdPb7U5_4col%, piStdPb7U5_4eCol%, piStdPb6U8_4col%, piStdPb6U8_4eCol%
Public piStdPb7U5Pb6U8_4rhoCol%, piStdPb76_4Col%, piStdPb76_4eCol%
Public piPb46_7col%, piPb46_7eCol%, piPb46_8col%, piPb46_8eCol%
Public piU8Pb6_4col%, piU8Pb6_4ecol%, piPb76_4col%, piPb76_4eCol%
Public piStdCom6_4col%, piStdCom6_7col%, piStdCom6_8col%
Public piStdCom8_4col%, piStdCom8_7col%, piCom6_4col%, piCom6_7col%, piCom6_8col%
'Public piCom6_4ecol%, piCom6_7ecol%, piCom6_8ecol%
Public piStdCom8col%, piCom8_4col%, piCom8_7col% ', piCom8_4ecol%, piCom8_7ecol%
Public piPb86_4col%, piPb86_4ecol%, piPb86_7col%, piPb86_7ecol%, piStdRadPb86col%, piStdRadPb86ecol%

Public piAgePb6U8_4col%, piAgePb6U8_4ecol%, piAgePb6U8_7col%, piAgePb6U8_7ecol%
Public piAgePb8Th2_7col%, piAgePb8Th2_7ecol%, piAgePb8Th2_4col%, piAgePb8Th2_4ecol%
Public piAgePb7U5_8col, piAgePb7U5_8ecol, piAgePb6U8_8col%, piAgePb6U8_8ecol%
Public piAgePb76_8col%, piAgePb76_8ecol%

Public piIzoom%, piParentEleStdN%, piFirstRatCol%
Public piNscans%, piGrpDateType%, piLwrIndx%, piFileNpks%
Public piStdCorrType%, piOverCtCorrType%, piGrpCommPbSKagetype%
Public piBkrdPkOrder%, piRefPkOrder%, pi204PkOrder%, pi206PkOrder%, pi207PkOrder%
Public pi238PkOrder%, pi208PkOrder%, pi232PkOrder%, pi248PkOrder%, pi254PkOrder%, pi264PkOrder%
Public pi270PkOrder%, piNsChars%, piNgChars%, piNshortList%, piWLrej%, piFormRes%
Public pi46ratOrder%, pi76ratOrder%, pi86ratOrder%, piU1Th2%, piTaskNum%
Public piNumDauPar%, piSqidNumCol%, piNshtsIn%, piParentIso%
Public piNoDupePkN%(), piSpotNum%
Public piNumConcStdSpots%, piLastN%, piSelectedConstIndx% ', piSmoothingWindow%
Public piMswdCt%, piLowessDeltaPcol%, piLowessMeasCol%, piUnDriftCorrConstCol%
Public piAllSigDeltaP%, piAllDeltaPcol%, piTrimCt%

' ARRAY INTEGER
Public piaPpmUcol%(0 To 1), piaPpmThcol%(0 To 1) ' piaComDauCol%(0 To 1),
Public piaOverCts4Col%(7 To 8), piaOverCts46Col%(7 To 8), piacorrAdeltCol%(7 To 8)
Public piaOverCts46eCol%(7 To 8)
'Public piaPb76_4col%(0 To 1), piaPb76_4eCol%(0 To 1)
'Public piaCom6_4col%(0 To 1), piaCom6_7col%(0 To 1), piaCom6_8col%(0 To 1)
'Public piaCom8_4ecol%(0 To 1), piaCom8_7ecol%(0 To 1)
Public piaAgePb76_4Col%(0 To 1), piaAgePb76_4eCol%(0 To 1), piaTh2U8col%(0 To 1)
Public piaTh2U8ecol%(0 To 1), piaRadDauCol_4%(6 To 8), piaRadDauCol_7%(6 To 8), piaRadDauCol_8%, piaAcol%(1 To 2), piaAeCol%(1 To 2)
Public piaSacol%(1 To 2), piaSaEcol%(1 To 2), piaStdUnCorrAcol%(1 To 2), piaStdUnCorrAerCol%(1 To 2)
Public piaSageCol%(1 To 2), piaSageEcol%(1 To 2), piaRadPb86col%(0 To 1) ', piaRadPb86eCol%(0 To 1)
Public piaEqCol%(), piaEqEcol%(), piaEqPkOrd%()
Public piaIsoRat%(), piaIsoRatCol%(), piaIsoRatEcol%(), piaNumSpots%(0 To 1)
Public piaStartSpotIndx%(0 To 1), piaEndSpotIndx%(0 To 1), piaSpotIndx%(0 To 1)
Public piaSpotCt%(0 To 1), piaIsoRatOrder%(), piaCPScol%(), piaBrakType%()
Public piaEqPkUndupeOrd%(), piaEqnRats%(), piaNeqnTerms%(), piaIsoRatsPkOrd%(), piaSwapCols%()
Public piaConcStdSpots%(), piaSpots%(), piaFileNscans%()

' SIMPLE LONG
Public plSpotOutputRw&, plHdrRw&, plSBMzero&, plOutputRw&, plFileNlines&

' ARRAY LONG
Public plaFirstDatRw&(0 To 1), plaLastDatRw&(0 To 1), plaConcStdRows&()
Public plaSpotNameRowsRaw&(), plaSpotNameRowsCond&()

' SIMPLE SINGLE

' SIMPLE DOUBLE
Public pdAgeStdAge#, pdStdAgePbPb#, pdStdAgeThPb#, pdStdAgeUPb#, pdComm76#, pdComm86#, pdComm64#
Public pdConcStdPpm#, pdMinProb#, pdMinFract#, pdMeanParentEleA#
Public pdBkrdCPS#, pdNetCps204#, pdNetCps206#, pdNetCpsCtRef#
Public pdTotCps204#, pdTotCps206#, pdStdPbPbAge#, pdDeadTimeSecs#

' ARRAY DOUBLE
Public pdaPkCts#(), pdaPkT#(), pdaPkNetCps#(), pdaPkFerr#(), pdaSBMcps#(), pdaIntT#(), pdaPkMass#()
Public pdaUsrConstVal#(), pdaTrimMass#(), pdaTrimTime#(), pdaSbmDeltaPcnt#()
Public pdaFileNominal#(), pdaFileMass#(), pdaTotCps#(), pdaNetCps#()

' OBJECTS
Public pwDatBk As Workbook, phRatSht As Worksheet
Public pwTaskBook As Workbook, phStdSht As Worksheet, phSamSht As Worksheet
Public phGrafSht As Worksheet, pwGrafBk As Workbook, pwCopyCatWbk As Workbook
Public phCondensedSht As Worksheet

'Simple Range
Public prSubsSpotNameFr As Range, prSubsetEqnNa As Range
Public prConstsRange As Range, prConstNames As Range, prConstValues As Range

' SQUID-DEFINED TYPES
Public puTask As SquidTask, puTaskCat As SquidTaskCatalog, puTrail As TaskEditTrail

Dim poAppObject As New AppClass

Type RawData
  baCenteredPeak() As Boolean
  saSpotName As String:        sDate As String:            sTimeOfDay As String
  saNukeLabels() As String:    saDetector() As String
  iNscans As Integer:          iNpeaks As Integer
  lSBMzero As Long
  dDeadTime As Double:         dPrimaryBeam As Double
  dQt1y As Double:             dQt1z As Double
  dStageX As Double:           dStageY As Double:          dStageZ As Double
  daTrueMass() As Double:      daTrimMass() As Double:     daIntegrTimes() As Double
  daWaitTimes() As Double:     daPkCts() As Double:        daSBMcts() As Double
  daPkSigmaMean() As Double:   daTimeStamp() As Double
End Type

Type Lowess
  bPercentErrs As Boolean
  iWindow As Integer:     iActualWindow As Integer
  dMean As Double:        dSigmaMean As Double:       dExtSigma As Double
  dMSWD As Double:        dSmoothedMSWD As Double:    dProbfit As Double
  dInitialMSWD As Double
  daX() As Double:        daY() As Double:            daYsig() As Double
  rX As Range:            rY As Range:                rYsig As Range
End Type

Type TaskEditTrail
  bFromName As Boolean:      bFromRuntable As Boolean
  bFromIsoRatios As Boolean: bFromUPbSpecial As Boolean
  bFromEquations As Boolean
End Type

Type SquidTaskCatalog
  baTypes() As Boolean
  saMinerals() As String:          saCreators() As String:        saNames() As String
  saFileNames() As String:         saDescr() As String:           saEqnNaList() As String
  saEqnSubsetNaList() As String:   saEqnNa() As String:           saEqnSubsetNa() As String
  saNomiMassList() As String:      saTrueMassList() As String
  iNameCol As Integer:             iFileNameCol As Integer:       iCreatorCol As Integer
  iMinCol As Integer:              iTypeCol As Integer:           iDescrCol As Integer
  iNominalCol As Integer:          iTrueCol As Integer:           iNeqnsCol As Integer
  iEqnNaCol As Integer:            iFileBkrdPkOrder As Integer:   iSubsNaCol As Integer
  iTotNtasksCol As Integer:        iUPbNtasksCol As Integer:      iGenNtasksCol As Integer
  iNpksCol As Integer:             iNumAllTasks As Integer
  iaNpeaks() As Integer:           iaNeqns() As Integer:          iaNumTasks(0 To 1) As Integer
  lGenNtasksRw As Long:            lFirstRw As Long:              lFirstUPbRw As Long
  lLastUPbRw As Long:              lFirstGenRW As Long:           lLastGenRW As Long
  lTotNtasksRw As Long:            lUPbNtasksRw As Long
  daNominalMass() As Double:       daTrueMass() As Double
End Type

Type ParamSwitch
  SC As Boolean:          FO As Boolean
  LA As Boolean:          ST As Boolean
  SA As Boolean:          Ar As Boolean
  Nu As Boolean:          ZC As Boolean
  HI As Boolean
  ArrNrows As Integer:    ArrNcols As Integer
End Type

Type Autocharts
  sXname As String:        sYname As String
  iXcol As Integer:        iYcol As Integer
  bAutoscaleX As Boolean:  bAutoscaleY As Boolean
  bZeroXmin As Boolean:    bZeroYmin As Boolean
  bLogX As Boolean:        bLogY As Boolean
  bRegress As Boolean:     bAverage As Boolean
End Type

Type SquidTask
  bIsUPb As Boolean:          bDirectAltPD As Boolean
  baSolverCall() As Boolean:  baCPScol() As Boolean
  baHiddenMass() As Boolean   'baHiddenEqn() As Boolean
  sDescr As String:           sDateCreated As String
  sLastModified As String:    sLastAccessed As String
  sName As String:            sFileName As String
  sBySquidVersion As String:  sMineral As String          '   e.g. 204,206  207,206, 208/206
  sCreator As String
  saIsoRats() As String:      saEqns() As String
  saEqnNames() As String:     saNuclides() As String       ' e.g. "196","204Pb", respectively
  saSubsetSpotNa() As String: saConstNames() As String
  iNpeaks As Integer:         iNrats As Integer
  iNeqns As Integer:          iNumAutoCharts As Integer
  iNconsts As Integer:        iParentIso As Integer
  lHdrRw As Long:             lSBMzero As Long:         lOutputRw As Long
  lFirstTaskRw As Long:       lLastTaskRw As Long:      lCreatedByRw As Long
  lFirstRowRw As Long:        lLastRowRw As Long:       lFileNameRw As Long
  lTypeRw As Long:            lNameRw As Long:          lDescrRw As Long
  lDefByRw As Long:           lLastRevRw As Long:       lMineralRw As Long
  lNuclidesRw As Long:        lTrueMassRw As Long:      lNominalMassRw As Long
  lRatiosRw As Long:          lRatioNumRw As Long:      lRatioDenomRw As Long
  lRefmassRw As Long:         lBkrdmassRw As Long:      lParentNuclideRw As Long
  lDirectAltPDrw As Long:     lPrimUThPbEqnRw As Long:  lSecUThPbEqnRw As Long
  lThUeqnRw As Long:          lPpmparentEqnRw As Long:  lEqnsRw As Long
  lEqnNamesRw As Long:        lEqnSwSTrw As Long:       lEqnSwSArw As Long
  lEqnSwSCrw As Long:         lEqnSwLArw As Long:       lEqnSwNUrw As Long
  lEqnSwFOrw As Long:         lEqnSwARrw As Long:       lEqnSwARrowsRw As Long
  lEqnSwARcolsRw As Long:     lEqnSwZCrw As Long:       lEqnSubsNaFrRw As Long
  lConstNamesRw As Long:      lConstValsRw As Long:     lEqnSwHIrw As Long
  lCPScolsRW As Long:         lHiddenMassRw As Long
  laAutoGrfRw(peMaxAutochts) As Long
  daTrueMass() As Double:     daNominal() As Double
  dRefTrimMass As Double:     daNmDmIso() As Double       ' (N,2) NumIso/DenomIso
  dBkrdMass As Double
  vaConstValues() As Variant
  uaSwitches() As ParamSwitch: uaAutographs() As Autocharts
End Type

Sub ReProcess(NewReducedWbk As Boolean, ExistingCondensedSht As Boolean, _
              DeleteReducedShts As Boolean)
' Determine whether the active workbook contains a valid SQ2 condensed-data worksheet;
' If so, delete any associated processed-data worksheets so the workbook is ready for
'   re-processing.
' 09/05/11 -- Rewrite.
' 09/06/16 -- Lines assigning pbLinfit(Special/Eqns/Rats), pbSecularTrend, &
'             pbCalc8corrConcPlotRats moved to Geochron and GenIsoRunStartup to
'             ensure that the assignments are carried out in all cases.
Dim NotDeleted As Boolean, AllowDelete As Boolean, GotSam As Boolean, GotStd As Boolean
Dim i%, Co%, Rw&
Dim Sht As Worksheet, Arr() As Variant

GetInfo

Arr = Array("Within-Spot Ratios", "Trim Masses", _
            "Autocharts", "Task", "Data-reduction params")

NoUpdate
On Error GoTo 0
FindCondensedSheet ExistingCondensedSht
FindStdOrSampleSheets GotSam, GotStd
AllowDelete = GotSam Or GotStd 'Not NewReducedWbk

If ExistingCondensedSht Then
  Set pwDatBk = ActiveWorkbook
  If True Or AllowDelete Then Call DiscardExistingWkShts(AllowDelete)
  pbRecycleCondensedSht = AllowDelete
  NewReducedWbk = Not AllowDelete
  phCondensedSht.Activate
Else
  pbRecycleCondensedSht = False
End If

DeleteReducedShts = AllowDelete And Not NewReducedWbk
NoUpdate False
End Sub

Sub UPbGeochronStartup()
' Initial procedure for processing a U/Pb geochronology PD or XML raw-data file.
'  (Calls GeochronSetup, initiates data-processing)
' 09/03/14 -- Pass the new "NotForGrouping" parameter to GetNameList as TRUE to
'             so that case is ignored in creating the trimmed sample-name list
' 09/05/11 -- Modify for rewritten ReProcess sub, using the new ExistingCondensedSht,
'             NewReducedWbk, and DeleteReducedShts Boolean variables.

Dim BadFile As Boolean, GotIsoplot As Boolean, DoAgain As Boolean
Dim ExistingCondensedSht As Boolean, NewReducedWbk As Boolean
Dim DeleteReducedShts As Boolean
Dim DirIn$, RawDatFolder$, DriveIn$, MustCreateCopy As Boolean, FileLines$()
Dim RawDat() As RawData

pbUPb = True: foUser("sqGeochron") = True
piLwrIndx = -4
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


StartSquid:
DoAgain = False
NewReducedWbk = foUser("SlowReprocess")
ReProcess NewReducedWbk, ExistingCondensedSht, DeleteReducedShts

LoadNewFile:

If Not ExistingCondensedSht Then  ' ie make an entire new processed-data workbook
  DriveIn = fsCurDrive
  DirIn = CurDir
  RawDatFolder = Trim(foUser("sqPDfolder"))
  If RawDatFolder <> "" Then ChDirDrv RawDatFolder

  Do
    ' Not re-using an open workbook's data - open a new file.
    GetFile RawDat, FileLines, BadFile
  Loop Until Not BadFile

  ChDirDrv DirIn
End If

'  How many chars to use in recognizing standard-spot names.
piNsChars = fvMinMax(Val(foUser("nschars")), 1, 6)
piNgChars = fvMinMax(Len(foUser("siAgeStdAge")), 1, 6)
RefreshSampleNames
pbFromSetup = True
'GetNameList , , , True ' Puts all spot-names into groups of like first NsChars, orders them by size.
ManCalc
psStN$ = "'" & pscStdShtNa & "'!": psStS = "'" & pscSamShtNa & "'!" ': KRL$ = "K.R. Ludwig, "
psProgName = fhSquidSht.[ProgName]
'Cbars pscSq, False  ' Hide the SQUID toolbar
If Not fbIsFresh Then BuildTaskCatalog

Do
  On Error GoTo 0

  If Workbooks.Count > 0 And NewReducedWbk And ExistingCondensedSht Then 'fbIsCondensedSheet Then
    phCondensedSht.Copy
    Set phCondensedSht = ActiveSheet
    Set pwDatBk = ActiveWorkbook
  End If

  GeochronSetup.Show

  If piFormRes = peCancel Then
    DoAgain = True
    GoTo StartSquid
  ElseIf piFormRes = peLoadNewFile Then
    ExistingCondensedSht = False
    GoTo LoadNewFile
  ElseIf piFormRes = peUndefined Then
    CrashEnd , "in Sub GenIsoRunStartup.  piFormRes = peUndefined"
  End If

  pbCanDriftCorr = (piNumDauPar = 1 And pbSecularTrend)
  psConcStdNa = Trim(psUstandard) ' psConcStdNa seems to go corrupt unless do this dance
  't1$ = "Invalid Common-Pb "
  pbBothStd = ((LCase(psConcStdNa)) = LCase(psAgeStdNa) Or psConcStdNa = "") ' Age & conc Std are the same.
  SquidGeochron DoAgain, RawDat, FileLines, ExistingCondensedSht
Loop Until Not DoAgain

End Sub

Sub SquidGeochron(DoAgain As Boolean, RawDat() As RawData, FileLines$(), _
                  ExistingCondensedSht As Boolean)
' Master procedure for processing raw-data PD or XML files for U-Pb geochronology.
' 09/07/09 -- Major rewrite of section dealing with existing and newly-added cols:
'             7cor46, 8cor46, 4corcom6, 7corcom6, 8corcom6, 4corcom8, 7corcom8, 4cor86
'             7cor86, 4corppm6, 7corppm6, 8corppm6, 4corppm8, 7corppm8

Dim IgnoredChangedRuntable As Boolean, CanChart As Boolean
Dim NotDone As Boolean, IsSample As Boolean, DidReject As Boolean
Dim IsCalibrConst As Boolean, Bad As Boolean, IsPbTh As Boolean

Dim OrigEqn$, Hdr$, HdrAlias$, ComRat$, DateStr$, CPbRat$
Dim Term1$, Term2$, Term3$, term4$, term5$, term6$, term7$, t1$, t2$, t3$, FinalTerm1$, FinalTerm2$
Dim sTmp$, SpotName$, BadMsg$, UncorrCalibConstCol$, UncorrCalibConstErCol$
Dim EqnResu$, EqnFerro$, MassPos$, MsgFrag1$, MsgFrag2$, PkMass$
Dim Numer$, Denom$, ShtNa$(), ModEqn$()

Dim i%, j%, DauParNum%, Nsheets%, InitialCalcSetting%, ShtNum%, DatRow%, iStd%
Dim RobAvRow%, RobAvCol%, OverCtCol%, WhichPbIso%, RatNum%, TmpRej%
Dim MaxDPnum%, EqNum%, SpotNum%, OutputNcols%, tmpCol%, Std1Sam2%, Startt%, Endd%
Dim OutputNrows%, NameLen%, PkNum%, PkCt%, Nspots%, ColNum%, Indx%, ChtIndx%
Dim ValCol%, ErCol%, ConcordiaCol%, p%, c%, Col%, BadSbm%(0 To 1), OutpCol%

Dim LastRawRow&, FirstSpotRow&, Rw&, r&, Frw&, Lrw&

Dim AvCalibrConst#, AvCalibrConst0#, Cw#, EqnRes#, EqnFerr#, OverCtsDeltaPb7corr#
Dim OverCtsDeltaPb7corrEr#, OverCtsDeltaPb7minusEr#, Seconds#
Dim FirstSecond#, DeltBk#, MinDeltBK#
Dim SbmOffs#(), SbmOffsErr#(), SbmPk#(), StdA#(1 To 2), StdAferr#(1 To 2)
Dim Ratios#(), RatioFractErrs#(), Adrift#(), AdriftErr#()
Dim EqRes#(), EqFerr#(), SbmDelta#()

Dim CalibConst1 As Range, CleanedConst As Range
Dim DataRange As Range, BiwtAv As Range, RawConstRange As Range
Dim CorrConstRange As Range
Dim SumDataRange As Variant, PDrowFields() As Variant
Dim Sht As Worksheet, SqSht(-1 To 0) As Worksheet
Dim WtdMeanAchartObj As ChartObject, vTmp As Variant, ChtObj As ChartObject

DoAgain = False
puTask.iNpeaks = 0
If Not ExistingCondensedSht Then piFileNpks = 0 ' 09/06/11 -- added to force InhaleRawData+ParsePD
If Workbooks.Count > 0 Then InitialCalcSetting = foAp.Calculation
' so can restore when done
foAp.DisplayCommentIndicator = xlCommentIndicatorOnly
NoUpdate
ShowStatusBar
NoUpdate
ManCalc
Alerts False

pbRatioDat = (foUser("ratiodat") = True)              ' /
pbLinfitSpecial = (foUser("linfitspecial") = True)    '|
pbLinfitRatsDiff = (foUser("linfitratsdiff") = True)  '| 09/06/18 -- added
pbLinfitEqns = (foUser("linfiteqns") = True)          '| 09/06/16 -- moved from
pbLinfitRats = (foUser("linfitrats") = True)          '|   ReProcess sub so that is
pbSecularTrend = foUser("SecularTrend")               '|   always implemented.
pbCalc8corrConcPlotRats = foUser("CalcFull8corrErrs") ' \

If Not ExistingCondensedSht Then
  CondenseRawData piFileNpks, pdaFileMass(), False, RawDat, FileLines
End If
Erase RawDat
'pdDeadTimeSecs = 0 ' maybe
LocateStdRows ' Locate the age- and concentration standard spots, catalogue
              '  in which rows each starts.
' Get info on Task and peaks
FirstSpotRow = plaSpotNameRowsCond(1)
psSpotName = psaSpotNames(1)
' 09/06/10 -- added the "and ...." in the following line
pbCanDriftCorr = pbCanDriftCorr And piaNumSpots(1) > peMinDriftcorrNumSpots

With puTask
  .iNpeaks = piFileNpks
  FindStr "dead time", , tmpCol, 2 + FirstSpotRow, picDatCol, , , , , True
  pdDeadTimeSecs = Cells(4 + FirstSpotRow, tmpCol) / pdcBillion

  If pdDeadTimeSecs > 0 Then
    FindStr "sbm zero", , tmpCol, 3 + FirstSpotRow, picDatCol, , , , , True
    plSBMzero = Cells(picDatRowOffs + FirstSpotRow, tmpCol)
  End If

  If pbPDfile And pdDeadTimeSecs = 0 Then
    ParseLine Cells(FirstSpotRow + 1, 1), vTmp, p, ","
    pdDeadTimeSecs = Val(vTmp(4)) / pdcBillion
    plSBMzero = Val(Mid$(vTmp(5), 1 + Len("sbm zero ")))
  End If

  ReDim pdaPkMass(1 To .iNpeaks), pdaTotCps(1 To .iNpeaks)
  ReDim piaCPScol(1 To .iNpeaks), psaCPScolHdr(1 To .iNpeaks)
End With

For PkNum = 1 To puTask.iNpeaks
  pdaPkMass(PkNum) = pdaFileMass(PkNum)
Next PkNum

EqnDetails  ' Determine ratios & their numerator-denominator nuclides,
            '  peak-order for these nuclides...

plHdrRw = 6: piParentEleStdN = 0
With puTask
  .iNpeaks = piFileNpks
  ReDim pbStdRej(1 To piaNumSpots(1), 1 To .iNrats)

  If piaNumSpots(0) > 0 Then
    ReDim pbSamRej(1 To piaNumSpots(0), 1 To .iNrats)
  Else
    pbStdsOnly = True
  End If

  ShowStatusBar

  ReDim Ratios(1 To .iNpeaks), RatioFractErrs(1 To .iNpeaks)
  ReDim pdaTrimMass(1 To .iNpeaks, 2000)
  ReDim pdaTrimTime(1 To .iNpeaks, 1 To 2000)
  ReDim Adrift(1 To piNumDauPar, 1 To piaNumSpots(1))
  ReDim AdriftErr(1 To piNumDauPar, 1 To piaNumSpots(1))
  ReDim pdaSbmDeltaPcnt(1 To .iNpeaks, 1 To piNumAllSpots)
  ReDim ModEqn(piLwrIndx To .iNeqns)
  If .iNeqns > 0 Then ReDim EqRes(1 To .iNeqns), EqFerr(1 To .iNeqns)

  ' Get concentration Std info
  If Len(psConcStdNa) > 0 And .saEqns(-4) <> "" Then
    GetConcStdData
  End If
End With

piTrimCt = 0:   piaSpotCt(0) = 0: piaSpotCt(1) = 0
If pbRatioDat Then
  CreateRatioDatSheet
  phCondensedSht.Activate
End If
' piSpotNum (range from 1 to total# spots analyzed) is the index
'  of the total vector of Spots, in time sequence.

' NumAgestdSpots is the number of Age-Std spots.
'    piaNumSpots  "   "    "    "  Sample    "  .
'    piaNumSpots is the total number of spots in the run.

' piaSpots(1) contains the piSpotNum values of the Age-Std spots.
' piaSpots(0)   "       "     "      "    "   "  Sample       .

' AgeStdSpotIndx is the index of piaSpots(1)() of the current
'  Age-Std spot being processed, & ranges from 1 to NumAgeStdSpots.

' SamSpotIndx is the index of piaSpots(0)() of the current
'  Sample spot being processed, & ranges from 1 to NumSamSpots.

' piaSpotCt(1) is the number of Age-Std spots processed so far.
'    piaSpotCt(0) "   "    "    "  Sample    "       "     "   " .

' piStartingSpot is the Index of the first spot to be processed.
'   piEndingSpot "   "   "    "   "  last   "   "  "      "    .

' StartAgeStdSpotIndx is the index of the First Age-Std spot to be processed.
'    StartSamSpotIndx "   "    "   "   "    "   Sample    "  "  "     "     .

' EndAgeStdSpotIndx is the index of the Last Age-Std spot to be processed.
'    EndSamSpotIndx "   "    "   "   "    "  Sample   "    " "     "     .

' LowerBracket indicates whether the first spot to be processed is the
'  first spot of the run (SpotNum=1) or a specified Age-Std spot.

' UpperBracket indicates whether the last spot to be processed is the
'  last spot of the run (SpotNum=NumSpots) or a specified Age-Std spot.

If foUser("olower") Or foUser("oboth") Then

  For SpotNum = 1 To piaNumSpots(1)
    SpotName = psaSpotNames(piaSpots(1, SpotNum))
    NameLen = Len(SpotName)
    sTmp = Left$(foUser("startstd"), NameLen)

    If SpotName = Left$(sTmp, NameLen) Then
      piaStartSpotIndx(1) = SpotNum: Exit For
    End If

  Next SpotNum

  For piaStartSpotIndx(0) = 1 To piaNumSpots(0)

    If piaSpots(0, piaStartSpotIndx(0)) > piaSpots(1, piaStartSpotIndx(1)) Then
      Exit For
    End If

  Next piaStartSpotIndx(0)

Else
  piaStartSpotIndx(0) = 1: piaStartSpotIndx(1) = 1
End If

If foUser("oUpper") Or foUser("oboth") Then

  For SpotNum = piaNumSpots(1) To 2 Step -1
    SpotName = psaSpotNames(piaSpots(1, SpotNum))
    NameLen = Len(SpotName)
    sTmp = Left$(foUser("endstd"), NameLen)

    If SpotName = Left$(sTmp, NameLen) Then
      piaEndSpotIndx(1) = SpotNum: Exit For
    End If

  Next SpotNum

  If piaEndSpotIndx(1) = 0 Then piaEndSpotIndx(1) = piaNumSpots(1)

  If piaEndSpotIndx(1) = piaNumSpots(1) Then
    piaEndSpotIndx(0) = piaNumSpots(0)
  Else

    For piaEndSpotIndx(0) = piaNumSpots(0) To 2 Step -1
      If piaSpots(0, piaEndSpotIndx(0)) < piaSpots(1, piaEndSpotIndx(1)) Then Exit For
    Next piaEndSpotIndx(0)

  End If
Else
  piaEndSpotIndx(0) = piaNumSpots(0)
  piaEndSpotIndx(1) = piaNumSpots(1)
End If

For Indx = 0 To 1
  piaSpotIndx(Indx) = piaStartSpotIndx(Indx) - 1
Next Indx

CreateSheets True, SqSht
If Not pbStdsOnly Then CreateSheets 0, SqSht

With puTask
  ReDim piaSwapCols(piLwrIndx To .iNeqns)

  For EqNum = piLwrIndx To .iNeqns

    For iStd = -1 To pbStdsOnly
      GetRelocColnum EqNum, piaSwapCols(EqNum), .saEqns(EqNum), _
                     ModEqn(EqNum), (iStd)  '.saEqns(EqNum)
    Next iStd

  Next EqNum

End With

psaStOrSa(-1) = "Standard spots ": psaStOrSa(0) = "Sample spots "
pbStd = False
If piStdCorrType > 0 Then foUser("ShowOverCtCols") = True
CollateUserConstants

' Start of Std-Sample Loop --------------------------------------------------------------

Do ' Loop A
  pbStd = Not pbStd Or pbStdsOnly: pbDone = False
  plSpotOutputRw = plHdrRw

  Do  ' Loop B (First spot-loop)

    Do ' Loop C
      piaSpotCt(-pbStd) = 1 + piaSpotCt(-pbStd)
      piaSpotIndx(-pbStd) = piaSpotIndx(-pbStd) + 1
      piSpotNum = piaSpots(-pbStd, piaSpotIndx(-pbStd))
      ParseRawData piSpotNum, True, IgnoredChangedRuntable, DateStr, True, , True

      If IgnoredChangedRuntable And piNscans > 1 Then
        MsgBox "Run Table changes at spot#" & StR(piSpotNum) & " -- terminating.", , pscSq
        CrashEnd
      End If

    Loop Until Not IgnoredChangedRuntable ' Loop C

    With puTask
      ReDim pbRatsPlaced(1 To .iNpeaks, 1 To piNscans)
      If .iNeqns > 0 Then
        ReDim pbEqnsPlaced(1 To .iNeqns, 1 To piNscans)
      End If
    End With

    StatBar psaStOrSa(pbStd) & ", pass 1: " & psSpotName
    DateStr = psaSpotDateTime(piSpotNum)
    ParseTimedate DateStr$, Seconds
    If FirstSecond = 0 Then FirstSecond = Seconds
    SqSht(pbStd).Activate
    plSpotOutputRw = 1 + plSpotOutputRw
    CFs plSpotOutputRw, 1, psSpotName: CFs plSpotOutputRw, piDateTimeCol, DateStr$
    CF plSpotOutputRw, piHoursCol, foAp.Fixed((Seconds - FirstSecond) / 3600, 3)
    GetRatios Ratios(), RatioFractErrs(), False, BadSbm()  ' Get & place isoratios
    ' Place this row's isotope ratios
    SqSht(pbStd).Activate
    PlaceRawRatios plSpotOutputRw, Ratios, RatioFractErrs
    plHdrRw = flHeaderRow(pbStd)
    foAp.Calculate
    CheckForSolver

    With puTask

      For EqNum = 1 To .iNeqns

        If Not .baSolverCall(EqNum) Then

          If fbOkEqn(EqNum) Then ' All column eqns not having the LA switch (for this row)
            With .uaSwitches(EqNum)

              If Not .LA And Not .SC Then

                If .FO Or piaEqnRats(EqNum) = 0 Or .Ar Then  ' Put in sequence with existing data cols
                  Formulae puTask.saEqns(EqNum), EqNum, pbStd, plSpotOutputRw, _
                           piaEqCol(pbStd, EqNum), plSpotOutputRw
                  ' If specified as by FOrmula rather than by calculated value
                  On Error Resume Next
                  If Not .FO Then EqnResu = Cells(plSpotOutputRw, piaEqCol(pbStd, EqNum))
                  On Error GoTo 0
                Else
                  piSpotOutputCol = piaEqCol(pbStd, EqNum)
                  EqnInterp puTask.saEqns(EqNum), EqNum, EqnRes, EqnFerr, 1, 0
                  EqnResu = fsS(EqnRes)
                  EqnFerro = fsS(100 * EqnFerr)
                  ' results of eqn(eqnum) in eqcol(Std,eqnum)
                  CFs plSpotOutputRw, piaEqCol(pbStd, EqNum), EqnResu
                  CFs plSpotOutputRw, piaEqEcol(pbStd, EqNum), EqnFerro
                End If

                If piaSwapCols(EqNum) <> 0 Then
                  HandleSwapCols EqNum, plSpotOutputRw, pbStd
                End If

              End If
            End With ' .uaSwitches(eqnum)

          End If     ' If fbOkEqn(eqnum) Then

        End If       ' Not .baSolverCall(EqNum)

      Next EqNum

    End With         ' puTask
  Loop Until piaSpotIndx(-pbStd) = piaEndSpotIndx(-pbStd) ' Loop B (end of 1st spot-loop)

  plHdrRw = flHeaderRow(pbStd)
  foAp.Calculate
  Frw = plaFirstDatRw(-pbStd)
  Lrw = plaLastDatRw(-pbStd)

  For EqNum = 1 To puTask.iNeqns  ' Now place any single-cell eqns not marked LAst
    With puTask

      If .baSolverCall(EqNum) Then
        TaskSolverCall EqNum
      Else
        With .uaSwitches(EqNum)

          If fbOkEqn(EqNum) And (.SC Or .ArrNrows = 1) And Not .LA Then
            Formulae puTask.saEqns(EqNum), EqNum, pbStd, _
                    Frw, piaEqCol(pbStd, EqNum)
          End If

        End With
      End If

    End With
  Next EqNum

  plSpotOutputRw = plHdrRw
  piaSpotIndx(-pbStd) = piaStartSpotIndx(-pbStd) - 1
  piaSpotCt(-pbStd) = 0  ' Start of second spot-loop
'q1

  Do   ' Loop D (calculate & place, row-by-row, the Daughter-Parent ratios/constants)

    Do ' Loop E
      piaSpotCt(-pbStd) = 1 + piaSpotCt(-pbStd)
      piaSpotIndx(-pbStd) = piaSpotIndx(-pbStd) + 1
      piSpotNum = piaSpots(-pbStd, piaSpotIndx(-pbStd))
      ParseRawData piSpotNum, False, IgnoredChangedRuntable
    Loop Until Not IgnoredChangedRuntable ' Loop E
'q2
    SqSht(pbStd).Activate
    plSpotOutputRw = 1 + plSpotOutputRw: plOutputRw = plSpotOutputRw
    StatBar psaStOrSa(pbStd) & ", pass 2: " & psSpotName
    MaxDPnum = piNumDauPar
    If MaxDPnum = 2 And puTask.bDirectAltPD And puTask.saEqns(-2) = "" Then MaxDPnum = 1
    foAp.Calculate

    For DauParNum = 1 To MaxDPnum ' 208/232 (or 206/238).
      piSpotOutputCol = piaEqCol(pbStd, -DauParNum)

      EqnInterp puTask.saEqns(-DauParNum), -DauParNum, EqnRes, EqnFerr, 1, TmpRej

      ValCol = IIf(pbStd, piaStdUnCorrAcol(DauParNum), IIf(pbCanDriftCorr, _
                    piUnDriftCorrConstCol, piaAcol(DauParNum)))

      ErCol = IIf(pbStd, piaStdUnCorrAerCol(DauParNum), piaAeCol(DauParNum))
      SqSht(pbStd).Activate
      CFs plSpotOutputRw, ValCol, fsS(CSng(EqnRes))

      If EqnRes <> pdcErrVal And EqnFerr <> pdcErrVal Then _
         CFs plSpotOutputRw, ErCol, fsS(CSng(100 * EqnFerr))
      If DauParNum = 1 Then piWLrej = TmpRej
    Next DauParNum

    foAp.Calculate
    StdElePpm pbStd, plSpotOutputRw                               ' U (or Th) concs
'q3
    If pbHasTh And pbHasU Then

        If piNumDauPar = 1 Then
          ThUfromFormula pbStd, plSpotOutputRw                    ' Calc 232Th/238U
          SecondaryParentPpmFromThU pbStd, plSpotOutputRw         ' Th (or U) concs
          'Tot68_82_fromA
        Else
          ' Sample spots:
          ' 232/238 & tot 206/238-208/232 from A(6/8)-A(8/2)
          ' For pbU,  produces tot206/238, tot238/206, tot208/232
          ' For pbTh, produces tot208/232, tot206/238
          If Not pbStd Then Tot68_82_fromA plSpotOutputRw
          ThUfromA1A2 pbStd, plSpotOutputRw, True
          ' NOTE: must recalc later because WtdMeanA1/2 range doesn't yet exist

          If Not pbStd Then
            SecondaryParentPpmFromThU pbStd, plSpotOutputRw
          End If

        End If

    End If
'q4
    If pbStd Then ' Standard A

      If foUser("ShowOverCtCols") Then
        OverCountColumns (plSpotOutputRw)
      End If

      ' calculate & place apparent 204 overcounts
      If pbU Then
        ComRat = " Pb46 *" & psaC64(1)
      ElseIf piPb86col > 0 Then ' Alpha0/Alpha, Gamm0/Gamma
        ComRat = " Pb46 *" & psaC84(1) & "/ Pb86 "
      End If

      For DauParNum = 1 To piNumDauPar  ' Correct the Pb/U or Pb/Th const for comm-Pb
        sTmp = "(" & fsS(DauParNum) & ")"
        UncorrCalibConstCol = " sUncorrAcol" & sTmp & " "
        UncorrCalibConstErCol = " sUncorrAecol" & sTmp & " "
        IsPbTh = ((pbU And DauParNum = 2) Or (pbTh And DauParNum = 1))
        CPbRat = "sComm" & IIf(pbStd, "1", "0") & "_" & IIf(IsPbTh, "84", "64")

        ' New as of 09/07/09 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        If piStdCorrType = 0 Then
          Term1 = "=(1- Pb46 "
          If IsPbTh Then Term1 = Term1 & "/ Pb86 "
          FinalTerm1 = Term1 & "*" & CPbRat & ")*" & UncorrCalibConstCol ' 4-corrected
        Else     ' Place 4-, 7- or 8-corr calibr const

          Select Case piStdCorrType
            Case 1:                               ' 7-corrected
              If IsPbTh Then
                FinalTerm1 = "=(1- Overcts46(7) / Pb86 *" & CPbRat & ")*" & UncorrCalibConstCol
              Else
                FinalTerm1 = "=(1- Overcts46(7) *" & CPbRat & ")*" & UncorrCalibConstCol
              End If
            Case 2:                               ' 8-corrected
              FinalTerm1 = "=(1- Overcts46(8) *" & CPbRat & ")*" & UncorrCalibConstCol
          End Select

        End If

        PlaceFormulae FinalTerm1, plSpotOutputRw, piaSacol(DauParNum)
        ' Now place the calibr. const errors
        Term1 = "": Term2 = "": Term3 = "": term4 = ""

        If piStdCorrType = 0 And piPb46col > 0 Then ' 4-corrected
          Term1 = UncorrCalibConstErCol & "^2+(" & CPbRat & "/("
          Term1 = Term1 & IIf(IsPbTh, " Pb86 ", "1") ' add in the 206/204 or 208/204 term
          Term1 = Term1 & "/ Pb46 -" & CPbRat & "))^2* Pb46e ^2"
          FinalTerm2 = "=Sqrt(" & Term1 & ")"
          ' Sqrt{UncalConst%err^2 +[Alpha0/(Alpha-Alpha0)]^2 *Pb46%err^2}   or
          ' Sqrt{UncalConst%err^2 +[Gamma0/(Gamma-Gamma0)]^2 *Pb46%err^2}, since
          '  Pb46%err should always be a good approximation of Pb48%err
        Else

          Select Case piStdCorrType
            Case 1:                                 ' 7-corrected
              If piPb46col > 0 And piPb76col > 0 Then
                If IsPbTh Then
                  Term1 = UncorrCalibConstErCol & "^2+(sComm1_84/( Pb86 / Overcts46col(7) "
                  Term2 = "-sComm1_84))^2*( Pb86ecol ^2+ Overcts46ecol(7) ^2)"
                  FinalTerm2 = "=SQRT(" & Term1 & Term2 & ")"
                Else
                  Term1 = UncorrCalibConstErCol & "^2+(sComm1_64/(1/ Overcts46col(7) "
                  Term2 = "-sComm1_64))^2* Overcts46ecol(7) ^2"
                  FinalTerm2 = "=SQRT(" & Term1 & Term2 & ")"
                End If
              End If
            Case 2:                                 ' 8-corrected
              If piPb46col > 0 And piPb86col > 0 And piaTh2U8col(1) > 0 Then
                ' ((Pb86-Stdrad86fact*Th2U8)/(gamma0*stdrad86fact*Th2U8))
                Numer = "( pb86col -Stdrad86fact* Th2U8col(1) )"
                ' (Gamma0-Alpha0*Stdrad86fact*Th2U8)
                Denom = "(Scomm1_84-Scomm1_64*StdRad86fact* Th2U8col(1) )"
                ' rawCalibrConst%err^2
                Term1 = UncorrCalibConstErCol & "^2"
                ' (alpha0*UncorrCalibrConst*8corr46/8corrCalibrConst)^2
                Term2 = "(Scomm1_64*" & UncorrCalibConstCol & "* Overcts46col(8) / sacol(1) )^2"
                ' (Pb86*Pb86%err/Numer)^2
                Term3 = "( Pb86Col * Pb86eCol /" & Numer & ")^2"
                ' (1/Numer+alpha0/Denom)
                term4 = "(1/" & Numer & "+scomm1_64/" & Denom & ")"
                ' stdrad86fact*Th2U8"
                term5 = "stdrad86fact* Th2U8col(1) * Th2U8ecol(1) "
                ' ((1/Numer+alpha0/Denom)*StdRad86fact*Th2U8*Th2U8%err)^2
                term6 = "(" & term4 & "*" & term5 & ")^2"
                ' StdRad86fact*Th2U8*Th2U8%err
                FinalTerm2 = "=sqrt(" & Term1 & "+" & Term2 & "*(" & Term3 & "+" & term6 & "))"
              End If
          End Select

        End If

        If FinalTerm2 <> "" Then
          PlaceFormulae FinalTerm2, plSpotOutputRw, piaSaEcol(DauParNum)
        End If
        ' End new as of 09/07/09 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

      Next DauParNum

    End If ' pbStd
'q5
    plHdrRw = flHeaderRow(pbStd)
  Loop Until piaSpotIndx(-pbStd) = piaEndSpotIndx(-pbStd) ' Loop D- end of 2nd spot-loop
'q6
  StatBar
  SqSht(pbStd).Activate                                 ' Std radiogenic 208Pb/206Pb

  For EqNum = 1 To puTask.iNeqns
    With puTask.uaSwitches(EqNum)

      If .SC And Not .LA And Not ((.SA And pbStd) Or (.ST And Not pbStd)) Then
        OkColWidth piaEqCol(pbStd, EqNum), 6, True
      End If

    End With
  Next EqNum

  ' New as of 09/07/09 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  If Not pbStd Then   ' 09/07/02 -- added
    If piPb46_7col > 0 Then     ' 7-corr Pb46
      Term1 = "=Pb46cor7( Pb76 ,sComm0_64,sComm0_74, AgePb6U8_7 )"
      PlaceFormulae Term1, Frw, piPb46_7col, Lrw
    End If
    If piPb46_8col > 0 Then     ' 8-corr Pb46
      Term1 = "=Pb46cor8( Pb86 , Th2U8(0) ,sComm0_64,sComm0_84, AgePb6U8_8 )"
      PlaceFormulae Term1, Frw, piPb46_8col, Lrw
    End If
  End If

  Term1 = "=100*"

  If pbStd Then
    If piStdCom6_4col > 0 Then
      Term2 = psaC64(1) & "* Pb46 "
      PlaceFormulae Term1 & Term2, Frw, piStdCom6_4col, Lrw
    End If
    If piStdCom6_7col > 0 Then
      Term2 = psaC64(1) & "* OverCts46(7) "
      PlaceFormulae Term1 & Term2, Frw, piStdCom6_7col, Lrw
    End If
    If piStdCom6_8col > 0 Then
      Term2 = psaC64(1) & "* OverCts46(8) "
      PlaceFormulae Term1 & Term2, Frw, piStdCom6_8col, Lrw
    End If
    If piStdCom8_4col > 0 Then
      Term2 = psaC84(1) & "/ Pb86 * Pb46 "
      PlaceFormulae Term1 & Term2, Frw, piStdCom8_4col, Lrw
    End If
    If piStdCom8_7col > 0 Then
      Term2 = psaC84(1) & "/ Pb86 * OverCts46(7) "
      PlaceFormulae Term1 & Term2, Frw, piStdCom8_7col, Lrw
    End If
  Else

    If piCom6_4col > 0 Then
      Term2 = psaC64(0) & "* Pb46 "
      PlaceFormulae Term1 & Term2, Frw, piCom6_4col, Lrw
    End If
    If piCom6_7col > 0 Then
      Term2 = psaC64(0) & "* Pb46_7 "
      PlaceFormulae Term1 & Term2, Frw, piCom6_7col, Lrw
    End If
    If piCom6_8col > 0 Then
      Term2 = psaC64(0) & "* Pb46_8 "
      PlaceFormulae Term1 & Term2, Frw, piCom6_8col, Lrw
    End If
    If piCom8_4col > 0 Then
      Term2 = psaC84(0) & "/ Pb86 * Pb46 "
      PlaceFormulae Term1 & Term2, Frw, piCom8_4col, Lrw
    End If
    If piCom8_7col > 0 Then
      Term2 = psaC84(0) & "/ Pb86 * Pb46_7 "
      PlaceFormulae Term1 & Term2, Frw, piCom8_7col, Lrw
    End If
  End If

  If piPb86col > 0 Then
    FinalTerm1 = "": FinalTerm2 = ""
    ' calculate radiogenic Pb86

    If piPb46col > 0 And (Not pbStd Or piStdCorrType = 0) And piPb86_4col > 0 Then
      ' calculate 4-corr rad86
      FinalTerm1 = "=( Pb86 / Pb46 -" & psaC84(-pbStd) & ")/(1/ Pb46 -" & psaC64(-pbStd) & ")"
      t1 = IIf(pbStd, "StdRadPb86", "Pb86_4")
      Term1 = "(( Pb86e /100* Pb86 )^2"                     '(SigmaPb86^2
      Term2 = "+( " & t1 & " *" & psaC64(-pbStd) & _
              "-" & psaC84(-pbStd) & ")^2" ' +(RadPb86*Alpha0-Gamma0)^2
      Term3 = "/(1-" & psaC64(-pbStd) & "* Pb46 )^2"        ' (1-Alpha0*Pb46)^2
      term4 = "*( Pb46e /100* Pb46 )^2)"                    ' *SigmaPb46^2)
      '(SigmaPb86^2+(RadPb86*Alpha0-Gamma0)^2*SigmaPb46^2)/(1-" & psaC64(-pbStd) & "* Pb46 )^2
      term5 = Term1 & Term2 & term4 & Term3
      FinalTerm2 = "=100*Sqrt(" & term5 & ")/abs( " & t1 & " )"
      OutpCol = IIf(pbStd, piStdRadPb86col, piPb86_4col)
      If OutpCol > 0 Then
        PlaceFormulae FinalTerm1, Frw, OutpCol, Lrw
        PlaceFormulae FinalTerm2, Frw, 1 + OutpCol, Lrw
      End If
    End If

    If piPb76col > 0 And (Not pbStd Or piStdCorrType = 1) And piPb86_4col > 0 Then
      ' calculate 7-corr rad86
      Term1 = "=( Pb86 / "
      Term2 = IIf(pbStd, "overcts46(7)", "Pb46_7")
      FinalTerm1 = Term1 & Term2 & " -" & psaC84(-pbStd) & ")/(1/ " & Term2 & " -" & psaC64(-pbStd) & ")"
      OutpCol = 0
      If pbStd Then
        Term1 = "=StdPb86radCor7per( Pb86 , Pb86e , Pb76 , Pb76e , "
        Term2 = " StdRadPb86 , overcts46(7) ,Std_76,sComm1_64,sComm1_74,sComm1_84)"
        FinalTerm2 = Term1 & Term2
        OutpCol = piStdRadPb86col
      ElseIf piPb6U8_totCol > 0 And piAgePb6U8_7col > 0 Then
        Term1 = "( Pb86 , Pb86e , Pb76 , Pb76e , Pb6U8_tot , Pb6U8_totE , AgePb6U8_7"
        Term2 = " ,sComm0_64,sComm0_74,sComm0_84)"
        FinalTerm2 = "=Pb86radCor7per" & Term1 & Term2
        OutpCol = piPb86_7col
      End If
      If OutpCol > 0 Then
        PlaceFormulae FinalTerm1, Frw, OutpCol, Lrw
        PlaceFormulae FinalTerm2, Frw, 1 + OutpCol, Lrw
      End If

    End If 'piPb76col > 0 And (Not pbStd Or piStdCorrType = 1)

    If FinalTerm1 <> "" Then
      PlaceFormulae FinalTerm1, Frw, OutpCol, Lrw
      If FinalTerm2 <> "" Then PlaceFormulae FinalTerm2, Frw, 1 + OutpCol, Lrw
    End If
    ' End new as of 09/07/09 >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

   End If  'piPb86col > 0

  plHdrRw = flHeaderRow(-pbStd) ' added 09/10/08; otherwise is zero

  If pbStd Then
    SqSht(pbStd).Activate
    NumFormatColumns pbStd
    StatBar "Std weighted-mean calcs"
  Else
    [samcommpb].Cut Cells(2, fiEndCol(plHdrRw) + 1)
  End If

  Columns(1).AutoFit
  Cells(1, 1) = "Isotope Ratios of " & IIf(pbStd, "Standards", "Samples")
  Cells(2, 1) = "(errors are 1s unless otherwise specified)"
  Cells(2, 1).Characters(14, 1).Font.Name = "Symbol"
  With Cells(1, 1).Font
    .Bold = True: .Size = 1.2 * .Size
  End With
  Fonts plHdrRw, 1, plHdrRw + IIf(pbStd, piaSpotCt(1), piaSpotCt(0)), 1, , True, xlRight
  ColWidth picAuto, 2, fiEndCol(r)

  If pbStd Then ' Do some formatting, WtdMean A calcs & chart, robust means
    r = flHeaderRow(True)              ' 09/06/09 -- added the 4 lines to left
    Col = fiEndCol(r)
    Fonts r, 1, , Col, , True, xlRight
    ColWidth picAuto, 2, Col

    If piaSpotCt(1) > 1 Then

      phStdSht.Activate
      AddName "StdHrs", True, plaFirstDatRw(1), piHoursCol, plaLastDatRw(1), piHoursCol
      With puTask
        If pbSbmNorm Then _
          ReDim SbmOffs#(1 To .iNpeaks), SbmOffsErr#(1 To .iNpeaks), SbmPk#(1 To piaSpotCt(1))
      End With
      WtdMeanAcalc BadSbm(), Adrift(), AdriftErr()
      OverCtMeans plaLastDatRw(1)

    End If 'piaSpotCt(1) > 1

    If piaSpotCt(1) <= 1 Then '2 Then ' Must quit
      NoUpdate False
      StatBar

      If pbFoundStdName And piaSpotCt(1) = 0 Then
        Term1$ = "": Term2$ = "Unable to parse the raw-data file."
      Else
        Term2$ = fsVertToLF(fsInQ(psAgeStdNa) & _
                 " -- |Did you correctly specify the Standard name?")
        If piaSpotCt(1) = 1 Then Term1$ = "Only 1 spot found with label similar to " _
          Else: Term1$ = "No spots found with labels similar to "
      End If

      CrashNoise
      MsgBox Term1$ & Term2$, vbOKOnly, pscSq
      Nsheets = Sheets.Count
      ReDim ShtNa(1 To Nsheets)

      For ShtNum = 1 To Nsheets
        ShtNa(ShtNum) = Sheets(ShtNum).Name
      Next ShtNum

      DoAgain = True: pbRecycleCondensedSht = True
      Exit Sub
    End If ' If piaSpotCt(1) <= 1'2

    If pbSbmNorm And BadSbm(1) < (piaSpotCt(1) * puTask.iNpeaks / 2) Then
      SBMdata SbmOffs(), SbmPk(), SbmOffsErr(), pdaSbmDeltaPcnt(), CanChart
      If CanChart Then AddSBMchart plHdrRw, SbmOffs(), SbmOffsErr()
    End If

  End If    ' If pbStd

  If pbHasTh And pbHasU And piNumDauPar = 2 Then

   ' now replace dummy 232/238 "1's with true formulae
    If pbStd And piNumDauPar = 2 Then
      Set CalibConst1 = frSr(1 + plaFirstDatRw(1), [WtdMeanA1].Column, _
                        Lrw)
      Clean DatRange:=CalibConst1, CleanedDat:=CleanedConst, _
            NumCleanRows:=0, AddStrikeThru:=True
      AvCalibrConst = foAp.Average(CalibConst1)
      If CleanedConst.Count < CalibConst1.Count Then p = 12 Else p = 0
    End If

    Do
      p = p + 1
      AvCalibrConst0 = AvCalibrConst

      For DatRow = Frw To Lrw
        ThUfromA1A2 pbStd, (DatRow), False
        ' (because WtdMeanA1 & 2 now exist).
      Next DatRow

      foAp.Calculate
      Redo
    If p > 12 Or Not (pbStd And piNumDauPar = 2) Then Exit Do
      AvCalibrConst = foAp.Average(CalibConst1)
      vTmp = 100 * Abs((AvCalibrConst - AvCalibrConst0) / AvCalibrConst)
    Loop Until p > 1 And vTmp < 0.001

    If piaPpmUcol(-pbStd) > 0 And piaPpmThcol(-pbStd) > 0 Then
      sTmp = fsS(-pbStd)
      Term1 = "= th2u8(" & sTmp & ") * ppmu(" & sTmp & ") *0.9678"
      ValCol = fiColNum("ppmth(" & sTmp & ")")

      If ValCol > 0 Then
        PlaceFormulae Term1, Frw, ValCol, Lrw
        Nformat piaPpmThcol(1), , True
      End If

    End If

    If pbStd Then
      FindStr pscRejectedSpotNums, r, c, 3 + plaLastDatRw(1), , 10 + plaLastDatRw(1)

      If r > 0 And c > 0 Then
        Fonts r + 1, c, , , vbRed, True, xlRight, 12, , , pscFluff, , "arial narrow"
      End If

    End If

  End If

Loop Until Not pbStd Or pbStdsOnly  ' Loop A (end of Std-Sample Loop
' --------------------------------------------------------------------------------

If pbCanDriftCorr And Not pbStdsOnly Then
  Set RawConstRange = frSr(plaFirstDatRw(0), piUnDriftCorrConstCol, plaLastDatRw(0))
  Set CorrConstRange = frSr(plaFirstDatRw(0), piaAcol(1), plaLastDatRw(0))
  DriftCorr RawConstRange, CorrConstRange
End If

' Calculate LA-equation columns & cells here
LastEquations True, True

If foUser("StdConcPlots") Then
  Col = fiEndCol(flHeaderRow(pbStd))
  StdRadiogenicAndAgeCols plaFirstDatRw(1), plaLastDatRw(1)
  Fonts flHeaderRow(pbStd), Col + 1, , Col + 5, , True, xlRight  ' 09/06/09 -- added
End If

If pbRatioDat Then phRatSht.Activate

If Not pbStdsOnly Then
  phSamSht.Activate
  StatBar "Sample calcs"
  ' More formatting

  Rows(plHdrRw).Font.Bold = True

  If piaSpotCt(0) > 0 Or piaSpotCt(1) > 0 Then
    Cells(1, 1).Select
    ' Refresh formulas so will be current - Kluge
    tmpCol = 0
    plSpotOutputRw = plaFirstDatRw(0)
    plOutputRw = plSpotOutputRw

    Do
      tmpCol = tmpCol + 1
    Loop Until Left$(Cells(plSpotOutputRw, tmpCol).Formula, 1) = "=" _
               Or tmpCol = peMaxCol

    If tmpCol < peMaxCol Then
      Do Until IsEmpty(Cells(plSpotOutputRw, 1))
        On Error Resume Next
        Cells(plSpotOutputRw, tmpCol).Formula = Cells(plSpotOutputRw, tmpCol).Formula
        On Error GoTo 0
        plSpotOutputRw = plSpotOutputRw + 1
      Loop
    End If

    plOutputRw = plSpotOutputRw
    If piaSpotCt(0) > 0 Then

      ' ---------------------------------------------------------------
      r = flHeaderRow(0)   ' 09/06/09 -- added to set first & last rows
      SamRadiogenicAndAgeCols plaFirstDatRw(0), plaLastDatRw(0)
      ' ---------------------------------------------------------------

    End If
  End If
  LastEquations False, True
End If ' pbStd

StatBar "formatting"
phStdSht.Activate
pbStd = True

For tmpCol = 2 + fvMax(piaSageCol(1), piaSageCol(2)) To fiEndCol(flHeaderRow(1))
  Hdr = Cells(plHdrRw, tmpCol).Text
  sTmp = Hdr
  Subst sTmp, vbLf

  If InStr(Hdr, "correl") Then
    RangeNumFor pscZd3, 1 + plHdrRw, tmpCol, flEndRow(tmpCol)
    ColWidth picAuto, tmpCol  ' 09/06/09 -- added
  ElseIf InStr(Hdr, "err") Then
    RangeNumFor pscErF, 1 + plHdrRw, tmpCol, flEndRow(tmpCol) 'pscZd2
    ColWidth picAuto, tmpCol ' 09/06/09 -- added
  ElseIf (Left$(Hdr, 1) = "%" And Right$(Hdr, 3) = "206") _
    Or (InStr(Hdr, "232Th") > 0 And InStr(Hdr, "/238U") > 0) Then

  ElseIf sTmp = "4-corr207Pb/206Pbage" Then
    tmpCol = tmpCol + 1
  ElseIf Cells(plHdrRw, tmpCol).Font.Color = vbBlack Then ' <> vbWhite Then
    IsSample = Not pbStd

    For WhichPbIso = 7 To 8
      IsSample = IsSample And Not (tmpCol = piaOverCts4Col(WhichPbIso) _
           Or tmpCol = piaOverCts46Col(WhichPbIso) _
           Or tmpCol = piacorrAdeltCol(WhichPbIso))
    Next WhichPbIso

    If IsSample Then
      Nformat tmpCol, , -pbStd
    End If

  End If

Next tmpCol

tmpCol = 3
IsCalibrConst = False

Do
  tmpCol = tmpCol + 1

  If InStr(LCase(Cells(plHdrRw, tmpCol + 1)), "calibr." & vbLf & "const") > 0 Then
    IsCalibrConst = True
  End If

Loop Until IsCalibrConst

For DauParNum = 1 To piNumDauPar
  ColWidth 0.5, piaSageEcol(DauParNum) + 1
Next DauParNum

If fbRangeNameExists("RbAv") Then

  If Cells(plaFirstDatRw(1), [rbav].Column).Formula = "" Then
    Set BiwtAv = [rbav]  ' "Robust averages", "Robust errors"
    With BiwtAv
      RobAvRow = .Row
      RobAvCol = .Column
      .Cut Cells(1, peMaxCol)
      ColWidth picAuto, RobAvCol
      .Cut Cells(RobAvRow, RobAvCol)
    End With
  End If

End If

DelCol peMaxCol   '  Final formatting & Warning messages

If pbSbmNorm And (BadSbm(0) > 0 Or BadSbm(1) > 0) Then
  MsgFrag1 = "Invalid SBM readings (of zero) were encountered in" _
    & StR(BadSbm(0) + BadSbm(1)) & " of the scans"
    If pbSbmNorm Then MsgFrag1 = MsgFrag1 & _
       fsVertToLF(":||reverted to simple interpolation in these cases.")
  MsgBox MsgFrag1, , pscSq
End If

If foUser("ShowOverCtCols") Then

  If piaOverCts4Col(7) > 0 Then

    phStdSht.Activate

    If IsNumeric([OverCtsDeltaP7corr]) And _
       IsNumeric([OverCtsDeltaP7corrEr]) Then
      ' Notify user if the 207-overcount correction is resolvably >0.3%
      OverCtsDeltaPb7corr = [OverCtsDeltaP7corr].Value
      OverCtsDeltaPb7corrEr = [OverCtsDeltaP7corrEr].Value
      OverCtsDeltaPb7minusEr = Abs(OverCtsDeltaPb7corr) - Abs(OverCtsDeltaPb7corrEr)

      If False And OverCtsDeltaPb7minusEr > 0.3 Then
        With foAp
          MsgFrag1 = "Overcounts on 204 appear to have a significant effect (" _
            & LTrim((Drnd(OverCtsDeltaPb7corr, 2))) & Chr(177) _
            & LTrim((Drnd(OverCtsDeltaPb7corrEr, 2))) & " percent)" _
            & " on the calibration constant."
          End With
        MsgBox MsgFrag1, , pscSq
      End If

    End If

  End If

  If Not foUser("ShowOverCtCols") Then

    For WhichPbIso = 7 To 8

      For ColNum = 1 To 3
        OverCtCol = Choose(ColNum, piaOverCts4Col(WhichPbIso), _
                    piaOverCts46Col(WhichPbIso), piacorrAdeltCol(WhichPbIso))
        If OverCtCol > 0 Then Columns(OverCtCol).Hidden = True
      Next ColNum

  Next WhichPbIso

  End If

  If Not foUser("ShowStdConcordiaCols") Then

    For ConcordiaCol = 1 To 9
      HdrAlias = Choose(ConcordiaCol, "Rad6Pb8U(1)", "Rad6Pb8Ue(1)", _
                        "Rad7Pb5U(1)", "Rad7Pb5Ue(1)", "ErrCorr(1)", _
        "rad8u6pb(1)", "rad8u6pbe(1)", "rad76(1)", "rad76e(1)")
      tmpCol = fiColNum(HdrAlias)
      If tmpCol > 0 Then Columns(tmpCol).Hidden = True
    Next ConcordiaCol

  End If

End If   ' foUser("ShowOverCtCols")

If Not pbStdsOnly Then
  pbStd = False: phSamSht.Activate
  foAp.Calculate
  NumFormatColumns 0
  Fonts 1, 5, , , vbWhite, , xlRight, 1, , , "SquidSampleData"
  AddButton 2, 7, "GroupButton", "Group Me", psTwbName & "GroupThis", _
            RGB(160, 255, 170), 16, 3.5
End If

If Not pbStd Then [samcommpb].Cut Cells(2, 14)
Zoom piIzoom
foAp.Calculate

If foUser("FreezeHeaders") Then

  If Not pbStdsOnly Then
    phSamSht.Activate: pbStd = False
    Cells(plaFirstDatRw(0), 2).Activate: Freeze
  End If

  phStdSht.Activate: ScrollW 1, 1
  Cells(1, 1).Select: Cells(plaFirstDatRw(1), 2).Select: Freeze
  NoUpdate
End If

pbStd = True: phStdSht.Activate
foAp.Iteration = False
ManCalc

For Std1Sam2 = 1 To 2 + pbStdsOnly
  If Std1Sam2 = 2 Then phSamSht.Activate

  For Nspots = 1 To Choose(Std1Sam2, piaNumSpots(1), piaNumSpots(0))

    For RatNum = 1 To puTask.iNrats

      If Std1Sam2 = 1 Then
        DidReject = pbStdRej(Nspots, RatNum)
      Else
        DidReject = pbSamRej(Nspots, RatNum)
      End If

      If DidReject Then
        Fonts plHdrRw + Nspots, piaIsoRatCol(RatNum), , _
              1 + piaIsoRatCol(RatNum), , True
      End If

Next RatNum, Nspots, Std1Sam2


'If Not pbStdsOnly Then phSamSht.Activate: pbStd = False ' 09/06/09 -- removed
phStdSht.Activate: pbStd = True
MsgFrag1 = "Ratios are " & IIf(pbSbmNorm, "", "NOT ") & "normalized to SBM ("
If pbSbmNorm And Not foUser("interpsbmnorm") Then MsgFrag1 = MsgFrag1 & "un-"
MsgFrag1 = MsgFrag1 & "interpolated)"

Term2 = "Spot values for Pb-U-Th Special equations calculated " & _
  IIf(pbLinfitSpecial, "at mid spot-time", "as spot average")
Term2 = Term2 & ", for other Task eqns " & _
  IIf(pbLinfitEqns, "at mid spot-time", "as spot average")
Term2 = Term2 & ", for isotope ratios of the same element " & _
  IIf(pbLinfitRats, "at mid spot-time", "as spot average")
Term2 = Term2 & ", for isotope ratios of different elements " & _
  IIf(pbLinfitRatsDiff, "at mid spot-time", "as spot average") ' 09/06/18 -- added

Cells(1, 7) = MsgFrag1
Fonts 1, 7, , , vbBlue, 0, xlLeft, 12
Cells(2, 7) = Term2
Fonts 2, 7, , , 128, 0, xlLeft, 12
ScrollW fvMax(plaFirstDatRw(1), ActiveSheet.[WtdMeanA1].Row - 26), 3
Zoom piIzoom
Cells(4, 1) = "from file:"
Fonts 4, 4, , , , True, xlLeft
Fonts 4, 2, , , RGB(150, 0, 0), True, xlLeft, , , , phCondensedSht.Name
Cells(3, 1) = phCondensedSht.Cells(2, 1) ' QQ is what?
Fonts 3, 1, , , , , xlLeft, 11
Fonts rw1:=3, Col1:=7, Clr:=RGB(0, 128, 0), Formul:="Task Name:  " & puTask.sName
sTmp = IIf(pbCanDriftCorr, "Corrected", "Uncorrected") & _
       " for secular drift of age standard"
Fonts rw1:=4, Col1:=7, Clr:=IIf(pbCanDriftCorr, vbRed, vbBlack), Formul:=sTmp

With fhSquidSht
  Fonts 5, 7, , , RGB(0, 0, 128), , xlLeft, 12, , , _
        "SQUID " & .[Version] & ", rev. " & .[revdate].Text
End With
Cbars pscSq, pbSqdBars

If pbDoMagGrafix Then TrimMassStuff plaLastDatRw(1) + 50

StatBar

For Std1Sam2 = 1 To 2 + pbStdsOnly
  If Std1Sam2 = 2 Then Set Sht = phSamSht Else Set Sht = phStdSht
  Sht.Select
  With ActiveWindow
    .DisplayWorkbookTabs = True
    .DisplayHorizontalScrollBar = True
    .TabRatio = 0.5
  End With
Next Std1Sam2

If Not pbStdsOnly Then
  phSamSht.Activate
  frSr(1 + flHeaderRow(0), 1, plaLastDatRw(0)).Columns.AutoFit ' 09/06/18 -- added
  HA xlRight, flHeaderRow(0)

  For i = 1 To 6
    j = Choose(i, piAgePb6U8_4col, piAgePb6U8_7col, piAgePb6U8_8col, _
              piaAgePb76_4Col(0), piAgePb8Th2_4col, piAgePb8Th2_7col, _
              piAgePb76_8col)
    If j > 0 Then
      frSr(flHeaderRow(0), j, , j + 1).HorizontalAlignment = xlRight
    End If
  Next i
End If

If Not (pbDoMagGrafix) Then
  pwDatBk.Activate
  phStdSht.Activate
End If

If foUser("StdConcPlots") Then
  StatBar "Std concordia plot"
  phStdSht.Activate
  piNumDauPar = 1 - puTask.bDirectAltPD
  SquidInvokedConcPlot phStdSht, NotDone

  If Not NotDone Then
    Set WtdMeanAchartObj = ActiveSheet.ChartObjects(psaWtdMeanAChartName(piNumDauPar))
    With foLastOb(ActiveSheet.Shapes)
      .Left = 40 + fnRight(WtdMeanAchartObj) - 30 * pbCanDriftCorr
      .Top = WtdMeanAchartObj.Top - 50
    End With
  End If

End If

StatBar "formatting"
PlaceEqnBox

If pbRatioDat Then
  With phRatSht

    For Col = 2 To fiEndCol(3)
      With .Cells(3, Col)
        sTmp = .Text
        With .Columns
          .ColumnWidth = 15
          .EntireColumn.AutoFit
          Cw = .ColumnWidth

          If Cw > 15 Then
            .ColumnWidth = 15
          ElseIf sTmp = "err" And Cw > 8 Then
            .ColumnWidth = 8
          ElseIf sTmp = "Time" And Cw > 6 Then
            .ColumnWidth = 6
          ElseIf Cw > 10 Then
            '.ColumnWidth = 10
          End If

        End With
      End With
    Next Col

    .Rows.AutoFit
  End With
End If

DelCol peMaxCol

If puTask.iNumAutoCharts > 0 Then
  CreateAutoCharts
End If

pbFromSetup = False
With foAp
  .Calculate
  .Calculation = xlAutomatic
End With
LoadStorePrefs 2

phStdSht.Activate
If foUser("AttachTask") Then
  AttachWorksheetInFileToOpenWorkbook puTask.sFileName, 0, , , , True
End If

If foUser("DatRedParamsSeparate") Then
  phStdSht.Activate
  MakeDatRedParamsSht
End If

If Not pbStdsOnly Then HideColumns False

foAp.Iteration = True
Blink
Set DataRange = Nothing
phStdSht.Activate
ScrollW fvMax(plaFirstDatRw(1), [WtdMeanA1].Row - 28), 12
Zoom piIzoom
StatBar "Done"
ActiveSheet.DisplayAutomaticPageBreaks = False
End Sub

Function fiGetRatNum%(NumerPkOrd%, DenomPkOrd%)
Dim i%, Num#, Den#
With puTask
  Num = .daNominal(NumerPkOrd)
  Den = .daNominal(DenomPkOrd)

  For i = 1 To .iNrats

    If Num = .daNmDmIso(1, i) And Den = .daNmDmIso(2, i) Then
      fiGetRatNum = i
      Exit Function
    End If

  Next i

End With
fiGetRatNum = 0
End Function

Sub InterpRat(ByVal NumPkOrd%, ByVal DenPkOrd%, RatioVal#, RatioFractErr#, BadSbm%(), _
  ByVal Std As Boolean, HasZerPk As Boolean) ', RatNum%)
' Calculate peak-height ratios

Const Numer = 1, Denom = 2, Small = 0.0000000001

Dim Singlescan As Boolean, UseSBM As Boolean, Bad As Boolean
Dim CanLinFit As Boolean, ZerPkCt() As Boolean, SameEle As Boolean
Dim Nu$, Nuclide$(1 To 2), Element$(1 To 2), tmp$, Nuke$, Ele$
Dim i%, j%, k%, p%, q%, aOrd%, bOrd%, Snum%, Rct%, Nr%, Sn1%, Num1Denom2%
Dim NumDenom%, RatNum%, MassDiff#
Dim Term1#, Term2#, RatioMean#, RatioMeanSig#, MSWD#, Probfit#, RatioSlope#
Dim SigmaRatioSlope#, RatioInter#, SigmaRatioInter#, CovSlopeInter#
Dim PkF#(), AvPkF#, RhoIJ#, f1#, f2#, ff1#, ff2#, aPk1#, bPk1#
Dim aPk2#, bPk2#, MidTime#, a1PkSig#, a2PkSig#, b1PkSig#, b2PkSig#
Dim ScanDeltaT#, bTfract#, aInterp#, bInterp#, aInterpFerr#, bInterpFerr#
Dim Rnum#, Rden#, a#, b#, CtgStatsRat#, CtgStatsFerr#, aNetCPS#, bNetCPS#
Dim TotT#, MeanT#, RatValVar#, RatValFvar#, TotNumerCts#, TotDenomCts#
Dim InterpA1#(), InterpA2#(), InterpB1#(), InterpB2#()
Dim a1Pk#(), a2Pk#(), b1Pk#(), b2Pk#(), InterpRatVal#(), RatValFerr#(), RatValSig#()
Dim RatioInterpTime#(), SigRho#(), aPkCts#(1 To 2), bPkCts#(1 To 2)

If piNscans = 0 Then Exit Sub
If NumPkOrd = 0 Or DenPkOrd = 0 Then
  MsgBox "Sorry, you've hit a SQUID bug (NumPkOrd or DenPkOrd passed as zero to InterpRat)."
  CrashEnd
End If

SameEle = False
With puTask ' 09/06/18 -- added
  Nuclide(1) = .saNuclides(NumPkOrd)
  Nuclide(2) = .saNuclides(DenPkOrd)
  MassDiff = Abs(.daTrueMass(NumPkOrd) - .daTrueMass(DenPkOrd))

  If MassDiff < 10 Then

    For NumDenom = 1 To 2
      Nu = Nuclide(NumDenom)
      p = InStr(Nu, " ")
      If p > 0 Then Nu = Left(Nu, p - 1)
      Nuke = ""

      For p = 1 To Len(Nu)
        Ele = Mid(Nu, p)
        tmp = Left(Ele, 1)
        If tmp <> "." And Not fbIsNum(tmp) Then Nuke = Ele: Exit For
      Next p

      Element(NumDenom) = Nuke
    Next NumDenom

    If Element(1) = Element(2) Then SameEle = True
  End If
End With

Nr = piNscans - 1
Singlescan = (piNscans = 1)
UseSBM = pbSbmNorm: HasZerPk = False
TotNumerCts = 0: TotDenomCts = 0
RatNum = fiGetRatNum(NumPkOrd, DenPkOrd)

If DenPkOrd > NumPkOrd Then
  aOrd = NumPkOrd: bOrd = DenPkOrd ' 2(204)   4(206)
Else
  aOrd = DenPkOrd: bOrd = NumPkOrd ' 4(206)   5(207)
End If

If (pbUPb And Std) Or Not pbUPb Then TrackSBM Std, UseSBM, BadSbm()

1:

ReDim InterpRatVal(1 To piNscans), RatioInterpTime(1 To piNscans)
ReDim RatValSig(1 To piNscans), ZerPkCt(1 To piNscans), a1Pk(1 To piNscans)
ReDim a2Pk(1 To piNscans), b1Pk(1 To piNscans), b2Pk(1 To piNscans)

If Not Singlescan Then
  ReDim RatValFerr(1 To piNscans), PkF(1 To Nr), InterpB1(1 To Nr), InterpB2(1 To Nr)
  ReDim SigRho(1 To Nr, 1 To Nr), InterpA1(1 To Nr), InterpA2(1 To Nr)

  ' Examples:
  'A: DenomPkOrd>NumPkOrd, eg 204/206 -- aOrd=2 (204), bOrd=4 (206)
  'B: NumPkOrd>DenomPkOrd, eg 207/206 -- aOrd=4,(206), bOrd=5 (207)

  For i = 1 To Nr
    PkF(i) = (pdaPkT(bOrd, i) - pdaPkT(aOrd, i)) / _
             (pdaPkT(aOrd, i + 1) - pdaPkT(aOrd, i))
  'A: [T(206,1)-T(204,1)]/[T(204,2)-T(204,1)]
  'B: [T(207,1)-T(206,1)]/[T(206,2)-T(206,1)]
  Next i

  AvPkF = foAp.Average(PkF)
  f1 = (1 - AvPkF) / 2
  f2 = (1 + AvPkF) / 2
  RhoIJ = (1 - AvPkF ^ 2) / (1 + AvPkF ^ 2) / 2
End If

For i = 1 To piNscans
  TotNumerCts = TotNumerCts + pdaPkNetCps(NumPkOrd, i) * pdaIntT(NumPkOrd)
  TotDenomCts = TotDenomCts + pdaPkNetCps(DenPkOrd, i) * pdaIntT(DenPkOrd)
Next i

If TotNumerCts < 32 Or TotDenomCts < 32 Or Singlescan Then

  If TotDenomCts = 0 Then

    If TotNumerCts = 0 Then
      RatioVal = pdcTiny
    Else
      RatioVal = 1E+16
    End If

    RatioFractErr = 1
  ElseIf TotNumerCts = 0 Then
    RatioVal = pdcTiny
    RatioFractErr = 1
  Else
    RatioVal = (TotNumerCts / pdaIntT(NumPkOrd)) / (TotDenomCts / pdaIntT(DenPkOrd))
    RatioFractErr = sqR(1 / Abs(TotNumerCts) + 1 / Abs(TotDenomCts))
  End If

  ReDim RatioInterpTime(1 To 1), InterpRatVal(1 To 1), RatValFerr(1 To 1)
  RatioInterpTime(1) = (pdaPkT(aOrd, 1) + pdaPkT(bOrd, piNscans)) / 2
  InterpRatVal(1) = RatioVal
  RatValFerr(1) = RatioFractErr

  PlaceRats psSpotName, piSpotNum, 1, RatNum, RatioInterpTime, _
            InterpRatVal, RatValFerr
  Exit Sub
End If

Rct = 0
For Snum = 1 To Nr

  Sn1 = Snum + 1
  TotT = pdaPkT(aOrd, Snum) + pdaPkT(aOrd, Sn1) _
       + pdaPkT(bOrd, Snum) + pdaPkT(bOrd, Sn1)
  MeanT = TotT / 4

  RatioInterpTime(Snum) = MeanT
  ZerPkCt(Snum) = False: ZerPkCt(Sn1) = False
  HasZerPk = False

  For NumDenom = 1 To 2 ' total counts for peaks
    k = Snum + NumDenom - 1
    aNetCPS = pdaPkNetCps(aOrd, k) '   A: 204   B: 206
    bNetCPS = pdaPkNetCps(bOrd, k) '   A: 206   B: 207
    If aNetCPS = pdcErrVal Or bNetCPS = pdcErrVal Then
      HasZerPk = True: ZerPkCt(k) = True: GoTo NextScanNum
    End If
    aPkCts(NumDenom) = aNetCPS * pdaIntT(aOrd)
    bPkCts(NumDenom) = bNetCPS * pdaIntT(bOrd)

    If UseSBM Then

      If pdaSBMcps(aOrd, k) <= 0 Or pdaSBMcps(aOrd, k) = pdcErrVal Or _
        pdaSBMcps(bOrd, k) <= 0 Or pdaSBMcps(bOrd, k) = pdcErrVal Then
          ZerPkCt(k) = True: GoTo NextScanNum
      End If

    End If

  Next NumDenom

  For k = 1 To 2
    NumDenom = Choose(k, 2, 1)
    a = aPkCts(k): b = aPkCts(NumDenom)
    If a <= 0 And b > 16 Then ZerPkCt(Snum + k - 1) = True
    a = bPkCts(k): b = bPkCts(NumDenom)
    If a <= 0 And b > 16 Then ZerPkCt(Snum + k - 1) = True
  Next k

  If ZerPkCt(Snum) Or ZerPkCt(Sn1) Then GoTo NextScanNum

  aPk1 = pdaPkNetCps(aOrd, Snum) ' 204(1)      206(1)
  bPk1 = pdaPkNetCps(bOrd, Snum) ' 206(1)      207(1)
  aPk2 = pdaPkNetCps(aOrd, Sn1)  ' 204(2)      206(2)
  bPk2 = pdaPkNetCps(bOrd, Sn1)  ' 206(2)      207(2)

  If UseSBM Then
    aPk1 = aPk1 / (pdaSBMcps(aOrd, Snum))
    bPk1 = bPk1 / (pdaSBMcps(bOrd, Snum))
    aPk2 = aPk2 / (pdaSBMcps(aOrd, Sn1))
    bPk2 = bPk2 / (pdaSBMcps(bOrd, Sn1))
  End If

  ScanDeltaT = pdaPkT(aOrd, Sn1) - pdaPkT(aOrd, Snum)
  ' T204(2)-T204(1)      T206(2)-T206(1)

  bTfract = (pdaPkT(bOrd, Snum) - pdaPkT(aOrd, Snum))
  ' T206(1)-T204(1)      T207(1)-T206(1)
  PkF(Snum) = bTfract / ScanDeltaT
  ff1 = (1 - PkF(Snum)) / 2
  ff2 = (1 + PkF(Snum)) / 2
  aInterp = ff1 * aPk1 + ff2 * aPk2
  bInterp = ff2 * bPk1 + ff1 * bPk2

  Term1 = (ff1 * pdaPkFerr(aOrd, Snum)) ^ 2
  Term2 = (ff2 * pdaPkFerr(aOrd, Sn1)) ^ 2
  aInterpFerr = sqR(Term1 + Term2) '       204      206

  Term1 = (ff1 * pdaPkFerr(bOrd, Snum)) ^ 2
  Term2 = (ff2 * pdaPkFerr(bOrd, Sn1)) ^ 2
  bInterpFerr = sqR(Term1 + Term2) '       206      207

  Rnum = IIf(NumPkOrd < DenPkOrd, aInterp, bInterp) ' 2(204)   5(206)
  Rden = IIf(DenPkOrd < NumPkOrd, aInterp, bInterp) ' 4(206)   5(207)

  If Rden <> 0 Then
    Rct = Rct + 1
    InterpRatVal(Rct) = Rnum / Rden
    ' Approx. internal ratio variance as sum of variances from the first pks
    a1Pk(Snum) = aPk1:   b1Pk(Snum) = bPk1
    a2Pk(Snum) = aPk2:   b2Pk(Snum) = bPk2
    InterpA1(Snum) = aInterp:  InterpB1(Snum) = bInterp
    a1PkSig = pdaPkFerr(aOrd, Snum) * aPk1
    a2PkSig = pdaPkFerr(aOrd, Sn1) * aPk2
    b1PkSig = pdaPkFerr(bOrd, Snum) * bPk1
    b2PkSig = pdaPkFerr(bOrd, Sn1) * bPk2

    If UseSBM Then
      a1PkSig = sqR(a1PkSig ^ 2 + aPk1 ^ 2 / pdaSBMcps(aOrd, Snum) / pdaIntT(aOrd))
      a2PkSig = sqR(a2PkSig ^ 2 + aPk2 ^ 2 / pdaSBMcps(aOrd, Sn1) / pdaIntT(aOrd))
      b1PkSig = sqR(b1PkSig ^ 2 + bPk1 ^ 2 / pdaSBMcps(bOrd, Snum) / pdaIntT(bOrd))
      b2PkSig = sqR(b2PkSig ^ 2 + bPk2 ^ 2 / pdaSBMcps(bOrd, Sn1) / pdaIntT(bOrd))
    End If

    If aInterp = 0 Or bInterp = 0 Then
      RatValFerr(Rct) = 1
      RatValSig(Rct) = pdcTiny
      SigRho(Rct, Rct) = pdcTiny
    Else
      Term1 = ((f1 * a1PkSig) ^ 2 + (f2 * a2PkSig) ^ 2)
      Term2 = ((f2 * b1PkSig) ^ 2 + (f1 * b2PkSig) ^ 2)
      RatValFvar = Term1 / aInterp ^ 2 + Term2 / bInterp ^ 2
      RatValVar = RatValFvar * InterpRatVal(Rct) ^ 2
      RatValFerr(Rct) = sqR(RatValFvar)
      RatValSig(Rct) = fvMax(Small, sqR(RatValVar))
      SigRho(Rct, Rct) = RatValSig(Rct)

      If Rct > 1 Then

        If ZerPkCt(Snum - 1) Then
          RhoIJ = 0
        Else
          RhoIJ = (1 - PkF(Snum) ^ 2) / (1 + PkF(Snum) ^ 2) / 2
        End If

        SigRho(Rct, Rct - 1) = RhoIJ
        SigRho(Rct - 1, Rct) = RhoIJ
      End If

    End If

  End If

NextScanNum: Next Snum

If Rct = 0 Then
  RatioVal = pdcErrVal
  RatioFractErr = pdcErrVal
  Exit Sub
ElseIf Rct = 1 Then                                                 ' *
  RatioVal = InterpRatVal(1)

  If RatioVal = 0 Then
    RatioVal = pdcTiny: RatioFractErr = 1
  Else
    RatioFractErr = RatValFerr(1)
  End If

Else

  ReDim Preserve InterpRatVal(1 To Rct), RatValFerr(1 To Rct)
  ReDim Preserve RatioInterpTime(1 To Rct), RatValSig(1 To Rct)

  If pbRatioDat Then
    PlaceRats psSpotName, piSpotNum, 1, RatNum, RatioInterpTime, _
              InterpRatVal, RatValFerr
  End If


  CanLinFit = (pbLinfitRats And SameEle) Or _
              (pbLinfitRatsDiff And Not SameEle) ' 09/06/18 -- added

  If CanLinFit And Rct > 3 Then
    ' Error-wtd regression of ratios vs time, ratio-err @ mid burn-time,
    '  using time-adjacent ratio err-correls.
    WtdLinCorr 2, Rct, InterpRatVal(), SigRho(), MSWD, Probfit, 0, _
              RatioInter, SigmaRatioInter, Bad, RatioSlope, _
              SigmaRatioSlope, CovSlopeInter, RatioInterpTime

    MidTime = (pdaPkT(puTask.iNpeaks, piNscans) + pdaPkT(1, 1)) / 2
    RatioMean = RatioSlope * MidTime + RatioInter
    RatioMeanSig = sqR((MidTime * SigmaRatioSlope) ^ 2 + _
      SigmaRatioInter ^ 2 + 2 * MidTime * CovSlopeInter)

  Else ' Error-wtd avg, using time-adjacent ratio err-correls

    WtdLinCorr 1, Rct, InterpRatVal, SigRho, MSWD, Probfit, 0, _
               RatioMean, RatioMeanSig, Bad

  End If

  If Bad Then
    RatioVal = pdcErrVal
    RatioFractErr = pdcErrVal
  ElseIf RatioMean = 0 Then
    RatioVal = pdcTiny
    RatioFractErr = 1
  Else
    RatioVal = RatioMean
    RatioFractErr = fvMax(pdcTiny, RatioMeanSig) / Abs(RatioVal)
  End If

End If   ' RatCt>1

End Sub

Sub ParseRawData(ByVal SpotNumber&, ByVal FirstPass As Boolean, _
  IgnoredChangedRuntable As Boolean, Optional sDate$ = "", _
  Optional GetTrim As Boolean = True, Optional DelMT As Boolean = False, _
  Optional GetTrimSigma As Boolean = False)
' Extracts data for a single SHRIMP spot, correcting for background counts;
'  assumes cursor is on the sample-name column & row of the raw data

Dim s$
Dim Npks%, PkNum%, ScanNum%, PkCt%, a%, b%, Col%, CommaPos%
Dim LastRow&, DatR&, HdrR&, NameRw&, Ncts&, r1&
Dim StartTime#, u#, v#, Seconds#, MeanNetCPS#, NetPkSigma#, tmp#
Dim CtgStatsSigma#, ObsSigma#, SumtotCPSnoerr#
Dim totCPSnoerr#(), zx#(), NetCps#(), TotSBMcts#()
Dim ParsedLine() As Variant

On Error GoTo 0
If SpotNumber = 0 Then Exit Sub
LastRow = flEndRow(1)
IgnoredChangedRuntable = False
NameRw = plaSpotNameRowsCond(SpotNumber)
HdrR = NameRw + picDatRowOffs - 1 ' last hdr line
psSpotName$ = psaSpotNames(SpotNumber)

If FirstPass Then
  sDate = psaSpotDateTime(SpotNumber)
  ParseTimedate sDate, Seconds
  StartTime = Seconds
Else

End If

GetNameDatePksScans NameRw, , , Npks, piNscans

If Npks <> puTask.iNpeaks Then
  If Not IgnoredChangedRuntable And FirstPass Then
    MsgBox "Run Table changes from" & StR(puTask.iNpeaks) & " mass stations to" & _
    StR(Npks) & " at spot#" & StR(piSpotNum) & ".  " & _
    "Ignoring all spots not having the original Run Table.", , pscSq
  End If
  IgnoredChangedRuntable = True: Exit Sub
ElseIf Npks < 2 Or puTask.iNpeaks < 2 And Not FirstPass Then
  MsgBox "Invalid number of pks obtained when parsing the PD file.", , pscSq
  Rows(HdrR).Select: End
ElseIf piNscans < 1 Then
  If FirstPass Then MsgBox "Ignoring spot" & StR(SpotNumber) & _
    " -- fewer than 2 scans.", , pscSq
  IgnoredChangedRuntable = True: Exit Sub
End If

a = puTask.iNpeaks: b = piNscans
ReDim PkSig#(1 To a, 1 To b), pdaIntT(1 To a), pdaPkCts(1 To a, 1 To b)
ReDim pdaPkT(1 To a, 1 To b), pdaPkNetCps(1 To a, 1 To b), totCPS#(1 To a, 1 To b), zx(1 To b)
ReDim totCPSnoerr(1 To b), pdaPkFerr(1 To a, 1 To b), TotSBMcts(1 To a, 1 To b)
ReDim NetCps(1 To b), pbCenteredPk(1 To a), pdaSBMcps(1 To a, 1 To b)

SumtotCPSnoerr = 0
For ScanNum = 1 To piNscans

  If GetTrim And FirstPass Then
    piTrimCt = 1 + piTrimCt
    CheckTrimCtNum piTrimCt, pdaTrimMass, pdaTrimTime
  End If

  For PkCt = 1 To puTask.iNpeaks
    Col = picDatCol + 5 * (PkCt - 1)

    If ScanNum = 1 Then
      pbCenteredPk(PkCt) = (Cells(HdrR, 1 + Col).Font.Underline _
                         = xlUnderlineStyleSingle)
    End If

    DatR = HdrR + ScanNum
    pdaPkT(PkCt, ScanNum) = Cells(DatR, Col)
    With Cells(DatR, Col + 1)

      If .NumberFormat = fsRejFormat Then
        pdaPkCts(PkCt, ScanNum) = pdcErrVal
      Else
       pdaPkCts(PkCt, ScanNum) = .Value
      End If

    End With

    PkSig(PkCt, ScanNum) = Cells(DatR, Col + 2)
    TotSBMcts(PkCt, ScanNum) = Cells(DatR, Col + 3)

    If GetTrim And FirstPass Then
      pdaTrimMass(PkCt, piTrimCt) = Cells(DatR, Col + 4)
      pdaTrimTime(PkCt, piTrimCt) = (pdaPkT(PkCt, ScanNum) + StartTime) / 3600
    End If

    pdaIntT(PkCt) = Cells(HdrR - 1, Col)
    If pdaIntT(PkCt) <= 0 Then GoTo 1
    ' Total counts/second for ith pk of jth scan
    u = pdaPkCts(PkCt, ScanNum)
    totCPS(PkCt, ScanNum) = IIf(u = pdcErrVal, pdcErrVal, u / pdaIntT(PkCt))
    pdaSBMcps(PkCt, ScanNum) = TotSBMcts(PkCt, ScanNum) / pdaIntT(PkCt) - plSBMzero

    If PkCt = piBkrdPkOrder Then
      u = totCPS(PkCt, ScanNum)
      totCPSnoerr(ScanNum) = IIf(u = pdcErrVal Or Not IsNumeric(u), 0, u)
      SumtotCPSnoerr = SumtotCPSnoerr + totCPSnoerr(ScanNum)
    End If

  Next PkCt

Next ScanNum

If piBkrdPkOrder > 0 Then
  v = SumtotCPSnoerr / piNscans

  If v < 10 Then
    pdBkrdCPS = v
  Else
    TukeysBiweight totCPSnoerr(), piNscans, pdBkrdCPS, 9
  End If

Else
  pdBkrdCPS = 0
End If

For PkCt = 1 To puTask.iNpeaks

  If PkCt <> piBkrdPkOrder Then

    For ScanNum = 1 To piNscans
      u = totCPS(PkCt, ScanNum)

      If u = pdcErrVal Then
        pdaPkNetCps(PkCt, ScanNum) = pdcErrVal
      Else
        pdaPkNetCps(PkCt, ScanNum) = u - pdBkrdCPS ' Bkrd-corrected ith pk, jth scan
      End If

      v = Abs(pdaPkNetCps(PkCt, ScanNum))

      If v < 0.000001 Then  ' 10/04/02 -- these 2 lines added to restrict ~0 cps to reasonable errors
        pdaPkFerr(PkCt, ScanNum) = 1
      ElseIf u <> pdcErrVal Then   ' Fractional err in net pk-ht jth scan ith pk
        tmp = Abs(pdaPkCts(PkCt, ScanNum))

        If piBkrdPkOrder > 0 Then
          tmp = tmp + (pdaIntT(PkCt) / pdaIntT(piBkrdPkOrder)) ^ 2 * Abs(pdBkrdCPS)
        End If

        NetPkSigma = sqR(tmp)
        pdaPkFerr(PkCt, ScanNum) = NetPkSigma / (v * pdaIntT(PkCt))
      Else
        pdaPkFerr(PkCt, ScanNum) = pdcTiny
      End If

    Next ScanNum

  End If

Next PkCt

If FirstPass Then
  With puTask

    For PkCt = 1 To .iNpeaks

      If .baCPScol(PkCt) Then

        For ScanNum = 1 To piNscans
          NetCps(ScanNum) = pdaPkNetCps(PkCt, ScanNum)
        Next ScanNum

        With foAp
          MeanNetCPS = foAp.Average(NetCps)
        End With
        pdaTotCps(PkCt) = MeanNetCPS + pdBkrdCPS

        If pbUPb Then
          If PkCt = pi206PkOrder Then pdTotCps206 = pdaTotCps(PkCt)
        ElseIf PkCt = piRefPkOrder Then
          pdNetCpsCtRef = MeanNetCPS
        End If

      End If

    Next PkCt

  End With
End If

Exit Sub
1: MsgBox "Row" & StR(HdrR) & " appears to be corrupted -- cannot parse", , pscSq
End Sub

Sub GetIsoplot(GotIsoplot)  ' Load the Isoplot.xla Add-in if not already done

Dim RevNaFrag$, Msg$, AddinName$, IsoName$, SqVersion$, IsoTitle$, IsoComments$, SqComments$
Dim IsoRev$, IsoRevDate As Date, SqRevDate As Date, IsoVersion$, SqRefs$, PathAndName$, SqTitle$
Dim AddInNum%, Iexist%, IsoNum%, Exist%, Pass%, p%, q%, IsoplotCt%, i%, IsoAddinNum%
Dim Ver#, VerText$, TestVerText$, TestVer#
Dim IsoAddin As AddIn

On Error GoTo 0

GotIsoplot = False: AddInNum = 0:  Iexist = 2

With ThisWorkbook.BuiltinDocumentProperties
  SqTitle = .Item("title")
  SqComments = LCase(.Item("Comments"))
End With
SqVersion = fhSquidSht.[Version]
p = InStr(SqComments, "references")
SqRefs = Trim(Mid(SqComments, p + 10))

' 11/02/03 - added ..............................
' Note that addin names seem to be case sensitive!
With AddIns
  For IsoAddinNum = 1 To .Count
    Set IsoAddin = AddIns(IsoAddinNum)
    If LCase(IsoAddin.Name) = SqRefs Then
      If IsoAddin.Installed Then
        GotIsoplot = True
        Exit For
      End If
    End If
  Next IsoAddinNum
End With
' ...............................................

' 11/02/03 - commented out ------------------
'On Error Resume Next
'GotIsoplot = AddIns(SqRefs).Installed
'On Error GoTo 0
' --------------------------------------------

If Not GotIsoplot Then ' 11/02/03 - added


  IsoplotCt = 0

  For AddInNum = 1 To AddIns.Count             ' Search for either "isoplot.xla" or
    AddinName = LCase(AddIns(AddInNum).Name)   '  "isoplotN.xla" where N = 2 to 9.
    StatBar "Searching add-ins:   " & AddinName

    If Right$(AddinName, 4) = ".xla" Then
      RevNaFrag = StrReverse(Mid$(StrReverse(AddinName), 5))
      Exist = AddIns(AddInNum).Installed
      ' Returns True if installed, False if present but not installed,
      ' error if not present (& not installed).
      ' Exist may be falselyu true, leading to errors if not trapped!

      If Exist Then
        PathAndName = ""
        On Error Resume Next
        PathAndName = Workbooks(AddinName).FullName
        On Error GoTo 0
        If PathAndName = "" Then Exist = False
      End If

      If Left$(RevNaFrag, 7) = "isoplot" And IsoplotCt = 0 Then ' And Exist Then
          On Error GoTo 1
          With Workbooks(AddinName).BuiltinDocumentProperties
            IsoTitle = .Item("title")
            IsoComments = LCase(.Item("Comments"))
          End With

          If AddinName = SqRefs Then

            If Not Exist Then
              AddIns(AddInNum).Installed = True
              'foUser("Splash") = True
            End If

            If AddIns(AddInNum).Installed Then
              GotIsoplot = True
            Else
              MsgBox "Unable to load required Isoplot add-in (" & SqRefs & ")", 1, pscSq
              'foUser("Splash") = True
              End
            End If

          ElseIf Exist Then
            MsgBox "Un-installing incompatible Isoplot version" & vbLf & vbLf & _
              PathAndName & vbLf & vbLf & "Please ignore any warning messages."
            Alerts False
            On Error Resume Next
            Workbooks(AddinName).Close
            AddIns(AddInNum).Installed = False
            On Error GoTo 0
            Alerts True
            MsgBox AddinName & " un-installation complete."
            'foUser("Splash") = True
          End If

      ElseIf Exist And Left$(AddinName, 5) = "squid" Then
        TestVerText = Mid$(AddinName, 6)
        Subst TestVerText, "."
        TestVer = Val(TestVerText)
        VerText = fhSquidSht.[Version]
        Subst VerText, "."
        Ver = Val(VerText)

        If TestVer < Ver And AddinName <> LCase(ThisWorkbook.Name) Then
          Msg = "Un-installing conflicting SQUID version" & vbLf & vbLf & _
            PathAndName & vbLf & vbLf & "Please ignore any warning messages."
          Msg = Msg & vbLf & vbLf & "(if SQUID-2 crashes, please un-install all" _
          & " SQUID add-ins, close Excel, then re-install this version)."
          MsgBox Msg, , pscSq
          Alerts False
          On Error Resume Next
          Workbooks(AddinName).Close
          AddIns(AddInNum).Installed = False
          On Error GoTo 0
          Alerts True
          MsgBox AddinName & " un-installation complete."
          'foUser("Splash") = True
        End If

      End If

    End If

1    On Error GoTo 0
  Next AddInNum

End If

StatBar

If Not GotIsoplot Then
  On Error Resume Next
  IsoTitle = Workbooks(SqRefs).BuiltinDocumentProperties("title")
  On Error GoTo 0

  If IsoTitle <> "" Then
    GotIsoplot = True
  Else
    On Error GoTo 2
    AddIns.Add ThisWorkbook.Path & "\" & SqRefs
    AddIns(SqRefs).Installed = True
    GotIsoplot = True
2   On Error GoTo 0

    If Not GotIsoplot Then
      MsgBox "Unable to find and load a compatible version of Isoplot.  Cannot run SQUID-2.", , pscSq
      End
    End If

  End If

End If

Exit Sub


3 On Error GoTo 0
IsoplotCt = 1 + IsoplotCt
GoTo 1
End Sub

Sub GetInfo(Optional DoRedim As Boolean = True)
' Define fundamental public SQUID-2 variables.
Dim GotStd As Boolean, GotSam As Boolean
Dim Dsch%, s$, CellA1$, d$(1 To 2), p$(1 To 2), e$(1 To 2), Da$(0 To 1)
Dim i%, j%, k%, m%, w As Worksheet

psSqrt = Chr(214): psPm1sig = pscPm & "1" & "sigma"
piIzoom = 100:  psWname$ = "": pbFromSetup = False
psBrQL = "[" & pscQ: psBrQR = pscQ & "]"
GetfsOpSys

Set prSubsSpotNameFr = foUser("NameFrags")
Set prSubsetEqnNa = foUser("SubsetEqns")

If DoRedim Then
  ReDim piaIsoRat(1 To peMaxRats, 1 To 2), piaIsoRatCol%(1 To peMaxRats)
  ReDim piaIsoRatEcol%(1 To peMaxRats)
End If

On Error Resume Next
psWname$ = ActiveWorkbook.Name
On Error GoTo 0

d(1) = "206":    p(1) = "238"
d(2) = "208":    p(2) = "232"
e(1) = "U":      e(2) = "Th"
Da(0) = "sComm0_":  Da(1) = "sComm1_"

For Dsch = 1 To 2
  psaPDdaMass(Dsch) = d(Dsch):          psaPDpaMass(Dsch) = p(Dsch)
  psaPDdaNuke(Dsch) = d(Dsch) & "Pb":   psaPDpaNuke(Dsch) = p(Dsch) & e(Dsch)
  psaPDeleRat(Dsch) = "Pb/" & e(Dsch):  psaPDele(Dsch) = e(Dsch)
  psaPDnumRat(Dsch) = d(Dsch) & "/" & p(Dsch)
  psaPDrat(Dsch) = d(Dsch) & "Pb/" & p(Dsch) & e(Dsch)
  psaPDrat_(Dsch) = d(Dsch) & "Pb|/" & p(Dsch) & e(Dsch)
  psaPDradRat(Dsch) = d(Dsch) & "Pb*/" & p(Dsch) & e(Dsch)
Next Dsch

For i = 0 To 1
  psaC64(i) = Da(i) & "64": psaC76(i) = Da(i) & "76"
  psaC86(i) = Da(i) & "86": psaC84(i) = Da(i) & "84"
  psaC74(i) = Da(i) & "74"
Next i

Set prConstsRange = foUser("Constants")
Set prConstNames = foUser("Constantnames")
Set prConstValues = foUser("constantvalues")
puTask.sBySquidVersion = fhSquidSht.[Version]
AddStrikeThru
End Sub

Sub ColInc(Col%, ByVal ColHeader$, ByVal ColVarName$, ByVal ColIncr%, _
  RatioColNum%, Optional Index, Optional Condition As Boolean = True, _
  Optional ErColNum As Integer, Optional ErColName)
' Define the various column-related variables and increment the Col% parameter.

' Col is the column# of the last-occuped column of the output data-sheet.
' ColHeader is the name to be used as a column header.
' RatioColNum is the col# assigned to the new column.
' Index, Ix2 are indexes of Index x Ix2 variables for U-Pb.
' Condition indicates whether or not the column is to be assigned at all.
' ErColNum is the col# assigned to the error of the column-parameter.

If Not Condition Then Exit Sub
Col = 1 + Col
RatioColNum = Col

If pbUPb Then ' Make entry into the column-index sheet.
  ColIndx Col, ColHeader, ColVarName, Index, (ColIncr = 2), ErColName
End If

If ColIncr = 2 Then Col = 1 + Col: ErColNum = Col
End Sub

Sub OpenTaskEditor() ' Invoke the Task Editor form, return to calling context.
GetInfo
pbEditingTask = False
TaskEditor.Show

If piFormRes = peCancel Or piFormRes = peBack Then

  If pbFromSetup Then

    If pbUPb Then
      UPbGeochronStartup
    Else
      GenIsoRunStartup
    End If

  Else
    End
  End If

End If

End Sub

Sub ParseTimedate(When$, Seconds#) ' Parse date/time from SHRIMP raw-data file
'Type 1:  14:31:36 30/10/1997   Type2:  30/10/97   14:31:36
' Always day/month/year

Dim t$, tm$, Mo$
Dim Hour#, Minute#, c#, Year#, Month#, DayVal#, i#, LastMd#
Dim Secs#, md#(1 To 12), MonthDays#
Dim r() As Variant

pbMonthDayYear = False
Mo$ = "JanFebMarAprMayJunJulAugSepOctNovDec"
tm$ = "302831303130303130313031"

For i = 1 To 12
  md(i) = Val(Mid$(tm$, 1 + (i - 1) * 2, 2))
Next i

t$ = "     ---"
    t$ = Trim(When$)
Do ' fsStrip any doubled spaces
  i = InStr(t$, "  ")
If i = 0 Then Exit Do
  t$ = Left$(t$, i) & Mid$(t$, i + 2)
Loop
On Error GoTo Pdone
ParseLine When, r(), , "-"
Year = Val(r(1))
Month = Val(r(2))
c = InStr(r(3), ",")
If c = 0 Then c = InStr(r(3), " ")
DayVal = Val(Left$(r(3), c - 1))
tm = Trim(Mid$(r(3), c + 1))
c = InStr(tm$, ":")
Hour = Val(Left$(tm$, c - 1))
tm$ = Mid$(tm$, 1 + c)
c = InStr(tm$, ":")
If c = 0 Then ' No seconds in date/time string for some reason

  Minute = Val(tm$): Secs = 0
Else
  Minute = Val(Left$(tm$, c - 1))
  Secs = Val(Mid$(tm$, c + 1))
End If

If Year Mod 4 = 0 Then md(2) = 29
MonthDays = 0

For i = 1 To Month - 1
  MonthDays = MonthDays + md(i)
Next i

If Month > 1 Then LastMd = md(Month - 1) Else LastMd = md(12)
Seconds = (Year - 1990) * pdcSecsPerYear + MonthDays * pdcSecsPerDay + _
       DayVal * pdcSecsPerDay + Hour * pdcSecsPerHour + Minute * 60# + Secs
Exit Sub

Pdone: On Error GoTo 0
When$ = "": Seconds = 0
End Sub

Sub LastUserSCcols(ByVal LastN%, ByVal StdCalc As Boolean)
' Fill data-columns for equations with LA and SC switches =TRUE.
Dim StdIn As Boolean
Dim EqNum%, Col%, Rw&

If LastN > 0 Then  ' ie if any switches include .LA AND .SC
  If pbUPb And StdCalc Then phStdSht.Activate Else phSamSht.Activate

  For EqNum = 1 To puTask.iNeqns  ' Single-cell equations with the LAst switch
    With puTask

      If .baSolverCall(EqNum) Then
        TaskSolverCall EqNum
      Else
        With .uaSwitches(EqNum)
          If pbUPb Then StdIn = pbStd: pbStd = StdCalc

          If fbOkEqn(EqNum) And .LA And .SC Then
            Col = piaEqCol(StdCalc, EqNum)
            Rw = fvMax(plaFirstDatRw(-StdCalc), 1 + _
                       Cells(plaLastDatRw(-StdCalc), Col).End(xlUp).Row)
            Formulae puTask.saEqns(EqNum), EqNum, pbStd, Rw, Col
            RangeNumFor "General", Columns(Col)
            Col = 1 + Col
          End If

          If pbUPb Then pbStd = StdIn
        End With
      End If

    End With
  Next EqNum

End If
foAp.Calculate
End Sub

Sub HandleSwapCols(ByVal EqNum%, ByVal OutputRow&, ByVal Std As Boolean)
' Swap cells for Task Equations having the redirect ("<=>") symbol.

Dim tB As Boolean, IsErrCol As Boolean
Dim OriginalColCellFormula$, tmp1$, tmp2$, NewColCellAddr$
Dim OriginalColCellAddr$, NewColHdr$, OriginalColHdr$, PossibleErrColHdr$
Dim NewEqnCol%, OriginalEqnCol%, PossibleNewErrCol%, PossibleOriginalErrCol%
Dim OriginalErrColCellFormula$, NewColCellFormula$, NewColCellErrorFormula$
Dim NewColCell As Range, OriginalColCell As Range, PossibleErrColHdrCell As Range
Dim PossibleErrColCell As Range

NewEqnCol = piaEqCol(Std, EqNum)
OriginalEqnCol = Abs(piaSwapCols(EqNum)) - (piaSwapCols(EqNum) < 0)
PossibleNewErrCol = piaEqEcol(Std, EqNum)
PossibleOriginalErrCol = 1 + OriginalEqnCol

Set NewColCell = Cells(OutputRow, NewEqnCol)
NewColCellFormula = NewColCell.Formula
NewColCellAddr = NewColCell.Address(False)

Set OriginalColCell = Cells(OutputRow, OriginalEqnCol)
OriginalColCellFormula = OriginalColCell.Formula
OriginalColCellAddr = OriginalColCell.Address(False)
OriginalColHdr = Cells(plHdrRw, OriginalEqnCol).Formula

Set PossibleErrColHdrCell = Cells(plHdrRw, 1 + OriginalEqnCol)
PossibleErrColHdr = fsStrip(PossibleErrColHdrCell.Formula)
IsErrCol = (InStr(PossibleErrColHdr, "%err") > 0) And Not puTask.uaSwitches(EqNum).FO
If PossibleNewErrCol > 0 Then Set PossibleErrColCell = Cells(OutputRow, PossibleNewErrCol)
OriginalErrColCellFormula = Cells(OutputRow, PossibleOriginalErrCol).Formula

Subst NewColCellFormula, OriginalColCellAddr, NewColCellAddr
' irrelevant if a number.
' store the unrelocated, original result in the rawcells array, BY VALUE
CFs OutputRow, NewEqnCol, OriginalColCellFormula
CFs OutputRow, OriginalEqnCol, NewColCellFormula

If IsErrCol Then
  NewColCellErrorFormula = ""
  On Error Resume Next
  NewColCellErrorFormula = Evaluate(OriginalErrColCellFormula)
  On Error GoTo 0
  CFs OutputRow, PossibleOriginalErrCol, NewColCellErrorFormula
  CFs OutputRow, PossibleNewErrCol, OriginalErrColCellFormula
End If

NewColHdr = OriginalColHdr
Subst NewColHdr, "RAW", "raw", "Raw", "raw"

NewColHdr = "RAW" & vbLf & NewColHdr

If NewColHdr <> Cells(plHdrRw, NewEqnCol).Formula Then
  CFs plHdrRw, NewEqnCol, NewColHdr
  ' add "raw" to the original but relocated col-hdr.
End If

End Sub

Sub LastEquations(ByVal StdCalc As Boolean, ByVal IncludeSingleCells As Boolean)
' Process Task Equations with LAst switch=TRUE for both sample & Std
Dim StdIn As Boolean, SpotNa$, EqNum%, SpotsN%, Col%, Rw&

StdIn = pbStd: pbStd = StdCalc
If StdCalc Then phStdSht.Activate Else phSamSht.Activate
foAp.Calculate
With puTask

  For EqNum = 1 To .iNeqns

    If Not .baSolverCall(EqNum) Then
      With .uaSwitches(EqNum)

        If .LA And Not .SC And ((pbStd And Not .SA) Or (Not pbStd And Not .ST)) Then

          If piaEqCol(StdCalc, EqNum) = 0 Then
            piaEqCol(StdCalc, EqNum) = 1 + fiEndCol(plHdrRw)

            If StdCalc Then
              SpotsN = ActiveSheet.[agestdage].Column
              If SpotsN - piaEqCol(StdCalc, EqNum) < 8 Then Columns(SpotsN - 5).Insert
            End If

          End If

          For SpotsN = piaStartSpotIndx(-StdCalc) To piaEndSpotIndx(-StdCalc)
            'To IIf(.sc, 0, piaSpotIndx(-StdCalc))
            Rw = plaFirstDatRw(-StdCalc) + SpotsN - piaStartSpotIndx(-StdCalc)
            SpotNa = psaSpotNames(piaSpots(-StdCalc, SpotsN))
            StatBar psaStOrSa(StdCalc) & ", pass 3, eqn" & StR(EqNum) & "     " & SpotNa

            If fbIsInSubset(SpotNa, EqNum) Then

              If Not StdCalc Then ' EqCol(StdCalcSpotsN) may be valid only for the Std sheet
                FindStr Phrase:=puTask.saEqnNames(EqNum), ColFound:=Col, _
                        RowLook1:=plHdrRw, RowLook2:=plHdrRw, _
                  LegalSheetNameOnly:=True
                If Col = 0 Then Col = piaEqCol(StdCalc, EqNum)
              Else
                Col = piaEqCol(StdCalc, EqNum)
              End If

              Formulae puTask.saEqns(EqNum), EqNum, StdCalc, Rw, Col, Rw

              If piaSwapCols(EqNum) <> 0 Then
                HandleSwapCols EqNum, Rw, StdCalc
              End If

              Cells(Rw, piaEqCol(StdCalc, EqNum)).Interior.Color = vbWhite
            Else
              Cells(Rw, piaEqCol(StdCalc, EqNum)).Interior.Color = peLightGray
            End If

          Next SpotsN

        End If

      End With
    End If

  Next EqNum

End With

If IncludeSingleCells Then
  StatBar IIf(StdCalc, "Std", "Sample") & " age calcs"
  LastUserSCcols piLastN, StdCalc  ' For +LA+SC eqns
End If

pbStd = StdIn
End Sub

Sub ShowToolbarHelp()
ToolbarHelp.Show
End Sub

Sub QuitSquid() ' Un-install SQUID-2
Dim Msg$, ExcelAddin As Excel.AddIn

Msg = "Do you really want to un-install the SQUID2 add-in?" & vbLf & _
  vbLf & "(you can reinstall SQUID later at any time)"
If MsgBox(Msg, vbYesNo, pscSq) <> vbYes Then Exit Sub

For Each ExcelAddin In Application.AddIns

  If ExcelAddin.Name = ThisWorkbook.Name Then
    Alerts False
    On Error Resume Next
    ExcelAddin.Close
    ExcelAddin.Installed = False
    Exit Sub
  End If

Next ExcelAddin
End Sub

Sub SquidToolbarHelp() ' Show HELP for the SQ2 toolbar.
ToolbarHelp.Show
End Sub

Sub InitWbkOpenHandler()
Set poAppObject.AppEvents = Application
End Sub

Sub WBopen(Wb)
' Event-handler sub for the WorkbookOpen event.
CleanupSquidRefs False
End Sub
