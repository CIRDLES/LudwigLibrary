Attribute VB_Name = "Main"
' Isoplot -- Main module
' Must NOT be Option Private Module!!!!
Option Explicit: Option Base 1
' Main module to Isoplot, a VBA program for calculation & graphical presentation
' of radiogenic-isotope data.
' K.R. Ludwig, Berkeley Geochronology Center.

'App.VBE.VBProjects("IsoplotEx").VBComponents("Main").Name = "Main"
' Set P = Application.VBE.VBProjects("IsoplotEx").References
' Set x = P.Item(6): P.Remove x
'-------------------------------------------------------------------------------------
'Application.MacroOptions Macro:="iso193e.xls!Lambda", Description:= _
 "Returns decay const. for the long-lived nuclide whose mass is 'NuclideMass'." _
 , ShortcutKey:=""

Public Const MAXLOG = 1E+308, MINLOG = 1E-307, MAXEXP = 709, EndRow = 65536, IsoCalc = True
Public Const Million = 1000000#, BadT = -1.23456789, MwF = 2, Iso = "Isoplot"
Public Const Und = "_", Hun = 100, Thou = 1000
Public Const Log_10 = 2.30258509299405, pi = 3.14159265358979, TwoPi = 6.28318530717959, Log_2 = 0.693147180559945
Public Const General = "General", IsoLarge = 9.9E+31, IsoSmall = 1.01010101010101E-32

Public Const Black = 0, Brown = 13209, OliveGreen = 13107, DkGreen = 13056, DarkTeal = 6697728
Public Const DarkBlue = 8388608, Indigo = 10040115, Gray80 = 3355443, DarkRed = 128, Orange = 26367
Public Const DarkYellow = 32896, Green = 32768, Teal = 8421376, Blue = 16711680, BlueGray = 10053222
Public Const Gray50 = 8421504, Red = 255, LightOrange = 39423, Lime = 52377, SeaGreen = 6723891
Public Const Aqua = 13421619, LightBlue = 16737843, Violet = 8388736, Gray40 = 9868950, Pink = 16711935
Public Const Gold = 52479, Yellow = 65535, BrightGreen = 65280, Turquoise = 16776960, SkyBlue = 16763904
Public Const Plum = 6697881, Gray25 = 12632256, Rose = 13408767, Tan_ = 10079487, LightYellow = 10092543
Public Const LightGreen = 13434828, LightTurquoise = 16777164, PaleBlue = 16764057, Lavender = 16751052
Public Const LavenderBlue = 16764108, Straw = 13434879, Mauve = 8421631, Cyan = 16776960, White = 16777215

Public Anch As DialogSheet, sc As SeriesCollection, ChrtDat As Worksheet, DatSht As Worksheet
Public YorkRes As Object, ResBox As Object, IsoChrt As Object, ArChrt As Object, DatWbk As Workbook
Public UisoRange() As Range, xErRange As Range, yErRange  As Range, RangeIn(9) As Range
Public CurvRange() As Range, TikRange() As Range, Oo As Range, skV As Range, IsoPlotTypes As Range
Public NoSuper, NoUpdate As Boolean, ShapesOK As Boolean, AddToPlot As Boolean, StackIso As Boolean
Public Eellipse As Boolean, eCross As Boolean, Ebox As Boolean, StraightLine As Boolean, ShowResidPlot As Boolean
Public Canceled As Boolean, Mac As Boolean, Windows As Boolean, Dim3 As Boolean, WideMargins As Boolean
Public AutoScale As Boolean, Anchored As Boolean, Regress As Boolean, OtherXY As Boolean, MacExcelX As Boolean
Public Normal As Boolean, Inverse As Boolean, DoPlot As Boolean, RefChord As Boolean, HeaderRow As Boolean
Public ConcConstr As Boolean, RowWise As Boolean, ColWise As Boolean, Cdecay As Boolean, AutoScale0 As Boolean
Public XYlim As Boolean, CanReject As Boolean, UseriesPlot As Boolean, MTrow() As Boolean, SmNdIso As Boolean
Public ArgonPlot As Boolean, ConcPlot As Boolean, ConcAge As Boolean, WtdAvPlot As Boolean
Public PbAnchor As Boolean, AgeAnchor As Boolean, ErrCorrsReqd As Boolean, ColorPlot As Boolean
Public PlotProj As Boolean, Linear3D As Boolean, Planar3D As Boolean, ClipEllipse As Boolean
Public CumGauss As Boolean, ConcAgePlot As Boolean, Cdecay0 As Boolean, NoPts As Boolean, ArInset As Boolean
Public AbsErrs As Boolean, BoldedData() As Boolean, ArChron As Boolean, ArPlat As Boolean, ArKa As Boolean
Public DoMC As Boolean, DetailsShown As Boolean, Sbar As Boolean, Stacked As Boolean, StackedUseries As Boolean
Public PbTicks As Boolean, PbTickLabels As Boolean, Ierror As Boolean, PbPlot As Boolean, DoMix As Boolean
Public Regress0 As Boolean, DoShape As Boolean, LineAgeTik As Boolean, ArIso As Boolean, ArSpect As Boolean
Public PbGrowth As Boolean, uMultipleEvos As Boolean, uUseTiks As Boolean, KCaIso As Boolean, AgeExtract As Boolean
Public YoungestDetrital As Boolean
Public uPlotIsochrons As Boolean, uLabelTiks As Boolean, uEvoCurve As Boolean, LabelUcurves As Boolean
Public ToolbarVisible As Boolean, Robust As Boolean, RobustType As Integer, BadConst As Boolean
Public UThPbIso As Boolean, ArgonStep As Boolean, ArRestricted As Boolean, ForcedPlateau As Boolean
Public ClassicalIso As Boolean, HasShapes As Boolean, AutoRescale As Boolean, BandBehind As Boolean
Public Aspline As Boolean, Nspline As Boolean, SplineLine As Boolean, SecularEquil As Boolean, WtdAvXY As Boolean
Public PlotErrEnv As Boolean, InvertPlotType As Boolean, ProbPlot As Boolean, AnyCurve As Boolean
Public AutoSort As Boolean, WasPlat As Boolean, AskInfo As Boolean, NumericPrefs As Boolean, GraphicsPrefs As Boolean
Public pm As String * 1, pmm As String * 3, SigLev%, N&, ndCols%, Xcolumn%, AnchorPt%
Public Isotype%, SymbCol%, SymbRow&, CurvTikInter#, Nser%, FromSquid As Boolean
Public excSymb%, excClr&, UsType%, excClrInd%, RightCol%, TopRow&, FirstCurvTik%, OtherIndx%
Public SymbClr%, SymbType%, PbType%, ArType%, Dseries%, DSeriesN%(), Nbins%
Public Ncurves%, AxisLthick%, NforcedSteps%, ExcelVersion, pFirst&, pLast&
Public ValidRow&(), HdrRow&, UisochPos%, Pvlines As Boolean, DatClr&(), ExcelCalc&
Public Uratio#, Lambda235#, Lambda238#, Lambda232#, Lambda87#, Lambda147#, Lambda187#, Lambda176#, Lambda40#
Public Lambda231#, Lambda226#, Lambda210#, Lambda227#, Lambda234#, Lambda230#, LambdaDiff, LambdaK, NumErr
Public Air4036!, ArMinSteps%, ArMinProb#, ArMinGas#, Lambda235err#, Lambda238err#, Lambda234err#, Lambda230err#
Public Xspred#, Yspred#, MinProb#, AnchorErr#, Anchor76#, AnchorAge#, ProjZ#
Public InpDat#(), iLambda(), yfResid#(), AgeSpred#, MinX#, MaxX#, MinY#, MaxY#, MinAge#, MaxAge#
Public MinCurvAge#, MaxCurvAge#, CurvAgeSpred#, uEvoCurvLabelAge#, BinStart#, BinWidth#
Public AnchorT1#, AnchorT2#, Xerror#, Yerror#, Zerror#, Rhos#(3), SigYinit#, Jay#, Jerror#, Jperr#
Public pbAlpha0#, pbBeta0#, pbGamma0#, pbMu#, pbKappaMu#, pbStartAge#, Nuke$(), RadNuke$(), AllNuke$()
Public AxX$, AxY$, AxZ$, DatSheet$, DatBook$, PlotDat$, PlotName$, Irange$, viv$, Dsep As String * 1
Public Lir$, Uir$, Msw$, Lint$, Uint$, AgeRes$, PlCap$, ExcelVer$, qq As String * 1, Sqrt As String * 1
Public uFirstGamma0#, Ugamma0#(), IntSl#(), ErrRho#()
Public PlotBoxLeft#, PlotBoxWidth#, PlotBoxTop#, PlotBoxHeight#
Public PlotBoxRight#, PlotBoxBottom#, TuPbAlpha0#, TuPbBeta0#
Public PubObj(4) As Object, PubText$(4), PubInt%(4), PubVar(4)
Public CurvWithDce As Boolean, ArChronSteps(2) As Integer, ForceFill As Boolean
Public pDots As Boolean, pBars As Boolean, pInpSig As Boolean, pSig As Boolean
Public pRegress As Boolean, pSigma!, HistoStacked As Boolean, FromModeless As Boolean
Public Opt As IsoplotOptions, yf As Yorkfit, Cmisc As ConcordiaAgesMisc
Public LambdaRef As Variant, HardRej As Boolean, XL2007 As Boolean, ShowRes As Boolean

Public Yrat#(2, 2), vcXY#(2, 2), Crs#(41), PbR0#(0 To 2)
Public PbExp#(0 To 2), MuIsh#(0 To 2), ParRat#(0 To 2), PbLambda#(0 To 2)

Public ct&, FromInit As Boolean, FromIso As Boolean

Type ConcordiaAgesMisc
  Xconc As Double:  Yconc As Double:  NoLerr As Boolean
End Type

Type IsoplotOptions
  PlotboxBorder As Boolean:     SheetClr As Long
  PlotboxClr As Long:           AgeTikSymbClr As Long
  AgeTikSymbFillClr As Long:    IsochClr As Long
  UseriesIsochClr As Long:      UseriesIsochStyle As Integer
  IsochStyle As Integer:        CurvClr As Long
  AxisTikLabelFont As String:   AxisTikLabelFontSize As Integer
  AxisNameFont As String:       AxisNameFontSize As Integer
  AxisAutoTikSpace As Boolean:  AgeTikFont As String
  AgeTikFontSize As Integer:    IsochResFont As String
  IsochResFontSize As Integer:  IsochResboxShadw As Boolean
  IsochResboxRnd As Boolean:    ConcResboxRnd As Boolean
  CurveRes As Integer:          AgeTikSymbol As Integer
  AxisThickLine As Boolean:     ClipEllipse As Boolean
  AxisTickCross As Boolean:     ConcLineThick As Integer
  AgeTikSymbSize As Integer:    SimplePlotSymbSize As Integer
  AlwaysPlot2sigma As Boolean:      EndCaps As Boolean
  IsochLineThick As Integer
End Type

Type wWtdAver
  IntMean As Double:    IntMeanErr95 As Double:  IntMeanErr2sigma As Double
  ExtMean As Double:    ExtMeanErr95 As Double:  Ext2Sigma  As Double
  MSWD As Double:       Probability As Double:   ExtMeanErr68 As Double
  BiwtMean As Double:   BiWtSigma As Double:     BiWtErr95 As Double
  Median As Double:     MedianPlusErr As Double: MedianMinusErr As Double:  MedianConf As Double
  ChosenMean As Double: ChosenErr As Double:     ChosenErrPercent As Double
End Type

Type DataPoints
  X As Double: Xerr As Double
  y As Double: Yerr As Double
  z As Double: Zerr As Double
  RhoXY As Double: RhoXZ As Double: RhoYZ As Double
End Type

Type Yorkfit
  Slope As Double:      SlopeError As Double
  Intercept As Double:  InterError As Double
  Xbar As Double:       Ybar As Double
  MSWD As Double:       Prob As Double
  ErrSlApr As Double:   ErrIntApr As Double:   RhoInterSlope As Double
  ErrSlincSc As Double: ErrIntincSc As Double: Emult As Double
  Xinter As Double:     XinterErr As Double
  Model As Integer:     WtdResid() As Double
  LwrSlope As Double:   UpprSlope As Double
  LwrInter As Double:   UpprInter As Double
  Ntrials As Long
  SlopeXZ As Double:    Zinter As Double
  LwrAge As Double:     UpprAge As Double
  LwrXinter As Double:  UpprXinter As Double
  LwrZinter As Double:  UpprZinter As Double
  LwrSlopeXZ As Double: UpprSlopeXZ As Double
End Type

Type Curves
  Ncurvtiks As Integer:  CurvTik() As Double
  NcurvEls() As Integer: Nisocs As Integer
  NageElls As Integer:   CurvTikPresent() As Boolean
End Type

Sub StartFromInit()
Attribute StartFromInit.VB_ProcData.VB_Invoke_Func = " \n14"
Dim ShtIn As Worksheet, rw&

FromInit = True: FromIso = False: FromSquid = False
Set ShtIn = ActiveSheet

With MenuSht
  .Activate
  On Error GoTo 1
  If IsError(.[nodisclaimer]) Then
1    rw = LastRow(3) + 6
    Cells(rw, 3) = "NoDisclaimer"

    With Cells(rw + 1, 3)
      .Value = False
      .Name = "NoDisclaimer"
    End With

  End If

  On Error GoTo 0
  If Not .[nodisclaimer] Then Disclaimer.Show
End With

ShtIn.Activate
Isoplot
End Sub

Sub StartFromIso()
Attribute StartFromIso.VB_ProcData.VB_Invoke_Func = " \n14"
FromInit = False: FromIso = True: FromSquid = False
HardRej = False
Isoplot
End Sub

Sub StartFromSquid()
FromInit = False: FromIso = False: FromSquid = True: DoPlot = True
AutoScale = True: ConcPlot = True: Ncurves = 1: ColorPlot = True
DoShape = True: Eellipse = True: AutoSort = 0: AutoRescale = 0
ReDim CurvRange(1), TikRange(1)
AssignIsoVars
Isoplot , 1, False, False
End Sub

'ISOPLOT
Sub Isoplot(Optional AddData, Optional Itype, Optional Inver, Optional DimThree, Optional LinDim)
Attribute Isoplot.VB_ProcData.VB_Invoke_Func = "i\n14"
' User interface & data entry
Dim BadYork As Boolean, tB As Boolean, tB1 As Boolean, tb2 As Boolean, EsymbPlotted As Boolean
Dim Failed As Boolean, GotStatus As Boolean
Dim Wcap$, tmp$, Rbx$, s1$, s2$
Dim Bct%, iCt%, TxbN%, bIndx%()
Dim tC&, i&, j&, k&, Nrej&, cca&, ccb&, Ntrials&
Dim PlAge#, PlAgeErr#, pFirst%, pLast%, GasFract#, SlpEq#, IntEq#, SlpEqEr#, IntEqEr#, PlInit#, plInitErr#
Dim Np#(), pp#(), xyProj#(0 To 5), Gas#()
Dim wX() As Variant, wXerr() As Variant, Wrejected() As Variant
Dim cb As Object, cbI As Object, Sht As Object, rBox As Object, Txb As Object
Dim wa As wWtdAver

Xcalc
GetConsts True

If Workbooks.Count > 0 Then

  If Ash.ProtectContents Then
    MsgBox "Can't work with protected sheets." & viv$ & _
      "(to unprotect, select Tools/Protection from the Excel menu-bar)", _
      vbOKOnly, Iso
    ExitIsoplot
  End If

End If

If Mac And Val(Menus("MbAvailable")) < 0.3 Then _
    MsgBox "Insufficient memory to reliably run Isoplot", , Iso

If Workbooks.Count = 0 Then Workbooks.Add
AddToPlot = IIf(IM(AddData), False, True)
ShowResidPlot = False
ReDim Preserve Nuke$(1), RadNuke$(1), AllNuke$(1)
StatBar
NoUp False
Failed = True: HasShapes = False: DetailsShown = False

If AddToPlot Then

  For Each Sht In Sheets
    If Sht.Name = DatSheet$ Then Failed = False: Exit For
  Next

  If Failed Then
    MsgBox "The new data must be in the same Workbook as the Chart, " & _
     vbLf & "and in the same Worksheet (" & qq + DatSheet$ & qq & _
     ") as the original data.", , Iso
    KwikEnd
  End If

  StP.DropDowns("dIsoType") = Isotype
  DatSht.Select

Else

  If Not FromSquid Then
    If tB Then ThisWbk.Workbook_Open
    i = -12345
    On Error Resume Next: i = Ash.Type: On Error GoTo 0

    If i = xlWorksheet Then
      If TypeName(Selection) <> "Range" Then Range("A1").Select
    ElseIf i = xlXYScatter Then
      RequestAdd
    Else
      MsgBox "You must start from the worksheet containing" & vbLf & _
          "the data to be plotted or added.", , Iso
      ExitIsoplot
    End If

  End If

End If

DatBook$ = Awb.Name

If Not AddToPlot Then
  Set DatWbk = Awb
  Set DatSht = Ash: DatSheet$ = DatSht.Name
End If

Set Anch = DlgSht("Anch")
Set skV = Menus("StaceyKramers").Cells

On Error GoTo Tested ' Crash if "Break on all errors" option is checked"
Error 9999
Tested: On Error GoTo 0

RefChord = False: WasPlat = False

If IsoCalc Then
  On Error Resume Next: App.Calculation = xlCalculationManual: On Error GoTo 0
End If

If NIM(Itype) Then Isotype = Itype
If NIM(Inver) Then Inverse = Inver
Normal = Not Inverse

If NIM(DimThree) Then
  Dim3 = DimThree
  If Dim3 And NIM(LinDim) Then Linear3D = LinDim: Planar3D = Not Linear3D
End If

Do
  If Not GotStatus And Not AddToPlot Then GetStorStatus GotStatus, NIM(Itype)
  If Not FromSquid Then Setup NIM(Itype)
Loop Until Not Canceled Or FromSquid

Err = 0
Randomize Timer

If AgeExtract Then
  ZirconAgeExtractor True

ElseIf YoungestDetrital Then
  Detrital True

ElseIf Stacked Or StackedUseries Then

  If Not StackedUseries And ndCols <> 2 Then _
    MsgBox "The " & qq & "Ages of Stacked Beds" & qq & " routine requires " & _
    "2 data-columns (age and age-error)." & viv$ & "You entered" & Str(ndCols) & ".", _
    vbOKOnly, Iso: ExitIsoplot

  SetupBracket True, StackedUseries

ElseIf DoMix Then
  Mix True

ElseIf WtdAvPlot Then
  WeightedAverage N, wa, Nrej, Wrejected(), (N > 4)
  ShowWtdAv wa, N, Nrej, Wcap$
  InsertWtdResids DlgSht("WtdAv")

ElseIf ArPlat And Regress Then
   ArgonJ
   PlateauCalc (N), PlAge, PlAgeErr, pFirst, pLast, GasFract, PlCap$, Gas()

ElseIf ArChron Then
  PlateauChron (N), PlAge, PlAgeErr, PlInit, plInitErr, pFirst, pLast, GasFract, PlCap$, Gas()
  If ArIso And Not ArSpect Then Call ConvertArData((N))

ElseIf Regress Then

  If Dim3 Or ConcAge Or (OtherXY And WtdAvXY) Then
    If Not ConcAge Then ReDim Np(N, 5)         ' for projected pts
    If Linear3D Or ConcAge Or (OtherXY And WtdAvXY) Then ReDim pp(1, 5) ' for XY-plane intercept
  End If

  On Error GoTo CalcsDone
  Failed = False

  If N > 1 Or ConcAge Then Solve N, BadYork, Np(), xyProj(), Failed

  If ConcAge Or (OtherXY And WtdAvXY) Then _
    InsertWtdResids DlgSht("xyWtdAv"), "xyWtd" & vbLf & "Resids"

  If Failed Or ConcAge Then GoTo PlotTheData

  If UseriesPlot And Not Dim3 And N = 1 And UsType = -1 Then '&&
    SinglePointThUage
    Regress = False: GoTo PlotTheData
  End If

CalcsDone:  On Error GoTo 0
  tB = (ConcPlot And Crs(8) = 0 And Crs(9) = 0 And Not (Dim3 And Linear3D))

  If (BadYork And Not Robust) Or tB Then
    tmp$ = "No Model-" & sn$(Crs(23)) & " age solution for these data"
    If MsgBox(tmp$, vbOKCancel, Iso) <> vbOK Then ExitIsoplot
    If BadYork Then Regress = False

  ElseIf UseriesPlot And (Not Dim3 And UsType = -1 And N > 1) _
    Or (Dim3 And UsType > 0 And UsType < 4 And N = 1) _
    Or (Not Dim3 And (UsType = 1 Or UsType = 4) And N = 0) Then
    MsgBox "No calculations relevant for this plot.", , Iso
    Regress = False

  ElseIf Not Dim3 And N = 1 Then
    SinglePointThUage InpDat(1, 1), InpDat(1, 2), InpDat(1, 3), InpDat(1, 4), InpDat(1, 5)
  ElseIf OtherXY And WtdAvXY Then

  Else

    If Not (Dim3 And Linear3D) Then  '&&
      Set YorkRes = DlgSht("YorkRes"): Set ResBox = DlgSht("IsoRes")
      DetailsShown = False: PlotErrEnv = False

      If UseriesPlot Then

        With yf
          SlpEq = .Slope: SlpEqEr = .SlopeError  '&&
          IntEq = .Intercept: IntEqEr = .InterError
          If UsType = 1 Then Swap SlpEq, IntEq: Swap SlpEqEr, IntEqEr
          SinglePointThUage SlpEq, SlpEqEr / .Emult, 1, 0, 0, .MSWD, .Prob, .Emult
        End With

      Else
        If Robust Then ShowRobust

        If OtherXY And iLambda(OtherIndx) = 0 And Not Robust Then
          RegresDetailsPick
        ElseIf PbPlot And PbType = 2 Then
          ResboxProc
        ElseIf Not (Robust And OtherXY And iLambda(OtherIndx) = 0) Then
          SetupIsoRes Rbx$

          Do
            ShowBox ResBox, True
          If Not AskInfo Then Exit Do
            Caveat_Isores
          Loop

          BandBehind = IsOn(ResBox.OptionButtons("oBehind"))
          LineAgeTik = IsOn(ResBox.OptionButtons("oLines"))
          DlgSht("ConcScale").OptionButtons("oLines") = LineAgeTik
        End If

      End If

      If IsOn(ResBox.CheckBoxes("cShowRes")) And Not OtherXY _
        And (Not PbPlot Or PbType = 1) Then

        If DetailsShown Then
          Set Txb = Ash.DrawingObjects
          TxbN = Txb.Count: Set Txb = Txb(TxbN)
          If TxbN > 0 Then AddResBox Rbx$, 0, 0, Straw, Txb.Left + Txb.Width
        Else
          AddResBox Rbx$
        End If

      End If

      If DoMC Then
        Ntrials = MinMax(Thou, 30000, Val(ResBox.EditBoxes("eNtrials").Text))

        If ConcPlot And (IsOn(ResBox.CheckBoxes("cWLE")) Or IsOn(ResBox.CheckBoxes("cWLE_MC"))) Then
          Cdecay0 = (Lambda238err > 0 And Lambda235err > 0)
          Cdecay = Cdecay0
        Else
          Cdecay = False
        End If

        If ConcPlot Then
          MonteCarloConcInterErrs Ntrials, N, _
            (Anchored And IsOn(DlgSht("Anch").OptionButtons("oContinuous")))

        ElseIf ArgonPlot Then
          ArgonMonteCarlo Val(ResBox.EditBoxes("eNtrials").Text), s1$, s2$
          i = InStr(Rbx$, "Age = "): j = InStr(Rbx$, "intercept: "): k = InStr(Rbx$, "MSWD")
          Rbx$ = Left(Rbx$, i + 5) & s1$ & "  (MonteCarlo)" & vbLf & _
            "40/36 intercept: " & s2$ & Mid(Rbx$, k - 1)
          ShowBox ResBox, True

          If IsOn(ResBox.CheckBoxes("cShowRes")) Then
            Set Txb = Ash.DrawingObjects
            Set Txb = Last(Txb)
            AddResBox Rbx$, 0, 0, Aqua, Txb.Left + Txb.Width
          End If

        End If

      End If

    End If

    If Dim3 And Linear3D Then
      tmp$ = IIf(UseriesPlot, "Uiso", IIf(ConcPlot, "ConcLinType", "3dLineRes"))
      Set rBox = DlgSht(tmp$)
    Else
      Set rBox = ResBox
    End If
    InsertWtdResids rBox
  End If

End If

PlotTheData:
If ShowResidPlot Then InvokeLinearizedProb True
If Stacked Or DoMix Or Not DoPlot Then ExitIsoplot

If SplineLine Then
  If N < 3 Then MsgBox "Need 3 or more points for Spline curves": ExitIsoplot

  If Nspline Then

    For i = 1 To N - 1
      If Abs(InpDat(i, 1) - InpDat(i + 1, 1)) < IsoSmall Then
        MsgBox "Can't construct spline curve for data-pairs with identical X-values"
        ExitIsoplot
      End If

    Next i

  End If

End If

Workbooks(DatBook$).Colors = TW.Colors
If Not AddToPlot Then SymbCol = 3: SymbRow = 1 ' reserve 1st 2 cols for info on plot
NoUp
EsymbPlotted = False

If AddToPlot Then
  Set ChrtDat = Sheets(PlotDat$)
  ChrtDat.Visible = True: ChrtDat.Select
Else
  Sheets.Add:  PlotDat$ = "PlotDat"
  MakeSheet PlotDat$, ChrtDat
End If

If ProbPlot Then

  With DlgSht("ProbPlot")
    .EditBoxes("eFirst").Text = "1"
    .EditBoxes("eLast").Text = tSt(N)
  End With

  ProbDiag_click

  Do
    If Not DialogShow("ProbPlot") Then ExitIsoplot
    If AskInfo Then ShowHelp "ProbPlotHelp"
  Loop Until ProbDiagOK And Not AskInfo

End If

If WtdAvPlot Or ProbPlot Then
  WtdAverPlot N, wa, Wrejected(), Nrej, Wcap$

ElseIf CumGauss Then
  GaussCumProb N, (ndCols = 2)
  With ActiveWindow: .Zoom = 100: .Zoom = 400: End With ' To fix VBA bug
  Ach.ChartArea.Select: ActiveWindow.Zoom = True

ElseIf ArPlat Or (ArChron And ArSpect And Not WasPlat) Then
  PlotArSteps (N), pFirst, pLast, PlAge, PlCap$, Gas()
  If ArChron Then Call AddErrSymbSizeNote(True)
  Set ArChrt = Ach

  If ArChron And ArIso Then
    ConvertArData (N)
    GoTo PlotTheData
  End If

Else
  If ArChron Then AutoScale = True
  Set YorkRes = DlgSht("YorkRes"): Set ResBox = DlgSht("IsoRes")
  ConstructPlot N

  If OtherXY And Regress And IsOn(YorkRes.CheckBoxes("cShowRes")) Then
    On Error GoTo 22
    Last(DatSht.Shapes).Copy
    Ash.Paste

    With Last(Ash.Shapes)
      .Left = Ach.Axes(2).Left + Ach.Axes(1).Width - .Width - 15
      .Top = Ach.Axes(1).Top - .Height - 15
    End With

    If Not IsOn(YorkRes.CheckBoxes("cShowRes")) Then Last(DatSht.Shapes).Delete

22:    On Error GoTo 0

  End If
  EsymbPlotted = (Eellipse Or eCross Or Ebox)

  If (Dim3 And Regress) Or (ConcAge And ConcAgePlot) Then
    tmp$ = "plotting projected ellipse"

    If PlotProj And N > 0 And Not ConcAge Then ' Plot projected data pts as ellipse with next color.
      Dseries = 1 + Dseries
      ReDim Preserve DSeriesN(Dseries)
      DSeriesN(Dseries) = N
      ' If data are shown as a simple plotting symbol, use an err-ellipse
      '  for the projected data-pts; but if an error-symbol is used for the
      '  data, use the same symbol (different color) for the projected pts.
      If excSymb <> 0 Then excSymb = 0: Eellipse = True
      CreatePlotdataSource Np(), 2 + (SigLev = 1 And Not Opt.AlwaysPlot2sigma)
      ChrtDat.Visible = False

      If ColorPlot Then
        i = 1 + excClrInd

        With MenuSht

          If i >= (.Range("ColorIndex").Count - 1) _
            Or .Range("Colors").Cells(i) = "font color" Then
            tC = RGB(140, 110, 130) 'Peuce
          Else
            tC = .Range("ClrStuff").Cells(i, 3).Interior.Color
          End If

        End With

      Else
        tC = IIf(DoShape, Gray25, Black)
      End If

      PlotDataPoints xlHairline, xlContinuous, tC, False, tmp$, 1
    End If

    If ((Linear3D And PlotProj) Or (ConcAge And ConcAgePlot)) Or (OtherXY And WtdAvXY) And N > 1 Then
      ' Plot err-ellipse of XY intercept or wtd xy mean?

      If (UseriesPlot And UsType = 3) Or (ConcPlot And Not ConcConstr) Or _
       (ConcAge And xyProj(2) <> 0) Or (OtherXY And WtdAvXY) Then
        Dseries = 1 + Dseries
        ReDim Preserve DSeriesN(Dseries)
        DSeriesN(Dseries) = 1
        tB = False  ' Plot proj. ellipse at 1-sigma instead of 2/95%-conf?

        If ConcPlot And ConcAge And IsOn(DlgSht("xyWtdAv").OptionButtons("oShow1")) Then
          tB = True
        ElseIf Dim3 And SigLev = 1 And Not Opt.AlwaysPlot2sigma Then
          If (ConcPlot And Not ConcConstr) Or (UseriesPlot And UsType = 3) Then tB = True
        ElseIf OtherXY And WtdAvXY Then
          tB = True
        End If

        For j = 1 To 5
          pp(1, j) = xyProj(j)

          If (j = 2 Or j = 4) And tB And xyProj(0) <> 0 Then
            pp(1, j) = pp(1, j) / xyProj(0)
          End If

        Next j

        Eellipse = True: excSymb = 0: eCross = False: Ebox = False
        StraightLine = False: Nspline = False: Aspline = False: SplineLine = False
        CreatePlotdataSource pp(), 1, "creating projected ellipse"
        ChrtDat.Visible = False
        cca = Gray25: ccb = Opt.PlotboxClr
        If (ColorPlot And ccb <> Cyan) Or ccb = cca Then cca = Cyan
        PlotDataPoints xlMedium, xlGray50, cca, False, tmp$, 2
        EsymbPlotted = True
      End If

    End If

  End If

End If

AddErrSymbSizeNote EsymbPlotted
ExitIsoplot
End Sub

Sub MakeSheet(NewSheetName$, WhichSheet As Object)
Attribute MakeSheet.VB_ProcData.VB_Invoke_Func = " \n14"
' Create a new sheet with the name NewSheetName$ plus the next available #
'  (e.g. Flarp1, Flarp2, Flarp3...), & return WhichSheet as an object referring
'  to the new sheet.
Dim Na$, i%
Na$ = Ash.Name: i = Val(Mid$(Na$, 6))

Do
  On Error Resume Next
  Sheets(Na$).Name = NewSheetName$ & sn$(i)
If Err = 0 Then Exit Do

  If Err <> 1004 And Err <> 1005 Then
    MsgBox "Can't create sheet " & NewSheetName$ & ": " & Str(Err), vbCritical, Iso
    ExitIsoplot
  ElseIf i = 99 Then
     MsgBox "Too many sheets in this Workbook", vbCritical, Iso
     ExitIsoplot
  End If

  Err = 0
  i = i + 1 ' Sheet w. this name probably already exists - increment counter
Loop

On Error GoTo 0
NewSheetName$ = Ash.Name
If Len(DatBook$) = 0 Then DatBook$ = Awb.Name
Set WhichSheet = Awb.Sheets(NewSheetName$)
On Error Resume Next
ActiveSheet.DisplayAutomaticPageBreaks = False
End Sub

Private Sub RequestAdd() ' Invoked when Isoplot is invoked from an existing chart
Dim OK As Boolean, tB As Boolean, Ra As Object, c As Object

Xcalc
Set c = ActiveSheet ' The chart for adding points
GetOpSys
GetPlotInfo OK

If Not OK Then
  MsgBox "Not a valid Isoplot chart, or" & vbLf & _
    "source-data is missing/corrupt, or" & vbLf & _
    "source data-sheet has been renamed.", , Iso
  KwikEnd
End If

PlotIdentify

If WtdAvPlot Or ArgonStep Then
  MsgBox "Can't add data to this type of plot", , Iso
  KwikEnd
End If

GetConsts True
AddToPlot = True
Set Ra = DlgSht("Add Points")
Ra.Labels("lAddPts").Text = "Press this button to add data-points from " _
  & DatSheet$ & " to the plot"
ShowBox Ra, True
c.Select
AddPoints
Rcalc
End Sub

Sub GetDatSheetName() ' Find out the name of the source-data sheet for a chart
Attribute GetDatSheetName.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, i%, eX%

PlotName$ = Ash.Name
On Error GoTo BadSheet
s$ = Ach.SeriesCollection(1).Formula
On Error GoTo 0
eX = InStr(s$, "!"): i = eX

Do
  i = i - 1
  If i = 0 Then GoTo BadSheet
Loop Until Mid$(s$, i, 1) = ","

PlotDat$ = Mid$(s$, i + 1, eX - i - 1)
Set ChrtDat = Sheets(PlotDat$)

With ChrtDat
  Irange$ = .Cells(12, 2).Text
  DatSheet$ = .Cells(1, 2).Text
End With

Set DatSht = Sheets(DatSheet$)
Exit Sub

BadSheet: MsgBox "Not an Isoplot chart", , Iso
KwikEnd
End Sub

Sub MoveChart()  ' Move/shrink an Isoplot chart from a separate sheet to the source-data sheet
Attribute MoveChart.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, cn$, b%, i%, RtCol%, TopRow&
Dim SerC As Object, DatSht As Worksheet, r As Range

Xcalc
NoUp
GetOpSys
Set IsoChrt = Ash ' The chart-sheet

If IsoChrt.Type = xlWorksheet Then
  MsgBox "Can't attach a chart already embedded in a worksheet", , Iso
  Exit Sub
ElseIf IsoChrt.Type <> xlXYScatter Then
  MsgBox "This chart was not created by Isoplot", , Iso
  Exit Sub
End If

s$ = IsoChrt.SeriesCollection(1).Formula
b = InStr(s$, "!"): i = b

Do
  i = i - 1
Loop Until Mid$(s$, i, 1) = ","

Set ChrtDat = Sheets(Mid$(s$, i + 1, b - i - 1)) ' The hidden PlotDat sheet
DatSheet$ = ChrtDat.Cells(1, 2).Text
On Error GoTo DataSheetMissing
Set DatSht = Sheets(DatSheet$)  ' The source-data sheet
On Error GoTo 0

With ChrtDat.Range(ChrtDat.Cells(12, 2).Text) ' The range in the PlotDat sheet
  RtCol = .Column + .Columns.Count
  TopRow = .Row
End With

cn$ = "ChartToData" & String(-Mac, "2")
On Error Resume Next
IsoChrt.Shapes(cn$).Cut
On Error GoTo 0
IsoChrt.ChartArea.Copy

With DatSht
  .Select
  Set r = .Cells(TopRow, RtCol)
  .Paste
  i = .ChartObjects.Count

  If i = 0 Then _
    MsgBox "Bug in Excel -- can't attach chart as a Chart-Object", , Iso: ExitIsoplot

  With .ChartObjects(i)
    .Width = .Width / 2: .Height = .Height / 2
    .Left = r.Left: .Top = r.Top
    .Activate
  End With

End With

With Ach

  For Each SerC In .SeriesCollection

    With SerC
      If .Border.LineStyle <> xlNone Then .Border.Weight = xlHairline
      If .MarkerStyle <> xlNone Then .MarkerSize = .MarkerSize / 2
    End With

  Next

  For Each SerC In .TextBoxes

    With SerC
      .Font.Size = .Font.Size * 0.9
      .AutoSize = True
    End With

  Next

  For Each SerC In .Shapes
    If SerC.Line.Visible Then SerC.Line.Weight = 0.25
  Next

  With .ChartArea
    .Border.LineStyle = xlContinuous: .Shadow = True
  End With

  If .Shapes.Count > 0 Then RescaleOnlyShapes False, True
End With

ActiveWindow.Visible = False
Cells(TopRow, Min(1, RtCol - 1)).Select
NoAlerts
IsoChrt.Delete
ExitIsoplot

DataSheetMissing: On Error GoTo 0
MsgBox "Source-data sheet " & qq & DatSheet$ & qq + _
  " has been deleted from this workbook" & vbLf & _
  "(or possibly renamed).", , Iso
ExitIsoplot
End Sub

Sub ChangeFontSize(Obj As Object, ByVal SizeFact!)
Attribute ChangeFontSize.VB_ProcData.VB_Invoke_Func = " \n14"
With Obj.Font
  .Size = .Size * SizeFact
End With
End Sub

Sub ShowBox(d As Object, Optional CanQuit = False)
Attribute ShowBox.VB_ProcData.VB_Invoke_Func = " \n14"
Dim b As Boolean ' Show dialog-box D, after enabling screen-updating.

ViM CanQuit, False
Canceled = False: AskInfo = False ' If CancelButton is pressed & CanQuit then exit isoplot.
App.ScreenUpdating = (NoUpdate <> True)
' If screen updating is enabled & the spreadsheet is large with recalculations
'  required, get annoying screen-redraws/recalcs after each dialog box is closed;
'  but if updating is disabled, end up with blank spots on the spreadsheet when the
'  dialog box is moved (eg to see range values).
'On Error GoTo Done
b = DialogShow(d)
If Canceled And CanQuit Then ExitIsoplot
NoUp
Exit Sub

done: ExitIsoplot
End Sub

Sub KwikEnd()
Attribute KwikEnd.VB_ProcData.VB_Invoke_Func = " \n14"
SetCalc ExcelCalc
On Error Resume Next
ActiveSheet.DisplayAutomaticPageBreaks = False
EmptyClipboard
NoAlerts False
End
End Sub

Sub ExitIsoplot()  ' Restore defaults, store status, quit.
Attribute ExitIsoplot.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, Na$(), i%, u%, SortType%, Calc&, SquidInst As Variant

If DoPlot Then

  If AddToPlot Then
    StPc("cAutosc") = AutoScale0
    StPc("cCalculate") = Regress0
    AutoScale = AutoScale0: Regress = Regress0

    With ChrtDat

      If SymbCol > 0 Then
        .Cells(4, 2) = SymbCol
        .Cells(16, 2) = Max(1, SymbRow)
      End If

      .Visible = False
    End With

  ElseIf Not FromSquid Then
    PutPlotInfo
  End If

End If

If Not FromSquid Then GetStorStatus
If Mac Then s$ = ".xla"
On Error Resume Next
SquidInst = AddIns("squid" & s$).Installed
' True if installed, False if present but not installed,
' error if not present (& not installed).
On Error GoTo 0

If SquidInst = True Then
  Calc = xlCalculationManual
Else
  Calc = ExcelCalc
End If

On Error GoTo BadSheet

If DoPlot And Ash.Type <> xlWorksheet Then

  With Ach

    If HasShapes Then
      If AutoRescale And AutoSort Then
        RescaleAndOrderShapes
      ElseIf AutoRescale Then
        RescaleOnlyShapes False
      ElseIf AutoSort Then
        OrderOnly False
      ElseIf ConcPlot And Cdecay And Not BandBehind Then
        ReDim Na$(.Shapes.Count) ' Put concordia band and ticks at front

        For i = 1 To .Shapes.Count: Na$(i) = .Shapes(i).Name: Next i

        For i = 1 To Ach.Shapes.Count
          s$ = Na$(i)

          With .Shapes(s$)

            If InStr(s$, PlotDat$) Then ' find last Und delimiter

              For u = Len(s$) To 1 Step -1
                If Mid$(s$, u, 1) = Und Then Exit For
              Next u

              SortType = Val(Mid$(s$, 1 + u))

              If SortType = 2 Then
                .ZOrder msoBringToFront
              End If

            End If

          End With

        Next i

      End If

    End If

    .Deselect
  End With

End If

If Not FromSquid Then Cleanup True
StatBar
SetCalc Calc
App.DisplayStatusBar = Sbar
NoAlerts False
On Error Resume Next
ActiveSheet.DisplayAutomaticPageBreaks = False

If FromSquid Then
  EmptyClipboard
  Exit Sub
End If

BadSheet: KwikEnd
End Sub

Sub LineInd(r As Range, Optional SeriesLabel = "IsoLine", Optional RowInc = 1)
Attribute LineInd.VB_ProcData.VB_Invoke_Func = " \n14"
' For data-series that are not shape-convertible
'  ("IsoLine" -- defines a simple line, not an outline),
'  or should in some other way be uniquely identified.
' Put identifying label at bottom of the range.
Dim i%

ViM RowInc, 1
ViM SeriesLabel, "IsoLine"

With r

  For i = 1 To 2
    r(RowInc + .Rows.Count, i) = SeriesLabel
  Next i

End With

End Sub

Sub GetNuclides(u%, Optional AllNukes = False) ' Initialize nuclides-strings
Attribute GetNuclides.VB_ProcData.VB_Invoke_Func = " \n14"
Dim f$, s$, nU$(1), T%, i%

f$ = IIf(AllNukes, "Nuclides", "RadNuclides")
Static q$

ViM AllNukes, False

If Len(q) = 0 Then

  With MenuSht.Range(f)

    For i = 1 To .Rows.Count
      q = q & .Cells(i, 1)
    Next i

  End With

End If

u = 0: s = q

Do
  u = u + 1
  T = InStr(s, " ")
  ReDim Preserve Nuke(u)
  Nuke(u) = Left(s, T - 1)
  s = Mid(s, T + 1): T = InStr(s, " ")
  Nuke(u) = Nuke(u) & Left(s, T - 1)
  s = Mid(s, T + 1)
Loop Until Len(s) = 0

If AllNukes Then ReDim AllNuke(u) Else ReDim RadNuke(u)

For i = 1 To u
  If AllNukes Then AllNuke(i) = Nuke(i) Else RadNuke(i) = Nuke(i)
Next i

End Sub
