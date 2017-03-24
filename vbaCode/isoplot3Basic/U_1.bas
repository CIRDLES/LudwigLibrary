Attribute VB_Name = "U_1"
' ISOPLOT module U_1
Option Private Module
Option Explicit: Option Base 1
Dim InputRange, PsI As Object, ErI As Object, pbG As Object, xyzE As Object
Dim DoEditDiseq As Boolean
Dim zGrp As Object, zEdb As Object, zChB As Object, zLbl As Object
Dim zOptB As Object, zDropD As Object, dIsoType As Object

Sub Setup(Optional PlotSpecified As Boolean = False)
Attribute Setup.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Nrows%, Ycol%, Zcol%, qC%, Na%
Dim RangeInVals#(), NoNum As Boolean, v, nc%, s1$
Dim ct%, i%, j%, k%, M%
Dim nR%, tB As Boolean, r%, c%, ci As Range, df As Object
Dim MoreUevos As Object, Ob As Object, pL As Object, Eb As Object, Tp$, tp1$, Ra$, tmp1, tmp2
Dim tOb As Object, Top As Object, tDb As Object, AxLab As Object, Ux As Object, UxO As Object
Dim Uax As Object, La As Object, tB1 As Boolean, tb2 As Boolean, cb As Object
Dim Beh As Object, axt As Object
Static DoneThis As Boolean
ViM PlotSpecified, False
Set PsI = DlgSht("ConcScale")
Set ErI = DlgSht("ErrInp"):  Set pL = DlgSht("PlotLimits")
Set pbG = DlgSht("PbGrowth")
AssignD , StP, zEdb, zChB, zOptB, zLbl, zGrp, , , , zDropD
AssignD "AxLab", AxLab, Eb, cb
If DoneThis And Isotype > 0 And Isotype < 21 Then GoTo Setup_1
Tp$ = "K.R. Ludwig, Berkeley Geochronology Center,  "
StP.TextBoxes("AuthorDate").Text = Tp$ & Menus("RevDate").Text

With StP.TextBoxes("VersionText")
  .Text = "version " & Menus("VerNum")
  i = InStr(UCase(.Text), "ALPHA")
  If i = 0 Then i = InStr(UCase(.Text), "BETA")

  If i > 0 Then
    With .Characters(i, 5).Font
      .Bold = True: .Color = vbRed
    End With
  End If

End With

AssignIsoVars
' On the Mac, if the range-input cell is specified as xlReference, so that the user can select a range
'  whilst the dialog box is present, any subsequent attempt to press any option button or textbox in
'  this dialog box will cause an error.
zEdb("eRange").InputType = IIf(Mac, xlText, xlReference)
DoneThis = True

Setup_1:
ReDim CurvRange(1), TikRange(1)
AxisLthick = IIf(Opt.AxisThickLine, xlMedium, xlThin)

If AddToPlot Then
  If SigLev = 1 Then zOptB("o1sigma") = xlOn Else zOptB("o2sigma") = xlOn
  If AbsErrs Then zOptB("oAbsolute") = xlOn Else zOptB("oPercent") = xlOn
Else
  ConcConstr = False
End If

If TypeName(Selection) = "Range" Then
  RangeCheck nR, nc, Na
  If nR > 0 And ((nc > 1 And nc < 7) Or nc = 8 Or nc = 9) Then _
    zEdb("eRange").Text = rStr(Selection.Address, "$")
End If

If Not AddToPlot And Not PlotSpecified Then
  WhatKindOfData Na, nR, zOptB, zDropD
  If Isotype = 0 And nc = 3 Then Isotype = 18
End If

Start:
Cdecay0 = (Lambda235err > 0 Or Lambda238err > 0)
' Find out if there is a header row to the data which specifies the type
'  of data/plot/isochron & the type & sigma-level of errors.
Canceled = False
PreProcessIsoPlotSetup PlotSpecified

Do

  Do
    NumericPrefs = False: GraphicsPrefs = False
    ShowBox StP, True

    If NumericPrefs Then
      EditConsts
    ElseIf GraphicsPrefs Then
      LoadUserForm Graphics
      Graphics.Show
    ElseIf Canceled Then
      PostProcessIsoPlotSetup
      ExitIsoplot
    End If

  Loop Until Not (NumericPrefs Or GraphicsPrefs)

  Cdecay0 = (Lambda235err > 0 Or Lambda238err > 0)
  PostProcessIsoPlotSetup
  PlotIdentify
  If WtdAvPlot Or CumGauss Or ProbPlot Then Regress = 0: Anchored = 0
  Ra$ = zEdb("eRange").Text
  Irange$ = Ra$
  If Len(Trim(Irange$)) = 0 Then GoTo RngCk
  On Error GoTo BadRng

  If Left(Ra, 1) = "[" Then
    i = InStr(Ra, "]")
    If i = 0 Then GoTo BadRng
    Tp = Mid(Ra, 2, i - 2)
    Workbooks(Left(Tp, i)).Activate
    Ra = Mid(Ra, i + 1)
    i = InStr(Ra, "!")

    If i > 0 Then
      Tp = Left(Ra, i - 1)
      Sheets(Tp).Activate
      Ra = Mid(Ra, i + 1)
    End If

  End If

  Range(Ra$).Select
  GoTo RngCk

BadRng: On Error GoTo 0
  MsgBox "Invalid data-range", , Iso
  GoTo Start

RngCk: On Error GoTo 0
  RangeCheck nR, nc, Na
  Na = 0

  Do
    i = InStr(Ra$, ","): k = Len(Ra$)
    j = IIf(i > 1, i - 1, k)
    Na = 1 + Na
    NoPts = (LTrim(Ra$) = "")
    If Not NoPts Then Set RangeIn(Na) = Range(Left$(Ra$, j))
  If i < 2 Or Na = 9 Or i = k Then Exit Do
    Ra$ = Mid$(Ra$, 1 + i)
  Loop

  If (ConcPlot Or PbGrowth) And DoPlot Then Ncurves = 1

  If PbGrowth Then

    With pbG

      Do
        ShowBox pbG, True
      If .EditBoxes("eNewT0").Text = .EditBoxes("lT0").Text _
          Or Len(LTrim(.EditBoxes("eNewT0").Text)) = 0 Then Exit Do
        IsSKatAge
      Loop

      CalcPbgrowthParams
      PbTicks = IsOn(.CheckBoxes("cTix"))
      PbTickLabels = IsOn(.CheckBoxes("cLabels"))
    End With

  End If

  If AddToPlot Then

    If OtherXY Then

      With Sheets(PlotName$)

        With .Axes(xlCategory)
          If .HasTitle Then AxX$ = .AxisTitle.Text Else AxX$ = ""
        End With

        With .Axes(xlValue)
          If .HasTitle Then AxY$ = .AxisTitle.Text Else AxY$ = ""
        End With

      End With

    Else
      AxX$ = Menus("AxXn").Cells(Isotype): AxY$ = Menus("AxYn").Cells(Isotype)
    End If

  ElseIf Not WtdAvPlot And Not AgeExtract And Not YoungestDetrital Then
    i = Isotype
    ' if a Concordia Age plot, use Concordia axis-names ~!

    If (OtherXY And Not (WtdAvXY And Not DoPlot)) Or CumGauss Then

      Do
        N = nR
        ShowBox AxLab, True

        If CumGauss And IsOn(cb("cInclHist")) Then
          DoShape = IsOn(cb("cFilledBIns"))

          If IsOn(cb("cAutoBins")) Then
            Nbins = MinMax(-9999, 9999, Val(Eb("eBinNumWidth").Text))
            If Nbins > 2 And Nbins < 1201 Then Exit Do
            Tp$ = "Number of histogram cells must be between 3 and 1200"
            If MsgBox(Tp$, vbOKCancel, Iso) = vbCancel Then ExitIsoplot

          Else
            BinWidth = EdBoxVal(Eb("eBinNumWidth"))
            BinStart = EdBoxVal(Eb("eBinStart"))
            If BinWidth > 0 Then Exit Do
            If MsgBox("Bin width must be >0", vbOKCancel, Iso) = vbCancel Then ExitIsoplot
          End If

        Else

          Exit Do
        End If

      Loop

      AxX$ = Eb("eXlabel").Text: AxY$ = Eb("eYlabel").Text
      If Dim3 Then AxZ$ = Eb("eZlabel").Text

      If IsNumeric(Eb("eLambda").Text) And Not CumGauss And Not ProbPlot Then
        j = StP.DropDowns("dIsoType").Value
        iLambda(j) = EdBoxVal(Eb("eLambda"))
      End If

    ElseIf Normal Then
      AxX$ = Menus("AxXn").Cells(i): AxY$ = Menus("AxYn").Cells(i)
      If Dim3 Then AxZ$ = Menus("AxZn").Cells(i)
    Else

      AxX$ = Menus("AxXi").Cells(i): AxY$ = Menus("AxYi").Cells(i)
      If Dim3 Then AxZ$ = Menus("AxZi").Cells(i)
    End If

  End If

  If NoPts Then AutoScale = False: GoTo Ptype

  If RowWise Then ' That is, multiple areas select different data-rows with same #cols
    Nrows = nR: ndCols = RangeIn(1).Columns.Count
  Else            ' That is, multiple areas select different data-columns with same #rows
    Nrows = Min(32766, RangeIn(1).Rows.Count): ndCols = nc
    Nrows = Min(Nrows, ActiveCell.SpecialCells(xlLastCell).Row - RangeIn(1).Row + 1)
    ' In case user selected a whole column
    RangeIn(1).Select ' To restore active cell position
  End If

  If Dim3 Then
    If ndCols = 3 Then Ycol = 2: Zcol = 3 Else Ycol = 3: Zcol = 5
  ElseIf WtdAvPlot Or CumGauss Or ProbPlot Or AgeExtract Or YoungestDetrital Then
    j = 2 '+ CumGauss

    If ndCols > j Then
      If WtdAvPlot Then
        Tp$ = "Weighted-averages"
      ElseIf CumGauss Then
        Tp$ = "Probability density plots"
      ElseIf ProbPlot Then
        Tp$ = "Probability plots"
      ElseIf AgeExtract Then
        Tp$ = "Age extractions"
      ElseIf YoungestDetrital Then
        Tp$ = "Youngest detrital zircons"
      End If

      Tp$ = Tp$ & " require only" & Str(j) & " data columns" & vbLf & "(values and errors)"
      MsgBox Tp$, , Iso: GoTo Start
    End If

  ElseIf ArChron And (ndCols < (6 + Inverse) Or ndCols > 6) Then
    tp1 = IIf(Inverse, "39/40-36/40", "39/36-40/36")
    tp1 = tp1 & " plateau-isochrons require "
    tp1$ = tp1 & IIf(Inverse, "5 (or 6, if error-correl. is included)", "6")
    Tp$ = tp1 & " data-columns."
    MsgBox Tp$, , Iso: GoTo Start

  ElseIf ArPlat And Not ArChron And ndCols <> 3 Then
    MsgBox "3 Columns (%gas, Age, Error) required for Ar-Ar age spectrum", , Iso: GoTo Start

  ElseIf Stacked And ndCols <> 2 Then
    MsgBox "The Stacked Ages routine requires 2 data columns " & _
      "(ages and errors).", , Iso: GoTo Start

  ElseIf StackedUseries And (ndCols < 4 Or ndCols > 5) Then
    MsgBox "Either 4 (230/238,err,234/238,err) or 5 (...,err correl) data columns." & _
      vbLf & "required for U-series data from stacked samples.", , Iso: GoTo Start

  ElseIf ndCols > 5 And Not ArChron Then
    Tp$ = "X-Y data-range must include 2, 4, or 5 columns"

    If (ConcPlot Or UseriesPlot) And Not Regress Then
      Tp$ = Tp$ & vbLf & "(check the CALCULATE box to see 3D regression options)"
    End If

    MsgBox Tp$, , Iso: GoTo Start

  ElseIf ndCols = 2 Then
    Ycol = 2 ' X-Y data, Y in 2nd col

  ElseIf ArChron Then
    Ycol = 4

  Else
    Ycol = 3 '  "    " ,    " 3rd  "
  End If

  If UseriesPlot And Nrows < 2 Then Dim3 = False ' Single-point ThU age

  If CumGauss Or ConcAge Or (ConcPlot And Anchored) Or (PbPlot And PbGrowth And Not Regress) Then
    i = 1 ' - CumGauss
  ElseIf ArgonStep Then
    i = 4
  'ElseIf UseriesPlot then And Not Dim3 Then
  '  i = 3
  ElseIf UseriesPlot Then
    If Regress Then i = 2 Else i = 1 - DoPlot
  ElseIf ProbPlot Or AgeExtract Or YoungestDetrital Then
    i = 6
  Else
    i = 2
  End If

  If Dim3 And Not Linear3D Then i = i + 1
  tB = (PbPlot And PbGrowth And Not Regress)
  If Not Regress And Not ProbPlot And i > 1 Then i = 1

  If Nrows < 1 Or (Nrows < i And ((Regress Or tB) And Not (UseriesPlot And Regress))) Then
    Tp$ = IIf(i > 1, "s", "")
    MsgBox "Data range must include at least " & sn$(i) & " row" & Tp$, vbExclamation, Iso
    GoTo Start
  End If

  ReDim RangeInVals(Nrows, ndCols), MTrow(Nrows), DatClr(Nrows)
  TopRow = 0:  ct = 0: RightCol = 0: Xcolumn = RangeIn(1).Column
  RowColWise RangeInVals(), Nrows, Na, nc, ct
  N = ct - Anchored
  If N < 2 And AutoScale And DoPlot Then AutoScale = False
  tB = PbGrowth And PbPlot And Not Regress
  i = 1 - (WtdAvPlot Or (Regress And Not ConcAge And Not (UseriesPlot And Not Dim3)))
  If tB Or (UseriesPlot And Not DoPlot And Not Dim3) Then i = 1

  If N < i Then
    Tp$ = "Need at least" & Str(i) & " data " & Plural("point", i)
    MsgBox Tp$, , Iso
    Canceled = True: Exit Sub

  ElseIf Robust And N < 5 Then
    MsgBox "Need at least 5 points for a robust regression", , Iso: ExitIsoplot
  End If

  ArPlat = (ArgonStep And ndCols = 3)
  qC = 5 - 4 * Dim3 - 1 * ArChron
  If (N - Anchored) < 2 And Not ConcAge And Not (UseriesPlot And Not Dim3) Then Regress = False:
  If N < 2 Then AutoScale = False

  ForceFill = (N > 250 And Not DoShape And DoPlot And (Ebox Or Eellipse))

  If ForceFill Then
    DoShape = True
    tmp1 = IIf(Ebox, "box", "ellipse")
    MsgBox "The upper limit on the number of unfilled error-" & tmp1 & " symbols is 250," _
      & vbLf & "so Isoplot will used filled symbols set to 100% transparency." _
      & vbLf & vbLf & "Remember to Rescale after any changes to the plot size or scale." _
      , , "Isoplot"
  End If

  ReDim InpDat(N, qC)

  If ArgonStep And ndCols = 3 Then
    ReDim BoldedData(N)
    NforcedSteps = 0: ForcedPlateau = False
  End If

  For i = 1 To N + Anchored

    For j = 1 To ndCols

      If Regress Or WtdAvPlot Then   ' Make sure all errors are nonzero
        If (ArgonStep And (j = 3 Or _
         ((ndCols = 5 Or ndCols = 6) And j = 3 Or j = 5))) Or _
         (ndCols > 1 And j = 2) Or (ndCols > 3 And j = 4) Or (Dim3 And j = 6) Then

          If RangeInVals(i, j) = 0 Or (Not AbsErrs And RangeInVals(i, j - 1) = 0) Then _
            MsgBox "Errors must be nonzero", , Iso: GoTo Start
        End If

      End If

      If (j > (4 - 2 * Dim3) And Not ArChron) Or (ArChron And j = 6) Then

        If Abs(RangeInVals(i, j)) > 1 Then
           MsgBox "Error correlations cannot exceed 1", vbExclamation, Iso
           GoTo Start
        ElseIf Normal And ConcPlot And Not Dim3 And RangeInVals(i, 5) <= 0 Then
           Tp$ = "Error correlations are ALWAYS >0 for conventional concordia"
           MsgBox Tp$, vbExclamation, Iso
           GoTo Start
        End If

      End If

    Next j

  Next i

Ptype:

  If Anchored Or RefChord Then
    AnchorProc

    Do
      ShowBox Anch, True
      AnchorProc
      If IsNumeric(Anch.EditBoxes("eAge").Text) Then ' Protect against blank entry

        If AgeAnchor Then
          If Normal And (AnchorAge >= 0 And AnchorAge < 6000) Then Exit Do
          If Inverse And AnchorAge <= 6000 And (AnchorAge - Abs(AnchorErr)) >= 0.1 Then Exit Do
        ElseIf PbAnchor Then
          If Anchor76 > 0.04 And Anchor76 < 2 Then Exit Do
        ElseIf RefChord Then
          Anchored = False: N = N - 1
        End If
      End If

      Tp = ""

      If RefChord Then
        tB1 = IIf(Inverse, AnchorT1 > 0 And AnchorT2 > 0, AnchorT1 >= 0 And AnchorT2 >= 0)
        tb2 = (AnchorT1 < 9000 And AnchorT2 < 9000)
        If tB1 And tb2 Then Exit Do
        Tp = "Cannot plot zero or negative ages on a Tera-Wasserburg plot"
      ElseIf AgeAnchor And Inverse And (AnchorAge - Abs(AnchorErr)) <= 0 Then
        Tp = "Unreasonable Anchor value" & vbLf & _
          "(age" & pm & "err must be >0.1 for Tera-Wasserburg concordia)"
      End If

      If Tp <> "" Then MsgBox Tp, , Iso Else Exit Do
    Loop

  ElseIf UseriesPlot Then 'And (N > 1 Or Not Dim3) Then
    Set Ux = DlgSht("3dU"): Set Uax = Menus("UseriesAxes").Cells: Set UxO = Ux.OptionButtons

    For i = 1 To 6
      Tp$ = ""

      For j = 1 To 3
         Tp$ = Tp$ & Uax(i, j).Text & "  "
      Next j

      UxO(i).Text = Tp$
    Next i

    For i = 1 To 3: UxO(i).Enabled = Dim3: Next i
    For i = 4 To 6: UxO(i).Enabled = Not Dim3: Next i
    Ux.DrawingObjects("rec2dBox").Enabled = Not Dim3

    If UsType = 0 Then
      UsType = IIf(Dim3, 1, -1)
    End If

    If UsType > 0 And UsType < 4 Then UxO(UsType) = xlOn

    If Not Dim3 Then
      If IsOn(UxO(1)) Then UxO(5) = xlOn
      If IsOn(UxO(2)) Or IsOn(UxO(3)) Then UxO(4) = xlOn
      If UsType = 4 Then UxO("oType4") = xlOn
    End If

    Proc3dU
    Ux.Labels("lIsoExp").Visible = DoPlot

    If Dim3 Then
      Ux.Labels("lIsoExp").Text = "(isochron plotted is the X-Y projection)"
      'Ux.DialogFrame.Text = "U-Series Plot/Isochron"
      If IsOn(UxO(4)) Then UxO(1) = xlOn
      If IsOn(UxO(5)) Then UxO(1) = xlOn
    ElseIf UsType = -1 Then
       Ux.Labels("lIsoExp").Text = "(3D required for isochron)"
       'Ux.DialogFrame.Text = "U-Series Evolution Plot"
       UxO(4) = xlOn
    End If

    If N > 1 Then

      Do
        ShowBox Ux, True
        If AskInfo Then Call ShowHelp("3dUhelp")
      Loop Until Not AskInfo

    End If

    If Dim3 Then
      i = UsType

      If i = -1 Then
        i = 4
      ElseIf i = 4 Then
        i = 6
      End If

      AxX$ = Uax(i, 1).Text: AxY$ = Uax(i, 2).Text: AxZ$ = Uax(i, 3).Text
      PlotProj = IsOn(Ux.CheckBoxes("cPlotProj"))

    ElseIf UsType = -1 Then
      AxX$ = Uax(4, 1).Text: AxY$ = Uax(4, 2).Text

    ElseIf UsType = 4 Then
      AxX$ = Uax(6, 1).Text: AxY$ = Uax(6, 2).Text

    End If

    If UsType <> 3 Then PlotProj = False

    If DoPlot And uEvoCurve Then
      LoadUserForm UevoT
      Ncurves = 1
      ReDim Ugamma0(1)

      With UevoT ' Order of control assignment is important!
        .cMultCurves = Menus("UevoMultCurves")
        .cLabelCurves = Menus("UevoLabelCurves")
        .cLabelTicks = Menus("UevoLabelTicks")
        .oInside = Menus("UevoInside")
        .oOutside = Not .oInside:
        .eGamma0_1 = UeT(Menus("UevoGamma0_1"))
        .eMaxAge = UeT(Menus("UevoMaxT"))
        .eLabelKa = UeT(Menus("UevoLabelKa"))
        .eTickInterval = UeT(Menus("UevoTickInterval"))
        .oIsochrons = Menus("UevoIsochrons")
        .oAgeTicks = Menus("UevoAgeTicks")
        .oNeither = Not (.oIsochrons Or .oAgeTicks)
      End With

      UevoProcT
      Ugamma0(1) = uFirstGamma0

      Do
        UevoT.Show
        tp1$ = "Maximum age cannot ": Tp$ = ""

        If MaxAge > 2000 Then
          Tp$ = tp1$ & "exceed 2000 ka"
        ElseIf MaxAge < 1 Then
          Tp$ = tp1$ & "be less than 1 ka"
        ElseIf uFirstGamma0 <= 0 Or uFirstGamma0 > 100 Then
          Tp$ = "Initial 234/238 must be between 0 and 100"
        ElseIf CurvTikInter <= 0 Then
          Tp$ = "You must enter an age-interval for the isochrons or age-ticks"
        ElseIf (uUseTiks Or uPlotIsochrons) And MaxAge / CurvTikInter > 127 Then
          Tp$ = "Age ticks are too closely-spaced - " & vbLf & _
                "please increase the age-tick interval"
          UevoT.eTickInterval.Text = tSt(MaxAge / Hun)
        ElseIf LabelUcurves And uEvoCurvLabelAge = 0 Then
          Tp$ = "You must specify at what age along the" & vbLf _
            & "evolution curves to label the curves."
        End If

        If Tp$ <> "" Then MsgBox Tp$, , Iso
      If Tp$ = "" Then Exit Do
        LoadUserForm UevoT
      Loop

      If uMultipleEvos Then
        ShowBox DlgSht("MoreUevos"), True
        Set MoreUevos = DlgSht("MoreUevos").EditBoxes
        j = MoreUevos.Count
        ReDim Preserve Ugamma0(1 + j)

        For i = 1 To j
          v = EdBoxVal(MoreUevos(i))

          If v > 0 And v < 100 Then
            Ncurves = 1 + Ncurves
            Ugamma0(Ncurves) = v
          End If

        Next i

        ReDim CurvRange(Ncurves), TikRange(Ncurves)
      End If

    End If

  End If

  If ConcPlot And Inverse And Dim3 And Linear3D And Regress Then
    ConcLinType_click

    Do
      ShowBox DlgSht("ConcLinType"), True
    If Not AskInfo Then Exit Do
      ShowHelp "ConcLinTypeHelp"
    Loop

    ConcLinType_click
  End If

  tB1 = (AutoScale And Not Regress And ConcPlot)
  tb2 = (Not AutoScale And Not WtdAvPlot And Not AgeExtract And Not YoungestDetrital _
         And Not CumGauss And Not ProbPlot And Not ArgonStep)

  If (tB1 Or tb2) And DoPlot And Not AddToPlot Then

    If ConcPlot Then
      tb2 = Not (tB1 And Cdecay0)

      If tb2 Then
        Set Beh = PsI.OptionButtons
        PsI.CheckBoxes("cWLE").Enabled = Cdecay0
        PsI.GroupBoxes("gWLE").Enabled = Cdecay0
        Beh("oLines") = LineAgeTik
        Beh("oCircles").Caption = ConcSymbInfo
        DlgSht("IsoRes").OptionButtons("oLines") = LineAgeTik
        WithDcErrs2_click
        XYlim = False
        'PsI.CheckBoxes("cAutoscale") = xlOff
        'Autoscale2_click

        Do
          ShowBox PsI, True
        If XYlim Or AutoScale Then Exit Do
          ProcConcScale NoNum
        Loop Until Not NoNum Or tB1

        If DoShape And Cdecay0 Then
          If IsOff(Beh("oBehind")) And IsOff(Beh("oFront")) Then Beh("oBehind") = xlOn
          BandBehind = IsOn(Beh("oBehind"))
        End If

        CurvWithDce = IsOn(PsI.CheckBoxes("cWLE"))
        LineAgeTik = IsOn(Beh("oLines")) And Not CurvWithDce
        DlgSht("IsoRes").OptionButtons("oLines") = LineAgeTik

      ElseIf Cdecay0 Then
        On Error GoTo 0
        LoadUserForm DCerrsOnly
        DCerrsOnly.Show
        If Canceled Then ExitIsoplot
        DlgSht("IsoRes").CheckBoxes("cWLE") = Cdecay
      End If

    End If

    If Not ConcPlot Or XYlim Then
      Tp$ = IIf(Len(AxX$) < 11, "    ", "")

      If ConcPlot And InvertPlotType Then

        If Inverse Then
          AxX$ = Menus("AxN")(1): AxY$ = Menus("AxYn")(1)
        Else
          AxX$ = Menus("AxXi")(1): AxY$ = Menus("AxYi")(1)
        End If

      End If

      With pL
        .Labels("lMinX").Text = Tp$ & "Min. " & AxX$
        .Labels("lMaxX").Text = Tp$ & "Max. " & AxX$
        .Labels("lMinY").Text = Tp$ & "Min. " & AxY$
        .Labels("lMaxY").Text = Tp$ & "Max. " & AxY$
      End With

      Do
        ShowBox pL, True
      If AutoScale Then Exit Do
        Set tDb = pL.EditBoxes
        MinX = EdBoxVal(tDb("eMinX")): MaxX = EdBoxVal(tDb("eMaxX"))
        MinY = EdBoxVal(tDb("eMinY")): MaxY = EdBoxVal(tDb("eMaxY"))
        If MinX > MaxX Then Swap MinX, MaxX
        If MinY > MaxY Then Swap MinY, MaxY
      Loop Until MinX < MaxX And MinY < MaxY

      If ConcPlot And XYlim Then

        If (Inverse And Not InvertPlotType) Or (Normal And InvertPlotType) Then
          If InvertPlotType Then Inverse = Not Inverse: Normal = Not Inverse
          If MinY <= 0 Then MinAge = 0 Else MinAge = ConcYage(MinY)
          MinAge = Max(MinAge, ConcXage(MaxX))
          If MinX <= 0 Then MaxAge = 6000 Else MaxAge = ConcXage(MinX)
          MaxAge = Min(MaxAge, ConcYage(MaxY))
        Else
          If InvertPlotType Then Inverse = Not Inverse: Normal = Not Inverse
          MinAge = Max(ConcXage(MinX), ConcYage(MinY))
          MaxAge = Min(ConcXage(MaxX), ConcYage(MaxY))
        End If

        If (MaxAge - MinAge) > 10 And (MaxAge - MinAge) <= 0.1 Then MaxAge = 1.1 * MinAge
        If InvertPlotType Then Inverse = Not Inverse: Normal = Not Inverse
      End If

    End If

  End If

  If NoPts Then N = 0: Exit Sub

  If Not WtdAvPlot And Not CumGauss And Not AgeExtract _
   And Not YoungestDetrital And Not ProbPlot Then

    If ArPlat And ndCols <> 3 Then
      Tp$ = "3 columns (%gas, Age, age-err) required for Ar-Ar step heating"
      MsgBox Tp$, , Iso: GoTo Start

    ElseIf ArChron And Not ArPlat And ((Normal And ndCols <> 6) Or (Inverse And ndCols <> 5 And ndCols <> 6)) Then
      tp1$ = IIf(Normal, "5 or 6", "6")
      Tp$ = "3 columns required in data range" & vbLf & "for simple Argon-Argon step-heating,"
      Tp$ = Tp$ & viv$ & tp1$ & " columns for Argon-Argon Step-Heating+Isochron PlateauChron"
      MsgBox Tp$, , Iso: GoTo Start

    ElseIf Dim3 Then

      If Not (ndCols = 3 Or ndCols = 6 Or ndCols = 9) Then
        Tp$ = "X-Y-Z data-range must include 3, 6, or 9 columns "
        MsgBox Tp$, , Iso
        GoTo Start
      End If

      Set xyzE = DlgSht("xyzErrs")
      Tp$ = IIf(AbsErrs, "Absolute e", "%E")
      xyzE.GroupBoxes("gErrs").Text = Tp$ & "rrors"

      For i = 1 To 6
        tB = (ndCols = 3 Or i > 3)
        xyzE.EditBoxes(i).Enabled = tB: xyzE.Labels(i).Enabled = tB
      Next i

    ElseIf Not Robust And Not Stacked And Not DoMix And Not ArgonStep Then

      With ErI

        For i = 1 To 3
          .EditBoxes(i).Enabled = True: .Labels(i).Enabled = True
        Next i

        Tp$ = IIf(AbsErrs, " errors", " %errs")
        .Labels("lXerrs").Text = AxX$ & Tp$: .Labels("lYerrs").Text = AxY$ & Tp$
        Tp$ = "Errors and error correlations"
        If N > 1 Then Tp$ = Tp$ & " for all points"
        .DialogFrame.Text = Tp$
      End With

    End If

  End If

  If Dim3 Then
  ' -------------
  ElseIf Stacked Then

    If StackedUseries And ndCols < 4 Or ndCols > 5 Then
      MsgBox "Data range must include either 4 or 5 columns", , Iso
      GoTo Start
    ElseIf ndCols <> 2 Then
      MsgBox "The " & qq & "Ages of Stacked Beds" & qq & " routine requires " & _
        "2 data-columns (age and age-error)." & viv$ & "You entered" & Str(ndCols) & ".", _
        vbOKOnly, Iso
      GoTo Start
    End If

  ElseIf Not ArgonStep Then

    If ndCols < (2 + ProbPlot + CumGauss) Then
      MsgBox "Data range must include 2 columns (values and errors)", , Iso
      GoTo Start

    ElseIf ndCols <> 2 And ndCols <> 4 And ndCols <> 5 And Not CumGauss And (Stacked Or DoMix) Then
      MsgBox "Data range must include 2 columns only (ages and errors)", , Iso
      GoTo Start

    ElseIf ndCols = 4 Then

      With ErI
        .EditBoxes("eXerrs").Enabled = False: .EditBoxes("eYerrs").Enabled = False
        .Labels("lXerrs").Enabled = False:    .Labels("lYerrs").Enabled = False
        .EditBoxes("eXerrs").Text = "":       .EditBoxes("eYerrs").Text = ""
      End With

    ElseIf ProbPlot Then

      If ndCols <> 1 And ndCols <> 2 Then
        MsgBox "Data range must include 1 or 2 columns", , Iso
        GoTo Start
      End If

    ElseIf ndCols <> 2 And ndCols <> 5 And Not CumGauss Then
      MsgBox "X-Y data range must include 2, 4, or 5 columns", , Iso
      GoTo Start
    End If

  End If

  If ConcPlot Then
    If InvertPlotType And AddToPlot And Inverse Then ErrCorrsReqd = True
    If Normal And Not AddToPlot Then ErrCorrsReqd = True
  ElseIf Normal And (PbPlot Or ArgonPlot) Then
    ErrCorrsReqd = True
  End If

  If Robust Or Stacked Or DoMix Or WtdAvPlot Or CumGauss Then ErrCorrsReqd = False
  If ArPlat And Not ArChron Then ErrCorrsReqd = False

  If ndCols < (4 - 2 * Dim3 - ErrCorrsReqd) And Not _
    (WtdAvPlot Or CumGauss Or ProbPlot Or ArPlat Or WtdAvPlot _
     Or AgeExtract Or YoungestDetrital Or Stacked Or DoMix) Then

    If (Regress And Not Robust) Or ((eCross Or Ebox) And ndCols = 2) Or Eellipse Then

      Do
        tB = False

        If Dim3 Then
          ShowBox xyzE, True
        Else
          ShowBox ErI, True
        End If

        ProcErr tB
      Loop Until Not tB

    End If

    If UseriesPlot And UsType <> -1 And Regress And Not AbsErrs Then
      tB = (Xerror / SigLev > 1.5) And Rhos(1) = 0

      If Dim3 Then tB = tB And Yerror / SigLev > 1.5 And Zerror / SigLev > 1.5 _
        And Rhos(1) = 0 And Rhos(2) = 0
      If tB Then MsgBox "Ignoring error correlations for alpha-spectrometric" & vbLf _
       & "data will result in inaccurate isochron ages.", , Iso
    End If

  End If

  For i = 1 To N + Anchored
    InpDat(i, 1) = RangeInVals(i, 1)

    If ArPlat Then
      r = RangeIn(1).Row - 1 + i - HeaderRow
      tB = False

      For j = 1 To ndCols
        c = RangeIn(1).Column - 1 + j
        If Cells(r, c).Font.Bold Then tB = True
      Next j

      BoldedData(i) = tB
      If tB Then NforcedSteps = 1 + NforcedSteps
    End If

    If Not WtdAvPlot And Not AgeExtract And Not YoungestDetrital And Not CumGauss _
      And Not ProbPlot Then InpDat(i, 3) = RangeInVals(i, Ycol)
    If Dim3 Then InpDat(i, 5) = RangeInVals(i, Zcol)

    If ArChron Then
      For j = 2 To ndCols: InpDat(i, j) = RangeInVals(i, j): Next j
    End If

    If Not WtdAvPlot And Not AgeExtract And Not YoungestDetrital _
     And Not CumGauss And Not ProbPlot And Not ArPlat And _
      Not Stacked And Not DoMix And ndCols = 2 - Dim3 Then
       'AbsErrs = False
       j = 2 - ArChron
       InpDat(i, j) = Xerror: InpDat(i, j + 2) = Yerror
       If Dim3 Then InpDat(i, 6) = Zerror

    ElseIf Not Stacked And Not DoMix And Not ArChron Then
       If ndCols > 1 Then InpDat(i, 2) = RangeInVals(i, 2)

       If Not WtdAvPlot And Not AgeExtract And Not YoungestDetrital _
        And Not CumGauss And Not ArPlat And Not ProbPlot Then
         InpDat(i, 4) = RangeInVals(i, 4)
       End If

       If Dim3 Then InpDat(i, 6) = RangeInVals(i, 6)
    End If

    If Dim3 Then

      If ndCols < 7 Then
        InpDat(i, 7) = Rhos(1): InpDat(i, 8) = Rhos(2)
        InpDat(i, 9) = Rhos(3)
      Else
        InpDat(i, 7) = RangeInVals(i, 7): InpDat(i, 8) = RangeInVals(i, 8)
        InpDat(i, 9) = RangeInVals(i, 9)
      End If

    ElseIf Not WtdAvPlot And Not AgeExtract And Not YoungestDetrital _
      And Not CumGauss And Not ProbPlot And Not ArPlat And _
      Not Stacked And Not DoMix Then

      If ndCols < 5 Then

        j = 5 - 3 * ArChron
        InpDat(i, j) = Rhos(1)

      Else
        InpDat(i, 5) = RangeInVals(i, 5)

        If ArChron And ndCols > 5 Then
          InpDat(i, 6) = RangeInVals(i, 6)
        End If

        If Normal And (ConcPlot Or PbPlot) Then

          If InpDat(i, 5) <= 0 Then
            Tp$ = "conventional-concordia data"
            If PbPlot Then Tp$ = "such Pb-Pb data"
            MsgBox "Error correlations for " & Tp$ & " are ALWAYS >0", vbExclamation, Iso
            GoTo Start
          End If

        End If

      End If

    End If

    If Not AbsErrs Then
      M = 4 + ArPlat - 2 * ArChron - 2 * Dim3 + 2 * (WtdAvPlot Or CumGauss)

      For j = 2 - ArPlat - ArChron To M Step 2
        InpDat(i, j) = Abs(InpDat(i, j) / Hun * InpDat(i, j - 1))
      Next j

    End If

    If SigLev <> 1 Then

       For j = 2 - ArPlat - ArChron To _
         4 + ArPlat - ArChron - 2 * Dim3 + 2 * (WtdAvPlot Or CumGauss) _
           Step 2
            InpDat(i, j) = InpDat(i, j) / SigLev
       Next j

    End If

  Next i

  If ConcPlot And Not Dim3 And Not InvertPlotType Then ' Check to make sure Conc. data-type is correct
    ReDim ConcNorm#(N), ConcTW#(N), cc(2, N), Ni$(-1 To 0)

    For i = 1 To N

      For j = -1 To 0
        If InpDat(i, 1) > 0 And InpDat(i, 3) > 0 Then
         cc(j + 2, i) = Abs(ConcXage(InpDat(i, 1), j) / ConcYage(InpDat(i, 3), j)) - 1
        End If
      Next j

      ConcNorm(i) = cc(2, i): ConcTW(i) = cc(1, i)
    Next i

    tmp1 = Abs(iMedian(ConcNorm())): tmp2 = Abs(iMedian(ConcTW()))
    tB1 = False: tb2 = False                                               ' /
                                                                           '|
    If Normal Then                                                         '|
      If tmp2 > 0 Then tB1 = (tmp1 > 1 And tmp2 < 1) Or (tmp1 / tmp2 > 10) '| 2010/11/3 -- mod to
    ElseIf tmp1 > 0 Then                                                   '|   protect against
      tB1 = (tmp2 > 1 And tmp1 < 1) Or (tmp2 / tmp1 > 10)                  '|   div-by-zero
    End If                                                                 ' \

    If tB1 And Not InvertPlotType Then
      Ni$(0) = "Inverse": Ni$(-1) = "Normal"
      s1$ = "You specified " & qq + Ni$(Normal) & qq & " Concordia, but your data " _
        & "seems more suited to " & qq & Ni$(Inverse) & qq & ".  Do you want to switch?"
      i = MsgBox(s1$, vbYesNoCancel, Iso)

      If i = vbCancel Then
        ExitIsoplot
      ElseIf i = vbYes Then
        Normal = Not Normal: Inverse = Not Normal
      End If

      CnameAssign
    End If

  End If

  If ArPlat And NforcedSteps > 1 Then ForcedPlateau = True

  If Anchored Then
    tmp1 = 0.00000001

    If PbAnchor Then
      InpDat(N, 1) = tmp1: InpDat(N, 3) = Anchor76: InpDat(N, 2) = tmp1
      InpDat(N, 4) = AnchorErr - 0.000001 * (AnchorErr = 0)
    ElseIf AgeAnchor Then
      InpDat(N, 1) = ConcX(AnchorAge): InpDat(N, 3) = ConcY(AnchorAge)
      InpDat(N, 2) = tmp1: InpDat(N, 4) = tmp1
    End If

  End If

  If ConcPlot And InvertPlotType And Not Dim3 And DoPlot Then ' Transform data
    If AddToPlot Then Inverse = Not Inverse: Normal = Not Normal

    If False And Inverse Then   ' Check for uniformly near-zero rho's
      tB = True

      For i = 1 To N
        If Abs(InpDat(i, 5)) > 0.001 Then tB = False
      Next i

      If tB Then ' Warn that the near-zero rho's must be accurate.
        DlgSht("InvertPtype").CheckBoxes("NoShowRho").Visible = False
        ShowBox DlgSht("InvertPtype"), True
      End If

    End If

    For i = 1 To N ' transform the x- y-values, errors, and rho's
      ConcConvert InpDat(i, 1), InpDat(i, 2), InpDat(i, 3), InpDat(i, 4), InpDat(i, 5), Inverse, tB
      ' ConcConvert will display a warning if any error correlations are impossible (tB=true)
      If tB Then ExitIsoplot
    Next i

    Inverse = Not Inverse: Normal = Not Inverse ' Transformed data are OK - switch plot type
    CnameAssign ' Switch axis names
  End If

  Exit Sub
Loop

End Sub

Private Sub ProcErr(BadER As Boolean)  ' Check error values & rho's for validity
Dim d As Object, i%, BadRhos%, tmp$, j%, v#
Dim b As Boolean, De As Object
If Dim3 Then Set d = xyzE Else Set d = ErI
Set De = d.EditBoxes
'D can be either the "ErrInp" or the "XZYerrs" dialog box.
' ErrInp.De(i):  1=eXerrs, 2=eYerrs, 3=eRhos
' XYZerrs.De(i): 1=eXerrs, 2=eYerrs, 3=eZerrs, 4=eXYrhos, 5=eXZrhos, 6=eYZrhos
Xerror = EdBoxVal(De(1))
Yerror = EdBoxVal(De(2))

For i = 3 To 3 - 3 * Dim3
  v = EdBoxVal(De(i))

  If Dim3 Then

    If i = 3 Then
      Zerror = v
    Else
      Rhos(i - 3) = v
    End If

  Else
    Rhos(1) = v
  End If

Next i

For i = 1 To 1 - 2 * Dim3
  If Abs(Rhos(i)) > 1 Then BadRhos = i
Next i

b = (Xerror = 0 Or Yerror = 0)
b = IIf(Dim3, (b Or Zerror = 0) And ndCols = 3, b And ndCols = 2)
If b Then BadER = True

If BadRhos Then
  MsgBox "Error correlations cannot exceed 1", vbExclamation, Iso
  BadER = True
ElseIf ErrCorrsReqd And Rhos(1) <= 0 Then
  tmp$ = IIf(Dim3, "E", "X-Y e") & "rror correlations for this plot are ALWAYS nonzero"
  MsgBox tmp$, vbExclamation, Iso
  BadER = True
ElseIf b Then
  MsgBox "Errors of zero not allowed", , Iso
  i = -(Xerror = 0) - 2 * (Yerror = 0) = 3 * (Dim3 And Zerror = 0)
  BadER = True
End If

End Sub

Private Sub ProcConcScale(NotOk As Boolean)  ' Handle Conc.-plot scale dialog-box click
Dim s$, Pe As Object
Set Pe = PsI.EditBoxes
NotOk = True
s$ = Pe("eMinAge").Text
If Not IsNumeric(s$) Then Exit Sub
MinAge = Val(s$)
s$ = Pe("eMaxAge").Text
If Not IsNumeric(s$) Then Exit Sub
MaxAge = Val(s$)

If MinAge >= MaxAge Then
  If MinAge = MaxAge Then MaxAge = 0.95 * MaxAge: MinAge = 1.05 * MinAge
  Swap MinAge, MaxAge
  Pe("eMaxAge").Text = Pe("eMinAge").Text
  Pe("eMinAge").Text = s$
End If

If ((Inverse And Not InvertPlotType) Or (Normal And InvertPlotType)) _
  And MinAge = 0 Then MinAge = MaxAge / 10
'Cdecay0 = IsOn(PsI.CheckBoxes("cWLE"))
If (MaxAge - MinAge) > 0.2 Or XYlim Then NotOk = False
End Sub

Private Sub ProcPbGrowth() ' Handle Single-stage Pb-growth plot dialog-box click
Dim Eb As Object
Set Eb = pbG.EditBoxes
pbAlpha0 = EdBoxVal(Eb(1)):   pbBeta0 = EdBoxVal(Eb(2))
pbGamma0 = EdBoxVal(Eb(3)):  pbMu = EdBoxVal(Eb(4))
pbKappaMu = EdBoxVal(Eb(5)): pbStartAge = EdBoxVal(Eb(6))
CalcPbgrowthParams
End Sub

Function EdBoxVal(EditBox As Object)
Dim v As Variant
v = 0

With EditBox

  If .Visible And .Enabled Then
    If IsNumeric(.Text) Then v = Val(.Text)
  End If

End With

EdBoxVal = v
End Function
Private Sub IsStaceyKramers() ' Handle Single-stage Pb-growth plot dialog-box click
Dim Eb As Object, i%
Set Eb = pbG.EditBoxes

For i = 1 To 6: Eb(i).Text = skV(i): Next i

Eb(7).Text = ""
CalcPbgrowthParams
End Sub

Private Sub IsSKatAge() ' Handle Single-stage Pb-growth plot dialog-box click
Dim Eb As Object, i%, j%, v, NewStartAge
Set Eb = pbG.EditBoxes
CalcPbgrowthParams
NewStartAge = EdBoxVal(Eb(7))
NewStartAge = MinMax(0, 4600, NewStartAge)
Eb(7).Text = sn$(NewStartAge)
'Alpha=Alpha0+Mu*[Exp(Lambda238*T0)-Exp(Lambda238*T)]

For i = 1 To 3
  j = i - 1
  v = EdBoxVal(Eb(i)) + MuIsh(j) * (PbExp(j) - Exp(PbLambda(j) * NewStartAge))
  Eb(i).Text = Sp(v, -3)
Next i

Eb(6).Text = Eb(7).Text
If Not Mac Then SendKeys vbTab
End Sub
Function NFont(s As Object)
Attribute NFont.VB_ProcData.VB_Invoke_Func = " \n14"
Dim v#
If IsNumeric(s) Then v = s Else v = Val(s.Text)
NFont = MinMax(4, 40, v)
End Function

Sub FillList(c As Object, Source As Range, ListIndx, ByVal IndxAsIndex As Boolean, Optional LastItem)
Attribute FillList.VB_ProcData.VB_Invoke_Func = " \n14"
' Populate the list of a ComboBox control with a single-column range
' (Emulates the RowSource property not supported by Macs)
Dim i%, FoundIndx As Boolean, LastInd%
LastInd = IIf(IM(LastItem), Source.Rows.Count, LastItem)

For i = 1 To LastInd
  c.AddItem Source(i)

  If Not IndxAsIndex Then

    If Not FoundIndx And Source(i) = ListIndx Then
      c.ListIndex = i: FoundIndx = True
    End If

  End If

Next i

If IndxAsIndex Then
  c.ListIndex = ListIndx
ElseIf Not FoundIndx Then
  c.ListIndex = -1
End If

End Sub
Function Match(ByVal v, r As Range) ' find position of V in range R
Attribute Match.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i&

For i = 1 To r.Rows.Count
  If r(i, 1) = v Then Match = i: Exit Function
Next i

Match = 0
End Function
Sub NumericPrefsPushed()
Attribute NumericPrefsPushed.VB_ProcData.VB_Invoke_Func = " \n14"
NumericPrefs = True: Canceled = False: AskInfo = False
End Sub
Sub GraphicsPrefPushed()
Attribute GraphicsPrefPushed.VB_ProcData.VB_Invoke_Func = " \n14"
GraphicsPrefs = True: Canceled = False: AskInfo = False
End Sub
Private Sub EditConsts() ' Setup and show the "Consts" user form.
Dim i%, d As Range, s1$, s2$, MAP As Range, ML As Range, Lambdas As Range, dc As Range
Set ML = Menus("DecayConstPerrs"): Set d = Menus("DecayConsts").Cells
Set Lambdas = Menus("Lambdas")
Set dc = Menus("DecayConsts")
Set MAP = Menus("ModelAgeParams").Cells

Do
  LoadUserForm Consts
  Consts.Show
Loop Until Not BadConst

AssignDirectLambdaNames
NoUp
StoreMenuSheet

For i = 1 To Lambdas.Count ' iLambda is indexed to plot-type, and is in 1/myr

  If LambdaRef(i) <> 0 Then
    iLambda(i) = dc(LambdaRef(i)) * Million
    Lambdas(i) = iLambda(i)
  End If

Next i

End Sub

Sub AssignDirectLambdaNames()
Lambda232 = Menus("Lambda232") * Million
Lambda234 = Menus("Lambda234")
Lambda235 = Menus("Lambda235") * Million
Lambda238 = Menus("Lambda238") * Million
Lambda87 = Menus("Lambda87") * Million
Lambda147 = Menus("Lambda147") * Million
Lambda187 = Menus("Lambda187") * Million
Lambda176 = Menus("Lambda176") * Million
Lambda40 = Menus("Lambda40") * Million
Lambda230 = Menus("Lambda230")
Lambda234 = Menus("Lambda234")
Lambda231 = Menus("Lambda231")
Lambda226 = Menus("Lambda226")
Lambda210 = Menus("Lambda210")
Lambda238err = Menus("Lambda238perr") / 200 * Lambda238
Lambda235err = Menus("Lambda235perr") / 200 * Lambda235
Lambda234err = Menus("Lambda234perr") / 200 * Lambda234
Lambda230err = Menus("Lambda230perr") / 200 * Lambda230
LambdaDiff = Lambda230 - Lambda234
LambdaK = Lambda230 / LambdaDiff
End Sub

Private Sub Isotype_click() ' Handle "Plot-type" dropdowns selection, Isoplot-setup dialog.
Dim i%, j%, b1 As Boolean, b2 As Boolean, b3 As Boolean, b4 As Boolean, s1$, s2$
Dim r As Boolean, T$, Rin As Object, nc%, Ra As Range
Set Rin = zEdb("eRange")
i = dIsoType: r = IsOn(zChB("cCalculate"))

If Ash.Type = xlWorksheet And i = 20 Then
  MsgBox "You must first create the plot on which to add the function", , Iso
  i = 14 ' Arbitrary Fn - can only do from a Chart sheet
  dIsoType = i ' Change to "other x-y plot"
End If

Menus("Isotype") = i
b1 = (i < 3 Or (i > 7 And i < 13) Or i = 14 Or i = 19)
zOptB("oNormal").Visible = b1: zOptB("oInverse").Visible = b1
Rprompt i
zEdb("eRange").Enabled = (i <> 20)
Dim2Dim3
ErrType
b1 = (i < 15)
zChB("cCalculate").Visible = (i <> 16 And i <> 17)

If i = 13 Then ' U-series
  On Error GoTo 1
  Set Ra = Range(Rin.Text)

  For j = 1 To Ra.Areas.Count
    nc = nc + Ra.Areas(j).Columns.Count
  Next j

1:  On Error GoTo 0

  If nc = 3 Or nc = 6 Or nc = 9 Then ' 3D
    With zOptB("o3D"): .Visible = True: .Enabled = True: .Value = xlOn: End With
    With zOptB("oLinear3D"): .Visible = True: .Enabled = True: .Value = xlOn: End With
    With zOptB("oPlanar3D"): .Visible = True: .Enabled = False: .Value = xlOff: End With
    zOptB("oInverse") = xlOn: zOptB("oNormal") = xlOff
  End If

ElseIf i = 16 Or i = 17 Then ' Prob density or Lin Prob
  zChB("cPlot") = xlOn: zChB("cColor").Enabled = True

ElseIf i = 15 Or i = 21 Or i = 22 Or i = 23 Or i = 24 Or i = 25 Then
  zChB("cCalculate") = xlOn
End If

If i = 21 Or i = 22 Or i = 23 Then zChB("cPlot") = xlOff: zChB("cColor").Enabled = False
zChB("cPlot").Enabled = (i < 20 And i <> 16 And i <> 17)
zChB("cCalculate").Enabled = (i <> 24)
RobustOK
AnchorBox
Symbols
zChB("cAutoSc").Enabled = (b1 And IsOn(zChB("cPlot")) And i < 17)
ShowFill i, IsOn(zChB("cPlot")), zDropD("dSymbol")
ConcAgeBox
End Sub

Private Sub ShowFill(ByVal Itype%, ByVal DoPlot As Boolean, ByVal Symb%)
Dim A As Boolean, b As Boolean, c As Boolean, i%
i = Itype
A = False '(i = 16) ' Always show (prob density)
b = (i < 17 Or i = 18 Or i = 19) 'Optional
c = (Symb < 8 And Symb <> 2) ' Symbols which can be filled

With zChB("cShapes")
  .Enabled = (A Or (DoPlot And b And c))
  .Visible = .Enabled
End With

End Sub

Sub ConcAgeBox()
Attribute ConcAgeBox.VB_ProcData.VB_Invoke_Func = " \n14"

With zChB("cConcAge")
  .Visible = (dIsoType = 1)
  .Enabled = (.Visible And IsOn(zChB("cCalculate")) And _
   (IsOff(zOptB("o3d")) Or Not zOptB("o3d").Visible))
  If Not .Enabled Then zChB("cConcAge") = xlOff
End With

End Sub

Private Sub D2D3_click()
Dim2Dim3 ' Handle "2D/3D" option-buttons click of Isoplot-setup dialog.
AnchorBox
ConcAgeBox
End Sub

Private Sub DoPlot_click() ' Handle "Plot" checkbox-click from Isoplot Setup dialog.
Dim P As Boolean, r As Boolean, i%, s%, b As Boolean, pp As Boolean
i = dIsoType:  s = zDropD("dSymbol")
P = IsOn(zChB("cPlot")): r = IsOn(zChB("cCalculate"))
b = (i < 17 Or i = 18 Or i = 19)
ShowFill i, P, s

If i = 16 And Not P Then zChB("cPlot") = xlOn: Exit Sub

If Not P And Not r Then zChB("cPlot") = xlOn: P = True
zChB("cAutoSc").Enabled = (i < 15 And P)
zChB("cColor").Enabled = P

With zChB("cShapes")
  .Visible = (P And b)
  .Enabled = .Visible
End With

ErrType
Symbols
AnchorBox
InvertConcType
End Sub

Private Sub Calculate_click() ' Handle "Isochron" checkbox-click from Isoplot-Setup dialog.
Dim i%, c As Object
i = dIsoType: Set c = zChB("cCalculate")

If i = 15 Or i = 21 Or i = 22 Or i = 23 Or i = 24 Or (IsOff(c) And IsOff(zChB("cPlot"))) Then c = xlOn

RobustOK
Dim2Dim3
ErrType
AnchorBox
ConcAgeBox
End Sub

Private Sub ColorPlot_click()  ' Handle click on the Color Plot box of the
Dim i%, b As Boolean '  Isoplot-Setup dialog.
i = dIsoType

b = (i <> 15 And i <> 16 And i <> 17 And i <> 18 And IsOn(zChB("cColor")) And IsOn(zChB("cPlot")))

zDropD("dSymbClr").Enabled = b: zLbl("lSymbClr").Enabled = b
End Sub

Private Sub SymbType_click()
ErrType ' Handle implications of selecting a plot-symbol type
Symbols ' in the isoplot-setup dialog.
End Sub

Private Sub NormalInverse_click() ' Handle "2D/3D" option-buttons click  of isoplot-setup dialog.

If dIsoType = 13 And IsOn(zOptB("o3D")) Then ' U-series; only Inverse allowed
  zOptB("oInverse") = xlOn:  zOptB("oNormal") = xlOff
ElseIf dIsoType = 19 Then
  Rprompt dIsoType
Else
  Dim2Dim3
  AnchorBox
  Symbols
  InvertConcType
End If

End Sub

Private Sub PreProcessIsoPlotSetup(Optional PlotSpecified As Boolean = False)
Dim b As Boolean, i%, r As Boolean, s%, ci$, P$, q As Range, Lop As Range
Dim oCalc As Object, Oplot As Object, nAP As Boolean, G As Object, ttop!
Dim oPlan As Object, Olin As Object, oShp As Object, cColor As Object, Cols%
ViM PlotSpecified, False

With Selection
  For i = 1 To .Areas.Count: Cols = Cols + .Areas(i).Columns.Count: Next i
End With

nAP = Not AddToPlot: ci$ = "ClrNames"
Set dIsoType = zDropD("dIsotype")
Set oPlan = zOptB("oPlanar3D"): Set Olin = zOptB("oLinear3D")
Set oCalc = zChB("cCalculate"): Set Oplot = zChB("cPlot"): Set oShp = zChB("cShapes")
zDropD("dSymbClr").ListFillRange = ci$
Set Lop = Menus("LastOpSys") ' Initialize the plot-type dropdowns if necessary
zChB("cRobust").Caption = IIf(Mac, "robust reg", "Robust regr.")
ttop = zGrp("gNdim").Top
zGrp("gPlotType").Top = ttop: zGrp("gPlanar").Top = ttop
zGrp("gAbs").Visible = False: zGrp("gSig").Visible = False

If (Mac And Lop.Text <> "Mac") Or (Windows And Lop.Text <> "Windows") Then
  P$ = IIf(Windows, "WinIsoTypes", "MacIsoTypes")
  Set q = Menus(P$)
  For i = 1 To q.Rows.Count: IsoPlotTypes(i) = q(i): Next i
  Lop = IIf(Mac, "Mac", "Windows")
End If

If Ash.Type = xlWorksheet And Isotype = 20 Then
  dIsoType = i ' Arbitrary Fn - can only do from a Chart sheet
  Isotype = 14: dIsoType = 14 ' Change to "other x-y plot"
End If

If Isotype = 0 Then
  Isotype = 1:      SymbType = 1:     SymbClr = 1
  AbsErrs = False:  Inverse = False:  Dim3 = False
  Planar3D = False: SigLev = 2:       DoPlot = True
  Regress = True:   ColorPlot = True: AutoScale = True: ProbPlot = False
  Anchored = False: PbGrowth = False: Robust = False: CanReject = False
End If

If AbsErrs Then zOptB("oAbsolute") = xlOn Else zOptB("oPercent") = xlOn
If SigLev = 1 Then zOptB("o1sigma") = xlOn Else zOptB("o2sigma") = xlOn
zOptB("o3D") = Dim3: zOptB("o2D") = Not Dim3
zOptB("oInverse") = Inverse: zOptB("oNormal") = Not Inverse
If IsOff(Oplot) And IsOff(oCalc) Then oCalc = xlOn
i = Isotype: dIsoType = i
oCalc.Enabled = True

If i = 23 Or i = 24 Or i = 25 Then
  ' Mix, TuffZirc, Detrital zirc
  oCalc.Enabled = False: oCalc = xlOn
  Oplot = xlOff: Oplot.Enabled = False
  DoPlot = False
End If

Regress = IsOn(oCalc)

If i = 21 Or i = 22 Then ' Stacked
  DoPlot = False: Oplot = xlOff: oCalc = xlOn
ElseIf ((Regress And nAP) Or i = 15) And i <> 16 And i <> 17 And i <> 24 Then
  Oplot = DoPlot: oCalc = xlOn
Else
  Oplot = xlOn:   DoPlot = True
End If

Oplot.Enabled = (nAP And i <> 21 And i <> 22 And i <> 23 And i <> 24 _
                  And i <> 25 And i <> 16 And i <> 17)

zChB("cAutoSc") = AutoScale
b = (i < 3 Or (i > 7 And i < 13) Or i = 14 Or i = 19) Or Cols = 8
zOptB("oNormal").Visible = b: zOptB("oInverse").Visible = b
Rprompt i
zEdb("eRange").Enabled = (i <> 20)
Dim2Dim3
ErrType
Symbols
b = (i < 15 And DoPlot)
zChB("cAutoSc").Enabled = (b And nAP)
oCalc.Visible = (i <> 16 And i <> 17)
RobustOK

If AddToPlot Then
  zChB("cColor") = ColorPlot
  zDropD("dSymbClr").Enabled = (ColorPlot And i < 15)
Else
  zDropD("dSymbClr").Enabled = (DoPlot And ColorPlot And i <> 24 And i <> 23 And i <> 18 And i <> 19)
End If

zLbl("lSymbClr").Enabled = zDropD("dSymbClr").Enabled
zChB("cColor").Enabled = ((DoPlot Or i = 21) And nAP)
If i = 15 Then CanReject = IsOn(zChB("cAnchored"))
DoPlot = IsOn(Oplot)
AnchorBox
dIsoType.Enabled = nAP
zOptB("oNormal").Enabled = nAP: zOptB("oInverse").Enabled = nAP

If PlotSpecified And Dim3 And Isotype = 1 Then
  zOptB("oPlanar3D") = Planar3D: zOptB("oLinear3D") = Linear3D
End If

If oPlan.Visible And IsOff(oPlan) And IsOff(Olin) Then Olin = xlOn
oShp = DoShape
ShowFill i, DoPlot, zDropD("dSymbol")
ConcAgeBox
InvertConcType
ColorPlot_click
End Sub

Private Sub PostProcessIsoPlotSetup()
 ' Get variables defining the data & plot type from the completed isoplot-setup dialog box.
Dim i%, s%, b As Boolean, FC$
Dim oSymb As Object, Olin As Object, oPlan As Object, oCalc As Object, o3d As Object, oAnch As Object
Set oSymb = zDropD("dSymbol"): Set Olin = zOptB("oLinear3D"): Set oPlan = zOptB("oPlanar3D")
Set oCalc = zChB("cCalculate"): Set o3d = zOptB("o3D"): Set oAnch = zChB("cAnchored")
Isotype = dIsoType:  i = Isotype
PlotIdentify
Regress = ((IsOn(oCalc) And oCalc.Visible) And Not Stacked And Not DoMix)
SymbType = oSymb: s = SymbType

Inverse = (IsOn(zOptB("oInverse")) And ((i = 18 Or i = 19) Or (i < 3 Or (i > 7 And i < 15))))

Normal = Not Inverse
SymbClr = zDropD("dSymbClr")
Dim3 = (IsOn(o3d) And o3d.Visible)
Linear3D = (IsOn(Olin) And Olin.Visible)
Planar3D = (IsOn(oPlan) And oPlan.Visible)
AbsErrs = IsOn(zOptB("oAbsolute"))
SigLev = 2 + IsOn(zOptB("o1sigma"))
ColorPlot = IsOn(zChB("cColor"))
DoPlot = ((IsOn(zChB("cPlot")) Or AddToPlot) And i <> 20)
Menus("doPlot") = DoPlot
b = IsOn(oAnch)
Anchored = False: PbGrowth = False: WtdAvXY = False
If Not WtdAvPlot Then CanReject = False

If oAnch.Visible Then

  Select Case i
    Case 1:    Anchored = b
    Case 8, 9: If Normal Then PbGrowth = b
    Case 14:   WtdAvXY = b
    Case 15:   CanReject = b
  End Select

End If

b = (i < 15)
AutoScale = (IsOn(zChB("cAutoSc")) And zChB("cAutoSc").Enabled)

If oSymb.Visible And oSymb.Enabled Then
  Eellipse = False: eCross = False: Ebox = False: excSymb = 0
  StraightLine = False: Nspline = False: Aspline = False: SplineLine = False

  Select Case oSymb
    Case 1: Eellipse = True
    Case 2: eCross = True
    Case 3: Ebox = True
    Case 4: excSymb = xlSquare
    Case 5: excSymb = xlDiamond
    Case 6: excSymb = xlTriangle
    Case 7: excSymb = xlCircle
    Case 8: excSymb = xlX
    Case 9: excSymb = xlPlus
    Case 10: StraightLine = True
    Case 11: Nspline = True
    Case 12: Aspline = True
  End Select

  SplineLine = (Nspline Or Aspline)
  DoShape = (ShapesOK And IsOn(zChB("cShapes")))

  If ColorPlot Then
    excClrInd = zDropD("dSymbClr").ListIndex
    'FC$ = Strip(Menus("Colors").Cells(excClrInd).Text, " ")
    'If FC$ = "FontColor" Then
    '  excClr = -1
    'ElseIf LCase$(FC$) = "black" Then
    '  excClr = vbBlack ' because vbBlack=0 = no color
    'Else
    '  excClr = MenuSht.Shapes(FC$).Fill.ForeColor.RGB
    'End If
    excClr = Menus("ClrStuff")(excClrInd, 4)
  Else
    excClr = vbBlack
  End If

End If

RobustOK
InvertConcType
AnyCurve = (Isotype = 20)

If Not Canceled Then

  If Stacked Then

    With DlgSht("Bracket").OptionButtons
      .Item("o" & tSt(SigLev) & "sigma") = xlOn
      If AbsErrs Then .Item("oAbs") = xlOn Else .Item("oPercent") = xlOn
    End With

  End If

End If

End Sub

Private Sub Robust_click()
Dim r As Object
ErrType

If dIsoType = 14 Then ' Other X-Y
  Set r = zChB("cRobust")

  With zChB("cAnchored")
    .Enabled = IsOff(r)
    If Not .Enabled Then .Value = xlOff
  End With

End If

End Sub

Private Sub ErrType()  ' Handle error-type elements of isoplot-setup dialog
Dim b As Boolean, i%, c As Object

i = dIsoType: Set c = zChB("cRobust")
b = (i > 14 Or (IsOn(zChB("cCalculate")) And Not ((c.Visible And c.Enabled And IsOn(c)))))
b = b Or (IsOn(zChB("cPlot")) And zDropD("dSymbol") < 4)
b = b And i <> 20 ' AnyCurve
zGrp("gErrors").Enabled = b
zOptB("o1sigma").Enabled = b:  zOptB("o2sigma").Enabled = b
zOptB("oPercent").Enabled = b: zOptB("oAbsolute").Enabled = b
End Sub

Private Sub Dim2Dim3() ' Handle 2d/3d & linear/planar opt-buttons of Isoplot-setup dialog
Dim i%, j%, b As Boolean, z2 As Object, z3 As Object
Dim zi As Object, zN As Object, zP As Object, zL As Object

Set z2 = zOptB("o2D"): Set z3 = zOptB("o3D"): Set zi = zOptB("oInverse")
Set zN = zOptB("oNormal"): Set zP = zOptB("oPlanar3D"): Set zL = zOptB("oLinear3D")

b = Not AddToPlot
z2.Enabled = b: z3.Enabled = b: zP.Enabled = b: zL.Enabled = b
If AddToPlot Then Exit Sub
i = dIsoType
b = ((i = 1 And zi = xlOn) Or i = 13 Or i = 14) ' Useries, OtherXY
If Not b Then z2 = xlOn: z3 = xlOff
b = (b And IsOn(zChB("cCalculate")))
zGrp("gNdim").Visible = b: z2.Visible = b: z3.Visible = b
b = (b And IsOn(z3))

zL.Visible = b: zGrp("gPlanar").Visible = b: zP.Visible = b: zL.Enabled = b

If i = 13 Then
  zP.Enabled = False: zL = xlOn
  If IsOn(z3) Then zi = xlOn: zN = xlOff
Else
  zP.Enabled = zP.Visible
End If

Rprompt i
ConcAgeBox
InvertConcType
RobustOK
End Sub

Private Sub AnchorBox()   'Handle "Anch"-box of isoplot setup dialog
Dim i%, d2 As Boolean, b As Boolean, T$, AC As String * 1
Dim Norml As Boolean, Pb As Boolean, A As Object, Calc As Boolean, Plt As Boolean
' Visible if wtd av, calculated normal Pb-Pb isochron, or calculated 2-D Concordia plot

i = dIsoType: Pb = (i = 8 Or i = 9): Calc = IsOn(zChB("cCalculate"))
Plt = IsOn(zChB("cPlot"))
Norml = IsOn(zOptB("oNormal"))
d2 = (IsOff(zOptB("o3D")) Or Not zOptB("o3d").Visible)

b = ((i = 1 And d2) Or (i = 14 And Calc) Or i = 15) ' 2-D concordia, OtherXY, or WtdAv
b = (b Or (Pb And Norml And Plt And Not AddToPlot)) ' "normal" Pb-Pb plot?

b = b And IsOff(zChB("cConcAge"))
Set A = zChB("cAnchored")

With A
  .Visible = b: .Enabled = b

  If b Then

    Select Case i
      Case 1:    b = Anchored:  T$ = "Anchored   ": AC = "d"
      Case 8, 9: b = PbGrowth:  T$ = "PbGrowth  ":  AC = "g"
      Case 14
        T$ = IIf(Mac And Int(ExcelVersion) >= 10, "XY WtdAv  ", "XY Wtd Avg  ")
        AC = "d"
      Case 15:   b = CanReject: T$ = "Reject OK?":  AC = "j"
    End Select

    If Not b Then A = xlOff
    .Text = T$: .Accelerator = AC
  End If

End With

End Sub

Private Sub Anchor_click()  'Handle "Anch"-box click of isoplot setup dialog
Dim b As Boolean, r As Object

b = IsOn(zChB("cAnchored"))

Select Case dIsoType
  Case 1:    Anchored = b
  Case 8, 9: If IsOn(zOptB("oNormal")) Then PbGrowth = b
  Case 14
    WtdAvXY = b
    Set r = zChB("cRobust")
    r.Enabled = Not b
    If b Then r.Value = xlOff
  Case 15:   CanReject = b
  Case Else:
End Select

End Sub

Private Sub Symbols() ' Handle plot-symbol elements of isoplot-setup dialog
Dim b As Boolean, i%, P As Boolean, r As Boolean, s%
Dim c As Boolean, s1$, s2$, s3$, s4$

i = dIsoType: s = zDropD("dSymbol")
r = IsOn(zChB("cCalculate")): P = IsOn(zChB("cPlot"))
c = IsOn(zChB("cColor"))
b = (i < 15 And P) ' not WtdAv, CumProb, Ar-Ar Step, AnyCurve, Mix, Bracket
zGrp("gSymbols").Enabled = b:   zLbl("lSymbol").Enabled = b
zDropD("dSymbol").Enabled = b:  zDropD("dSymbClr").Enabled = (b And c)
zLbl("lSymbClr").Enabled = (b And c)
ShowFill i, P, s

If zDropD("dSymbol").Enabled Then
  b = (P And IsOn(zOptB("oNormal")) And (i = 1 Or i = 8 Or i = 9 Or i = 17))
  ' Concordia, Ar-Ar Step, Pb-Pb
  s1$ = "error cross": s2$ = "error box"

  If b Then
    s1$ = "[" & s1$ & "]": s2$ = "[" & s2$ & "]"
    If s = 2 Or s = 3 Then zDropD("dSymbol").ListIndex = 1
  End If

  If IsOn(zChB("cPlot")) Then
    With Menus("Symbols"): .Cells(2) = s1$: .Cells(3) = s2$: End With
  End If

End If

If b And s < 4 Then ErrType
End Sub

Sub RangeCheck(Nrows%, Ncols%, Nareas%)
Attribute RangeCheck.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, Bad As Boolean, a1 As Range, a2 As Range, s As Range, b As Boolean, j%
Dim NM%, Nw%, r As Range, tRw%, tCo%, tC As Range
Dim Frow&, Lrow&, Fcol%, Lcol%, Ecol%

Nrows = 0: Ncols = 0:  Ecol = 256
On Error GoTo BadRange
Set r = Selection
On Error GoTo 0
' Select contiguous numeric range if a single cell is selected
' Not perfect, needs work!
Nareas = r.Areas.Count

If Nareas = 1 And r.Count = 1 Then
  Set r = r.CurrentRegion
  NM = r.Columns.Count: Nw = r.Rows.Count

  If Nw > 1 Then
    Set tC = ActiveCell
    tRw = tC.Row: tCo = tC.Column

    Do
      tRw = tRw - 1
    Loop Until tRw <= 1 Or Not IsNumber(Cells(Max(1, tRw), tCo))

    Frow = Max(1, tRw + 1): tRw = tC.Row
    If Frow = 2 And IsNumber(Cells(1, tCo)) Then Frow = 1

    Do
      tRw = tRw + 1
    Loop Until tRw >= EndRow Or Not IsNumber(Cells(Min(EndRow, tRw), tCo))

    Lrow = Max(1, tRw - 1): tRw = tC.Row
    If Lrow = EndRow - 1 And IsNumber(Cells(EndRow, tCo)) Then Lrow = EndRow

    Do
      tCo = tCo - 1
    Loop Until tCo <= 1 Or Not IsNumber(Cells(tRw, Max(1, tCo)))

    Fcol = MinMax(1, 256, tCo + 1)
    If Fcol = 2 And IsNumber(Cells(tRw, 1)) Then Fcol = 1

    Do
      tCo = tCo + 1
    Loop Until tCo >= Ecol Or Not IsNumber(Cells(tRw, Min(Ecol, tCo)))

    Lcol = MinMax(1, 256, tCo - 1)
    If Lcol = Ecol - 1 And IsNumber(Cells(tRw, Ecol)) Then Lcol = Ecol
    Set r = sR(Frow, Fcol, Lrow, Lcol)
    r.Select
  End If

End If

With r
  ColWise = True

  For i = 2 To Nareas
    If .Areas(i).Rows.Count <> .Areas(1).Rows.Count Then RowWise = False
    If .Areas(i).Row <> .Areas(1).Row Then ColWise = False
  Next i

  RowWise = Not ColWise: Bad = False

  For i = 2 To Nareas
    Set a1 = .Areas(i - 1): Set a2 = .Areas(i)

    If ColWise Then
      If a1.Row <> a2.Row Or a1.Rows.Count <> a2.Rows.Count Then Bad = True
    Else
      If a1.Column <> a2.Column Or a1.Columns.Count <> a2.Columns.Count Then Bad = True
    End If

  Next i

  If Bad Then MsgBox "Invalid range for Isoplot", , Iso: ExitIsoplot
  ' Find out how many rows & columns of data.  If legal, parse into a A1-style
  '  range spec.

  For i = 1 To Nareas

    With .Areas(i)

      If ColWise Then
        If i = 1 Then Nrows = Min(32766, .Rows.Count)
        Ncols = Ncols + .Columns.Count
      Else
        If i = 1 Then Ncols = .Columns.Count
        Nrows = Nrows + .Rows.Count
      End If

    End With
  Next i
End With

If Nrows > 0 And ((Ncols > 0 And Ncols < 7) Or Ncols = 9) Then _
  StP.EditBoxes("eRange").Text = rStr(Selection.Address, "$")
Exit Sub

BadRange: On Error GoTo 0
MsgBox "Selection is not a valid worksheet range", , Iso
ExitIsoplot
End Sub

Sub IncludeDecayConstErrsClick() ' OptionButton of MonteCarloErrors group-box clicked in IsoRes
Attribute IncludeDecayConstErrsClick.VB_ProcData.VB_Invoke_Func = " \n14"
Dim r As Object, Ch As Object, Op As Object, Front As Object, Behind As Object, b As Boolean
AssignD "IsoRes", ResBox, , Ch, Op
Set Front = Op("oFront"): Set Behind = Op("oBehind")
If IsOff(Behind) And IsOff(Front) Then Behind = xlOn

If DoPlot Then
  b = IsOn(Ch("cShowRes"))
ElseIf Cdecay0 Then
  Behind.Enabled = True: Front.Enabled = True
End If

End Sub

Sub AddPoints()
Attribute AddPoints.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s$, o As Object, c As Object, A
On Error GoTo NotAchartsheet
Set sc = Ach.SeriesCollection
On Error GoTo 0
GetDatSheetName

If Ash.Type <> xlXYScatter Then
  MsgBox "Data can only be added to a full-sheet Isoplot chart", , Iso
  KwikEnd
End If

GetPlotInfo 0

With StP
  Set o = .OptionButtons: Set c = .CheckBoxes
  .DropDowns("dIsotype") = Isotype
End With

ConcAgePlot = ConcAge
c("InvertPlot") = InvertPlotType
c("cConcAge") = (ConcAge And ConcPlot And Not Dim3)
If (ConcPlot And InvertPlotType) Then Inverse = Not Inverse: Normal = Not Inverse
o("oNormal") = Not Inverse: c("cColor") = ColorPlot
o("o3D") = Dim3: o("oPlanar3D") = Planar3D
c("InvertPlot").Enabled = False
c("cShapes") = DoShape
PlotIdentify

If WtdAvPlot Or ProbPlot Or ArgonStep Or DoMix Then
  A = Array("Weighted average", "Probability", "Step-heating", "Gaussian-unmixing")
  s$ = A(-WtdAvPlot - 2 * ProbPlot - 3 * ArgonStep - 4 * DoMix)
  MsgBox "You can't add data to a " & s$ & " plot", , Iso: KwikEnd
ElseIf CumGauss Then
  LoadUserForm AddHisto
  AddHisto.Show
End If

With StP
  AutoScale0 = IsOn(.CheckBoxes("cAutoSc"))
  DoShape = IsOn(.CheckBoxes("cShapes"))
  AutoScale0 = IsOn(.CheckBoxes("cAutoSc"))
  Regress0 = IsOn(.CheckBoxes("cCalculate"))
  .CheckBoxes("cCalculate") = xlOff
End With

Isoplot True
Exit Sub

NotAchartsheet: MsgBox "This is not a chart created by Isoplot", , Iso
End Sub

Sub LabelPoints()  ' Label the data points of an Isoplot chart
Attribute LabelPoints.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, k%, P%, q%, M%, r%
Dim lRange As Range, Dlb As Object, T$, s$(), Dup() As Boolean, ri As Range, rj As Range
Dim Npts%, LrA As Object, W&, b As Boolean, Ser() As Object
Dim Pts As Object, OK As Boolean, Csh As Object, Dlg As Object
Dim NdSer%, CshSC As Object, db As Object, DS As Object, Op As Object, L As Object
NoUp
Set Csh = Ash: Set Dlg = DlgSht
GetOpSys

If Csh.Type = xlWorksheet Then
  On Error Resume Next
  Selection.Activate  ' in case selection but not activated
  On Error GoTo 0
End If

GetPlotInfo OK
PlotIdentify

If Not OK Then
  MsgBox "Not an Isoplot chart, or" & vbLf & _
    "hidden PlotDat sheets are missing/corrupt, or" & vbLf & _
    "source data-sheet has been renamed.", , Iso
  KwikEnd
ElseIf CumGauss Or ArgonStep Or ProbPlot Then
  MsgBox "Can't label points for this type of plot", , Iso
  KwikEnd
End If

Csh.Activate
NdSer = NumDataSeries(Ach)

If NdSer = 0 Then
  MsgBox "There are no data-points in this chart" & _
    vbLf & "that were plotted by Isoplot.", , Iso
  KwikEnd
End If

Set CshSC = Ach.SeriesCollection
'=SERIES("IsoDat1",PlotDat4!$C$1:$C$12,PlotDat4!$D$1:$D$12,1)
'=SERIES("IsoDat2",PlotDat4!$C$1:$C$12,PlotDat4!$D$1:$D$12,2)
ReDim Sname$(NdSer), s$(NdSer), Dup(NdSer), Ser(NdSer)
LoadUserForm Series

For i = 1 To NdSer       ' Parse out the range addreses (in the PlotDat sheet)
  T$ = CshSC(i).Formula  '  of the different data-series.
  P = InStr(T$, ","): q = InStr(RevStr(T$), ","): M = Len(T$)
  s$(i) = Mid(T$, 1 + P, M - q - P)
Next i

For i = 1 To NdSer ' Look for any duplicate series
  Dup(i) = False

  For j = 1 To NdSer

    If i <> j And Not Dup(i) Then
      Set ri = Range(s$(i)): Set rj = Range(s$(j))

      If ri.Count = rj.Count And ri.Areas.Count = rj.Areas.Count Then
        b = True

        For k = 1 To ri.Areas.Count

          For M = 1 To ri.Areas(k).Rows.Count

            For r = 1 To ri.Areas(k).Columns.Count
              If ri.Areas(k)(M, r) <> rj.Areas(k)(M, r) Then b = False: Exit For
            Next r

          Next M
        Next k

        If b Then Dup(j) = True

      End If

    End If

  Next j

Next i

k = 0

With Series ' Put the non-duplicate series in the dropdown

  For i = 1 To NdSer

    If Not Dup(i) Then
      k = 1 + k
      Set Ser(k) = CshSC("IsoDat" & tSt(i))
      Sname$(k) = " " & Str(k) + Space(16) + Str(Ser(k).Points.Count)
      .sOrder.AddItem Sname$(k)
    End If

  Next i

  .sOrder.ListIndex = 0
  b = (k > 1)
  If Not b Then Set DS = CshSC("IsoDat1")
  .sOrder.Visible = b: .Oplot.Visible = b
  .Npts.Visible = b:   .SelectPts.Visible = b
  .Show
End With

If Canceled Then KwikEnd
Set DS = Ser(PubInt(1)): Set Pts = DS.Points
Npts = Pts.Count

1:

Do
  LoadUserForm DatLab
  NoUp False
  DatSht.Activate
  On Error Resume Next
  DatLab.LabelRange = Selection.Address
  On Error GoTo 0
  DatLab.Show
  If Canceled Then KwikEnd
  NoUp
  PubObj(2).Select
  On Error GoTo 2
  Set lRange = PubObj(1)
  lRange.Select
  On Error GoTo 0
If lRange.Count = Npts Then Exit Do
  MsgBox "Number of Labels must match number of plotted points (=" & _
    tSt(Npts) & ")", , Iso
Loop

Csh.Activate

With DS
  .ApplyDataLabels Type:=xlShowLabel, LegendKey:=False
  With .DataLabels.Font: .Name = "Arial Narrow": .Size = 10: End With
End With

Set LrA = lRange.Areas

Select Case PubInt(2)
  Case 1: W = xlLabelPositionAbove
  Case 2: W = xlLabelPositionRight
  Case 3: W = xlLabelPositionBelow
  Case 4: W = xlLabelPositionLeft
End Select

k = 0

For i = 1 To LrA.Count

  For j = 1 To LrA(i).Count
    k = 1 + k

    With Pts(k).DataLabel
      .Text = LrA(i).Cells(j).Text
      .Position = W
    End With

Next j, i

If Ash.Type = xlXYScatter Then Ach.Deselect
Exit Sub

2: On Error GoTo 0
MsgBox "Invalid range (try selecting again with mouse)", , Iso
GoTo 1
End Sub

Private Sub WithDcErrs2_click()
Dim b1 As Boolean, b As Boolean

With DlgSht("concscale")
  b1 = IsOn(.CheckBoxes("cWLE"))
  b = (DoShape And Cdecay0 And b1)
  .OptionButtons("oBehind").Enabled = b
  .OptionButtons("oFront").Enabled = b
  .GroupBoxes("gAgeTicks").Enabled = Not b1
  .OptionButtons("oCircles").Enabled = Not b1
  .OptionButtons("oLines").Enabled = Not b1
End With

End Sub

Private Sub WithDcErrs3_click()
Dim b As Boolean, o As Object

With DlgSht("IsoRes")
  Set o = .OptionButtons
  b = (DoShape And Cdecay0 And IsOn(.CheckBoxes("cWLE")))
  o("oBehind").Enabled = b: o("oFront").Enabled = b
  .GroupBoxes("gAgeTicks").Enabled = Not b
  o("oCircles").Enabled = Not b: o("oLines").Enabled = Not b

  If Not b And o("oCircles") = xlOff And o("oLines") = xlOff Then o("oCircles") = xlOn

End With

End Sub

Private Function rStr(ByVal Phrase$, ByVal ReplaceThis$, Optional ByVal WithThis = "") As String
ViM WithThis, ""
rStr = ApSub(Phrase$, ReplaceThis$, (WithThis))
End Function

Private Sub Autoscale2_click()
Dim d As Boolean, c As Object, L As Object, e As Object, Bu As Object, Grp As Object

AssignD "ConcScale", , e, c, , L, Grp, , Bu
d = IsOff(c("cAutoscale"))
Bu("bXYlim").Enabled = d: Grp("gAgeLim").Enabled = d
L("lMinAge").Enabled = d: L("lMaxAge").Enabled = d
L("lDefBy").Enabled = d:  L("lLims").Enabled = d
e("eMinAge").Enabled = d: e("eMaxAge").Enabled = d
AutoScale = Not d
End Sub

Private Sub Autoscale3_click()
AutoScale = True: Canceled = False
End Sub

Private Sub InvertConcType()
Dim b As Boolean

b = (dIsoType = 1 And IsOn(zOptB("o2D")))

With zChB("InvertPlot")
  .Visible = (dIsoType = 1 And IsOn(zOptB("o2D")) And IsOn(zChB("cPlot")))

  If .Visible Then
    .Text = "plot as " & IIf(IsOn(zOptB("oNormal")), "Tera-Wasserburg", "conv. concordia     ")
  ElseIf dIsoType = 1 And IsOn(zOptB("o3D")) Then
    .Value = xlOff
  End If

  InvertPlotType = IsOn(.Value)
  .Enabled = Not AddToPlot
End With

End Sub

Private Sub CnameAssign()

If Normal Then
  AxX$ = Menus("AxXn").Cells(Isotype): AxY$ = Menus("AxYn").Cells(Isotype)
Else
  AxX$ = Menus("AxXi").Cells(Isotype): AxY$ = Menus("AxYi").Cells(Isotype)
End If

End Sub

Private Sub ConcLinType_click()
Dim c As Object, L As Object, o As Object, e As Object, d As Object
AssignD "ConcLinType", d, e, c, o, L

With d
  If Not DoPlot Then c("cShowProj") = xlOff
  c("cShowProj").Enabled = DoPlot
  ConcConstr = IsOff(o("oUnConstr"))
  e("eAlpha0").Enabled = ConcConstr: e("eBeta0").Enabled = ConcConstr
  L("lAlpha0").Enabled = ConcConstr: L("lBeta0").Enabled = ConcConstr
  PlotProj = IsOn(c("cShowProj"))
  TuPbAlpha0 = EdBoxVal(e("eAlpha0")): TuPbBeta0 = EdBoxVal(e("eBeta0"))
End With

End Sub

Sub GetStorStatus(Optional GotStatus, Optional PlotSpecified As Boolean = False) ' Get/Store user choices/constants
Attribute GetStorStatus.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, vv, cc As Object, d As Object, P$, GetStat As Boolean, Lam, b(8) As Boolean
Dim MAP As Object, tB1 As Boolean, tb2 As Boolean, j%, oF As Range, po As Object, pn$
Static Gotted As Boolean
ViM PlotSpecified, False
NoAlerts

With MenuSht
  Set oF = .Range("OptFonts")
  Set Oo = .Range("Options"):  Set MAP = .Range("ModelAgeParams").Cells
End With

GetStat = NIM(GotStatus)
If GetStat And Gotted Then GotStatus = True: Exit Sub
GotStatus = False

With MenuSht


  If GetStat Then
    If FromSquid Then
      SigLev = 1: AbsErrs = False: AutoRescale = True
    Else

      If Not PlotSpecified Then
        Isotype = .Range("IsoType")
        Inverse = .Range("Inverse")
      End If

      Normal = Not Inverse
      SigLev = .Range("SigmaLevel"):   DoShape = .Range("FilledSymbols")
      ColorPlot = .Range("ColorPlot"): Anchored = .Range("Anchored")
      AutoScale = .Range("AutoScale"): DoPlot = .Range("DoPlot")
      Regress = .Range("Regress"):     AbsErrs = .Range("AbsErrs")
      AutoSort = .Range("Autosort"):   AutoRescale = .Range("AutoRescale")
      On Error Resume Next
      LineAgeTik = .Range("LineAgeTik")
    End If

    On Error GoTo 0

  Else
    .Range("IsoType") = Isotype: .Range("SigmaLevel") = SigLev
    .Range("Inverse") = Inverse: .Range("ColorPlot") = ColorPlot
    .Range("AutoScale") = AutoScale: DoPlot = .Range("DoPlot")
    .Range("Regress") = Regress: .Range("AbsErrs") = AbsErrs
    .Range("Anchored") = Anchored: .Range("FilledSymbols") = DoShape
    .Range("LineAgeTik") = LineAgeTik
    .Range("Autosort") = AutoSort: .Range("AutoRescale") = AutoRescale
  End If

  If GetStat Then
    ArMinGas = .Range("ArMinGas"): Air4036 = .Range("_Air4036")
    ArMinProb = .Range("ArMinProb"): ArMinSteps = .Range("ArMinSteps")
    NoSuper = .Range("NoSuper")
    If Air4036 <= 0 Then Air4036 = 295.5
    StackIso = .Range("StackIso")
    NoUpdate = False '.Range("NoUpdate")
    WideMargins = False '.Range("WideMargins")
    Opt.AxisThickLine = Oo(24): Opt.IsochLineThick = Oo(32)
  Else
    If Air4036 <= 0 Then Air4036 = 295.5
    .Range("ArMinGas") = ArMinGas: .Range("_Air4036") = Air4036
    .Range("ArMinSteps") = ArMinSteps: .Range("ArMinProb") = ArMinProb
    .Range("NoSuper") = NoSuper
    .Range("NoUpdate") = False: .Range("StackIso") = StackIso
    .Range("WideMargins") = False 'WideMargins
    Oo(24) = Opt.AxisThickLine ': Oo(32) = Opt.IsochLineThick
  End If

  If GetStat Then

    If FromSquid Then
      CurvWithDce = False
    Else
      CurvWithDce = .Range("CurvWithDce"): BandBehind = .Range("BandBehind")
      'Pvlines = .Range("Pvlines")
    End If

  Else
    .Range("CurvWithDce") = CurvWithDce: .Range("BandBehind") = BandBehind
    '.Range("Pvlines") = Pvlines
  End If

  If GetStat Then
    GotStatus = True: Gotted = True
    'For i = 1 To oO.Count: oO(i) = oD(i): Next i
    i = Oo(29)
    If i < 1 And i > 5 Then Oo(29) = 2 ' Opt.ConcLineThick
  End If

End With

GetStoreDone: On Error Resume Next
Close #1
NoAlerts False
On Error GoTo 0
End Sub

Function IsOn(Obj) As Boolean
Attribute IsOn.VB_ProcData.VB_Invoke_Func = " \n14"
IsOn = (Obj = xlOn)
End Function

Function IsOff(ByVal v) As Boolean
Attribute IsOff.VB_ProcData.VB_Invoke_Func = " \n14"
IsOff = (v = xlOff)
End Function

Sub Sigma1Click()
Attribute Sigma1Click.VB_ProcData.VB_Invoke_Func = " \n14"
zOptB("o1Sigma") = xlOn: zOptB("o2Sigma") = xlOff
End Sub

Sub Sigma2Click()
Attribute Sigma2Click.VB_ProcData.VB_Invoke_Func = " \n14"
zOptB("o2Sigma") = xlOn: zOptB("o1Sigma") = xlOff
End Sub

Sub PercentClick()
zOptB("oAbsolute") = xlOff: zOptB("oPercent") = xlOn
End Sub

Sub AbsClick()
Attribute AbsClick.VB_ProcData.VB_Invoke_Func = " \n14"
zOptB("oAbsolute") = xlOn: zOptB("oPercent") = xlOff
End Sub

Sub RowColWise(RangeInVals#(), ByVal Nrows%, ByVal Na%, nc%, ct%)
Attribute RowColWise.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, j%, k%, M%, MtCt%, tB As Boolean, Msg$, IgnoreWarning As Boolean

Msg = "Your input-data range uses a " & Chr(34) & "percent" & Chr(34) & _
      " number-format, so that the underlying values (used by Isoplot) of such cells " _
      & "is 100x less the displayed value." & vbLf & vbLf & _
     "Knowing this, do you still wish to proceed?"

If RowWise Then       ' Select valid data-rows, ignore others
  RightCol = RangeIn(1).Column + ndCols - 1
  IgnoreWarning = False

  For j = 1 To Na

    With RangeIn(j)

      For i = 1 To .Rows.Count
        MtCt = 0: MTrow(i) = False

        For k = 1 To ndCols

          With .Cells(i, k)

            If Not IgnoreWarning And Right(.NumberFormat, 1) = "%" Then
              If MsgBox(Msg, vbYesNo, "Isoplot") = vbNo Then
                ExitIsoplot
              Else
                IgnoreWarning = True
              End If

            End If

            tB = .Font.Strikethrough

            If IsNumeric(.Value) And Not tB Then
              If CDbl(.Value) = 0 Then MtCt = 1 + MtCt
            Else
              MtCt = ndCols:  MTrow(i) = True
            End If

          End With

        Next k

        If MtCt < ndCols Then
          If Not MTrow(i) And TopRow = 0 Then TopRow = .Row + i - 1
          ct = 1 + ct

          For k = 1 To ndCols
            RangeInVals(ct, k) = Min(CDbl(.Cells(i, k).Value), 1E+38)

            If k = 1 Then

              With .Cells(i, 1).Font
                DatClr(ct) = .Color
                If DatClr(ct) = xlAutomatic Or Not ColorPlot Then DatClr(ct) = Black
                If .Bold Then DatClr(ct) = -DatClr(ct)
              End With

              ReDim Preserve ValidRow(ct)
              ValidRow(ct) = .Cells(i, 1).Row
            End If

          Next k

        End If

      Next i

    End With

  Next j

Else ' Colwise

  For i = 1 To Nrows
    MtCt = 0: MTrow(i) = False

    For j = 1 To Na

      With RangeIn(j)
        nc = .Columns.Count
        RightCol = Max(RightCol, .Column + nc - 1)

        For k = 1 To nc

          With .Cells(i, k)

            If Not IgnoreWarning And Right(.NumberFormat, 1) = "%" Then

              If MsgBox(Msg, vbYesNo, "Isoplot") = vbNo Then
                ExitIsoplot
              Else
                IgnoreWarning = True
              End If

            End If

            tB = .Font.Strikethrough

            If IsNumeric(.Value) And Not tB Then
              If CDbl(.Value) = 0 Then MtCt = 1 + MtCt
            Else
              MtCt = ndCols:  MTrow(i) = True
            End If

          End With

        Next k

      End With

    Next j

    If TopRow = 0 And Not MTrow(i) Then TopRow = RangeIn(1).Row + i - 1

    If MtCt < ndCols Then

      ct = 1 + ct: M = 0

      For j = 1 To Na

        With RangeIn(j)
          nc = .Columns.Count

          For k = 1 To nc
            M = M + 1
          If M > ndCols Then Exit For
            RangeInVals(ct, M) = Min(CDbl(.Cells(i, k).Value), 1E+38)

            If k = 1 Then

              With .Cells(i, 1).Font
                DatClr(ct) = .Color
                If DatClr(ct) = xlAutomatic Or Not ColorPlot Then DatClr(ct) = Black
                If .Bold Then DatClr(ct) = -DatClr(ct)
              End With

              ReDim Preserve ValidRow(ct)
              ValidRow(ct) = .Cells(i, 1).Row
            End If

          Next k

        End With

      Next j

    End If

  Next i

End If

End Sub

Sub Rprompt(ByVal PlotType%)
Attribute Rprompt.VB_ProcData.VB_Invoke_Func = " \n14"
Dim j%, z As Object, s$
Set z = StP.OptionButtons
If IsOn(z("o3D")) Then j = 3 Else j = 1 - IsOn(z("oInverse"))
s$ = Menus("InputRangePrompt")(PlotType, j)
StP.Labels("lRangeExp").Text = App.Substitute(s$, "|", vbLf)
End Sub

Function Sp(ByVal v, ByVal PwrTen%, Optional SIGN = False)
Attribute Sp.VB_ProcData.VB_Invoke_Func = " \n14"
' Round number to -PwrTen decimal places, put in string, with + or - if specified.
ViM SIGN, False
Sp = sn$(Prnd(v, PwrTen), SIGN)
End Function

Sub AddSymbCol(ByVal Incr%)
Attribute AddSymbCol.VB_ProcData.VB_Invoke_Func = " \n14"
SymbCol = SymbCol + Incr

If SymbCol > 256 Then
  SymbRow = 200 + SymbRow: SymbCol = 5
End If

End Sub

Private Sub ConcAgeClick()

With zChB("cAnchored")
  .Visible = IsOff(zChB("cConcAge"))
  .Enabled = .Visible
End With

End Sub

Private Sub RobustOK()
Dim i%, d3 As Boolean, Dim3 As Object, Linear As Boolean
Dim OK As Boolean, Calc As Object, zOptB As Object, zChB As Object

AssignD "IsoSetup", , , zChB, zOptB
Set Dim3 = zOptB("o3d"): Set Calc = zChB("cCalculate")

i = zDropD("dIsotype"): d3 = (IsOn(Dim3) And Dim3.Enabled And Dim3.Visible)
OK = (IsOn(Calc) And Calc.Visible And Calc.Enabled And _
      i > 2 And i < 15 And i <> 13)

With zChB("cRobust")
  .Enabled = OK: .Visible = OK
  If Not OK Then .Value = xlOff
End With

Robust = IsOn(zChB("cRobust"))
End Sub

Sub testUserForms() ' Load, format, & show each user form
Dim i

Set DatSht = Ash: Set Oo = Menus("Options")
GetOpSys

For i = 11 To 11

Select Case i
  Case 1: LoadUserForm AddHisto: AddHisto.Show
  Case 2: LoadUserForm Consts: Consts.Show
  Case 3: LoadUserForm DatLab: DatLab.Show
  Case 4: LoadUserForm DCerrsOnly: DCerrsOnly.Show
  Case 5: LoadUserForm FalseClr: FalseClr.Show
  Case 6: LoadUserForm Graphics: Graphics.Show
  Case 7: LoadUserForm Help: Help.Show
  Case 8: LoadUserForm Jinput: Jinput.Show
  Case 9: LoadUserForm Series: Series.Show
  Case 10: LoadUserForm Transp: Transp.Show
  Case 11: LoadUserForm TuffZirc: TuffZirc.Show
  Case 12: LoadUserForm UevoT: UevoT.Show
End Select

Next i

End Sub
Public Function UeT(c As Variant)
Dim v#
v = Val(c)
If v <= 0 Then UeT = "" Else UeT = v
End Function
