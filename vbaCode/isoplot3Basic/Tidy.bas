Attribute VB_Name = "Tidy"
Option Private Module
Option Explicit: Option Base 1
Dim ParamSht As Worksheet
Function RevStr$(ByVal s$)
Attribute RevStr.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i%, T$
For i = Len(s$) To 1 Step -1
  T$ = T$ & Mid(s$, i, 1)
Next i
RevStr$ = T$
End Function
Sub zzz()
Dim A, b, c, d
Set A = ThisWorkbook.DialogSheets
For b = 33 To A.Count
    StatBar Str(b) & "   " & A(b).Name
    DialogShow A(b)
Next
End Sub

Sub ShowDialog(Dbox As Variant, Optional DBvalue)
Dim N$, b As Boolean, Obj As Boolean, NoHelp As Boolean

b = (Not ConcPlot Or Not AutoScale)
Obj = IsObject(Dbox)

If Obj Then N$ = Dbox.Name Else N$ = Dbox

SetDialogBox N$

Select Case N$
  Case "IsoRes":    ConfigIsoRes
  Case "AxLab":     NoHelp = True: ConfigAxLab
  Case "ArStepAge": ConfigArStepAge
  Case "Anch"
  Case Else:        NoHelp = True
End Select

Do
  'On Error GoTo 1 ' In case window-"x" is clicked
  'Canceled = True
  On Error Resume Next
  DBvalue = DlgSht(N$).Show
  On Error GoTo 0
If NoHelp Or Not AskInfo Then Exit Do

  Select Case N$
    Case "IsoRes":    Caveat_Isores
    Case "ArStepAge": Help_ArStepAge
    Case "Anch":      ShowHelp "AnchHelp"
  End Select

Loop

1:
End Sub

Function DialogShow(Dbox As Variant)
Attribute DialogShow.VB_ProcData.VB_Invoke_Func = " \n14"
Dim v As Variant
ShowDialog Dbox, v
DialogShow = v
End Function

Sub SetDialogBox(ByVal DBname$) ' Set the parameters of all controls in all dialog sheets
Attribute SetDialogBox.VB_ProcData.VB_Invoke_Func = " \n14"
Dim P$, s$
s$ = OpSys

If Left(s$, 7) = "Windows" Then
  P$ = "DialogsWin"
ElseIf Left(s$, 9) = "Macintosh" Then
  P = IIf(Int(Version(True)) >= 10, "DialogsMacX", "DialogsMac")
Else
  MsgBox "Unrecognized operating system (" & s$ & ")"
  ExitIsoplot
End If

'p = "DialogsWin"
'Mac = True: p = "DialogsMac": MacExcelX = False
'Mac = True: p = "DialogsMacX": MacExcelX = True
Windows = Not Mac

With TW
  Set ParamSht = .Sheets(P$)
  DboxSet .Sheets(DBname$)
End With

End Sub

Sub DboxSet(DialogSht As DialogSheet) ' Set the parameters of all of the controls in a particular dialog sheet
Attribute DboxSet.VB_ProcData.VB_Invoke_Func = " \n14"
Dim Row%, Pcell As Range, tB As Object, N$

Row = 2
Set Pcell = ParamSht.Cells
N$ = DialogSht.Name
Set tB = DialogSht.TextBoxes

Do

  With Pcell(Row, 1)
    If .Text = N$ And .Font.Bold = True Then Exit Do
  End With

  Row = 1 + Row
Loop

Row = 1 + Row
If tB.Count > 0 Then DtoolSet DialogSht, tB, "TextBoxes", Row
DtoolSet DialogSht, DialogSht.Shapes, "Shapes", Row
End Sub

Sub DtoolSet(DialogSht As DialogSheet, Ctrl As Object, ByVal Na$, Row%)
Attribute DtoolSet.VB_ProcData.VB_Invoke_Func = " \n14"
' Set the parameters of a particular control-type in a secified dialog sheet
Dim Ncontrols%, Control As Object, i%, CtrlNa$, CtrlExists As Boolean, Tx$, MacLabel As Object
Dim Qfalse As Worksheet, Pcell As Range, P%, q%, s$, T$, z$, r%, tB As Boolean, Vis As Boolean
Dim Mna$, j%, ct%, En As Boolean, d As DialogSheet, fs!

Set Qfalse = TW.Sheets("bftsplk")
Set Pcell = ParamSht.Cells
Set d = DialogSht

Do Until Pcell(Row, 1) = Na$
  Row = Row + 1
Loop

fs = -2 * MacExcelX
Ncontrols = Pcell(Row, 2)

For i = Row To Row + Ncontrols - 1
  CtrlExists = False
  Set Control = Qfalse
  CtrlNa$ = Pcell(i, 3).Text
  On Error GoTo Next_i:
  Set Control = Ctrl(CtrlNa$)
  On Error GoTo 0

  With Control
    Vis = .Visible
    If MacExcelX And Na = "TextBoxes" Then .Visible = False
    .Left = Pcell(i, 4).Value
    .Top = Pcell(i, 5).Value
    .Width = Pcell(i, 6).Value - fs
    .Height = Pcell(i, 7).Value

    If Na$ = "TextBoxes" Then
      Tx = .Caption
      En = .Enabled
      If Tx = "ts" & Chr(214) & "MSWD" Then Tx = "tsigmaSqrt(MSWD)"
      ' in case last user went from Mac to Windows
      Mna = "L_" & .Name
      j = 0

      With d ' Delete any previously Mac-added labels
             '  to susbstitute for texboxes.
        Do
          j = j + 1

          If .Labels(j).Name = Mna Then
            .Labels(j).Delete: j = j - 1
          End If

        Loop Until j = .Labels.Count

      End With

     If Not Vis Then GoTo Next_i

     If Not MacExcelX Then
        On Error Resume Next

        With .Font
          .Name = Pcell(i, 8)
          .Size = Pcell(i, 9).Value
          .Bold = Pcell(i, 10)
          .Italic = Pcell(i, 11)
        End With

        ConvertSymbols Control
        .ShapeRange.Fill.ForeColor.RGB = IIf(Windows, RGB(212, 208, 200), RGB(224, 224, 224))

      Else ' Excel 2004 (Mac) does not support dialogsheet textboxes --
           '  must convert to labels.
        Set MacLabel = d.Labels.Add(.Left, .Top, .Width, .Height)

        With MacLabel
          .Name = Mna
          r = InStr(Tx, "(=tsigmasqrtMSWD")

          If r > 0 Then
            .Width = 10 + .Width
            .Left = .Left - 12
            Tx = Left(Tx, r - 1) & "(t*sig*sqrtMSWD)"
          End If

          If d.Name = "Bracket" And Control.Name = "tStrat" Then

          End If

          .Text = Tx: .Enabled = En
          .Visible = Vis
        End With

        .Visible = False
      End If

    End If

  End With

Next_i: On Error GoTo 0

Next i

End Sub

Sub ConfigIsoRes() ' Contract size of Isores dialog box if not an autoscaled concordia plot
Attribute ConfigIsoRes.VB_ProcData.VB_Invoke_Func = " \n14"
Dim t1!, T!, f As Object, G As Object ' (put MonteCarlo controls on top of un-visible concordia ticks controls)
Dim d As Object, c As Object, e As Object, s$, P$, q$(), i%, b As Boolean
Dim L As Object, Sp As Object, ct!, Bu As Object, En As Object, Cm As Object, Op As Object


If ConcPlot And AutoScale Then Exit Sub

AssignD "IsoRes", d, e, c, Op, L, G, , Bu, Sp, Dframe:=f

t1 = Bottom(c("cShowRes")) + 12
T = G("gMC").Top - t1
G("gMC").Top = t1

If ArgonPlot Then
  Set En = e("eNtrials")
  Set Cm = c("cMC")
  Cm.Top = Cm.Top - T
  ct = Cm.Top
  G("gMC").Height = 2.8 * Cm.Height
  c("cLIgtZero").Visible = False
  L("lNtrials").Top = ct
  Sp("sNtrials").Top = ct
  En.Top = ct
  f.Height = Bottom(G("gMC")) + 5 - f.Top

ElseIf ConcPlot And (Not DoPlot Or Not G("gAgeTicks").Visible) Then
  RsT c("cMC"), T:       RsT c("cLIgtZero"), T
  RsT c("cWLE_MC"), T
  RsT L("lNtrials"), T:  RsT e("eNtrials"), T
  RsT G("gHisto"), T:    RsT Sp("sNtrials"), T
  RsT c("cHisto"), T:    RsT e("eNbins"), T
  RsT Sp("sNbins"), T:   RsT L("lNbins"), T
  RsT Op("oLower"), T:   RsT Op("oUpper"), T

  RsT Op("oSeparateSheet"), T
  RsT Op("oDataSheet"), T
  RsT G("gPlotWhere"), T: RsT G("gWhichInter"), T
  f.Height = f.Height - T
Else
  f.Height = Bottom(c("cShowRes")) - f.Top
End If

End Sub

Sub RsT(Tool As Object, ByVal Offs!) ' Shift a dialog-box tool up
Attribute RsT.VB_ProcData.VB_Invoke_Func = " \n14"
With Tool: .Top = .Top - Offs: End With

End Sub

Function Bottom(o As Object) As Single
Attribute Bottom.VB_ProcData.VB_Invoke_Func = " \n14"
Bottom = o.Top + o.Height
End Function

Function Right_(o As Object) As Single
Attribute Right_.VB_ProcData.VB_Invoke_Func = " \n14"
Right_ = o.Left + o.Width
End Function

Sub ConfigArStepAge()
Attribute ConfigArStepAge.VB_ProcData.VB_Invoke_Func = " \n14"
Dim s As Object, c As Object, df As Object, b As Object, L As Object, G As Object

AssignD "ArStepAge", s, , c, , L, G, , b, , , , , , df

If Not DoPlot Or Not ArChron Then
  df.Height = Bottom(c("cShowRes")) - df.Top + 5
End If

b("bHelp").Enabled = ArChron
'L("lJay").Visible = True
End Sub

Sub ConfigAxLab()
Attribute ConfigAxLab.VB_ProcData.VB_Invoke_Func = " \n14"
Dim AxLab As Object, Eb As Object, cb As Object, La As Object, df As Object
Dim tB1 As Boolean, tb2 As Boolean
Dim Tp$, HistoOnly As Boolean, s2$

AssignD "AxLab", AxLab, Eb, cb, , La, Dframe:=df
s2$ = "          X axis"

If CumGauss Then

  If RangeIn(1).Columns.Count = 1 Then
    Tp$ = "Histogram Plot"
    HistoOnly = True
  Else
    Tp$ = "Probability Density Plot"
    HistoOnly = False
  End If

  s2$ = "X-Axis Label"
ElseIf ProbPlot Then
  Tp$ = "Probability Plot"
  s2$ = "Y-Axis Label"
Else
  Tp$ = IIf(Dim3, "X-Y (-Z) Axis Labels", "X-Y Axis Labels")
End If

df.Caption = Tp$
La("lXlabel").Text = s2$
tB1 = Dim3 And Not CumGauss
tb2 = Regress And Not Dim3 And Not CumGauss And Not WtdAvXY
La("lYlabel").Visible = Not CumGauss And Not ProbPlot
Eb("eYlabel").Visible = La("lYlabel").Visible
La("lZlabel").Visible = tB1
Eb("eZlabel").Visible = tB1
La("lLambda1").Visible = tb2
La("lLambda2").Visible = tb2
Eb("eLambda").Visible = tb2

With cb("cInclHist")
  .Visible = CumGauss
  .Enabled = Not HistoOnly
  If HistoOnly Then .Value = xlOn
End With

cb("cAutoBins").Visible = CumGauss
cb("cFilledBins").Visible = CumGauss
La("lBinStart").Visible = CumGauss
Eb("eBinStart").Visible = CumGauss
La("lBinNumWidth").Visible = CumGauss
Eb("eBinNumWidth").Visible = CumGauss
'If HeaderRow Then eb("eXlabel").Text = AxX$
If CumGauss Then Call InclHistClick
Nbins = 0: BinWidth = 0: BinStart = 0

With df

  If tb2 Then
    .Height = Bottom(La("lLambda2")) + 10 - .Top
  ElseIf tB1 Then
    .Height = Bottom(Eb("eZlabel")) + 10 - .Top
  ElseIf CumGauss Then
    .Height = Bottom(Eb("eBinNumWidth")) + 10 - .Top
  Else
    .Height = Bottom(La("lLambda2")) + 10 - .Top
  End If

End With

End Sub

Function ClrIndx(ByVal ClrRGB&)
Attribute ClrIndx.VB_ProcData.VB_Invoke_Func = " \n14"
' Return the color index corresponding to an RGB color value
Dim r%, c%

Select Case ClrRGB

  Case xlNone, xlAutomatic
    ClrIndx = ClrRGB
  Case Is <= 0
    ClrIndx = 0
  Case Else
    On Error Resume Next

    With Menus("ClrRGB")
      r = .Find(ClrRGB).Row

      If r = 0 Then
        ClrIndx = 0
      Else
        ClrIndx = MenuSht.Cells(r, .Column - 1)
      End If

    End With

  End Select

End Function

Sub InitIndex(IndX, ByVal N&)
Dim i&
ReDim IndX(N)

For i = 1 To N
  IndX(i) = i
Next i

End Sub

Sub SortCol(Arr#(), SortedVect#(), ByVal Nrows&, ByVal Col%, _
  Optional IndX, Optional PutBack As Boolean = False)
Dim i&, Ix&()
ReDim SortedVect(Nrows)

For i = 1 To Nrows

  If Col > 0 Then
    SortedVect(i) = Arr(i, Col)
  Else
    SortedVect(i) = Arr(i)
  End If

Next i

If NIM(IndX) Then
  InitIndex Ix, Nrows
  QuickIndxSort SortedVect(), Ix()
  ReDim IndX(Nrows)

  For i = 1 To Nrows
    IndX(i) = Ix(i)
  Next i

Else
  QuickSort SortedVect()
End If

If PutBack Then

  For i = 1 To Nrows
    Arr(i, Col) = SortedVect(i)
  Next i

  On Error Resume Next
  Erase SortedVect
End If

End Sub

Sub QuickIndxSort(Vect#(), IndX&(), _
  Optional ByVal LeftInd& = -2, Optional ByVal RightInd& = -2)

Dim i&, j&, MidInd&, TestVal#
If LeftInd = -2 Then LeftInd = LBound(Vect)
If RightInd = -2 Then RightInd = UBound(Vect)

If LeftInd < RightInd Then
  MidInd = (LeftInd + RightInd) \ 2
  TestVal = Vect(MidInd)
  i = LeftInd
  j = RightInd

  Do

    Do While Vect(i) < TestVal
      i = i + 1
    Loop

    Do While Vect(j) > TestVal
      j = j - 1
    Loop

    If i <= j Then
      SwapElements Vect(), i, j
      Swap IndX(i), IndX(j)
      i = i + 1
      j = j - 1
    End If

  Loop Until i > j

  ' Optimize sort by sorting smaller segment first
  If j <= MidInd Then
    QuickIndxSort Vect, IndX(), LeftInd, j
    QuickIndxSort Vect, IndX(), i, RightInd
  Else
    QuickIndxSort Vect, IndX(), i, RightInd
    QuickIndxSort Vect, IndX(), LeftInd, j
  End If

End If

End Sub

Sub QuickSort(Vect As Variant, _
  Optional ByVal LeftInd& = -2, Optional ByVal RightInd& = -2)

Dim i&, j&, MidInd&, TestVal#

If LeftInd = -2 Then LeftInd = LBound(Vect)
If RightInd = -2 Then RightInd = UBound(Vect)

If LeftInd < RightInd Then
  MidInd = (LeftInd + RightInd) \ 2
  TestVal = Vect(MidInd)
  i = LeftInd
  j = RightInd

  Do
    Do While Vect(i) < TestVal
      i = i + 1
    Loop

    Do While Vect(j) > TestVal
      j = j - 1
    Loop

    If i <= j Then
      SwapElements Vect, i, j
      i = i + 1
      j = j - 1
    End If

  Loop Until i > j

  ' Optimize sort by sorting smaller segment first
  If j <= MidInd Then
    QuickSort Vect, LeftInd, j
    QuickSort Vect, i, RightInd
  Else
    QuickSort Vect, i, RightInd
    QuickSort Vect, LeftInd, j
  End If

End If

End Sub

' Used in QuickSort function
Private Sub SwapElements(Items As Variant, ByVal Item1&, ByVal Item2&)
Dim temp#
temp = Items(Item2)
Items(Item2) = Items(Item1)
Items(Item1) = temp
End Sub

Function RefEditVal(Ctrl As Object) As Double
Attribute RefEditVal.VB_ProcData.VB_Invoke_Func = " \n14"
Dim v, s$

v = BadT
On Error GoTo 1

With Ctrl
  If IsNumber(.Text) Then
    v = .Value
  Else
    s$ = Range(Ctrl.Text).Text
    If IsNumber(s$) Then v = Val(s$)
  End If

End With

1: On Error GoTo 0
RefEditVal = v
End Function

Sub SetUform(u As Object, Uname$)
Dim f, s$, i, Got As Boolean

f = Array("AddHisto", "Consts", "DatLab", "DCerrsOnly", "FalseClr", "Graphics", _
  "Help", "Jinput", "Series", "Transp", "TuffZirc", "UevoT")
Got = False

Do
  s$ = InputBox("Name of user form?")
  If s$ = "" Then KwikEnd

  For i = 1 To UBound(f)

    If LCase(f(i)) = LCase(s$) Then
      Got = True
      Exit For
    End If

  Next i
Loop Until Got

Uname$ = f(i)

Select Case i
  Case 1:  Set u = AddHisto
  Case 2:  Set u = Consts
  Case 3:  Set u = DatLab
  Case 4:  Set u = DCerrsOnly
  Case 5:  Set u = FalseClr
  Case 6:  Set u = Graphics
  Case 7:  Set u = Help
  Case 8:  Set u = Jinput
  Case 9:  Set u = Series
  Case 10: Set u = Transp
  Case 11: Set u = TuffZirc
  Case 12: Set u = UevoT
End Select

End Sub

Sub GetFormInfo()
' Read the parameters of each control of a specified User Form and store then
'  in the appropriate (preexisting) range in the UserFrms sheet.

Dim u As Object, f, Uname$, i, j, NaR As Range, G As UserForm
Dim cOff%, r$(), N%, Rr As Range

Set Oo = Menus("options")
Set DatSht = Ash
GetOpSys

SetUform u, Uname$

cOff = IIf(Left(OpSys, 7) = "Windows", 0, 7)
TW.Sheets("UserFrms").Activate

With u
  N = .Controls.Count
  ReDim r(N + 1, 7), Na(N + 1, 1)
  Na(1, 1) = .Name
  r(1, 1) = .Font.Name:   r(1, 2) = tSt(.Font.Size)
  r(1, 3) = tSt(.Height): r(1, 4) = tSt(.Left)
  r(1, 5) = tSt(.Top):    r(1, 6) = tSt(.Width)
  On Error Resume Next

  For i = 0 To N - 1
    j = i + 2

    With .Controls(i)
      Na(j, 1) = .Name
      r(j, 1) = .Font.Name:   r(j, 2) = tSt(.Font.Size)
      r(j, 3) = tSt(.Height): r(j, 4) = tSt(.Left)
      r(j, 5) = tSt(.Top):    r(j, 6) = tSt(.Width)
      r(j, 7) = .Font.Bold
    End With

  Next i

  On Error GoTo 0
End With

With TW.Sheets("UserFrms")
  i = 4

  Do Until Cells(i, 1) = Na(1, 1) Or IsEmpty(Cells(i, 1))
    i = 1 + i
  Loop

  Set NaR = .Range(Cells(i, 1), Cells(i + N, 1))
  Set Rr = .Range(Cells(i, 2 + cOff), Cells(i + N, 8 + cOff))
  NaR = Na
  Rr = r
  Range(Cells(i, 1), Cells(i + N, 15)).Name = Uname$
End With

End Sub

Function Nff(ByVal v#, Optional Sigfigs% = 3, Optional AsNumber As Boolean = False)
Attribute Nff.VB_ProcData.VB_Invoke_Func = " \n14"
Dim f$, b As Boolean
b = (v = 0 Or (v > 0.001 And v < 1000))
f$ = FloatingKluge(v, IIf(b, -1, Sigfigs))
If AsNumber Then Nff = Val(f$) Else Nff = f$
End Function

Sub AssignValue(AssignTo, Textbox As Object, ByVal Description$, ByVal MinVal#, _
 ByVal MaxVal#, DefaultVal#, Optional Refused As Boolean, _
 Optional Nonzero As Boolean = False, Optional Positive As Boolean = False)
Attribute AssignValue.VB_ProcData.VB_Invoke_Func = " \n14"
' Controls acceptable input into UserForm textboxes

Dim s$, P$, v#, UseDef As Boolean, mn$, Mx$

P$ = "" ' -- using default value of " & Nff(DefaultVal, 7)
s$ = Description$ & " required" & P$
UseDef = True
Refused = False

If Not IsNumeric(Textbox) Then
  MsgBox "Numeric input for " & s$, , Iso
Else
  v = Textbox.Value
  MinVal = Drnd(MinVal, 3): MaxVal = Drnd(MaxVal, 3)

  If v = 0 And Nonzero Then
    MsgBox "Nonzero input for " & s$, , Iso
  ElseIf v < 0 And Positive Then
    MsgBox "Positive value for " & s$, , Iso
  ElseIf v < MinVal Or v > MaxVal Then
    mn$ = Nff(MinVal)
    Mx$ = Nff(MaxVal)
    MsgBox Description$ & " must be between " & mn$ & " and " & Mx$ & P$, , Iso
  Else
    UseDef = False
    AssignTo = v
  End If

End If

If NIM(Refused) Then Refused = UseDef
If UseDef Then AssignTo = DefaultVal
If NIM(Refused) Then Refused = UseDef
End Sub

Sub ShowAllDlg()
Attribute ShowAllDlg.VB_ProcData.VB_Invoke_Func = " \n14"
Dim N$, i, j, o(14)
N$ = InputBox("Dialog-sheet name")

AssignD N$, o(1), o(2), o(3), o(4), o(5), o(6), o(7), o(8), _
            o(9), o(10), o(11), o(12), o(13), o(14)

For i = 2 To 13

  For j = 1 To o(i).Count
    o(i)(j).Visible = True
    On Error Resume Next
    o(i)(j).Enabled = True
    On Error GoTo 0
  Next j

Next i

End Sub

Sub ChartMargins(ChartSheet As Object, LeftMargin!, RightMargin!, _
  TopMargin!, BottomMargin!)
Attribute ChartMargins.VB_ProcData.VB_Invoke_Func = " \n14"
' Get approximate margins (unused space, in points)of a chart sheet.
' DOES NOT COUNT THE AXIS-TITLE HEIGHT

Dim s As Object, MinLeft!, MinTop!, MaxLeft!, MaxTop!

With ChartSheet
  MinLeft = .PlotArea.Left
  MaxLeft = Right_(.PlotArea)
  MinTop = .PlotArea.Top
  MaxTop = Bottom(.PlotArea)

  If .Shapes.Count > 0 Then
    On Error Resume Next

    For Each s In .Shapes
      MinLeft = Min(MinLeft, s.Left)
      MaxLeft = Max(MaxLeft, Right_(s))
      MinTop = Min(MinTop, s.Top)
      MaxTop = Max(MaxTop, Bottom(s))
    Next s

    On Error GoTo 0

    With .Axes(xlCategory)
      If .HasTitle Then MaxTop = Max(MaxTop, .AxisTitle)
    End With

  End If

  LeftMargin = MinLeft
  RightMargin = .ChartArea.Width - MaxLeft
  TopMargin = MinTop
  BottomMargin = .ChartArea.Height - MaxTop
End With

End Sub

Sub CropPicture(Pic As Object, ByVal L, ByVal r, ByVal T, ByVal b)
Attribute CropPicture.VB_ProcData.VB_Invoke_Func = " \n14"
' Crop a picture by the specified number of Left/Right/Top/Bottom points

With Pic.PictureFormat
  .CropLeft = L
  .CropRight = r
  .CropTop = T
  .CropBottom = b
End With

End Sub

Sub Greek()
' Replace all spelled-out english Greek characters in a range
'   by their greek-character equivalent.
Dim P%, i%, j%, k%, M%, Le%
Dim r As Range, e, G, Delim, d(2) As Boolean, Cap As Boolean
Dim s$, ss$, s1$, s2$, A$(2), ts$

Delim = Array(32, 10, 37, 40, 41, 44, 59, 45, 48, 49, 50, _
              51, 52, 53, 54, 55, 56, 57, 124, 177)
Const sY = "symbol"
e = Array("alpha", "beta", "gamma", "delta", "epsilon", "kappa", "lambda", "mu", _
    "nu", "omega", "pi", "psi", "phi", "rho", "sigma", "tau", "eta", "zeta", "chi", _
    "iota", "theta", "xi", "upsilon")
G = Array("a", "b", "g", "d", "e", "k", "l", "m", _
    "n", "w", "p", "y", "j", "r", "s", "t", "h", _
    "z", "c", "i", "q", "x", "u")
Set r = Selection

For i = 1 To r.Count
  s$ = r(i).Text

  For j = 1 To UBound(e)
    P = InStr(LCase(s$), e(j))

    If P > 0 Then

      If e(j) = "sigma" And P > 2 Then
        ' Remove dashes/spaces from all "1-sigma/2-sigma" and "1 sigma/2 sigma"
        ss$ = Mid(s$, P - 2, 2)
        s1$ = Left(ss$, 1)
        s2$ = Right(ss$, 1)

        If s1$ = "1" Or s1$ = "2" Then

          If s2$ = " " Or s2$ = "-" Then
            s$ = Left(s$, P - 2) & Mid(s$, P)
            P = P - 1
          End If

        End If

      End If

      Le = Len(e(j))
      d(1) = (P = 1)
      If Not d(1) Then A$(1) = Mid(s$, P - 1, 1)
      d(2) = ((Len(s$) - P) = (Le - 1))
      If Not d(2) Then A$(2) = Mid(s$, P + Le, 1)

      For k = 1 To 2

        If Not d(k) Then
          For M = 1 To UBound(Delim)
            If Asc(A$(k)) = Delim(M) Then
              d(k) = True
              Exit For
            End If

          Next M

        End If

      Next k

      If d(1) And d(2) Then
        Cap = (s$ = UCase(s$))
        s$ = Left(s$, P - 1) & G(j) & Mid(s$, P + Len(e(j)))
        If Cap Then s$ = UCase(s$)

        With r(i)
          .Formula = s$
          .Characters(P, 1).Font.Name = sY
        End With

      End If

    End If

  Next j

Next i

End Sub

Sub SymbCorr(s$, ByVal SymbStart%, ByVal SymbLen%, ByVal G$, r As Range)
With r
  .Characters.Text = Left(s$, SymbStart - 1) & G$ & Mid(s$, SymbStart + SymbLen)
  .Characters(SymbStart, 1).Font.Name = "symbol"
  s$ = .Characters.Text
End With
End Sub

Sub SetCalc(ByVal Calc&) ' Set the Excel Calculation mode
If Workbooks.Count > 0 And Calc <> 0 Then
  On Error Resume Next
  App.Calculation = Calc
End If
End Sub

Function Qcalc&() ' Query the Excel calculation mode
If Workbooks.Count > 0 Then
  Qcalc = App.Calculation
Else
  Qcalc = xlCalculationAutomatic
End If
End Function

Sub MatCopy(ArrToCopy, ArrCopied)
Dim i&, j&, sR&, sc&, nR&, nc&

On Error GoTo Vect
sc = LBound(ArrToCopy, 2)
On Error GoTo 0

sR = LBound(ArrToCopy, 1)
nR = UBound(ArrToCopy, 1)
nc = UBound(ArrToCopy, 2)

ReDim ArrCopied(sR To nR, sc To nc)

For i = sR To nR

  For j = sc To nc
    ArrCopied(i, j) = ArrToCopy(i, j)
Next j, i

Exit Sub

Vect: On Error GoTo 0
VectCopy ArrToCopy, ArrCopied
End Sub

Sub VectCopy(VectToCopy, VectCopied)
Dim i&, s&, N&
s = LBound(VectToCopy, 1)
N = UBound(VectToCopy, 1)
ReDim VectCopied(s To N)

For i = s To N
  VectCopied(i) = VectToCopy(i)
Next i
End Sub

Function Plural(ByVal s$, ByVal Num%)
Plural = s$ & IIf(Num > 1, "s", "")
End Function
'='C:\Documents and Settings\kludwig\Application Data\Microsoft\AddIns\Iso2.49x.xla'!agepb76(0.1)

Sub AssignIsoVars()

Set Oo = Menus("Options")

With Opt
  .PlotboxBorder = Oo(1):         .SheetClr = Oo(2)
  .PlotboxClr = Oo(3):            .AgeTikSymbClr = Oo(4)
  .AgeTikSymbFillClr = Oo(5):     .UseriesIsochClr = Oo(6)
  .UseriesIsochStyle = Oo(7):     .IsochStyle = Oo(8)
  .CurvClr = Oo(9):               .AxisNameFont = Oo(10)
  .AxisNameFontSize = Oo(11):     .AxisTikLabelFont = Oo(12)
  .AxisTikLabelFontSize = Oo(13): .AxisAutoTikSpace = Oo(14)
  .IsochResFont = Oo(15):         .IsochResFontSize = Oo(16)
  .AgeTikFont = Oo(17):           .AgeTikFontSize = Oo(18)
  .IsochResboxShadw = Oo(19):     .IsochResboxRnd = Oo(20)
  .ConcResboxRnd = Oo(21):        .IsochClr = .UseriesIsochClr
  .CurveRes = Oo(22):             .AgeTikSymbol = Oo(23)
  .AxisThickLine = Oo(24):        .ClipEllipse = Oo(25)
  .AxisTickCross = Oo(26):        .AgeTikSymbSize = Oo(27)
  .SimplePlotSymbSize = Oo(28):   .ConcLineThick = Oo(29)
  .AlwaysPlot2sigma = Oo(30):         .EndCaps = Oo(31)
  .IsochLineThick = Oo(32)
End With

Awb.Colors = TW.Colors
AxisLthick = IIf(Opt.AxisThickLine, xlMedium, xlThin)
Application.DisplayStatusBar = True
End Sub

Sub UmontSpinClick()
With DlgSht("ThUage")
  .EditBoxes("eNtrials").Text = .Spinners("sNtrials").Value
End With
End Sub

Function iAverage(ByVal v As Variant) As Double
iAverage = Sum(v) / (UBound(v) - LBound(v) + 1)
End Function

Sub DeleteSheet()    ' Delete active sheet and, if an Isoplot chart,
Dim P As Object, s$  '  the hidden PlotDat sheet.

NoAlerts
Set P = ActiveSheet

If Ash.Type <> xlWorksheet Then
  On Error GoTo NoGo
  s = Ach.SeriesCollection(1).Formula
  s = Mid$(s, 1 + InStr(s, ","))
  s = Left$(s, InStr(s, "!") - 1)

  If Left(s, 7) = "PlotDat" Then
    With Sheets(s)
      If Not .Visible And .Cells(1, 1) = "Source sheet" Then .Delete
    End With

  End If

End If

NoGo: On Error GoTo 0
If Sheets.Count > 1 Then P.Delete
End Sub

Sub d()
Dim DialogSht As DialogSheet, Ctrl As Object, Na$, Row%
Dim Ncontrols%, Control As Object, i%, CtrlNa$, CtrlExists As Boolean, Tx$, MacLabel As Object
Dim Qfalse As Worksheet, Pcell As Range, P%, q%, s$, T$, z$, r%, tB As Boolean, Vis As Boolean
Dim Mna$, j%, ct%, Dna$, ii%
GetOpSys

Set ParamSht = ThisWorkbook.Sheets("dialogsmacx")
Row = 1
Na = "TextBoxes"
Set Qfalse = ThisWorkbook.Sheets("bftsplk")
Set Pcell = ParamSht.Cells

For ii = 1 To ThisWorkbook.DialogSheets.Count
  Set DialogSht = ThisWorkbook.DialogSheets(ii)
  Dna = DialogSht.Name
  Application.StatusBar = Dna

  Do Until Pcell(Row, 1) = Dna: Row = Row + 1: Loop

  Row = 1 + Row

  If Pcell(Row, 1) <> Na Then GoTo NextSht

  Ncontrols = Pcell(Row, 2)

  For i = Row To Row + Ncontrols - 1
    CtrlExists = False
    Set Control = Qfalse
    CtrlNa$ = Pcell(i, 3).Text
    On Error GoTo Next_i:
    Set Control = DialogSht.TextBoxes(CtrlNa$)
    On Error GoTo 0

    With Control
      Vis = .Visible
      .Visible = True
      .Left = Pcell(i, 4).Value
      .Top = Pcell(i, 5).Value
      .Width = Pcell(i, 6).Value
      .Height = Pcell(i, 7).Value

      If Na$ = "TextBoxes" Then
        Tx = .Caption
        If Tx = "ts" & Chr(214) & "MSWD" Then Tx = "tsigmaSqrt(MSWD)"
        Mna = "L_" & .Name
        j = 0

        With DialogSht ' Delete any previously Mac-added labels

          Do           '  to susbstitute for texboxes.
            j = j + 1

            If .Labels(j).Name = Mna Then
              .Labels(j).Delete
              j = j - 1
            End If

          Loop Until j = .Labels.Count

        End With

        Set MacLabel = DialogSht.Labels.Add(.Left, .Top, .Width, .Height)

        With MacLabel
          .Name = Mna
          r = InStr(Tx, "tsigmasqrtMSWD")

          If r > 0 Then
            Tx = Left(Tx, r - 1) & "t*sigma*sqrt[MSWD]"
          End If

          .Text = Tx
        End With

        .Visible = False
      End If

    End With

Next_i:   On Error GoTo 0
  Next i

NextSht:
Next ii

Application.StatusBar = ""
End Sub
