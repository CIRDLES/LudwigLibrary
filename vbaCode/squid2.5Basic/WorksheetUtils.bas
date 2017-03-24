Attribute VB_Name = "WorksheetUtils"
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
' 09/04/15 -- Add the UnhideRehide toolbar button-invoked sub to toggle the Hidden status
'             of CPS and Task Equation columns specified as hidden in the Task definition.
Option Explicit
Option Base 1

Sub NameRange(Rname$, Optional StdOrSam, Optional RowOrRange, Optional Column_, _
    Optional RangeValue, Optional NumFor$, Optional RangeLabel = True, Optional Comment As String)
' Put name range to left of a range, bold, right-aligned.
' If range value & number format are passed, put those in the range.
Dim Row_&

If fbIM(RowOrRange) Then
  Row_ = Range(Rname).Row: Column_ = Range(Rname).Column
Else

  If fbNIM(RowOrRange) And fbIM(Column_) And TypeName(RowOrRange) = "Range" Then
    Row_ = RowOrRange.Row: Column_ = RowOrRange.Column
  Else
    Row_ = RowOrRange
  End If

  If Row_ = 0 Or Column_ = 0 Then
    MsgBox "SQUID error in Sub NameRange -- Row_ = 0 Or Column_ = 0", , pscSq
    CrashEnd
  End If

  If Rname <> "" Then AddName Rname, (StdOrSam), (Row_), (Column_)
  Cells(Row_, Column_) = RangeValue
  Fonts Row_, Column_, , , , False, xlLeft
  If fbNIM(NumFor) Then Call RangeNumFor((NumFor), Row_, Column_)
End If

If RangeLabel Then

  If fbNIM(StdOrSam) Then
    If StdOrSam Then phStdSht.Activate Else phSamSht.Activate
  End If

  Cells(Row_, Column_ - 1) = Rname
  Fonts Row_, Column_ - 1, , , , True, xlRight
End If

If fbNIM(Comment) And Comment <> "" Then Note Row_, Column_, Comment
End Sub

Function flHeaderRow&(ByVal ForStd As Boolean, Optional UpperRowLim& = 20, _
  Optional NoErr As Boolean = False, Optional NoActivate As Boolean = False)
' Find header-row (first double-underlined)
Dim GotSheet As Boolean, i%, Nsheets%, RowCt&, SheetIn As Worksheet, Sht As Worksheet

Set SheetIn = ActiveSheet
GotSheet = False
If NoErr Then On Error Resume Next

If ForStd And Not NoActivate And ActiveSheet.Name <> pscStdShtNa Then
  On Error Resume Next

  For Each Sht In ActiveWorkbook.Worksheets
    If Sht.Name = pscStdShtNa Then
      GotSheet = True: Sht.Activate: Exit For
    End If
  Next Sht

  If Not GotSheet Then
    MsgBox "Unable to activate the StandardData worksheet.", , pscSq
    End
  End If
End If

On Error GoTo 0
RowCt = 1 ' Find first double-underlined cell, first/last data-rows

Do Until Cells(RowCt, 1).Borders(xlBottom).LineStyle = xlDouble
  RowCt = RowCt + 1
  If RowCt > UpperRowLim Then flHeaderRow = 0: Exit Function
Loop

plaFirstDatRw(-ForStd) = 1 + RowCt
flHeaderRow = RowCt
plaLastDatRw(-ForStd) = flEndRow
SheetIn.Activate
Exit Function

1: On Error GoTo 0
CrashNoise
CrashEnd
End Function

Function fiEndCol%(ByVal AtRow&, Optional Wkbook, Optional WkSheet)
' Find last occuped column
Dim Nk$, Nh$
If AtRow = 0 Then AtRow = plHdrRw
VIM Wkbook, ActiveWorkbook
VIM WkSheet, ActiveSheet
Nk = Wkbook.Name: Nh = WkSheet.Name
Worksheets(Nh).Activate
fiEndCol = Workbooks(Nk).Worksheets(Nh).Cells(AtRow, _
           peMaxCol).End(xlToLeft).Column
End Function

Function flEndRow&(Optional ByVal Column_% = 1, Optional Wkbook, Optional WkSheet)
' Find last occuped row
Dim Nk$, Nh$
VIM Wkbook, ActiveWorkbook
VIM WkSheet, ActiveSheet
Nk = Wkbook.Name: Nh = WkSheet.Name

flEndRow = Workbooks(Nk).Worksheets(Nh).Cells(pemaxrow, _
            Column_).End(xlUp).Row
End Function

Function fsVertToLF$(ByVal s$) ' Replace | char with linefeed char
fsVertToLF = fsApSub("|", s$, vbLf)
End Function

Sub BorderLine(ByVal BrdrInd%, ByVal NumLines%, rw1 As Variant, _
  Optional Col1, Optional rw2, Optional Col2)
Dim Line_Style% ' Put specified border-line in a range

Select Case NumLines
  Case 2: Line_Style = xlDouble
  Case xlNone: Line_Style = xlNone
  Case Else: Line_Style = xlContinuous
End Select

frSr(rw1, Col1, rw2, Col2).Borders(BrdrInd).LineStyle = Line_Style
End Sub

Sub RangeNumFor(ByVal NumFormat$, Optional rw1 = 0, Optional Col1 = 0, _
  Optional rw2 = 0, Optional Col2 = 0)
' Specify number format of a range
DefVal rw1, 0
DefVal Col1, 0
DefVal rw2, 0
DefVal Col2, 0

If fbIM(Col1) Or Col1 = 0 Then
  If TypeName(rw1) <> "Range" Then Exit Sub
  frSr(rw1).NumberFormat = NumFormat$
ElseIf Not (rw1 = 0 And Col1 = 0) Then
  frSr(rw1, Col1, rw2, Col2).NumberFormat = NumFormat ' Utility for cleaner code
End If

End Sub

Sub AddName(ByVal Name$, ByVal Std As Boolean, ByVal rw1&, ByVal Col1%, _
  Optional rw2&, Optional Col2%)
Dim RangeAddr$  ' For cleaner code - Adds whether Standard or Sample worksheet
Dim Sht As Worksheet, TempRange As Range

CheckRC 110, rw1, Col1
If Std Then Set Sht = phStdSht Else Set Sht = phSamSht
Set TempRange = frSr(rw1, Col1, rw2, Col2)
RangeAddr = TempRange.Address
Sht.Names.Add Name$, "=" & RangeAddr
End Sub

Sub StatBar(Optional Phrase, Optional Count)
Dim Spaces$
With Application
  If Not .DisplayStatusBar Then .DisplayStatusBar = True

  If fbIM(Phrase) Then
    .StatusBar = ""
  Else
    Spaces = Space(10)

    If fbIM(Count) Then
      Spaces = Spaces & Phrase & " . . ."
    Else
      Spaces = Spaces & Phrase & Space(fvMax(5, 30 - Len(Phrase))) & fsS(CLng(Count))
    End If

    .StatusBar = Spaces
  End If

End With
End Sub

Sub FreezeValues() ' Convert formulae to values for active sheet
Dim Col%, RowCt_a%, RowCt_b%, EmptyCellCt%, a&, Shp As Object

If Workbooks.Count = 0 Then Exit Sub
GetInfo
If ActiveSheet.Type <> xlWorksheet Then Exit Sub
NoUpdate
RowCt_a = 1: Col = 1

Do
  RowCt_a = 1 + RowCt_a
  If RowCt_a = 20 Then Exit Sub
Loop Until Cells(RowCt_a, 1).Borders(xlBottom).LineStyle = xlDouble

foAp.CalculateBeforeSave = False ' !*** Otherwise will CRASH Excel when workbook is saved ***!

With Cells
  '.MergeCells = False ' 09/06/18 -- commented out
  .Copy
  .PasteSpecial Paste:=xlValues
End With

foAp.CutCopyMode = False
RowCt_b = 0: EmptyCellCt = 0

Do
  RowCt_b = 1 + RowCt_b
  EmptyCellCt = EmptyCellCt - IsEmpty(Cells(RowCt_b, 1))
  If Rows(RowCt_b).Hidden Then Rows(RowCt_b).Delete: RowCt_b = RowCt_b - 1
Loop Until EmptyCellCt = 50

For RowCt_b = 1 To peMaxCol
  If Columns(RowCt_b).Hidden Then Columns(RowCt_b).Delete: RowCt_b = RowCt_b - 1
Next RowCt_b

For Each Shp In ActiveSheet.Shapes ' Delete buttons
  With Shp
    If .OnAction <> "" Then .Delete
  End With
Next

For Each Shp In ActiveSheet.Comments: Shp.Delete: Next
Cells(RowCt_a, Col).Select
End Sub

Sub FreezeAll() ' Convert formulae to values for all sheets
Dim Sht As Worksheet

If Workbooks.Count = 0 Then Exit Sub
NoUpdate

For Each Sht In Worksheets
  With Sht

    If .Visible Then
      StatBar .Name
      .Select
      FreezeValues
    End If

  End With
Next Sht

ActiveSheet.Activate
StatBar
End Sub

Sub CheckRC(ByVal ErrNum%, Optional rw1, Optional Col1, Optional rw2, Optional Col2)
Dim ErInc% ' Are the Row# and Col# specified for a range valid?

If fbNIM(rw1) And TypeName(rw1) <> "Range" Then
  If rw1 <= 0 Or rw1 > pemaxrow Then ErInc = 1
End If

If fbNIM(Col1) And TypeName(Col1) <> "Range" Then
  If (Col1 <= 0 Or Col1 > peMaxCol) Then ErInc = 2
End If

If fbNIM(rw2) Then
  If rw2 <= 0 Or rw2 > pemaxrow Then ErInc = 3
End If

If fbNIM(Col2) And TypeName(Col2) <> "Range" Then
  If Col2 <= 0 Or Col2 > peMaxCol Then ErInc = 4
End If

If ErInc Then
  MsgBox "SQUID error in Sub CheckRC, ErInc=" & StR(ErInc), , pscSq
  End
End If

End Sub

Function fbNIM(ByVal Param) As Boolean ' Not ismissing
fbNIM = Not fbIM(Param)
End Function

Function fbIM(ByVal Param) As Boolean ' ismissing
fbIM = (IsMissing(Param) Or IsNull(Param))
End Function

Sub VIM(Param, DefaultVal, Optional Ndim = 0)
Dim i%, j%, Lower%(2), Upper%(2)

If fbNIM(Param) Then Exit Sub

Select Case Ndim
  Case 0:

    If IsObject(DefaultVal) Then
      Set Param = DefaultVal
    Else
      Param = DefaultVal
    End If

  Case 1
    Lower(1) = LBound(DefaultVal): Upper(1) = UBound(DefaultVal)
    ReDim Param(Lower(1) To Upper(1))

    For i = Lower(1) To Upper(1)
      Param(i) = DefaultVal(i)
    Next i

  Case 2

    For i = 1 To 2
      Lower(i) = LBound(DefaultVal, i)
      Upper(i) = UBound(DefaultVal, i)
      ReDim Param(Lower(1) To Upper(1), Lower(2) To Upper(2))
    Next i

    For i = Lower(1) To Upper(1)
      For j = Lower(2) To Upper(2)
        Param(i, j) = DefaultVal(i, j)
    Next j, i

  Case Else
    MsgBox "Bad VIM call": End
End Select

End Sub

Sub IntClr(ByVal Clr&, Optional rw1, Optional Col1, Optional rw2, Optional Col2)
' 10/03/12 -- Rewrite error handling - would logically report error no matter what !!!!
Dim Msg$
Msg = "Squid error in Sub IntClr -- "

If TypeName(rw1) <> "Range" Then
  If fbNIM(rw1) And rw1 <= 0 Then Msg = Msg & "rw1<=0":    GoTo Bad
  If fbNIM(Col1) And Col1 <= 0 Then Msg = Msg & "Col1<=0": GoTo Bad
End If

On Error GoTo Bad
frSr(rw1, Col1, rw2, Col2).Interior.Color = Clr ' Set interior color of a range
Exit Sub

Bad: On Error GoTo 0
MsgBox Msg, , pscSq: CrashEnd
End Sub

Sub Fonts(Optional rw1, Optional Col1, Optional rw2, Optional Col2, _
  Optional Clr, Optional Bold, Optional HorizAlign, _
  Optional Size, Optional StrkThru, Optional Underline, Optional Formul, _
  Optional NumFormat, Optional FontName, Optional MergeTheCells, _
  Optional italic, Optional Phrase, Optional InteriorColor)
' 09/05/06 -- Add the optional "InteriorColor" parameter
Dim TmpRange As Range ' specify attributes of a range

If fbIM(Col1) Then
  Set TmpRange = rw1
Else
  Set TmpRange = frSr(rw1, Col1, rw2, Col2)
  On Error Resume Next
End If

On Error Resume Next
With TmpRange
  With .Font
    If fbNIM(Clr) Then .Color = Clr
    If fbNIM(Bold) Then .Bold = Bold
    If fbNIM(Size) Then .Size = Size
    If fbNIM(StrkThru) Then .Strikethrough = StrkThru
    If fbNIM(FontName) Then .Name = FontName
    If fbNIM(italic) Then .italic = italic
    If fbNIM(Underline) Then .Underline = Underline
  End With
  If fbNIM(HorizAlign) Then .HorizontalAlignment = HorizAlign
  If fbNIM(Formul) Then .Formula = Formul
  If fbNIM(NumFormat) Then .NumberFormat = NumFormat
  If fbNIM(MergeTheCells) Then .MergeCells = MergeTheCells
  If fbNIM(Phrase) Then .Formula = Phrase
  If fbNIM(InteriorColor) Then .Interior.Color = InteriorColor
End With

On Error GoTo 0
End Sub

Function frSr(Optional rw1 As Variant, Optional ByVal Col1 = 0, Optional ByVal rw2 = 0, _
  Optional ByVal Col2, Optional WkSht) As Range  ' specify a range
' Return a range regardless of how specified
' Forbidden: zero values for both rw1 and Col1
Dim RowOrRange As Boolean, TmpRange As Range

RowOrRange = fbNIM(rw1)
DefVal Col1, 0
DefVal Col2, 0
DefVal rw1, 0
DefVal rw2, 0

If RowOrRange And TypeName(rw1) = "Range" Then
  Set TmpRange = rw1
ElseIf RowOrRange And TypeName(rw1) = "String" Then
  Set TmpRange = Range(rw1)
Else
  If rw2 = 0 And rw1 > 0 Then rw2 = rw1
  If Col2 = 0 And Col1 > 0 Then Col2 = Col1

  If rw1 = 0 Then
    Set TmpRange = Range(Columns(Col1), Columns(Col2))
  ElseIf Col1 = 0 Then
    Set TmpRange = Range(Rows(rw1), Rows(rw2))
  Else
    Set TmpRange = Range(Cells(rw1, Col1), Cells(rw2, Col2))
  End If

End If

If fbNIM(WkSht) Then
  Set frSr = WkSht.Range(TmpRange.Address)
Else
  Set frSr = TmpRange
End If

End Function

Sub Pad(ByVal c%, ByVal rw1&, ByVal rw2&)
' Pad a numeric string with leading & trailing spaces so that a fixed-width font
'  looks correctly justified with respect to decimal points.

Dim HasPm As Boolean, s$
Dim t%, i%, Dr%, dL%, MdR%, MdL%, d%
Dim r As Range

CheckRC 140, rw1, , rw2
Set r = frSr(rw1, c, rw2)

For t = 1 To 2

  If t = 1 Then
    With r

      If .NumberFormat <> "Null" Then
        HasPm = (InStr(.NumberFormat, Chr(177)) <> 0)
        .NumberFormat = "@"
      End If

    End With
  End If

  For i = 1 To r.Cells.Count

    If Not IsEmpty(r(i)) Then
      s$ = Trim(r(i)): d = InStr(s$, ".")

      If d > 0 Then
        dL = Len(Left$(s$, d - 1)): Dr = Len(s$) - d + 1
      Else
        dL = Len(s$): Dr = 0
      End If

      If t = 1 Then
        MdL = fvMax(MdL, dL): MdR = fvMax(MdR, Dr)
      Else
        s$ = Space(MdL - dL) & s$ & Space(MdR - Dr)
        If HasPm Then s$ = Chr(177) & s$
        r(i) = s$
      End If

    End If

  Next i

Next t

End Sub

Sub Box(ByVal FirstRow As Variant, Optional FirstCol, _
  Optional LastRow, Optional piLastCol, Optional Clr, _
  Optional dbl = False, Optional None = False, _
  Optional Wkbook, Optional WkSheet, Optional SingleThick = False)
' 09/10/08 -- added the SingleThick parameter

Dim i%, r As Range, Nk$, Nh$
VIM Wkbook, ActiveWorkbook
VIM WkSheet, ActiveSheet
Nk = Wkbook.Name: Nh = WkSheet.Name

If TypeName(FirstRow) = "Range" Then
  Set r = FirstRow
ElseIf TypeName(FirstRow) = "String" Then
  Set r = Workbooks(Nk).Worksheets(Nh).Range(FirstRow)
Else
  CheckRC 150, FirstRow, FirstCol, LastRow, piLastCol
  DefVal LastRow, FirstRow
  DefVal piLastCol, FirstCol
  Set r = Workbooks(Nk).Worksheets(Nh).Range(frSr(FirstRow, _
          FirstCol, LastRow, piLastCol).Address)
End If

With r
  If fbNIM(Clr) Then .Interior.Color = Clr

  For i = 7 To 10
    With .Borders(i)

      If dbl Then
        .LineStyle = xlDouble
      ElseIf None Then
        .LineStyle = xlLineStyleNone
      ElseIf SingleThick Then
        .Weight = xlMedium
      Else
        .Weight = xlThin
      End If

    End With
  Next i

End With
End Sub

Function fsLSN$(ByVal Na$) ' Replace illegal sheet-name characters
Dim f As String * 6, r As String * 6, i%
f = "/\:*?-": r = "><|~&@"

For i = 1 To Len(f)
  Subst Na$, Mid$(f, i, 1), Mid$(r, i, 1)
Next i

fsLSN = Na$
End Function

Sub StdResFmt(ByVal r&, ByVal cc%, ByVal Cv%, ByVal Caption$, Optional Content, Optional NumFor, Optional HorizAlign = xlLeft)
CheckRC 160, r, cc ' Place a number-formatted value in a validity-checked range
Cells(r, cc).Formula = Caption$
Fonts r, Cv, , , , , HorizAlign, , , , Content, NumFor
End Sub

Function fsRa$(ByVal plaFirstDatRw%, ByVal Col1%, Optional Row2% = 0, _
  Optional Col2% = 0, Optional Abso% = 0) ' Specify a range, relative or absolute
Dim RowAbs As Boolean, ColAbs As Boolean

RowAbs = (Abso = 1 Or Abso = 3)
ColAbs = (Abso = 2 Or Abso = 3)
If Row2 = 0 Then Row2 = plaFirstDatRw
If Col2 = 0 Then Col2 = Col1

fsRa = Range(Cells(plaFirstDatRw, Col1), Cells(Row2, Col2)).Address(RowAbs, ColAbs)
End Function

Sub CF(ByVal Row&, ByVal Col%, ByVal v#, Optional Perr)
Dim Mult# ' Fill cell with a double-prec. value, optionally multiplied by 100

If Col = 0 Or v = pdcErrVal Or v = CSng(pdcErrVal) Then Exit Sub

Mult = 1
If fbNIM(Perr) Then
  If Perr <> 0 Then Mult = 100
End If
Cells(Row, Col) = Mult * v
End Sub

Sub PlaceHdr(ByVal Row&, ByVal Col%, ByVal HdrName$, Optional Comment)
' 09/07/02 -- created
Dim r&, Hdr$, ShtIn As Worksheet, rs%, Test$, Le%, LeH%
If Col = 0 Then Exit Sub

Set ShtIn = ActiveSheet
fhColIndx.Activate
HdrName = Trim(HdrName)
LeH = Len(HdrName)
rs = 1

Do
  FindStr HdrName, r, , rs, 1, 999, 1
  If r = 0 Then
    ShtIn.Activate
    Exit Sub
  End If
  Test = Trim(Cells(r, 1))
  Le = Len(Test)
  rs = 1 + r
Loop Until Le = LeH

Hdr = Trim(Cells(r, 3))
ShtIn.Activate
CFs Row, Col, Hdr, True, Comment
End Sub

Sub CFs(ByVal Row&, ByVal Col%, ByVal CellContents$, _
  Optional Lfs = False, Optional Comment) ' Specify cell contents , error-checked & other things
  Dim tc$, se
If Col = 0 Or CellContents = "" Then Exit Sub ' probably an undefined column-name

If IsNumeric(CellContents) Then
  tc = Trim(CellContents)
  If tc = fsS(pdcErrVal) Or tc = fsS(CSng(pdcErrVal)) Then Exit Sub
End If

If Lfs Then ' Specify contents of a cell
  Cells(Row, Col) = fsVertToLF(CellContents)
Else
  Cells(Row, Col) = CellContents
End If

If fbNIM(Comment) Then Note Row, Col, (Comment)
End Sub

Sub Sheetname(Na$) ' If sheet name exists, add " #" to it.
Dim Dup As Boolean, Sht$, s$, Ls%, Ln%, N#, m#, Sh As Worksheet
Ln = Len(Na$): m = 0: Dup = False

For Each Sh In Worksheets
  Sht$ = LCase(Sh.Name)
  Ls = Len(Sht$)

  If Ls >= Ln Then

    If LCase(Left$(Sht$, Ln)) = LCase(Na$) Then
      Dup = True
      s$ = Mid$(Sht$, 1 + Ln)
      If IsNumeric(s$) Then N = Val(s$) Else N = 0
      m = fvMax(m, N)
    End If

  End If

Next Sh

If Dup Then Na$ = Na$ & StR(1 + m)
End Sub

Sub DelCol(Col1%, Optional Col2) 'delete one or more colums
CheckRC 190, , Col1, , Col2
If fbIM(Col2) Then
  Columns(Col1).Delete
Else
  Range(Columns(Col1), Columns(Col2)).Delete
End If
End Sub

Sub AddButton(ByVal Row, ByVal Col, ByVal Name$, ByVal Phrase$, ByVal Macro$, _
  Optional Clr, Optional FontSize! = 11, Optional ShadowOffs = 2.5)
' To a worksheet, add a button calling a SQUID sub.
Dim c%, r&, L#, t#, s As Object

r = Row: c = Col
DefVal Clr, peLightGray

If r >= 0 And c >= 0 Then
  With Cells(r, c)
    L = .Left + (Col - c) * .Width
    t = .Top + (Row - r) * .Height
  End With
Else
  L = -r: t = -c
End If

With ActiveSheet

  For Each s In .Shapes ' Only add if no same-named shape already exists
    If s.Name = Name$ Then Exit Sub
  Next s

  .Shapes.AddShape(msoShapeBevel, L, t, 1, 1).Select

  With Selection
    .Characters.Text = Phrase$: .AutoSize = True
    .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    c = InStr(Macro, "!")
    If c Then Macro = Mid$(Macro, 1 + c)
    .OnAction = Macro
    .Name = Name$
    With .Font
      .Name = "Arial": .Size = FontSize
    End With

    With .ShapeRange
      .Fill.ForeColor.RGB = Clr: .Line.Weight = 1
      With .Shadow
        .Type = msoShadow14
        .IncrementOffsetX ShadowOffs
        .IncrementOffsetY ShadowOffs
      End With
    End With

  End With

End With

Cells(Row, Col).Activate
End Sub

Sub AddRedoButton(ByVal r, ByVal c)    ' Add a "ReDo" button
CheckRC 220, r, c
AddButton r, c, "Redo", " Redo ", "Redo"
End Sub

Function flVisRow&(ByVal r&) ' next or current not-hidden row

Do While Rows(r).Hidden
  r = 1 + r
Loop

flVisRow = r
End Function

Sub DelSheet(Optional Sht)
If fbIM(Sht) Then Set Sht = ActiveSheet

If Sheets.Count > 1 Then
  Alerts False: Sht.Delete: Alerts True
End If

End Sub

Sub Cbars(ByVal CbarName$, ByVal Visible As Boolean) ' Make a specified command-bar visible
On Error Resume Next
foAp.CommandBars(CbarName$).Visible = Visible
On Error GoTo 0
End Sub

Sub CleanupSquidRefs(Optional ShowMsgbox As Boolean = True)
' Delete all references to Squid*.xla! and Isoplot*,xla so that all worbook references are
'  to the squid/isoplot on the current computer.
' 09/05/09 -- Add the optional ShowMsgBox parameter and UpdatedRefs variable for use with
'             the WorkbookOpen even-handler call to this sub.
Dim EncounteredError As Boolean
Dim CellFormula$, ReferenceToDelete$, tmp$, ErrMsg$, NewCellFormula$, d As String * 1
Dim ShtType%, p%, q%, LocCt%, FoundCellCol%, tB As Boolean, UpdatedRefs As Boolean
Dim LastFoundCellCol%, ArrayFormulaCol1%, ArrayFormulaCol2%, InstLoc%()
Dim FoundCellRow&, LastFoundCellRow&, ArrayFormulaRow1&, ArrayFormulaRow2&
Dim ShtIn As Variant, WkSht As Worksheet, c As Range, Buttn As Object

If Workbooks.Count = 0 Then Exit Sub

NoUpdate
ManCalc    ' Must not change back!
Set ShtIn = ActiveSheet
d = "\"
EncounteredError = False

For Each WkSht In ActiveWorkbook.Worksheets
  With WkSht
    ShtType = 0: On Error Resume Next: ShtType = .Type: On Error GoTo 0
    .Activate
    StatBar "Updating " & .Name
    .Visible = True
    FoundCellRow = 1:     FoundCellCol = 1
    LastFoundCellRow = 0: LastFoundCellCol = 0

    Do While FoundCellRow > LastFoundCellRow Or _
        (FoundCellRow = LastFoundCellRow And FoundCellCol > LastFoundCellCol)
      Cells(FoundCellRow, FoundCellCol).Activate
      LastFoundCellRow = FoundCellRow: LastFoundCellCol = FoundCellCol
      Set c = Cells.Find("'", ActiveCell, xlFormulas, xlPart, xlByRows, xlNext, False)
      If LCase(TypeName(c)) <> "range" Then Exit Do
      c.Select
      CellFormula = LCase(CStr(c.Formula)) ' 09/06/13 -- added the "Cstr" to prevent crash if contains an error
      FoundCellRow = ActiveCell.Row: FoundCellCol = ActiveCell.Column
      StatBar "Updating " & .Name & "   " & StR(FoundCellRow) & "   " & StR(FoundCellCol)

      Do
        AllInstanceLoc "'", CellFormula, InstLoc, LocCt
        If LocCt < 2 Then Exit Do
        tmp = Mid$(CellFormula, InstLoc(1), InstLoc(2) - InstLoc(1))
        ' 09/06/13 -- added "squid.xla" to line below
        tB = InStr(CellFormula, "isoplot") = 0 And _
             InStr(CellFormula, "squid2") = 0 And _
             InStr(CellFormula, LCase(pscSq)) = 0 And _
             InStr(CellFormula, "squid.xla") = 0

        If tB Then Exit Do

        UpdatedRefs = True

        For q = 1 To LocCt - 1 Step 2
          p = InStr(CellFormula, ".xla'")

          If p = (InstLoc(q + 1) - 4) And InstLoc(q) < p Then
            ReferenceToDelete = fsSubStr(CellFormula, InstLoc(q), 1 + InstLoc(q + 1))

            If InStr(ReferenceToDelete, "squid") > 0 Or _
               InStr(ReferenceToDelete, "isoplot") > 0 Then
              NewCellFormula = CellFormula
              Subst NewCellFormula, ReferenceToDelete
              ErrMsg = ""
              On Error Resume Next
              Cells(FoundCellRow, FoundCellCol).Formula = NewCellFormula
              CellFormula = NewCellFormula
              ErrMsg = Error
              On Error GoTo 0

              If ErrMsg = "You cannot change part of an array." Then
                ArrayFormulaRow1 = FoundCellRow: ArrayFormulaCol1 = FoundCellCol
                ArrayFormulaRow2 = FoundCellRow: ArrayFormulaCol2 = FoundCellCol

                Do While Cells(1 + ArrayFormulaRow2, FoundCellCol).Formula = CellFormula
                  ArrayFormulaRow2 = 1 + ArrayFormulaRow2
                Loop

                Do While Cells(FoundCellRow, 1 + ArrayFormulaCol2).Formula = CellFormula
                  ArrayFormulaCol2 = 1 + ArrayFormulaCol2
                Loop

                frSr(ArrayFormulaRow1, ArrayFormulaCol1, ArrayFormulaRow2, _
                   ArrayFormulaCol2).FormulaArray = NewCellFormula
              ElseIf ErrMsg <> "" Then
                EncounteredError = True
              End If

            End If

          End If

        Next q

      Loop Until LocCt < 3

    Loop

    For Each Buttn In .Shapes
      With Buttn
        CellFormula = LCase(.OnAction)

        If CellFormula <> "" Then
          AllInstanceLoc "'", CellFormula, InstLoc, LocCt

          If LocCt > 1 Then
            p = InStr(CellFormula, ".xla'")
            If p = (InstLoc(2) - 4) And InstLoc(1) < p Then
              On Error Resume Next
              ReferenceToDelete = fsSubStr(CellFormula, InstLoc(1), 1 + InstLoc(2))
              NewCellFormula = CellFormula
              Subst NewCellFormula, ReferenceToDelete
              .OnAction = NewCellFormula
            End If

          End If

          On Error GoTo 0
        End If
      End With
    Next Buttn

  End With 'WkSht
Next WkSht

ShtIn.Activate
StatBar

If EncounteredError Then
  MsgBox "Unable to update all Isoplot & Squid references.", , pscSq
Else
  If UpdatedRefs Or ShowMsgbox Then MsgBox "References to Squid and Isoplot updated.", , pscSq
End If

End Sub

Sub WidenCols()
ChangeCols 1
End Sub

Sub ChangeCols(ByVal DeltaWidth!, Optional ByVal EndColRow)
Dim c%, cc%, e%, w1!, w2!
If Workbooks.Count = 0 Then Exit Sub
NoUpdate
If fbIM(EndColRow) Then e = 255 Else e = fiEndCol((EndColRow))

For c = e To 1 Step -1
  With Columns(c)

    If .ColumnWidth >= 1 And Not .Hidden Then
      w1 = .ColumnWidth
      w2 = w1 + DeltaWidth
      If w1 <> 0 Then .ColumnWidth = fvMax(0.1, w2)
    End If

  End With
Next c

End Sub

Sub NarrowCols()
ChangeCols -1
End Sub

Sub Note(r As Variant, Optional c, Optional s$) ' Add a comment to a range
Dim Rng As Range
If TypeName(r) = "Range" Then Set Rng = r Else Set Rng = Cells(r, c)

With Rng
  On Error Resume Next
  .AddComment
  With .Comment
    .Text Text:=s
    If Len(s) > 80 Then .Shape.Width = 100
  End With
  On Error GoTo 0
End With

End Sub

Sub ScrollW(ByVal Row&, ByVal Col%) ' Specify scrollrow & scrollcolumn
If Row < 1 Or Col < 1 Then Exit Sub

With ActiveWindow
  On Error Resume Next
  .ScrollRow = Row: .ScrollColumn = Col
End With
End Sub

Sub Freeze(Optional ByVal Yes As Boolean = True)
With ActiveWindow
  .FreezePanes = False
  If Yes Then .FreezePanes = True
End With
End Sub

Sub Zoom(ByVal z%)
On Error Resume Next
ActiveWindow.Zoom = z
On Error GoTo 0
End Sub

Function fhTempTask() As Worksheet
Set fhTempTask = ThisWorkbook.Sheets("TempTask")
End Function

Function fhTempCat() As Worksheet
Set fhTempCat = ThisWorkbook.Sheets("TempCat")
End Function

Function frUserTaskList(ByVal IsUPb As Boolean) As Range
Set frUserTaskList = foUser("Last" & IIf(IsUPb, "UPb", "General") & "Tasks")
End Function

Function foUser(Optional ByVal Range_, Optional ByVal NoObjectBack As Boolean)
' If no params, return the User sheet;
' If Range_ is an object or a string, return the indicated range in the User sheet
Dim WkSht As Worksheet
Set WkSht = ThisWorkbook.Sheets("User")
DefVal NoObjectBack, True
On Error GoTo Bad

With WkSht

  If fbIM(Range_) Then
    On Error GoTo 0
    Set foUser = WkSht
  ElseIf TypeName(Range_) = "Range" Then
    Set foUser = .Range(Range_.Name)
  ElseIf NoObjectBack Then
    foUser = .Range(Range_)
  Else
    Set foUser = .Range(Range_)
  End If

End With
1 Exit Function

Bad: On Error GoTo 0
MsgBox "Error in foUser function: reference to range " & fsInQ(Range_), , pscSq
CrashEnd
End Function

Function foToolbars(Optional ByVal Range_, Optional ByVal NoObjectBack As Boolean)
' If no params, return the User sheet;
' If Range_ is an object or a string, return the indicated range in the User sheet
Dim WkSht As Worksheet
Set WkSht = ThisWorkbook.Sheets("Toolbars")
DefVal NoObjectBack, True

With WkSht

  If fbIM(Range_) Then
    Set foUser = WkSht
  ElseIf TypeName(Range_) = "Range" Then
    Set foUser = .Range(Range_.Name)
  ElseIf NoObjectBack Then
    foUser = .Range(Range_)
  Else
    Set foToolbars = .Range(Range_)
  End If

End With

End Function

Function fsApSub$(ReplaceThis$, InThis$, WithThis$, Optional Instance)
fsApSub = foAp.Substitute(InThis$, ReplaceThis$, WithThis$, Instance)
End Function

Sub ColWidth(ByVal Width, Col1, Optional Col2, Optional MinWidth)
If Col1 = 0 Then Exit Sub
CheckRC 250, , Col1, , Col2
Dim r As Range

If TypeName(Col1) = "Range" Then
  Set r = Col1
ElseIf Col1 = 0 Then
  Exit Sub
Else
  If fbIM(Col2) Then Col2 = 0
  If Col2 = 0 Then Col2 = Col1
  Set r = frSr(, Col1, , Col2)
End If

With r.Columns

  If Width = picAuto Then
    .AutoFit

    If fbNIM(MinWidth) Then
      .ColumnWidth = fvMax(.ColumnWidth, MinWidth)
    End If

  Else
    .ColumnWidth = Width
  End If

End With
End Sub

Sub HA(Alignment%, Optional rw1, Optional Col1, Optional rw2, Optional Col2)
frSr(rw1, Col1, rw2, Col2).HorizontalAlignment = Alignment
End Sub

Sub NoGridlines()
If Workbooks.Count > 0 Then ActiveWindow.DisplayGridlines = False
End Sub

Function fiColNum%(ByVal ColName, Optional Ix1, Optional Ix2, _
  Optional StrippedStr As Boolean = False, _
  Optional IgnoreErr As Boolean = False)
' Using the ColIndx sheet of this addin, determine the column whose header is ColName
Dim t$, u$, Cn%, Col%, p%, c$, q%, r&
c = LCase(Trim(ColName)): p = InStr(c, "("): Cn = 0

If LCase(Right$(c, 3)) = "col" And p = 0 Then
  c = Left$(c, Len(c) - 3)
ElseIf p > 3 Then

  If Mid$(c, p - 3, 3) = "col" And p > 0 Then
    c = Left$(c, p - 4) & Mid$(c, p): GoTo 1
  End If

End If

If fbNIM(Ix1) Then
  c = c & "(" & fsS(Ix1)
  If fbNIM(Ix2) Then c = c & "," & fsS(Ix2)
  c = c & ")"
End If

1: r = 0
With ThisWorkbook.Sheets("ColIndex")

  If StrippedStr Then ' Look in the column-header column

    Do
      r = r + 1
      t = fsStrip(.Cells(r, 3), , , , , , True, True, True, True, , True)
      If t = LCase(ColName) Then
        fiColNum = .Cells(r, 2): Exit Function
      End If
    Loop Until .Cells(r + 1, 3) = ""
     fiColNum = 0
    Exit Function

  Else ' Look in squid-assigned variable names

    Do
      r = r + 1
      t = .Cells(r, 1)
      p = InStr(t, "(")
      If LCase(Right$(t, 3)) = "col" And p = 0 Then
        t = Left$(t, Len(t) - 3)
      ElseIf p > 3 Then
        If Mid$(t, p - 3, 3) = "col" And p > 0 Then
          t = Left$(t, p - 4) & Mid$(t, p): GoTo 1
        End If
      End If

    Loop Until t = c Or r = 999

    If r = 999 Then
      If IgnoreErr Then
        r = 0: Exit Function
      Else
        GoTo NoCol
      End If

    End If

    fiColNum = .Cells(r, 2): Exit Function
  End If

End With
NoCol: On Error GoTo 0
End Function

Sub SqE(ByVal ErNum%)
Dim X%
MsgBox "Squid Error " & fsS(ErNum), , pscSq
On Error GoTo 0

If fbNotDbug Then
  CrashEnd
Else
  X = 1 / 0
End If

End Sub

Function fsCP$(ByVal ColName$, ByVal Ix1%, Optional Ix2) '
Dim s$
s = "(" & fsS(Ix1)
If fbNIM(Ix2) Then s = s & "," & fsS(Ix2)
s = s & ") "
fsCP = " " & ColName & s
End Function

Function fsAddr$(ByVal Rw&, ByVal Col%, Optional RowAbsolute = False, Optional ColAbsolute = True)
fsAddr = Cells(Rw, Col).Address(RowAbsolute, ColAbsolute)
End Function

Function fsFcell$(ByVal Equation$, ByVal Row&, Optional RowAbsolute As Boolean = False)
'Converts a source-string into a cell-formula, extracting the cell column #s
' from column-reference strings, which must be bracketed with spaces.  All other
'  symbols must NOT have spaces in between.
' Also, the first characters in Equation$ must not be a space-bracked column-ref.
'Example:  Eqns$ = "= Pb86 *( Pb86 - Pb46col *sComm84)/(1- Pb46col *sComm64)"
' becomes "= $J7*( $J7- $F7*sComm84)/(1- $F7
Dim s$, LeftPart$, e$, m As String * 1, SpacePos1%, SpacePos2%, ColNum%
e = Equation
s = "": m = " "
On Error GoTo 0

Do
  SpacePos1 = InStr(Equation, m) ' 1st space

  If SpacePos1 > 0 Then
    s = s & Left$(Equation, SpacePos1) ' carve out column name
    Equation = LTrim(Mid$(Equation, SpacePos1)) ' remaining eqn
    SpacePos2 = InStr(Equation, m) ' next space

    If SpacePos2 Then
      LeftPart = Left$(Equation, SpacePos2)
      ColNum = fiColNum(Trim(LeftPart))

      If ColNum = 0 Then
        MsgBox "Colname-index error for " & fsInQ(Trim(LeftPart))
        CrashEnd
      End If

      s = Trim(s) & Cells(Row, ColNum).Address(RowAbsolute)
      Equation = LTrim(Mid$(Equation, SpacePos2))
    End If

  End If

Loop Until SpacePos1 = 0

fsFcell = Trim(s) & Trim(Equation)
End Function

Public Function fbRangeNameExists(ByVal RangeName$, _
       Optional Wkbook, Optional WkSheet) As Boolean
' Is the specified name an existing range name?
Dim RangeAddress$, Nk$, Nh$
VIM Wkbook, ActiveWorkbook
VIM WkSheet, ActiveSheet
Nk = Wkbook.Name: Nh = WkSheet.Name
fbRangeNameExists = False
On Error GoTo 5
RangeAddress = Workbooks(Nk).Sheets(Nh).Range(RangeName).Address
fbRangeNameExists = True
5: On Error GoTo 0
End Function

Function fiFindHeader%(ByVal Header$, Optional Sheetname = "", Optional FromStd As Boolean)
' Looks for Header$ in either active sheet or specified sheet.

Dim StdCalc As Boolean, s$, ColNum%, LastCol%, HdrRow&, ShtIn As Worksheet

Set ShtIn = ActiveSheet

If fbNIM(FromStd) Then
  StdCalc = FromStd
Else
  StdCalc = pbStd
End If

Sheets(IIf(StdCalc, pscStdShtNa, pscSamShtNa)).Activate
HdrRow = flHeaderRow(StdCalc)
Header = fsStrip(s:=Header, IgCase:=False, IgVertSlash:=True)
LastCol = fiEndCol(HdrRow)

For ColNum = 1 To LastCol
  s = fsStrip(s:=Cells(HdrRow, ColNum).Text, IgCase:=False, IgVertSlash:=True)

  If LCase(s) = LCase(Header) Then
    fiFindHeader = ColNum
    Exit For
  End If

Next ColNum

If ColNum > LastCol Then fiFindHeader = 0
HdrRow& = plaLastDatRw(-StdCalc)
ShtIn.Activate
End Function

Sub LocateStdRows()
' Locate the PD file rows whose spot-names are those of specified standards.
Dim Age$, Conc$, N$, Nc$, Na$, tmp$, Nn%, i%, LA%, LC%, r&

ReDim plaConcStdRows(1 To 999), piaSpots(0 To 1, 1 To 999), piaConcStdSpots(1 To 999)
Age = LCase(psAgeStdNa): Conc = LCase(psConcStdNa)
LC = Len(Conc): LA = Len(Age)
piaNumSpots(1) = 0: piNumConcStdSpots = 0: piaNumSpots(0) = 0

For i = 1 To piNumAllSpots
  r = plaSpotNameRowsCond(i)
  tmp = LCase(psaSpotNames(i))
  Subst tmp, " "
  Na = Left$(tmp, LA)

  If Na <> Age Then
    piaNumSpots(0) = 1 + piaNumSpots(0)
    piaSpots(0, piaNumSpots(0)) = i
  End If

  If Na = Age Then
    piaNumSpots(1) = 1 + piaNumSpots(1)
    piaSpots(1, piaNumSpots(1)) = i
  End If

  If LC > 0 Then
    Nc = Left$(tmp, LC)

    If Nc = Conc Then
      piNumConcStdSpots = 1 + piNumConcStdSpots
      plaConcStdRows(piNumConcStdSpots) = r
      piaConcStdSpots(piNumConcStdSpots) = i
    End If

  End If

Next i

If piaNumSpots(0) = 0 Then
  pbStdsOnly = True
ElseIf piNumConcStdSpots > 0 Then
  ReDim Preserve plaConcStdRows(1 To piNumConcStdSpots)
  ReDim Preserve piaConcStdSpots(1 To piNumConcStdSpots)
End If

Nn = fvMax(piaNumSpots(1), piaNumSpots(0))
ReDim Preserve piaSpots(0 To 1, 1 To Nn)
End Sub

Function fsOptNumFor$(ByVal Row1&, ByVal Row2&, ByVal Col%, _
                      Optional Short As Boolean = False)
' return the "optimum" number format of a range
Dim i%, j%, d%, p%, Ndec%
Dim m#, s$, t#, X#, dMad#
Dim c As Range
Dim v() As Variant, vv As Variant

Const LogTen = 2.30258509299405
X = IIf(Short, 10, 1): j = 0
CheckRC 260, Row1, Col, Row2
ReDim v(1 To Row2 - Row1 + 1)

If Row1 = Row2 Then
  vv = Cells(Row1, Col)

  If IsNumeric(v) Then
    If CLng(vv) = vv Then fsOptNumFor = "General": Exit Function
  End If

End If

For i = Row1 To Row2
  Set c = Cells(i, Col)

  If IsNumeric(c) Then

    If Not IsEmpty(c) And c.Text <> "" Then
      j = j + 1
      v(j) = Val(c.Value)
    End If

  End If

Next i

If j > 0 Then
  ReDim Preserve v(1 To j)
  m = Abs(foAp.Median(v))

  Select Case m
    Case 0
      s = pscZq
    Case Is >= 1000000 / X
      s = "0.0E+0"
    Case Is >= 100000 / X
      s = "0E+0"
    Case Is >= 100 / X
      s = pscZq
    Case Is < 0.001 * X
      s = "0.0E+0"
    Case Else

      If j = 1 Then
        p = 3 - Int(Log(m) / Log(10))
        s = "0." & String(fvMax(1, p + (X <> 1)), pscZq)
      Else
        dMad = Abs(MAD(v))

        If dMad > 0 Then
          i = 1 - fiIntLogAbs(dMad)
          Ndec = fvMax(0, i) + Short
          s = 0
          If Ndec > 0 Then s = s & "." & String(Ndec, "0")
        Else
          s = pscGen
        End If

      End If

  End Select

Else
  s = pscGen
End If

fsOptNumFor = s
End Function

Sub Nformat(Optional Col% = 0, Optional ErCol = False, Optional Std = 0, _
  Optional Frmat$ = "", Optional Short As Boolean = False, _
  Optional RestrictedRange)
' Number-format a range, either auto or specified
Dim f$, i%, N%, t%, sL%, vct%, m#, v() As Variant, X As Variant, r As Range, SumV
ReDim v(1 To 9999)

If Col > 0 Then
  Set r = frSr(plaFirstDatRw(Abs(Std)), Col, plaLastDatRw(Abs(Std)))
ElseIf IsMissing(RestrictedRange) Then
  Exit Sub
Else
  Set r = RestrictedRange
End If

With r
  N = .Rows.Count

  For i = 1 To N
    X = .Cells(i, 1)

    If IsNumeric(X) Then
      vct = 1 + vct
      v(i) = X
      sL = sL + Len(.Cells(i, 1).Text)
    End If

  Next i

  If vct = 0 Or sL = 0 Then Exit Sub
  ReDim Preserve v(1 To vct)
  SumV = foAp.Sum(v)

  If Int(SumV) = SumV And Int(v(1)) = v(1) And Int(v(vct)) = v(vct) Then
    .NumberFormat = "General"
    If fbNIM(Frmat) Then Frmat = .NumberFormat
    Exit Sub
  End If

  If vct > 1 Then
    m = Abs(foAp.Median(v))
  Else
    m = Abs(v(1))
  End If

  If ErCol Then

    If m < 0.000001 Then
      f = "0": .HorizontalAlignment = xlCenter
    ElseIf m < 0.1 Then
      f = pscZd3
    ElseIf m < 0.01 Then
      f = "0.0000"
    Else
      f = pscErF 'pscZd2 ' 09/06/09 -- mod
    End If

  Else

    If m = 0 Then
      If UBound(v) > 1 Then m = foAp.Average(v) Else m = v(1)
    End If

    Select Case m
      Case Is < 1E-18: f = "0": .HorizontalAlignment = xlCenter
      Case Is < 0.00001, Is > 100000
        f = "0.0E+0"
      Case Is > 1000: f = pscZq
      Case Is > 100: f = "0.0"
      Case Else
        t = fiIntLogAbs(m)
        f = "0." & String(fvMax(3 + 2 * Short - t, 1), pscZq)
    End Select

  End If

  .NumberFormat = fsQq(f)
  If fbNIM(Frmat) Then Frmat = .NumberFormat
End With
End Sub

Sub NumFormatColumns(ByVal Std As Boolean) ' Number-format the U-Pb data columns
' 09/04/01 -- Change statbar parse to "Excel is recalculating the spreadsheet..."
Dim LastHdrWasAge As Boolean, b As Boolean, IsInt As Boolean, HdrIsAge As Boolean
Dim LcHdr$, LastHdr$, Hdr$, sp1$, sp2$, sp2a$, sp3$, sq$, t1$, t2$, q$, h$
Dim j%, k%, i%, s%, nV%, NvV%, c%
Dim Frw&, Lrw&, Rw&
Dim v#
Dim cc As Range, Co As Range
Dim va As Variant, Vra() As Variant, vv As Variant
Dim ccf As Font

StatBar "Optimizing significant figures in columns"
s = -Std
Frw = plaFirstDatRw(s)
Lrw = plaLastDatRw(s)
sp1 = "[>1E-6]0.0E+0;[<-1E-6]-0.0E+0;--- "
sp2 = "[>=.01]0.0"
sp2a = "[>=.01]0;[<.1]--- "
sp3 = fsQq("+0;$-$0;0")
sq = fsQq("$+$0.00$%$;$-$0.00$%$;$0$")

HA xlCenter, plHdrRw, piDateTimeCol
RangeNumFor "dd mmm, yyyy  hh:mm", , piDateTimeCol
ColWidth 20, piHoursCol
ColWidth picAuto, piHoursCol
RangeNumFor pscZd2, , piHoursCol

If pbXMLfile Then
  If piStageXcol > 0 Then HA xlCenter, plHdrRw, piStageXcol
  If piStageYcol > 0 Then HA xlCenter, plHdrRw, piStageYcol
  If piStageZcol > 0 Then HA xlCenter, plHdrRw, piStageZcol
  If piQt1yCol > 0 Then HA xlCenter, plHdrRw, piQt1yCol
  If piQt1Zcol > 0 Then HA xlCenter, plHdrRw, piQt1Zcol
End If

RangeNumFor pscZd2, Frw, piBkrdCtsCol, Lrw
RangeNumFor pscZd2, Frw, piPb204ctsCol, Lrw
RangeNumFor pscZq, Frw, piPb206ctsCol, Lrw
RangeNumFor sp1, Frw, piPb46col, Lrw

If piPb46col > 0 Then

  For Rw = Frw To Lrw
    va = Cells(Rw, piPb46eCol)

    If IsNumeric(va) Then
      v = fvMin(va, 9999)
      Cells(Rw, piPb46eCol) = fvMax(0.000000001, v)
    End If

   Next Rw

  va = foAp.Median(frSr(Frw, piPb46eCol, Lrw))

  If IsNumeric(va) Then

    For Rw = Frw To Lrw

      If Cells(Rw, piPb46eCol) = 9999 Then
        t1 = " ---"
      Else
        t1 = IIf(va < 2, sp2, sp2a)
      End If

      RangeNumFor t1, Rw, piPb46eCol
    Next Rw

  End If

  On Error Resume Next
  For i = 1 To 5
    If Std Then
      j = Choose(i, piStdCom6_4col, piStdCom6_7col, piStdCom6_8col, _
                 piStdCom8_4col, piStdCom8_7col)
    Else
      j = Choose(i, piCom6_4col, piCom6_7col, piCom6_8col, _
                 piCom8_4col, piCom8_7col)
    End If
    If j > 0 Then
      va = foAp.Median(frSr(Frw, j, Lrw))
      If IsNumeric(va) Then
        h = IIf(va < 0.1, pscZd3, pscZd2)
        RangeNumFor h, , j
      End If
    End If
  Next i
  On Error GoTo 0
End If

If Std And piPb46col > 0 Then

  For i = 7 To 8
    Nformat piaOverCts4Col(i), , True, , True
    Nformat piaOverCts46Col(i), , True, , True
    Nformat piacorrAdeltCol(i), , True, , True ' xcept for this col
  Next i

Else
  Nformat piaPpmUcol(s), , Std
  'Nformat piaPpmThcol(s),  , Std
End If

StatBar "Excel is recalculating the spreadsheet..."
foAp.Calculate
StatBar "formatting"
nV = Lrw - Frw + 1
If nV = 0 Then Exit Sub
LastHdr = "": LastHdrWasAge = False
plHdrRw = flHeaderRow(pbStd)
piLastCol = fiEndCol(plHdrRw)

For c = fvMin(4, 1 + piPb206ctsCol) To piLastCol
  Set Co = frSr(Frw, c, Lrw)
  Set cc = Cells(plHdrRw, c)
  Set ccf = cc.Font
  Hdr = cc.Formula
  Subst Hdr, vbLf
  LcHdr = LCase(Hdr)
  vv = Cells(Frw, c)
  IsInt = False

  If IsNumeric(vv) Then
    If Abs(vv) < 2147483647 Then ' 09/06/09 -- added.
      If CLng(vv) = vv Then IsInt = True
    End If
  End If

  b = (Hdr <> "" And c <> piPb46col And c <> piPb46eCol And _
       c <> piBkrdCtsCol And c <> piPb204ctsCol)

  If b Then
    If Std Then
      b = (c <> piStdCom6_4col And c <> piStdCom8_4col And c <> piStdCom6_7col And _
           c <> piStdCom6_8col And c <> piStdCom8_7col)
    Else
      b = (c <> piCom6_4col And c <> piCom8_4col And c <> piCom6_7col And _
           c <> piCom6_8col And c <> piCom8_7col)
    End If
  End If

  If Std And b Then

    For j = 7 To 8
      b = b And Not (c = piaOverCts4Col(j) Or c = piaOverCts46Col(j) Or _
          c = piacorrAdeltCol(j))
    Next j

  End If

  If b Then
    LastHdrWasAge = HdrIsAge
    HdrIsAge = (InStr(Hdr, "Age(Ma)") > 0 Or Right$(LcHdr, 3) = "age")

    If Hdr = "%Dis-cor-dant" Then
      q = sp3
    ElseIf Left$(Hdr, 4) = "%err" Then
      NvV = 0
      ReDim Vra(1 To nV)

      For j = 1 To nV
        va = Co(j)

        If IsNumeric(va) Then

          If va <> 0 Then
            NvV = 1 + NvV
            Vra(NvV) = va
          End If

        End If

      Next j

      va = foAp.Median(Vra)

      If IsNumeric(va) Then

        Select Case va
          Case 0: q = "0"
          Case Is < 0.01: q = "0.0000"
          Case Is < 0.1: q = pscZd3
          Case Is < 1: q = pscErF 'pscZd2 '"0.00"
          Case Is < 10: q = pscZd1 '"0.0"
          Case Is < 100000: q = "0"
          Case Else: q = "0E+0"
        End Select

      Else
        q = "General"
      End If

    ElseIf HdrIsAge Then
      q = pscAgeFormat
    ElseIf (Hdr = pscPm & "1s" Or Hdr = "1serr") And LastHdrWasAge Then
      q = pscAgeErrFormat
    ElseIf ccf.italic And ccf.Color = vbBlue And IsInt Then
      q = "General"
    Else
      b = (InStr(Hdr, "err") > 0 Or Hdr = pscPm)
      q = fsOptNumFor(Frw, Lrw, c, b)
    End If

    RangeNumFor q, Frw, c, Lrw

    If q = sp3 Then
      frSr(Frw, c, Lrw).InsertIndent 1
    End If

    If Std Then

      For j = 1 To piNumDauPar
        If c = piaSacol(j) Then psaCalibConstNumFor(j) = q
      Next j

    End If

  End If

  If Std And cc.Font.italic Then ' 09/06/09 -- deleted "And Bold"
    frSr(1 + plaFirstDatRw(1), c, plaLastDatRw(1)).Interior.Color = RGB(216, 216, 216)
  End If

  LastHdr = Hdr
Next c

If pbUPb Then
  Cells.FormatConditions.Delete
  frSr(Frw, 4, Lrw, fiEndCol(plHdrRw)).Select
  With Selection
    ' whites out 232/238, 4overcts8, 46from8, 8constdelta because all #REF at this point
    .FormatConditions.Add Type:=xlExpression, Formula1:="=ISERROR(" & Cells(Frw, 4).Address(0, 0) & ")"
    .FormatConditions(1).Font.Color = vbWhite
  End With
End If

For Rw = Frw To Lrw

  For c = 4 To fiEndCol(Rw)
    If Not Columns(c).Hidden Then  ' 09/07/21 -- added
    With Cells(Rw, c)

      If Len(.Text) > 8 Then
        .NumberFormat = "0.0E+0"
      End If

    End With
    End If
Next c, Rw

For c = 2 To fiEndCol(flHeaderRow(Std))  ' 09/07/21 -- added
  If Not Columns(c).Hidden Then ColWidth picAuto, c
Next c
'ColWidth picAuto, 2, fiEndCol(flHeaderRow(Std))  ' 09/07/21 -- commented out
End Sub

Sub DupeNames(NewShtWbkName$, Sheet1Workbook2%)
' If the specified sheet already exists or a Workbook whose name is
' NewShtWbkName$ is currently loaded, add _1, _2 ... to the name.
' If a Sheet, then rename the sheet (but leave "NewShtWbkName" unchanged).
' If a Workbook, return "NewShtWbkName" as the new name (but don't save it).

Dim Found As Boolean
Dim s$, f$, ShtsWbksNa$, ShtWbkInNa$, NewName$
Dim i%, ShtsWbksN%, LenNa%, sw%, p%(), q%()
Dim Obj As Object, ShtWbk As Object, ShtWbkIn As Object

sw = Sheet1Workbook2

Select Case sw
  Case 1: Set ShtWbk = Sheets
  Case 2: Set ShtWbk = Workbooks
  Case Else
    MsgBox "SQUID code error: " & fsInQ("Sheet1Workbook2") & _
           " passed as" & StR(sw) & "in sub DupNames."
    End
End Select

ShtsWbksN = ShtWbk.Count
ShtsWbksNa = LCase(NewShtWbkName)
LenNa = Len(ShtsWbksNa)
ReDim q(1 To ShtsWbksN), p(1 To ShtsWbksN)
Set ShtWbkIn = Choose(sw, ActiveSheet, ActiveWorkbook)
ShtWbkInNa = ShtWbkIn.Name

For i = 1 To ShtsWbksN
  s = LCase(ShtWbk(i).Name)

  If Left$(s, LenNa) = ShtsWbksNa Then
    Found = True

    If Len(s) = LenNa Then
      p(i) = 2
    Else
      f = Mid$(s, 2 + LenNa)

      If Mid$(s, LenNa + 1, 1) = " " And fbIsNum(f) Then
        p(i) = 1 + f
      End If

    End If

  End If

Next i

If Not Found Then Exit Sub

BubbleSort p(), q(), True
ShowStatusBar
ManCalc

For i = 1 To ShtsWbksN

  If p(i) Then
    NewName = NewShtWbkName & StR(p(i))

    If sw = 1 Then
      ShtWbk(q(i)).Name = NewName
    ElseIf sw = 2 Then
      NewShtWbkName = NewName
    End If

    Exit For
  End If

Next i

End Sub

Sub FormulaEval(ByVal Formula, ByVal EqNum%, ByVal ScanNum%, PkHts#(), _
  PkHtFerr#(), NumericRes#, NumericFerr#)
' PkHts() and PkHtFerr() must be indexed in order of scanning
' Evaluate numerically an algebraic expression
Const DeltMult = 1.0001, Numerator = 1, Denominator = 2

Dim MTcell As Boolean, GotUnduped As Boolean, b As Boolean, ErColRef() As Boolean
Dim PertF$, LastPertEq$, f0$, PertEq$, s$, tmp$, Form$, tForm$, IsEq As Boolean
Dim Numer$(), Denom$(), ErColRefAddr$(), EqRef$, RefEqNum%, CellAddr$
Dim UnDupPkOrd%, q%, PkOrd%, RatNum%, NdpPk%, MaxConsts%
Dim fDelt#, fVar#, PertVal#, RatVal#(50), Num#(), Den#(), PratVal#()
Dim j%, k%, VarLen%, d1%, d2%, Col1%, Col2%, nt%, Rct%, Nr%, Nrt%, Nrn%
Dim tA As Variant, tB As Variant, tc As Variant, TD As Variant, NumRes As Variant
Dim TestCell As Range, RatCt%

Static BypassEmptyCell(-4 To 99) As Boolean

f0 = Formula

 ' 10/04/02 - added to correct any range names mistakenly enclosed in ["  "]
StripBracketedRangenames f0, Nrn

On Error GoTo 0

Formula = f0
Nr = piaEqnRats(EqNum)
nt = piaNeqnTerms(EqNum)
Nrt = Nr + nt

ReDim Num(1 To Nrt), Den(1 To Nrt), Numer(1 To Nrt), Denom(1 To Nrt), Rat(1 To Nrt)
ReDim ErColRefAddr$(1 To Nrt), ErColRef(1 To Nrt), PratVal(1 To Nrt)
NumericRes = 0: NumericFerr = 0: RatCt = 0

For RatNum = 1 To Nrt
  Col1 = InStr(Formula, psBrQL): Col2 = InStr(Formula, psBrQR)
  d1 = InStr(Formula, "["):      d2 = InStr(Formula, "]")

  If piaBrakType(RatNum, EqNum) = 1 Then  ' isotope ratio
    RatCt = 1 + RatCt
    j = piaEqPkOrd(EqNum, RatNum, Numerator)   'RatCt, Numerator)
    k = piaEqPkOrd(EqNum, RatNum, Denominator) 'RatCt, Deniminator)
    SqBrakExtract Formula, s
    ErColRef(RatNum) = (Left$(s, 1) = pscPm) 'RatCt) = (Left$(s, 1) = pscPm)
    IsEq = False

    If ErColRef(RatNum) Then  'RatCt) Then
      With puTask
        tmp = StR(.daNmDmIso(1, 1)) & "/" & StR(.daNmDmIso(2, 1))
        FindStr Trim(tmp), , j, plHdrRw
      End With
      ' FindStr Trim(puTask.saIsoRats(RatCt)), , j, plHdrRw
      tA = Cells(plOutputRw, j).Address(0, 0)
      ErColRefAddr(RatCt) = tA
      s = tA
    Else
      Num(RatCt) = PkHts(j, ScanNum):  Den(RatCt) = PkHts(k, ScanNum)
      Numer(RatCt) = StR(Num(RatCt)):  Denom(RatCt) = StR(Den(RatCt))
      If Den(RatCt) = 0 Then
        NumericRes = 1E+32: NumericFerr = 0: Exit Sub
      End If
      RatVal(RatCt) = Num(RatCt) / Den(RatCt)
      Rat(RatCt) = StR(RatVal(RatCt))
      s = Rat(RatCt)
    End If

  Else ' eqn index
    IsEq = True
    s = Mid$(Formula, d1 + 1, d2 - d1 - 1)
    s = fsLegalName(s, True, True, False)

    If Not BypassEmptyCell(EqNum) Then
      MTcell = False

      If fbRangeNameExists(s) Then
        Subst s, Chr(34)
        Set TestCell = Range(s)
      Else ' a column-header reference?

        If Col1 = d1 And Col2 = d2 - 1 Then
          Subst s, Chr(34)
        End If

        ' find the column# of the column-header
        FindStr Phrase:=fsLegalName(s), ColFound:=k, RowLook1:=plHdrRw, _
          ColLook1:=2, WholeWord:=True

        If k = 0 Then
          MTcell = True
        ElseIf plSpotOutputRw > 0 Then            ' replace the column-cell reference by its value
          Set TestCell = Cells(plSpotOutputRw, k) '  in f0
          On Error Resume Next
          s = TestCell.Value
          On Error GoTo 0
          Subst f0, Mid$(Formula, d1, d2 - d1 + 1), s
        End If

      End If

      If Not MTcell Then
        If IsEmpty(TestCell) Then MTcell = True
      End If

      If MTcell And EqNum > 0 Then

        If puTask.uaSwitches(EqNum).Nu Then
          s = "User equation " & fsS(EqNum) & " cannot be numerically evaluated" _
            & " because no values for range or column " & s & " have yet been " & _
            "calculated." & pscLF2 & "To avoid this problem, set the FO switch " & _
            "when defining the equation." & pscLF2 & "Abandon data reduction now?"
          If MsgBox(s, vbYesNo, pscSq) <> vbNo Then CrashEnd
          BypassEmptyCell(EqNum) = True
        End If

      End If

    End If

  End If

  Formula = Left$(Formula, d1 - 1) & s & Mid$(Formula, d2 + 1)

Next RatNum

NumRes = Evaluate(Formula)
If IsError(NumRes) Or Not fbIsNum(NumRes) Then
 NumericRes = pdcErrVal: Exit Sub
End If
NumericRes = NumRes

If NumericRes = 0 Then Exit Sub

fVar = 0: PertF = Formula: LastPertEq = ""

For NdpPk = 1 To piNoDupePkN(EqNum)
  GotUnduped = False
  PertEq = f0: UnDupPkOrd = piaEqPkUndupeOrd(EqNum, NdpPk)

  For RatNum = 1 To Nr
    If Not ErColRef(RatNum) Then PratVal(RatNum) = RatVal(RatNum)
  Next RatNum

  For RatNum = 1 To Nr
    If ErColRef(RatNum) Then

    Else
      For k = Numerator To Denominator
        PkOrd = piaEqPkOrd(EqNum, RatNum, k)

        If PkOrd = UnDupPkOrd Then
          GotUnduped = True

          If k = 1 Then ' Numerator
            PratVal(RatNum) = Num(RatNum) * DeltMult / Den(RatNum)
          Else
            PratVal(RatNum) = Num(RatNum) / (Den(RatNum) * DeltMult)
          End If

          Exit For ' numerator <> denominator
        End If

      Next k

    End If

  Next RatNum

  For RatNum = 1 To Nr
    d1 = InStr(PertEq, "["): d2 = InStr(PertEq, "]")

    If ErColRef(RatNum) Then
      s = Range(ErColRefAddr(RatNum))
    Else
      s = StR(PratVal(RatNum))
    End If

    PertEq = Left$(PertEq, d1 - 1) & s & Mid$(PertEq, d2 + 1)
  Next RatNum

  If GotUnduped Then
    Subst PertEq, psBrQL, , psBrQR  ' in case named user-equation present
    PertVal = Evaluate(PertEq)  ' Numeric result of perturbing one peak
    If NumericRes = 0 Then NumericRes = pdcTiny
    fDelt = PertVal / NumericRes - 1 ' fractional difference
    ' kluge to evade an Excel2001 bug.
    tA = PkHtFerr(PkOrd, ScanNum)
    tB = DeltMult - 1
    tc = fDelt * fDelt
    TD = (tA / tB) ^ 2 * tc
    ' fractional internal variance
    fVar = fVar + TD
    LastPertEq = PertEq
  End If

Next NdpPk

NumericFerr = sqR(fVar)
End Sub

Function fiInstanceLoc%(ByVal InString$, ByVal LocNum%, ByVal ThisChar$)
' Return the position of the LocNum-th instance of ThisChar
Dim t$, i%, j%

For i = 1 To Len(InString)
  t = Mid$(InString, i, Len(ThisChar))

  If t = ThisChar Then
    j = 1 + j
    If j = LocNum Then fiInstanceLoc = i: Exit Function
  End If

Next i

fiInstanceLoc = 0
End Function

Sub AllInstanceLoc(ByVal fsSubStr$, ByVal Phrase$, InstLoc%(), LocCt%)
' phrase=ABC{[DEFGaa{[zz{[  p=4  LocCt=1  fiInstanceLoc(1)=4  Le=5
' phrase=DEFGaa{[zz{[       p=7  LocCt=2  fiInstanceLoc(2)=12 Le=13
' phrase=zz[{               p=3  LocCt=3  fiInstanceLoc(3)=16
' phrase=""
Dim s$, Le%, Ls%, p%
Ls = Len(fsSubStr): Le = 0: LocCt = 0

Do
  p = InStr(Phrase, fsSubStr)

  If p > 0 Then
    LocCt = 1 + LocCt
    ReDim Preserve InstLoc(1 To LocCt)
    InstLoc(LocCt) = p + Le
    Le = Le + p - 1 + Ls
    Phrase = Mid$(Phrase, p + Ls)
  End If

Loop Until p = 0

End Sub

Function fiInstanceCount%(ByVal Char$, ByVal Phrase$) ' Return the number of occurences of Char in Phrase
Dim i%, ct%

For i = 1 To Len(Phrase)
  If Mid$(Phrase, i, 1) = Char Then ct = 1 + ct
Next

fiInstanceCount = ct
End Function

Sub PlaceFormulae(ByVal Formula$, ByVal FirstRow&, ByVal Col%, Optional ByVal LastRow&)
Dim r& ' Put Formula in the specified range
If LastRow = 0 Or fbIM(LastRow) Then LastRow = FirstRow
If fbNotDbug Then On Error Resume Next
frSr(FirstRow, Col, LastRow) = fsFcell(Formula, FirstRow)
End Sub

Function fsRejFormat$() ' format to indicate a rejected point
Dim s$
s = "[Red]" & """" & "REJ" & """"
fsRejFormat = s & ";" & s & ";" & s
End Function

Sub sCopyPaste(ByVal ShapeName$, Optional SourceBook, Optional SourceSheet, _
  Optional DestRow, Optional DestCol, Optional DestBookName, Optional DestShtName)
' To avoid problems with Mac and Virtual PC
Dim Sbk As Workbook, Dbk As Workbook, Ssht As Worksheet, Dsht As Worksheet

If fbIM(SourceBook) Then Set Sbk = ActiveWorkbook Else Set Sbk = SourceBook
If fbIM(DestBookName) Then Set Dbk = ActiveWorkbook Else Set Dbk = DestBookName
If fbIM(SourceSheet) Then Set Ssht = ActiveSheet Else _
                          Set Ssht = Sbk.Sheets(SourceSheet)
If fbIM(DestShtName) Then Set Dsht = ActiveSheet Else _
                          Set Dsht = Dbk.Sheets(DestShtName)
Ssht.Shapes(ShapeName).Copy

If True Then
  Ssht.Paste
  foLastOb(Ssht.Shapes).Cut
End If

If fbNIM(DestRow) And fbNIM(DestCol) Then
  Dsht.Activate
  Cells(DestRow, DestCol).Select
End If

Dsht.Paste
End Sub

Function fbLegalEq(ByVal Equation$, ByVal CellParam As Boolean, _
  Isorats$(), Optional BadSwap As Boolean = False) As Boolean
' Is "Equation" in legal Squid & Excel format?

' 09/04/10 -- Add code to detect SC-equation names present as their range-names to be.

Dim BadSht As Boolean, BadWbk As Boolean, AlphEq As Boolean
Dim BadW As Boolean, BadS As Boolean, NumEq As Boolean, cP As Boolean
Dim ConstN$, tmp$, HardwiredColHdr$, TestColHdr$
Dim SwapCol$, e$, d$, s$, s1$, s0$, t$, Eq$
Dim FuncArr$(), Wna$(), Sna$()
Dim pb%, pv%, qb%, eb%, qv%, IndxType%, i%, p%, q%
Dim Indx%, Lett%, ErNumber%, No%, Nm%, Nrefs%
Dim ConstV#
Dim SquidRangeNames As Range, MultiInputFunctions As Range, ColumnHeaders As Range
Dim Result As ErrObject
Dim res As Variant, eRes As Variant, Tres As Variant
Dim SyntaxErr1 As Variant, SyntaxErr2 As Variant, ErrDivZero As Variant

With ThisWorkbook.Sheets("ColHdrs")
  Set ColumnHeaders = .[ColumnHeaders]
  Set MultiInputFunctions = .[MultiInputFunctions]
  Set SquidRangeNames = .[SquidRangeNames]
End With

Tres = True: d$ = Chr(34)
ErrDivZero = CVErr(2007): SyntaxErr1 = CVErr(2015): SyntaxErr2 = CVErr(2029)
Eq = Trim(Equation)
s = Eq: AlphEq = False: NumEq = False: fbLegalEq = False
FindWbkShtRefs LCase(Eq), Nrefs, Wna, Wna, BadWbk, BadSht ', s
If BadWbk Or BadSht Then fbLegalEq = False: Exit Function

s = LCase(Eq)
p = InStr(s, "<=>")

If p > 0 Then
  tmp = Mid$(Eq, 3 + p)
   ExtractEqnRef tmp, SwapCol, Indx, IndxType, Isorats

  If IndxType = peColumnHeader Then
    TestColHdr = LCase(fsLegalName(SwapCol))
    With ColumnHeaders

      For i = 1 To .Rows.Count
        HardwiredColHdr = LCase(fsLegalName(.Cells(i, 1)))

        If HardwiredColHdr = TestColHdr Then
          MsgBox "Sorry, you can't swap the " & fsInQ(SwapCol) & " column.", , pscSq
          BadSwap = True: fbLegalEq = False
          Exit Function
        End If

      Next i

    End With
  End If

  s = LCase(Left$(s, p - 1))
End If

Do
 s0 = s
 Subst s, "+]", "]"
 s = fsBracketReplace(s, " " & fsRvar(CellParam) & " ", True)
Loop Until s0 = s

For i = 1 To peMaxConsts
  ConstN = LCase(prConstNames(i))
  If ConstN = "" Then Exit For
  ConstV = prConstValues(i)
  Subst s, ConstN, Abs(Rnd)
Next i

With SquidRangeNames

  For i = 1 To .Rows.Count
    e = .Cells(i, 1)
    Subst s, e, "1"
  Next i

End With

With puTask ' look for Range names from SC eqns
  For i = 1 To .iNeqns

    If .uaSwitches(i).SC Then
      t = LCase(.saEqnNames(i))
      p = InStr(s, t)

      If p > 0 Then
        If p > 2 Then s1 = Mid(s, p - 2, 2)
        If p > 2 And s1 <> "[" & Chr(34) Then
          Subst s, t, Abs(Rnd)
        End If
      End If

    End If

  Next i
End With


If InStr(s, "choose") = 0 And InStr(s, "biweight") = 0 Then

  For i = 1 To MultiInputFunctions.Count
    e = LCase(MultiInputFunctions(i))
    Subst s, e, "Max(1,", "Max(1,(", "Max(1,"
  Next i

  On Error Resume Next
  res = Evaluate(s)

  If fbNoNum(res) And IsError(res) Then
    eRes = CVErr(res)
    If eRes = SyntaxErr1 Or eRes = SyntaxErr2 Then Tres = False
  End If

End If

fbLegalEq = Tres
End Function

Function fsRvar$(ByVal cP As Boolean)
If cP Then fsRvar = fsRandArray Else fsRvar = fsS(Rnd)
End Function

Function fsRandArray() As String
fsRandArray = fsS(Rnd) & "," & fsS(Rnd) & "," & fsS(Rnd)
End Function

Sub FindDelCol(ByVal Header$, ByVal ColsToRight%, ByVal RowLook1, _
  Optional ColLook1 = 1, Optional RowLook2, Optional ColLook2 = 255, _
  Optional CaseSensitive As Boolean = False, _
  Optional InclLineFeeds As Boolean = False)
' find col-hdr Header & delete
Dim ColFound%, RowFound&

FindStr Header, , ColFound, RowLook1, ColLook1, RowLook2, ColLook2, CaseSensitive, _
   InclLineFeeds
If ColFound Then DelCol ColFound, ColFound + ColsToRight
End Sub

Function fbIsInSubset(ByVal SpotName$, ByVal EqNum%)
' Is spotname correspond to a specified name-fragment for equation EqNum?

Dim Nfrag$, i%, Le

Nfrag = fsStrip(prSubsSpotNameFr(EqNum))
SpotName = fsStrip(s:=SpotName, IgNonAlphaNum:=True)
Le = Len(Nfrag)

If Le = 0 Or Left$(SpotName, Le) = Nfrag Then
  fbIsInSubset = True
Else
  fbIsInSubset = False
End If

End Function

Function fnLeftTopRowCol(ByVal Row1Col2%, ByVal TopLeft!) As Single
' Returns row or column corresponding to topleft

Dim c%, r&, w!, h!, L!, t!, Result!, fract!

With ActiveSheet

  If Row1Col2 = 1 Then

    For r = 1 To pemaxrow
      t = .Rows(r).Top

      If TopLeft < t Then
        h = .Rows(r - 1).Height
        fract = (TopLeft - .Rows(r - 1).Top) / h
        Result = r - 1 + fract
        Exit For
      End If

    Next r

  Else

    For c = 1 To peMaxCol
      L = .Columns(c).Left

      If TopLeft < L Then
        w = .Columns(c - 1).Width
        fract = (TopLeft - .Columns(c - 1).Left) / w
        Result = c - 1 + fract
        Exit For
      End If

    Next c

  End If

End With
fnLeftTopRowCol = Result
End Function

Sub OkColWidth(ByVal Col%, ByVal OkWidth!, ByVal StdCalc As Boolean)
Cells(plaFirstDatRw(-StdCalc), Col).Cut
Cells(1, 256).Select
On Error GoTo 1
ActiveSheet.Paste
On Error GoTo 0

With Columns(Col)
  .ColumnWidth = 16
  .AutoFit
  .ColumnWidth = fvMax(.ColumnWidth, OkWidth)
End With

Cells(1, 256).Cut
Cells(plaFirstDatRw(-StdCalc), Col).Select
ActiveSheet.Paste
1: On Error GoTo 0
End Sub

Function frRowColExtract(ByVal Row1Col2%, SourceRange As Range, _
                         ByVal RowColNum&) As Range

With SourceRange

  Select Case Row1Col2
    Case 1
      Set frRowColExtract = Range(.Cells(RowColNum, 1), _
          .Cells(RowColNum, .Columns.Count))
    Case 2
      Set frRowColExtract = Range(.Cells(1, RowColNum), _
           .Cells(.Rows.Count, RowColNum))
  End Select

End With
End Function

Function fbSheetExist(ShtName$) As Boolean
Dim Sht As Worksheet

For Each Sht In ActiveWorkbook.Worksheets

  If LCase(ShtName) = LCase(Sht.Name) Then
    fbSheetExist = True
    Exit Function
  End If

Next Sht

fbSheetExist = False
End Function

Sub CollateUserConstants()

Dim Extr$, Eq$, s$, IndxType%
Dim EqNum%, Rw#, ParamIndx%, ConstNum, Col%, ColEnd%, HdrRow%
Dim IndxUsed%(peMaxConsts)
Dim ConstVals#
Dim Consts As Range

If pbUPb Then phStdSht.Activate Else phSamSht.Activate
ReDim psaUsrConstNa(1 To peMaxConsts), pdaUsrConstVal(1 To peMaxConsts)
piNconstsUsed = 0

With puTask

  For EqNum = piLwrIndx To .iNeqns

    If .saEqns(EqNum) <> "" Then

      Do
        ExtractEqnRef .saEqns(EqNum), Extr, ParamIndx, IndxType

        If IndxType = pePrefsConstant Then
          ParamIndx = -ParamIndx - 1000
        ElseIf IndxType = peBothConstant Then
          ParamIndx = -ParamIndx - 3000
        Else
          Exit Do
        End If

        For ConstNum = 1 To piNconstsUsed
          If ParamIndx = IndxUsed(ConstNum) Then Exit For
        Next ConstNum

        If ConstNum > piNconstsUsed Then
          piNconstsUsed = 1 + piNconstsUsed
          psaUsrConstNa(piNconstsUsed) = fs_(prConstNames(ParamIndx))
          ConstVals = prConstValues(ParamIndx)
          pdaUsrConstVal(piNconstsUsed) = ConstVals
          IndxUsed(piNconstsUsed) = ParamIndx
        End If

        Subst .saEqns(EqNum), "<" & Extr & ">", "_" & Extr
      Loop

    End If

  Next EqNum

End With

If piNconstsUsed > 0 Then
  ReDim Preserve psaUsrConstNa(1 To piNconstsUsed)
  ReDim Preserve pdaUsrConstVal(1 To piNconstsUsed)
  plHdrRw = flHeaderRow(pbUPb)
  ColEnd = 2 + fiEndCol(HdrRow)

  For Rw = plHdrRw To HdrRow + 5 + piNconstsUsed
    Col = fiEndCol(Rw) + 2
    If Col > ColEnd Then ColEnd = Col
  Next Rw

  Set Consts = frSr(1 + HdrRow, ColEnd, piNconstsUsed + HdrRow, 1 + ColEnd)

  With Consts

    For ConstNum = 1 To piNconstsUsed
      With .Item(ConstNum, 1)
        .Formula = "_" & psaUsrConstNa(ConstNum) & " ="
        .HorizontalAlignment = xlRight
      End With

      With .Item(ConstNum, 2)
        .Formula = pdaUsrConstVal(ConstNum)
        .HorizontalAlignment = xlLeft
        .Name = "_" & psaUsrConstNa(ConstNum)
      End With

    Next ConstNum

    .NumberFormat = "General"
    .Name = "ConstantsUsed"
  End With

Else
  Erase psaUsrConstNa, pdaUsrConstVal
End If

End Sub

Sub CreateNewWorkbook(Optional NumberOfSheets% = 1)
foAp.SheetsInNewWorkbook = 1
Workbooks.Add
NoGridlines
foAp.SheetsInNewWorkbook = NumberOfSheets
End Sub

Sub BlankZeroCells(CellsRange As Range)
Const Clr = 12632256
On Error GoTo Bad

With CellsRange
  .FormatConditions.Delete
  .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
  With .FormatConditions(1)
    .Font.Color = Clr
    .Interior.Color = Clr
  End With
End With

Bad:
End Sub

Sub MakeDatRedParamsSht()
Dim i%, j%
Dim Rw&, rwStdConc&, rwUPb&, rwThPb&, rwConsts&, rwCPb&
Dim SourceSht As Worksheet, DestSht As Worksheet

Set SourceSht = ActiveSheet
On Error Resume Next
ActiveWorkbook.Sheets("Data-reduction params").Delete
On Error GoTo 0
Sheets.Add Before:=Sheets(1)
ActiveSheet.Name = "Data-reduction params"
NoGridlines

With SourceSht

  If pbUPb Then
    rwCPb = 11
    Rw = rwCPb + 1
    frSr(Rw, 3, 3 + Rw).NumberFormat = "@"
    frSr(Rw, 3, 3 + Rw).NumberFormat = "@"
    Cells(Rw, 2) = .[sComm1_64]
    Cells(Rw, 1) = "206/204"
    Cells(Rw, 4) = .[sComm1_74]
    Cells(Rw, 3) = "207/204"
    Rw = Rw + 1
    Cells(Rw, 4) = .[sComm1_84]
    Cells(Rw, 3) = "206/204"
    Cells(Rw, 2) = .[sComm1_76]
    Cells(Rw, 1) = "207/206"
    Rw = Rw + 1
    Cells(Rw, 2) = .[sComm1_86]
    Cells(Rw, 1) = "208/206"
    HA xlRight, rwCPb, 3, rwCPb + 2
    HA xlLeft, rwCPb + 1, 4, rwCPb + 2
    Rw = 13
    rwStdConc = Rw
    On Error GoTo UPbStd
    Cells(Rw, 2) = .[ConcStdPpm].Value
    Cells(Rw, 1) = .[ConcStdPpm].Cells(1, 0)
    Rw = 1 + Rw
    Cells(Rw, 2) = .[ConcStdConst].Value
    Cells(Rw, 1) = .[ConcStdConst].Cells(1, 0)
    Rw = Rw + 2
UPbStd:
    rwUPb = Rw
    On Error GoTo ThPbStd
    Cells(Rw, 1) = .[agestdage].Cells(1, 0)
    Cells(Rw, 2) = .[agestdage].Value
    Rw = 1 + Rw
    Cells(Rw, 1) = .[StdUPbRatio].Cells(1, 0)
    Cells(Rw, 2) = .[StdUPbRatio].Value
    Rw = 1 + Rw
    Cells(Rw, 1) = .[std_76].Cells(1, 0)
    Cells(Rw, 2) = .[std_76].Value
    Rw = 2 + Rw
ThPbStd:
    rwThPb = Rw
    On Error GoTo Consts
    Cells(Rw, 1) = .[stdthpbratio].Cells(1, 0)
    Cells(Rw, 2) = .[agestdage].Value
    Rw = 1 + Rw
    Cells(Rw, 1) = .[StdRad86fact].Cells(1, 0)
    Cells(Rw, 2) = .[StdRad86fact].Value
    Rw = 2 + Rw
  End If

Consts:
  'On Error Resume Next
  rwConsts = IIf(pbUPb, Rw, 8)
  Cells(rwConsts, 1) = "Task Constants"
  With puTask

    If .iNconsts = 0 Then
      Cells(1 + rwConsts, 1) = "None"
    Else

      For i = 1 To .iNconsts
        Cells(rwConsts + i, 1) = .saConstNames(i)
        Cells(rwConsts + i, 2) = .vaConstValues(i)
      Next i

    End If

  End With

  On Error GoTo 0
  HA xlRight, Columns(1)
  HA xlLeft, 1, 1, 6
  HA xlLeft, Columns(2)
  HA xlLeft, Columns(3)
  HA xlLeft, Columns(4)
  ColWidth picAuto, 1, 4
  ColWidth 6, 2
  Fonts rwConsts, 1, , , , True, xlRight
  HA xlRight, rwConsts, 1

  If pbUPb Then
    Fonts rwCPb, 2, , , , True, xlCenter, , , , , , , , , _
          "Common Pb for Age Std"
    HA xlCenter, rwCPb, 2
  End If

  HA xlCenter, rwConsts, 2
  Cells(1, 2) = .Cells(3, 1)
  Cells(2, 1) = "From file:"
  Cells(2, 2) = .Cells(4, 2)

  If pbUPb Then
    Cells(3, 1) = "Spot values for Pb-U-Th Special equations calculated a" & _
      IIf(pbLinfitSpecial, "t mid spot-time", "s spot average")
  End If

  Cells(3 - pbUPb, 1) = "Spot values for other Task equations calculated a" & _
    IIf(pbLinfitEqns, "t mid spot-time", "s spot average")
  Cells(4 - pbUPb, 1) = "Spot values for isotope ratios of similar elements calculated a" & _
    IIf(pbLinfitRats, "t mid spot-time", "s spot average")
  Cells(5 - pbUPb, 1) = "Spot values for isotope ratios of different elements calculated a" & _
      IIf(pbLinfitRatsDiff, "t mid spot-time", "s spot average") ' 09/06/18 - -added

  For i = 3 To 5
    Cells(i + 3 - pbUPb, 1) = .Cells(i, 7)
  Next i

End With
HA xlLeft, 1, 1, 9, 1
Rw = 1

With puTask

  For i = IIf(pbUPb, -4, 1) To .iNeqns

    If i <> 0 Then

      If .saEqns(i) <> "" Then
        Rw = 1 + Rw
        Cells(Rw, 11) = psaEqShow(i, 3)
        Cells(Rw, 12) = psaEqShow(i, 2)
        Cells(Rw, 13) = psaEqShow(i, 1)
      End If

    End If

  Next i

End With

ColWidth picAuto, 11, 13
HA xlRight, , 11, , 11
HA xlCenter, , 12, , 12
HA xlLeft, , 13, , 13
Fonts 1, 2, 2, 2, , True
Fonts 1, 12, , , , True, xlCenter, , , , , , , , , "Task Equations"
Box 1, 1, 8 - pbUPb, 10, 13434828, , True
If pbUPb Then Box rwCPb, 1, 3 + rwCPb, 4, 13434828, , True
Box rwConsts, 1, puTask.iNconsts + rwConsts, 2, 13434828, , True

If pbUPb Then
  Box rwStdConc, 1, 1 + rwStdConc, 2, 13434828, , True
  Box rwUPb, 1, 2 + rwUPb, 2, 13434828, , True
  Box rwThPb, 1, 1 + rwThPb, 2, 13434828, , True
End If

Box 1, 11, 1 + puTask.iNeqns - 3 * pbUPb, 13, peStraw, , True
SourceSht.Activate
End Sub

Sub NoGrids()
NoGridlines
End Sub

Sub HideColumns(Optional IsStd As Boolean = False)
' 09/04/10 -- Add code for hidden CPS columns

Dim EqNum%, PkNum%, Col%, Hdr&

If IsStd Then
  phStdSht.Activate
Else
  phSamSht.Activate
End If

Hdr = flHeaderRow(IsStd)
With puTask

  For EqNum = 1 To .iNeqns

    If .uaSwitches(EqNum).HI And Not ((.uaSwitches(EqNum).SA And IsStd) Or _
      (.uaSwitches(EqNum).ST And Not IsStd)) Then

      FindStr .saEqnNames(EqNum), , Col, Hdr, , Hdr, 256

      If Col > 0 Then
        HideMarkCol 1, Col ' 09/07/21 -- added
'        Columns(Col).Hidden = True
'        With Cells(1, Col).Borders(xlEdgeTop)
'          .Weight = xlHairline
'          .Color = vbWhite
'        End With
        If fsStrip(Cells(Hdr, 1 + Col)) = "%err" Then
          HideMarkCol 1, 1 + Col ' 09/07/21 -- added
'          Columns(1 + Col).Hidden = True
'          With Cells(1, Col).Borders(xlEdgeTop)
'            .Weight = xlHairline
'            .Color = vbWhite
'          End With
        End If
      End If

    End If

  Next EqNum

  For PkNum = IIf(IsStd, 8, 4) To .iNpeaks

    If .baCPScol(PkNum) And .baHiddenMass(PkNum) Then
      HideMarkCol 1, piaCPScol(PkNum) ' 09/07/21 -- added
'      Columns(piaCPScol(PkNum)).Hidden = True
'      With Cells(1, piaCPScol(PkNum)).Borders(xlEdgeTop)
'        .Weight = xlHairline
'        .Color = vbWhite
'      End With
    End If
  Next PkNum

End With
End Sub

Public Sub UnhideRehideColumns()
Dim Col%, EndCol%, HdrRow&

HdrRow = foAp.Max(1, flHeaderRow(0))
EndCol = fiEndCol(HdrRow)
NoUpdate

For Col = 1 To EndCol
  If fbIsSquidHid(1, Col) Then  ' 09/07/21 -- added
'  With Cells(1, Col).Borders(xlEdgeTop)
'    If .Weight = xlHairline And .Color = vbWhite Then
      With Columns(Col)
        .Hidden = Not .Hidden
      End With
'     End If
'  End With
  End If
Next Col

End Sub

Function fhColIndx() As Worksheet
' 09/07/02 -- created
Set fhColIndx = ThisWorkbook.Sheets("ColIndex")
End Function

Function fsColHdr$(ByVal ColIndxNa$, Optional HdrColNum)
' 09/07/02 -- created
Dim rs%, Le%, LeH%, r%, Test$, ShtIn As Worksheet
Set ShtIn = ActiveSheet
fhColIndx.Activate
LeH = Len(ColIndxNa)
rs = 1

Do
  FindStr ColIndxNa, r, , rs, 1, 999, 1
  If r = 0 Then
    fsColHdr = pscNoCol
    ShtIn.Activate
    Exit Function
  End If
  Test = Trim(Cells(r, 1))
  Le = Len(Test)
  rs = 1 + r
Loop Until Le = LeH

If Not IsMissing(HdrColNum) Then HdrColNum = Cells(r, 2)
fsColHdr = Cells(r, 3)
ShtIn.Activate
End Function

Sub FindDriftCorrRanges(Optional RawConstRange As Range, _
    Optional DriftCorrectedRange As Range, Optional RawCol% = 0, _
    Optional CorrCol% = 0)
' '09/07/17 -- added
Dim Hrw%

Set phSamSht = ActiveWorkbook.Sheets(pscSamShtNa)
phSamSht.Activate
Hrw = flHeaderRow(False)

If Not IsMissing(RawConstRange) Then ' 09/07/21 -- added
  FindStr "rawcalibconst", , RawCol, Hrw, , Hrw
End If

FindStr "calibr.const", , CorrCol, Hrw, 1 + RawCol, Hrw

If CorrCol = 0 Then ' 09/07/21 -- removed "or corrcol"
  MsgBox "Unable to locate one or both of the sample-sheet's 'Calibr. Const' column-header.", , pscSq
  CrashEnd
End If

If Not IsMissing(RawConstRange) And RawCol > 0 Then ' 09/07/21 -- mod
  Set RawConstRange = frSr(plaFirstDatRw(0), RawCol, plaLastDatRw(0))
End If

If Not IsMissing(DriftCorrectedRange) Then ' 09/07/21 -- added
  Set DriftCorrectedRange = frSr(plaFirstDatRw(0), CorrCol, plaLastDatRw(0))
End If

End Sub

Sub HideMarkCol(Rw&, Col%)   '  09/07/21 -- added
If Col > 0 Then
  With Cells(Rw, Col).Borders(xlEdgeTop)
    .Weight = xlHairline
    .Color = vbWhite
  End With
  Columns(Col).Hidden = True
End If
End Sub

Function fbIsSquidHid(Rw&, Col%) As Boolean ' 09/07/21 -- added
If Col > 0 Then
  With Cells(Rw, Col).Borders(xlEdgeTop)
    fbIsSquidHid = (.Weight = xlHairline And .Color = vbWhite)
  End With
End If
End Function

Sub StripBracketedRangenames(StringIn$, Optional NumExtracted% = 0)
 ' 10/04/02 - added this sub, called by Sub FormulaEval to correct any range
 '  names mistakenly enclosed in ["   "]
Dim s$, Extr$, ExtrPhrases$(), p%, q%, i%
NumExtracted = 0
s = StringIn

Do
  p = InStr(s, psBrQL)
  q = InStr(s, psBrQR)
If p = 0 Or q = 0 Or q < p Then Exit Do
  Extr = Mid(s, p + 2, q - p - 2)
  s = Left(s, p - 1) & Extr & Mid(s, q + 2)
  NumExtracted = 1 + NumExtracted
  ReDim Preserve ExtrPhrases(NumExtracted)
  ExtrPhrases(NumExtracted) = Extr
Loop

For i = 1 To NumExtracted
  s = ExtrPhrases(i)
  If fbRangeNameExists(s) Then
    Subst StringIn, psBrQL & s & psBrQR, ExtrPhrases(i)
  End If
Next i

End Sub
