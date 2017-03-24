Attribute VB_Name = "StringUtils"
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
' 09/03/03 -- All lower array bounds explicit

Option Explicit
Option Base 1

Function fsInsertStr$(ByVal Host$, ByVal Guest$, ByVal StartPos, ByVal EndPos)
fsInsertStr = Left$(Host, StartPos - 1) & Guest & Mid$(Host, EndPos + 1)
End Function

Function fsS$(ByVal v#)  ' Return space-trimmed number-to-string
fsS = Trim(StR(v))
End Function

Sub TwoSigText(Optional SigLev = 2)
With ActiveChart.TextBoxes.Add(0, 0, 1, 1)
  .AutoSize = True: .AutoscaleFont = False
  .Text = fsS(SigLev) & "s error bars": .Font.Name = "Arial"
  .Characters(Start:=2, Length:=1).Font.Name = "Symbol"
  .VerticalAlignment = xlTop: .HorizontalAlignment = xlRight
  .Font.Size = 11:
  .Left = ActiveChart.ChartArea.Width - .Width
End With
End Sub

Function fsStrip$(ByVal s$, Optional IgCase As Boolean = True, _
  Optional IgSpaces As Boolean = True, Optional IgDashes As Boolean = True, _
  Optional IgSlashes As Boolean = True, Optional IgSht As Boolean = False, _
  Optional IgCommas As Boolean = False, Optional IgColons As Boolean = False, _
  Optional IgSemicolons As Boolean = False, Optional IgPeriods As Boolean = False, _
  Optional IgLinefeeds As Boolean = True, Optional IgVertSlash = False, _
  Optional IgNonAlphaNum = False)
' fsStrip " ", "/", "-" from a string (for Std-name recognition)
Dim i%
If IgSpaces Then Subst s, " "

If IgNonAlphaNum Then

  For i = 33 To 200
    If i < 48 Or (i > 57 And i < 65) Or (i > 90 And i < 97) Or i > 122 Then
      Subst s, Chr(i)
    End If
  Next i

Else
  If IgSlashes Then Subst s, "/"
  If IgDashes Then Subst s, "-"
  If IgCommas Then Subst s, ","
  If IgSemicolons Then Subst s, ";"
  If IgColons Then Subst s, ":"
  If IgPeriods Then Subst s, "."
  If IgLinefeeds Then Subst s, vbLf
  If IgVertSlash Then Subst s, "|"
  If IgSht Then
    Subst s, "?", , "*", , ":"
  End If
End If

If IgCase Then s = LCase(s)
fsStrip = s
End Function

Function fsRatioHdrStr$(ByVal NumerIso#, ByVal DenomIso#)
If NumerIso < 0 Or DenomIso < 0 Then
  MsgBox "SQUID error in Function fsIsRatL$ -- NumerIso or DenomIso passed as zero.", , pscSq
  CrashEnd
End If
fsRatioHdrStr = fsS(Drnd(NumerIso, 5)) & "|/" & fsS(Drnd(DenomIso, 5))
End Function

Function fbIsNum(ByVal v, Optional NonzeroOnly = False, _
  Optional PositiveOnly = False) As Boolean
fbIsNum = False

If IsNumeric(v) Then
  If PositiveOnly Then
    fbIsNum = (v > 0)
  ElseIf NonzeroOnly Then
    fbIsNum = (v <> 0)
  Else
    fbIsNum = True
  End If
End If

End Function

Function fbIsNumChar(ByVal v, Optional IntOnly, Optional NoSign) As Boolean
Dim s$, a%
s = Trim(v): fbIsNumChar = False
DefVal IntOnly, False
DefVal NoSign, False
If s = "" Then Exit Function
a = Asc(s)
If (a > 47 And a < 58) Or _
  (Not NoSign And (a = 43 Or a = 45) Or _
  (Not IntOnly And a = 46)) Then fbIsNumChar = True
End Function

Function fbIsAllNumChars(ByVal Phrase$, Optional IntOnly As Boolean = False, _
  Optional NoSign As Boolean = False) As Boolean
Dim i%

If Phrase = "" Then
  fbIsAllNumChars = False: Exit Function
Else

  For i = 1 To Len(Phrase$)
    If Not fbIsNumChar(Mid$(Phrase, i, 1), IntOnly, NoSign) Then _
      fbIsAllNumChars = False: Exit Function
  Next i

End If

fbIsAllNumChars = True
End Function

Function fbIsAlphaChar(ByVal Ch) As Boolean
Dim p
p = Asc(UCase(Ch))
fbIsAlphaChar = (p > 64 And p < 91)
End Function

Function fbIsAllAlphaChars(ByVal Schar$) As Boolean
Dim s$, a%, i%

For i = 1 To Len(Schar)
  s = LCase(Trim(Mid$(Schar, i, 1)))
  a = Asc(s)
  If (a < 97 Or a > 122) Then fbIsAllAlphaChars = False: Exit Function
Next i

fbIsAllAlphaChars = True
End Function

Function fbNoNum(ByVal q) As Boolean
fbNoNum = Not IsNumeric(q)
End Function

Function fbIsBkrd(ByVal s$) As Boolean
Dim b As Boolean
b = True

Select Case Left$(LCase(s$), 4)
  Case "back"
  Case "bkgr"
  Case "bkrd"
  Case "zero"
  Case "base"
  Case "bkr"
  Case "bk"
  Case Else: b = False
End Select

fbIsBkrd = b
End Function

Sub Cpos(ContainerStr$, ByVal ItemStr$, StrPos%(), Nitems%)
' Find all locations if ItemStr within ContainerStr, and put in
'  the StrPos array.
' Delimiters at beginning & end of string are ignored, as are
'  adjacent delimiters.
' ContainterStr is returned as space-trimmed with repeated elimiters removed!!!!
Dim i%, p%, q%, d$, e$, f$

Nitems = 0: q = 0
ContainerStr = Trim(ContainerStr)

If Len(ItemStr) = 1 Then

  For i = 1 To Len(ContainerStr)
    e = Mid$(ContainerStr, i, 1)

    If i = 1 Then
      d = d & e
    Else
      f = Mid$(ContainerStr, i - 1, 1)
      If e <> ItemStr Or f <> e Then d = d & e
    End If

  Next i

Else
  d = ContainerStr
End If

ContainerStr = d

Do
  p = InStr(d, ItemStr)
If p = 0 Then Exit Do
  Nitems = 1 + Nitems
  ReDim Preserve StrPos(1 To Nitems)
  q = p + q
  StrPos(Nitems) = q
  d = Mid$(d, 1 + p)
Loop

End Sub

Sub NumDenom(ByVal Rat$, Numer#, Denom#)
Dim p%
Numer = 0: Denom = 0
p = InStr(Rat, "/")

If p > 1 Then
  Numer = Val(Left$(Rat, p - 1))
  Denom = Val(Mid$(Rat, p + 1))
End If
End Sub

Function fbIsNumber(ByVal Phrase$) As Boolean ' Does Phrase contain only numbers (or ".")?
Dim s$, Le%, i%, Asc_%

Phrase = Trim(Phrase)
s = Left$(Phrase, 1)
If Phrase = "" Then fbIsNumber = False: Exit Function

If s = "-" Or s = "+" Then Phrase = Mid$(Phrase, 2)
s = Left$(Right$(Phrase, 4), 2)
Le = Len(Phrase)

If s = "E-" Or s = "E+" Then
  Phrase = Left$(Phrase, Le - 4)
Else
  s = Left$(Right$(Phrase, 3), 2): Le = Len(Phrase)
  If s = "E-" Or s = "E+" Then Phrase = Left$(Phrase, Le - 3)
End If

For i = 1 To Len(Phrase)
  Asc_ = Asc(Mid$(Phrase, i, 1))
  If Asc_ = 47 Or Asc_ < 46 Or (Asc_ > 57 And (Asc_ <> 69 Or i > 1)) Then _
    fbIsNumber = False: Exit Function
Next i

fbIsNumber = True
End Function

Function fsExtractPart(ByVal Phrase$, ByVal Instance%, ByVal Delim$, Optional Delim2)
Dim q%, Loc%(), s$, LocCt%

VIM Delim2, Delim
AllInstanceLoc Delim, Phrase, Loc, LocCt

If LocCt >= Instance Then
  s = Mid$(Phrase, 1 + Loc(Instance))
  q = InStr(s, Delim2)
  If q > 0 Then
    fsExtractPart = Left$(s, q - 1)
  End If
Else
  fsExtractPart = ""
End If

End Function

Function fsSubStr$(ByVal s$, ByVal StartAt%, ByVal EndAt%)
fsSubStr = Mid$(s, StartAt, EndAt - StartAt + 1)
End Function

Function fsExtrLeft$(ByVal Phrase$, ByVal Delim$, Optional LastDelim As Boolean = False)
Dim Rv$, s$, p%
s = ""

If LastDelim Then
  Rv = StrReverse(Phrase)
  p = InStr(Rv, Delim)
  If p > 0 Then s = StrReverse(Mid$(Rv, p + 1))
Else
  p = InStr(Phrase, Delim)
  If p > 0 Then s = Left$(Phrase, p - 1)
End If

fsExtrLeft = s
End Function

Function fsExtrRight$(Phrase$, Delim$, Optional LastDelim As Boolean = False)
Dim Rv$, s$, p%
s = ""

If LastDelim Then
  Rv = StrReverse(Phrase)
  p = InStr(Rv, Delim)
  If p > 0 Then s = StrReverse(Left$(Rv, p - 1))
Else
  p = InStr(Phrase, Delim)
  If p > 0 Then s = Mid$(Phrase, p + 1)
End If
fsExtrRight = s

End Function

Sub ParseLine(ByVal Inp$, Outp As Variant, Optional N% = 0, Optional Delim)
' Parse a string into an array of values  (numeric, string, or mixed)
' Delim can be either the ASCII code of a character or the character itself.
Dim s$, d$, Nitems%, Le%, Le0%, a%(), i&, Nn&, zOutp As Variant

If Len(Trim(Inp$)) = 0 Then
  N = 0
  Exit Sub
End If

If fbIM(Delim) Then
  d = Chr(9)
ElseIf IsNumeric(Delim) Then
  d = Chr(Delim)
Else
  d = Delim
End If

Inp = Trim(Inp)

If d = " " Then
  Do ' remove consecutive space-delimiters
    Le0 = Len(Inp)
    Subst Inp, d & d, d
    Le = Len(Inp)
  Loop Until Le = Le0
End If

zOutp = Split(Inp, d)
Nn = 1 + UBound(zOutp)
ReDim Outp(1 To Nn)

For i = 1 To Nn
  Outp(i) = zOutp(i - 1)
Next
If Nn < (2 ^ 15) Then
  N = Nn
End If
Exit Sub

Le = Len(d)
s = Trim(Inp$): N = 0
If s = "" Or s = d Then Exit Sub

If InStr(Inp, d) > 0 Then
  Cpos Inp, d, a(), Nitems
  If Nitems = 0 Then Exit Sub
  N = 1 + Nitems
  ReDim Outp(1 To N)

  For i = 2 To N - 1
    Outp(i) = Mid$(Inp, Le + a(i - 1), a(i) - a(i - 1) - Le)
  Next i

  Outp(1) = Left$(Inp, a(1) - 1)
  Outp(N) = Mid$(Inp, Le + a(N - 1))
Else
  ReDim Outp(1 To 1)
  N = 1: Outp(1) = Inp
End If

If N > 1 And Right$(Inp, 1) = d Then
  N = N - 1
  ReDim Preserve Outp(1 To N)
End If

End Sub

Sub GetSaveTaskVal(ByVal Get1Save2%, ParamVal, Optional ByVal ParamRow, _
  Optional ByVal ParamName, Optional ByVal LookInCol% = 1, Optional RowFound&, _
  Optional ByVal ParamCol% = 2, Optional Bad As Boolean = False)
Dim Rw&, v As Variant

If fbNIM(ParamRow) Then
  If fbIM(ParamName) Then MsgBox "GetSaveTaskVal needs either " & _
  "  ParamRow or ParamName.", pscSq: End
  Rw = ParamRow
Else
  FindStr ParamName, Rw, , 1, LookInCol, 99, LookInCol
End If

If fbNIM(RowFound) Then RowFound = Rw
If Rw = 0 Then Bad = True: Exit Sub
v = Cells(Rw, ParamCol).Formula
If v = "TRUE" Or v = "FALSE" Then v = CBool(v)
If v = "" Then v = 0
pwTaskBook.Sheets(1).Columns(1).ColumnWidth = 24
frSr(1 + puTask.laAutoGrfRw(peMaxAutochts), 1, 9999, 256).Clear

If Get1Save2 = 1 Then
  ParamVal = v
Else
  Cells(Rw, ParamCol).Formula = v
End If
End Sub

Sub FindStr(ByVal Phrase$, Optional RowFound, Optional ColFound, Optional ByVal RowLook1 = 1, _
  Optional ByVal ColLook1 = 1, Optional ByVal RowLook2, Optional ByVal ColLook2 = 255, _
  Optional CaseSensitive As Boolean = False, Optional InclLineFeeds As Boolean = False, _
  Optional InclSpaces = False, Optional InclDashes = False, Optional InclPeriods = False, _
  Optional InclSlashes = False, Optional InclColons = False, Optional InclSemicolons = False, _
  Optional InclVertSlashes = False, Optional WholeWord = False, Optional InclAsterisks = True, _
  Optional LegalSheetNameOnly = False, Optional LegalRangeNameOnly As Boolean = False, _
  Optional InclCommas As Boolean = False, Optional ByVal ColIndxName$ = "")

' Find the rowfound/colfound of a string in a range
Dim tB As Boolean, CellIndx&, ra As Range, s$

If Phrase = "" Then GoTo Failed
DefVal RowLook2, RowLook1
If Not CaseSensitive Then Phrase = LCase(Phrase)

If LegalSheetNameOnly Or LegalRangeNameOnly Then
  Phrase = fsLegalName(Phrase, LegalRangeNameOnly, , False)
Else
  If Not InclSpaces Then Subst Phrase, " "
  If Not InclPeriods Then Subst Phrase, "."
  If Not InclDashes Then Subst Phrase, "-"
  If Not InclSlashes Then Subst Phrase, "/"
  If Not InclColons Then Subst Phrase, ":"
  If Not InclCommas Then Subst Phrase, ","
  If Not InclSemicolons Then Subst Phrase, ";"
  If Not InclVertSlashes Then Subst Phrase, "|"
  If Not InclAsterisks Then Subst Phrase, "*"
  If InclLineFeeds Then
    Phrase = fsVertToLF(Phrase) ' change  to linefeed
  Else
    Subst Phrase, "|", , vbLf
  End If
End If
If Not InclSpaces Then Subst Phrase, " ", ""
If RowLook1 = 0 Or ColLook1 = 0 Or RowLook2 = 0 Or _
  (fbNIM(ColLook2) And ColLook2) = 0 Then GoTo Failed
CheckRC 100, RowLook1, ColLook1, RowLook2, ColLook2
On Error GoTo Failed

Set ra = frSr(RowLook1, ColLook1, RowLook2, ColLook2)

For CellIndx = 1 To ra.Cells.Count
  s = ""
  On Error Resume Next
  s = IIf(CaseSensitive, ra(CellIndx).Text, LCase(ra(CellIndx).Text))
  On Error GoTo 0

  If s <> "" Then
    If Not InclLineFeeds Then
      Subst s, "|", , vbLf
    End If

    If LegalSheetNameOnly Or LegalRangeNameOnly Then
      s = fsLegalName(s, LegalRangeNameOnly, , False)
    Else
      If Not InclSlashes Then Subst s, "/"
      If Not InclPeriods Then Subst s, "."
      If Not InclColons Then Subst s, ":"
      If Not InclSemicolons Then Subst s, ";"
      If Not InclCommas Then Subst s, ","
      If Not InclVertSlashes Then Subst s, "|"
      If Not InclAsterisks Then Subst Phrase, "*"
    End If

    If Not InclSpaces Then Subst s, " ", ""
    If Not InclDashes Then Subst s, "-", ""

    If WholeWord Then
      tB = (s = Phrase)
    Else
      tB = (InStr(s, Phrase) > 0)
    End If

    If tB Then
      RowFound = ra(CellIndx).Row
      ColFound = ra(CellIndx).Column
      RefreshColIndx ColIndxName, RowFound
      Exit Sub
    End If

  End If

Next CellIndx

Failed: RowFound = 0: ColFound = 0
End Sub

Sub ErrFor(v#, e#, Optional ErrSigFigs As Integer = 2)
' Return e as 2 sig-figs, sigfigs of v to match
Dim Le#, tv#, e1#, e2#, absE#, Delt#, Small#

If e <= 0 Then
  e = 0
  v = Drnd(v, 3)
Else
  tv = v: absE = Abs(e): Small = 0.000000000001
  With foAp
    Le = -fdLog10(Abs(absE))
    tv = .Fixed(v, ErrSigFigs + Le)
  End With
  If v <> 0 And tv = 0 Then tv = Drnd(v, 1)
  e1 = Drnd(absE, 1): e2 = Drnd(absE, 2)
  Delt = Abs(e1 - e2)

  If absE < 10 And Delt < Small Then
    If absE >= 1 And absE < 10 Then e2 = e2 & "."
    e2 = e2 & pscZq
  End If

  v = tv: e = e2
End If
End Sub

Sub SigConv(rw1 As Variant, Optional Col1, Optional rw2, Optional Col2)
' Replace "sigma" with greek-letter sigma
Dim q As String * 1, s$, p%, ra As Range

CheckRC 120, rw1, Col1, rw2, Col2
q = "�"

For Each ra In frSr(rw1, Col1, rw2, Col2).Cells
  With ra
    s$ = fsApSub("sigma", .Text, q)
    p = InStr(s, q)

    If p > 0 Then
      .Formula = s
      With .Characters(p, 1)
        .Font.Name = "Symbol": .Text = "s"
      End With
    End If

  End With
Next ra
End Sub

Sub Subst(Phrase, This$, Optional WithThis$ = "", Optional This2$ = "", _
  Optional WithThis2$ = "", Optional This3$ = "", Optional WithThis3$ = "", _
  Optional This4$ = "", Optional withThis4$ = "")
Dim This_$, WithThis_$, ThisIndx%

For ThisIndx = 1 To 4
  This_ = Choose(ThisIndx, This, This2, This3, This4)
  WithThis_ = Choose(ThisIndx, WithThis, WithThis2, WithThis3, withThis4)

  If This_ <> "" Then
    Phrase = Replace(Phrase, This_, WithThis_)
  End If

Next ThisIndx

End Sub

Sub SymbChar(ByVal Rw&, ByVal c%, ByVal CharPos%)
Cells(Rw, c).Characters(CharPos, 1).Font.Name = "Symbol"
End Sub

Function fsQq$(ByVal s$) 'Replace $ char with "
fsQq = fsApSub("$", s$, pscQ)
End Function

Function fsBracketReplace(ByVal Phrase$, ByVal ToBeContained$, _
  Optional AllDelims As Boolean = False) As String
Dim BrkRepl$, DelimIndx%, p%, q%, Ndelims%, Delims As Variant

Ndelims = 1 - 2 * AllDelims
Delims = Array("[]", "{}", "<>")
BrkRepl = Phrase

For DelimIndx = 1 To Ndelims
  p = InStr(BrkRepl, Mid$(Delims(DelimIndx), 1, 1))
  q = InStr(BrkRepl, Mid$(Delims(DelimIndx), 2, 1))
  If q > p Then BrkRepl = Left$(BrkRepl, p - 1) & ToBeContained & Mid$(BrkRepl, q + 1)
Next DelimIndx

fsBracketReplace = BrkRepl
End Function

Function fsStrExtract$(ByVal Phrase$, ByVal StartDelim$, ByVal EndDelim$, _
      Optional Trimm As Boolean = True, Optional Stripp As Boolean = False, _
      Optional StringArrayV, Optional NumberExtracted%)

Dim ExtractedString$, p%, q%

p = InStr(Phrase, StartDelim)
q = InStr(Phrase, EndDelim)

If p > 0 And q > p Then
  ExtractedString = Mid$(Phrase, p + Len(StartDelim), q - p - Len(EndDelim))

  If Stripp Then
    ExtractedString = fsStrip(s:=ExtractedString, IgCommas:=True, _
      IgColons:=True, IgSemicolons:=True, IgPeriods:=True, _
      IgLinefeeds:=True, IgNonAlphaNum:=True)
  ElseIf Trimm Then
    ExtractedString = Trim(ExtractedString)
  End If

  If fbNIM(StringArrayV) Then
    ParseLine ExtractedString, StringArrayV, NumberExtracted, " "
    fsStrExtract = ExtractedString
  Else
    fsStrExtract = ExtractedString
    NumberExtracted = 1
  End If
Else
  fsStrExtract = ""
  NumberExtracted = 0
End If

End Function

Sub SqBrakQuExtract(ByVal s$, Extracted$, Optional N%, Optional Trimm As Boolean = True, _
  Optional Stripp As Boolean = False, Optional Arr)
Extracted = fsStrExtract(s, psBrQL, psBrQR, Trimm, Stripp, Arr, N)
End Sub

Sub SqBrakExtract(ByVal s$, Extracted$, Optional N%, Optional Trimm As Boolean = True, _
  Optional Stripp As Boolean = False, Optional Arr)
Extracted = fsStrExtract(s, "[", "]", Trimm, Stripp, Arr, N)
End Sub

Sub CurlyExtract(Phrase$, Extracted$, Optional N%, Optional Trimm As Boolean = True, _
  Optional Stripp As Boolean = False, Optional Arr, Optional RemoveFromPhrase As Boolean = False)
Extracted = fsStrExtract(Phrase, "{", "}", False, False, Arr, N)

If RemoveFromPhrase Then
  Subst Phrase, "{" & Extracted & "}"
End If

If Trimm Then Extracted = Trim(Extracted): Phrase = Trim(Phrase)
End Sub

Sub Clean(DatRange As Range, CleanedDat As Range, NumCleanRows%, _
  Optional ZeroesOK As Boolean = False, Optional BlankOk As Boolean = False, _
  Optional AllColsOK As Boolean = False, Optional BothNegPos As Boolean = True, _
  Optional StrikeThruOK As Boolean = False, Optional AllComers As Boolean = False, _
  Optional AddStrikeThru As Boolean = False)
' Returns Cleandat as array cleaned of all noncomplying rows.

Dim First As Boolean, OkCel As Boolean
Dim Nareas%, Ncols%, OKcol%, Col%, AreaIndx%, TempNum%, CleanedRowCt%
Dim Rw&, RowCt&, v#
Dim Cel As Range, Area As Range, Crow As Range

First = True
With DatRange
  Ncols = IIf(AllColsOK, .Columns.Count, 1)
  CleanedRowCt = 0
  Nareas = .Areas.Count

  For AreaIndx = 1 To Nareas
    TempNum = 1 + CleanedRowCt
    Set Area = .Areas(AreaIndx)
    With Area
      RowCt = .Rows.Count

      For Rw = 1 To RowCt
        TempNum = CleanedRowCt + 1
        OKcol = 0

        For Col = 1 To Ncols
          Set Cel = .Item(Rw, Col)
          OkCel = False

          If BlankOk Or Cel.Formula <> "" Or AllComers Then

            If IsNumeric(Cel) Or AllComers Then
              With Cel
                v = Cel.Value

                If ZeroesOK Or v <> 0 Or AllComers Then
                  If BothNegPos Or v > 0 Or AllComers Then

                    If StrikeThruOK Or Not .Font.Strikethrough Or AllComers Then
                      OKcol = 1 + OKcol
                      OkCel = True
                    End If

                  End If

                End If
              End With

            End If

          End If

          If Not OkCel And AddStrikeThru Then
            Cel.Font.Strikethrough = True
          End If

          If OKcol < Col Then Exit For
        Next Col

        If OKcol = Ncols Then
          CleanedRowCt = 1 + CleanedRowCt
          Set Crow = Range(.Item(Rw, 1), .Item(Rw, Ncols))

          If First Then
            Set CleanedDat = Crow
            First = False
          Else
            Set CleanedDat = Union(CleanedDat, Crow)
          End If

        End If

      Next Rw

    End With
  Next AreaIndx

End With
NumCleanRows = CleanedRowCt
End Sub

Public Function BiWt(RangeIn, Optional Tuning, Optional SingleValOut As Boolean = False, _
  Optional AllComers As Boolean = False, Optional IsoplotStyle As Boolean = False)

Dim CleanCt%, Area%, CleanedRowCt%, Rw&
Dim BiWtSigma#, BiWt95#, BiWtMean#, CleanedNumeric#()
Dim RI As Range, CleanedRange As Range, OutpRange(1 To 3, 1 To 2) As Variant

If TypeName(RangeIn) <> "Range" Then Exit Function

Set RI = RangeIn
Clean RI, CleanedRange, CleanedRowCt, , , , , , AllComers
If CleanedRowCt = 0 Then Exit Function
ReDim CleanedNumeric(1 To CleanedRowCt)
CleanCt = 0

With CleanedRange

  For Area = 1 To .Areas.Count
    With .Areas(Area)

      For Rw = 1 To .Rows.Count
        CleanCt = 1 + CleanCt
        CleanedNumeric(CleanCt) = .Rows(Rw)
      Next Rw

    End With
  Next Area

End With

VIM Tuning, 6
'ImN Tuning, 6
TukeysBiweight CleanedNumeric, CleanCt, BiWtMean, (Tuning), BiWtSigma, BiWt95

If SingleValOut Then
  BiWt = BiWtMean
ElseIf IsoplotStyle Then
  OutpRange(1, 1) = BiWtMean:  OutpRange(1, 2) = "Biwt Mean"
  OutpRange(2, 1) = BiWtSigma: OutpRange(2, 2) = "Biwt Sigma"
  OutpRange(3, 1) = BiWt95:    OutpRange(3, 2) = pscPm & "95%conf"
Else
  OutpRange(1, 1) = BiWtMean
  OutpRange(1, 2) = BiWtSigma
End If

BiWt = OutpRange
End Function

Sub CreateMergedSubsetRange(ByVal EqNum%, MergedRangeName$, Col%, _
  ByVal StdCalc As Boolean, LastRow&)

Dim NumPieces%, Rw&, Nrows&
Dim ColRange As Range, MergedRange As Range, Spot_Name As Variant

' 09/06/09 -- add line below
Rw = flHeaderRow(-StdCalc) ' to set plaFirstDatRw

Set ColRange = frSr(plaFirstDatRw(-StdCalc), Col, LastRow)
Nrows = ColRange.Rows.Count
NumPieces = 0

For Rw = 1 To Nrows

  If pbUPb Then
    Spot_Name = Cells(plaFirstDatRw(-StdCalc) + Rw - 1, 1)
  Else
    Spot_Name = psaSpotNames(Rw)
  End If

  If fbIsInSubset(Spot_Name, EqNum) Then

    If NumPieces = 0 Then
      Set MergedRange = ColRange(Rw)
    Else
      Set MergedRange = Union(MergedRange, ColRange(Rw))
    End If

    NumPieces = 1 + NumPieces
  End If

Next Rw

If NumPieces > 0 Then
  MergedRangeName = MergedRange.Address
Else
  MergedRangeName = ""
End If
End Sub

Sub Formulae(ByVal Eqn$, ByVal EqNum%, ByVal StdCalc As Boolean, _
   Optional ByVal OutputRow = 0, Optional ByVal OutputCol% = 0, _
   Optional ByVal InputRow% = 0)
' 09/06/10 -- Rewrite code associated with external workbook and/or worksheet references
' Either OutputCol or OutputRow must be specified!

' Parse square-bracketed expressions in user-defined equation# EqNum,
'   replace with the appropriate column- or cell-addresses,
'   and place the parsed formula in the appropriate cell.

' Parse each bracketed expression serially, in each determining if:
'  1) A worksheet reference in the form ([TestData.PD]SampleData!$B$8:$C$8)
'     (the parens must have been placed by the user when defning the equation),
'  2) An isotope ratio index in the form B or AB (1 or 2 alphabetic chars),
'  3) A user-defined equation index# in the form of 1 or 2 numeric chars,
'  4) An EXISTING column header,
'  5) An EXISTING range name,
'  6) A yet-to-be defined range name.

Dim RefToOtherSht As Boolean, tB As Boolean, RefIsToStd As Boolean, ErCol As Boolean
Dim OneCellRef As Boolean, ByRangeName As Boolean, PossibleRangeName As Boolean, StdRef As Boolean
Dim IsExistingRangeRef As Boolean, BadWbk As Boolean, BadSht As Boolean, SCref As Boolean

Dim EqFrag$, BrkExtr$, Msg$, TestRef$, tmp$, tmp2$, Na$, IndxRangeName$, ShtRef$
Dim WbkRef$, WkshtRef$, EqnIn$, EqnLeft$, EqnRight$, TestWkSht$, WbkNa$(), WshtNa$()
Dim Col%, OutputNcols%, OutputNrows%, i%, j%, ExclPos%, Nrefs%, Ref%
Dim LeftSqBrkPos%, RightSqBrkPos%, LeftBrk%, RightBrk%, IndxType%, Indx%, EqnIndx%
Dim SingleQuotePos%, LeftParenSqBrk%, p%, q%, h%, LeftPos%, RightPos%
Dim FirstRow&, LastRow&
Dim OutputCellRange As Range, OutputRangeNameCell As Range, IndxRange As Range
Dim OtherSheet As Worksheet, WkshtIn As Worksheet, WbkIn As Workbook, TestWbk As Workbook

' NOTE: Chr(34) = "
'       Chr(147)= �
'       Chr(148)= �

FirstRow = plaFirstDatRw(-StdCalc): LastRow = plaLastDatRw(-StdCalc)

EqnIn = Eqn
Eqn = EqnIn

p = InStr(Eqn, "<=>")
If p > 0 Then Eqn = Left$(Eqn, p - 1)

Subst Eqn, Chr(147), pscQ, Chr(148), pscQ
Set WbkIn = ActiveWorkbook
Set WkshtIn = ActiveSheet

If pbUPb Then
  If WkshtIn.Name = pscStdShtNa And Not pbStdsOnly Then
    Set OtherSheet = phSamSht
  Else
    Set OtherSheet = phStdSht
  End If
End If

FindWbkShtRefs Eqn, Nrefs, WbkNa, WshtNa, BadWbk, BadSht

For Ref = 1 To Nrefs          ' Extract workbook references in the eqn

  If WbkNa(Ref) <> "" Then
    TestRef = WbkNa(Ref)
    On Error GoTo BadWbkRef   ' Is the workbook loaded?
    Set TestWbk = Workbooks(TestRef)
    On Error GoTo 0           ' "Hide" the workbook-indicating brackets
    Subst Eqn, "[" & TestRef & "]", "###" & fsS(Ref) & "###"
  End If

Next Ref

With puTask.uaSwitches(EqNum)
  If .Ar And .ArrNrows > 1 Then .SC = True

  If .SC Then
    On Error Resume Next
    foAp.Calculate
    On Error GoTo 0

    If pbUPb Then
      LastRow = flEndRow(1)
    Else
      LastRow = FirstRow + piaNumSpots(0) - 1
    End If

  End If


  Do
    OneCellRef = False: ByRangeName = False
    PossibleRangeName = False
    EqFrag = "" ' First hide any worksheet refs
    LeftSqBrkPos = InStr(Eqn, "["):     RightSqBrkPos = InStr(Eqn, "]")
    LeftBrk = InStr(Eqn, "[" & pscQ):   RightBrk = InStr(Eqn, pscQ & "]")
    LeftParenSqBrk = InStr(Eqn, "([")
    ' Indx=0 if nothing in brackets, -n if an isot-rat#, +n if an eqn#,
    '  1000 + Colnum if an existing column-header, -1000-ConstNum if a defined constant.

    With puTask
      ExtractEqnRef Eqn, BrkExtr, Indx, IndxType, .saIsoRats, .saEqnNames, _
         , , , , ErCol, , WbkRef, WkshtRef
    End With
    Msg = ""
    IsExistingRangeRef = False: Col = 0: OneCellRef = True
    tmp = fsLegalName(BrkExtr, True, , False)

    If fbRangeNameExists((tmp)) Then
      IsExistingRangeRef = True
      If IndxType = peEquation Then Col = piaEqCol(StdCalc, Indx)
    Else

      Select Case IndxType
        Case pePrefsConstant Or peTaskConstant Or peBothConstant
          MsgBox "SQUID error: ref to constant found in Sub Formulae.", , pscSq
          CrashEnd

        Case peRatio
          Indx = -Indx     ' Indx=ratio-index#
          Col = piaIsoRatCol(Indx)
          If .SC And Not (.Ar And .ArrNrows = 1) Then OneCellRef = False

        Case peColumnHeader  ' a column-header
          Indx = Indx - 1000
          Col = Indx

        Case peEquation      ' Indx=eqnIndex#
          Col = piaEqCol(StdCalc, Indx)
          If .SC Then ByRangeName = True
          RefToOtherSht = False
          If pbUPb And ((StdCalc And puTask.uaSwitches(Indx).SA) Or _
            (Not StdCalc And puTask.uaSwitches(Indx).ST)) Then RefToOtherSht = True
          p = InStr(Eqn, BrkExtr): q = 0
          WbkRef = "": ShtRef = ""

          If p > 4 Then ' Does a worksheet reference preceed the extracted reference?

            If Mid(Eqn, p - 4, 4) = "'!" & psBrQL Then
              q = p - 4
            ElseIf Mid(Eqn, p - 3, 3) = "'![" Then
              q = p - 3
            End If

            If q > 0 Then                         ' Yes, it does.
              tmp = StrReverse(Left(Eqn, q - 1))  ' Does a workbook reference preceed the
              p = InStr(tmp, "'")                 '   worksheet reference?
              TestRef = ""
              If p = 0 Then GoTo BadWbkRef
              tmp = StrReverse(Left(tmp, p - 1))  ' Look for wbk ref of the form ###2###

              If Left(tmp, 3) = "###" Then
                tmp2 = Mid(tmp, 4, 1)

                If IsNumeric(tmp2) And Mid(tmp, 7, 1) = "#" Then
                  q = Val(Mid(tmp, 4))                 ' = the wbk-reference index for WbkNa()

                  If q > 0 Then
                    WbkRef = WbkNa(q)
                    Set TestWbk = Workbooks(WbkRef)    ' No need to test again
                    TestRef = Mid(tmp, 8)
                    On Error GoTo BadWkShtRef           ' Does the specified worksheet exist in the
                    tmp = TestWbk.Sheets(TestRef).Name  '   referenced workbook?
                    On Error GoTo 0
                    ShtRef = TestRef
                  End If

                End If

              Else
                TestRef = tmp                          ' Does the sheet exist in the active workbook?
                On Error GoTo BadWkShtRef
                TestWkSht = WbkIn.Sheets(TestRef).Name ' 10/04/04 - added the ".Name" to fix error
                On Error GoTo 0
                ShtRef = TestRef
              End If

            End If

          End If

          ShtRef = LCase(ShtRef)

          If pbUPb Then
            StdRef = StdCalc Or ShtRef = LCase(pscStdShtNa)
          Else
            StdRef = False
          End If

          Col = piaEqCol(StdRef, Indx)

          If Col = 0 And Not ByRangeName Then ' range does not exist yet
            Col = fiFindHeader(BrkExtr, , 0 * True)

            If Col = 0 Then
              tmp = LCase(fsLegalName(BrkExtr, True, , False))

              For i = 1 To puTask.iNeqns

                If puTask.saEqnNames(i) <> "" Then
                  If LCase(fsLegalName(puTask.saEqnNames(i), True, , False)) = tmp Then Exit For
                End If

              Next i

              If i < puTask.iNeqns Then
                Indx = i
                Col = piaEqCol(StdCalc, Indx)
                If puTask.uaSwitches(i).SC Then OneCellRef = False
              End If

            End If

          End If

        Case Else
          IsExistingRangeRef = fbRangeNameExists(fsLegalName(BrkExtr, True, , False))
          If Not IsExistingRangeRef Then PossibleRangeName = True
      End Select

    End If

    If (ByRangeName Or IsExistingRangeRef) Then

      If fbRangeNameExists(BrkExtr) Then
        IndxRangeName = fsLegalName(BrkExtr, True, , False)
        Set IndxRange = Range(IndxRangeName)
        On Error GoTo 1

        If IndxRange.Count = 1 Then
          OneCellRef = True
1:        On Error GoTo 0
        End If

      End If

'    ElseIf (IndxType = peRatio Or IndxType = peEquation Or IndxType = peColumnHeader) _
'            And (.SC Or (.Ar And .ArrNrows > 1)) Then
    ElseIf (IndxType = peRatio Or IndxType = peEquation Or IndxType = peColumnHeader) _
            And (.SC Or (.Ar And .ArrNrows > 1)) Then
      OneCellRef = False
    End If

    If IndxType = peEquation And Indx > EqNum And EqNum > 0 And .Nu Then 'Or (Not .SC And Not .AR)) Then
      tmp = "User equation " & fsS(EqNum) & " cannot be numerically evaluated because some of its " _
        & vbLf & "required values have not yet been calculated." & pscLF2 & _
        "To avoid this problem, add the FO switch to the equation definition." & _
        pscLF2 & "Abandon data reduction now?"
      If MsgBox(tmp, vbYesNo, pscSq) <> vbNo Then CrashEnd
    End If

    If .SC And Not OneCellRef And prSubsSpotNameFr(EqNum) <> "" Then
      ' Construct a (discontinuous) range address that refers to only those numeric cells
      '   which are in the subset specified by prSubsSpotNameFr(EqNum) and are not
      '   StruckThrough.
      CreateMergedSubsetRange EqNum, EqFrag, Col, StdCalc, (LastRow)

      If EqFrag = "" Then Exit Sub ' No spot names in the worksheet match the subset name

      ' PERMITS NATIVE EXCEL FUNCTIONS TO WORK ON DISCONTINUOUS RANGES
      '  BUT NOT USER FUNCTIONS.
    ElseIf IsExistingRangeRef Or PossibleRangeName Or ByRangeName Then
      tB = (Indx > 0 And Indx <= puTask.iNeqns)

      If tB And IndxType = peEquation Then

        If puTask.uaSwitches(Indx).SC Then
          EqFrag = Cells(FirstRow, Col).Address
        Else
          EqFrag = frSr(FirstRow, Col, LastRow).Address
        End If

      Else
        EqFrag = fsLegalName(BrkExtr, True, , False)
      End If

    ElseIf IndxType = peRatio Or IndxType = peEquation Or IndxType = peColumnHeader Then

      If OneCellRef Then

        If Col > 0 Then
          On Error Resume Next ' in case no samsht yet
          If RefToOtherSht Then Sheets(IIf(StdCalc, pscSamShtNa, pscStdShtNa)).Activate
          tmp = foAp.Clean(LCase(fsLegalName(Cells(flHeaderRow(pbStd), Col).Formula, False)))
          If RefToOtherSht Then Sheets(IIf(StdCalc, pscStdShtNa, pscSamShtNa)).Activate
          On Error GoTo 0
          SCref = False

          For i = 1 To puTask.iNeqns
            Na = fsLegalName(puTask.saEqnNames(i), False)
            If LCase(Na) = tmp Then

              If puTask.uaSwitches(i).SC Or (puTask.uaSwitches(i).Ar _
              And puTask.uaSwitches(i).ArrNrows = 1) Then
                SCref = True
                Exit For
              End If

            End If

          Next i

        End If

        tmp = ""

        If pbUPb And Col = 0 Then

          If (StdCalc And puTask.uaSwitches(Indx).SA) Or (Not StdCalc And puTask.uaSwitches(Indx).ST) Then

            If TestRef = "" Then
              tmp = IIf(StdCalc, "'SampleData'!", "'StandardData'!")
              Col = piaEqCol(Not StdCalc, Indx)
            End If

          End If

         End If

        If SCref Then ' reference to a single-cell equation
          EqFrag = tmp & Cells(FirstRow, Col - ErCol).Address(True)
        Else
          EqFrag = tmp & Cells(OutputRow, Col - ErCol).Address(False)
        End If

      Else
        EqFrag = frSr(FirstRow, Col - ErCol, LastRow).Address(False)
      End If

    End If

'    If And EqFrag = "" Then    ' ' 09/06/09 -- commented out -- the eqn could be simply a number
'      MsgBox "The bracketed expression " & Chr(34) & " " & BrkExtr & " " & Chr(34) & " in Task Equation" & _
'        StR(EqNum) & vbLf & "(saved as " & Chr(34) & "  " & TestRef & "  " & Chr(34) & ") " & vbLf & _
'        "does not refer to any defined range name," & vbLf & _
'        "equation, column-header, or ratio." & pscLF2 & _
'        "Aborting data-reduction.", , pscSq
'      CrashEnd
'    End If

    If RefToOtherSht Then
      EqFrag = IIf(StdCalc, "'" & pscStdShtNa & "'!", "'" & pscSamShtNa & "'!") & EqFrag
    End If

    If LeftSqBrkPos > 0 Then 'And Mid$(Eqn, LeftSqBrkPos + 1, 1) <> Chr(34) Then
      LeftPos = LeftSqBrkPos - 1

      If LeftBrk = LeftSqBrkPos Then
        RightPos = RightBrk + 2
      Else
        RightPos = RightSqBrkPos + 1
      End If

      EqnLeft = Left$(Eqn, LeftPos)
      EqnRight = Mid$(Eqn, RightPos)
      Eqn = EqnLeft & EqFrag & EqnRight 'EqnLeft & RefAddr & EqFrag & EqnRight
    End If

NextDo:

  Loop Until InStr(Eqn, "[") = 0

  Subst Eqn, "$$$", "[", "&&&", "]"

' Now place the output of the parsed formula in the correct cell.
' If .SC and .LA then must have both outputrow and outputcol selected
' If .SC and not .LA then place in cell just below the eqn name col-header
' If .AR and .LA then
' If .AR and not .LA then
' If .FO then place the parsed expression
' If .NU or not .FO then place the numerically evaluated expression
' If .SC then name the output cell as a range (same as col-header if not .LA)
  OutputNrows = IIf(.Ar, .ArrNrows, 1): OutputNcols = IIf(.Ar, .ArrNcols, 1)
  Set OutputCellRange = frSr(OutputRow, OutputCol, _
      OutputRow + OutputNrows - 1, OutputCol + OutputNcols - 1)

  If .SC Or .LA Or .Ar Then
    foAp.Calculate
    tmp = puTask.saEqnNames(EqNum)
    'CurlyExtract Phrase:=tmp, Extracted:="", RemoveFromPhrase:=True

    If .SC Or .Ar Then                      '  !!!!Check if ok for all scenarios
      With OutputCellRange(1, 1)
        Na = fsLegalName(tmp, True, , False)
        p = InStr(Na, "!")
        If p > 0 Then Na = Mid$(Na, p + 1)
        p = InStr(Na, "||")
        If p > 0 Then Na = Left$(Na, p - 1)
        .Name = Na
        With .Font:
          .Bold = True: .Color = RGB(0, 0, 128): .italic = True
        End With
      End With

      If OutputNrows > 1 Or OutputNcols > 1 Then  ' Name all cells of the output

        For i = 1 To .ArrNrows                    '  range by adding #rows#cols to
          For j = 1 To .ArrNcols                  '  the main (r1c1) name -e.g. ppmU12, ppmU13...
            OutputCellRange(i, j).Name = Na & fsS(i) & fsS(j)
        Next j, i

      End If

    End If

    If .LA Then
      Subst tmp, "|", " "

      If .SC Then
        Set OutputRangeNameCell = Cells(plHdrRw, OutputCol)
      Else
        Set OutputRangeNameCell = Cells(plHdrRw, OutputCol)
      End If

      If (.LA And OutputRangeNameCell = "") Then
        OutputRangeNameCell.Formula = tmp
      End If

    Else
      Set OutputRangeNameCell = frSr(plHdrRw, OutputCol)
      If Not .Ar Then Subst tmp, "|", vbLf
    End If

    If .Ar Then
      Fonts rw1:=OutputRangeNameCell, Bold:=True, italic:=True, Clr:=vbBlue
    Else
      If pbUPb Then tB = (piaSwapCols(EqNum) = 0) Else tB = True

      If tB Or OutputRow = plaFirstDatRw(-StdCalc) Then
        Fonts rw1:=OutputRangeNameCell, Formul:=tmp, Bold:=True, italic:=True, Clr:=vbBlue
      End If

    End If

  End If

  For Ref = 1 To Nrefs
    If WbkNa(Ref) <> "" Then
      Subst Eqn, "###" & fsS(Ref) & "###", "[" & WbkNa(Ref) & "]"
    End If
  Next Ref

  On Error Resume Next ' in case nonexistent range-name, etc
  p = InStr(LCase(Eqn), "sqmad(")
  If p > 0 Then Eqn = Left$(Eqn, p - 1) & "%%%%%%" & Mid$(Eqn, p + 6)
  p = InStr(LCase(Eqn), "mad(")
  If p > 0 Then Eqn = Left$(Eqn, p - 1) & "sqMAD(" & Mid$(Eqn, p + 4)
  Subst Eqn, "%%%%%%", "sqMAD("

  If .Ar Then
    On Error Resume Next
    With OutputCellRange

      If puTask.uaSwitches(EqNum).FO Then
        .FormulaArray = "=" & Eqn
      Else
        .FormulaArray = Evaluate(Eqn)
      End If

      On Error GoTo 0
      With .Font: .Color = RGB(0, 0, 128): .italic = True: End With
    End With
  ElseIf .FO Then
    OutputCellRange = "=" & Eqn
  Else
    OutputCellRange = Evaluate(Eqn)
  End If

  OutputCellRange.ColumnWidth = fvMax(6, OutputCellRange.ColumnWidth)
  On Error GoTo 0
End With
Exit Sub

NoName: Exit Sub
Msg = "Equation " & fsS(EqNum) & " refers either to a column header (" & _
  BrkExtr & ")" & vbLf & "whose data-column does not (yet?) exist, or to a" & _
  vbLf & "range name that does not and will not exist."
ComplainCrash Msg
Exit Sub

BadWbkRef: On Error GoTo 0
tmp = IIf(TestRef = " ", "", " (" & TestRef & ") ")
MsgBox "The workbook referenced by Equation" & StR(EqNum) & _
 tmp & "must be loaded to run this Task.", , pscSq
 ActiveSheet.Delete
End
BadWkShtRef: On Error GoTo 0
MsgBox "The worksheet referenced by Equation" & StR(EqNum) & _
 " (" & TestRef & ") does not seem to exist.", , pscSq
 ActiveSheet.Delete
End
End Sub

Function fiExtrRatnumFromRatLett%(ByVal ContainerStr$)
' return isotope-ratio# from a "[H]"-type string
Dim tmp%, Le%, i%, N%
tmp = 0
ContainerStr = LCase(Trim(ContainerStr))
Le = Len(ContainerStr)

If fbIsAllAlphaChars(ContainerStr) Then

  For i = 1 To Le
    N = Asc(Mid$(ContainerStr, i, 1)) - 96
    tmp = tmp + N * 26 ^ (Le - i)
  Next i

End If

fiExtrRatnumFromRatLett = tmp
End Function

Sub ExtractEqnRef(ByVal Phrase$, IndxStr$, IndxNum%, Optional IndxType, _
  Optional Isorats, Optional EquaNames, Optional InQuotes As Boolean = False, _
  Optional InstanceCt% = 1, Optional NoIsorats As Boolean = False, _
  Optional ExtractedPhrase, Optional fbIsErCol, Optional RefType = 0, _
  Optional WbkName, Optional WkShtName)

' ExtractedIndxNum=0 if nothing in brackets,
' -n if an isotope-ratio#,
' +n if an eqn#,
' 1000 + Colnum if an existing column-header,
' -1000-PrefsConstNum if a Preferences-defined constant,
' -2000-TaskConstNum if a Task-defined constant,
' -3000-PrefsConstNum if both.

' ApparentRefType= 1 if ratio/eq/const index, 2 if literal

' Isorats is list of Task's isotope ratios, EquaNames the eqn names
Dim ConstInTask As Boolean, ConstInPrefs As Boolean
Dim Ex$, s$, s0$, LeftBr$, RtBr$, LenBr%, s1$, s2$, s3$, s4$, Extr%, Repl$
Dim p%, q%, NumRats%, NumEqns%, LenEx%, RatIndx%, EqnIndx%, ConstIndx%, TconstIndx%

Extr = 0: Ex = "": RefType = 0: IndxType = 0
VIM WbkName, ""
VIM WkShtName, ""

fbIsErCol = False: ConstInTask = False: ConstInPrefs = False
s0 = Phrase
LeftBr = IIf(InQuotes, "[" & pscQ, "[")
RtBr = IIf(InQuotes, pscQ & "]", "]")
LenBr = 1 - InQuotes
p = fiInstanceLoc(s0, InstanceCt, LeftBr)

If p > 0 Then   ' if Phrase contains [" (inquotes=true) or [ (...=False)
  s = Mid$(s0, p): q = InStr(s, RtBr)
  If p > 0 And q > 0 Then
    Ex = Trim(Mid$(s, 1 + LenBr, q - LenBr - 1))  ' the bracketed string

    If p > 3 Then
      If Mid(s0, p - 2, 2) = "'!" Then
        s = StrReverse(Left(s0, p - 3))
        p = InStr(s, "'")
        WkShtName = StrReverse(Left(s, p - 1))
      End If
    End If

    If Left$(Ex, 1) = pscPm Then
      Ex = Mid$(Ex, 2)
      fbIsErCol = True
    End If

    LenEx = Len(Ex)

    If LenEx <= 2 Then             ' a 1 or 2 letter or number string?
      s1 = LCase(Trim(Left$(Ex, 1)))
      If LenEx = 2 Then s2 = LCase(Trim(Mid$(Ex, 2, 1)))

      If fbIsNumber(s1) And (LenEx = 1 Or fbIsNumber(s2)) Then
        Extr = Val(Ex)              ' an Equation index-number
        RefType = 1
        IndxType = peEquation
      ElseIf fbIsAllAlphaChars(s1 & s2) Then ' an isotope-ratio index-letter
        Extr = -fiExtrRatnumFromRatLett(s1 & s2)
        RefType = 1
        IndxType = peRatio
      End If

    Else                            ' A literal defined isotope ratio (eg "206/204")?
      Subst Ex, pscQ
      p = InStr(Ex, "/")

      If p > 0 Then
        If IsNumeric(Left$(Ex, p - 1)) Then
          If IsNumeric(Mid$(Ex, p + 1)) Then RefType = 2
        End If
      End If

      If fbNIM(Isorats) Then
        NumRats = UBound(Isorats)
        Subst Ex, pscQ

        For RatIndx = 1 To NumRats
          If Isorats(RatIndx) = Ex Then Exit For
        Next RatIndx

        If RatIndx <= NumRats Then
          Extr = -RatIndx                            ' yes, is a literal ratio
          RefType = 2
          IndxType = peRatio
        End If
      End If

      If Extr = 0 And fbNIM(EquaNames) Then
        s1 = LCase(fsStrip(Ex, , , , False, , True, , , True, , True))
        NumEqns = UBound(EquaNames)

        For EqnIndx = 1 To NumEqns            ' a defined-equation name?
          s4 = EquaNames(EqnIndx)
          If s4 <> "" Then
            s2 = LCase(fsLegalName(s4))
            If s1 = s2 Then Exit For
          End If
        Next EqnIndx

        If EqnIndx <= NumEqns Then
          Extr = EqnIndx                      ' yes, is a defined-eqn name
          RefType = 1
          IndxType = peEquation

        ElseIf plHdrRw > 0 Then
          If Cells(plHdrRw, 3) = "Hours" Then ' Sample or Std worksheet?
            EqnIndx = FindHeaderCol(s1)       ' An already-placed column-header string?

            If EqnIndx > 0 Then
              Extr = 1000 + EqnIndx
              IndxType = peColumnHeader
            End If

          End If

        End If

      End If

    End If

  End If

End If

If Extr = 0 Then
  p = fiInstanceLoc(s0, InstanceCt, "<")
  q = fiInstanceLoc(s0, InstanceCt, ">")
  'p = InStr(s0, "<"): q = InStr(s0, ">")

  If p > 0 And (q - p) > 1 Then
    Ex = Mid$(s0, p + 1, q - p - 1)
    s1 = fs_(Ex, True)

    If fbIsAllNumChars(s1) Then
      ConstIndx = Val(s1)
      s2 = Trim(fs_(prConstNames(ConstIndx)))

      If s2 = "" Then
        IndxStr = "": IndxNum = 0: Exit Sub
      Else
        ConstInPrefs = True: Ex = s1
        s1 = fs_(prConstNames(ConstIndx), True)
      End If

      RefType = 1
    End If

    If RefType <> 1 Then

      For ConstIndx = 1 To peMaxConsts
        s3 = fs_(prConstNames(ConstIndx), True)

        If s3 = "" Then
          Exit For
        ElseIf s1 = s3 Then
          ConstInPrefs = True
          RefType = 2
          Exit For
        End If

      Next ConstIndx

    End If

    For TconstIndx = 1 To puTask.iNconsts
      s3 = fs_(puTask.saConstNames(TconstIndx), True)

      If s1 = s3 Then
        ConstInTask = True
        RefType = 2
        Exit For
      End If

    Next TconstIndx

    If ConstInPrefs And Not ConstInTask Then
      Extr = -1000 - ConstIndx
      IndxType = pePrefsConstant
    ElseIf ConstInTask And Not ConstInPrefs Then
      Extr = -2000 - TconstIndx
      IndxType = peTaskConstant
    ElseIf ConstInPrefs And ConstInTask Then
      Extr = -3000 - ConstIndx
      IndxType = peBothConstant
    Else
      IndxType = peUndefinedConstant
    End If

  End If

End If

If fbNIM(ExtractedPhrase) And Ex <> "" Then
  Repl = ""
  ExtractedPhrase = Phrase

  If Extr < -1000 Or Extr = 0 Then
    Repl = "<" & Ex & ">"
  ElseIf Extr < 0 And fbNIM(Isorats) Then
    Repl = Isorats(-Extr)
  ElseIf InStr(Ex, "/") > 0 Then
    Repl = psBrQL & Ex & psBrQR
  ElseIf (Extr > 0 And Extr <= peMaxEqns) Or (Extr < 0 And Extr >= -peMaxRats) Then
    Repl = "[" & Ex & "]"
  End If

  If Repl <> "" Then Subst ExtractedPhrase, Repl
End If
IndxStr = Ex: IndxNum = Extr
End Sub

Function FindHeaderCol(ByVal EqnStr$, Optional ByVal SamAndStd As Boolean = False)
' Find equation number of an equation-string
Dim Stand As Boolean, s$, s1$, Col%, tmp%, Le%
Dim ShtIn As Worksheet

tmp = 0:     s = Trim(EqnStr)
Le = Len(s): Set ShtIn = ActiveSheet
Stand = (ActiveSheet.Name = pscStdShtNa)

If Le = 1 Or Le = 2 Then

  If fbIsNumChar(Left$(s, 1)) And fbIsNumChar(Mid$(s, Le, 1)) Then
    tmp = Val(s)
  End If

End If

If tmp = 0 Then
  Subst s, pscQ
  ' first look in active sheet
  s1 = fsStrip(s, , , , , , True, , , True) 'fsLegalName(s)
  HdrRow = flHeaderRow(Stand)
  FindStr Phrase:=s1, ColFound:=Col, RowLook1:=HdrRow, ColLook1:=1, LegalRangeNameOnly:=False, _
    RowLook2:=HdrRow, WholeWord:=True

  If SamAndStd And Col = 0 And pbUPb Then ' Then in inactive (either Std or sample) sheet

    If Not foUser("StdsOnly") Then
      On Error GoTo 1
      Sheets(IIf(Stand, pscSamShtNa, pscStdShtNa)).Activate
      HdrRow = flHeaderRow(Not Stand)
      FindStr Phrase:=s1, ColFound:=Col, RowLook1:=HdrRow, ColLook1:=2, LegalSheetNameOnly:=True, _
        RowLook2:=HdrRow, WholeWord:=True
1:    On Error GoTo 0
      ShtIn.Activate
      HdrRow = flHeaderRow(Stand)
    End If

  End If

  tmp = Col
End If

FindHeaderCol = tmp
End Function

Function fbIsErCol(ByVal ColNum%) As Boolean
Dim Ecol%, EcolHdr$
Ecol = 1 + ColNum
EcolHdr = Cells(plHdrRw, Ecol)

If InStr(EcolHdr, "+/-") = 0 And InStr(EcolHdr, "+-") = 0 And InStr(EcolHdr, pscPm) = 0 _
   And InStr(EcolHdr, "err") = 0 Then
  fbIsErCol = False
Else
  fbIsErCol = True
End If

End Function

Sub oldFindWbkShtRefs(ByVal Phrase$, Nrefs%, WbkNames$(), WkshtNames$(), _
  Optional BadWbk As Boolean = False, Optional BadSht As Boolean = False, _
  Optional StrippedPhrase = "")
' 09/06/09 -- Modified so >1 wbk/wksht ref can be extracted!
Dim WbkName$, ShtName$, StrPhr$, BothRef$, LcStrPhr$, s$, AscS%, Test$
Dim LeftBrak%, RightBrak%, ExclamPt%, Apost%, p%
' [WbkName.xls]WkshtName!$X$6  OR '[WbkName.xls]WkshtName'!$X$6
' Book12!ky
' 'Book12'!ky
' [book12]standarddata!ky
' '[book12]standarddata'!ky

Nrefs = 0
Phrase = Trim(LCase(Phrase))
Subst Phrase, "'"

Do
  WbkName = "": ShtName = ""
  StrPhr = Trim(Phrase)
  LcStrPhr = LCase(StrPhr)
  ExclamPt = InStr(StrPhr, "!")

If ExclamPt < 2 Then Exit Do
  Nrefs = 1 + Nrefs
  BothRef = ""
  ReDim Preserve WbkNames(Nrefs), WkshtNames(Nrefs)

  If Mid(LcStrPhr, ExclamPt - 1, 1) = "'" Then
    Apost = InStr(LcStrPhr, "'")
    If Apost > (ExclamPt - 2) Then BadSht = True: Exit Sub
    BothRef = Mid(LcStrPhr, Apost + 1, ExclamPt - Apost - 2)
    Subst LcStrPhr, "'" & BothRef & "'!"
  Else
    '  BadSht = True: Exit Sub  ' 10/04/20 commented out

    Test = Left(LcStrPhr, ExclamPt - 1)

    For p = Len(Test) To 1 Step -1
      s = Mid(Test, p, 1)
      AscS = Asc(s)

      If s <> "_" And (AscS < 97 Or AscS > 122) Then
        ShtName = Mid(Test, p + 1)
        Exit For
      End If

    Next p

    If p = 0 Then ShtName = Test
    Subst LcStrPhr, ShtName & "!"
    BothRef = Test

  End If

  LeftBrak = InStr(BothRef, "[")

  If LeftBrak > 0 Then
    RightBrak = InStr(BothRef, "]")
    If RightBrak = 0 Then BadWbk = True: Exit Sub
    WbkName = Mid(BothRef, 2, RightBrak - 2)
    If ShtName = "" Then
      ShtName = Mid(BothRef, 1 + RightBrak, ExclamPt - RightBrak - 1)
    End If
  ElseIf ShtName <> "" Then
    ShtName = Left(BothRef, ExclamPt - 1)
  End If

  WbkNames(Nrefs) = WbkName
  If ShtName = "" Then
    BadSht = True: BadWbk = True: Exit Sub
  Else
    WkshtNames(Nrefs) = ShtName
    Phrase = StrPhr
  End If

Loop

StrippedPhrase = Phrase
End Sub
Sub Test()
Dim s$, a$(), b$(), c$, N%
s = "Log(WtdMean)*'OtherSht1'!A1+'[OtherBk2]OtherSht2'!B2"
 FindWbkShtRefs s, N, a, b
End Sub

Sub FindWbkShtRefs(ByVal Phrase$, Nrefs%, WbkNames$(), WkshtNames$(), _
  Optional BadWbk As Boolean = False, Optional BadSht As Boolean = False) ', _
'  Optional StrippedPhrase = "")
' 09/06/09 -- Modified so >1 wbk/wksht ref can be extracted!
' 10/04/20 -- recoded for recognition of more ways of workbook/worksheet refs
Dim WbkName$, WkShtName$, s$, AscS%, Test$, LastS$, ss$
Dim LeftBrak%, RightBrak%, ExPt%, Apost%, p%, tmp$
' [WbkName.xls]WkshtName!$X$6  OR '[WbkName.xls]WkshtName'!$X$6
'    Book12!ky                  NOT Allowed
'    'Book12'!ky                NOT Allowed
'     [book12]standarddata!ky   OK
'    '[book12]standarddata'!ky  OK

Nrefs = 0
s = Trim(LCase(Phrase))          ' say  "Log(WtdMean)*'OtherSht'!A1+'[OtherBk]OtherSht'!B2"
Subst s, "'", , psBrQL, , psBrQR ' now  "Log(WtdMean)*OtherSht!A1+[OtherBk]OtherSht!B2"

Do
  WbkName = "": WkShtName = ""
  LastS = s
  ExPt = InStr(s, "!")
If ExPt < 2 Then Exit Do
  Nrefs = 1 + Nrefs
  LeftBrak = InStr(s, "[")

  ReDim Preserve WbkNames(Nrefs), WkshtNames(Nrefs)
  Test = Left(s, ExPt - 1)       ' say  "Log(WtdMean)*OtherSht"
  tmp = ""

  For p = Len(Test) To 1 Step -1
    ss = Mid(Test, p, 1)

    If Not fbIsLegalName(ss) Then
      tmp = Left(s, p)           ' say  "Log(WtdMean)*"
      WkShtName = Mid(Test, p + 1)
      Exit For
    End If


  Next p

  s = tmp & Mid(s, ExPt + 1)    ' say  "Log(WtdMean)*A1+[OtherBk]OtherSht!B2"

  If LeftBrak > 0 And LeftBrak < ExPt Then
    LeftBrak = InStr(s, "[")
    RightBrak = InStr(s, "]")
    If LeftBrak > RightBrak Or RightBrak = 0 Then BadWbk = True: Exit Sub

    WbkName = Mid(s, 1 + LeftBrak, RightBrak - LeftBrak - 1)
    s = Left(s, LeftBrak - 1) & Mid(s, RightBrak + 1)
  End If

If LastS = s Then Exit Do

  WbkNames(Nrefs) = WbkName
  WkshtNames(Nrefs) = WkShtName

Loop

End Sub

Function fsInQ$(ByVal Phrase$, Optional AsLowercase As Boolean = False)
Dim s$
s = pscQ & Phrase & pscQ
If AsLowercase Then s = LCase(s)
fsInQ = s
End Function

Function fs_(ByVal Phrase$, Optional AsLowercase As Boolean = False)
Dim s$
s = Phrase

If s = "" Or s = "_" Then
  s = ""
Else
  If Left$(s, 1) = "_" Then s = Mid$(s, 2)
  If AsLowercase Then s = LCase(s)
End If

fs_ = s
End Function

Sub GetCondensedShtInfo(Optional CondensedSht)
Dim i%, Co%, Nsp%, Rw&, CondSht As Worksheet

If IsMissing(CondensedSht) Then
  Set CondSht = ActiveSheet
Else
  Set CondSht = CondensedSht
End If


Set phCondensedSht = CondSht
Nsp = Val(Cells(4, picDatCol))
piNumAllSpots = Nsp
pbXMLfile = (Left(Cells(1, picRawfiletypeCol), 3) = "XML")
pbShortCondensed = (Right(Cells(1, picRawfileFirstcolCol), 5) = "short")

ReDim plaSpotNameRowsCond(1 To Nsp), psaSpotNames(1 To Nsp)
ReDim piaFileNscans(1 To Nsp), psaSpotDateTime(1 To Nsp)

For i = 1 To Nsp
  Rw = Cells(i + 1, peReadyCol)
  plaSpotNameRowsCond(i) = Rw
  GetNameDatePksScans Rw, psaSpotNames(i), psaSpotDateTime(i), piFileNpks, piaFileNscans(i)
Next i

ReDim pdaFileMass(1 To piFileNpks)
Rw = plaSpotNameRowsCond(1)

For i = 1 To piFileNpks
  Co = picDatCol + (i - 1) * 5 + 1
  pdaFileMass(i) = Cells(Rw + picDatRowOffs - 1, Co)
Next i

CalcNominalMassVal True
End Sub

Sub GetNameDatePksScans(ByVal Rw&, Optional SpotName$, Optional SpotDateTime$, _
 Optional SpotNpeaks%, Optional SpotNscans%, Optional SpotDateOnly$, _
 Optional SpotTimeOnly$)
Dim GotSht As Boolean, ND$, PS$, DateTime$, p%

FindCondensedSheet GotSht, True
If Not GotSht Then
  MsgBox "Can't find raw-data worksheet in active workbook.", , pscSq
  End
End If

ND = Cells(Rw, picNameDateCol)
PS = Cells(Rw, picPksScansCol)
RemoveDblChars ND
p = InStr(ND, ", 2")

If fbNIM(SpotName) Then
  SpotName = Trim(Left$(ND, p - 1))
End If

If fbNIM(SpotDateTime) Or fbNIM(SpotDateOnly) Or fbNIM(SpotTimeOnly) Then
  DateTime = Trim(Mid$(ND, p + 1))
  If fbNIM(SpotDateTime) Then SpotDateTime = Trim(Mid$(ND, p + 1))
  If fbNIM(SpotDateOnly) Then SpotDateOnly = fsExtrLeft(DateTime, " ")
  If fbNIM(SpotTimeOnly) Then SpotTimeOnly = fsExtrRight(DateTime, " ", True)
End If

If fbNIM(SpotNpeaks) Then
  SpotNpeaks = Val(fsExtrLeft(PS, ","))
End If

If fbNIM(SpotNscans) Then
  SpotNscans = Val(fsExtrRight(PS, ","))
End If
End Sub
