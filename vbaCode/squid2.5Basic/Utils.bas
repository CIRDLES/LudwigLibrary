Attribute VB_Name = "Utils"
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
'' Module UTILS for SQUID
Option Explicit
Option Base 1

Sub NoUpdate(Optional DontUp As Boolean = True)
With Application

  If DontUp Then
    .ScreenUpdating = False
  Else
    .ScreenUpdating = True
  End If

End With
End Sub

Sub ManCalc(Optional Manual As Boolean = True)
On Error Resume Next
foAp.Calculation = IIf(Manual, xlCalculationManual, xlCalculationAutomatic)
End Sub

Sub TickFor() ' Find shortest number-format for Y-axis tick labels

Dim s$, d%, L%, v#, MaxD#, t#, MinS#, MaxS#

With ActiveChart.Axes(xlValue)
  MinS = .MinimumScale: MaxS = .MaximumScale
  t = .MajorUnit: MaxD = 0

  If MinS = MaxS Then
    MsgBox "SQUID error in Sub TickFor during axis-scaling.", , pscSq
    CrashEnd
  End If

  v = Drnd(MinS, 6)

  Do Until v > MaxS
    s$ = fsS(v): L = Len(s$)
    d = InStr(s$, ".")
    If d > 0 Then MaxD = fvMax(MaxD, L - d)
    v = Drnd(v + t, 6)
  Loop

  If MaxD = 0 Then
    s$ = pscZq
  Else
    s$ = "0." & String(MaxD, pscZq)
  End If

  .TickLabels.NumberFormat = s$
End With
End Sub

Sub AddStrikeThru()
' If the StrikeThrough button isn't already on the formatting toolbar, Add it.
Dim Exists As Boolean, i%, b As Object, c As Object

Set b = foAp.CommandBars("Formatting").Controls
Exists = False

For Each c In b
  If c.ID = 290 Then Exists = True: Exit For  ' Strikethru
Next c

If Not Exists Then b.Add Type:=msoControlButton, ID:=290, Before:=6
Exists = False
Set b = foAp.CommandBars("Standard").Controls

For Each c In b
  If c.ID = 960 Then Exists = True: Exit For ' Calculate Now
Next c

If Not Exists Then b.Add Type:=msoControlButton, ID:=960, Before:=13
ClearObj b, c
End Sub

Sub Alerts(ByVal Yes As Boolean)
foAp.DisplayAlerts = Yes
End Sub

Function fsOpSys$()
fsOpSys = foAp.OperatingSystem
End Function

Function foLastOb(Ob As Object) As Object
Set foLastOb = Ob(Ob.Count)
End Function

Sub SetLastChart(Cht As ChartObject, NoChart As Boolean)
' Set Cht to be last-created chart if a chart exists
Dim ct%
ct = ActiveSheet.ChartObjects.Count

If ct Then
  Set Cht = ActiveSheet.ChartObjects(ct): NoChart = False
Else
  NoChart = True
End If

End Sub

Function fhSquidSht() As Worksheet
Set fhSquidSht = ThisWorkbook.Sheets("Squid")
End Function

Sub SortStr(Unsorted$(), Sorted$(), Indx, Optional Descending As Boolean = False)
Dim Lb%, i&, N&, j&

N = UBound(Unsorted)
Lb = LBound(Unsorted)

ReDim Preserve Sorted(Lb To N), Indx(Lb To N)

For i = Lb To N
  Sorted(i) = Unsorted(i)
Next i

BubbleSort Sorted, Indx, Descending
End Sub

Sub DupeSpotNames(SpotNamesIn$())
' NOT USED
Dim NewName$, OldName$, FileLine$, SortedNames$()
Dim c%, N%
Dim Indx&(), i&, r&, j&

N = UBound(SpotNamesIn)
SortStr SpotNamesIn, SortedNames, Indx

For i = 2 To N

  If SortedNames(i) = SortedNames(i - 1) Then
    j = Indx(i)
    OldName = SortedNames(i)
    NewName = OldName & "_A"
    r = plaSpotNameRowsCond(j)
    FileLine = Cells(r, 1).Text
    Subst FileLine, OldName, NewName
    Cells(r, 1) = FileLine
    SpotNamesIn(j) = NewName
  End If

Next i

End Sub

Sub RefreshSampleNames()

Dim i&, sNames As Range, Nspots%, tB As Boolean, Sht As Worksheet

Set sNames = foUser("SpotNames")
sNames.Clear
tB = True
Set Sht = ActiveSheet

On Error Resume Next
tB = (psaSpotNames(1) = "")
On Error GoTo 0

If tB Or piNumAllSpots = 0 Then FindCondensedSheet tB

On Error Resume Next
Sht.Activate
On Error GoTo 0

For i = 1 To piNumAllSpots
  sNames(i) = psaSpotNames(i)
Next i

With sNames
  .NumberFormat = "@"
  .HorizontalAlignment = xlCenter
End With

sNames.Sort Key1:=sNames(1)
Set sNames = Nothing

End Sub

Sub GetNameList(Optional NumInGroups, Optional NoGroups As Boolean, _
  Optional Pregroup As Boolean = False, Optional NotForGrouping As Boolean = False, _
  Optional FixedNchars As Boolean = False)


' 09/03/01 -- Add Nss and LargestGrp vars, force the group-finding algorithm to keep reducing
'              #chars to group be until at least one group of 2 or more spots is found
'             (if none found, NoGroups is returned as TRUE).
' 09/03/14 -- Add the "NotForGrouping" optional parameter.  If NotForGrouping (ie called from
'              GeochronSetup or GenIsoSetup), ignore case and spaces but nothing else.
'              Repair crept-in errors in algorithm between lines 1-2
' 10/11/21 -- Added the optional FixedNchars variable

' Put trimmed-sample names in User sheet
Dim s$, s0$, LastS$, ttn$()
Dim i%, j%, Nc%, Nss%, c%, r%, LargestGrp%, si%(), tnu%()
Dim ns&
Dim sNames As Range, Tnames As Range, WbkIn As Workbook

If pbFromSetup Then
  Nc = piNgChars
Else
  Nc = piNsChars ' 09/08/06
End If

Nc = Nc + 1    ' Nc is the number of leftmost-characters in the sample names
LastS = ""

Set WbkIn = ActiveWorkbook

With foUser.[SpotNames]
  r = .Row
  c = .Column
End With

If piNumAllSpots = 0 Then 'QQ
  ns = Val(Cells(4, picDatCol))
  piNumAllSpots = ns
Else
  ns = piNumAllSpots
End If

foUser.Activate
Set sNames = frSr(r, c, flEndRow(c))
Set Tnames = foUser("trimmedspotnames")

If Nc = 0 Then Nc = 3

' Starting with piNgchars -1 (or pbNschars - 1), shorten the clipped spot names
'  until the largest group of clipped names is 2 or more

If pbUPb Then ' 03/16/09 -- added

  Do
    LargestGrp = 0
    Nc = Nc - 1

    With Tnames
      frSr(.Row, .Column, fvMax(.Row, flEndRow(.Column))).ClearContents
    End With

    Set Tnames = foUser(frSr(Tnames.Row, Tnames.Column, r + piNumAllSpots - 1).Address)
    Tnames.Clear
    sNames.NumberFormat = "@": Tnames.NumberFormat = "@"
    Tnames.HorizontalAlignment = xlCenter
    sNames.Sort Key1:=sNames.Cells(1, 1), MatchCase:=False
    piNshortList = 0

    ReDim NumInGroups(1 To 1)

1
    For i = 1 To piNumAllSpots

      If NotForGrouping Then  ' mod 3/14/09
        s0 = fsStrip(sNames(i).Text, True, True)
      Else
        s0 = fsStrip(sNames(i).Text, pbIgCase, pbIgSpaces, pbIgDashes, _
           pbIgSlashes, PbIgCommas, pbIgColons, pbIgSemicolons, pbIgPeriods)
      End If

      s = LCase(Left$(s0, Nc)) ' 09/08/06 added Lcase
      c = (piNshortList = 0)

      If s <> "" Then

        If c Or LastS <> s Then
          piNshortList = 1 + piNshortList
          Tnames(piNshortList) = s

          ReDim Preserve NumInGroups(1 To piNshortList)

          NumInGroups(piNshortList) = 1
        Else
          NumInGroups(piNshortList) = 1 + NumInGroups(piNshortList)
        End If

        LargestGrp = fvMax(LargestGrp, NumInGroups(piNshortList))

      End If

      LastS = s
    Next i

2
     If LargestGrp > 1 Then

      If Pregroup Then

        ReDim si(1 To piNshortList), ttn(1 To piNshortList), tnu(1 To piNshortList)

        For i = 1 To piNshortList
          ttn(i) = Tnames(i)
        Next i

        For i = 1 To piNshortList
          tnu(i) = NumInGroups(i)
        Next i

        BubbleSort tnu, si(), True
        Nss = piNshortList

        For i = 1 To piNshortList
          r = si(i)
          c = tnu(i)

          If c < peMinNumInGroup Then
            Nss = i - 1
            Exit For
          End If

          Tnames(i) = ttn(r)
          NumInGroups(i) = tnu(i)
        Next i

        For i = Nss + 1 To piNshortList
          Tnames(i) = ""
        Next i

        piNshortList = Nss
        If Nss > 0 Then ReDim NumInGroups(1 To Nss)
      End If

    End If

  Loop Until LargestGrp > 1 Or Nc = 1 Or FixedNchars ' 10/11/21 -- added "Or FixedNchars"

End If

NoGroups = (LargestGrp = 1 And Nc = 1)

If NoGroups Then
  MsgBox "Cannot group these spots by name.", , pscSq
ElseIf pbFromSetup Then
  piNgChars = Nc
Else
  piNsChars = Nc
End If

foUser(IIf(pbFromSetup, "NgChars", "Nschars")) = Nc ' 09/08/06
ClearObj sNames, Tnames
On Error Resume Next
WbkIn.Activate
End Sub

Function fdDeadTimeCorrCPS#(ByVal MeasuredCPS As Variant, Optional Deadtime)
Dim Denom#
DefVal Deadtime, pdDeadTimeSecs
If Deadtime = 0 Then fdDeadTimeCorrCPS = MeasuredCPS: Exit Function
Denom = MeasuredCPS * Deadtime

If Denom <= 0 Then
  fdDeadTimeCorrCPS = 0
Else
  fdDeadTimeCorrCPS = MeasuredCPS / (1 - MeasuredCPS * Deadtime)
End If

End Function

Sub CrashEnd(Optional LeaveWindowOpen As Boolean = False, Optional Msg$ = "")
If Msg <> "" Then MsgBox "Squid error " & Msg, , pscSq
On Error Resume Next
Alerts False
StatBar
phCondensedSht.Activate
End
End Sub

Sub ComplainCrash(ByVal Complaint$)
'CrashNoise
MsgBox Complaint, vbInformation, pscSq
CrashEnd
End Sub

Sub DefVal(a As Variant, DefaultVal As Variant)
If fbIM(a) Then a = DefaultVal
End Sub

Sub CopySbox(ByVal BoxName$, ByVal FirstRow&, ByVal FirstCol%, piLastCol%)

Dim Lrow&, b As Range

If FirstCol = 0 Then CrashEnd , "in Sub CopySbox -- FirstCol passed as zero."

Set b = foUser(BoxName)
b.Copy Cells(FirstRow, FirstCol)
piLastCol = FirstCol + b.Columns.Count - 1
Lrow = FirstRow + b.Rows.Count - 1
frSr(FirstRow, FirstCol, Lrow, piLastCol).Name = BoxName
Set b = Nothing
End Sub

Function fdPrnd#(ByVal Number, ByVal Power As Integer) ' Return a # rounded to the
fdPrnd = foAp.Round(Number, -Power)                    '  specified power-of-10
End Function

Sub GetConcStdData() ' Get concentration Std info

Dim IgnoredChangedRuntable As Boolean
Dim TD$, p1%, p2%, i&, Sp&
Dim ConcStdConstSum#, temp1#, temp2#

ConcStdConstSum = 0: piParentEleStdN = 0: IgnoredChangedRuntable = False

For i = 1 To piNumConcStdSpots
  Sp = piaConcStdSpots(i)
  ParseRawData Sp, False, IgnoredChangedRuntable, TD, False, False

  If Not IgnoredChangedRuntable Then
    piSpotOutputCol = 0
    EqnInterp puTask.saEqns(-4), -4, temp1, temp2, 1, p1, True

    If temp1 <> 0 And temp1 <> pdcErrVal Then
      piParentEleStdN = 1 + piParentEleStdN
      StatBar "Getting " & psaPDele(piU1Th2) & "-pbStd data ", piParentEleStdN
      ConcStdConstSum = ConcStdConstSum + temp1
    End If

  End If

Next i

If piParentEleStdN > 0 Then
  pdMeanParentEleA = ConcStdConstSum / piParentEleStdN
End If
End Sub

Sub ExtractRowDblDbl(Arry#(), Vect#(), ByVal Row&, Optional ZerDim As Boolean = False)
Dim i%, L%, u%
L = LBound(Arry, 2): u = UBound(Arry, 2)

If ZerDim Then
  ReDim Vect(L To u)

  For i = L To u
    Vect(i) = Arry(Row, i)
  Next i

Else
  ReDim Vect(L To u, 1)

  For i = L To u
    Vect(i, 1) = Arry(Row, i)
  Next i

End If

End Sub

Sub ExtractColDblDbl(Arry#(), Vect#(), ByVal Col%, Optional ZerDim As Boolean = False)
Dim i%, L%, u%
L = LBound(Arry, 2): u = UBound(Arry, 2)

If ZerDim Then
  ReDim Vect(L To u)

  For i = L To u
    Vect(i) = Arry(i, Col)
  Next i

Else
  ReDim Vect(L To u, 1)

  For i = L To u
    Vect(i, 1) = Arry(i, Col)
  Next i

End If

End Sub

Sub SetArrayVal(SetValue, Arr())

Dim Dim2 As Boolean
Dim i&, j&, L1&, U1&, L2&, U2&

L1 = LBound(Arr, 1)
U1 = UBound(Arr, 1)

On Error GoTo 1
L2 = LBound(Arr, 2)
U2 = UBound(Arr, 2)
Dim2 = True
1: On Error GoTo 0

If Dim2 Then

  For i = L1 To U1
    For j = L2 To U2
      Arr(i, j) = SetValue
  Next j, i

Else

  For i = L1 To U1
    Arr(i) = SetValue
  Next i

End If

End Sub

Sub FormatAge(ByVal AgeCol%, ByVal FontClr&)

If AgeCol > 0 Then
  Fonts , AgeCol, , 1 + AgeCol, FontClr, True
  SymbChar plHdrRw, 1 + AgeCol, 2
End If

End Sub

Sub TrackSBM(ByVal Std As Boolean, UseSBM, BadSbm%())

Dim i%, j%, ii%, jj%, Cts#, s1#, s2#, MeanSBM#

With puTask
  ReDim NormSbm(1 To .iNpeaks)

  For i = 1 To .iNpeaks
    s2 = 0

    For j = 1 To piNscans
      Cts = pdaSBMcps(i, j) '* pdaIntT(i)

      If Cts <= 0 Then
        BadSbm(IIf(pbUPb, -Std, 1)) = 1 + BadSbm(IIf(pbUPb, -Std, 1))
        UseSBM = False: Exit Sub
      End If

      s2 = s2 + Cts
    Next j

    NormSbm(i) = s2 / piNscans ' Mean SBM at each mass station
  Next i

  MeanSBM = foAp.Average(NormSbm)

  For i = 1 To .iNpeaks
    ' Mean %deviation at each mass station from mean SBM
    pdaSbmDeltaPcnt(i, piSpotNum) = 100 * (NormSbm(i) - MeanSBM) / MeanSBM
  Next i

End With
End Sub

Function fbNotDbug() As Boolean ' if debugging return TRUE
fbNotDbug = True 'False ' Not (Right$(LCase(ThisWorkbook.Name), 4) = ".xls")
End Function

Sub ShowStatusBar(Optional Yes = True) ' Enable Excel's Status Bar?
foAp.DisplayStatusBar = Yes
End Sub

Sub Blink()    ' Force screen update, then inhibit
NoUpdate False
NoUpdate
End Sub

Function fnRight(b As Object) As Single
fnRight = b.Left + b.Width
End Function

Function fnBottom(b As Object) As Single
fnBottom = b.Top + b.Height
End Function

Sub FillList(c As Object, Source, Optional ListIndx, _
  Optional ByVal IndxAsIndex As Boolean, Optional ByVal LastItem, _
  Optional ByVal NoAct As Boolean = False, Optional IsOK)
' Populate the list of a ComboBox control with a single-column range

Dim FoundIndx As Boolean, i%, LastInd%, tB As Boolean

DefVal IndxAsIndex, 0
DefVal ListIndx, 0
c.RowSource = ""
c.Clear
On Error GoTo 10
i = Source.Count
GoTo 11
10: i = UBound(Source, 1)
11: On Error GoTo 0
LastInd = IIf(fbIM(LastItem), i, LastItem)

For i = 1 To LastInd

  If IsMissing(IsOK) Then
    tB = True
  Else
    tB = IsOK(i)
  End If

  If tB Then
    c.AddItem Source(i)

    If Not IndxAsIndex Then

      If Not FoundIndx And Source(i) = ListIndx Then
        On Error GoTo 1
        c.ListIndex = i: FoundIndx = True
        On Error Resume Next
      End If

    End If
  End If

Next i

1:
If IndxAsIndex Then
  c.ListIndex = ListIndx
ElseIf Not FoundIndx Then
  c.ListIndex = 0
End If
End Sub

Sub GetfsOpSys()
psExcelVersion = foAp.Version
piIzoom = 75
psStdFont = "Arial"
piNshtsIn = foAp.SheetsInNewWorkbook
psTwbName = "'" & ThisWorkbook.Name & "'!"
End Sub

Sub CopyDVect(VectorIn(), VectorOut#())
Dim i%, u%, L%
L = LBound(VectorIn): u = UBound(VectorIn)
ReDim Preserve VectorOut(L To u)
For i = L To u: VectorOut(i) = VectorIn(i): Next
End Sub

Sub CopyIVect(VectorIn(), VectorOut%())
Dim i%, u%, L%
L = LBound(VectorIn): u = UBound(VectorIn)
ReDim Preserve VectorOut(L To u)
For i = L To u: VectorOut(i) = VectorIn(i): Next
End Sub

Sub CopySVect(VectorIn(), VectorOut$())
Dim i%, u%, L%
L = LBound(VectorIn): u = UBound(VectorIn)
ReDim Preserve VectorOut(L To u)
For i = L To u: VectorOut(i) = VectorIn(i): Next
End Sub

Sub OpenPreferences()
GetInfo
Preferences.Show
End Sub

Public Function flFindRow&(ByVal StartRow, Optional WhNotTrue_UntlIsFalse As Boolean = True, _
  Optional Txt = "", Optional Col = 1, Optional StartChar = 1, Optional StringLen = 0, _
  Optional CaseSensitive = False, Optional DelBlankRows = True, Optional NoRowLim = False)
  ' If WhileNotTrue_UntilIsFalse=True, then find first succeeding row that does
  '  NOT match Trim(Mid$(Txt,StartChar,StringLen)).
  ' If WhileNotTrue_UntilIsFalse=False, then find first succeeding row that DOES
  '  match Trim(Mid$(Txt,StartChar,StringLen)).
  ' Starts looking 1 row BELOW Startrow

Dim Got As Boolean, tB As Boolean, s$, RwIn&, r&, NumBlnkRows&

RwIn = StartRow: r = RwIn: NumBlnkRows = 0
If Not CaseSensitive Then Txt = LCase(Txt)

Do
  s = Cells(r + 1, Col)

  If Trim(s) = "" And DelBlankRows Then
    On Error GoTo 1 ' won't work if merged
    Rows(r + 1).Delete
    NumBlnkRows = 1 + NumBlnkRows
    If NumBlnkRows > 20 Then r = 0: Exit Do
1:  On Error GoTo 0

  Else
    If Not CaseSensitive Then s = LCase(s)
    If StringLen = 0 Then StringLen = 99
    s = Trim(Mid$(s, StartChar, StringLen))
    tB = (s = Txt)
    r = r + 1
    If WhNotTrue_UntlIsFalse Xor tB Then Exit Do
  End If

Loop Until r > 65534

If r > 65534 Then r = 0
flFindRow = r
End Function

Sub CreateRatioDatSheet()
Dim Sht As Worksheet
Sheets.Add
Set phRatSht = ActiveSheet
Alerts False

For Each Sht In ActiveWorkbook.Worksheets

  If LCase(Sht.Name) = "within-spot ratios" Then
    Sht.Delete: Exit For
  End If

Next Sht

phRatSht.Name = "Within-Spot Ratios"
NoGridlines
HA xlRight, 2, 2, pemaxrow, peMaxCol
ActiveWindow.Zoom = 85
phCondensedSht.Activate
End Sub

Sub PlaceRats(SpotNa$, ByVal SpotNum%, ByVal Rat1Eqn2%, ByVal RatEqNum%, _
               RatEqTime#(), RatEqVal#(), RatEqFerr#())
Dim s$
Dim c%, i%, p%, RightCol%, RatEqCol%
Dim Rw&, SpotR&
Dim Sh As Worksheet, ShtIn As Worksheet

If Not pbRatioDat Then Exit Sub
Set ShtIn = ActiveSheet
phRatSht.Activate

SpotR = flEndRow(1) - 2

With puTask

  If .bIsUPb And Rat1Eqn2 = 2 And RatEqNum < 0 Then ' a U-Pb "Special" equation

    FindStr SpotNa, SpotR, , , , 32767, 1, True, True, True, True, True, True, True, True, True, True
    ' 09/10/08 -- added the "WholeWord=True" in line above.

    If SpotR = 0 Or Cells(SpotR + 2, 1) <> "Spot#" & StR(SpotNum) Then
      ShtIn.Activate  ' 09/10/08 -- added
      Exit Sub
    End If

    RatEqCol = fiEndCol(SpotR) + 1

  ElseIf SpotR <= 0 Then                ' empty worksheet
    SpotR = 3
    RatEqCol = 2
    p = .iNrats
    If .bIsUPb Then p = p + 1 - .bDirectAltPD

    For i = 1 To .iNeqns
      With .uaSwitches(i)
        If .Nu And Not .SC And Not .LA Then p = p + 1
      End With
    Next i

    RightCol = fvMin(peMaxCol, 2 + 3 * p)

    For c = 2 To RightCol Step 6
      frSr(, c, , c + 2).Font.Color = 8388608
      frSr(, c + 3, , c + 5).Font.Color = 128
    Next c

    HA xlRight, , 1
    HA xlLeft, 1, 1
    With Cells(1, 1)
      .Value = "Within-spot ratio data"
      .Font.Size = 14
    End With

  ElseIf Cells(SpotR, 1) <> SpotNa Then ' find row with spot name

    For i = 2 To fiEndCol(SpotR)
      SpotR = fvMax(SpotR, flEndRow(i))
    Next i

    SpotR = SpotR + 2
    RatEqCol = 2
  Else

    RatEqCol = fiEndCol(SpotR) + 1
  End If

  If Cells(SpotR, 1) = "" Then           ' no spot name yet
    Cells(SpotR, 1) = SpotNa
    Cells(SpotR + 1, 1) = psaSpotDateTime(SpotNum)
    Cells(SpotR + 2, 1) = "Spot#" & StR(SpotNum)
    Rows(SpotR).Font.Bold = True
    Cells(SpotR + 1, 1).Font.Size = 8
    Cells(SpotR + 2, 1).Font.Bold = True
  End If

  For i = 1 To UBound(RatEqTime)
    Cells(SpotR + i, RatEqCol) = RatEqTime(i)
  Next i

  Cells(SpotR, RatEqCol) = "Time"
  RatEqCol = RatEqCol + 1
  Rows(SpotR).WrapText = True

  If Rat1Eqn2 = 2 Then
    If .bIsUPb And RatEqNum < 0 Then

      Select Case RatEqNum
        Case -1
          s = psaUThPbConstColNames(pbStd, 1)
        Case -2
          s = psaUThPbConstColNames(pbStd, 2)
        Case -3
          s = "232Th|/238U"
        Case -4
          s = "ppm|U"
      End Select

      Subst s, "|", " "
      Cells(SpotR, RatEqCol) = s 'fsVertToLF(s)
    Else
      'Subst .saEqnNames(NumPkOrd), "|", vbLf, "/", " /"
      Cells(SpotR, RatEqCol) = .saEqnNames(RatEqNum)
    End If
  Else

    If Rat1Eqn2 = 1 Then
      s = .saIsoRats(RatEqNum)
    Else
      s = .saEqns(RatEqNum)
    End If

    Cells(SpotR, RatEqCol) = s
  End If

  Cells(SpotR, 1 + RatEqCol) = "err"
  Rw = SpotR
  BorderLine 4, 2, SpotR, RatEqCol - 1, , RatEqCol + 1

  For i = 1 To UBound(RatEqVal)
    Rw = Rw + 1
    Cells(Rw, RatEqCol) = RatEqVal(i)
    Cells(Rw, RatEqCol + 1) = Abs(RatEqFerr(i) * RatEqVal(i))
  Next i

End With

ShtIn.Activate
ClearObj Sh, ShtIn
End Sub

Function fbInBox(ByVal PtLeft, ByVal PtTop, Boxx As Shape) As Boolean
fbInBox = (PtLeft > Boxx.Left And PtLeft < fnRight(Boxx) And _
    PtTop > fnBottom(Boxx) And PtTop < Boxx.Top)
End Function

Function fbOverlapAny(Boxx As Object)
Dim bBox!, rBox!, Shp As Object

bBox = fnBottom(Boxx): rBox = fnRight(Boxx)

For Each Shp In ActiveSheet
  With Boxx
    If fbInBox(.Left, bBox, Shp) Or fbInBox(.Left, .Top, Shp) Or _
       fbInBox(rBox, .Top, Shp) Or fbInBox(rBox, bBox, Shp) Then
      fbOverlapAny = True:  Exit Function
    End If
  End With
Next Shp

fbOverlapAny = False
Set Shp = Nothing
End Function

Sub ConvertIndexBracketed(EqnIn$, EqnOut$)
' Replace nonstd quote-symbols,[�X] with (�X)
' 10/04/21 modified to use wbk/wkshet refs in the [WbkName]WkshtName!A1 format
'   rather than '[WbkName]WkshtName'!A1
Dim ErCol As Boolean, BadWbk As Boolean, BadWksht As Boolean, ShtRefOnly() As Boolean
Dim WbkStr$, ExtrStr$, NoBrakExtr$, PmCleaned$, Msg$, WbkName$(), ShtName$()
Dim BrL%, BrR%, sqL%, sqR%, ExclPt%, Ref%, p%, Le%, Num%, LoopCt%, Nrefs%
Dim s$, AscS%, tmp$
' NOTE: Chr(34) = ", Chr(147)= �, Chr(148)= �
Subst EqnIn, Chr(147), pscQ, Chr(148), pscQ

EqnOut = EqnIn
FindWbkShtRefs EqnOut, Nrefs, WbkName, ShtName, BadWbk, BadWksht ', WbkStr

If Nrefs > 0 Then ' Temporarily replace workbook/worksheet refs with ### or @@@
  ReDim ShtRefOnly(Nrefs)
  '  Say start with                  Log(WtdMean)*'OtherSht'!A1+'[OtherBk]OtherSht'!B2
  Subst EqnOut, "'"        ' now is  Log(WtdMean)*OtherSht!A1+[OtherBk]OtherSht!B2

  For Ref = 1 To Nrefs
    ExclPt = InStr(EqnOut, "!")
    ExclPt = InStr(EqnOut, "!")
    tmp = ""

    For p = ExclPt - 1 To 1 Step -1 ' Find 1st char of the wksht name
      s = Mid(EqnOut, p, 1)
      AscS = Asc(s)

      If fbIsLegalName(s) Then
        tmp = Left(EqnOut, p + 2)
        Exit For
      End If

    Next p

    EqnOut = tmp & "###" & Mid(EqnOut, ExclPt + 1)
      ' So now is  Log(WtdMean)*###A1+[OtherBk]OtherSht!B2  (first Ref)
      ShtRefOnly(Ref) = True

    If BrL > 0 Then  ' Deal with 1st WkBk ref
      ' Say start with Log(WtdMean)*[OtherBk]OtherSht!B2+OtherSht!A1
      ExclPt = InStr(EqnOut, "!")
      BrR = InStr(EqnOut, "]")
      EqnOut = Left$(EqnOut, BrL - 1) & "@@@" & Mid$(EqnOut, BrR + 1)
      ShtRefOnly(Ref) = False
      ' So now is Log(WtdMean)*@@@B2+OtherSht!A1            (first Ref)
    End If

  Next Ref
End If

' Replace [" and "] with $$$ and &&&
Subst EqnOut, psBrQL, "$$$", psBrQR, "&&&"
LoopCt = 0

Do ' Deal with error-column refs, ie [�Ref]

  LoopCt = 1 + LoopCt
  ErCol = False
  BrL = InStr(EqnOut, "[")
  BrR = InStr(EqnOut, "]")
  If BrL = 0 Or BrR = 0 Then Exit Do
  ExtrStr = Mid$(EqnOut, BrL, BrR - BrL + 1)
  NoBrakExtr = Mid$(ExtrStr, 2, Len(ExtrStr) - 2)
  PmCleaned = NoBrakExtr
  Subst PmCleaned, pscPm
  If Len(PmCleaned) < Len(NoBrakExtr) Then ErCol = True

  If Len(PmCleaned) < 3 Then
    With puTask

      If fbIsAllAlphaChars(PmCleaned) Then
        Num = fiLettToNum(PmCleaned)

        If Num > 0 And Num <= .iNrats Then
          PmCleaned = .saIsoRats(Num)
        Else
          PmCleaned = "???"
        End If

      ElseIf fbIsAllNumChars(PmCleaned) Then
        Num = Val(PmCleaned)

        If Num > 0 And Num <= .iNeqns Then
          PmCleaned = .saEqnNames(Num)
          Subst PmCleaned, " ", , "|"
        Else
          PmCleaned = "???"
        End If

      End If

    End With
    If ErCol Then PmCleaned = pscPm & PmCleaned
    Subst EqnOut, ExtrStr, "(" & PmCleaned & ")"
  End If

Loop Until InStr(EqnOut, "[") = 0 Or LoopCt > 99

If LoopCt > 99 Then
  CrashEnd , "-- the Task equation   " & _
              EqnIn & "   does not seem to be legal."
End If

' Put back original non-ref non-quote-enclosing square brackets as L/R parens
Subst EqnOut, "$$$", "(", "&&&", ")"

For Ref = 1 To Nrefs ' Put back original workbook/worksheet references

  If ShtRefOnly(Ref) Then
    Subst EqnOut, "###", "'" & ShtName(Ref) & "'!"
  Else
    Subst EqnOut, "@@@", "'[" & WbkName(Ref) & "]" & ShtName(Ref) & "'!"
  End If

Next Ref
End Sub

Sub PlaceEqnBox()

Dim tB As Boolean, SwB As Boolean
Dim SwapRat$, EqTest$, SwStr$, Eq$, ts1$, ts2$, BoxStr$, AddEq$
Dim SwitchStr$(), EqTot$()
Dim Pad%, BoxStrLen%, p%, q%, LeE%, k%, EqFragLen%, EqFragN%
Dim Eqn%, c%, MaxL%, Nspc%, i%, EqNum%, EqCt%, CharGroup%, EqLen%, MaxLeSS%
Dim TbxBottomRow%, TbxLeftCol%, ConstsBottomRow%, ConstsLeftCol%
Dim LeSS%()
Dim r&
Dim Ls!
Dim dc As Range, ExtPerr As Range
Dim SwNa As Variant, SwArray As Variant
Dim Tbx As TextBox, WtdAvChart As Object

With puTask

  If pbUPb Then
    r = [extperra1].Row + 4
  ElseIf .iNeqns = 0 Then
    Exit Sub
  Else
    r = 1 + plHdrRw
  End If

  MaxL = 0: EqCt = 0: ts2 = "": MaxLeSS = 0

  ReDim EqTot(piLwrIndx To .iNeqns), SwitchStr(piLwrIndx To .iNeqns), LeSS(piLwrIndx To .iNeqns)

  SwNa = Array("ST", "SA", "SC", "LA", "FO", "NU", "AR")

  For EqNum = piLwrIndx To .iNeqns

    If EqNum <> 0 Then
      Subst .saEqns(EqNum), " "
      ConvertIndexBracketed .saEqns(EqNum), Eq  ' 10/04/20 -- added
      SwapRat = ""

      If EqNum > 0 Then

        If pbUPb Then
          If piaSwapCols(EqNum) > 0 Then

          End If
        End If

        With .uaSwitches(EqNum)
          SwArray = Array(.ST, .SA, .SC, .LA, .FO, .Nu, .Ar)
        End With
        SwStr = ""

        For i = 1 To 7
          If SwArray(i) Then
            SwStr = SwStr & SwNa(i) & " "
          End If
        Next i

        If SwStr <> "" Then
          SwStr = "{" & Trim(SwStr) & "}"
        End If

        If EqNum > 0 And Eq <> "" Then

          ts1 = prSubsSpotNameFr(EqNum)
          If ts1 <> "" Then
            ts2 = " <" & ts1 & ">"
            SwStr = ts2 & SwStr
          End If

        End If

        SwitchStr(EqNum) = SwStr
        LeSS(EqNum) = Len(SwStr)
      End If

      MaxL = fvMax(MaxL, Len(Eq))
      MaxLeSS = fvMax(MaxLeSS, LeSS(EqNum))
      EqTot(EqNum) = Eq
    End If

  Next EqNum

End With

Set Tbx = ActiveSheet.TextBoxes.Add(0, 0, 10, 10)
With Tbx
  .Name = "Equations"
  BoxStr = "Task Equations for " & puTask.sName & pscLF2
  .Font.Name = "Lucida Console": .Font.Size = 11: .Font.Bold = False

  For EqNum = piLwrIndx To puTask.iNeqns

    If EqNum <> 0 Then
      Eq = EqTot(EqNum)

      If pbUPb Then
        tB = (puTask.bDirectAltPD And EqNum = -3)
        tB = tB Or (Not puTask.bDirectAltPD And EqNum = -2)
      Else
        tB = False
      End If

      If Not tB Then
        EqCt = 1 + EqCt: EqLen = Len(Eq)
        LeE = 0: ts2 = ""

        Select Case EqNum
          Case -1: ts1 = IIf(pbU, "206Pb/238U", "208Pb/232Th") & " const."
          Case -2: ts1 = IIf(pbTh, "206Pb/238U", "208Pb/232Th") & " const."
          Case -3: ts1 = "232Th/238U const."
          Case -4: ts1 = IIf(pbU, "U", "Th") & " concentration"
          Case Else
            ts1 = puTask.saEqnNames(EqNum)
            k = InStr(ts1, "||")
            If k > 0 Then ts1 = Left$(ts1, k - 1)
            Subst ts1, vbLf, " ", "|", " ", " "
        End Select

        AddEq = Eq & Space(1 + MaxL - Len(Eq) + MaxLeSS - LeSS(EqNum))
        AddEq = AddEq & SwitchStr(EqNum)
        BoxStr = BoxStr & AddEq & "  " & ts1 & vbLf
        psaEqShow(EqNum, 1) = Eq
        psaEqShow(EqNum, 2) = SwitchStr(EqNum)
        psaEqShow(EqNum, 3) = puTask.saEqnNames(EqNum)
        Subst psaEqShow(EqNum, 3), " ", , vbLf
      End If

    End If

  Next EqNum

  ts1 = Left$(BoxStr, Len(BoxStr) - 1): MaxL = Len(ts1)
  .Text = Left$(ts1, 200)
  For CharGroup = 1 To 1 + foAp.Fixed((Len(ts1) - 200) / 200, 0)
    .Characters(200 * CharGroup).Insert String:=Mid$(ts1, 200 * CharGroup, 200)
  Next CharGroup
  .AutoSize = False
  With .ShapeRange.TextFrame
    .MarginLeft = 35: .MarginRight = 25
    .MarginTop = 35:  .MarginBottom = 22
  End With
  .AutoSize = True: .Interior.Color = RGB(255, 200, 222)
  p = InStr(ts1, vbLf) - 1
  q = Len(puTask.sName)
  .Characters.Font.Size = 11
  .Characters(1, p).Font.italic = True
  .Characters(20, q).Font.FontStyle = "Bold Italic"
  .Characters(1, p).Font.Size = 13

  If pbUPb Then

    If fbRangeNameExists("ThPbStdAgesRatios") Then
      .Left = [ThPbStdAgesRatios].Left
      .Top = fnBottom([ThPbStdAgesRatios]) + 22
    ElseIf fbRangeNameExists("upbstdagesratios") Then
      .Left = [upbstdagesratios].Left
      .Top = fnBottom([upbstdagesratios]) + 22
    End If

    If .Left = 0 Then .Left = [stdcommpb].Left
  Else
    .Left = Columns(piLastCol + 3).Left + 5
    .Top = Rows(r).Top
  End If

  .Width = 7 + .Width

  If piNconstsUsed > 0 Then
    TbxBottomRow = fnLeftTopRowCol(1, fnBottom(Tbx))
    TbxLeftCol = fnLeftTopRowCol(2, Tbx.Left) + 1

    With [constantsused]
      .Cut Destination:=Cells(3 + TbxBottomRow, TbxLeftCol)
      Set dc = frSr(.Row - 1, .Column - 1, .Row + piNconstsUsed, .Column + 2)
      Box dc, , , , vbYellow, True

      For i = 1 To piNconstsUsed
        With Range(.Cells(i, 2), .Cells(i, 3))
          .Merge
          .HorizontalAlignment = xlLeft
        End With
      Next i

    End With

  End If

End With
ClearObj WtdAvChart, ExtPerr
End Sub

Sub CheckTrimCtNum(ByVal TrimCt&, TrimMass#(), TrimTime#())
Dim p%, q%
p = UBound(TrimMass, 1): q = UBound(TrimMass, 2)
If TrimCt >= 0.9 * q Then
  ReDim Preserve TrimMass(1 To p, 1 To q + 1000), TrimTime(1 To p, 1 To q + 1000)
End If
End Sub

Function fbOkEqn(ByVal EqNum%) As Boolean
Dim OK As Boolean

With puTask.uaSwitches(EqNum)

  If .SC Then
    OK = True
  Else
    OK = (prSubsSpotNameFr(EqNum) = "" Or fbIsInSubset(psSpotName, EqNum))
  End If

  If pbUPb And OK Then
    OK = OK And ((.ST And pbStd) Or (.SA And Not pbStd) Or Not (.ST Or .SA))
  End If

  fbOkEqn = OK
End With

End Function

Sub CreateSheets(ByVal ForStd As Boolean, SqSht() As Worksheet)
Dim ShtExists As Boolean, t$, Na$, t1$, t2$, i%, TabClr&

For i = 1 To piNumDauPar

  If ForStd And piStdCorrType >= 0 And piStdCorrType <= 2 And piPb46col > 0 Then
    t = fsS(4 - 3 * (piStdCorrType = 1) - 4 * (piStdCorrType = 2)) & "-corr|"
  Else
    t = ""
  End If

  t = t & IIf((pbU And i = 1) Or (pbTh And i = 2), "206Pb|/238U", "208Pb|/232Th")
  t = t & "|calibr.|const"
  psaUThPbConstColNames(ForStd, i) = t
Next i

Sheets.Add
t2 = ActiveSheet.Name

If ForStd Then
  Set phStdSht = ActiveSheet: Set SqSht(-1) = phStdSht
Else
  Set phSamSht = ActiveSheet
  Set SqSht(0) = phSamSht
End If

t1 = IIf(ForStd, pscStdShtNa, pscSamShtNa)
TabClr = IIf(ForStd, 8388736, 32768)
DupeNames t1, 1

' 09/05/18 -- Mod to protect against std sht or sam sht already present
ShtExists = False
Na = IIf(ForStd, pscStdShtNa, pscSamShtNa)

On Error GoTo 1
Sheets(Na).Activate
ShtExists = True
1 On Error GoTo 0
If ShtExists Then Sheets(Na).Delete

With Sheets(t2)
  .Activate
  .Tab.Color = TabClr
  .Name = IIf(ForStd, pscStdShtNa, pscSamShtNa)
End With

NoGridlines
With Cells.Font: .Name = psStdFont: .Size = 11: End With
Zoom piIzoom
plHdrRw = 6
plaFirstDatRw(-ForStd) = 1 + plHdrRw
plaLastDatRw(-ForStd) = plHdrRw + 1

If ForStd Then
  AssignUPbColNumbers
  phStdSht.Activate
  CreateStdConstBoxes
Else
  phSamSht.Activate
  CreateSampleCommPbBox
End If

PlaceUPbHeaders plHdrRw, ForStd
Columns(1).NumberFormat = "@"
End Sub

Sub CleanData(RangeIn, RangeVals#(), Nrows%, Optional NoZero As Boolean = False, _
  Optional NoNeg As Boolean = False)
' RangeIn is a rectangular range, RangeVals is the numeric values in RangeIn,
'  excluding rows with one or more non-numeric cells.

Dim tB As Boolean, Rng As Boolean
Dim i%, j%, Nc%, Nr%, ct%
Dim Tdat#()
Dim v() As Variant

If TypeName(RangeIn) = "Range" Then
  Nc = RangeIn.Columns.Count
  Nr = RangeIn.Rows.Count
  Rng = True
Else
  Nc = UBound(RangeIn, 2)
  Nr = UBound(RangeIn, 1)
  Rng = False
End If

ct = 0
ReDim Tdat(1 To Nr, 1 To Nc), v(1 To Nc)

For i = 1 To Nr
  tB = True

  For j = 1 To Nc
    v(j) = RangeIn(i, j)
    tB = IsNumeric(v(j))
    If tB Then tB = fbIsNumber(v(j))

    If tB Then
      If (v(j) = 0 And NoZero) Or (v(j) < 0 And NoNeg) Then tB = False
    End If

    If Not tB Then Exit For
  Next j

  If tB Then
    ct = 1 + ct
    For j = 1 To Nc: Tdat(ct, j) = v(j): Next
  End If

Next i

If Nc = 1 Then
  ReDim Preserve RangeVals(1 To ct)
  For i = 1 To ct: RangeVals(i) = Tdat(i, 1): Next
Else
  ReDim Preserve RangeVals(1 To ct, 1 To Nc)

  For i = 1 To ct
    For j = 1 To Nc
      RangeVals(i, j) = Tdat(i, j)
  Next j, i

End If

Nrows = ct
End Sub

Sub ClearObj(a As Object, Optional b, Optional c, Optional d, _
  Optional e, Optional f, Optional g, Optional h, Optional i, Optional j)
If fbNIM(a) Then Set a = Nothing
If fbNIM(b) Then Set b = Nothing
If fbNIM(b) Then Set c = Nothing
If fbNIM(b) Then Set d = Nothing
If fbNIM(b) Then Set e = Nothing
If fbNIM(b) Then Set f = Nothing
If fbNIM(b) Then Set g = Nothing
If fbNIM(b) Then Set h = Nothing
End Sub

Function foAp() As Excel.Application
Set foAp = Excel.Application
End Function

Sub TaskSolverCall(Optional EqNum% = 0)
' Invokes Excel's Solver (via SolveThis) to minimize the value
' in the CellToMinimize by varying the values in the CellsToVary range.

Dim RevTerm$, Paren$, v$, vv$
Dim i%, j%, r&, c%, p%, q%, u%, Start%, Endd%, Nranges%
Dim RangeSpec As Variant, CellToMinimize As Variant
Dim CellsToVary As Variant, ToVary As Variant

foAp.Calculate
With puTask
  Start = IIf(EqNum > 0, EqNum, 1)
  Endd = IIf(EqNum > 0, EqNum, .iNeqns)

  For i = Start To Endd

    If .baSolverCall(i) Then
      ParseLine .saEqns(i), RangeSpec, Nranges, ";"

      If Nranges = 2 Then
        v = StrReverse(RangeSpec(1))
        p = InStr(v, ")"): q = InStr(v, "(")

        If p = 1 And q < 10 Then
          vv = StrReverse(Left$(v, q))
          u = InStr(vv, ",")
          r = fvMax(1, Val(Mid$(vv, 2)))
          c = fvMax(1, Val(Mid$(vv, 1 + u)))
          RangeSpec(1) = StrReverse(Mid$(v, 1 + q))
        Else
          r = 1: c = 1
        End If

        Set CellToMinimize = Range(fsLegalName(RangeSpec(1), True))(r, c)
        ParseLine RangeSpec(2), ToVary, Nranges, ","
        Set CellsToVary = Range(fsLegalName(ToVary(1), True))

        For j = 1 To Nranges
          Set CellsToVary = Union(CellsToVary, Range(fsLegalName(ToVary(j), True)))
        Next j

        SolveThis CellToMinimize, CellsToVary, 2
      End If

    End If

  Next i

End With
End Sub

Sub KillTheForm()
 Unload Splash
End Sub

Function fbDataIsDriftCorr()
Dim tB As Boolean, Co%, Rw&, HdrRowStd&, ShtIn As Worksheet

If pbSecularTrend Then
  tB = True
Else
  tB = False
  Set ShtIn = ActiveSheet
  Sheets(pscStdShtNa).Activate
  HdrRowStd = flHeaderRow(-1)
  FindStr "corrected for secular drift of age standard", Rw, Co, 1, 1, HdrRowStd - 1, 99
  If Rw <> 0 Then tB = (Left$(Cells(Rw, Co), 2) <> "Un")
  ShtIn.Activate
End If

fbDataIsDriftCorr = tB
End Function

Sub GetScreenResolution(WidthPixels&, HeightPixels&)
Dim Sys As New CSysInfo ' p. 533, Excel 2002 VBA

On Error Resume Next
With Sys
  HeightPixels = .ScreenHeight
  WidthPixels = .ScreenWidth
End With
On Error GoTo 0
End Sub
