Attribute VB_Name = "ScanGraphics"
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

' Scan graphics for SHRIMP raw-data
' K.R. Ludwig, Berkeley Geochrono
Option Explicit
Option Base 1

Dim mnChtWidth!, mnChtHeight!, msTmpSqGraf$, msTmpSqGrafS$
Dim mraTr() As Range, mraCr() As Range, mdMaxX#
Dim miNumSpotsToChart%, miNumSpotsCharted%, miNumSpots% ', miStartSpot%

Sub AllSpotScanGraphics()
CreateScanGraphics , True
End Sub

Sub SingleSpotScanGraphics()
CreateScanGraphics
End Sub

Sub CreateScanGraphics(Optional ByVal NotFromMenu As Boolean = False, Optional AllSpots As Boolean = False)
' assumes active cell is the sample name

Dim CondensedSht As Boolean, tB1 As Boolean, tB2 As Boolean, tB3 As Boolean, tB5 As Boolean
Dim FromMenu As Boolean, tB4 As Boolean, GotSht As Boolean, StoreTempBks As Boolean
Dim SingleSpot As Boolean, StdSpot As Boolean, StoreWbk As Boolean

Dim s$, Ws$, rnx$, rny$, DrNa$, Wna$, RangeNa$, ShNa$, DrA$(), saNames$(), mraTrAddr$(), mraCrAddr$()

Dim i%, j%, k%, c%, Col1%, Npks%, p%, OkCt%, FirstSBMrw%, LastSBMrw%
Dim SpotNum%, PkNum%, SpotNpks%, NumGrafBks%, SbmCt%, tCo%
Dim WbkStoredCt%, SpotsPerSht%, Nscans%, WbkCt%, NumAreas%, NumCells%

Dim r&, Rw&, rw1&, rw2&, tRw#, FirstDatRw&, LastDatRw&, SpotR&
Dim SBMrow&, Nrows&, CondRH&(), Crows&(), CondShtNameRows&()

Dim SpotPks#()

Dim Dr() As Range, SigR() As Range, OkR As Range, SelR As Range, SBMrange As Range
Dim mraTr() As Range, mraCr() As Range

Dim ShtIn As Worksheet, Gbks() As Workbook
Dim Ob As Object, v  As Variant

If Workbooks.Count = 0 Then Exit Sub
If ActiveSheet.Type <> xlWorksheet Then Exit Sub

tB1 = (Cells(1, 10) = ("SQUID grouped-sample sheet"))
tB2 = (InStr(Cells(1, 1), "Isotope Ratios of S") = 1)
tB3 = (Cells(2, 1) = "(errors are 1s unless otherwise specified)")
tB4 = (Cells(1, 1) = "Isotope Ratios")
tB5 = Not (tB2 Or tB4) Or Not tB3

Set ShtIn = ActiveSheet
CondensedSht = fbIsCondensedSheet
Set SelR = Selection

If CondensedSht Then
  GetCondensedShtInfo
ElseIf Not tB1 Then
  If tB5 Then Exit Sub
  If piFileNpks = 0 Then
    FindCondensedSheet tB1
    If Not tB1 Or piFileNpks = 0 Then
      MsgBox "The active workbook doesn't appear to contain a SQUID2-condensed raw-data sheet.", , pscSq
      End
    End If 'lseIf Not tB5 Then
      'Exit Sub 'CondensedSht = True
    'End If
  End If
End If

SpotsPerSht = foUser("NumTmpShts")
GetInfo
ShowStatusBar
NoUpdate
StdSpot = (ActiveSheet.Name = pscStdShtNa)

If LCase(ActiveSheet.Name) = "within-spot ratios" Then
  MsgBox "You can only request scan graphics from a worksheet " & _
         "with valid sample names in the first column.", , pscSq
  End
End If

FromMenu = Not NotFromMenu
If Workbooks.Count = 0 Then CreateNewWorkbook
If Not CondensedSht Then plHdrRw = flHeaderRow(StdSpot, , True)

If CondensedSht Then
  Set phCondensedSht = ActiveSheet
  Set pwDatBk = ActiveWorkbook
  GotSht = True
Else
  FindCondensedSheet GotSht
  If GotSht Then
    Set phCondensedSht = ActiveSheet
  End If
End If

If Not GotSht Then
  MsgBox "Can't find the raw PD- or XML- data sheet.", , pscSq: End
End If

SpotNpks = piFileNpks  ' 09/06/14 -- move from earler, when condensed-sht info might not
miNumSpots = 0         '             have been acquired, to here.

Nrows = flEndRow
ReDim CondRH(1 To piNumAllSpots)
NoUpdate
StatBar "Wait"
With ActiveSheet

  For r = 1 To piNumAllSpots
    tB1 = .Rows(r + 1).Hidden
    CondRH(r) = tB1
    If tB1 Then .Rows(r + 1).Hidden = False
  Next r

End With

If CondensedSht Or AllSpots Then
  miNumSpots = IIf(AllSpots, piNumAllSpots, 1)
  ReDim saNames(1 To miNumSpots), CondShtNameRows(1 To miNumSpots)

  For r = 2 To piNumAllSpots + 1
    Rw = plaSpotNameRowsCond(r - 1)

    If AllSpots Then
      CondShtNameRows(r - 1) = Rw
      saNames(r - 1) = Cells(Rw, picNameDateCol)
    ElseIf Rw > SelR.Row Then
      If r < 3 Then End
      CondShtNameRows(1) = plaSpotNameRowsCond(r - 2)
      saNames(1) = psaSpotNames(r - 2)
      Exit For
    End If

  Next r

Else
  ShtIn.Activate
  With SelR
    NumAreas = .Areas.Count
    NumCells = .Cells.Count
    ReDim Crows(1 To NumCells)

    For i = 1 To NumAreas
      For j = 1 To .Areas(i).Cells.Count
        miNumSpots = 1 + miNumSpots
        Crows(miNumSpots) = .Areas(i)(j).Row
    Next j, i

  End With
  ReDim saNames(1 To miNumSpots)
End If

If Not (CondensedSht Or AllSpots) Then
  ReDim CondShtNameRows(1 To miNumSpots)

  For i = 1 To miNumSpots
    saNames(i) = Cells(Crows(i), 1)
  Next i

  phCondensedSht.Activate
  On Error Resume Next

  For i = 1 To miNumSpots

    For j = 1 To piNumAllSpots
      If psaSpotNames(j) = saNames(i) Then
        CondShtNameRows(i) = plaSpotNameRowsCond(j)
        Exit For
      End If
    Next j

  Next i

  On Error GoTo 0
End If

With ActiveSheet
  For r = 1 To piNumAllSpots
    .Rows(r + 1).Hidden = CondRH(r)
  Next r
End With

SingleSpot = (miNumSpots = 1)
StoreTempBks = (Not SingleSpot And foUser("StoreTmpBks"))

Do
  If Workbooks.Count = 0 Then CreateNewWorkbook
  NoUpdate
  Npks = SpotNpks
If Npks > 0 Then Exit Do
  MsgBox "Not a valid raw-data PD file.", , pscSq
  End
Loop

Set pwDatBk = ActiveWorkbook
Set phCondensedSht = ActiveSheet
ReDim ScanPks(1 To SpotNpks)

For i = 1 To SpotNpks

  If FromMenu Or SingleSpot Then
    ScanPks(i) = pdaFileNominal(i) ' Drnd(SpotPks(i), 4)
  Else
    ScanPks(i) = pdaFileNominal(i)
  End If

Next i

If StoreTempBks Then
  msTmpSqGraf = fsSquidUserFolder & "tmpSquid"
  msTmpSqGrafS = msTmpSqGraf & fsPathSep
  On Error Resume Next
  Kill msTmpSqGrafS & "*.*"
  MkDir msTmpSqGraf
  On Error GoTo 0
End If

Zoom 85
phCondensedSht.Activate
Cells(1, 1).Activate
NoUpdate
Col1 = picDatCol
j = 1 + SpotNpks: WbkCt = 0
NumGrafBks = -Int(-miNumSpots / SpotsPerSht)

ReDim mraTr(1 To j), mraCr(1 To j), Dr(1 To j), SigR(1 To j - 1)
ReDim mraTrAddr$(1 To j), mraCrAddr$(1 To j)

If NumGrafBks > 1 Then ReDim Gbks(1 To NumGrafBks)
ReDim DrA(1 To miNumSpots, 1 To 2)
NoUpdate
If NumGrafBks > 2 Then CreateNewWorkbook
WbkCt = 1
Set pwGrafBk = ActiveWorkbook

If NumGrafBks > 2 Then

  Set Gbks(WbkCt) = pwGrafBk

  For i = 1 To ActiveWorkbook.Sheets.Count - 1
    ActiveSheet.Delete
  Next

  phCondensedSht.Activate

End If

miNumSpotsToChart = miNumSpots
WbkStoredCt = 0

'If miNumSpots > 20 Then
'  EscapeTrap.Show vbModeless
'End If

For SpotNum = 1 To miNumSpots
  StatBar IIf(SingleSpot, SpotNum, miNumSpots - SpotNum + 1)
  SpotR = CondShtNameRows(SpotNum)
  phCondensedSht.Activate
  GetNameDatePksScans CondShtNameRows(SpotNum), psSpotName, , , Nscans
  FirstDatRw = SpotR + picDatRowOffs
  LastDatRw = FirstDatRw + Nscans - 1
  pwGrafBk.Activate
  AddGrafSht SpotNum, NumGrafBks < 3
  SbmCt = Nscans * SpotNpks
  SBMrow = FirstDatRw
  Set SBMrange = frSr(1, picSBMcol, SbmCt, 1 + picSBMcol) '???
  With phCondensedSht
    k = 0

    For i = 1 To SpotNpks
      For j = 1 To Nscans
        k = k + 1
        r = SBMrow + j - 1
        c = picDatCol + (i - 1) * 5
        SBMrange(k, 1) = .Cells(r, c)
        SBMrange(k, 2) = .Cells(r, c + 3) / .Cells(FirstDatRw - 2, c)
    Next j, i

  End With
  SBMrange.Sort SBMrange(1, 1)
  Fonts SBMrange, , , , Hues.peMedGray

  For PkNum = 1 To SpotNpks + 1
    phCondensedSht.Activate

    If PkNum > SpotNpks Then ' sbm
      With SBMrange
        c = .Column
        FirstSBMrw = .Row
        LastSBMrw = .Rows.Count
      End With
    Else
      c = Col1 + 5 * (PkNum - 1)
      Set SigR(PkNum) = frSr(FirstDatRw, c + 2, LastDatRw)
    End If

    If LastDatRw <= FirstDatRw Then
      MsgBox "Not a SQUID-created Sample or Standard reduced-data worksheet", , pscSq
      End
    End If

    If PkNum > SpotNpks Then
      rw1 = FirstSBMrw: rw2 = LastSBMrw
    Else
      rw1 = FirstDatRw: rw2 = LastDatRw
    End If

    If PkNum > SpotNpks Then            ' 09/06/14 -- added the ' chars for Wna & ShNa
      Wna = "'[" & pwGrafBk.Name & "]"
      ShNa = phGrafSht.Name & "'!"
    Else
      Wna = "'[" & pwDatBk.Name & "]"
      ShNa = phCondensedSht.Name & "'!"
    End If

    Ws = Wna & ShNa
    OkCt = 0

' -------------------------------------------------------
' 09/05/18 added lines below so can deal with rejected scans
    For r = rw1 To rw2
      Set OkR = frSr(r, c, , c + 2)
      If OkR(2).NumberFormat <> fsRejFormat Then
        OkCt = 1 + OkCt
        If OkCt = 1 Then
          Set mraTr(PkNum) = OkR(1)
          Set mraCr(PkNum) = OkR(2)
        Else
          Set mraTr(PkNum) = Union(mraTr(PkNum), OkR(1))
          Set mraCr(PkNum) = Union(mraCr(PkNum), OkR(2))
        End If
      Else

      End If
    Next r

    rnx = mraTr(PkNum).Address
    p = InStr(rnx, ",")

    If p = 0 Then
      rnx = Ws & rnx
    Else ' Add workbook & sheet name
      rnx = Ws & Left(rnx, p) & Ws & Mid(rnx, p + 1)
    End If

    rny = mraCr(PkNum).Address
    p = InStr(rny, ",")

    If p = 0 Then
      rny = Ws & rny
    Else ' Add workbook & sheet name
      rny = Ws & Left(rny, p) & Ws & Mid(rny, p + 1)
    End If
' -----------------------------------------------------------
    DrNa = rnx & "," & rny
    Set Dr(PkNum) = Range(DrNa)

    If PkNum > SpotNpks Then
      phGrafSht.Activate
      mdMaxX = SBMrange(SbmCt, 1)
    Else
      tRw = mraTr(PkNum).Row - 1
      tCo = mraTr(PkNum).Column
      Do Until Cells(tRw, tCo) = "Secs"
        tRw = tRw - 1
      Loop
      mdMaxX = fvMax(10, Cells(tRw + Nscans, tCo))
    End If

    pwGrafBk.Activate
    phGrafSht.Activate
    FormatCharts DrNa, PkNum, SpotNpks, SpotNum%, False
  Next PkNum

  phCondensedSht.Activate
  With Dr(SpotNpks)
    DrA(SpotNum, 1) = .Address
  End With
  On Error GoTo 0

  phGrafSht.Activate
  p = 1 + foLastOb(ActiveSheet.ChartObjects).BottomRightCell.Row
  Fonts p, 1, , , RGB(0, 0, 96), , , 10, , , "Spot#" & StR(SpotNum) & ", " & psSpotName
  Fonts p + 1, 1, , , RGB(0, 0, 96), , , 8, , , fsInQ(phCondensedSht.Name) & _
    " rows" & StR(FirstDatRw) & "-" & fsS(LastDatRw)
  Fonts p + 2, 1, , , 128, , xlLeft, 8, , , "Units are total counts"
  frSr(p, 1, p + 2).IndentLevel = 2
  Blink
  miNumSpotsCharted = SpotNum '+ 1 - miStartSpot
  StoreWbk = (miNumSpotsCharted = miNumSpotsToChart Or _
              miNumSpotsCharted Mod SpotsPerSht = 0 Or _
             (SpotNum = piNumAllSpots) Or SingleSpot)
  StoreWbk = StoreWbk And NumGrafBks > 2

  If StoreWbk Then

    If StoreTempBks Then
      WbkStoredCt = 1 + WbkStoredCt
      Alerts False
      s = fsS(WbkStoredCt) & ".xls"
      StatBar "Saving temporary workbook"
      pwGrafBk.SaveAs msTmpSqGrafS & s, , , , False
      pwGrafBk.Close
      StatBar
      Alerts True
    End If

    If (StoreWbk And StoreTempBks) Then
      If WbkCt >= NumGrafBks Then Exit For
      CreateNewWorkbook
      WbkCt = 1 + WbkCt
      Set pwGrafBk = ActiveWorkbook
      Set Gbks(WbkCt) = ActiveWorkbook
      Alerts False

      If NumGrafBks > 1 Then
        Do Until ActiveWorkbook.Sheets.Count = 1
          ActiveSheet.Delete
        Loop
      End If

    End If

    Set phGrafSht = ActiveSheet
    phCondensedSht.Activate
  End If

  Cells(1, 1).Select
  If pbEscapeSquid Then
    End
  End If
Next SpotNum

If NumGrafBks > 2 Then
  NoUpdate
  AssembleScanGrafix msTmpSqGrafS, piNumAllSpots, WbkStoredCt, SpotsPerSht, _
                     Gbks, StoreTempBks, SingleSpot

  For Each Ob In ActiveSheet.Shapes
    If Left$(Ob.Name, 4) = "Rejb" Then
      Ob.Delete
    End If
  Next Ob

End If
Alerts True
StatBar
End Sub

Sub CreateChart(SourceRange$)
Dim Xtik#, X#, ActCht As Chart
Charts.Add ' Create the basic, unformatted scan chart
Set ActCht = ActiveChart
With ActCht
  .ChartType = xlXYScatterSmooth: .HasLegend = False
  .SetSourceData Source:=Range(SourceRange), PlotBy:=xlColumns
  ActCht.Activate
  .Location Where:=xlLocationAsObject, Name:=phGrafSht.Name
End With
Tick mdMaxX, Xtik

Do
  X = X + Xtik
Loop Until X >= mdMaxX

With ActiveChart
  .HasTitle = False
  With .Axes(xlCategory)
    .HasMajorGridlines = False: .HasMinorGridlines = False
    .MaximumScale = X: mdMaxX = X
  End With
  With .Axes(xlValue)
    .HasMajorGridlines = True: .HasMinorGridlines = False
  End With
End With

With phGrafSht.Shapes: .Item(.Count).LockAspectRatio = True: End With
End Sub

Sub FormatCharts(DataRange$, ByVal PkNum%, ByVal Npks%, SpotNum%, _
  CFshape, Optional ChtCap)
Dim SBM As Boolean, ChtName$, ChtOnAct$, ChtCaption$
Dim Co%, Rw&
Dim Xspace!, Yspace!, ChtLeft!, ChtTop!
Dim DatRange As Range, Cht As Object

Yspace = 20: Xspace = 5
NoUpdate
mnChtWidth = 0.92 * ActiveWindow.Width / 4 - Xspace
mnChtHeight = 0.88 * ActiveWindow.Height / 3 - Yspace
CreateChart DataRange
Set Cht = ActiveSheet
Cht.Cells(1, 1).Activate
ChtLeft = Xspace + (PkNum - 1) Mod 4
ChtTop = Yspace + (PkNum - 1) \ 4
SBM = (PkNum > Npks)
Set DatRange = Range(DataRange)

If fbNIM(ChtCap) Then
  ChtCaption = ChtCap
End If

If SBM Then
  ChtCaption = "SBM  "
ElseIf PkNum = piBkrdPkOrder Then
  ChtCaption = "Bkgrd"
ElseIf ChtCaption = "" Then
  Rw = -1 ' because of possible rejected spots
  Do Until DatRange(Rw, 2) = "mass"
    Rw = Rw - 1
  Loop

  ChtCaption = DatRange(Rw + 1, 2)
End If

ChtName = "Spot#" & fsS(SpotNum)
If SBM Then ChtName = ChtName & " SBM" Else _
  ChtName = ChtName & " Mass station" & StR(PkNum)
ChtLeft = ((PkNum - 1) Mod 4) * mnChtWidth + Xspace ' creating
ChtTop = ((PkNum - 1) \ 4) * mnChtHeight + Yspace ' creating

With foLastOb(ActiveSheet.ChartObjects)
  .Height = mnChtHeight: .Width = mnChtWidth ' creating
  .Top = ChtTop: .Left = ChtLeft ' creating
  .Activate
End With

If LCase(Right$(ChtName, 3)) = "sbm" Then
  ChtOnAct = "SBMclick"
Else
  ChtOnAct = "ChartRej" & fsS(PkNum)
End If

FchartFrag DataRange, CFshape, SBM, ChtCaption, ChtName, Cht, ChtOnAct
End Sub

Sub AssembleScanGrafix(Fpath$, ByVal NumSpots%, LastWbkCt%, ByVal SpotsPerSht%, _
  Gbks() As Workbook, ByVal StoreTempBks As Boolean, _
  ByVal SingleSpot As Boolean)

Dim FileExists As Boolean
Dim t$, s$
Dim i%, f%, q%
Dim nwS As Object, Shp As Shape

Alerts False
StatBar "   Assembling charts from disk..."
NoUpdate
q = 0

For i = 1 To LastWbkCt
    StatBar "Combining temporary files... " & StR(LastWbkCt - i + 1) '1 + i - miStartSpot)
    If StoreTempBks Then
      OpenWorkbook Fpath & fsS(i) & ".XLS", FileExists
      If Not FileExists Then MsgBox "Couldn't load " & Fpath, , pscSq: End
      Set Gbks(i) = ActiveWorkbook
    End If
    Set nwS = Gbks(i).Sheets

    For f = nwS.Count To 1 Step -1
      t = nwS(f).Name

      If Left$(t, 5) <> "Sheet" Then
        q = q + 1
        nwS(f).Copy After:=phCondensedSht
      End If

    Next f

    Gbks(i).Close
Next i

If StoreTempBks Then
  Alerts False
  On Error Resume Next
  Kill msTmpSqGrafS & "*.*"
  RmDir msTmpSqGraf
  On Error GoTo 0
End If

For Each Shp In ActiveSheet.Shapes
  s = Shp.Name
  If Left$(s, 4) = "rbox" Then
    Shp.Delete
  End If
Next

End Sub

Sub AddGrafSht(ByVal ScanShtCt%, Optional AddAfter As Boolean = False)
Dim s$, ct%, e&

If AddAfter Then
  Sheets.Add After:=phCondensedSht
Else
  Sheets.Add
End If

Set phGrafSht = ActiveSheet
s = fsLegalRangeAndSheetName(psSpotName, True)
DupeNames s, 1
NoGrids
NoUpdate
Alerts False
On Error GoTo 0

1: On Error GoTo 2
  ActiveSheet.Name = s:  phGrafSht.Name = s
  Exit Sub
2: e = Err.Number
On Error GoTo 0

If e = 1004 And ct < 99 Then
  ct = 1 + ct
  s = s & "-" & fsS(ct)
  GoTo 1
Else
  MsgBox "Error in assigning sheet-names for scan-graphics" & vbLf & "Must terminate."
End If
End Sub

Sub PicConvert() ' Converts last chart in a sheet to pictures
Dim L!, t!, w!, h!, Ch As ChartObject
NoUpdate
Set Ch = foLastOb(ActiveSheet.ChartObjects)
With Ch
  L = .Left: t = .Top: w = .Width: h = .Height
End With

Ch.CopyPicture
ActiveSheet.Paste
Ch.Delete

With foLastOb(ActiveSheet.Shapes)
  .Left = L: .Top = t: .Height = h: .Width = w
End With
End Sub

Sub RemoveDblChars(Phrase$, Optional CharsToUndouble$ = " ")
Dim Le%, Le0%, ChD$
ChD = CharsToUndouble & CharsToUndouble
Le = Len(Phrase)
Do
  Le0 = Le
  Subst Phrase, ChD, CharsToUndouble
  Le = Len(Phrase)
Loop Until Le = Le0
End Sub

Sub FindCondensedSheet(GotSheet As Boolean, Optional DontGetInfo As Boolean = False)
Dim Sht As Worksheet

GotSheet = False
If Workbooks.Count = 0 Then Exit Sub
GotSheet = fbIsCondensedSheet

If Not GotSheet Then

  For Each Sht In ActiveWorkbook.Worksheets
    If fbIsCondensedSheet(Sht) Then
      GotSheet = True
      Sht.Activate
      Exit For
    End If

  Next Sht

End If

If GotSheet And Not DontGetInfo Then
  GetCondensedShtInfo
End If

End Sub

Sub FindStdOrSampleSheets(GotSample As Boolean, GotStd As Boolean)
Dim Sht As Worksheet
GotSample = False: GotStd = False

If Workbooks.Count > 0 Then

  For Each Sht In ActiveWorkbook.Worksheets

    If Sht.Name = pscSamShtNa Then
      Set phSamSht = Sht: GotSample = True
    ElseIf Sht.Name = pscStdShtNa Then
      Set phStdSht = Sht: GotStd = True
    End If

    If GotSample And GotStd Then Exit Sub
  Next Sht

End If

End Sub

Sub SBMclick()
MsgBox "You can't reject SBM data.", , "SQUID"
End Sub

Sub ChartRej1()
SetupRejbox 1
End Sub
Sub ChartRej2()
SetupRejbox 2
End Sub
Sub ChartRej3()
SetupRejbox 3
End Sub
Sub ChartRej4()
SetupRejbox 4
End Sub
Sub ChartRej5()
SetupRejbox 5
End Sub
Sub ChartRej6()
SetupRejbox 6
End Sub
Sub ChartRej7()
SetupRejbox 7
End Sub
Sub ChartRej8()
SetupRejbox 8
End Sub
Sub ChartRej9()
SetupRejbox 9
End Sub
Sub ChartRej10()
SetupRejbox 10
End Sub
Sub ChartRej11()
SetupRejbox 11
End Sub
Sub ChartRej12()
SetupRejbox 12
End Sub
Sub ChartRej13()
SetupRejbox 13
End Sub
Sub ChartRej14()
SetupRejbox 14
End Sub
Sub ChartRej15()
SetupRejbox 15
End Sub
Sub ChartRej16()
SetupRejbox 16
End Sub
Sub ChartRej17()
SetupRejbox 17
End Sub
Sub ChartRej18()
SetupRejbox 18
End Sub
Sub ChartRej19()
SetupRejbox 19
End Sub
Sub ChartRej20()
SetupRejbox 20
End Sub
Sub ChartRej21()
SetupRejbox 21
End Sub
Sub ChartRej22()
SetupRejbox 22
End Sub
Sub ChartRej23()
SetupRejbox 23
End Sub
Sub ChartRej24()
SetupRejbox 24
End Sub
Sub ChartRej25()
SetupRejbox 25
End Sub
Sub ChartRej26()
SetupRejbox 26
End Sub
Sub ChartRej27()
SetupRejbox 27
End Sub
Sub ChartRej28()
SetupRejbox 28
End Sub
Sub ChartRej29()
SetupRejbox 29
End Sub
Sub ChartRej30()
SetupRejbox 30
End Sub
Sub ChartRej31()
SetupRejbox 31
End Sub
Sub ChartRej32()
SetupRejbox 32
End Sub
Sub ChartRej33()
SetupRejbox 33
End Sub
Sub ChartRej34()
SetupRejbox 34
End Sub
Sub ChartRej35()
SetupRejbox 35
End Sub
Sub ChartRej36()
SetupRejbox 36
End Sub
Sub ChartRej37()
SetupRejbox 37
End Sub
Sub ChartRej38()
SetupRejbox 38
End Sub
Sub ChartRej39()
SetupRejbox 39
End Sub
Sub ChartRej40()
SetupRejbox 40
End Sub
Sub ChartRej41()
SetupRejbox 41
End Sub
Sub ChartRej42()
SetupRejbox 42
End Sub
Sub ChartRej43()
SetupRejbox 43
End Sub
Sub ChartRej44()
SetupRejbox 44
End Sub
Sub ChartRej45()
SetupRejbox 45
End Sub
Sub ChartRej46()
SetupRejbox 46
End Sub
Sub ChartRej47()
SetupRejbox 47
End Sub
Sub ChartRej48()
SetupRejbox 48
End Sub
Sub ChartRej49()
SetupRejbox 49
End Sub
Sub ChartRej50()
SetupRejbox 50
End Sub

Sub rej1()
RejectPoint 1
End Sub
Sub rej2()
RejectPoint 2
End Sub
Sub rej3()
RejectPoint 3
End Sub
Sub rej4()
RejectPoint 4
End Sub
Sub rej5()
RejectPoint 5
End Sub
Sub rej6()
RejectPoint 6
End Sub
Sub rej7()
RejectPoint 7
End Sub
Sub rej8()
RejectPoint 8
End Sub
Sub rej9()
RejectPoint 9
End Sub
Sub rej10()
RejectPoint 10
End Sub
Sub rej11()
RejectPoint 11
End Sub
Sub rej12()
RejectPoint 12
End Sub
Sub rej13()
RejectPoint 13
End Sub
Sub rej14()
RejectPoint 14
End Sub
Sub rej15()
RejectPoint 15
End Sub
Sub rej16()
RejectPoint 16
End Sub
Sub rej17()
RejectPoint 17
End Sub
Sub rej18()
RejectPoint 18
End Sub
Sub rej19()
RejectPoint 19
End Sub
Sub rej20()
RejectPoint 20
End Sub
Sub rej21()
RejectPoint 21
End Sub
Sub rej22()
RejectPoint 22
End Sub
Sub rej23()
RejectPoint 23
End Sub
Sub rej24()
RejectPoint 24
End Sub
Sub rej25()
RejectPoint 25
End Sub
Sub rej26()
RejectPoint 26
End Sub
Sub rej27()
RejectPoint 27
End Sub
Sub rej28()
RejectPoint 28
End Sub
Sub rej29()
RejectPoint 29
End Sub
Sub rej30()
RejectPoint 30
End Sub
Sub rej31()
RejectPoint 31
End Sub
Sub rej32()
RejectPoint 32
End Sub
Sub rej33()
RejectPoint 33
End Sub
Sub rej34()
RejectPoint 34
End Sub
Sub rej35()
RejectPoint 35
End Sub
Sub rej36()
RejectPoint 36
End Sub
Sub rej37()
RejectPoint 37
End Sub
Sub rej38()
RejectPoint 38
End Sub
Sub rej39()
RejectPoint 39
End Sub
Sub rej40()
RejectPoint 40
End Sub
Sub rej41()
RejectPoint 41
End Sub
Sub rej42()
RejectPoint 42
End Sub
Sub rej43()
RejectPoint 43
End Sub
Sub rej44()
RejectPoint 44
End Sub
Sub rej45()
RejectPoint 45
End Sub
Sub rej46()
RejectPoint 46
End Sub
Sub rej47()
RejectPoint 47
End Sub
Sub rej48()
RejectPoint 48
End Sub
Sub rej49()
RejectPoint 49
End Sub
Sub rej50()
RejectPoint 50
End Sub

Function flRejBoxClr(ByVal Rejected As Boolean)
flRejBoxClr = IIf(Rejected, RGB(255, 212, 212), RGB(202, 202, 255))
End Function

Sub SetupRejbox(ByVal MassPos%)
Dim ErasedRejBoxes As Boolean, Rejected() As Boolean
Dim s$, f$, ScF$
Dim i%, p%, q%, Rct%, SpotCol%, OrigN%, CoCt%, NewN%, OldMassPos%, Nscans%
Dim SpotR1&, SpotR2&, OrigR1&, OrigR2&
Dim L!, wW!, bw!, bH!, PlotboxW!, PlotBoxT!
Dim CtsR As Range, OrigR As Range, NewR As Range, SecsR As Range, xyR As Range, ScRa As Range
Dim CoH() As ChartObject, Cht As ChartObject, Chh As ChartObjects
Dim Shp As Shape, Shps As Shapes

NoUpdate
ErasedRejBoxes = False
With ActiveSheet
  Set Shps = .Shapes
  Set Chh = .ChartObjects
End With
CoCt = Chh.Count

If CoCt > 0 Then
  ReDim CoH(1 To CoCt)
  For i = 1 To CoCt: Set CoH(i) = Chh(i): Next
End If

For Each Shp In Shps
  s = Shp.Name: f = Right$(s, 1)

  If Left$(s, 4) = "rbox" And fbIsNumber(f) Then
    If OldMassPos = 0 And Val(f) = 1 Then FindChartFromRejbox OldMassPos
    Shp.Delete: ErasedRejBoxes = True
  End If

Next Shp

If OldMassPos = MassPos Then Exit Sub
Set phGrafSht = ActiveSheet
FindStr "Spot#", i, , 1, , 999, 1, True
If i = 0 Then MsgBox "Can't find Spot# label on this sheet", , pscSq: End
s = Mid$(Cells(i + 1, 1).Formula, 2)
p = InStr(s, pscQ & " rows ")
Set phCondensedSht = Sheets(Left$(s, p - 1))
'SpotNum = Val(Mid$(Cells(i, 1).Text, 6))
s = Cells(i + 1, 1).Text
p = InStr(s, "rows"): q = InStr(s, "-")
OrigR1 = Val(Mid$(s, p + 4))
OrigR2 = Val(Mid$(s, q + 1))
SpotCol = picDatCol + 1 + (MassPos - 1) * 5

If ErasedRejBoxes Then
  q = 0
  With CoH(OldMassPos)
    s = Right$(.Name, 1)
    If fbIsNumber(s) Then q = Val(s)
  End With
  ActiveSheet.ChartObjects(MassPos).Select
  ScF = ActiveChart.SeriesCollection(1).Formula
  s = StrReverse(ScF)
  s = StrReverse(Mid$(s, 1 + InStr(s, ",")))
  ScF = Mid$(s, 1 + InStr(s, ","))
  Set ScRa = Range(ScF)
  phCondensedSht.Activate
  SpotR1 = ScRa.Row
  SpotR2 = SpotR1 + ScRa.Rows.Count - 1
  phGrafSht.Activate
End If

Set Cht = ActiveSheet.ChartObjects(MassPos)
phCondensedSht.Activate
'SpotNpks = Val(fsExtractPart(Cells(OrigR1 - 2, 1), 3, ","))
Nscans = Val(fsExtractPart(Cells(OrigR1 - 2, 1), 2, ","))
Set CtsR = frSr(OrigR1, SpotCol, OrigR2)
'ChtCap = CtsR(0)
Rct = 0

For i = 1 To Nscans
  If CtsR(i).NumberFormat = fsRejFormat Then Rct = 1 + Rct
Next i

If Rct > 0 Then
  SpotR1 = OrigR2 + 2
  SpotR2 = OrigR1 + Nscans - Rct - 1
Else
  SpotR1 = OrigR1: SpotR2 = OrigR2
End If

Set CtsR = frSr(SpotR1, SpotCol, SpotR2)
Set SecsR = frSr(SpotR1, SpotCol - 1, SpotR2)
Set xyR = Range(SecsR.Address & "," & CtsR.Address)
xyR.Rows.Hidden = False
phGrafSht.Activate
Cht.Select

With ActiveChart
  PlotboxW = .Axes(1).Width
  PlotBoxT = Cht.Top + .Axes(2).Top
  GetSourceInfo .SeriesCollection(1), OrigR, OrigN, NewR, NewN, _
    Rejected(), 0
End With

phGrafSht.Activate
L = Cht.Left + Cht.Width * 0.075
bw = PlotboxW / OrigN: bH = fvMin(16, bw)
L = Cht.Left + Cht.Width * 0.055
wW = Cht.Width / OrigN * 0.8

For i = 1 To OrigN
  MakeRejBoxCell i, Rejected(i), L + (i - 1) * wW, PlotBoxT - bH + 3, wW, bH
  Selection.OnAction = "Rej" & fsS(i)
Next i

Cells(3, 3).Activate
End Sub

Sub RejectPoint(ByVal RejScan%)
' 09/06/09 -- replace line 2 with line 1 to avoid clearing last scan
Dim Rejecting As Boolean, Rejected() As Boolean
Dim s$, i%, p%, Col1%, MassPos%, RejCt%, OrigN%, NewN%
Dim nR1&, nR2&
Dim Xmin#, Ymin#, Xmax#, Ymax#
Dim OrigR As Range, NewR As Range
Dim Ch As ChartObject

NoUpdate
FindChartFromRejbox MassPos
If MassPos = 0 Then Exit Sub
p = ActiveSheet.ChartObjects.Count

For i = 1 To p
  Set Ch = ActiveSheet.ChartObjects(i)
  s = Trim(Right$(Ch.Name, 2))

 If fbIsNumber(s) Then
   If Val(s) = MassPos Then Exit For
  End If

Next i
If i > p Then Exit Sub

Set phGrafSht = ActiveSheet
Ch.Activate
GetSourceInfo ActiveChart.SeriesCollection(1), OrigR, OrigN%, NewR, NewN%, Rejected(), RejCt%
On Error Resume Next

With OrigR(RejScan, 2)
  Rejecting = (.NumberFormat <> fsRejFormat)

  If Rejecting And ((1 + RejCt) / OrigN > (1 / 3)) Or RejCt > 2 Then
    MsgBox "Sorry, you can't reject any more points.", , pscSq
    phGrafSht.Activate:  phGrafSht.Cells(3, 3).Activate
    Exit Sub
  End If

  .NumberFormat = IIf(Rejecting, fsRejFormat, "0;0;0")
  .Font.Bold = Not (.Font.Bold)
End With

On Error GoTo 0
nR1 = OrigR.Row + OrigN + 1
nR2 = nR1 - 1
Col1 = OrigR.Column
1: frSr(nR1 - 1, Col1, nR1 + OrigN - 1, 2 + Col1).ClearContents

For i = 1 To OrigN

  If OrigR(i, 2).NumberFormat <> fsRejFormat Then
    nR2 = 1 + nR2
    Range(OrigR(i, 1), OrigR(i, 3)).Copy Cells(nR2, Col1)
  End If

Next i

Set NewR = frSr(nR1, Col1, nR2, 1 + Col1)
s = NewR.Address
NewR.Rows.Hidden = False
2: 'frSr(nR1 + OrigN + 1, Col1, nR1 + OrigN, 2 + Col1).ClearContents
phGrafSht.Activate: Ch.Activate

With ActiveChart
  .ChartArea.Select
  With .Axes(1)
    Xmin = .MinimumScale: Xmax = .MaximumScale
  End With
  With .Axes(2)
    Ymin = .MinimumScale: Ymax = .MaximumScale
  End With
  .SetSourceData Source:=NewR, PlotBy:=xlColumns
  With .Axes(1)
    .MinimumScale = Xmin: .MaximumScale = Xmax
  End With
  With .Axes(2)
    .MinimumScale = Ymin: .MaximumScale = Ymax
  End With

  If .SeriesCollection.Count = 2 Then
    .SeriesCollection(2).Delete ' Should not have to do this!!!!!!!!!!!
  End If

End With

ActiveWindow.Visible = False ' to take focus off massnum textbox
ActiveSheet.Shapes("rbox" & fsS(RejScan)).Select

With Selection
  With .ShapeRange.Fill
    .ForeColor.RGB = flRejBoxClr(Rejecting)
    .BackColor.RGB = vbWhite
    .TwoColorGradient msoGradientFromCenter, 2
  End With
End With

Cells(3, 3).Activate
End Sub

Sub MakeRejBoxCell(ByVal Num%, ByVal Rejected As Boolean, ByVal L!, ByVal t!, _
  Optional ByVal w! = 20, Optional ByVal h! = 20)

ActiveSheet.Shapes.AddShape(msoShapeRectangle, L, t, w, h).Select
With Selection
  .Characters.Text = fsS(Num)
  On Error Resume Next
  With .Characters(1, 1).Font
    .Size = 0.6 * h
    .Name = "Arial"
  End With
  On Error GoTo 0
  With .ShapeRange
    With .Fill
      .ForeColor.RGB = flRejBoxClr(Rejected)
      .BackColor.RGB = vbWhite
      .TwoColorGradient msoGradientFromCenter, 2
    End With
  End With
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
  .Name = "rbox" & fsS(Num)
End With
End Sub

Sub GetSourceInfo(SC As Object, OrigR As Range, OrigN%, _
  NewR As Range, NewN%, Rejected() As Boolean, RejCt%)

Dim Dsh$, f$, s$, t$
Dim Col1%, Col2%, c%, p%, q%, i%
Dim oR1&, oR2&

f = SC.Formula
On Error Resume Next
ActiveChart.Deselect            ' Otherwise, can't access any cells
ActiveSheet.Cells(1, 1).Select  '  or range objects in the active sheet.
On Error GoTo 0
' 1) =SERIES(,'!JNA_A221'!H8:H13,'!JNA_A221'!I8:I13,1)
' 2) =SERIES(,  JNA_A232 !H69:H70, JNA_A232 !I69:I70,1)
'    =SERIES(,'!_!JNA_A221'!H8:H13,'!_!JNA_A221'!I8:I13,1)
FindStr "Spot#", q, , 1, 1, 999, 1, True

If q > 0 Then FindStr pscQ & " rows ", p, , q, 1, q + 99, 1, True

If q = 0 Or p = 0 Then
  MsgBox "Can't locate the sheet name of these plots.", , pscSq
  End
End If

s = ActiveSheet.Cells(p, 1)
p = InStr(s, pscQ & " rows")
If p < 2 Then End
Dsh = Mid$(s, 2, p - 2)
' 1) JNA_A221'!H8:H13,'!JNA_A221'!I8:I13,1)
' 2) JNA_A232!H69:H70,JNA_A232!I69:I70,1)
Set phCondensedSht = Sheets(Dsh)
phCondensedSht.Activate

p = InStr(f, Dsh) + Len(Dsh)
s = Mid$(f, 1 + p)
If Left$(s, 1) = "!" Then s = Mid$(s, 2)
p = InStr(s, ",")
t = Left$(s, p)
p = InStr(s, Dsh) + Len(Dsh)
s = LTrim(Mid$(s, 1 + p))
If Left$(s, 1) = "!" Then s = LTrim(Mid$(s, 2))
If Left$(s, 1) = "," Then s = LTrim(Mid$(s, 2))
p = InStr(s, ",")
s = Left$(s, p - 1)
s = t & s
' 09/05/18 -- added line below to prevent crash
If Right(s, 1) = ")" Then s = Left(s, Len(s) - 1)
Set NewR = Range(s)
Col1 = NewR.Column: c = Col1 + 1
Col2 = c + 1: oR1 = NewR.Row
NewN = NewR.Rows.Count

Do Until InStr(Cells(oR1 - 2, c), "mass") > 0
  oR1 = oR1 - 1
Loop

oR2 = oR1: RejCt = 0

Do Until Cells(oR2 + 1, c) = ""
  oR2 = oR2 + 1
Loop

Set OrigR = frSr(oR1, Col1, oR2, Col2)
OrigN = OrigR.Rows.Count
ReDim Rejected(1 To OrigN)

For i = 1 To OrigN
  Rejected(i) = (OrigR(i, 2).NumberFormat = fsRejFormat)
  RejCt = RejCt - Rejected(i)
Next i
End Sub

Sub FchartFrag(DataRange$, ByVal CFshape As Boolean, ByVal SBM As Boolean, _
               ByVal ChtCap$, ByVal ChtName$, Cht As Object, ByVal ChtOnAct$)
Dim LoCts As Boolean
Dim s$, ErrBarRangeAddr$, p%, rw1&, rw2&
Dim SigR As Range, DatRange As Range
Dim AL As Axis, Tbx As TextBox, SC As Series, Ach As Chart, ChtObj As ChartObject

NoUpdate
Set DatRange = Range(DataRange)
Set Ach = ActiveChart
Set ChtObj = foLastOb(ActiveSheet.ChartObjects)

If Not SBM Then
  ActiveSheet.Cells(1, 1).Select
  rw1 = DatRange.Row
  rw2 = rw1 + DatRange.Rows.Count - 1
  p = DatRange.Column
  Set SigR = frSr(rw1, p + 2, rw2, , phCondensedSht)
  ErrBarRangeAddr = "=[" & pwDatBk.Name & "]" & phCondensedSht.Name & "!" & SigR.Address
  ChtObj.Activate
End If

With ActiveChart
  Set SC = .SeriesCollection(1)
  With .ChartArea
    .Interior.Color = vbWhite
    .Border.LineStyle = xlNone
  End With
  Set AL = .Axes(xlValue)

  With AL
    .MinimumScale = 0
    LoCts = (.MaximumScale <= 10)
    If LoCts Then .MaximumScale = 10: .MajorUnit = 1
    If SBM Then .MaximumScale = .MaximumScale + .MajorUnit
    On Error GoTo 1
    .TickLabels.Font.Size = 8
    On Error GoTo 0
    .MajorTickMark = xlCross: .MinorTickMark = xlNone: .TickLabelPosition = xlNone
    With .MajorGridlines.Border
      .Color = peLightGray
      .LineStyle = xlContinuous
    End With
  End With

  With .Axes(xlCategory)
    .MajorTickMark = xlCross: .MinorTickMark = xlNone: .TickLabelPosition = xlNone
    .MinimumScale = 0
    .MaximumScale = fvMax(mdMaxX, .MaximumScale)
  End With

  With .PlotArea
    .Left = 0: .Top = 0
    .Interior.Color = 16777164
    .Border.Color = vbBlack
  End With

  Dim LineWt%, SymbSize%, BkClr&, ForeClr&

  FormatSeriesCol SC, , xlContinuous, vbBlack, IIf(SBM, xlHairline, xlThin), _
     IIf(AL.MaximumScale <= 20, False, True), xlCircle, IIf(SBM, 3, 6), vbRed, vbWhite

  If Not SBM Then
    FormatErrorBars SC, 2, SigR, vbRed, xlThin, False
  End If

  On Error GoTo 0
  .Shapes.AddTextbox(1, 1, 1, 1, 1).Select
  Set Tbx = Selection

  With Tbx
    .AutoSize = True
     On Error GoTo 1
    .Font.Size = 9
    On Error GoTo 0
    .Text = Format(AL.MaximumScale, " 0,0")
    .Interior.Color = RGB(200, 255, 200)
    .ShapeRange.Fill.Transparency = 0.5
    With .ShapeRange.TextFrame
      .MarginLeft = 0: .MarginRight = 0
      .MarginTop = 0: .MarginBottom = 0
    End With
    .Top = 0: .Name = "MaxY"
  End With

  ' Label the chart (ScanPks station or SBM) at lower right
  .Shapes.AddTextbox(1, 0, 0, 1, 1).Select
  Set Tbx = Selection

  With Tbx
    .Text = ChtCap: .AutoSize = True
    On Error GoTo 1
    .Font.Size = 13
    On Error GoTo 0
    .Interior.Color = vbWhite
    .ShapeRange.Fill.Visible = False
    .Name = "Nuclide"
  End With

  With .PlotArea
    Tbx.Top = .Top + .Height - Tbx.Height + 2
    Tbx.Left = .Left + .Width - Tbx.Width + 1
  End With

  If CFshape Then
    ActiveSheet.Shapes(ChtName).Delete
    With foLastOb(ActiveSheet.ChartObjects)
      .OnAction = ChtOnAct
      .Name = ChtName & "cht"
    End With
  Else
    ActiveWindow.Visible = False ' otherwise ChtCap may stay selected
    If SBM Then
      PicConvert
    End If
  End If

End With

If SBM Then Exit Sub

If Not CFshape Then
  With ActiveSheet
    Set Cht = foLastOb(ActiveSheet.ChartObjects)
  End With
  With Cht
    .Name = ChtName
    .OnAction = ChtOnAct
  End With
End If

ActiveSheet.Cells(3, 3).Activate
Exit Sub
1: On Error GoTo 0
MsgBox "Sorry, Excel has not allocated enough memory to continue.", , pscSq
End
End Sub

Sub FindChartFromRejbox(ChartNum%)
Dim RejboxBottom!, RejboxLeft!, i%, N%, Chts As ChartObjects

Set Chts = ActiveSheet.ChartObjects: N = Chts.Count

With ActiveSheet.Shapes("rbox1")
  RejboxLeft = .Left
  RejboxBottom = .Top + .Height
End With

For i = 1 To N
  With Chts(i)

    If Abs(RejboxBottom - .Top) < 10 Then

      If Abs(RejboxLeft - .Left) < 30 Then
        ChartNum = i: Exit Sub
      End If

    End If

  End With
Next i

End Sub
