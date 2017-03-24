Attribute VB_Name = "RawDataFiles"
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
' Module RawDataFiles
' 09/03/02 -- All lower array bounds explicit
' 09/05/11 -- Add the TooManyMassStations sub, call from ParsePD & ParseXML subs.
Option Explicit
Option Base 1

Sub InhaleRawdata(ByVal IsPDfile As Boolean, FileLine$(), Bad As Boolean)
Dim GotPks As Boolean, GotDupe As Boolean, tB As Boolean
Dim t$, ti$, ei$, mi$, si$, ai$, Typ$, s1$, s2$, Trans$, Dup$, DupeN%, DupeStr$, Ndupe$, BadMsg$
Dim p%, j%, Np%, Lai%, Lsi%, Lti%, Lmi%, Nsp%, PkNum%, PkCt%, StartPk%
Dim i&, Lin&, NameRw&, r&, NL&
Dim Outp As Variant
Const PNA = "<par name=", Fcup = " (Faraday cup used instead of counter)"

ReDim psaSpotNames(1 To 999), plaSpotNameRowsRaw(1 To 999), piaFileNscans(1 To 999)
Nsp = 0: piFileNpks = 0
StartPk = 1: GotPks = False

pbXMLfile = Not IsPDfile
pbPDfile = Not pbXMLfile
pbAlwaysShort = Not foUser("LongCondensed")

If pbXMLfile Then
  ti = PNA & fsInQ("title") & " value=" & pscQ
  mi = PNA & fsInQ("measurements") & " value="
  ai = PNA & fsInQ("amu") & " value="
  ei = pscQ & " />"
  si = PNA & fsInQ("scans") & " value="
  Lsi = Len(si) + 2: Lai = Len(ai) + 2
  Lti = Len(ti) + 1: Lmi = Len(mi) + 2
End If

pbShortCondensed = pbAlwaysShort

On Error Resume Next
Close #1
On Error GoTo 0
NoUpdate False
t = IIf(pbXMLfile, "   (please be patient)", "")
StatBar "Loading " & fsExtrRight(psRawFileName, "\", True) & t
Open psRawFileName For Input As #1

Lin = 0: Bad = False
Typ = IIf(IsPDfile, "PD", "XML")
Trans = "Transferring " & Typ & " file"
On Error GoTo BadFile

If pbXMLfile Then
  Line Input #1, t
  ParseLine t, Outp, , vbLf
  NL = UBound(Outp)
  ReDim FileLine(1 To NL)

  For i = 1 To NL

    If Outp(i) <> "" Then
      Lin = 1 + Lin
      FileLine(Lin) = CleanLine(Outp(i))
    End If

  Next i

Else
  ReDim FileLine(1 To 2 ^ 24)

  Do While Not EOF(1)
    Line Input #1, t

    If t <> "" Then
      Lin = Lin + 1

      If Lin Mod 1000 = 0 Then
        StatBar Trans, Lin
      End If

      FileLine(Lin) = CleanLine(t)
    End If

  Loop

End If
NL = Lin
ReDim Preserve FileLine(1 To NL)

For r = 1 To NL

    t = FileLine(r)

    If r Mod 1000 = 0 Then
      StatBar "Wait", r
    End If

    If IsPDfile Then

       If r > 1 Then

        If FileLine(r - 1) = "***" Then
          Nsp = 1 + Nsp
          p = InStr(t, ",")
          plaSpotNameRowsRaw(Nsp) = r
          psaSpotNames(Nsp) = Left$(t, p - 1)
          GotPks = False
        End If

        If Nsp > 0 And piaFileNscans(Nsp - (Nsp = 0)) = 0 Then

          If Left$(t, 4) = "set " Then
            ai = Mid$(t, InStr(t, ",") + 1)
            piaFileNscans(Nsp) = Val(ai)

            If piFileNpks = 0 Then
              mi = Mid$(t, InStr(t, " scans,") + 8)
              piFileNpks = Val(mi)
              ReDim pdaFileMass(1 To piFileNpks)
            End If

          End If

        ElseIf NameRw = 0 Then
          If Left$(t, 5) = "Name " Then NameRw = r

        ElseIf Left$(t, 5) <> "Name " And InStr(t, " CUP ") > 10 Then
          ParseLine t, Outp, Np, " "

          For j = 1 To Np ' All pks must be on counter
            If j > 8 And Outp(j) = "CUP" Then
              Nsp = Nsp - 1
              BadMsg = "Ignoring Spot " & psaSpotNames(Nsp) & Fcup
              If MsgBox(BadMsg, vbOKCancel, pscSq) = vbCancel Then CrashEnd
              Nsp = Nsp - 1
              GoTo Nextr
            End If
          Next j

        Else

          For PkCt = StartPk To piFileNpks

            If pdaFileMass(PkCt) = 0 Then

              If r = (NameRw + PkCt) Then
                pdaFileMass(PkCt) = Val(Mid$(t, 11))
                If PkCt < piFileNpks Then StartPk = PkCt Else GotPks = True
                Exit For
              End If

            End If

          Next PkCt

        End If

      End If ' r>1

    Else ' is an XML file
      If InStr(t, ti) > 0 Then            ' Spot name
        p = InStr(t, ei)

        If p > 0 Then
          Nsp = 1 + Nsp
          plaSpotNameRowsRaw(Nsp) = r
          psaSpotNames(Nsp) = Trim(Mid$(t, Lti, p - Lti))
          GotPks = False
        End If

      ElseIf piFileNpks = 0 Then          ' #peaks

        If InStr(t, mi) > 0 Then
          piFileNpks = Val(Mid$(t, Lmi))
          ReDim pdaFileMass(1 To piFileNpks)
        End If

      ElseIf InStr(LCase(t), "sc_detector") > 0 Then

        tB = Fparam(1, t, "sc_detector", , s1)
        If tB And s1 <> "counter" Then
          BadMsg = "Ignoring Spot " & psaSpotNames(Nsp) & Fcup
          If MsgBox(BadMsg, vbOKCancel, pscSq) = vbCancel Then CrashEnd
          Nsp = Nsp - 1
          GoTo Nextr
        End If

      ElseIf piFileNpks > 0 Then          ' Pk masses

        For PkCt = StartPk To piFileNpks

          If pdaFileMass(PkCt) = 0 Then

            If InStr(t, ai) > 0 Then
              pdaFileMass(PkCt) = Val(Mid$(t, Lai))
              If PkCt < piFileNpks Then StartPk = PkCt Else GotPks = True
              Exit For
            End If

          End If

        Next PkCt

      End If

      If InStr(t, si) > 0 Then ' #scans
        piaFileNscans(Nsp) = Val(Mid$(t, Lsi))
      End If

    End If 'pd or xml

Nextr:
Next r

If NL > (1 + pemaxrow) Then pbShortCondensed = True

If Nsp = 0 Or NL = 0 Then GoTo BadFile
On Error GoTo 0
plFileNlines = NL
piNumAllSpots = Nsp
StatBar
ReDim Preserve FileLine(1 To plFileNlines)
ReDim Preserve psaSpotNames(1 To Nsp), plaSpotNameRowsRaw(1 To Nsp), piaFileNscans(1 To Nsp)
ReDim SortedNames$(1 To Nsp), Indx%(1 To Nsp)

' Look for duplicate names, rename "...dup1, ...dup2", etc. if found
For i = 1 To Nsp
    SortedNames(i) = psaSpotNames(i)
    Indx(i) = i
Next i

BubbleSort SortedNames, Indx, , True
Dup = Chr(133) & "dup"

Do
  GotDupe = False

  For i = Nsp To 2 Step -1
    s1 = SortedNames(i)
    s2 = SortedNames(i - 1)

    If s1 = s2 Then
      p = InStr(s1, Dup)
      Ndupe = Right$(s1, 1)

      If p = 0 Or Not IsNumeric(Ndupe) Then
        DupeStr = Dup & "1"
      Else
        DupeN = 1 + fvMax(1, Val(Ndupe))
        p = InStr(s2, Dup)
        If p > 0 Then s2 = Left$(s2, p - 1)
        DupeStr = Dup & fsS(DupeN)
      End If

      SortedNames(i) = s2 & DupeStr
      psaSpotNames(Indx(i)) = SortedNames(i)
      'If Workbooks.Count > 0 Then Cells(i, 1) = SortedNames(i)
      GotDupe = True
    End If

  Next i

Loop Until Not GotDupe

Exit Sub

BadFile: On Error GoTo 0
MsgBox "Error in parsing " & psRawFileName
Bad = True
End Sub

Sub ParsePD(RawDat() As RawData, PDm$(), Nlines&, Bad As Boolean, BadMsg$, _
  Optional Spot1Only As Boolean = False)
' 09/04/01 -- Inhibit re-parsing of the Spot Names (InhaleRawData names & dupe-corrs retained)
Dim xs$, tmp$, tB As Boolean
Dim SpotNum%, ScanNum%, Nscans%, Nf%, PkNum%, Npks%, Nspots%
Dim k%, m%, Nsp%, Nsc%, p1%, p2%, p3%, p4%, spN%, q%, Col%
Dim j&, i&, b%, r&
Dim SbmTotCts#, PkTotCts#, CorrCts#, PkSigmaMean#, SigmaSbmTotCts#
Dim Cps#, XkTotCts#, SigmaPkTotCts#
Dim PkCts#(1 To 10), SBMcts#(1 To 10), Fields As Variant

Nsp = 0

If piNumAllSpots = 0 Then
  ' An already-condensed PD file must be open when this sub is called
  For i = 1 To Nlines

    If PDm(i) = "***" Then
      Nsp = 1 + Nsp
      ReDim Preserve plaSpotNameRowsCond(1 To Nsp)
      plaSpotNameRowsCond(Nsp) = i + 1
    End If

  Next i

  piNumAllSpots = Nsp
Else
  Nsp = piNumAllSpots
End If

ReDim RawDat(1 To Nsp), psaSpotDateTime(1 To Nsp)

For SpotNum = 1 To IIf(Spot1Only, 1, piNumAllSpots)
  With RawDat(SpotNum)
    spN = SpotNum

    If spN Mod 2 = 0 Then
      q = 2 * ((Nsp - spN) \ 2)
      StatBar "Parsing PD file", q
    End If

    r = plaSpotNameRowsRaw(spN)
    ParseLine CleanLine(PDm(r)), Fields, Nf, ","
    'psaSpotNames(spN) = Fields(1)
    .saSpotName = psaSpotNames(spN) ' Fields(1)
    .sDate = Fields(2)
    .sTimeOfDay = Fields(3)
    psaSpotDateTime(spN) = .sDate & "," & .sTimeOfDay

    Do
      r = r + 1
      If r > Nlines Then b = 1: GoTo Bad
      xs = CleanLine(PDm(r))
    Loop Until Left$(xs, 4) = "set "

    ParseLine xs, Fields, Nf, ","
    .iNscans = Val(Fields(2))
    Nsc = .iNscans
    .iNpeaks = Val(Fields(3))
    Npks = .iNpeaks
    If Npks > MinOrMax.peMaxNukes Then Call TooManyMassStations("PD", Npks)

    If SpotNum = 1 Then
      piFileNpks = Npks
      ReDim pdaFileMass(1 To Npks)
    End If

    .dDeadTime = Val(Fields(4)) / pdcBillion
    .lSBMzero = Val(Mid$(Fields(5), Len("sbm zero ") + 1))

    ReDim .saNukeLabels(1 To Npks), .daTrueMass(1 To Npks)
    ReDim .daTrimMass(1 To Npks, 1 To Nsc), .daIntegrTimes(1 To Npks), .daWaitTimes(1 To Npks)
    ReDim .saDetector(1 To Npks), .daTimeStamp(1 To Npks, 1 To Nsc), .daPkCts(1 To Npks, 1 To Nsc)
    ReDim .daSBMcts(1 To Npks, 1 To Nsc), .daPkSigmaMean(1 To Npks, 1 To Nsc), .baCenteredPeak(1 To Npks)
    If SpotNum = 1 Then ReDim pbCenteredPk(1 To Npks)

    Do
      r = r + 1
      If r > Nlines Then b = 2: GoTo Bad
      xs = LCase(PDm(r))
      p1 = InStr(xs, "name")
      p2 = InStr(xs, "amu offset")
      p3 = InStr(xs, "time s")
      p4 = InStr(xs, "delay s")
    Loop Until p1 = 1 And p2 > 5 And p3 > p2 And p4 > p3

    For PkNum = 1 To .iNpeaks
      Do
        r = r + 1
        If r > Nlines Then b = 3: GoTo Bad
        xs = PDm(r)
        tmp = Left$(xs, 11)
        Subst tmp, " "
        If Len(tmp) < Len(xs) Then xs = tmp & " " & Mid$(xs, 12)
        ParseLine xs, Fields, Nf, " "
      Loop Until Nf >= 11

      .saNukeLabels(PkNum) = Fields(1)
      .daTrueMass(PkNum) = Fields(2)
      If SpotNum = 1 Then pdaFileMass(PkNum) = Fields(2)
      .daIntegrTimes(PkNum) = Fields(4)
      .daWaitTimes(PkNum) = Fields(5)
      .baCenteredPeak(PkNum) = (Fields(7) = "YES")
      If SpotNum = 1 Then pbCenteredPk(PkNum) = .baCenteredPeak(PkNum)
      .saDetector(PkNum) = LCase(Fields(11))

      If .saDetector(PkNum) <> "counter" Then
        BadMsg = "Error -SQUID failed to intercept a Faraday Cup spot."
        MsgBox BadMsg, vbOKOnly, pscSq
        CrashEnd
      End If


    Next PkNum

    Do
      r = r + 1
      If r > Nlines Then b = 5: GoTo Bad
      xs = Left$(CleanLine(PDm(r)), 18)
    Loop Until xs = "AMU-B calibration "

    For ScanNum = 1 To Nsc

     For i = 1 To 2 * .iNpeaks Step 2

      Do
        r = r + 1
        If r > Nlines Then b = 6: GoTo Bad
        PkNum = (i + 1) \ 2
        ParseLine PDm(r), Fields, Nf, " "
      Loop Until Nf = 14

      .daTrimMass(PkNum, ScanNum) = Fields(2)
      .daTimeStamp(PkNum, ScanNum) = Fields(3)

      For m = 1 To 10
        PkCts(m) = Fields(4 + m)
      Next m

      r = r + 1
      ParseLine PDm(r), Fields, Nf, " "
      If Nf <> 10 Then b = 7: GoTo Bad

      For m = 1 To 10
        SBMcts(m) = Fields(m)
      Next m

      Nf = 10

      PoissonOutliers 10, PkCts, SBMcts, PkTotCts, SigmaPkTotCts, _
              SbmTotCts, SigmaSbmTotCts, Nf, .daIntegrTimes(PkNum)

      Cps = PkTotCts / .daIntegrTimes(PkNum)
      CorrCts = .daIntegrTimes(PkNum) * fdDeadTimeCorrCPS(Cps, .dDeadTime) ' Total cts, dead-time corr
      .daPkCts(PkNum, ScanNum) = CorrCts
      .daPkSigmaMean(PkNum, ScanNum) = SigmaPkTotCts
      .daSBMcts(PkNum, ScanNum) = SbmTotCts
    Next i

  Next ScanNum

  End With
NextSpotNum:
Next SpotNum

StatBar
Exit Sub

Bad:
On Error GoTo 0

If SpotNum = Nsp Then
  piNumAllSpots = piNumAllSpots - 1
Else
  MsgBox "Corrupt or otherwise unparseable PD file" & vbLf & StR(b), , pscSq
  End
End If

End Sub

Sub ParseShortCondensed(Spot1Only As Boolean, RawDat() As RawData)
Dim GotSht As Boolean, tmp$, PkNum%, SpotNum%
Dim ScanNum%, spN%, Nsp%, Nsc%, Npks%, Col%
Dim r&, q&, Rw&, m&


FindCondensedSheet GotSht

If Not GotSht Then
  MsgBox "Can't find raw-data worksheet in active workbook.", , pscSq
  End
End If

Set pwDatBk = ActiveWorkbook
Set phCondensedSht = ActiveSheet

Nsp = Val(Cells(4, picDatCol))
piNumAllSpots = Nsp
ReDim plaSpotNameRowsCond(1 To Nsp)

For r = 1 To Nsp
  plaSpotNameRowsCond(r) = Cells(r + 1, 2)
Next r

ReDim RawDat(1 To Nsp), psaSpotDateTime(1 To Nsp), psaSpotNames(1 To Nsp)

For SpotNum = 1 To IIf(Spot1Only, 1, piNumAllSpots)
  With RawDat(SpotNum)
    spN = SpotNum
    q = Nsp - spN + 1
    StatBar "Parsing PD file", q
    r = plaSpotNameRowsCond(spN)
    GetNameDatePksScans r, .saSpotName, , Npks, Nsc, .sDate, .sTimeOfDay
    .iNpeaks = Npks: .iNscans = Nsc

    If SpotNum = 1 Then
      piFileNpks = Npks
      ReDim pbCenteredPk(1 To Npks)
    End If

    ReDim pdaFileMass(1 To Npks), .saNukeLabels(1 To Npks), .daTrueMass(1 To Npks)
    ReDim .daTrimMass(1 To Npks, 1 To Nsc), .daIntegrTimes(1 To Npks), .daWaitTimes(1 To Npks)
    ReDim .saDetector(1 To Npks), .daTimeStamp(1 To Npks, 1 To Nsc), .daPkCts(1 To Npks, 1 To Nsc)
    ReDim .daSBMcts(1 To Npks, 1 To Nsc), .daPkSigmaMean(1 To Npks, 1 To Nsc), .baCenteredPeak(1 To Npks)

    psaSpotNames(spN) = .saSpotName
    psaSpotDateTime(spN) = .sDate & "," & .sTimeOfDay
    r = r + picDatRowOffs - 1 ' last header-row, just above 1st scan-data

    For PkNum = 1 To .iNpeaks
      Col = picDatCol + 5 * (PkNum - 1)
      .daIntegrTimes(PkNum) = Cells(r - 1, Col)
      .daTrueMass(PkNum) = Cells(r, 1 + Col)
      If SpotNum = 1 Then pdaFileMass(PkNum) = .daTrueMass(PkNum)

      For ScanNum = 1 To Nsc
         m = r + ScanNum
        .daTimeStamp(PkNum, ScanNum) = Cells(m, Col)
        .daPkCts(PkNum, ScanNum) = Cells(m, 1 + Col)
        .daPkSigmaMean(PkNum, ScanNum) = Cells(m, 2 + Col)
        .daSBMcts(PkNum, ScanNum) = Cells(m, 3 + Col)
        .daTrimMass(PkNum, ScanNum) = Cells(m, 4 + Col)
      Next ScanNum

    Next PkNum

    Col = picDatCol + 5 * Npks
    r = r + 1
    .dDeadTime = Cells(r, Col)
    .lSBMzero = Cells(r, 1 + Col)

    If pbXMLfile Then
      .dStageX = Cells(r, 2 + Col)
      .dStageY = Cells(r, 3 + Col)
      .dStageZ = Cells(r, 4 + Col)
      .dQt1y = Cells(r, 5 + Col)
      .dQt1z = Cells(r, 6 + Col)
      .dPrimaryBeam = Cells(r, 7 + Col)
    End If

  End With
Next SpotNum

End Sub

Function Fparam(ByVal Par1Dat2%, ByVal Fs$, ByVal ParName$, Optional vParam, _
  Optional sParam, Optional Simple As Boolean = False, Optional ByVal StartDelim, _
  Optional ByVal EndDelim) As Boolean
Dim t$, s$, p%, Le%, Tx$
Fparam = False

If Simple Then
  Le = Len(ParName)

  If Left$(Fs, 2 + Le) = "<" & ParName & ">" Then

    If fbNIM(vParam) Then
      vParam = 0
    ElseIf fbNIM(sParam) Then
      sParam = ""
    End If

    Fparam = True
  End If

Else
  If fbNIM(StartDelim) Then
    p = InStr(Fs, StartDelim)
    If p = 0 Then Exit Function
    Fs = Mid$(Fs, p + Len(StartDelim))
  End If

  If fbNIM(EndDelim) Then
    p = InStr(Fs, EndDelim)
    If p = 0 Then Exit Function
    Fs = Left$(Fs, p - Len(EndDelim))
  End If

  If fbNIM(StartDelim) Or fbNIM(EndDelim) Then
    If fbNIM(vParam) Then
      vParam = Val(Fs)
    Else
      sParam = Fs
    End If
    Fparam = True
    Exit Function
  End If

  t = "<" & Choose(Par1Dat2, "par", "data") & " name=" & fsInQ(ParName)
  If Par1Dat2 = 1 Then
    t = t & " value=" & pscQ
  Else
    t = t & ">"
  End If

  Le = Len(t)

  If Left$(Fs, Le) = t Then
    Tx = Mid$(Fs, 1 + Le)

    If fbNIM(vParam) Then
      vParam = Val(Tx)
      Fparam = True
    Else
      p = InStr(Tx, pscQ & " />")
      If p > 0 Then Tx = Left$(Tx, p - 1)
      sParam = Tx
      Fparam = True
    End If

  End If

End If
End Function

Function CleanLine(ByVal FileLine$) As String
Dim Le%
FileLine = Replace(FileLine, pscQ & pscQ, pscQ)
Le = Len(FileLine)

If Left$(FileLine, 1) = pscQ Then
  FileLine = Mid$(FileLine, 2, Le - 2)
End If

CleanLine = FileLine
End Function

Sub ParseXML(RawDat() As RawData, Xm$(), Nlines&, _
  Bad As Boolean, BadMsg$, Optional Spot1Only As Boolean = False)
' 09/04/01 -- Inhibit re-parsing of the Spot Names (InhaleRawData names & dupe-corrs retained)
Dim tB As Boolean, xs$, Tx$
Dim SpotNum%, ScanNum%, Nscans%, Nf%, i%, p%, PkNum%, Npks%, Nspots%, spN%, q%
Dim j&, Nrw&, PeakRow&, RunsRow&, SetRow&
Dim SigmaSbmTotCts#, SbmTotCts#, PkTotCts#, SigmaPkTotCts#, Cps#, CorrCts#, PkSigmaMean#
Dim PkCts#(1 To 10), SBMcts#(1 To 10)
Dim CountArray() As Variant, SBMarray() As Variant, v As Variant

j = 0
Do
  j = j + 1
  xs = Xm(j)

  If Fparam(1, xs, "", , Tx, , "<!", ">") Then
    psXmlFileType = Tx
  ElseIf Fparam(1, xs, "software_version", , Tx, True) Then
    Tx = Mid$(xs, 19)
    p = InStr(Tx, "</")
    Tx = Left$(Tx, p - 1)
    psRawdatSoftwareVer = Tx
  ElseIf Fparam(1, xs, "runs", v, , True) Then
    piNumAllSpots = Val(Mid$(xs, 7))
  End If

Loop Until (psXmlFileType <> "" And psRawdatSoftwareVer <> "" And piNumAllSpots > 0) Or j = Nlines

ReDim RawDat(1 To piNumAllSpots), psaSpotDateTime(1 To piNumAllSpots), _
      plaSpotNameRowsCond(1 To piNumAllSpots)

For SpotNum = 1 To IIf(Spot1Only, 1, piNumAllSpots)
  q = piNumAllSpots - spN
  StatBar "Parsing XML file", q

  Do
    j = j + 1
  Loop Until Xm(j) = "<run>"

  spN = SpotNum
  With RawDat(spN)
    .dStageX = -1:   .dStageY = -1:   .dStageZ = -1
    .dDeadTime = -1: .lSBMzero = -1

    Do
      j = j + 1
      xs = Xm(j)
      tB = False

      If .saSpotName = "" Then
        tB = Fparam(1, xs, "title", , .saSpotName)
        If tB Then plaSpotNameRowsRaw(spN) = j
      End If

      If Not tB And .iNpeaks = 0 Then tB = Fparam(1, xs, "measurements", .iNpeaks)
      If Not tB And .iNscans = 0 Then tB = Fparam(1, xs, "scans", .iNscans)

      If Not tB And .dDeadTime < 0 Then
        tB = Fparam(1, xs, "dead_time_ns", v)
        If tB Then .dDeadTime = v / pdcBillion
      End If

      If Not tB And .lSBMzero < 0 Then tB = Fparam(1, xs, "sbm_zero_cps", .lSBMzero)
      If Not tB And .dStageX < 0 Then tB = Fparam(1, xs, "stage_x", .dStageX)
      If Not tB And .dStageY < 0 Then tB = Fparam(1, xs, "stage_y", .dStageY)
      If Not tB And .dStageZ < 0 Then tB = Fparam(1, xs, "stage_z", .dStageZ)

    Loop Until xs = "<entry>"

    Nscans = .iNscans
    Npks = .iNpeaks
    If Npks > MinOrMax.peMaxNukes Then Call TooManyMassStations("XML", Npks)

    ReDim .saNukeLabels(1 To Npks), .daTrueMass(1 To Npks)
    ReDim .daTrimMass(1 To Npks, 1 To Nscans), .daIntegrTimes(1 To Npks), .daWaitTimes(1 To Npks)
    ReDim .saDetector(1 To Npks), .daTimeStamp(1 To Npks, 1 To Nscans), .daPkCts(1 To Npks, 1 To Nscans)
    ReDim .daSBMcts(1 To Npks, 1 To Nscans), .daPkSigmaMean(1 To Npks, 1 To Nscans), .baCenteredPeak(1 To Npks)
    If SpotNum = 1 Then ReDim pbCenteredPk(1 To Npks)

    For PkNum = 1 To Npks

      Do
        tB = False
        j = j + 1
        xs = Xm(j)
        If .saNukeLabels(PkNum) = "" Then tB = Fparam(1, xs, "label", , .saNukeLabels(PkNum))
        If Not tB And .daTrueMass(PkNum) = 0 Then tB = Fparam(1, xs, "amu", .daTrueMass(PkNum))

        If Not tB Then
          tB = Fparam(1, xs, "centering_time_sec", v)

          If tB Then
            .baCenteredPeak(PkNum) = (v > 0)
            If SpotNum = 1 Then pbCenteredPk(PkNum) = (v > 0)
          End If

        End If

        If Not tB Then tB = Fparam(1, xs, "count_time_sec", .daIntegrTimes(PkNum))
        If Not tB Then tB = Fparam(1, xs, "delay_sec", .daWaitTimes(PkNum))

        If Not tB And .saDetector(PkNum) = "" Then
          tB = Fparam(1, xs, "sc_detector", , .saDetector(PkNum))

          If tB And .saDetector(PkNum) <> "counter" Then
            BadMsg = "Error -SQUID failed to intercept a Faraday Cup spot."
            MsgBox BadMsg, vbOKOnly, pscSq
            CrashEnd
          End If

        End If

      Loop Until xs = "</entry>" 'Or Xs = "</run_table>"

    Next PkNum

    Do
      j = j + 1
    Loop Until Xm(j) = "</run_table>"

    Do
      j = j + 1
    Loop Until Xm(j) = "<set>"

    .dQt1y = -1: .dQt1z = -1: .dPrimaryBeam = 9999

    Do
      j = j + 1
      xs = Xm(j)
      tB = False
      tB = Fparam(1, xs, "date", , .sDate)
      If Not tB And .sTimeOfDay = "" Then tB = Fparam(1, xs, "time", , .sTimeOfDay)
      If Not tB And .dQt1y < 0 Then tB = Fparam(1, xs, "qt1y", .dQt1y)
      If Not tB And .dQt1z < 0 Then tB = Fparam(1, xs, "qt1z", .dQt1z)
      If Not tB And .dPrimaryBeam = 9999 Then tB = Fparam(1, xs, "pbm", .dPrimaryBeam)
    Loop Until xs = "<scan number=" & fsInQ("1") & ">"

    psaSpotDateTime(spN) = .sDate & ", " & .sTimeOfDay

    For ScanNum = 1 To Nscans

      Do While Xm(j) <> "<scan number=" & fsInQ(fsS(ScanNum)) & ">"
        j = j + 1
      Loop

      For PkNum = 1 To Npks

        Do
          tB = False
          j = j + 1
          xs = Xm(j)

          If .daTrimMass(PkNum, ScanNum) = 0 Then
            tB = Fparam(1, xs, "trim_mass", .daTrimMass(PkNum, ScanNum))
          End If

          If Not tB And .daTimeStamp(PkNum, ScanNum) = 0 Then
            tB = Fparam(1, xs, "time_stamp_sec", .daTimeStamp(PkNum, ScanNum))
          End If

          If Not tB Then
            tB = Fparam(2, xs, .saNukeLabels(PkNum), 0)

            If tB Then
              p = InStr(xs, ">")
              Tx = Mid$(xs, p + 1)
              p = InStr(Tx, "</")
              Tx = Left$(Tx, p - 1)
              ParseLine Tx, CountArray, Nf, ","

              For i = 1 To 10
                PkCts(i) = CountArray(i)
              Next i

            End If

          End If

          If Not tB Then
            tB = Fparam(2, xs, "SBM", 0)

            If tB Then
              p = InStr(xs, ">")
              Tx = Mid$(xs, p + 1)
              p = InStr(Tx, "</")
              Tx = Left$(Tx, p - 1)
              ParseLine Tx, SBMarray, Nf, ","

              For i = 1 To 10
                SBMcts(i) = SBMarray(i)
              Next i

            End If

          End If

        Loop Until xs = "</measurement>"
        Nf = 10

        PoissonOutliers 10, PkCts, SBMcts, PkTotCts, SigmaPkTotCts, _
                  SbmTotCts, SigmaSbmTotCts, Nf, .daIntegrTimes(PkNum)

        Cps = PkTotCts / .daIntegrTimes(PkNum)
        CorrCts = .daIntegrTimes(PkNum) * fdDeadTimeCorrCPS(Cps, .dDeadTime) ' Total cts, dead-time corr
        .daPkCts(PkNum, ScanNum) = CorrCts
        .daPkSigmaMean(PkNum, ScanNum) = SigmaPkTotCts
        .daSBMcts(PkNum, ScanNum) = SbmTotCts
      Next PkNum

      Do
        j = 1 + j
      Loop Until Xm(j) = "</scan>"

    Next ScanNum

    Do
      j = 1 + j
    Loop Until Xm(j) = "</run>"

  End With
NextSpotNum: Next SpotNum

StatBar
Exit Sub

Bad:
MsgBox BadMsg, , pscSq
End
End Sub

Sub CondenseRawData(ByVal Npks%, PkMass#(), ForScanGrafix As Boolean, RawDat() As RawData, _
  FileLines$(), Optional GetStdNames As Boolean = False, Optional Bad As Boolean)

Dim NoMatchOK As Boolean
Dim s1$, s2$, tmp$, tmp1$, tmp2$, NumForm$, WbkNa$, BadMsg$, FL$(1 To 3)
Dim k%, Col%, SbmCt%, TimeCol%, EndCol%, Msg%, PkNum%, ScanNum%, SpotCt%, Npks0%
Dim Na1rw&, r&, rw1&, rw2&, Clr1&, Clr2&, SpotR&, FirstDatRw&, LastHdrRw&
Dim NameColWidth!, ScansPksColWidth!
Dim IntT#()
Dim HdrDat() As Variant

Clr2 = RGB(128, 0, 0): Clr1 = RGB(0, 0, 128)
If fbNIM(Bad) Then Bad = False

pdDeadTimeSecs = 0 ' Les Sullivan; overrides global pdDeadTimeSecs left behind from previous 'purple zircon' run
' Ensures that CondensedRawData "tot cts mass" values are repeatable irrespective of the sequence of 'blue rectangle'
' and purple zircon' data-reduction processes executed.

Do

  If Npks = 0 Or Workbooks.Count = 0 Then
    InhaleRawdata pbPDfile, FileLines, Bad
    If Bad Then End

    If pbXMLfile Then
      ParseXML RawDat, FileLines, plFileNlines, Bad, BadMsg
    Else
      ParsePD RawDat, FileLines, plFileNlines, Bad, BadMsg
    End If

    If Bad Then Exit Sub
    Npks = RawDat(1).iNpeaks
  Else

    plFileNlines = 0
    On Error Resume Next
    plFileNlines = UBound(FileLines)
    On Error GoTo 0
    Npks = IIf(plFileNlines = 0, 0, RawDat(1).iNpeaks)
  End If

Loop Until Npks > 0

With RawDat(1)
  piaFileNscans(1) = .iNscans
  ReDim PkMass(1 To Npks)

  For PkNum = 1 To Npks
    PkMass(PkNum) = .daTrueMass(PkNum)
  Next PkNum

End With

ReDim HdrDat(1 To 3, 1 To 4 + 5 * Npks - 6 * pbXMLfile)

If Not pbShortCondensed Then StatBar "Archiving " & IIf(pbXMLfile, "XML", "PD") & _
  "file lines"
Workbooks.Add
tmp = fsExtrRight(psRawFileName, "\", True)
Sheets(1).Name = Left$(fsLegalName(tmp, True), 31)
Set phCondensedSht = ActiveSheet
Set pwDatBk = ActiveWorkbook
NoGridlines
NoUpdate
Cells.Font.Size = 9
Columns(1).Font.Name = "Lucida console"

For r = 1 To 3
  FL(r) = FileLines(r)
Next r

If Not pbShortCondensed Then
  For r = 1 To plFileNlines
    Cells(1 + r, 1) = FileLines(r)
  Next r
End If

Erase FileLines

Npks0 = Npks
foAp.EnableCancelKey = xlInterrupt
NoMatchOK = False
WbkNa = ActiveWorkbook.Name

ReDim plaSpotNameRowsCond(1 To piNumAllSpots)
Zoom 100
Blink

For SpotCt = 1 To piNumAllSpots
  tmp = StR(piNumAllSpots - SpotCt)
  StatBar "Reformatting  " & WbkNa & "  " & tmp

  If pbShortCondensed Then

    If SpotCt = 1 Then
      SpotR = 6
    Else
      SpotR = plaSpotNameRowsCond(SpotCt - 1) + _
              picDatRowOffs + 2 * piaFileNscans(SpotCt - 1) + 1
    End If

  Else
    SpotR = fvMax(6, plaSpotNameRowsRaw(SpotCt))
  End If

  plaSpotNameRowsCond(SpotCt) = SpotR
  FirstDatRw = SpotR + picDatRowOffs
  LastHdrRw = FirstDatRw - 1
  Npks = RawDat(SpotCt).iNpeaks
  piNscans = RawDat(SpotCt).iNscans

  If Npks > peMaxNukes Then

    If MsgBox("The Run Table for Spot " & fsS(SpotCt) & _
      " uses " & Npks & " mass stations; SQUID is limited to " & _
      fsS(peMaxNukes) & ".", vbOKCancel, pscSq) = vbCancel _
      Then CrashEnd
    GoTo NextSpotCt
  ElseIf Npks <> Npks0 And Not NoMatchOK And Not ForScanGrafix Then
    Msg = MsgBox("The first analyzed spot has" & StR(Npks0) & _
          " mass stations, but spot#" & StR(SpotCt) & _
          " has" & StR(Npks) & ".  Continue anyway?", vbYesNo, pscSq)
    If Msg = vbYes Then NoMatchOK = True
      If Not NoMatchOK Then End
  End If

  ReDim pdaPkCts(1 To Npks, piNscans), pdaPkT(1 To Npks, 1 To piNscans), pdaSBMcps(1 To Npks, 1 To piNscans)
  ReDim PkDat(1 To piNscans, 5 * Npks), SbmDat(1 To piNscans * Npks, 1 To 2)
  ReDim IntT(1 To Npks)

  For ScanNum = 1 To piNscans

    For PkNum = 1 To Npks
      TimeCol = 1 + 5 * (PkNum - 1) ' PkDat col 1
      With RawDat(SpotCt)
        IntT(PkNum) = .daIntegrTimes(PkNum)
        PkDat(ScanNum, TimeCol) = .daTimeStamp(PkNum, ScanNum) ' time of integration
        PkDat(ScanNum, 4 + TimeCol) = .daTrimMass(PkNum, ScanNum)
      End With

      If ScanNum = 1 Then ' Make header row
        HdrDat(3, TimeCol) = "Secs"
        HdrDat(3, 1 + TimeCol) = Drnd(RawDat(SpotCt).daTrueMass(PkNum), 6)
        HdrDat(3, 2 + TimeCol) = "�1sig":    HdrDat(2, 3 + TimeCol) = "tot cts"
        HdrDat(1, 1 + TimeCol) = "tot cts":  HdrDat(2, TimeCol) = IntT(PkNum)
        HdrDat(2, 1 + TimeCol) = "mass":     HdrDat(3, 3 + TimeCol) = "SBM"
        HdrDat(3, 4 + TimeCol) = "mass":     HdrDat(2, 4 + TimeCol) = "trim"

        If PkNum = Npks Then
          Col = 5 * Npks + 1
          HdrDat(3, Col) = "ns"
          HdrDat(2, Col) = "Dead time"
          HdrDat(3, Col + 1) = "sbm zero"

          If pbXMLfile Then
            HdrDat(3, Col + 2) = "Stage X"
            HdrDat(3, Col + 3) = "Stage Y"
            HdrDat(3, Col + 4) = "Stage Z"
            HdrDat(3, Col + 5) = "Qt1y"
            HdrDat(3, Col + 6) = "Qt1z"
            HdrDat(2, Col + 7) = "Primary"
            HdrDat(3, Col + 7) = "beam (na)"
          End If

        End If

      End If

      SbmCt = PkNum + Npks * (ScanNum - 1)
      With RawDat(SpotCt)
        PkDat(ScanNum, TimeCol + 1) = .daPkCts(PkNum, ScanNum)
        PkDat(ScanNum, TimeCol + 2) = .daPkSigmaMean(PkNum, ScanNum)
        PkDat(ScanNum, TimeCol + 3) = .daSBMcts(PkNum, ScanNum)     ' total sbm cts
        SbmDat(SbmCt, 2) = .daSBMcts(PkNum, ScanNum) / IntT(PkNum)  ' sbm cps
        SbmDat(SbmCt, 1) = PkDat(ScanNum, TimeCol)
      End With
    Next PkNum

  Next ScanNum

  EndCol = picDatCol - 1 + 5 * Npks
  With frSr(SpotR + 1, picDatCol, SpotR + 3, EndCol + 4 - 6 * pbXMLfile)
    .Value = HdrDat: .Font.Bold = True
  End With
  r = SpotR + picDatRowOffs
  frSr(r, picDatCol, r + piNscans - 1, EndCol) = PkDat

  With RawDat(SpotCt)
    Cells(r, 1 + EndCol) = .dDeadTime * pdcBillion
    Cells(r, 2 + EndCol) = .lSBMzero

    If pbXMLfile Then
      Cells(r, 3 + EndCol) = .dStageX
      Cells(r, 4 + EndCol) = .dStageY
      Cells(r, 5 + EndCol) = .dStageZ
      Cells(r, 6 + EndCol) = .dQt1y
      Cells(r, 7 + EndCol) = .dQt1z
      Cells(r, 8 + EndCol) = Abs(.dPrimaryBeam)
    End If

  End With

'  NaRw = SpotR
'  plaSpotNameRowsCond(SpotCt) = NaRw

  If NameColWidth = 0 Then
    ColWidth picAuto, picNameDateCol
    NameColWidth = Columns(picNameDateCol).ColumnWidth
  End If

  If ScansPksColWidth = 0 Then
    ColWidth picAuto, picPksScansCol
    ScansPksColWidth = Columns(picPksScansCol).ColumnWidth
  End If

  s1 = psaSpotNames(SpotCt) & ", " & RawDat(SpotCt).sDate & "  " & RawDat(SpotCt).sTimeOfDay
  s2 = StR(Npks) & " peaks," & StR(piNscans) & " scans"
  Fonts SpotR, picNameDateCol, , , RGB(0, 96, 0), True, , 11, _
        , , s1, , , , True
  Fonts SpotR, picPksScansCol, , , 96, True, , 11, , , s2, , , , True

  For PkNum = 1 To Npks
    Col = 5 * (PkNum - 1) + 3
    frSr(SpotR, Col, SpotR + 3 + piNscans, Col + 4).Font.Color = _
         IIf(PkNum Mod 2, Clr1, Clr2)

    If RawDat(SpotCt).baCenteredPeak(PkNum) Then
      Fonts SpotR + picDatRowOffs - 1, Col + 1, , , , , , 9, , True
    End If

  Next PkNum

  If SpotCt < piNumAllSpots Then
    rw1 = LastHdrRw + piNscans + 1

    If pbShortCondensed Then
      rw2 = FirstDatRw + 2 * piaFileNscans(SpotCt) - 1
    Else
      rw2 = plaSpotNameRowsRaw(1 + SpotCt) - 2
    End If

    frSr(rw1, , rw2).Rows.Hidden = True
  End If

NextSpotCt: ActiveWindow.ScrollRow = LastHdrRw

Next SpotCt
StatBar "Formatting condensed sheet"

If pbXMLfile Then
  rw1 = 1: rw2 = flEndRow(picDatCol)
Else
  rw1 = plaSpotNameRowsCond(1)
  rw2 = flEndRow(picDatCol)
End If

HA xlCenter, rw1, picDatCol, rw2, picDatCol + 3 + Npks * 5 - 8 * pbXMLfile

If Not pbShortCondensed Then
  rw1 = 1 + flEndRow(picDatCol)
  rw2 = flEndRow(1)
  ColWidth picAuto, 1
  frSr(rw1, , rw2).Rows.Hidden = True
  Na1rw = plaSpotNameRowsCond(1)
  If Na1rw > 5 Then
    frSr(5, , Na1rw - 1).Hidden = True
  End If
End If

For Col = picDatCol To picDatCol - 1 + 5 * Npks
  k = 1 + (Col - picDatCol) Mod 5

  Select Case k
    Case 2: NumForm = "[>1000]0.0;[<100]0.0000;0.000"
    Case 3: NumForm = "[>100]0;[<10]0.00;0.0"
    Case 4: NumForm = "[>10000]0;[<1000]0.000;0.00"
    Case 5: NumForm = "[>100]0.000;[<10]0.00000;0.00000"
  End Select

  If k > 1 Then Columns(Col).NumberFormat = NumForm
  ColWidth picAuto, Col
Next Col

ColWidth picAuto, Col, Col + 3 - 6 * pbXMLfile
ColWidth picAuto, 1
ColWidth NameColWidth, picNameDateCol
ColWidth ScansPksColWidth, picPksScansCol

If ForScanGrafix Then
  Erase pdaPkCts:  Erase pdaPkT: Erase IntT
  Erase pdaSBMcps: Erase PkDat: Erase SbmDat
End If

tmp1 = IIf(pbPDfile, "PD  ", "XML ")
tmp2 = IIf(pbShortCondensed, "short", "long")
Fonts 1, picRawfiletypeCol, , , PeDarkRed, True, xlRight, , , , , , , , , tmp1 & " file,"
Fonts 1, picRawfileFirstcolCol, , , PeDarkRed, True, xlLeft, , , , , , , , , tmp2

For SpotCt = 1 To piNumAllSpots
  Cells(SpotCt + 1, 2) = plaSpotNameRowsCond(SpotCt)
Next SpotCt

Rows(plaSpotNameRowsCond(1)).RowHeight = 22

If pbPDfile Then
  Fonts 1, picDatCol, , , vbRed, True, xlLeft, , , , , , , , , FL(3)
  HA xlLeft, 2, picDatCol, 3
  Cells(2, picDatCol) = FL(1)
  Cells(3, picDatCol) = FL(2)
ElseIf pbXMLfile Then
  Fonts 1, picDatCol, , , vbBlue, False, xlLeft, 11, , , , , , , , psXmlFileType
  Fonts 2, picDatCol, , , RGB(0, 160, 0), False, xlLeft, 11, , , , , , , , psRawdatSoftwareVer
End If

Fonts 4, picDatCol, , , vbBlack, True, xlLeft, 11, , , fsS(piNumAllSpots) & " Spots"
HA xlLeft, 4, picDatCol
HA xlCenter, , 2
Fonts 1, peReadyCol, , , vbBlue, , xlLeft, , , , "squid ready", , , , True
ColWidth picAuto, 2
Columns("A:B").Hidden = True
ActiveSheet.Tab.Color = 128
ActiveWindow.ScrollRow = 1
NoUpdate False
StatBar "Done"
NoUpdate
End Sub

Sub CondenseSeveralFiles()

Dim Bad As Boolean, SaveEm As Boolean, Exists As Boolean
Dim XMLastext As Boolean, WasXML As Boolean
Dim s$, Na$, RawFileNa$, NaNoExt$, Ext$, Mtype$
Dim FileNameAndPath$, FileName$, FileNames$(), FileLines$()
Dim Npks%, i%, j%, p%, Nfiles%, LoopCt%, Le%, FileN%, FileCt%, Origin&
Dim PkMass#()
Dim FileNa As Variant, RawDat() As RawData

GetInfo
ChDirDrv foUser("sqPDfolder"), Bad

Mtype$ = "Raw-data files (*.pd; *.txt; *.xml),**.pd;.txt;.xml"
Alerts False
NoUpdate

FileNa = foAp.GetOpenFilename(FileFilter:=Mtype, MultiSelect:=True, _
          Title:="Select one or more PD files to condense/reformat")

On Error GoTo 1
Nfiles = UBound(FileNa)
If Nfiles = 0 Then Exit Sub

SaveEm = False
If Nfiles > 1 Then SaveEm = _
  (MsgBox("Save converted files (as *.XLS)?", vbYesNo, pscSq) = vbYes)
On Error GoTo 0

On Error GoTo 2

For i = 1 To Nfiles
  FileNameAndPath = FileNa(i)
  FileName = fsExtrRight(FileNameAndPath, "\", True)
  s = Right$(FileName, 4)
  pbXMLfile = (UCase(s) = ".XML")
  pbPDfile = Not pbXMLfile
  If Workbooks.Count > 0 Then
    With ActiveWorkbook
      If .Name <> ThisWorkbook.Name Then .Close
    End With
  End If
  'XMLconversion FileNameAndPath, Origin&, XMLastext, WasXML, Bad
  psRawFileName = FileNameAndPath
  Npks = 0

  CondenseRawData Npks, PkMass, True, RawDat, FileLines

  If SaveEm Then
    On Error GoTo 0
    Na = FileName 'ActiveWorkbook.Name
    LoopCt = 0

    Do
      LoopCt = 1 + LoopCt
      Le = Len(Na) - 3 + pbXMLfile
      NaNoExt = Left$(Na, Le)
      Ext = Mid$(Na, Le + 1)

      If fbFileNameExists(CurDir, NaNoExt & ".xls") Then
        Le = Len(NaNoExt)

        For j = Le To 1 Step -1
          p = Asc(Mid$(NaNoExt, j, 1))
          If p < 49 Or p > 57 Then Exit For
        Next j

        If Mid$(NaNoExt, j, 1) <> "_" Then
          NaNoExt = NaNoExt & "_"
          j = j + 1
        End If

        Na = Left$(NaNoExt, j) & fsS(LoopCt) & Ext
      Else
        Exit Do
      End If

    Loop

    Na = NaNoExt & ".XLS"
    On Error GoTo CouldntSave
    StatBar "Saving " & Na
    ActiveWorkbook.SaveAs Na, xlWorkbookNormal
    On Error GoTo 0
    ActiveWorkbook.Close
  End If

4: Next i
StatBar

1: Exit Sub
2: On Error GoTo 0
MsgBox "Error accessing file", , pscSq
3: On Error GoTo 0
MsgBox "Error condensing " & pscLF2 & FileNameAndPath, , pscSq
GoTo 4
CouldntSave: On Error GoTo 0
MsgBox "Unable to save " & Na, , pscSq
End Sub

Sub TooManyMassStations(FileType$, Npeaks%)
MsgBox "This PD file has" & Npeaks & " mass stations, which exceeds the maximum number " _
       & "of " & StR(peMaxNukes) & " allowed by SQUID-2.   Must abort.", , pscSq
CrashEnd
End
End Sub
