Attribute VB_Name = "FormSupport"
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
' 09/03/01 All lower array bounds explicit

Option Explicit
Option Base 1

Sub NameR(ByVal Npars%, FormControls As Controls, SortedControls() As Control, _
    ByVal Left1!, ByVal Top1!, ByVal Start1%, ByVal End1%, Optional NumCtrlsFound, _
    Optional ByVal Left2 = 0, Optional ByVal Top2 = 0, Optional ByVal Start2 = 0, _
    Optional ByVal End2 = 0, Optional ByVal DontRedim As Boolean = False, _
    Optional SortedControlNames)
' Takes up to 2 columns of like controls (say textboxes containing equations) all of
'  whose Top parfamters must be larger than Top1! or Top2!, and all with Left = Left1
'  or Left2.  The distribution into column1 or column2 is defined by the Start-End
'  index params.

Dim Names As Boolean
Dim tmpNa$()
Dim StartCtrl%, EndCtrl%, ToSortCt%, ict%, CtrlCt%, NumCtrlCt%, i%
Dim ToSortIndx&()
Dim Topp!, Leftt!
Dim ToSortCtrlTop() As Variant
Dim tmp() As Control, Ctrl As Control

If Not DontRedim Then
  ReDim SortedControls(1 To Npars)
End If
Names = fbNIM(SortedControlNames)
On Error GoTo 0
NumCtrlCt = 1 - (Start2 > 0)

For CtrlCt = 1 To NumCtrlCt
  ToSortCt = 0
  StartCtrl = Choose(CtrlCt, Start1, Start2)
  EndCtrl = Choose(CtrlCt, End1, End2)
  Leftt = Choose(CtrlCt, Left1, Left2)
  Topp = Choose(CtrlCt, Top1, Top2)
  ReDim tmp(StartCtrl To EndCtrl), ToSortCtrlTop(StartCtrl To EndCtrl), ToSortIndx(StartCtrl To EndCtrl)

  If Names Then
    If Not DontRedim Then ReDim SortedControlNames(StartCtrl To EndCtrl)
    ReDim tmpNa(StartCtrl To EndCtrl)
  End If

  For Each Ctrl In FormControls
    With Ctrl
      ' Do the control's Top and Left params meet the input specs?

      If fdPrnd(.Top - Topp, -2) >= 0 And Abs((.Left - Leftt)) < 0.001 Then
         ToSortCt = 1 + ToSortCt
         ict = StartCtrl + ToSortCt - 1
         Set SortedControls(ict) = Ctrl
         ToSortCtrlTop(ict) = .Top
         If Names Then SortedControlNames(ict) = Ctrl.Name
      End If

    End With
  Next Ctrl

  NumCtrlsFound = ToSortCt
  For i = 1 To ToSortCt: ToSortIndx(i) = i: Next
  QuickIndxSort ToSortCtrlTop, ToSortIndx

  For i = StartCtrl To EndCtrl
    If ToSortIndx(i) = 0 Then GoTo 1
    Set tmp(i) = SortedControls(ToSortIndx(i))
    If Names Then tmpNa(i) = SortedControlNames(ToSortIndx(i))
  Next i

  For i = StartCtrl To EndCtrl
    Set SortedControls(i) = tmp(i)
    If Names Then SortedControlNames(i) = tmpNa(i)
  Next i
Next CtrlCt

1:
End Sub

Sub CheckRatLetsEqNums(ByVal EqNum%, Eqn As Control, Warned$, OK As Boolean, RatCtrl, _
  EqNames, Optional SCSwitch, Optional FOswitch, Optional FutureEqnRefsOK As Boolean = False)
' Check the validity of [bracketed] and ["bracketed"] references to range names, column headedrs,
'  Task equations, Task isotope ratios, and Task constants in a Task equation.
Dim BadRat As Boolean, RangeLegal As Boolean, BadEq As Boolean, BadWbk As Boolean, BadWksht As Boolean
Dim HasBracketedRefs As Boolean, Lett(1 To 2) As Boolean, Numb(1 To 2) As Boolean
Dim Eq$, TrmRat$, EqRef$, RatRef$, IndxRef$, tmp$, Legal$, EqRefStripped$
Dim LcLegalNa$, Msg$, LegalEqRef$, BrRpos_1$, LastCheck$, LoopLock$
Dim WbkName$, WkShtName$, LegalRangenameOrColHdr$, ShtRef$
Dim EqIn$, UserEqRef$, TrueEqRef$, UcEq$, LcEq$, EqNa$(), WbNa$(), ShNa$()
Dim NumRats%, NumEqns%, CharPos%, ExclPos%, BrqLpos%, BrLpos%, BrqRpos%, BrRpos%
Dim RatIndx%, NameIndx%, EqIndx%, HdrIndx%, TotCtsIndx%, Nrefs%
Dim EqAsc%, Which%, NumR%, Nume%, LenEqRef%, NcolHdrs%, LenIndxRef%, b%, LoopCt%
Dim RatNum%(1 To 2), EqRefNum%(1 To 2), CharAsc%(1 To 2)
Dim ColumnHeaders As Range

' Check validity of delimiter-b racketed letters or numbers as indicating defined ratios or eqns.
For RatIndx = 1 To UBound(RatCtrl)
  If RatCtrl(RatIndx) = "" Then Exit For
Next
NumRats = RatIndx - 1

For NameIndx = 1 To UBound(EqNames)
  If EqNames(NameIndx) = "" Then Exit For
Next

NumEqns = NameIndx - 1
Set ColumnHeaders = ThisWorkbook.Sheets("ColHdrs").[ColumnHeaders]
NcolHdrs = ColumnHeaders.Count
EqIn = Eqn
LoopLock = "Loop-locked in sub CheckRatLetsEqNums, with Eqn =  " _
          & EqIn & " ." & pscLF2 & "Please notify Ken Ludwig."
UcEq = Trim(Eqn)
LcEq = LCase(UcEq)

1:
LastCheck = ""
OK = False

If fbIM(RatCtrl) Then
  Which = 2
ElseIf fbIM(EqNames) Then
  Which = 1
Else
  Which = 3
End If

HasBracketedRefs = (InStr(LcEq, psBrQL) > 0 And InStr(LcEq, psBrQR) > 0)

If HasBracketedRefs And fbNIM(FOswitch) Then
  FOswitch = True
End If

Subst LcEq, "[" & pscPm, "["
Subst LcEq, "[+", "["

LoopCt = 0 ' check for literal equation & ratio refs

Do
  LoopCt = 1 + LoopCt
  If LoopCt > 99 Then MsgBox LoopLock, , pscSq: CrashEnd
  ' Find wbk/Wksht refs & strip
  FindWbkShtRefs LcEq, Nrefs, WbNa, ShNa, BadWbk, BadWksht ', LcEq

  If BadWbk Or BadWksht Then '(BrqLpos > 0 And BrqRpos = 0) Or (BrqLpos = 0 And BrqRpos > 0) Then
    Msg = "The equation  " & Eqn & "  contains an improperly configured " _
        & "reference to a workbook or worksheet." 'expression within square brackets."
    MsgBox Msg, , pscSq
    GoTo Bad
  End If

  LcEq = LCase(Eq)
  ' Remove all literals (quote-bracketed) from LcEq
  BrqLpos = InStr(LcEq, psBrQL): BrqRpos = InStr(LcEq, psBrQR)

  If BrqLpos > 0 And BrqRpos > BrqLpos Then
    EqRef = Trim(Mid$(LcEq, BrqLpos + 2, BrqRpos - BrqLpos - 2))
    LastCheck = EqRef: RatIndx = 0
    EqRefStripped = EqRef
    Subst EqRefStripped, "/"
    Subst EqRefStripped, "."

    ' Is Eqref a literal, defined isotope-ratio?
    If fbIsAllNumChars(EqRefStripped, True, True) Then

      For RatIndx = 1 To NumRats ' a defined-ratio literal?
        TrmRat = Trim(RatCtrl(RatIndx))

        If TrmRat = "" Then
          NumRats = RatIndx - 1
          Exit For
        ElseIf TrmRat = EqRef Then
          LcEq = Left$(LcEq, BrqLpos - 1) & Mid$(LcEq, BrqRpos + 2)
          Exit For
        End If

      Next RatIndx

    End If

    LegalEqRef = LCase(fsLegalName(EqRef))
    LegalRangenameOrColHdr = LCase(fsLegalName(EqRef, True, , False))

    If (RatIndx = 0 Or RatIndx > NumRats) And Not pbDefiningNew Then

      For EqIndx = 1 To NumEqns ' Is LegalEqRef a literal Equation name?
        Legal = IIf(RangeLegal, LegalRangenameOrColHdr, LegalEqRef)
        LcLegalNa = LCase(fsLegalName(EqNames(EqIndx), True)) 'RangeLegal))

        If LcLegalNa = Legal Or LegalEqRef = Legal Then
          LcEq = Left$(LcEq, BrqLpos - 1) & Mid$(LcEq, BrqRpos + 2)
          Exit For
        End If

      Next EqIndx

      If EqIndx > NumEqns Then

        ' Is LegalRangenameOrColHdr a common U-Pb column header?
        If ((EqIndx = 0 Or EqIndx > NumEqns) Or HasBracketedRefs) Then

          For HdrIndx = 1 To NcolHdrs
            LcLegalNa = LCase(fsLegalName(ColumnHeaders(HdrIndx)))

            If LcLegalNa = LegalEqRef Then
              LcEq = Left$(LcEq, BrqLpos - 1) & Mid$(LcEq, BrqRpos + 2)
              Exit For
            End If

            If Not pbDefiningNew Then

              For b = 1 To NumEqns

                If LcLegalNa = EqNames(b) Then
                  LcEq = Left$(LcEq, BrqLpos - 1) & Mid$(LcEq, BrqRpos + 2)
                  Exit For
                End If

              Next b

            End If
          Next HdrIndx

          If HdrIndx > NcolHdrs Then

            For TotCtsIndx = 1 To piNtotctsHdrs
              LcLegalNa = LCase(fsLegalName(psaTotCtsHdrs(TotCtsIndx)))

              If LcLegalNa = LegalEqRef Then
                LcEq = Left$(LcEq, BrqLpos - 1) & Mid$(LcEq, BrqRpos + 2)
                Exit For
              End If

            Next TotCtsIndx

            If TotCtsIndx > piNtotctsHdrs Then
              Msg = psBrQL & EqRef & psBrQR & "  is not a valid equation " & _
                    "or isotope-ratio reference, and is probably(?) not a " & _
                    "column-header reference." & pscLF2 & "Proceed anyway?"

              If MsgBox(Msg, vbYesNo, pscSq) = vbYes Then
                LcEq = Left$(LcEq, BrqLpos - 1) & Mid$(LcEq, BrqRpos + 2)
                BrqLpos = 0
              Else
                GoTo Bad
              End If

            End If

          End If

        End If

      End If

    End If

  End If

  If BrqLpos > 0 Then

    If Mid$(LcEq, BrqLpos, 1) = "[" And pbUPb Then
      MsgBox "Can't identify the reference " & fsInQ(EqRef) & _
              " in Equation" & StR(EqNum), , pscSq
      BadEq = True: GoTo Bad
    ElseIf Not pbUPb Then
      'BrqLpos = 0
    End If

  End If

Loop Until BrqLpos = 0 Or BrqRpos = 0

LoopCt = 0 ' Now check for equation or ratio index-refs

Do
  LoopCt = 1 + LoopCt
  If LoopCt > 99 Then MsgBox LoopLock, , pscSq: CrashEnd
  ' All [" --- "] (quote-bracketed) strings removed:
  BrLpos = InStr(LcEq, "[")
  BrRpos = InStr(LcEq, "]")  ' Now sequentially identify & remove
  ExclPos = InStr(LcEq, "!") '  simple-bracketed strings
  BrRpos_1 = Mid$(LcEq, BrRpos + 1, 1)

  If BrLpos > 0 And BrRpos > BrLpos And (BrRpos - BrLpos) < 4 Then
    ' Is the bracketed string an eqn, e.g. [11] or ratio, e.g. [b], ref?
    IndxRef = Trim(fsSubStr(LcEq, BrLpos + 1, BrRpos - 1))
    LenIndxRef = Len(IndxRef)

    For CharPos = 1 To LenIndxRef

      CharAsc(CharPos) = Asc(fsSubStr(IndxRef, CharPos, CharPos))
      If Which <> 1 Then Numb(CharPos) = _
        (CharAsc(CharPos) > 47 And CharAsc(CharPos) < 58)
      If Which <> 2 Then Lett(CharPos) = _
        (CharAsc(CharPos) > 96 And CharAsc(CharPos) < 123)

      If CharPos = 2 Then
        tmp = "[" & IndxRef & "]"

        If (Which <> 1 And Numb(1) Xor Numb(2)) Or _
           (Which <> 2 And Lett(1) Xor Lett(2)) Then
          MsgBox tmp & " is not a valid isotope ratio- or equation-index.", , pscSq
          GoTo Bad
        End If

      End If

      If Lett(CharPos) Then                    ' An isotope-ratio index, e.g. [ab]?
        RatNum(CharPos) = CharAsc(CharPos) - 96

        If CharPos = LenIndxRef Then
          NumR = RatNum(CharPos)
          If LenIndxRef = 2 Then NumR = 26 + RatNum(CharPos)  ' Valid [ab]?
          If NumR > UBound(RatCtrl) Then BadRat = True: GoTo Bad

          If RatCtrl(NumR).Caption = "" Then
            MsgBox "[" & IndxRef & "]  is not a valid isotope-ratio index" _
              & " in Equation " & fsS(EqNum), , pscSq
            BadRat = True: GoTo Bad
          End If

        End If

      ElseIf Numb(CharPos) Then                ' An equation index, e.g. [12]?
        EqRefNum(CharPos) = CharAsc(CharPos) - 48

        If CharPos = LenIndxRef Then
          Nume = EqRefNum(CharPos)
          If LenIndxRef = 2 Then Nume = 10 + EqRefNum(CharPos)

          If Nume > UBound(EqNames) Then
            BadEq = True: GoTo Bad
          ElseIf EqNames(Nume) = "" Then
            BadEq = True: GoTo Bad
          End If

        End If

      Else
        GoTo Bad
      End If

    Next CharPos

    IndxRef = Mid$(LcEq, BrRpos + 1)
    LcEq = IndxRef

  ElseIf BrRpos > 0 And BrLpos < (BrRpos - 2) Then
    If BrRpos_1 = "'" And ExclPos > BrRpos Then

      If LCase(Mid$(LcEq, BrRpos - 4, 4)) <> ".xls" Then
        MsgBox "Workbook references must have the .XLS extension", , pscSq
        BadEq = True: GoTo Bad
      Else
        LcEq = Left$(LcEq, BrLpos - 1) & Mid$(LcEq, ExclPos + 1)
      End If

    Else

      UserEqRef = Mid$(LcEq, BrLpos, BrRpos - BrLpos + 1)
      TrueEqRef = psBrQL & Mid$(LcEq, BrLpos + 1, BrRpos - BrLpos - 1) & psBrQR
      Msg = "The reference in Equation" & StR(EqNum) & " to  " & _
        UserEqRef & "  is not legal." & pscLF2 & _
        "Did you intend to reference to be  " & TrueEqRef & "  ?"

      If MsgBox(Msg, vbYesNo, pscSq) = vbYes Then
        UcEq = LCase(UcEq)
        Subst UcEq, UserEqRef, TrueEqRef
        Eqn = UcEq
        LcEq = UcEq
        GoTo 1
      Else
        GoTo Bad
      End If

    End If

  End If

Loop Until BrLpos = 0 Or BrRpos = 0 Or (BrRpos - BrLpos) < 2

OK = True
Exit Sub
Bad: OK = False
On Error Resume Next
Eqn.SetFocus
End Sub

Sub GetRats(Ctrls As Controls, RatCtrl() As Control, Optional NumRats% = 0)
' Put isotope ratios strings into isotope ratio labels
Dim RatIndx%, RatStr$
GetTaskRatios
If fbNIM(NumRats) Then NumRats = puTask.iNrats

For RatIndx = 1 To peMaxRats
  With RatCtrl(RatIndx)

    If RatIndx > puTask.iNrats Then
      RatStr = ""
      .Enabled = False: .BackColor = &H8000000F
      .BorderColor = &H80000011
      .ForeColor = peGray
    Else
      RatStr = puTask.saIsoRats(RatIndx)
      .Enabled = True: .ForeColor = 0
      .ControlTipText = "Click on a ratio to insert it into the active Equation box"
    End If

    .Caption = RatStr
  End With
Next RatIndx

End Sub

Function fhAtWts() As Worksheet
Set fhAtWts = ThisWorkbook.Sheets("atomicweights")
End Function

Function frActiveUPbTaskName() As Range
Set frActiveUPbTaskName = foUser.[ActiveUPbTaskName]
End Function

Function frActiveUPbTaskNum() As Range
Set frActiveUPbTaskNum = foUser.[ActiveUPbTaskNum]
End Function

Function frActiveGeneralTaskNum() As Range
Set frActiveGeneralTaskNum = foUser.[ActiveGeneralTaskNum]
End Function

Function frActiveTaskNum() As Range
Set frActiveTaskNum = foUser.[ActiveTaskNum]
End Function

Function frActiveGeneralTaskName() As Range
Set frActiveGeneralTaskName = foUser.[ActiveGeneralTaskName]
End Function

Sub AssignEqnSwitchVals(Eqns$(), Neqns%, BadEqn%, ByVal UPbGeochron As Boolean, _
     STswitch, SAswitch, SCSwitch, LAswitch, FOswitch, NUswitch, HIswitch, ARswitch, ARrc)
' Copy Task Equation Switch settings from the Equations panel of the Task Editor into
'  the puTask variable
Dim rc1$, rc2$, Eqn$, EqNum%, SpacePos%

If Neqns = 0 Then Exit Sub
With puTask
  ReDim Preserve .uaSwitches(1 To Neqns)
  piLastN = 0

  For EqNum = 1 To Neqns
    Eqn = Eqns(EqNum)
    With .uaSwitches(EqNum)

      If Eqn <> "" Then
        If UPbGeochron Then
          .ST = (STswitch(EqNum) <> "")
          .SA = (SAswitch(EqNum) <> "")
          .LA = (LAswitch(EqNum) <> "")
          piLastN = piLastN - .LA
        Else
          .ST = False: .SA = False: .LA = False
        End If

        .SC = (SCSwitch(EqNum) <> ""): .FO = (FOswitch(EqNum) <> "")
        .Ar = (ARswitch(EqNum) <> ""): .Nu = (NUswitch(EqNum) <> "")
        .HI = (HIswitch(EqNum) <> "")
        If .Ar Then
          SpacePos = InStr(ARrc(EqNum), " ")

          If SpacePos > 0 Then
            rc1 = Left$(ARrc(EqNum), SpacePos - 1)
            rc2 = Mid$(ARrc(EqNum), SpacePos + 1)
          Else
            rc1 = Left$(ARrc(EqNum), 1)
            rc2 = Mid$(ARrc(EqNum), 2)
          End If

          .ArrNrows = fvMax(1, Val(rc1))
          .ArrNcols = fvMax(1, Val(rc2))

          If .ArrNrows = 0 Or .ArrNcols = 0 Then
            MsgBox "The AR switch in Equation" & StR(EqNum) & " requires the #rows/#cols " _
              & " of the output in the appropriate box.", , pscSq
            BadEqn = True: Exit Sub
          ElseIf False Then
            MsgBox "The SC switch for Equation" & StR(EqNum) & _
                   " is forbidden for multi-cell Array equations.", , pscSq
            .SC = False: BadEqn = EqNum: Exit Sub
          End If

        End If

      End If

      If Not (.FO Or .Nu) Then  ' and UPbGeochron
        If .SC And .LA Then

          .FO = True: .Nu = False
        ElseIf Not (.SC Or .LA) Then
          .Nu = True: .FO = False
        End If

        If .Nu And .FO Then
          MsgBox "The NU and FO switches are mutually exclusive.  Pick one or none.", , pscSq
          BadEqn = EqNum: Exit Sub
        End If

        If False And .Ar And Not .SC And .ArrNrows > 1 Then
          MsgBox "The number of output rows for array functions that " & _
          "are not specified as Single-Cell (SC) output must be 1." & vbLf _
          & vbLf & "Please correct output #rows/#cols for equation" & StR(EqNum) & ".", , pscSq
          BadEqn = EqNum: Exit Sub
        End If

      End If

    End With

  Next EqNum

End With
End Sub

Sub BiwtConvert(EqnString)
' Change all forms Task eqn references to the Biweight function to "sqBiweight"
'  to ensure that the SQUID-2 form of the function will be employed.  Do the
'  same for refs to the MAF function, changing them to "sqMAD".
Dim LcEqn$, Eqn$, p%, q%, MADpos%, CharCode%

If IsObject(EqnString) Then Eqn = EqnString.Text Else Eqn = EqnString
LcEqn = LCase(Eqn)
RemoveDblChars LcEqn, "sq"
p = InStr(LcEqn, "biweight(")
q = InStr(Eqn, "sqbiweight(")

If p > 0 And q = 0 Then
  Eqn = Left$(Eqn, p - 1) & "sqBiweight(" & Mid$(Eqn, p + 9)
  Subst Eqn, "biweight(", "sqBiweight("
End If

Do
  MADpos = InStr(LcEqn, "mad(")

  If MADpos > 0 Then

    If MADpos > 1 Then
      CharCode = Asc(Mid$(Eqn, MADpos - 1, 1))
    Else
      CharCode = 0
    End If

    If MADpos = 1 Or (MADpos > 1 And (CharCode < 97 Or CharCode > 122)) Then
      Eqn = Left$(Eqn, MADpos - 1) & "$$$" & Mid$(Eqn, MADpos + 3)
      LcEqn = LCase(Eqn)
    Else
      Exit Do
    End If

  End If

Loop Until MADpos = 0

Subst Eqn, "$$$", "sqMAD"
If IsObject(EqnString) Then EqnString.Text = Eqn Else EqnString = Eqn
End Sub

Sub CheckEqnSyntax(Eqn() As Control, Isorats$(), Optional EqNames, Optional Bad As Boolean = False)
' Determine if a Task eqn's references to likely output-sheet range-names and column-headers,
'  as well as Task constants & equations are correct & Excel-parseable.

Dim NoSheet As Boolean, BadSwap As Boolean
Dim Eq$, tmp$, Cna$, s1$, s2$, TrEqNa$, LegalEqNa$, Eq0$, Msg$
Dim i%, p%, EqIndx%, IndxType%, EqIndxStr$
Dim ConstIndx%, ExtrIndx%, SheetCt%, InstanceCt%, EqNum%, ColFuncIndx%, MultEqIndx%
Dim ColEq As Range, MultiEq As Range, Rarr As Variant

If puTask.iNeqns = 0 Then Exit Sub

ReDim Clen%(1 To peMaxConsts), ClenIndx%(1 To peMaxConsts)
With ThisWorkbook.Sheets("ColHdrs")
  Set MultiEq = .[MultiInputFunctions]
  Set ColEq = .[ColumnInputFunctions]
End With

If pbUPb Then
    Rarr = Array("wtdmeana1", "wtdmeana2", "wtdmeanaperr1", "wtdmeanaperr2", "scomm1_64", "scomm1_74", _
      "scomm1_84", "scomm1_76", "scomm1_86", "ConcStdPpm", "ConcStdConst", "AgeStdAge", "stdupbratio", _
      "Std_76", "StdAgeThPb", "stdrad86fact", "extperra1", "extperra2", "PbArr_2", "Aer_2", "Aadat_2", _
      "Aaerdat1_2", "Aaerdat2_2", "PbArr_2", "Aer_2", "Aadat_2", "Aaerdat1_2", "Aaerdat2_2", "%com206", _
      "ppmU", "232Th/238U", "204overcts/sec.(fr.207)", "204overcts/sec.(fr.208)", "208Pb*/206Pb*", _
      "204Pb/206Pb(fr.207)", "204Pb/206Pb(fr.208)", "7-corr206Pb/238Uconst.delta%", _
      "8-corr206Pb/238Uconst.delta%", "UncorrPb/Uconst", "4-corr207Pb/206Pbage", "207Pb*/206Pb*")
End If

Bad = False
For EqNum = 1 To puTask.iNeqns

  If InStr(LCase(puTask.saEqnNames(EqNum)), "<<solve>>") = 0 Then

    Msg = ""
    Eq0 = Trim(Eqn(EqNum))
    Eq = LCase(Eq0)

    With puTask.uaSwitches(EqNum)
      ' Check for: missing names, header-rangenames not starting with a #,

      If fbNIM(EqNames) Then
        TrEqNa = Trim(EqNames(EqNum))

        If TrEqNa = "" Then
          Msg = "All defined User Equations must be assigned a name."
          GoTo BadExit
        End If

        If .SC Then
          CurlyExtract TrEqNa, "", , , , , True
          LegalEqNa = fsLegalName(TrEqNa, True)

          If LegalEqNa = "" Then
            Msg = fsInQ("{" & TrEqNa & "}") & " isn't a legal spot-name reference."
            GoTo BadExit
          End If

          Alerts False
          NoSheet = False
          On Error GoTo 2
          SheetCt = ActiveWorkbook.Worksheets.Count
          GoTo 3
2:        NoSheet = True
          On Error GoTo 0
          CreateNewWorkbook
          Msg = fsInQ(LegalEqNa) & " is not a legal Excel range-name"
          If fbIsNumChar(Left$(LegalEqNa, 1)) Then Msg = Msg & " (can't start with a number)"
          Msg = Msg & " --" & pscLF2 & "Please assign another name to equation" & StR(EqNum) & "."
3:        On Error GoTo BadExit
          s1 = LegalEqNa
          Msg = s1 & "  is not a legal Equation name." ' 09/06/10 -- added
          Cells(1, 256).Name = s1
          If NoSheet Then ActiveWorkbook.Close
          On Error GoTo 0
          Msg = ""
          Alerts True
        End If   ' SC

        If .Nu Then ' does the eqn require any not-yet-calculated eqns?

          InstanceCt = 0
          Do
            InstanceCt = 1 + InstanceCt
            ExtractEqnRef Eq, "", EqIndx, IndxType, , EqNames, , InstanceCt

            If IndxType = peEquation Then

              If Not .LA Or puTask.uaSwitches(EqIndx).LA Then
                s1 = "Equation" & StR(EqNum): s2 = "Equation" & StR(EqIndx)
                Msg = s1 & " requires data from " & s2 & " at a point when the result(s) of " _
                 & s2 & " will not yet exist." & pscLF2 & "Please do one of the following:" _
                 & pscLF2 & "    1) Place " & s1 & " after " & s2 & "," & pscLF2 & _
                 "    2) Set the LA switch for " & s1 & ", or" & pscLF2 & "    3) Set " & _
                 "the FO switch for " & s1 & "."
                GoTo BadExit
              End If

            End If

          Loop Until EqIndx = 0

        End If     ' NU switched
      End If       ' Names passed? (fbNIM(EqNames))

      Subst Eq, "[" & pscPm, "["
      Msg = ""

      With prConstsRange
        '  Look from longest names to shortest to avoid
        '  "_Lambda238ka" being subst. with "_Lambda238"

        For ConstIndx = 1 To peMaxConsts
          Clen(ConstIndx) = Len(.Item(ConstIndx, 1))
        Next ConstIndx

        BubbleSort Clen, ClenIndx, True
      End With

      For ConstIndx = 1 To peMaxConsts  ' Underscored const names?
        Cna = fs_(prConstNames(ClenIndx(ConstIndx), True))
        If Cna = "" Then Exit For
        p = InStr(Eq, Cna)

        If p > 0 Then
         s1 = prConstValues(ClenIndx(ConstIndx))
          Subst Eq, Cna, Val(s1)
        End If

      Next ConstIndx

      If pbUPb Then

        For i = 1 To UBound(Rarr)
          Subst Eq, (Rarr(i)), "1"
        Next i

      End If

      If Not fbLegalEq(Eq, False, Isorats, BadSwap) Then

        If BadSwap Then
          GoTo BadExit
        Else
          tmp = fsQq("The equation  $" & Eq0 & "$  may refer to a nonexistent " _
          & "equation, or may not be in legal Excel format." & pscLF2 & _
          "Proceed anyway?")
          i = MsgBox(tmp, vbYesNo, pscSq)
          If i = vbNo Then GoTo BadExit
        End If

      End If

      If Not .SC And Not .Ar Then ' Insist on SC switch for column-input equations

        For ColFuncIndx = 1 To ColEq.Count
          EqIndxStr = LCase(ColEq(ColFuncIndx))

          If Left$(Eq, Len(EqIndxStr)) = EqIndxStr Then
            MsgBox "Unless used as (and useable as) an Array equation, the  " & _
              fsInQ(UCase(Left$(EqIndxStr, 1)) & Mid$(EqIndxStr, 2)) & _
              "  function requires the .SC switch.", , pscSq
            GoTo BadExit
          End If

        Next ColFuncIndx

      Else

        Dim FirstComma%, FirstLsq%, FirstRsq%, SecondRsq%, SecondLsq%, FirstLparen%, EqFr$

        For MultEqIndx = 1 To MultiEq.Count
          EqIndxStr = LCase(MultiEq(MultEqIndx))
          ' worry about having both "th230age" and "th230ageandinitial"
          p = InStr(Eq, EqIndxStr)

          If p > 0 Then
            EqFr = Eq                       ' an equation requiring at least 2 arguments
            Subst EqFr, EqIndxStr
            FirstLparen = InStr(EqFr, "(")  ' th230ageandinitial([1],[1])
            FirstLsq = InStr(EqFr, "[")
            FirstRsq = InStr(EqFr, "]")
            FirstComma% = InStr(EqFr, ",")
            SecondLsq = fiInstanceLoc(EqFr, 2, "[")
            SecondRsq = fiInstanceLoc(EqFr, 2, "]")

            If FirstComma > 0 Then
              If FirstLparen = 1 Then
                If FirstRsq < FirstLsq Or _
                   FirstComma < FirstRsq Or SecondLsq < FirstComma Or _
                   SecondRsq < SecondLsq Then
                  MsgBox "Please ensure that the  " & fsInQ(UCase(Left$(EqIndxStr, 1)) & Mid$(EqIndxStr, 2)) & _
                       "  function has at least 2 arguments.", , pscSq
                End If

              End If

            End If

          End If

         Next MultEqIndx

        End If

    End With 'puTask.uaSwitches(EqNum)

  End If ' not solver

Next EqNum
Exit Sub

BadExit: On Error Resume Next
Eqn(EqNum).SetFocus
On Error GoTo 0
If Msg <> "" Then MsgBox Msg, , pscSq
Bad = True
Alerts False
If NoSheet Then ActiveWorkbook.Close
Alerts True
End Sub

Sub RatClick(Ctrls As Controls, ByVal RatLett$, ByVal EqNum%, _
  EqCtrls() As Control, ByVal IndexClickResult%)
' Respond to a click on an Isotope Ratio box from the Equations panel
'  by appending the appropriated-formatted reference to the Equation.
Dim s$, bx As Control

If EqNum = 0 Then Exit Sub
Set bx = Ctrls("Lb" & RatLett)
s = bx.Caption

If s = "" Then
  Exit Sub
ElseIf IndexClickResult = 1 Then
  s = LCase(RatLett)
Else
  s = fsInQ(s)
End If

With EqCtrls(EqNum)
  .Value = .Value & "[" & s & "]"
End With
End Sub

Sub CalcNominalMassVal(Optional FromSheet As Boolean = False)
' Determine the "pdaFileNominal" masses for a Run Table, defined as being
'  the mass-value with the fewest after-decimal figures to distinguish
'  each mass from another.

' 09/11/12 -- modifed to correctly deal with duplicate masses and
'             masses crossing a power-of-ten boundary'

Dim Fline$(), i%, j%, UndupedN%
Dim Delt#, MinDelt#, mi#, mj#, tmp1#, tmp2#, tz%, tz0%
Dim RawDat() As RawData, UnDupedFileMass#(), SortedFileMass#()
Dim UndupedZ%(), FilemassZ%()

If piNumAllSpots = 0 And Not FromSheet Then
  InhaleRawdata pbPDfile, Fline, False
End If

ReDim pdaFileNominal(1 To piFileNpks), UnDupedFileMass(piFileNpks)
ReDim SortedFileMass(piFileNpks), UndupedZ(piFileNpks), FilemassZ(piFileNpks)

UndupedN = 0 ' Number of distinct masses in the run table

For i = 1 To piFileNpks
  SortedFileMass(i) = pdaFileMass(i)
Next i

QuickSort SortedFileMass

For i = 1 To piFileNpks - 1  ' Create a list of unduplicated masses
  tmp1 = SortedFileMass(i)

  If tmp1 <> SortedFileMass(1 + i) Then
    UndupedN = 1 + UndupedN
    UnDupedFileMass(UndupedN) = tmp1
  End If

Next i

UndupedN = 1 + UndupedN
UnDupedFileMass(UndupedN) = SortedFileMass(piFileNpks)

ReDim Preserve UnDupedFileMass(UndupedN)

For i = 1 To UndupedN - 1 ' Determine the rounding power-of-ten required to distinguish
                          '   each unduplicated mass from all others.
  MinDelt = 999

  mi = UnDupedFileMass(i)
  mj = UnDupedFileMass(i + 1)
  Delt = mj - mi
  If Delt < MinDelt Then MinDelt = Delt
  UndupedZ(i) = fvMin(0, Int(fdLog10(MinDelt)))

  If fdPrnd(mi, UndupedZ(i)) = fdPrnd(mj, UndupedZ(i)) Then
    UndupedZ(i) = fvMin(0, UndupedZ(i) - 1)
  End If

Next i

UndupedZ(UndupedN) = UndupedZ(UndupedN - 1)

For i = 1 To piFileNpks
  mi = pdaFileMass(i)

  For j = 1 To UndupedN
    If mi = UnDupedFileMass(j) Then FilemassZ(i) = UndupedZ(j)
  Next j

  If False And i > 1 Then
    tz = FilemassZ(i - 1)
    If FilemassZ(i) > tz Then FilemassZ(i) = tz
  End If

  pdaFileNominal(i) = fdPrnd(pdaFileMass(i), FilemassZ(i))
Next i

For i = 1 To piFileNpks     '  If mass i is 203.9324 and mass j is  203.9774, instead of rounding
                            '    the (nominal) masses to 203.93 and 204, round to 203.93 and 203.98
  For j = 1 To piFileNpks   '  That is, if any 2 pks have the same value when rounded to the less-
    If i <> j Then          '  precise mass, round them both to the same precision as the more.
      mj = fdPrnd(pdaFileMass(j), FilemassZ(i))
      If mj = pdaFileNominal(i) Then
        If FilemassZ(i) > FilemassZ(j) Then
          pdaFileNominal(i) = fdPrnd(pdaFileMass(i), FilemassZ(j))
        End If
      End If
    End If
  Next j
Next i

End Sub

Public Function fdNukeMass#(ByVal NuclideIn$, Optional ByVal Sigfigs, _
                            Optional ByVal CorrectCase = True)
' Input is the nuclide formula with elements separated by a space in the format
'  "90Zr2 16O"  or "142Ce2 31P 16O4 ++"
' Returns the atomic weight of the nuclide divided by its charge.

Dim Nuke$, q$, Charge$, Nuclide$, m As String * 1, Nu$()
Dim i%, p%, NukeIndx%, MassIndx%, NumNukes%, nq%, AmuNum%, Ln%, CharPos%
Dim w#, v#
Dim Amu As Range, tnu As Range

Nuke = "": NumNukes = 0
With fhAtWts
  Set Amu = .[Amu]: Set tnu = .[Nuclide] '.[Amu]
End With
Nuclide$ = Trim(NuclideIn) & " "
AmuNum = Amu.Count

Do
  NumNukes = 1 + NumNukes
  p = InStr(Nuclide, " ")
  ReDim Preserve Nu(1 To NumNukes)
  Nu(NumNukes) = Left$(Nuclide$, p - 1)
  Nuclide = LTrim(Mid$(Nuclide, p + 1))
Loop Until Not fbIsNum(Left$(Nuclide, 1))

For NukeIndx = 1 To NumNukes
  Nu(NukeIndx) = Trim(LCase(Nu(NukeIndx)))
  m = Right$(Nu(NukeIndx), 1)
  nq = IIf(m < "1" Or m > "9", 1, Val(m))
  Ln = Len(Nu(NukeIndx))
  q = Left$(Nu(NukeIndx), Ln + (nq > 1))
  Subst q, "+", , "-", , "ref", , "bkrd"

  For MassIndx = 1 To AmuNum
    If LCase(tnu(MassIndx)) = q Then
      p = InStr(Nuke, q)
      Ln = Len(tnu(MassIndx))
      Nu$(NukeIndx) = tnu(MassIndx) & Mid$(Nu(NukeIndx), 1 + Ln)
      Nuke = Nuke & " " & Nu(NukeIndx)
      Exit For
    End If
  Next MassIndx

  If MassIndx > AmuNum Then
    fdNukeMass = 0: Exit Function
  End If

  w = w + nq * Amu(MassIndx)
Next NukeIndx

i = 0: Charge = ""
Nuclide = Trim(NuclideIn)

For CharPos = Len(Nuclide$) To 1 Step -1
  m = Right$(Nuclide$, 1)
If m <> "+" And m <> "-" Then Exit For
  Charge = Charge & m
  Nuclide$ = Left$(Nuclide$, Len(Nuclide$) - 1)
  i = i + 1
Next CharPos

If CorrectCase Then Nuclide = Nuke & " " & Charge
v = w / IIf(i = 0, 1, i)
If fbNIM(Sigfigs) Then v = Drnd(v, Sigfigs)
fdNukeMass = v
End Function

Sub CreateColHdrs(tmpColHdrs$(), Nrats%, Rats$(), _
  Neqns%, Eqns$(), EqnNames$())
' Create output-worksheet column-headers

Dim tB As Boolean, s$, EqIndx%, RatIndx%, N%

ReDim tmpColHdrs(1 To 100)
N = 1
tmpColHdrs(1) = "Hours"
With puTask
  GetUPbPkOrders .iNpeaks, .dBkrdMass, .dRefTrimMass, .daNominal
End With

If piBkrdPkOrder > 0 Then _
  N = 1 + N: tmpColHdrs(N) = "Bkrd cts/sec"
If pbUPb And pi204PkOrder > 0 Then _
  N = 1 + N: tmpColHdrs(N) = "total 204 cts/sec"
If pbUPb And pi204PkOrder > 0 And pi206PkOrder > 0 Then _
  N = 1 + N: tmpColHdrs(N) = "total 206 cts/sec"

For RatIndx = 1 To Nrats
  tmpColHdrs(N + RatIndx) = Rats(RatIndx)
Next RatIndx

N = N + Nrats

For EqIndx = 1 + 5 * pbUPb To Neqns
  tB = False
  If EqIndx <> 0 Then
    s = EqnNames(EqIndx)

    If EqIndx > 0 Then
      With puTask.uaSwitches(EqIndx)
        tB = (Not .SC And Not .SA)
      End With
    End If

    If s <> "" Then

      If EqIndx > 0 And tB Then
        N = N + 1
        tmpColHdrs(N) = EqnNames(EqIndx)
      End If

    End If

  End If

Next EqIndx

If True Then
  tmpColHdrs(N + 1) = "Stage X"
  tmpColHdrs(N + 2) = "Stage Y"
  tmpColHdrs(N + 3) = "Stage Z"
  tmpColHdrs(N + 4) = "Qt1y"
  tmpColHdrs(N + 5) = "Qt1z"
  tmpColHdrs(N + 6) = "Primary Beam (na)" ' 09/06/11 -- added the " (na)"
  N = N + 6
End If

ReDim Preserve tmpColHdrs(1 To N)
End Sub

Function fiLettToNum%(ByVal Lett$)
' Convert a 1- or 2-character alphabetic index
'  to a number (e.g. "AB" to "27").
Lett = Trim(UCase(Lett))
If Len(Lett) = 2 Then
  fiLettToNum = Asc(Mid$(Lett, 2)) - 38
ElseIf Len(Lett) = 0 Then
  fiLettToNum = 0
Else
  fiLettToNum = Asc(Lett) - 64
End If
End Function

Function fsNumToLett$(ByVal Num)
' Convert an index-number to a 1- or 2-letter alphabetic index
If Num > 26 Then fsNumToLett = "A" & Chr(Num + 38) Else fsNumToLett = Chr(Num + 64)
End Function

Public Sub CheckNuclides(RatCtrl() As Control, MassCtrl() As Control, _
  GoBack As Boolean, Optional CheckMissingNuclides As Boolean = False, _
  Optional CheckAsters As Boolean = False)
' Do all numerator- and denominator-mass values correspond
'  to nuclides present in the current Task's Run Table?
Dim GotNu As Boolean, GotDe As Boolean
Dim s$, t$, m$, Msg$, Msg1$
Dim RatIndx%, NukeIndx%
Dim Nu#, de#

Msg1 = "Please either delete the ratio requiring a nuclide missing from the Run Table," _
    & vbLf & "or add the required nuclide to the Run Table before proceeding."
GoBack = False

For RatIndx = 1 To peMaxRats
  s = RatCtrl(RatIndx)

  If s <> "" Then

    If CheckAsters Then
      If InStr(s, "@@@") Then GoBack = True: Exit For
    End If

    If CheckMissingNuclides Then
      NumDenom s, Nu, de
      GotNu = False: GotDe = False

      For NukeIndx = 1 To peMaxNukes

        If MassCtrl(NukeIndx) = Nu Then
          GotNu = True
        ElseIf MassCtrl(NukeIndx) = de Then
          GotDe = True
        End If

        If GotNu And GotDe Then Exit For
      Next NukeIndx

      Msg = Msg1

      If Not GotNu Or Not GotDe Then

        If Nu = de Then
         Msg = "You can't ratio an isotope against itself,"
        ElseIf Not GotNu Then
          s = "@@@/" & fsS(de)
        Else
          s = fsS(Nu) & "/@@@"
        End If

        RatCtrl(RatIndx) = s
        GoBack = True
      End If

    End If

  End If

Next RatIndx
If GoBack Then MsgBox Msg, , pscSq
End Sub

Sub TimeWindowHandle(UsrForm As UserForm, ByVal NoAct As Boolean)
' Transfer Time-Window paramters from the Squid User sheet to the
'  Time Window User Form.
Dim b As Boolean
Dim SpotText$, TimeText$, StartTime$, EndTime$, StartSpot$, EndSpot$
Dim Startt$, Endd$, Nsp%, Clr&

Load TimeWindow
With TimeWindow
  If Not NoAct Then

    If foUser("startstd") = "" And foUser("endstd") = "" Then
      .oLower = False: .oUpper = False
      .oBoth = False:  .oNoBracket = True
      Nsp = .cmbStartStandard.ListCount
      Startt = .cmbStartStandard.List(0)
      Endd = .cmbStartStandard.List(Nsp - 1)
    End If

    .Show
  End If

  Startt = foUser("startstd"):         Endd = foUser("endstd")
  StartSpot = Trim(Left$(Startt, 17)): EndSpot = Trim(Left$(Endd, 17))
  StartTime = Mid$(Startt, 18):        EndTime = Mid$(Endd, 18)

  If foUser("oNoBracket") Or (Startt = "" And Endd = "") Then
    SpotText = "None - process all spots"
    TimeText = "(click button at right to define)"
  ElseIf foUser("oLower") Then
    SpotText = "Spots " & StartSpot & " to last spot"
    TimeText = StartTime & " to run end"
  ElseIf foUser("oUpper") Then
    SpotText = "First spot to " & EndSpot
    TimeText = "run start to " & EndTime
  ElseIf foUser("oBoth") Then
    SpotText = "Spots " & StartSpot & "  to  " & EndSpot
    TimeText = StartTime & "  to  " & EndTime
  End If

  b = (TimeText = ""): Clr = RGB(0, 0, 128)
  With UsrForm.lTimeWindow1
    .Caption = SpotText: .Top = IIf(b, 12, 7)
    .ForeColor = IIf(b, 0, Clr)
  End With
  With UsrForm.lTimeWindow2
    .Caption = TimeText:         .Top = 18
    .ForeColor = IIf(b, 0, Clr): .Visible = (TimeText <> "")
  End With
End With
End Sub

Sub TimeWindowParams(SpotNames$())
' Calculate starting- and ending-spots corresponding to the Time Window
'  parameters in the User sheet, and place them in the appropriate
'  Public variables.
Dim b As Boolean
Dim s1$, s2$, s1t$, s2t$
Dim ns%, piStartingSpot%, piEndingSpot%

ns = UBound(SpotNames)
s1 = foUser("startstd"): s2 = foUser("endstd")
s1t = Trim(Left$(s1, 17)): s2t = Trim(Left$(s2, 17))
piStartingSpot = 1: piEndingSpot = ns

If foUser("oNoBracket") Then
  foUser("startstd") = "":  foUser("endstd") = ""
  foUser("olower") = False: foUser("oUpper") = False: foUser("oboth") = False
  piStartingSpot = 1: piEndingSpot = ns
Else
  If foUser("oLower") Or foUser("oBoth") Then
    piStartingSpot = 0: b = False

    Do
      piStartingSpot = 1 + piStartingSpot
    Loop Until SpotNames(piStartingSpot) = s1t Or piStartingSpot >= (ns - 1 + 2 * pbUPb)

    foUser("startstd") = s1
  End If

  If foUser("oUpper") Or foUser("oBoth") Then

    piEndingSpot = piStartingSpot

    Do
      piEndingSpot = 1 + piEndingSpot
    Loop Until SpotNames(piEndingSpot) = s2t Or piEndingSpot = ns

    foUser("endstd") = s2
  End If

End If

If Not pbUPb Then
  piaStartSpotIndx(0) = piStartingSpot
  piaEndSpotIndx(0) = piEndingSpot
End If
End Sub

Sub ResetTrail()
' Reset the variables that track the last Task Editor panel
'  accessed.
With puTrail
  .bFromEquations = False: .bFromIsoRatios = False
  .bFromName = False:      .bFromRuntable = False
  .bFromUPbSpecial = False
End With
End Sub

Sub GetUPbPkOrders(ByVal Npks%, BkrdMass#, ByVal RefMass#, NominalMasses#())
Dim PkNum%, m# ' Determine values for the Run-Table order of the 4 crucial U-Pb peaks
piBkrdPkOrder = 0: piRefPkOrder = 0: pi204PkOrder = 0: pi206PkOrder = 0

For PkNum = 1 To Npks
  m = NominalMasses(PkNum)

  If m = BkrdMass Then
    piBkrdPkOrder = PkNum
  ElseIf m = RefMass Then
    piRefPkOrder = PkNum
  ElseIf m = 204 Then
    pi204PkOrder = PkNum
  ElseIf m = 206 Then
    pi206PkOrder = PkNum
  End If

Next PkNum
End Sub

Sub CheckParens(EqNum%, EqnCtrl() As Control, Bad As Boolean, Optional EqnNameCtrl, _
  Optional UPbEqns As Boolean = False)
' Validate a particular Task Equation during interaction with the Equation panel.
Dim Msg$, m1$, m2$, m3$, Es$, En$
Dim Rgtct%, LparenCt%, RbrakCt%, BrkErrType%, LgtCt%, LbrakCt%, RparenCt%
Dim Lgth%(), Rgth%(), Rparen%(), Lbrak%(), Rbrak%(), Lparen%()
Dim Eq As Control, EqNa As Control

Set Eq = EqnCtrl(EqNum)
If Not UPbEqns Then Set EqNa = EqnNameCtrl(EqNum): En = Trim(EqNa.Text)
Bad = True: Es = Trim(Eq.Text): Msg = ""

If Not UPbEqns And En = "" Then
  MsgBox "Please define a name for Equation" & StR(EqNum), , pscSq
  On Error Resume Next
  EqNa.SetFocus
  On Error GoTo 0
  Exit Sub
ElseIf InStr(Es, "@@@") Then
  Msg = "You must replace all @@@ with valid isotope-ratio references"
  GoTo BadExit
ElseIf InStr(Es, "//") > 0 Or InStr(Es, "**") > 0 Then
  Msg = "Equation contains illegal character pairs ( // or  **)"
  GoTo BadExit
End If

AllInstanceLoc "(", Es, Lparen, LparenCt
AllInstanceLoc ")", Es, Rparen, RparenCt

If LparenCt < RparenCt Then
  BrkErrType = 1
ElseIf RparenCt < LparenCt Then
  BrkErrType = 2
Else
  AllInstanceLoc "[", Es, Lbrak, LbrakCt
  AllInstanceLoc "]", Es, Rbrak, RbrakCt

  If LbrakCt < RbrakCt Then
    BrkErrType = 5
  ElseIf RbrakCt < LbrakCt Then
    BrkErrType = 6
  Else
    AllInstanceLoc "<", Es, Lgth, LgtCt
    AllInstanceLoc ">", Es, Rgth, Rgtct

    If LgtCt < Rgtct Then
      BrkErrType = 7
    ElseIf Rgtct < LgtCt Then
      BrkErrType = 8
    Else
      BrkErrType = 0
    End If

  End If

End If

If BrkErrType > 0 Then
  If BrkErrType Mod 2 = 0 Then m1 = "right" Else m1 = "Left"

  If BrkErrType < 3 Then
    m2 = "parenthesis"
  ElseIf BrkErrType = 4 Or BrkErrType = 4 Then
    m2 = "curly-brace"
    m3 = "The name of Equation" & StR(EqNum)
  ElseIf BrkErrType = 5 Or BrkErrType = 6 Then
    m2 = "square bracket"
  ElseIf BrkErrType = 7 Or BrkErrType = 8 Then
    m2 = "<>"
  End If

  If BrkErrType < 3 Or BrkErrType > 4 Then
    m3 = fsQq("Equation  $" & Es & "$ ")
  End If

  Msg = m3 & " is missing a " & m1 & " " & m2 & "."
  On Error Resume Next

  If UPbEqns Then
    Eq.SetFocus
  Else
    EqNa.SetFocus
  End If

  GoTo BadExit
End If

Bad = False
Exit Sub
BadExit: On Error GoTo 0
MsgBox Msg, , pscSq
End Sub

Sub CreateTotCtsColHdrArray()
' Create column headers for any "Total CPS" columns
Dim PkIndx%
piNtotctsHdrs = 0
With puTask

  For PkIndx = 1 To .iNpeaks

    If .baCPScol(PkIndx) Then
      piNtotctsHdrs = 1 + piNtotctsHdrs
      ReDim Preserve psaTotCtsHdrs(1 To piNtotctsHdrs)
      psaTotCtsHdrs(piNtotctsHdrs) = _
        "Total " & fsS(.daNominal(PkIndx)) & " cts/sec"
    End If

  Next PkIndx

End With
End Sub

Sub SimpleSubsetMatch(Nsubsets%, TaskName$, SubsetsLabel As Control)
' Assemble the string that will show the Subsets in the appropriate area
'  of the GeochronSetup or GenIsoStup panel.
Dim SubsetsString$, NameFrIndx%, NameFrCt%, ns%
ns = 0: SubsetsString = ""

With prSubsSpotNameFr
  NameFrCt = .Count

  For NameFrIndx = 1 To NameFrCt

    If Trim(.Cells(NameFrIndx, 1)) <> "" Then
      ns = 1 + ns
      SubsetsString = SubsetsString & Trim(.Cells(NameFrIndx, 1))
      SubsetsString = SubsetsString & "<" & fsS(NameFrIndx) & "> "
    End If

  Next NameFrIndx

End With
Nsubsets = ns
FormatSubsetLabel SubsetsString, Nsubsets, TaskName, SubsetsLabel
End Sub

Sub FormatSubsetLabel(Capt$, Nsubsets%, TaskName$, SubsetsLabel As Control)
Dim FontItalic As Boolean, s$, LabelForeClr&, LabelBackClr&
' Format the string that will show the Subsets in the appropriate area
'  of the GeochronSetup or GenIsoStup panel.
If Nsubsets = 0 Then
  s = "No Subsets"
  LabelForeClr = 0: LabelBackClr = peUformBclr
  pbNameFragsMatched = False: FontItalic = True
Else
  LabelForeClr = vbWhite: LabelBackClr = RGB(64, 64, 64)
  pbNameFragsMatched = True: FontItalic = False
  foUser("FragTaskType") = TaskName
End If

With SubsetsLabel
  .Caption = Capt: .Font.italic = FontItalic
  .ForeColor = LabelForeClr: .BackColor = LabelBackClr
End With
End Sub

Sub CheckUPbEqnRefs(Isorats$(), Missing As Boolean, EqnsIn$(), _
  TaskEqn() As Control, mbPbU As Boolean)
' Validate equations in the 'U-Pb Special" panel of the Task Editor.
Dim ChangedQ As Boolean, ChangedB As Boolean, Changed() As Boolean
Dim EqBrk$, Eq$, ExtrStr$, Msg$
Dim i%, j%, Indx%, ct%, IndxType%

ReDim Changed(piLwrIndx To peMaxEqns)

With puTask
  Missing = False

  For j = piLwrIndx To -1
    ChangedB = False: ChangedQ = False
    Eq = .saEqns(j)

    If Eq <> "" Then

      Do
        ExtractEqnRef Eq, ExtrStr, Indx, IndxType, Isorats, , False, , True

        If IndxType = peEquation And Indx <= .iNeqns Then
          EqBrk = "[" & ExtrStr & "]"

          If EqnsIn(Indx) = TaskEqn(Indx) Then
            Subst Eq, EqBrk, "<<<" & ExtrStr & ">>>"
          Else
            Subst Eq, EqBrk, "{{{}}}"
            ChangedB = True
          End If

        Else
          ExtrStr = ""
        End If

        Subst Eq, "<<<", "[": Subst Eq, ">>>", "]"
         ExtractEqnRef Eq, ExtrStr, Indx, IndxType, Isorats, .saEqnNames, True, , True

        If IndxType = peEquation Then
          EqBrk = psBrQL & ExtrStr & psBrQR

          If EqnsIn(Indx) = TaskEqn(Indx) Then
            Subst Eq, EqBrk, "<<<" & ExtrStr & ">>>"
          Else
            Subst Eq, EqBrk, "{{{}}}"
            ChangedQ = True
          End If

        Else
          ExtrStr = ""
        End If

        ct = 1 + ct
        If ct = 99 Then Error 9992
      Loop Until ExtrStr = ""

      Subst Eq, "<<<", psBrQL: Subst Eq, ">>>", psBrQR

      If ChangedB Or ChangedQ Then
        Subst Eq, "{{{}}}", "@@@"
        If ChangedB Or ChangedQ Then Changed(j) = True
      End If

    End If

  Next j

End With

For j = piLwrIndx To -1
  If Changed(j) Then Missing = True: Exit For
Next j

If Missing Then
  Msg = "Because you have changed or deleted one or more equations or equation names," _
    & vbLf & "you must modify the U-Pb equations panel to be consistent."
  MsgBox Msg, , pscSq
End If
End Sub

Sub FindCompatibleTask(TaskNumIn%, TryThisNameFirst$, NoAct As Boolean, _
  cbTask As Control, PksMatched As Boolean, NpksMatched As Boolean, _
  SortaMatched As Boolean, TaskName$, Npeaks%, Nomi#())
' 09/03/15 -- Delete subs of this name in GeochronSetup and GenIsoSetup, modify
'             to specifically pass Module variables from the original contexts.
' 09/03/17 -- Bracket "If .saNames(p, TaskNum) = TryThisNameFirst Then:TaskNum = 0: TryThisNameFirst = "":End If"
'             with "If i>0 .... End If"  to avoid subscr-out-of-range err when i=0.
' Using information in the Task Catalog, try to find a Task that is compatible with the
'  current PD or XML file's Run Table.

Dim RebuiltTaskCat As Boolean
Dim TaskNum%, j%, k%, p%, Ntasks%, RnumTest%

NoAct = False
p = -pbUPb
RebuiltTaskCat = False

With puTaskCat
  Ntasks = .iaNumTasks(p)

  For TaskNum = 0 To Ntasks

    If TaskNum = 0 Then

      For j = 1 To Ntasks
        If .saNames(p, j) = TryThisNameFirst Then
          TaskNum = j: Exit For
        End If
      Next j

    End If

    If TaskNum > 0 Then
      RnumTest = TaskNum
      Npeaks = .iaNpeaks(p, RnumTest)

      Do While Npeaks < 2
        If RebuiltTaskCat Then
          MsgBox "The file for Task " & fsInQ(.saNames(p, TaskNum)) & " has #peaks=0. " _
            & vbLf & "Please repair or delete.", , pscSq
          End
        Else
          BuildTaskCatalog
          RebuiltTaskCat = True
        End If
      Loop

      ReDim Nomi(1 To Npeaks)

      For k = 1 To Npeaks
        Nomi(k) = .daNominalMass(p, k, RnumTest)

        Do While Nomi(k) <= 0
          If Not RebuiltTaskCat Then
            BuildTaskCatalog
            RebuiltTaskCat = True
          Else
            MsgBox "The file for Task " & fsInQ(.saNames(p, TaskNum)) & _
              " has Nominal Mass #" & fsS(k) & "=0. " & vbLf & _
              "Please repair or delete.", , pscSq
            End
          End If
        Loop

      Next k

      PeaksMatched RnumTest, True, NpksMatched, Npeaks, Nomi, SortaMatched, PksMatched
  If PksMatched Then Exit For

    End If

    If TaskNum > 0 Then
      If .saNames(p, TaskNum) = TryThisNameFirst Then
        TaskNum = 0: TryThisNameFirst = ""
      End If
    End If
  Next TaskNum

End With

If PksMatched Then
  TaskNumIn = RnumTest
Else
  TaskNumIn = fvMin(fvMax(TaskNumIn, 1), Ntasks)
End If

TaskName = puTaskCat.saNames(p, TaskNumIn)
NoAct = True
cbTask.ListIndex = TaskNumIn - 1
NoAct = False
' cbTask_Change must follow in calling context
End Sub

Sub PeaksMatched(TaskNum%, Initializing, NpksMatched As Boolean, _
  Npeaks%, Nomi#(), SortaMatched As Boolean, PksMatched As Boolean)
' 09/03/15 -- Delete subs of this name in GeochronSetup and GenIsoSetup, modify
'             to specifically pass Module variables from the original contexts.
' Do the number of peaks and mass values of a Task match that of the loaded
'  PD or XML file?
Dim FromSheet As Boolean, MassesMatched As Boolean
Dim PkNum%, Mdelt#, MaxMdelt#, MaxPpmMassDelt#

NpksMatched = (Npeaks = piFileNpks)

If NpksMatched Or Initializing Then

  CalcNominalMassVal True
  FromSheet = fbIsCondensedSheet
  If FromSheet Then GetCondensedShtInfo

  If NpksMatched Then
    MaxMdelt = 0: MassesMatched = True
    MaxPpmMassDelt = foUser("MaxPpmMassDelt")

    For PkNum = 1 To piFileNpks

      If pdaFileNominal(PkNum) > 0 And Nomi(PkNum) > 0 Then
        Mdelt = Abs(pdaFileNominal(PkNum) / Nomi(PkNum) - 1) * pdcMillion
        If MaxMdelt < Mdelt Then MaxMdelt = Mdelt
      End If

    Next PkNum

    If MaxMdelt > MaxPpmMassDelt Then MassesMatched = False
    SortaMatched = (MaxMdelt < 100000)
  End If

End If

PksMatched = (NpksMatched And MassesMatched)
End Sub

Sub GetFileInfoForUPb(HasU As Boolean, HasTh As Boolean, Has206 As Boolean, _
    Has207 As Boolean, Has208 As Boolean, Has204 As Boolean)
Dim PkNum%

For PkNum = 1 To piFileNpks
  Select Case CInt(pdaFileNominal(PkNum)) ' 09/06/12 -- add the Cint
    Case 204:           Has204 = True
    Case 206:           Has206 = True
    Case 207:           Has207 = True
    Case 208:           Has208 = True
    Case 232, 248:      HasTh = True
    Case 238, 254, 270: HasU = True
  End Select
Next PkNum

End Sub
