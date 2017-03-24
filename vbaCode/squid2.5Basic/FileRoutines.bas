Attribute VB_Name = "FileRoutines"
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
' Module FileRoutines

' 09/03/01 -- all lower array bounds explicit
' 09/04/13 -- Add a Minimum Window parameter (for LOWESS smoothing) to the User sheet.
Option Explicit
Option Base 1

Sub LoadStorePrefs(ByVal Load1Store2%)
' Load or Store the SquidPreferences workbook and copy the User worksheet to or from
'  (respectively) the internal User worksheet of this add-in.
' 09/05/06 -- If no SquidUser file exists in the folder containing Squid2.xla, ask
'             user if SQ2 should create a new one with default Tasks and Prefs.
Dim bStore As Boolean, Bad As Boolean, FileDateOK As Boolean, FileExists As Boolean, tB As Boolean
Dim PreferencesFilePath$, PreferencesFilename$, SquidPath$, Msg$, tmp$, DefaultTask$
Dim TaskFileNa$
Dim j%, Co%, ErrNum%, Tnum%, Ttype%, Tmsg%, Rw&, rw1&, rw2&
Dim Sqsh As Worksheet, TmpSht As Worksheet
Dim Urange As Range
Dim Shp As Shape
Dim Nwb As Workbook
Dim UserVersion As Variant

bStore = (Load1Store2 = 2)
NoUpdate
Alerts False
PreferencesFilename = "SquidPreferences.xls"
' Default path for Squid user files is the folder in which Squid*.xla resides.
SquidPath = ThisWorkbook.Path
ChDirDrv fsSquidUserFolder, Bad, True, ErrNum
If Not Bad Then Bad = Dir(fsSquidUserFolder) = ""

If Bad Then
  Tmsg = MsgBox("There doesn't seem to be a non-empty SquidUser folder in the " _
    & "SQUID 2 program folder  (" & SquidPath & ")." & pscLF2 & _
    "Would you like SQUID to create a new SquidUser folder with " & _
    "default Tasks and Preferences?", vbYesNo, pscSq)
    
  If Tmsg = vbYes Then ' Make a new SquidUser folder & fill it with default files.
    MkDir fsSquidUserFolder
    ChDir fsSquidUserFolder
    foAp.CutCopyMode = False
    foUser.Copy
    ActiveWorkbook.SaveAs FileName:=PreferencesFilename
    ActiveWorkbook.Close
    
    For Ttype = 1 To 2
      Tnum = 0
      tmp = Choose(Ttype, "UPb", "Gen") & "Task ("
      
      Do
        DefaultTask = tmp & fsS(Tnum + 1) & ")"
        StatBar "saving " & DefaultTask
        Set TmpSht = Nothing
        On Error Resume Next
        Set TmpSht = ThisWorkbook.Sheets(DefaultTask)
        On Error GoTo 0
        
        If TmpSht Is Nothing Then
          Exit Do
        Else
          GetTaskRows
          TmpSht.Copy
          ActiveSheet.Name = "Task"
          TaskFileNa = Cells(puTask.lFileNameRw, peTaskValCol)
          ActiveWorkbook.SaveAs FileName:=fsSquidUserFolder & TaskFileNa
          ActiveWindow.Close
          Tnum = 1 + Tnum
        End If
      Loop
      
    Next Ttype
    
    StatBar
    
  Else
    End
  End If
  
End If

If Dir(fsSquidUserFolder) = "" Then
  MsgBox "No files exist in the SquidUser folder", , pscSq
  End
End If

ChDirDrv fsSquidUserFolder
PreferencesFilePath = fsSquidUserFolder & PreferencesFilename

If bStore Then
  ' Copy the User sheet from this add-in and store as the SquidPreferences workbook
  On Error GoTo 4
  foAp.CutCopyMode = False
  foUser.Copy
  ActiveWorkbook.SaveAs FileName:=PreferencesFilePath
  On Error GoTo 0
  ActiveWorkbook.Close
  Exit Sub
4: On Error Resume Next
  ActiveWorkbook.Close
  MsgBox "Unable to re-store Squid Preferences file", , pscSq
  GoTo 1
  
Else ' Open the SquidPreferences workbook and copy its User worksheet to
     '  this add-in's User worksheet.
  FileDateOK = False
  OpenWorkbook PreferencesFilePath, FileExists
  
  If FileExists Then
    On Error GoTo DateNotOK
    UserVersion = Val(fhSquidSht.[Version])
    If fbRangeNameExists("OldestAcceptableRevdate") Then
      FileDateOK = (UserVersion >= Val([OldestAcceptableRevdate]))
    End If
DateNotOK:
    On Error GoTo 0
  End If
  
  If Not FileDateOK Then
  
    If Not FileExists Then
      Msg = "No Preferences file in the SquidUser folder -- using default."
    Else
      Msg = "Updating Preferences file for compatibility with this SQUID version" _
           & vbLf & "                      (please re-specify your Preferences)"
      ActiveWorkbook.Close
    End If
    
    MsgBox Msg, , pscSq
    Set Sqsh = foUser
  Else
    Set Sqsh = ActiveWorkbook.ActiveSheet
    Set Nwb = ActiveWorkbook
    On Error GoTo 6
    
    For Each Shp In Nwb.Sheets(1).Shapes
      If Left$(Shp.Name, 7) <> "Comment" Then Shp.Delete
      ' otherwise end up with replicate shapes in foUser
    Next Shp
    
    Nwb.Sheets(1).Cells.Copy Destination:=foUser.Cells(1, 1)
    On Error Resume Next
    Nwb.Close
    
    ' The ranges below may have been added to the code since the stored
    '  Preferences file was created.  If so, create the ranges and put
    '  default values in them.
    AddUserRange "SecularTrend", False, 1, 3
    AddUserRange "SmoothingWindow", 10, 1, 1
    AddUserRange "AttachTask", False, 1, 1
    AddUserRange "DatRedParamsSeparate", False, 1, 1
    AddUserRange "SeparateAutochtSht", False, 1, 1
    AddUserRange "LinFitSpecial", True, 23, 1
    AddUserRange "LinFitRats", True, 23, 1
    AddUserRange "LinFitEqns", True, 23, 1
    AddUserRange "ZeroYmin", True, 1, 1
    AddUserRange "LongCondensed", True, 1, 1
    AddUserRange "OldestAcceptableRevdate", 2.31, 1, 1
    AddUserRange "AutoWindow", False, 1, 1
    AddUserRange "MinWindow", 4, 1, 1
    AddUserRange "Splash", True, 1, 1
    AddUserRange "NoUPbConstAutoreject", False, 1, 1
    AddUserRange "CompareGroupedUOUwithStd", False, 1, 1
    AddUserRange "LastPreferencesPage", 0, 1, 1
    AddUserRange "NoUThConcStd", False, 1, 1
    AddUserRange "Corr7ThPb", False, 1, 1
    'AddUserRange "GrpSKageType", 1, 1, 1            ' 09/12/06 commented out
    AddUserRange "Calc7corrPbThages", False, 1, 1
    AddUserRange "GrpCalc7corrPbThages", False, 1, 1
    AddUserRange "CalcFull8corrErrs", False, 1, 1
    AddUserRange "Corr8PbPb", False, 1, 1
    AddUserRange "LinFitRatsDiff", True, 23, 1       ' added 09/06/18
    AddUserRange "CPbSKage", True, 1, 0              ' added 09/10/08
    AddUserRange "CPbSpecType", True, 1, 1           ' added 09/10/08
    AddUserRange "ExtractAgeGroups", True, 1, 1      ' added 09/12/06
    AddUserRange "GrpCommPbSpecific", False, 1, 1    ' added 09/12/06
    AddUserRange "ExtractSpotNameGroups", True, 1, 1 ' added 09/12/06
    AddUserRange "RememberGroupNchars", False, 1, 1  ' added 10/11/18
    
    If Not fbRangeNameExists("ThPbStdAges", ThisWorkbook, foUser) Then
      With foUser
        With .[UPbStdAges]
            Co = .Column
            rw1 = .Row - 1
            rw2 = .Row + .Rows.Count - 1
            Rw = rw2 + 2
        End With
        .Range(frSr(rw1, Co, rw2).Address).Copy .Cells(Rw, Co)
        .Cells(Rw, Co) = "ThPbStdAges"
        With .Range(frSr(Rw + 1, Co, Rw + rw2 - rw1).Address)
          .Name = "ThPbStdAges"
          .ClearContents
        End With
      End With
    End If
    
  End If
  GoTo 7
6: On Error GoTo 0
  MsgBox "Unable to re-store User Preferences"
  GoTo 3
7: On Error GoTo 0

  If Not fbRangeNameExists("IgPeriods", ThisWorkbook, foUser) Then
    Set Urange = [AU3:AV6]
    ' In case a very old version of Preferences was stored.
    For j = 1 To 4
      Urange(j, 1) = Array("IgPeriods", "IgCommas", "IgColons", "IgSemicolons")(j)
      Urange(j, 2).Name = Urange(j, 1)
      Urange(j, 2) = False
      Urange(j, 1).HorizontalAlignment = xlRight
    Next j
    
  End If
  
3: On Error Resume Next
  With foUser
    .[LastUseDateThisUser] = Date
    .[sqDateVerUser] = fhSquidSht.[Version]
  End With
End If
Exit Sub

1: On Error GoTo 0
End Sub

Sub AddUserRange(RangeName$, RangeVal As Variant, ColNum%, RwOffset%)
Dim Rw&, rw1&, rw2&
foUser.Activate
On Error GoTo 1

If Not fbRangeNameExists(RangeName, ThisWorkbook, foUser) Then
  rw1 = flEndRow(ColNum, ThisWorkbook, foUser)
  rw2 = flEndRow(ColNum + 1, ThisWorkbook, foUser)
  Rw = RwOffset + fvMax(rw1, rw2)
  Box Rw, ColNum, , , vbYellow, , , ThisWorkbook, foUser
  With Cells(Rw, ColNum)
    .Name = RangeName
    .Formula = RangeVal
    .HorizontalAlignment = xlCenter
  End With
  Cells(Rw, 1 + ColNum) = RangeName
End If

1: On Error GoTo 0
End Sub

Sub GetRawData(Bad As Boolean)
' User selects a PD or XML file to load
Dim Mtype$, Origin&

Mtype$ = "Raw-data files (*.pd; *.txt; *.xml),**.pd;.txt;.xml"
Origin = xlWindows
Alerts False
NoUpdate
Bad = False

psRawFileName = foAp.GetOpenFilename(FileFilter:=Mtype$, _
  Title:="Select SHRIMP Datafile to open:")
  
If psRawFileName = "False" Then Exit Sub

pbXMLfile = UCase(Right$(psRawFileName, 4)) = ".XML"
pbPDfile = Not pbXMLfile
End Sub

Sub GetFile(RawDat() As RawData, FileLines$(), Optional Bad As Boolean = False)
' Displays the File-open dialog box, parses the resulting loaded PD/XML file.
Do

  Do
    GetRawData Bad
    If psRawFileName = "False" Then End
  Loop Until Not Bad
  
  piNumAllSpots = 0
  InhaleRawdata pbPDfile, FileLines, Bad
  If Bad Then Exit Sub
  
  If pbXMLfile Then
    ParseXML RawDat, FileLines, UBound(FileLines), Bad, ""
  ElseIf pbPDfile Then
    ParsePD RawDat, FileLines, UBound(FileLines), Bad, ""
  End If
  
  If Bad Then
    MsgBox "This file is not a valid SHRIMP PD or XML file.", , pscSq
    If Workbooks.Count > 0 Then
      With ActiveWorkbook
        If .Name <> ThisWorkbook.Name Then .Close
      End With
    End If
    Exit Sub
  End If
  
Loop Until Not Bad
  
CalcNominalMassVal
End Sub

Function fsPathSep$() ' returns the path separator for Windows
fsPathSep = Application.PathSeparator
End Function

Function fsGetDirectory$(Optional Msg)
' Allows user to specify a folder
Dim Path As String, pos%, r&, X&
Dim bInfo As BROWSEINFO

'   Root folder = Desktop
bInfo.pidlRoot = 0&
'   Title in the dialog
If IsMissing(Msg) Then
    bInfo.lpszTitle = "Select a folder."
Else
    bInfo.lpszTitle = Msg
End If
'   Type of directory to return
bInfo.ulFlags = &H1
'   Display the dialog
X = SHBrowseForFolder(bInfo)
'   Parse the result
Path = Space$(512)
r = SHGetPathFromIDList(ByVal X, ByVal Path)
If r Then
  pos = InStr(Path, Chr$(0))
  fsGetDirectory = Left$(Path, pos - 1)
Else
  fsGetDirectory = ""
End If
End Function

Sub GetFileInfo(PathFile$, Exists As Boolean, Optional LastModified As Date)
' Determines if a file exists and its last-modified date
Dim Sfo As Object, GotFile As Object
 
' Not reliable??
Set Sfo = CreateObject("Scripting.FileSystemObject")

If Right$(LCase(PathFile), 4) <> ".xls" Then
  PathFile = PathFile & ".xls"
End If

Exists = Sfo.FileExists(PathFile)

If Exists Then
  Set GotFile = Sfo.GetFile(PathFile)
  LastModified = FileDateTime(GotFile)
End If
End Sub

Sub OpenWorkbook(ByVal Filespec$, Optional Exists As Boolean, _
                Optional ShowWarning As Boolean = False)
' Open a workbook
GetFileInfo Filespec, Exists
On Error GoTo 1
If Exists Then Workbooks.Open Filespec
Exit Sub
1: Exists = False
If ShowWarning Then MsgBox "Unable to open " & fsInQ(Filespec), , pscSq
End Sub

Sub ChDirDrv(ByVal FullPath$, Optional Bad = False, _
             Optional NoErr As Boolean = True, Optional ErrNumber)
' Error-trapped ChangeDrive procedure
Dim Drive$
Bad = True
VIM ErrNumber, 0
Drive = fsCurDrive(FullPath)
On Error GoTo 1
ChDrive Drive
On Error GoTo 2
ChDir FullPath
Bad = False
Exit Sub

1: On Error GoTo 0
If Not NoErr Then MsgBox "Unable to access drive " & Drive, , pscSq
Exit Sub
2: If fbNIM(ErrNumber) Then ErrNumber = Err.Number
On Error GoTo 0
If Not NoErr Then MsgBox "Unable to access path " & FullPath, , pscSq
End Sub

Function SheetExists(ByVal Sheetname$, Optional Wkbook)
Dim Sht As Worksheet ' Determine if a worksheet exists in the active workbook
VIM Wkbook, ActiveWorkbook

For Each Sht In Wkbook.Worksheets
  If LCase(Sht.Name) = LCase(Sheetname) Then
    SheetExists = True: Exit Function
  End If
Next Sht

SheetExists = False
End Function

Function fbIsCondensedSheet(Optional WkSht) As Boolean
Dim Rw&, i%
' Determine if a worksheet is a SQUID-condensed PD or XML file
fbIsCondensedSheet = False
On Error GoTo 1
VIM WkSht, ActiveSheet

If Workbooks.Count > 0 Then

  If WkSht.Type = xlWorksheet Then
  
    If WkSht.Cells(1, peReadyCol) = "squid ready" Then
      WkSht.Activate
      fbIsCondensedSheet = True
      Set phCondensedSht = ActiveSheet
      Set pwDatBk = ActiveWorkbook
    End If
    
  End If
  
End If

1: On Error GoTo 0
End Function

Sub FileNamesInDir(Folder$, FileNames$(), FileCt%, OkExtensions%)
' Fill an array with the filenames in a specified folder that
'  have the extensions specified by the OkExtensions parameter.

' OkExtensions% values:
' 0 - all extensions
' 1 - XLS only
' 2 - PD and XML only
' 3 - PD only
' 4 - XML only

Dim Ext$, FileNa$, p%, Le%, OkX%, ct%, FolderIn$, DriveIn$
OkX = OkExtensions

If OkX < 0 Or OkX > 4 Then
  MsgBox "Code error:   OkExensions%  passed as" & StR(OkX) & " in Sub FileNames", , pscSq
  End
End If

DriveIn = fsCurDrive
FolderIn = CurDir
ChDirDrv Folder
FileCt = 0: ct = 0

Do

  If ct = 0 Then
    FileNa = Dir("")
  Else
    FileNa = Dir
  End If
  
  ct = 1 + ct
  Le = Len(FileNa)
If Le = 0 Then Exit Do
  p = InStr(StrReverse(FileNa), ".")
  Ext = LCase(Mid$(FileNa, Le - p + 2))
  
  If OkX = 0 Or (OkX = 1 And Ext = "xls") Or (OkX = 2 And (Ext = "pd" Or Ext = "xml")) Or _
    (OkX = 3 And Ext = "pd") Or (OkX = 4 And Ext = "xml") Then
    FileCt = 1 + FileCt
    ReDim Preserve FileNames(1 To FileCt)
    FileNames(FileCt) = FileNa
  End If
  
Loop

ChDirDrv FolderIn
End Sub

Function fbFileNameExists(Folder$, FileName$) As Boolean
' Simple determination of existence of a file
Dim FileNames$(), FileCt%, FileNum%

FileNamesInDir Folder, FileNames, FileCt, 0

For FileNum = 1 To FileCt
  If LCase(FileName) = LCase(FileNames(FileNum)) Then Exit For
Next FileNum

fbFileNameExists = (FileNum <= FileCt)
End Function

Sub DiscardExistingWkShts(AllowDelete As Boolean)
Dim DidDelete As Boolean, DoDelete&, Msg$, i%, Co%, Rw&
Dim ShtArr As Variant, TmpArr As Variant, Sht As Worksheet

TmpArr = Array("Standard and Sample worksheets", "Reduced-data worksheet")
ShtArr = Array("Within-Spot Ratios", "Trim Masses", _
        "Autocharts", "Task", "Data-reduction params")
        
Msg = "Discard existing " & TmpArr(2 + pbUPb) & " before processing?"
Alerts False
DoDelete = vbNo: AllowDelete = True
foAp.Calculation = xlManual

Do
  DidDelete = False
  
  For Each Sht In ActiveWorkbook.Worksheets
    With Sht
    
      If .Name = pscStdShtNa Or .Name = pscSamShtNa Then
        GoSub DeleteSheet
        Exit For
      Else
      
        For i = LBound(ShtArr) To UBound(ShtArr)
          If LCase(.Name) = LCase(ShtArr(i)) Then GoSub DeleteSheet
          If DidDelete Then Exit For
        Next i
        If DidDelete Then Exit For
        
        On Error GoTo 2
        Sht.Activate
        On Error GoTo 0
        FindStr "Spot Name", Rw, , 1, 1, 10
        
        If Rw > 0 Then ' If a Grouped Sample sheet, delete
          ' 09/10/15 -- Line below - changed "RowLook1:" & "ColLook1:" to unpassed params
          FindStr "SQUID grouped-sample sheet", Rw, Co, , , 1, 255
          If Rw > 0 Then GoSub DeleteSheet
          If DidDelete Then Exit For

        End If
        
      End If
      
    End With
  Next Sht
  
2  On Error GoTo 0
Loop Until Not DidDelete

Alerts True
Exit Sub

DeleteSheet:
If DoDelete = vbNo Then DoDelete = MsgBox(Msg, vbYesNoCancel, pscSq)
If DoDelete = vbCancel Then
  End
ElseIf DoDelete = vbYes Then
  Sht.Delete: DidDelete = True: DoDelete = vbYes
Else
  AllowDelete = False
  Exit Sub
End If

Return
End Sub

Sub CorrectEqnRefs(Phrase$, Optional IsUPb As Boolean = False)  ' 10/04/20 -- added
Dim p%, q%, NoApost$, WithApost$, s$, ss$, StdNot$, SamNot$, RefNa$, Nrefs%, i%
Dim LegalEqNa$, LegalRef$
' Remove apostrophes from workbook/worksheet references, e.g. '[WbkName]WkshtName'!
'  changed to [WbkName]WkshtName!
' If IsUPb then eliminate refs to Std & Sam wksht (valid if ref is a range name).
SamNot = "'" & LCase(pscSamShtNa) & "'!"
StdNot = "'" & LCase(pscStdShtNa) & "'!"
s = (Phrase)

Do
  ' Don't need explicit samsht/stdsht refs since ref is a named range
  If IsUPb Then Subst s, LCase(StdNot), , LCase(SamNot)
  NoApost = fsExtractPart(s, 1, "'", "'!")
  WithApost = "'" & NoApost & "'!"
  ss = s
  Subst ss, WithApost, NoApost & "!"
If ss = s Then Exit Do
  s = ss
Loop

'  All worksheet/workbook refs simplified - now
'  replace single-cell ["NA"] eqn refs with NA.

Do
  RefNa = "": ss = s
  SqBrakQuExtract s, RefNa$ ', Nrefs
  RefNa = Trim(LCase(RefNa))
  
  If RefNa <> "" Then
    
    If Not fbIsNumChar(RefNa) Then ' i.e. if not an isotope-ratio reference
    
      LegalRef = fsLegalName(LCase(RefNa))
      Subst s, RefNa, LegalRef
      
      With puTask
      
        For i = 1 To .iNeqns
          LegalEqNa = fsLegalName(LCase(.saEqnNames(i)))
          
          If LegalRef = LegalEqNa Then          ' Is the reference to the ith equation?
            If .uaSwitches(i).SC Then Exit For  ' And is this a single-cel (SC) equation "
          End If
          
        Next i
        
        If i > .iNeqns Then Exit Do              ' No, it's not
      End With
      
      Subst s, psBrQL & RefNa & psBrQR, RefNa    ' The ref is to a single-cell equation,
                                                 '   so remove it from ["  "] brackets
    End If                                       '   (as the output cell will be a named range).
    
  End If
  
Loop Until ss = s

Phrase = s
End Sub
