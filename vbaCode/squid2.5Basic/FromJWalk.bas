Attribute VB_Name = "FromJWalk"
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
Option Explicit
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'32-bit API declarations
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) _
  As Long

Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function PlaySound Lib "winmm.dll" _
  Alias "PlaySoundA" (ByVal lpszName As String, _
  ByVal hModule As Long, ByVal dwFlags As Long) As Long

Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000

Sub CrashNoise()
'PlayWAV "sound240.wav"
End Sub

Sub PlayWAV(wf)
Dim WAVFile As String, PA$
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
PA = "C:\Documents and Settings\KRL\My Documents\FromSony\SquidFolders\noise\"
WAVFile = PA & wf
'WAVFile = "dogbark.wav"
'WAVFile = "c:\spam2.wav" 'twb.path & "\" & WAVFile
Call PlaySound(WAVFile, 0&, SND_ASYNC Or SND_FILENAME)
End Sub

Sub ClearAllClipboard()
' from http://www.tek-tips.com/viewthread.cfm?qid=1227582&page=1
  OpenClipboard 0&
  EmptyClipboard
  CloseClipboard
End Sub

Sub Alarm() 'Cell, Condition)
Dim WAVFile As String
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
On Error GoTo ErrHandler
'If Evaluate(Cell.Value & Condition) Then
    WAVFile = "c:\sound240.wav" 'Edit this statement
    Call PlaySound(WAVFile, 0&, SND_ASYNC Or SND_FILENAME)
    'Alarm = True
    Exit Sub
'End If
ErrHandler:
'Alarm = False
End Sub
