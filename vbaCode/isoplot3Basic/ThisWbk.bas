Attribute VB_Name = "ThisWbk"
Option Private Module
Option Base 1: Option Explicit
Const IsoCap = "Iso&plot", StrkThrough = 290, NumB% = 4 ', CantDeleteIfInvokedFrom = -2147467259
Public Starting As Boolean, TT$(NumB), IsoToolbars As Range

Sub Workbook_Open()
Attribute Workbook_Open.VB_ProcData.VB_Invoke_Func = " \n14"
' Add an "Isoplot" dropdown menu to the Worksheet and Charts menu-bar
Dim i%, j%, k%, s$, r$, T$, Missing As Boolean, eR& ', IsoTB$(3, 2)
Dim b As Object, cf As Object, d As Object, e As Object, tCt%(2), tB1 As Boolean
Dim tr As Range, Std As Object, Got As Boolean, Twn$, TmpB As Boolean, bType%, NumDel%, TotToDel%
Dim IndX%(), Bct%, iCt%(2), DidDel%, P%, q%, Xct%, CbCt%, tB As Toolbar, TbBtn As ToolbarButton
Dim Tbars As Toolbars, Cbars As CommandBars, Cbar(2) As CommandBar, Tbar(3) As Control, BarBtn As Control
Dim MenuBarCtrlCt%, IsoButtonPos%, IsoTbars(3) As CommandBar
Const Ip = "iso*.xl?"
Starting = True
'i = 1 / 0
GetOpSys
Workbook_BeforeClose True
Twn$ = "'" & TW.Name & "'!"
If LCase(Right$(TW.Name, 4)) = ".xla" Then GetMenuSheet
With Application
  Set Cbars = .CommandBars
  Set Tbars = .Toolbars
End With
Set Cbar(1) = Cbars("worksheet menu bar")
Set Cbar(2) = Cbars("chart menu bar")
Set IsoToolbars = Menus("isotoolbars")
For i = 1 To 3
  s = IsoToolbars(i, 3).Text
  Set IsoTbars(i) = Cbars(s)
Next i
MenuBarCtrlCt = Cbar(1).Controls.Count
IsoButtonPos = Min(MenuBarCtrlCt, 10)
Cbars("Isoplot 3 Charts & Isochrons").Controls("invokeisoplot") _
  .Copy Bar:=Cbars("Worksheet Menu Bar"), Before:=IsoButtonPos
On Error Resume Next
'Cbars("Isoplot 4 Charts & Isochrons").Delete
'Cbars("Isoplot 4 Worksheet Tools & Isochrons").Delete
'Cbars("Isoplot 4 Chart Tools").Delete
For Each tB In Tbars ' Make all isoplot toolbars refer to this add-in
  s = tB.Name
  If InStr(LCase(s), "isoplot") > 0 Then
    For Each TbBtn In tB.ToolbarButtons
      s = TbBtn.OnAction
      k = InStr(s, "!")
      If k > 0 Then
        s = Mid(s, 1 + k)
        TbBtn.OnAction = s
      End If
    Next
  End If
Next
On Error GoTo 0
' Make activesheet a worksheet if not already.
For i = 1 To Workbooks.Count
  If Workbooks(i).Name <> ThisWorkbook.Name Then
    Workbooks(i).Activate
    If ActiveSheet.Type <> xlWorksheet Then
      For j = 1 To Sheets.Count
        With Sheets(j)
          If .Type = xlWorksheet And .Visible Then
            Sheets(j).Activate
            Exit For
          End If
        End With
      Next j
      If j > Sheets.Count Then Sheets.Add
    End If
  End If
Next i

Starting = False
StPc("cShapes").Visible = True
'On Error GoTo NoCanDo
For i = 2 To 1 Step -1
  Set tr = Menus(Choose(i, "Worksheet", "Chart") & "Menubar")
  tCt(i) = Cbars(i).Controls.Count
  For k = 1 To Cbars.Count
    TmpB = False
    With Cbars(k)
      If LCase(.Name) = LCase(tr(0, 1).Text) Then
        'On Error Resume Next
        For Each d In .Controls ' Make sure any Isoplot button refers to this program
          If d.Type = msoControlButton Then
            If d.TooltipText = "Run ISOPLOT" Then
              d.OnAction = "StartFromInit" 'Iso
            Else ' Parse out path and filename of macro, delete so macro refers to this worbook
              s$ = LCase(d.OnAction): r = RevStr(s$)
              P = InStr(r$, "!") - 1: q = InStr(r$, PathSep) - 1
              If P > 0 And q > 0 And (q - P - 2) > 0 Then
                T$ = Left(Right(s$, q), q - P - 2)
                On Error Resume Next
                If T$ Like Ip Then d.OnAction = Right(s$, P)
                On Error GoTo 0
              End If
            End If
          End If
        Next d
        .Visible = True
        'On Error GoTo NoCanDo
        ' Add the main Isoplot menu for either Worksheet or Chart sheet
        With .Controls.Add(Type:=msoControlPopup, Before:=IsoButtonPos, Temporary:=True)
          iCt(i) = .Index
          .Caption = IsoCap: j = 0
          Do
            j = 1 + j
            If Not IsEmpty(tr(j, 1)) Then
              ' Add an item in the Isoplot dropdown
              bType = IIf(tr(j, 1).Font.Bold, msoControlPopup, msoControlButton)
              With .Controls.Add(Type:=bType, Temporary:=True)
                .Caption = tr(j, 1)
                .BeginGroup = (tr(j, 1).Font.Underline = xlSingle)
                If bType = msoControlPopup Then ' Add a daughter menu
                  Do
                    With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                      .Caption = tr(j, 2): .OnAction = tr(j, 3)
                    End With
                  If IsEmpty(tr(j + 1, 2)) Or j = tr.Count Then Exit Do
                    j = 1 + j
                  Loop
                Else
                  .OnAction = tr(j, 3)
                End If
              End With
            End If
          Loop Until j = tr.Count / 2
        End With
        TmpB = True
      End If
    End With
    If TmpB Then Exit For
  Next k
Next i
Set d = Application.Toolbars
On Error Resume Next
'For i = 1 To 3
'  IsoTbars(i).Delete
'Next
'd("Isoplot Worksheet Tools").Delete
'd("Isoplot Charts & Isochrons").Delete
'd("Isoplot Chart Tools").Delete
'd("Isoplot 3 Tools").Delete
For Each d In Cbars(1).Controls
  If d.Caption = "&Tools" Then
    For Each e In d.Controls
      If e.Caption = "&Delete unused Isoplot hidden-sheets" Then
        e.Delete
      End If
    Next e
  End If
Next d
j = 1 - 2 * Mac    ' Either just before the Isoplot dropdown or @ end of the Std toolbar
On Error Resume Next
On Error GoTo 0
Set cf = Cbars("Formatting").Controls: Set Std = Cbars("Standard")
Missing = True ' Add strikethru button if missing
For Each b In cf
  If b.ID = StrkThrough Then Missing = False: Exit For
Next
If Missing Then cf.Add Type:=msoControlButton, ID:=StrkThrough, Before:=6
If Not MacExcelX Then
  On Error Resume Next
  For i = 1 To 3
    With Menus("IsoToolbars")
        Cbars(.Cells(i, 3).Text).Visible = .Cells(i, 2)
    End With
  Next i
  j = Cbars(1).Controls("invokeIsoplot").Index
  With Cbars("Isoplot 3 Worksheet Tools").Controls
    For i = 1 To .Count
      s = LCase(.Item(i).OnAction)
      k = InStr(s, "!")
      If k > 0 Then s = Mid(s, 1 + k)
      If s = "startfromiso" Then 'nit" Then
        .Item(i).Copy Bar:=Cbar(1), Before:=1 + j
        Exit For
      End If
    Next i
  End With
  For i = 1 To 3
    With IsoTbars(i)
      On Error Resume Next
      .Visible = IsoToolbars(i, 2)
      .Position = IsoToolbars(i, 5)
      .Left = IsoToolbars(i, 6)
      .Top = IsoToolbars(i, 7)
      On Error GoTo 0
    End With
  Next i
  On Error GoTo 0
End If
Application.DisplayAlerts = True
EmptyClipboard
'If MacExcelX Then
'  MsgBox "With Excel-X under OS-X, Isoplot3.xla must be un-installed before closing Excel." _
'    & vbLf & vbLf & "If you forget to do this, Excel will crash when re-opened unless " _
'    & "Isoplot3.xla is temporarily moved to a different folder.", vbOKOnly, "NOTE:"
'End If
'On Error Resume Next
'Application.CommandBars("isoplot 3 worksheet tools").Visible = True
Exit Sub
NoCanDo: MsgBox "Unable to add all Isoplot menu-items", , Iso
End Sub

Sub Workbook_BeforeClose(Cancel As Boolean)
Attribute Workbook_BeforeClose.VB_ProcData.VB_Invoke_Func = " \n14"
' Delete all "Isoplot" dropdowns from the Worksheet & Chart menu-bars
Dim i%, b As Object, q As Object, c$, cs$, Bct%, Cbars As CommandBars
Dim ITb As CommandBar
GetOpSys
TT(1) = "isoplot": TT(2) = "fromiso": TT(3) = "startfromiso": TT(4) = "startfrominit"
Set IsoToolbars = Menus("isotoolbars")
If Not Starting And Not Cancel Then
  On Error Resume Next
  For i = 1 To 3
    With App.CommandBars(IsoToolbars(i, 3).Text)
      IsoToolbars(i, 2) = .Visible
      IsoToolbars(i, 5) = .Position
      IsoToolbars(i, 6) = .Left
      IsoToolbars(i, 7) = .Top
    End With
  Next i
  On Error GoTo 0
End If
With App
  On Error Resume Next
  If Not Starting Then
    Set Cbars = .CommandBars
    Cbars("Isoplot 3 Charts & Isochrons").Delete
    Cbars("Isoplot 3 Worksheet Tools & Isochrons").Delete
    Cbars("Isoplot 3 Chart Tools").Delete
  End If
  For i = 1 To 6
    For Each b In .CommandBars(i).Controls
      c$ = b.Caption: cs$ = Strip(c$, "&")
      If c$ = IsoCap Or Len(c$) = 0 Then
        b.Delete
      ElseIf cs$ = "Tools" Then ' <> "" And cs$ <> "File" And cs$ <> "Window" Then
        On Error Resume Next
        For Each q In b.Controls
          c$ = Strip(LCase(q.Caption), "&")
          If InStr(c$, Iso) Then q.Delete
        Next
        On Error GoTo 0
      End If
    Next
  Next i
  For Bct = 1 To NumB
    For i = 1 To 6 ' Standard, Charts, Formatting,...
      DelButton i, msoControlButton, Bct
    Next i
  Next Bct
End With
If Not Starting And ThisWorkbook.IsAddin Then StoreMenuSheet
End Sub

Sub DelInitialize(FromInitialize As Boolean)
Attribute DelInitialize.VB_ProcData.VB_Invoke_Func = " \n14"
Dim M As Object
FromInitialize = False
With Application.CommandBars(1)
  For Each M In .Controls
    If M.Caption = "Initialize Iso&plot" Then
1:    On Error GoTo 3
      M.Delete
2:    On Error GoTo 0
    End If
  Next M
End With
Exit Sub
3: On Error GoTo 0
FromInitialize = True
GoTo 2
End Sub

Sub DelButton(ByVal CbarNum%, ByVal bType%, Bct%)
Dim IndX%, i%, cb As Object, Bu As Object, d As Boolean, nCt%, ButNum%
Set cb = App.CommandBars(CbarNum): Set Bu = cb.Controls
Do
  FindIsoButton cb, bType, IndX, TT(Bct)
If IndX = 0 Then Exit Do
  With Bu(IndX)
    If .Type = bType Then
1:     On Error GoTo 3
       .Delete
2:      On Error GoTo 0
    End If
  End With
Loop
Exit Sub
3: GoTo 2
End Sub

Sub FindIsoButton(ComBar As Variant, ByVal bType%, IndX%, Bstr)
Attribute FindIsoButton.VB_ProcData.VB_Invoke_Func = " \n14"
' t$ is .OnAction if bType=msoControlButton, .Caption if msoControlPopup
Dim b As Object, c As Object, nc%, Ltt%, OnAct$, Capt$, rtOnAct$
IndX = 0
On Error GoTo 1
If IsObject(ComBar) Then
  Set c = ComBar
Else
  Set c = App.CommandBars(ComBar)
End If
nc = c.Controls.Count
For Each b In c.Controls
  Capt = "": OnAct = "": rtOnAct = ""
  With b
    Ltt = Len(Bstr)
    If .Type = bType Then
      If .Type = msoControlPopup Then
        Capt = LCase(.Caption)
      ElseIf .Type = msoControlButton Then
        OnAct = LCase(.OnAction): Capt = LCase(.Caption)
        rtOnAct = Right(OnAct, Ltt)
      Else
        MsgBox "FindIsoButton bType error": KwikEnd
      End If
      If Capt = Bstr Or rtOnAct = Bstr Then
        IndX = .Index
        Exit Sub
      End If
    End If
  End With
Next b
1: On Error GoTo 0
End Sub

Sub StoreMenuSheet() ' Do when unloading isoplot.xla or changing Consts
Dim e, s$, ct%, Ra As Range, rMax&, cMax%, r&, c%, ww As Workbook
NoUp
NoAlerts
NoUp
On Error GoTo 3
Workbooks.Add
Set ww = Awb
Ash.Name = "Isoplot"
If Mac And False Then
  StatBar "wait ..."
  CornerRange MenuSht, rMax, cMax
  ww.Activate
  For r = 1 To rMax
    For c = 1 To cMax
      Ash.Cells(r, c) = MenuSht.Cells(r, c).Formula
  Next c, r
  If Sheets.Count = 1 Then Sheets.Add Else Sheets(2).Activate
  sR(1, 1, 99, 99).Interior.Color = RGB(128, 128, 128)
  With Cells(2, 2)
    .Formula = "Isoplot"
    With .Font
      .Color = RGB(64, 64, 64)
      .Size = 108: .Italic = True
    End With
  End With
  Ash.Name = "Logo"
  StatBar
Else
  MenuSht.Cells.Copy
  With Sheets("isoplot")
    .Cells.PasteSpecial Paste:=xlValues
    App.CutCopyMode = False
  End With
  TW.Sheets("Logo").Copy Before:=Sheets("isoplot") ' which user sees briefly during copying
End If
On Error Resume Next
Workbooks(Istat$).Close
NoAlerts
On Error GoTo 3
s = IsoPath & PathSep & Istat
1: ActiveWorkbook.SaveAs Filename:=s, Password:="Slurg", ReadOnlyRecommended:=False, _
     FileFormat:=xlNormal
On Error GoTo 0
ActiveWindow.Close
Exit Sub

2: If ct = 1 Then GoTo 3
On Error GoTo 3
Kill s$
ct = 1 + ct
GoTo 1

3: e = Error(Err)
On Error Resume Next
ActiveWorkbook.Close
MsgBox "Unable to store Isoplot status-file -- current settings have not been saved." _
  & vbLf & vbLf & "(" & e & ")", , Iso
End Sub

Sub GetMenuSheet() ' Do when loading isoplot.xla
Dim Got As Boolean, s$, e, i%, En, rT$(2, 12), FD, ThisVer$, IsostatVer$, Ad$
Dim r&, c%, mr&, MC%, Nsh%, Ra As Range, Rd As Range
Const DYI = "DetritalYoungthIndx"

If Not (Int(ExcelVersion) = "10" And Windows) Then NoUp
'If Workbooks.Count = 0 Then Workbooks.Add
Got = False
NoAlerts
NoUp
s$ = IsoPath & PathSep & Istat
Ad = Menus("vernum").Address
ThisVer = Menus("vernum").Text
For r = 1 To 2
  For c = 1 To 12
    rT(r, c) = MenuSht.Cells(r, c).Text
Next c, r
' The Isostat.xls file should be in the same folder as the add-in itself
If Not (Int(ExcelVersion) = "10" And Windows) Then NoUp
On Error GoTo badd
FD = FileDateTime(s)
Workbooks.Open Filename:=s$, ReadOnly:=False, Password:="Slurg"
If Workbooks.Count > 0 Then
  If LCase(Awb.Name) = Istat Then Got = True
End If
With Awb.Sheets
  Nsh = .Count
  For i = 1 To Nsh
    s = LCase(.Item(i).Name)
    If s = "isoplot" Or s = "menus" Then Exit For
  Next i
  If i > Nsh Then GoTo badd
  'Item(i).Cells.Copy
End With
Awb.Sheets(s).Select
IsostatVer = Range(Ad).Text
If (Val(ThisVer) = Val(IsostatVer)) Then
  If Mac Then
    StatBar "wait..."
    CornerRange Ash, mr, MC
    For r = 1 To mr
      For c = 1 To MC
        MenuSht.Cells(r, c) = Ash.Cells(r, c).Formula
    Next c, r
  Else
    For r = 1 To 2
      For c = 1 To 10
        Menus(Cells(r, c).Address).Formula = rT(r, c)
    Next c, r
  End If
  Set Ra = Range(Menus("isotoolbars").Address)
  If Ra(1, 1) = "" Then ' in case version mismatch
    For i = 1 To 3
      Ra(i, 1) = Choose(i, "WorksheetTools", "ChartsIsochrons", "Charttools")
      Ra(i, 2) = False
      Ra(i, 3) = "Isoplot 3 " & Choose(i, "Worksheet Tools", "Charts & Isochrons", "Chart Tools")
    Next i
  End If
  App.CutCopyMode = False
  Ash.Cells.Copy
  MenuSht.Cells.PasteSpecial Paste:=xlValues
End If
MenuSht.Activate
On Error GoTo 1
Set Rd = Range(Menus(DYI).Address)
If Rd(1, 1) = "" Then
1:  On Error GoTo 0
  With Menus("charttools")
    r = 2 + .Row
    c = .Column
  End With
  Do
    r = r + 1
  Loop Until MenuSht.Cells(r, c) = "" And MenuSht.Cells(r + 1, c) = ""
  MenuSht.Cells(r, c).Name = DYI
  MenuSht.Cells(r, c) = 1
  MenuSht.Cells(r - 1, c) = DYI
End If

NoAlerts
On Error Resume Next
If Got Then
  Workbooks(Istat).Close
  NoUp False
End If
StatBar
Exit Sub
badd: e = Error(Err): En = Err
On Error GoTo 0
NoUp False
NoAlerts False
'MsgBox "Can't " & IIf(en = 1004, "find old IsoStat ", "load Isoplot's status-") & _
'  "file -- using default settings and constants." & IIf(en = 1004, "", vbLf & vbLf & e)
End Sub

Sub EmptyClipboard()
Dim TempDat As New DataObject
On Error Resume Next
With TempDat
  .SetText Text:=Empty
  .PutInClipboard
End With
On Error GoTo 0
End Sub
