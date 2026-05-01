VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Object = "{621DDB00-A516-11E8-A658-0013D350667C}#3.2#0"; "tx4ole26.ocx"
Begin VB.Form frmAufga 
   Caption         =   "Aufgaben"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
      _Version        =   1048579
      _ExtentX        =   6588
      _ExtentY        =   3625
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin Tx4oleLib.TXTextControl TexCont4 
      Height          =   1815
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
      _Version        =   196610
      _ExtentX        =   5953
      _ExtentY        =   3201
      _StockProps     =   73
      BackColor       =   16777215
      Language        =   49
      BorderStyle     =   1
      BackStyle       =   1
      ControlChars    =   0   'False
      EditMode        =   0
      HideSelection   =   -1  'True
      InsertionMode   =   -1  'True
      MousePointer    =   0
      ZoomFactor      =   100
      ViewMode        =   3
      ClipChildren    =   0   'False
      ClipSiblings    =   -1  'True
      SizeMode        =   0
      TabKey          =   -1  'True
      FormatSelection =   0   'False
      VTSpellDictionary=   "C:\PROGRA~1\TEXTCO~1\TXTEXT~1.0AC\Bin\AMERICAN.VTD"
      ScrollBars      =   3
      PageWidth       =   12240
      PageHeight      =   15840
      PageMarginL     =   1440
      PageMarginT     =   1440
      PageMarginR     =   1440
      PageMarginB     =   1440
      PrintZoom       =   100
      PrintOffset     =   0   'False
      PrintColors     =   -1  'True
      FontName        =   "Arial"
      FontSize        =   12
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Baseline        =   0
      TextBkColor     =   16777215
      Alignment       =   0
      LineSpacing     =   100
      LineSpacingT    =   0
      FrameStyle      =   32
      FrameDistance   =   0
      FrameLineWidth  =   20
      IndentL         =   0
      IndentR         =   0
      IndentFL        =   0
      IndentT         =   0
      IndentB         =   0
      Text            =   ""
      WordWrapMode    =   1
      AllowUndo       =   -1  'True
      TextFrameMarkerLines=   -1  'True
      FieldLinkTargetMarkers=   0   'False
      PageOrientation =   0
      PageViewStyle   =   1
      FontSettings    =   0
      AllowDrag       =   0   'False
      AllowDrop       =   0   'False
      SelectionViewMode=   1
      SectionRestartPageNumbering=   0
      PermanentControlChars=   16
      RightToLeft     =   0   'False
      TextDirection   =   2
      Locale          =   1031
      Justification   =   1
      FrameColor      =   16777215
      FrameLineColor  =   0
      DocumentPermissions=   31
      SelectObjects   =   -1  'True
      IsTrackChangesEnabled=   0   'False
      IsFormulaCalculationEnabled=   -1  'True
      FormulaReferenceStyle=   0
   End
   Begin VB.TextBox txtDummy 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   20000
      Width           =   80
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   975
      Left            =   5160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
      _Version        =   1048579
      _ExtentX        =   1931
      _ExtentY        =   1720
      _StockProps     =   64
      Show3DBorder    =   2
      VisualTheme     =   0
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   480
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAufga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private FS As Form
Private AktCo As VB.Control
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private ChRow As XtremeReportControl.ReportRow
Private ChRws As XtremeReportControl.ReportRows
Private TxCoN As Tx4oleLib.TXTextControl
Private WithEvents TbBar As XtremeCommandBars.TabToolBar
Attribute TbBar.VB_VarHelpID = -1
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Const GWL_WNDPROC = (-4)
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const KEYEVENTF_KEYUP = &H2
Private Const CB_FINDSTRING = &H14C&
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private TmTag As String
Private TxSav As Boolean
Public SelTa As Long

Private clFen As clsFenster
Private clFil As clsFile

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub FClos()
On Error GoTo LiErr

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlIdi = False Then 'Idiotenmodus
    If GlRes = False Then 'Reset der Einstellungen
        clFen.FenSav
        If clFen.FeSta = 0 Then
            IniSetVal "Aufgaben", "FenLin", clFen.FeLin
            IniSetVal "Aufgaben", "FenObe", clFen.FeObn
            IniSetVal "Aufgaben", "FenBre", clFen.FeBre
            IniSetVal "Aufgaben", "FenHoh", clFen.FeHoh
        End If
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim DaPi1 As XtremeCalendarControl.DatePicker
Dim CmDat As XtremeCommandBars.CommandBarEdit

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set DaPi1 = FM.dtpDatu1
Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)

If IsDate(CmDat.Text) Then
    NeuDa = CmDat.Text
Else
    NeuDa = Date
End If

With DaPi1
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Left = 3280
    .Top = 1260
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            CmDat.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set DaPi1 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FLad(ByVal NeuDa As Boolean)
On Error GoTo OpErr
'Öffnet das Wiedervorlageformular

Dim AnzPo As Long
Dim IdxNr As Long
Dim StMan As Long
Dim Mld1, Tit1 As String
Dim CmMan As XtremeSuiteControls.ComboBox
Dim CmMit As XtremeSuiteControls.ComboBox
Dim PuBu1 As XtremeSuiteControls.PushButton
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmAufga
Set FS = frmWieder
Set CmMan = FS.cmbBehan
Set CmMit = FS.cmbMitar
Set PuBu1 = FS.btnDatu1
Set RpCon = FM.repCont1
Set ImMan = frmMain.imgManag
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

StMan = GlMan(GlSMa, 2)

If GlWaT = RibTab_Wart_Wied Then
    GlKoL = True
    GlKoS = False
    If NeuDa = True Then
        GlKoN = True
        PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
        CmMan.ListIndex = GlSMa - 1
        CmMit.ListIndex = GlSmI - 1
        CmMan.Tag = "1IDP"
        CmMit.Tag = "1IDM"
        frmWieder.Show vbModal
    Else
        GlKoN = False
        AnzPo = RpCon.Records.Count
        If AnzPo > 0 Then
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                Set RpCol = RpCls.Find(Ter_ID2)
                IdxNr = RpRow.Record(RpCol.ItemIndex).Value
                Wie_Lad IdxNr
                PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
                frmWieder.Show vbModal
                GlKoS = False
            End If
        Else
            Mld1 = "Es ist kein Wiedervorlageeintrag vorhanden den Sie öffnen könnten"
            Tit1 = "Kein Wiedervorlageeintrag"
            SPopu Tit1, Mld1, IC48_Forbidden
        End If
    End If
End If

GlKoL = False

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLad " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim DaPi1 As XtremeCalendarControl.DatePicker
Dim CmDat As XtremeCommandBars.CommandBarEdit

Set CmBrs = Me.comBar02
Set DaPi1 = Me.dtpDatu1

Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)

If DaPi1.Selection.BlocksCount > 0 Then
    NeuDa = DaPi1.Selection.Blocks(0).DateBegin
    CmDat.Text = NeuDa
End If

Set DaPi1 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FHilfe()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = ""
TeMai = ""
TeInh = ""
TeFus = ""

'SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub

Private Sub FLoe(ByVal LiTyp As Integer)
On Error GoTo GrErr
'Löscht die markierten Einträge

Dim RowNr As Long
Dim IdxNr As Long
Dim ReNum As Long
Dim GesZa As Long
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Tit1 = "Eintrag Entfernen"

Set FM = frmAufga
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

If GlRch(0, 23) = 0 Then
    WindowMess "Sie besitzen keine Berechtigung für diesen Vorgang", Dial3, Tit1, FM.hwnd
    Exit Sub
End If

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

GesZa = RpSel.Count

If GesZa = 0 Then
    Exit Sub
ElseIf GesZa = 1 Then
    Mld1 = "Möchten Sie den markierten Eintrag wirklich entfernen?"
ElseIf GesZa > 1 Then
    Mld1 = "Möchten Sie die " & GesZa & " markierten Einträge wirklich entfernen?"
End If

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then
    Screen.MousePointer = vbHourglass
    clFen.FenDsk 2
    
    S_WaLo LiTyp
    
    clFen.FenDsk 3
    Screen.MousePointer = vbNormal
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Set clFen = Nothing

Exit Sub

GrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoe " & Err.Number
Resume Next

End Sub
Private Sub FMenu()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim RetWe As Long
Dim KeyNa As String
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim ToTab As XtremeCommandBars.TabControlItem
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmTex As XtremeCommandBars.CommandBarEdit
Dim CmDat As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmMan As XtremeCommandBars.CommandBarComboBox
Dim CmMit As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmAcs = CmBrs.Actions
Set CmOpt = CmBrs.Options
Set ImMan = frmMain.imgManag

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmSta
    .Font.SIZE = 8
    Set CmPan = .AddPane(1)
    CmPan.Text = Date
    CmPan.Width = 100
    CmPan.Alignment = xtpAlignmentCenter
    Set CmPan = .AddPane(2)
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    .Visible = True
End With

With CmAcs
    Set CmAct = .Add(SY_SuWi1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuWi2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuWi3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuDat, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_SuBut, vbNullString, vbNullString, vbNullString, vbNullString)
End With

Set TbBar = CmBrs.AddTabToolBar("TabBar")

'______________________________________________________________________

Set ToTab = TbBar.InsertCategory(RibTab_Wart_Wied, "Wiedervorlage")
With ToTab
    .ToolTip = "Zeigt die Widervorlagen des ausgewählten Tages"
    .Visible = True
    If SelTa > 0 Then
        If SelTa = RibTab_Wart_Wied Then
            .Selected = True
        Else
            .Selected = False
        End If
    Else
        .Selected = True
    End If
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .Category = "Wiedervorlage"
        .ToolTipText = "Legt eine neue Wiedervorlage an"
        .BeginGroup = True
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Oeffnen, "Bearbeiten")
    With CmCon
        .Category = "Wiedervorlage"
        .ToolTipText = "Öffnet die markierte Wiedervorlage"
        .BeginGroup = True
        .IconId = IC24_Doc_Edit
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Wiedervorlage"
        .ToolTipText = "Löscht die markierten Wiedervorlage"
        .BeginGroup = True
        .IconId = IC24_Doc_Del
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Drucken, "Drucken")
    With CmCon
        .Category = "Wiedervorlage"
        .ToolTipText = "Druckt die Aufgabenliste"
        .ShortcutText = "F10"
        .BeginGroup = True
        .IconId = IC24_Printer_Ink
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Wiedervorlage"
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
    End With
End With

'______________________________________________________________________

Set ToTab = TbBar.InsertCategory(RibTab_Wart_Beha, "In Behandlung")
With ToTab
    .ToolTip = "Zeigt alle Patienten, desses Rechnungen noch nicht abgeschlossen wurden"
    .Visible = True
    If SelTa > 0 Then
        If SelTa = RibTab_Wart_Beha Then
            .Selected = True
        Else
            .Selected = False
        End If
    Else
        .Selected = False
    End If
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Bearbeiten, "Bearbeiten")
    With CmCon
        .Category = "In Behandlung"
        .ToolTipText = "Rechnungsdaten Bearbeiten"
        .BeginGroup = True
        .IconId = IC24_Mail_Edit
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Delete, "Entfernen")
    With CmCon
        .Category = "In Behandlung"
        .ToolTipText = "Löscht die markierte Rechnung"
        .BeginGroup = True
        .IconId = IC24_Mail_Del
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Close, "Verriegeln")
    With CmCon
        .Category = "In Behandlung"
        .ToolTipText = "Verriegelt die markierten Rechnungen"
        .BeginGroup = True
        .IconId = IC24_Mail_Close
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "In Behandlung"
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
    End With
End With

'______________________________________________________________________

Set ToTab = TbBar.InsertCategory(RibTab_Wart_Noti, "Notizblock")
With ToTab
    .ToolTip = "Zeigt den allgemeinen Notizblock des Mitarbeiters"
    .Visible = True
    If SelTa > 0 Then
        If SelTa = RibTab_Wart_Noti Then
            .Selected = True
        Else
            .Selected = False
        End If
    Else
        .Selected = False
    End If
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Notizblock"
        .ToolTipText = "Notiz Speichern"
        .BeginGroup = True
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Schrift, "Schriftart")
    With CmCon
        .Category = "Notizblock"
        .ToolTipText = "Schriftart wählen"
        .BeginGroup = True
        .IconId = IC24_Form_Edit
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Drucken, "Drucken")
    With CmCon
        .Category = "Notizblock"
        .ToolTipText = "Notiz Drucken"
        .BeginGroup = True
        .ShortcutText = "F10"
        .IconId = IC24_Printer_Ink
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Notizblock"
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
    End With
End With

'______________________________________________________________________

For Each CmBar In CmBrs
    If CmBar.Type = xtpBarTypeNormal Then
        Set CmCoS = CmBar.Controls
        For Each CmCon In CmCoS
            CmCon.Style = xtpButtonIconAndCaption
        Next CmCon
    End If
Next CmBar

'______________________________________________________________________

Set CmBar = CmBrs.Add("ID_Suche", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, SY_Cap01, "Mitarbeiter :")
    With CmCon
        .ToolTipText = "Wählen Sie den gewünschten Mitarbeiter"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
        .Visible = False
    End With
    
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Suche in :")
    With CmCon
        .ToolTipText = "Geben Sie bitte hier Ihre Suchanfrage ein"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With

    Set CmMit = .Add(xtpControlComboBox, SY_SuMit, vbNullString)
    With CmMit
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Wählen Sie den gewünschten Mitarbeiter"
        .ThemedItems = True
        .Visible = False
        .Width = 140
        For AktZa = 1 To UBound(GlMiA)
            .AddItem GlMiA(AktZa, 1)
            .ItemData(AktZa) = GlMiA(AktZa, 2)
        Next AktZa
        .ListIndex = GlSmI
    End With

    Set CmCom = .Add(xtpControlComboBox, SY_SuWi1, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Welches Datenfeld soll durchsucht werden?"
        .IconId = IC16_View
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 100
        .AddItem "Tagesdatum", 1
        .AddItem "Patientenname", 2
        .AddItem "Terminbetreff", 3
        .AddItem "Mandanten", 4
        .AddItem "Ausstehend", 5
        .ListIndex = 1
    End With

    Set CmCon = .Add(xtpControlLabel, SY_Plac2, Space$(1))
    CmCon.Visible = False
        
    Set CmCon = .Add(xtpControlLabel, SY_Cap04, " nach :")
    With CmCon
        .ToolTipText = "Tragen Sie hier Ihre Suchkriterien ein"
        .Style = xtpButtonCaption
    End With
    
    Set CmEdi = .Add(xtpControlEdit, SY_SuWi2, vbNullString)
    With CmEdi
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Suchbegriff..."
        .ToolTipText = "Geben Sie bitte hier den Suchbegriff ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonAutomatic
        .Width = 140
    End With

    Set CmCon = .Add(xtpControlButton, SY_SuZur, vbNullString)
    With CmCon
        .ToolTipText = "Vorherigen Suchen"
        .Style = xtpButtonIcon
        .IconId = IC16_Arrow_Left
        .Visible = False
    End With
    
    Set CmTex = .Add(xtpControlEdit, SY_SuTex, vbNullString)
    With CmTex
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Suchbegriff..."
        .ToolTipText = "Geben Sie bitte hier den Suchbegriff ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonAutomatic
        .Width = 140
        .Visible = False
    End With
    
    Set CmCon = .Add(xtpControlButton, SY_SuWei, vbNullString)
    With CmCon
        .ToolTipText = "Nächsten Suchen"
        .Style = xtpButtonIcon
        .IconId = IC16_Arrow_Right
        .Visible = False
    End With

    Set CmMan = .Add(xtpControlComboBox, SY_SuWi3, vbNullString)
    With CmMan
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Die Einträge welches Mandanten sollen angezeigt werden?"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 140
        .Visible = False
        For AktZa = 1 To UBound(GlThe)
            .AddItem GlThe(AktZa, 13)
            .ItemData(AktZa) = GlThe(AktZa, 0)
        Next AktZa
        .ListIndex = 1
    End With

    Set CmDat = .Add(xtpControlEdit, SY_SuDat, vbNullString)
    With CmDat
        .ToolTipText = "Wählen Sie das gewünschte Tagesdatum aus"
        .Style = xtpButtonAutomatic
        .Width = 90
        .EditStyle = xtpEditStyleCenter
    End With
    
    Set CmCon = .Add(xtpControlButton, SY_SuBut, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hierm, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Month
    End With
End With

'______________________________________________________________________

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7: .VisualTheme = xtpThemeOffice2013
    Case 8: .VisualTheme = xtpThemeOffice2013
    Case Else: .VisualTheme = xtpThemeRibbon
    End Select
    If GlSty = 8 Then 'Office 2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    ElseIf GlSty = 7 Then 'Office 2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Else
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End If
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = False
    .ToolTipContext.ShowOfficeBorder = True
    .ToolTipContext.ShowShadow = True
    .ToolTipContext.ShowTitleAndDescription False, xtpToolTipIconNone
    .ToolTipContext.Style = xtpToolTipResource
    .PaintManager.AutoResizeIcons = False
    .PaintManager.ClearTypeTextQuality = GlCle
    .PaintManager.EnableAnimation = GlMeA
    .PaintManager.FlatMenuBar = False
    .PaintManager.FlatToolBar = False
    .PaintManager.SelectImageInPopupBar = True
    .PaintManager.ShowShadow = True
    .PaintManager.ThemedCheckBox = True
    .PaintManager.ThemedStatusBar = True
    .PaintManager.ThickCheckMark = False
    .KeyBindings.Add 0, VK_F1, KY_F1
    .KeyBindings.Add 0, VK_F2, KY_F2
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
End With

With CmOpt
    .AltDragCustomization = False
    .AlwaysShowFullMenus = True
    .AutoHideUnusedPopups = False
    .ExpandDelay = 100
    .ExpandHoverDelay = 100
    .FloatToolbarsByDoubleClick = False
    .IconsWithShadow = False
    .KeyboardCuesShow = xtpKeyboardCuesShowAlways
    .KeyboardCuesUse = xtpKeyboardCuesUseMenuOnly
    .LargeIcons = False
    .LunaColors = GlLun
    .MaxPopupWidth = 0.5
    .OfficeStyleDisabledIcons = True
    .SetIconSize True, 24, 24
    .ShowExpandButtonAlways = False
    .ShowFullAfterDelay = True
    .ShowPopupBarToolTips = False
    .ShowTextBelowIcons = False
    .ShowKeyboardTips = True
    .SyncFloatingToolbars = True
    .ToolBarAccelTips = True
    .ToolBarScreenTips = True
    .UpdatePeriod = 100
    .UseAltNumPadKeys = False
    .UseDisabledIcons = True
    .UseFadedIcons = False
    .UseSharedImageList = False
    .UseSystemSaveBitsStyle = False
    .Animation = xtpAnimateWindowsDefault
    .Font.SIZE = 8
    .ComboBoxFont.SIZE = 8
End With

With TbBar
    .AllowReorder = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableAnimation = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 24, 24
    .ShowExpandButton = False
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .SetIconSize 24, 24
    Select Case GlSty
    Case 8:
        .TabPaintManager.Appearance = xtpTabThemeOffice2013
        .TabPaintManager.Color = xtpTabColorOffice2013
    Case 7:
        .TabPaintManager.Appearance = xtpTabThemeOffice2013
        .TabPaintManager.Color = xtpTabColorOffice2013
    Case Else:
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2007
        .TabPaintManager.Color = xtpTabColorResource
    End Select
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ButtonMargin.Top = 6
    .TabPaintManager.FixedTabWidth = 110
    .TabPaintManager.ButtonMargin.Bottom = 0
    .TabPaintManager.ButtonMargin.Left = 0
    .TabPaintManager.ButtonMargin.Right = 0
    .TabPaintManager.ClientFrame = xtpTabFrameSingleLine
    .TabPaintManager.ClientMargin.Bottom = 0
    .TabPaintManager.ClientMargin.Top = 0
    .TabPaintManager.ClientMargin.Left = 0
    .TabPaintManager.ClientMargin.Right = 0
    .TabPaintManager.ControlMargin.Top = 0
    .TabPaintManager.ControlMargin.Bottom = 0
    .TabPaintManager.ControlMargin.Left = 0
    .TabPaintManager.ControlMargin.Right = 0
    .TabPaintManager.HeaderMargin.Top = 0
    .TabPaintManager.HeaderMargin.Bottom = 0
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.HeaderMargin.Right = 0
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = True
    .TabPaintManager.HotTracking = True
    .TabPaintManager.Layout = xtpTabLayoutFixed
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.Font.SIZE = 8
End With

DoEvents
CmBrs.RecalcLayout
DoEvents
CmBrs.PaintManager.RefreshMetrics

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FPrint()
On Error GoTo OpErr
'Setzt die Aknkunftszeit und die Abrisezeit

Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set RpCon = FM.repCont1
Set TxCoN = FM.TexCont4
Set CmAcs = CmBrs.Actions

If GlWaT = RibTab_Wart_Noti Then
    TxCoN.PrintDialog "Druck"
Else
    RpCon.PrintReport 0
End If

Set RpCon = Nothing
Set TxCoN = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSet " & Err.Number
Resume Next

End Sub
Private Sub FReFi()
On Error GoTo AnErr
'Filtert die Patientenrechnungen

Dim AnzPo As Long
Dim IdxNr As Long
Dim Mld1, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCon = Me.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

AnzPo = RpCon.Records.Count

If AnzPo > 0 Then
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Select Case GlWaT
            Case RibTab_Wart_Wied: Set RpCol = RpCls.Find(0)
            Case RibTab_Wart_Beha: Set RpCol = RpCls.Find(1)
            End Select
            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
            Select Case GlAdO 'Adressenverwaltung Doppelklick
            Case 0: SReZe IdxNr
            Case 1: SKrZe IdxNr
            Case 2: FLad False
            End Select
        End If
    End If
Else
    Mld1 = "Es ist kein Termin vorhanden den Sie öffnen könnten. Legen Sie einen neuen Termin an"
    Tit1 = "Kein Termin"
    SPopu Tit1, Mld1, IC48_Forbidden
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FReFi " & Err.Number
Resume Next

End Sub
Private Sub FSet(ByVal EiTyp As Integer)
On Error GoTo OpErr
'Setzt die Aknkunftszeit und die Abrisezeit

Dim AnzPo As Long
Dim IdxNr As Long
Dim Mld1, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmAufga
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

AnzPo = RpCon.Records.Count
If AnzPo > 0 Then
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        Set RpCol = RpCls.Find(Ter_ID2)
        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
        
        Screen.MousePointer = vbHourglass
        clFen.FenDsk 2
        
        S_WaSe1 IdxNr, EiTyp, GlWaT
        
        clFen.FenDsk 3
        Screen.MousePointer = vbNormal
    End If
Else
    Mld1 = "Es ist kein Termineintrag vorhanden den Sie bearbeiten könnten"
    Tit1 = "Kein Termineintrag"
    SPopu Tit1, Mld1, IC48_Forbidden
End If

Set clFen = Nothing

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSet " & Err.Number
Resume Next

End Sub
Private Sub FSuFe()
On Error GoTo OrErr
'Öffnet das Suchformular oder springt in das Suchfeld

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmDat As XtremeCommandBars.CommandBarEdit
Dim CmTex As XtremeCommandBars.CommandBarEdit
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmBu1 As XtremeCommandBars.CommandBarControl
Dim CmCa4 As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmMan As XtremeCommandBars.CommandBarComboBox
Dim CmMit As XtremeCommandBars.CommandBarComboBox

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, SY_SuWi1, , True)
Set CmEdi = CmBrs.FindControl(CmEdi, SY_SuWi2, , True)
Set CmTex = CmBrs.FindControl(CmTex, SY_SuTex, , True)
Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)
Set CmMan = CmBrs.FindControl(CmMan, SY_SuWi3, , True)
Set CmMit = CmBrs.FindControl(CmMit, SY_SuMit, , True)
Set CmBu1 = CmBrs.FindControl(CmBu1, SY_SuBut, , True)
Set CmCa4 = CmBrs.FindControl(CmCa4, SY_Cap04, , True)

Select Case CmCom.ListIndex
Case 1:
    CmEdi.Visible = False
    CmDat.Visible = True
    CmBu1.Visible = True
    CmMan.Visible = False
    CmCa4.Visible = True
Case 2:
    CmEdi.Visible = True
    CmDat.Visible = False
    CmBu1.Visible = False
    CmMan.Visible = False
    CmCa4.Visible = True
    CmAcs(SY_SuWi2).Visible = True
Case 3:
    CmEdi.Visible = True
    CmDat.Visible = False
    CmBu1.Visible = False
    CmMan.Visible = False
    CmCa4.Visible = True
    CmAcs(SY_SuWi2).Visible = True
Case 4:
    CmEdi.Visible = False
    CmDat.Visible = False
    CmBu1.Visible = False
    CmMan.Visible = True
    CmCa4.Visible = True
    CmAcs(SY_SuWi3).Visible = True
Case 5:
    CmEdi.Visible = False
    CmDat.Visible = False
    CmBu1.Visible = False
    CmMan.Visible = False
    CmCa4.Visible = False
End Select

CmBrs.RecalcLayout
DoEvents

S_WaLa GlWaT

Set CmAcs = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFe " & Err.Number
Resume Next


End Sub
Private Sub FTabu(ByVal TaIdx As Long)
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmTex As XtremeCommandBars.CommandBarEdit
Dim CmDat As XtremeCommandBars.CommandBarEdit
Dim CmCa1 As XtremeCommandBars.CommandBarControl
Dim CmCa2 As XtremeCommandBars.CommandBarControl
Dim CmCa4 As XtremeCommandBars.CommandBarControl
Dim CmPl2 As XtremeCommandBars.CommandBarControl
Dim CmBu1 As XtremeCommandBars.CommandBarControl
Dim CmBu2 As XtremeCommandBars.CommandBarControl
Dim CmBu3 As XtremeCommandBars.CommandBarControl
Dim CmAus As XtremeCommandBars.CommandBarComboBox
Dim CmMan As XtremeCommandBars.CommandBarComboBox
Dim CmMit As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmAufga
Set TxCoN = FM.TexCont4
Set RpCon = FM.repCont1
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmEdi = CmBrs.FindControl(CmEdi, SY_SuWi2, , True)
Set CmTex = CmBrs.FindControl(CmTex, SY_SuTex, , True)
Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)
Set CmCa1 = CmBrs.FindControl(CmCa1, SY_Cap01, , True)
Set CmCa2 = CmBrs.FindControl(CmCa2, SY_Cap02, , True)
Set CmCa4 = CmBrs.FindControl(CmCa4, SY_Cap04, , True)
Set CmPl2 = CmBrs.FindControl(CmPl2, SY_Plac2, , True)
Set CmBu1 = CmBrs.FindControl(CmBu1, SY_SuBut, , True)
Set CmBu2 = CmBrs.FindControl(CmBu2, SY_SuZur, , True)
Set CmBu3 = CmBrs.FindControl(CmBu3, SY_SuWei, , True)
Set CmAus = CmBrs.FindControl(CmAus, SY_SuWi1, , True)
Set CmMan = CmBrs.FindControl(CmMan, SY_SuWi3, , True)
Set CmMit = CmBrs.FindControl(CmMit, SY_SuMit, , True)

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

If GlWLa = False Then
    Screen.MousePointer = vbHourglass
    clFen.FenDsk 2

    Select Case TaIdx
    Case 0:
            GlWaT = RibTab_Wart_Wied
            WaSpl RibTab_Wart_Wied
            S_WaLa RibTab_Wart_Wied
            CmEdi.Visible = False
            CmTex.Visible = False
            CmDat.Visible = True
            CmCa1.Visible = False
            CmCa2.Visible = True
            CmCa4.Visible = True
            CmPl2.Visible = False
            CmBu1.Visible = True
            CmBu2.Visible = False
            CmBu3.Visible = False
            CmAus.Visible = True
            CmMan.Visible = False
            CmMit.Visible = False
            RpCon.Visible = True
            TxCoN.Visible = False
    Case 1:
            GlWaT = RibTab_Wart_Beha
            WaSpl RibTab_Wart_Beha
            S_WaLa RibTab_Wart_Beha
            CmEdi.Visible = False
            CmTex.Visible = False
            CmDat.Visible = False
            CmCa1.Visible = False
            CmCa2.Visible = False
            CmCa4.Visible = False
            CmPl2.Visible = True
            CmBu1.Visible = False
            CmBu2.Visible = False
            CmBu3.Visible = False
            CmAus.Visible = False
            CmMan.Visible = False
            CmMit.Visible = False
            RpCon.Visible = True
            TxCoN.Visible = False
    Case 2:
            GlWaT = RibTab_Wart_Noti
            CmEdi.Visible = False
            CmTex.Visible = True
            CmDat.Visible = False
            CmCa1.Visible = True
            CmCa2.Visible = False
            CmCa4.Visible = False
            CmPl2.Visible = True
            CmBu1.Visible = False
            CmBu2.Visible = True
            CmBu3.Visible = True
            CmAus.Visible = False
            CmMan.Visible = False
            CmMit.Visible = True
            RpCon.Visible = False
            TxCoN.Visible = True
            If TxCoN.Text = vbNullString Then
                FTxLa
            End If
    End Select
    
    clFen.FenDsk 3
    Screen.MousePointer = vbNormal
End If

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then
    Exit Sub
End If

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F2: FLad False
Case KY_F3: FLad True
Case KY_F8: FTxSa
Case KY_F10: FPrint
Case KY_F11: Unload Me
Case SY_OP_Reset: FSet 1
Case SY_OP_Zuruck: FSet 2
Case SY_OP_Hinzufuegen: FLad True
Case SY_OP_Oeffnen: FLad False
Case SY_OP_Loeschen: FLoe GlWaT
Case SY_OP_Abbruch: Unload Me
Case SY_OP_Drucken: FPrint
Case SY_OP_Bearbeiten: FReFi
                       frmReEdit.Show vbModal
Case SY_OP_Delete: FReFi
                  SLoHa
Case SY_OP_Close: FReFi
                  frmReAbs.Show vbModal
Case SY_OP_Speichern: FTxSa
Case SY_OP_Schrift: FTxCo TolId
Case SY_SuBut: FKale
Case SY_SuDat: S_WaLa GlWaT
Case SY_SuWi1: FSuFe
Case SY_SuWi2: S_WaLa GlWaT
Case SY_SuWi3: S_WaLa GlWaT
Case SY_SuMit: FTxLa
Case SY_SuTex: FTxCo TolId
Case SY_SuZur: FTxCo TolId
Case SY_SuWei: FTxCo TolId
End Select

GlToo = False

End Sub

Private Sub FTxCo(ByVal TxFun As Long, Optional ByVal ColID As Long)
On Error GoTo PoErr

Dim SuStr As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmTex As XtremeCommandBars.CommandBarEdit

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont4

Set CmTex = CmBrs.FindControl(CmTex, SY_SuTex, , True)

SuStr = CmTex.Text

Select Case TxFun
Case SY_SuMit:
Case SY_SuTex: TxCoN.Find SuStr, 1
Case SY_SuZur: TxCoN.Find SuStr, -1, 1
Case SY_SuWei: TxCoN.Find SuStr, -1
Case SY_OP_Schrift: TxCoN.FontDialog
End Select

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxCO " & Err.Number
Resume Next

End Sub
Private Sub FTxLa()
On Error GoTo PoErr

Dim MitNr As Long
Dim FiNam As String
Dim DaNam As String
Dim GuiID As String
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmMit As XtremeCommandBars.CommandBarComboBox

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont4

Set CmMit = CmBrs.FindControl(CmMit, SY_SuMit, , True)

MitNr = CmMit.ItemData(CmMit.ListIndex)

TxCoN.ResetContents

For AktZa = 1 To UBound(GlMiK)
    If MitNr = GlMiK(AktZa, 2) Then
        GuiID = GlMiK(AktZa, 20)
        DaNam = "TD_" & SNaFi(GuiID) & ".txm"
        FiNam = GlDox & DaNam
        Exit For
    End If
Next AktZa

Set clFil = New clsFile

With clFil
    If Not IsNull(FiNam) And Not FiNam = vbNullString Then
        If clFil.FilVor(FiNam) = True Then
            TxCoN.Load FiNam, 0, 3
        End If
    End If
End With

Set TxCoN = Nothing
Set clFil = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxLa " & Err.Number
Resume Next

End Sub
Private Sub FTxSa()
On Error GoTo PoErr

Dim MitNr As Long
Dim FiNam As String
Dim DaNam As String
Dim GuiID As String
Dim AktZa As Integer
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmMit As XtremeCommandBars.CommandBarComboBox

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont4

Set CmMit = CmBrs.FindControl(CmMit, SY_SuMit, , True)

MitNr = CmMit.ItemData(CmMit.ListIndex)

For AktZa = 1 To UBound(GlMiK)
    If MitNr = GlMiK(AktZa, 2) Then
        GuiID = GlMiK(AktZa, 20)
        DaNam = "TD_" & SNaFi(GuiID) & ".txm"
        FiNam = GlDox & DaNam
        Exit For
    End If
Next AktZa

Tit1 = "Notiz Speichern"
Mld1 = "Soll die Notiz gespeichert werden?"

Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
If Frage = 6 Then

    Set clFil = New clsFile
    
    With clFil
        If Not IsNull(FiNam) And Not FiNam = vbNullString Then
            If clFil.FilVor(FiNam) = True Then
                .DaLoe = FiNam & vbNullChar
                .FilLoe
            End If
        End If
    End With
    
    Set clFil = Nothing
    
    TxCoN.Save FiNam, 0, 3
End If

TxSav = False

Set TxCoN = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTxSa " & Err.Number
Resume Next

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlWLa = False Then FTool Control.id
End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlWLa = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    WaPosi
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)

Dim AktTa As Long
Dim AktZa As Integer

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTeV = True Then 'Termine vorhanden
    If GlTpV = True Then 'Kalendermarker vorhanden
        For AktTa = 0 To GlKMa - 1 'Anzahl Kalendermatker
            If Day = Left$(GlTEr(0, AktTa), 10) Then
                For AktZa = 1 To UBound(GlTep) 'Kalendermarker
                    If GlTep(AktZa, 0) = GlTEr(1, AktTa) Then
                        Metrics.BackColor = GlTep(AktZa, 2)
                        Exit For
                    End If
                Next AktZa
            End If
        Next AktTa
    End If
End If

End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
    S_WaLa GlWaT
End Sub

Private Sub Form_Activate()
    WaPosi
End Sub

Private Sub Form_Load()
    
Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 14000
    .ClientMaxWidth = 11000
    .ClientMinHeight = 6600
    .ClientMinWidth = 6800
    .TopMost = True
End With

FMenu
GlWaT = SelTa

Set FrmEx = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmAufga = Nothing
End Sub
Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Select Case GlWaT
Case RibTab_Wart_Wied:
    If CDate(Row.Record(2).Value) < Date Then
        Metrics.ForeColor = vbRed
    End If
Case RibTab_Wart_Beha:
    If Row.GroupRow = False Then
        If Row.Record(Rec_Selekt).Value = 0 Then Metrics.Font.Bold = True
        Select Case Row.Record(Rec_Type).Value
        Case "M": Metrics.ForeColor = 16744448
        Case "L": Metrics.ForeColor = 33023
        Case "V": Metrics.ForeColor = 8421631
        Case "U": Metrics.ForeColor = 6604830
        Case Else:
            If Row.Record(Rec_Selekt).Value = 0 Then
                Metrics.ForeColor = GlRFa 'Farbe nicht abgeschlossene Rechnungen
            End If
            If Row.Record(Rec_Storniert).Value = True Then
                Metrics.Font.Strikethrough = True
                Metrics.ForeColor = 8421504
            End If
        End Select
    End If
End Select

End Sub

Private Sub repCont1_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
    Item.Tag = "@" & TmTag
    S_WaSa
End Sub

Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        FReFi
        If GlAdO < 2 Then 'Adressenverwaltung Doppelklick
            If GlWaL = True Then 'Wartezimmerliste Schließen
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FReFi
    If GlAdO < 2 Then 'Adressenverwaltung Doppelklick
        If GlWaL = True Then 'Wartezimmerliste Schließen
            Unload Me
        End If
    End If
End Sub
Private Sub repCont1_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)
    Item.Tag = "@" & TmTag
    S_WaSa
End Sub
Private Sub TbBar_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    FTabu Item.Index
End Sub

Private Sub TexCont4_PosChange()
    TxSav = True
End Sub
Private Sub txtDummy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
