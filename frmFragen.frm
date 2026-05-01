VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmFragen 
   Caption         =   "Fragebogen"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   320
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   564
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   2000
      Left            =   195
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   10995
      _Version        =   1048579
      _ExtentX        =   19394
      _ExtentY        =   3528
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   1575
         Left            =   200
         TabIndex        =   2
         Top             =   200
         Width           =   5505
         _Version        =   1048579
         _ExtentX        =   9701
         _ExtentY        =   2787
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   5200
      Left            =   105
      TabIndex        =   3
      Top             =   495
      Width           =   10995
      _Version        =   1048579
      _ExtentX        =   19394
      _ExtentY        =   9172
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtBeTex 
         Height          =   780
         Left            =   1400
         TabIndex        =   4
         Top             =   1360
         Width           =   9000
         _Version        =   1048579
         _ExtentX        =   15875
         _ExtentY        =   1376
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkVorga 
         Height          =   240
         Left            =   8200
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4500
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Checkboxfeldvorgabe"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbGrupe 
         Height          =   315
         Left            =   1400
         TabIndex        =   6
         Top             =   3380
         Width           =   4200
         _Version        =   1048579
         _ExtentX        =   7408
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtZeile 
         Height          =   350
         Left            =   1400
         TabIndex        =   7
         Top             =   3920
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "2"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtZeich 
         Height          =   350
         Left            =   1400
         TabIndex        =   8
         Top             =   4490
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1773
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "20"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   350
         Left            =   2420
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3920
         Width           =   250
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Max             =   8
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtZeile"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtSorte 
         Height          =   350
         Left            =   4280
         TabIndex        =   10
         Top             =   3920
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1773
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "1"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtFelNa 
         Height          =   350
         Left            =   6900
         TabIndex        =   11
         Top             =   3380
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6165
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         MaxLength       =   100
      End
      Begin XtremeSuiteControls.FlatEdit txtMaxZe 
         Height          =   350
         Left            =   4280
         TabIndex        =   12
         Top             =   4490
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1773
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "250"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkPflch 
         Height          =   240
         Left            =   6900
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4500
         Width           =   2205
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Pflichtwert"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   350
         Left            =   2420
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4490
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Max             =   40
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtZeich"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont3 
         Height          =   350
         Left            =   5300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3920
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Max             =   9999
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSorte"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont4 
         Height          =   350
         Left            =   5300
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4490
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Min             =   1
         Max             =   250
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtMaxZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtBezei 
         Height          =   780
         Left            =   1400
         TabIndex        =   17
         Top             =   400
         Width           =   9000
         _Version        =   1048579
         _ExtentX        =   15875
         _ExtentY        =   1376
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtVorga 
         Height          =   350
         Left            =   1400
         TabIndex        =   18
         Top             =   2320
         Width           =   9000
         _Version        =   1048579
         _ExtentX        =   15875
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         MaxLength       =   100
      End
      Begin XtremeSuiteControls.ComboBox cmbAbhen 
         Height          =   315
         Left            =   1400
         TabIndex        =   19
         Top             =   2840
         Width           =   9000
         _Version        =   1048579
         _ExtentX        =   15875
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtIdxNr 
         Height          =   350
         Left            =   6900
         TabIndex        =   20
         Top             =   3920
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin VB.Label lblLabl8 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Zeichenbreite :"
         Height          =   255
         Left            =   220
         TabIndex        =   31
         Top             =   4540
         Width           =   1100
      End
      Begin VB.Label lblLabl5 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Anz. Zeilen :"
         Height          =   240
         Left            =   220
         TabIndex        =   30
         Top             =   3940
         Width           =   1100
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Fragentext :"
         Height          =   240
         Left            =   220
         TabIndex        =   29
         Top             =   450
         Width           =   1100
      End
      Begin VB.Label lblLabl1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Vorgabetext :"
         Height          =   240
         Left            =   220
         TabIndex        =   28
         Top             =   2380
         Width           =   1100
      End
      Begin XtremeSuiteControls.Label lblLabl7 
         Height          =   240
         Left            =   5730
         TabIndex        =   27
         Top             =   3420
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Feldname :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Sortierung :"
         Height          =   240
         Left            =   3120
         TabIndex        =   26
         Top             =   3940
         Width           =   1100
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Max. Zeichen :"
         Height          =   255
         Left            =   3120
         TabIndex        =   25
         Top             =   4540
         Width           =   1100
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   240
         Left            =   220
         TabIndex        =   24
         Top             =   3420
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Gruppe :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   240
         Left            =   220
         TabIndex        =   23
         Top             =   1420
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Antworttext :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLabl3 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Abhängig von :"
         Height          =   240
         Left            =   220
         TabIndex        =   22
         Top             =   2860
         Width           =   1100
      End
      Begin XtremeSuiteControls.Label lblLabl9 
         Height          =   240
         Left            =   5730
         TabIndex        =   21
         Top             =   3940
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Fragen-Nr.:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   480
      Top             =   0
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
Attribute VB_Name = "frmFragen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private ColMa As XtremeCommandBars.ColorManager
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRec As XtremeReportControl.ReportRecord
Private RpRes As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private RpCol As XtremeReportControl.ReportColumn
Private RpCls As XtremeReportControl.ReportColumns
Private RpRow As XtremeReportControl.ReportRow
Private RpGrw As XtremeReportControl.ReportGroupRow
Private ChRow As XtremeReportControl.ReportRow
Private ChRws As XtremeReportControl.ReportRows
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
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

Private clFen As clsFenster

Private FoLad As Boolean

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
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
Private Sub cmbAbhen_Click()
    If FoLad = False Then
        FPruf
    End If
End Sub
Private Sub FaInit()
On Error GoTo InErr

Dim RetWe As Long
Dim KeyNa As String
Dim TreKy As String
Dim AktZa As Integer
Dim TmFnt As New StdFont
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim ToTab As XtremeCommandBars.TabControlItem
Dim ChVor As XtremeSuiteControls.CheckBox
Dim ChPfl As XtremeSuiteControls.CheckBox

Set FM = frmFragen
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set ChVor = FM.chkVorga
Set ChPfl = FM.chkPflch
Set CmBrs = FM.comBar02
Set RpCon = FM.repCont1
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag

KeyNa = "ToolTips"
TreKy = Left$(GlNod, 1)

TmFnt.Name = GlTFt.Name
TmFnt.SIZE = GlTFt.SIZE

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmSta
    .Font.SIZE = 8
    Set CmPan = .AddPane(1)
    CmPan.Text = vbNullString
    CmPan.Width = 100
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Style = SBPS_STRETCH
    .Visible = True
End With

Set TbBar = CmBrs.AddTabToolBar("TabBar")

Set ToTab = TbBar.InsertCategory(RibTab_Opti1, "Hauptdaten")
With ToTab
    .ToolTip = "Die Hauptdaten der gewählten Frage"
    .Selected = True
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neue Frage")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Legt eine Frage an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Speichert die aktuelle Frage"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Löscht die aktuelle Frage"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Hauptdaten"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Hauptdaten"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

Set ToTab = TbBar.InsertCategory(RibTab_Opti2, "Antwortvorgaben")
With ToTab
    .ToolTip = "Antwortvorgaben für die Frage"
    .Selected = False
End With
Set CmCoS = TbBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Sub_Neu, "Neue Antwort")
    With CmCon
        .Category = "Antwortvorgaben"
        .ToolTipText = "Legt eine neue Frage an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Sub_Sav, "Speichern")
    With CmCon
        .Category = "Antwortvorgaben"
        .ToolTipText = "Speichert alle Antwortvorgaben"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Sub_Loe, "Entfernen")
    With CmCon
        .Category = "Antwortvorgaben"
        .ToolTipText = "Löscht die markierte Antwort"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .Category = "Antwortvorgaben"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .Category = "Antwortvorgaben"
        .ToolTipText = "Schließt diesen Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

For Each CmBar In CmBrs
    If CmBar.Type = xtpBarTypeNormal Then
        Set CmCoS = CmBar.Controls
        For Each CmCon In CmCoS
            CmCon.Style = xtpButtonIconAndCaption
        Next CmCon
    End If
Next CmBar

'-----------------------------------------------------------------------------------------------------------

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
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Fragentyp :")
    With CmCon
        .ToolTipText = "Wählen Sie den passenden Fragentyp aus"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With
    Set CmCom = .Add(xtpControlComboBox, SY_SuCm1, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Wählen Sie den passenden Fragentyp aus"
        .IconId = IC16_View
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 120
        For AktZa = 1 To UBound(GlFrT)
            .AddItem GlFrT(AktZa)
            .ItemData(AktZa) = AktZa
        Next AktZa
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Nav_Vor, "Nach Oben")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .ToolTipText = "Verschiebt den Eintrag nach oben"
        .IconId = IC16_Arrow_Up
        .BeginGroup = True
        .Visible = False
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Nav_Zuru, "Nach Unten")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .ToolTipText = "Verschiebt den Eintrag nach unten"
        .IconId = IC16_Arrow_Down
        .BeginGroup = True
        .Visible = False
    End With
End With

'-----------------------------------------------------------------------------------------------------------

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = True
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = True
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Einträge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = False
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = GlAnZ
    .PaintManager.ThemedInplaceButtons = True
    .PaintManager.HorizontalGridStyle = xtpGridSolid
    .PaintManager.VerticalGridStyle = xtpGridSolid
    .PaintManager.FixedRowHeight = True
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With
'---

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

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
ChVor.BackColor = GlBak
ChPfl.BackColor = GlBak

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FaInit " & Err.Number
Resume Next

End Sub
Private Sub FClos()
On Error GoTo LiErr

Dim TypNr As Integer
Dim NoClo As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim CmTyp As XtremeCommandBars.CommandBarComboBox

Set FM = frmFragen
Set CmBrs = FM.comBar02
Set RpCo1 = FM.repCont1
Set RpRcs = RpCo1.Records

Set CmTyp = CmBrs.FindControl(CmTyp, SY_SuCm1, , True)

TypNr = CmTyp.ListIndex

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Select Case TypNr:
Case 1: NoClo = False 'Textfeld
Case 2: NoClo = True  'Auswahlfeld
Case 3: NoClo = True  'Ankreuzfeld
Case 4: NoClo = True  'Einfachauswahl
Case 5: NoClo = True  'Mehrfachauswahl
Case 6: NoClo = False 'Zwischentext
Case 7: NoClo = False 'Datumsfeld
End Select

If NoClo = True Then
    If RpRcs.Count = 0 Then
        F_Neu
        DoEvents
    End If
End If

If GlIdi = False Then 'Idiotenmodus
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "KatFrag", "FenLin", clFen.FeLin
        IniSetVal "KatFrag", "FenObe", clFen.FeObn
        IniSetVal "KatFrag", "FenBre", clFen.FeBre
        IniSetVal "KatFrag", "FenHoh", clFen.FeHoh
    End If
End If

Set RpRcs = Nothing
Set RpCo1 = Nothing
Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
Resume Next

End Sub
Private Sub FPruf()
On Error Resume Next

Dim RetWe As Long
Dim IdxNr As Long
Dim SubNr As Long
Dim TxIdx As XtremeSuiteControls.FlatEdit
Dim CmAbh As XtremeSuiteControls.ComboBox

Set FM = frmFragen
Set TxIdx = FM.txtIdxNr
Set CmAbh = FM.cmbAbhen

If TxIdx.Text <> vbNullString Then
    If IsNumeric(TxIdx.Text) = True Then
        If CLng(TxIdx.Text) > 0 Then
            IdxNr = CLng(TxIdx.Text)
        Else
            IdxNr = 0
        End If
    Else
        IdxNr = 0
    End If
Else
    IdxNr = 0
End If

If CmAbh.ListCount > 1 Then
    SubNr = CmAbh.ItemData(CmAbh.ListIndex)
Else
    SubNr = 0
End If

If IdxNr > 0 Then
    If IdxNr = SubNr Then
        RetWe = SendMessage(CmAbh.hwnd, CB_SETCURSEL, 0, ByVal 0&)
    End If
End If

End Sub

Private Sub FTabu(ByVal TaIdx As Long)
On Error GoTo AnErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmLab As XtremeCommandBars.CommandBarControl
Dim CmBu1 As XtremeCommandBars.CommandBarControl
Dim CmBu2 As XtremeCommandBars.CommandBarControl
Dim CmBu3 As XtremeCommandBars.CommandBarControl
Dim CmBu4 As XtremeCommandBars.CommandBarControl
Dim CmTyp As XtremeCommandBars.CommandBarComboBox

Set FM = frmFragen
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmLab = CmBrs.FindControl(CmLab, SY_Cap02, , True)
Set CmTyp = CmBrs.FindControl(CmTyp, SY_SuCm1, , True)
Set CmBu1 = CmBrs.FindControl(CmBu1, SY_OP_Sub_Neu, , True)
Set CmBu2 = CmBrs.FindControl(CmBu2, SY_OP_Sub_Loe, , True)
Set CmBu3 = CmBrs.FindControl(CmBu3, SY_OP_Nav_Vor, , True)
Set CmBu4 = CmBrs.FindControl(CmBu4, SY_OP_Nav_Zuru, , True)

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case TaIdx
Case 0:
    Rahm1.Visible = True
    Rahm2.Visible = False
    CmLab.Visible = True
    CmTyp.Visible = True
    CmBu1.Visible = False
    CmBu2.Visible = False
    CmBu3.Visible = False
    CmBu4.Visible = False
Case 1:
    Rahm1.Visible = False
    Rahm2.Visible = True
    CmLab.Visible = False
    CmTyp.Visible = False
    CmBu1.Visible = True
    CmBu2.Visible = True
    CmBu3.Visible = True
    CmBu4.Visible = True
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set clFen = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTabu " & Err.Number
Resume Next

End Sub

Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlKal = False Then FTool Control.id
End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlKal = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    FaPos
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: FaNeu 1, True, 0, True
Case KY_F8: F_Save
Case KY_F11: Unload Me
Case SY_OP_Hilfe: FHilfe
Case SY_OP_Hinzufuegen: FaNeu 1, True, 0, True
Case SY_OP_Speichern: F_Save
Case SY_OP_Loeschen: F_Loe
                     Unload Me
Case SY_OP_Abbruch: Unload Me
Case SY_OP_Sub_Neu: F_Act True
                    F_Neu
Case SY_OP_Sub_Sav: F_Save True
                    F_Act True
Case SY_OP_Sub_Loe: F_ReL
Case SY_OP_Nav_Vor: FMov True
Case SY_OP_Nav_Zuru: FMov False
Case SY_SuCm1: FType
End Select

GlToo = False

End Sub
Private Sub FType()
On Error GoTo LiErr

Dim IdxNr As Long
Dim TypNr As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmLab As XtremeCommandBars.CommandBarControl
Dim CmBu1 As XtremeCommandBars.CommandBarControl
Dim CmBu2 As XtremeCommandBars.CommandBarControl
Dim CmBu3 As XtremeCommandBars.CommandBarControl
Dim CmBu4 As XtremeCommandBars.CommandBarControl
Dim CmBu5 As XtremeCommandBars.CommandBarControl
Dim CmTyp As XtremeCommandBars.CommandBarComboBox
Dim RpCon As XtremeReportControl.ReportControl
Dim ChPfl As XtremeSuiteControls.CheckBox

Set FM = frmFragen
Set ChPfl = FM.chkPflch
Set RpCon = FM.repCont1
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmLab = CmBrs.FindControl(CmLab, SY_Cap02, , True)
Set CmTyp = CmBrs.FindControl(CmTyp, SY_SuCm1, , True)
Set CmBu1 = CmBrs.FindControl(CmBu1, SY_OP_Sub_Neu, , True)
Set CmBu2 = CmBrs.FindControl(CmBu2, SY_OP_Sub_Loe, , True)
Set CmBu3 = CmBrs.FindControl(CmBu3, SY_OP_Nav_Vor, , True)
Set CmBu4 = CmBrs.FindControl(CmBu4, SY_OP_Nav_Zuru, , True)
Set CmBu5 = CmBrs.FindControl(CmBu5, SY_OP_Sub_Sav, , True)

TypNr = CmTyp.ListIndex

Select Case TypNr
Case 1: 'Textfeld
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = True
    FM.txtVorga.Enabled = True
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = True
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = True
    CmBu1.Enabled = False
    CmBu2.Enabled = False
    CmBu3.Enabled = False
    CmBu4.Enabled = False
    CmBu5.Enabled = False
    RpCon.Enabled = False
Case 2: 'Auswahlfeld
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 3: 'Ankreuzfeld
    FM.chkVorga.Enabled = True
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = False
    FM.txtZeich.Enabled = False
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = False
    FM.updCont2.Enabled = False
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 4: 'Einfachauswahl
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = False
    FM.txtZeich.Enabled = False
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = False
    FM.updCont2.Enabled = False
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 5: 'Mehrfachauswahl
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = True
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = False
    FM.txtZeich.Enabled = False
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = False
    FM.updCont2.Enabled = False
    FM.updCont4.Enabled = False
    CmBu1.Enabled = True
    CmBu2.Enabled = True
    CmBu3.Enabled = True
    CmBu4.Enabled = True
    CmBu5.Enabled = True
    RpCon.Enabled = True
Case 6: 'Zwischentext
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = False
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = False
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = False
    CmBu1.Enabled = False
    CmBu2.Enabled = False
    CmBu3.Enabled = False
    CmBu4.Enabled = False
    CmBu5.Enabled = False
    RpCon.Enabled = False
Case 7: 'Datumsfeld
    FM.chkVorga.Enabled = False
    FM.txtBeTex.Enabled = True
    FM.txtVorga.Enabled = False
    FM.txtZeile.Enabled = True
    FM.txtZeich.Enabled = True
    FM.txtMaxZe.Enabled = True
    FM.updCont1.Enabled = True
    FM.updCont2.Enabled = True
    FM.updCont4.Enabled = True
    CmBu1.Enabled = False
    CmBu2.Enabled = False
    CmBu3.Enabled = False
    CmBu4.Enabled = False
    CmBu5.Enabled = False
    RpCon.Enabled = False
End Select

If TypNr = 6 Then
    ChPfl.Enabled = False
Else
    ChPfl.Enabled = True
End If

ChPfl.Value = xtpUnchecked

Set CmBrs = Nothing

If FM.txtBezei.Text <> vbNullString Then
    If TypNr = 1 Or TypNr = 5 Or TypNr = 7 Then
        If FM.txtBeTex.Text <> vbNullString Then
            F_Save True
        End If
    Else
        F_Save True
    End If
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FType " & Err.Number
Resume Next

End Sub
Private Sub Form_Activate()
    FaPos
End Sub
Private Sub Form_Load()
On Error Resume Next
    
Set FrmEx = Me.frmExtde

FoLad = True

With FrmEx
    .ClientMaxHeight = 8000
    .ClientMaxWidth = 15000
    .ClientMinHeight = 6660
    .ClientMinWidth = 10700
    .TopMost = True
End With

Set FrmEx = Nothing

FaInit
AFont Me

FoLad = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmFragen = Nothing
End Sub

Private Sub repCont1_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)

Dim TmTag As String
Dim TmpTg As String

TmpTg = Item.Tag
TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

Item.Tag = "@" & TmTag

F_Act

Item.Tag = TmpTg

End Sub
Private Sub repCont1_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)

Dim TmTag As String
Dim TmpTg As String

TmpTg = Item.Tag
TmTag = Mid$(Item.Tag, 2, Len(Item.Tag) - 1)

Item.Tag = "@" & TmTag

F_Act

Item.Tag = TmpTg

End Sub
Private Sub TbBar_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    If GlKal = False Then FTabu Item.Index
End Sub


Private Sub txtBeTex_GotFocus()
    Me.txtBeTex.SelStart = 0
    Me.txtBeTex.SelLength = Len(Me.txtBeTex.Text)
End Sub

Private Sub txtBeTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtBezei_GotFocus()
    Me.txtBezei.SelStart = 0
    Me.txtBezei.SelLength = Len(Me.txtBezei.Text)
End Sub

Private Sub txtBezei_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtFelNa_GotFocus()
    Me.txtFelNa.SelStart = 0
    Me.txtFelNa.SelLength = Len(Me.txtFelNa.Text)
End Sub

Private Sub txtMaxZe_Change()
On Error Resume Next

If Me.txtMaxZe.Text <> vbNullString Then
    If IsNumeric(Me.txtMaxZe.Text) = True Then
        If CInt(Me.txtMaxZe.Text) > 250 Then
            Me.txtMaxZe.Text = 250
        End If
    Else
        Me.txtMaxZe.Text = 250
    End If
Else
    Me.txtMaxZe.Text = 250
End If

End Sub

Private Sub txtMaxZe_GotFocus()
    Me.txtMaxZe.SelStart = 0
    Me.txtMaxZe.SelLength = Len(Me.txtMaxZe.Text)
End Sub

Private Sub txtMaxZe_LostFocus()
On Error Resume Next

If Me.txtMaxZe.Text <> vbNullString Then
    If IsNumeric(Me.txtMaxZe.Text) = True Then
        If CInt(Me.txtMaxZe.Text) > 250 Then
            Me.txtMaxZe.Text = 250
        End If
    Else
        Me.txtMaxZe.Text = 250
    End If
Else
    Me.txtMaxZe.Text = 250
End If

End Sub


Private Sub txtSorte_Change()
On Error Resume Next

If Me.txtSorte.Text <> vbNullString Then
    If IsNumeric(Me.txtSorte.Text) = True Then
        If CInt(Me.txtSorte.Text) > 999 Then
            Me.txtSorte.Text = 999
        End If
    Else
        Me.txtSorte.Text = 999
    End If
Else
    Me.txtSorte.Text = 999
End If

End Sub
Private Sub txtSorte_GotFocus()
    Me.txtSorte.SelStart = 0
    Me.txtSorte.SelLength = Len(Me.txtSorte.Text)
End Sub

Private Sub txtSorte_LostFocus()
On Error Resume Next

If Me.txtSorte.Text <> vbNullString Then
    If IsNumeric(Me.txtSorte.Text) = True Then
        If CInt(Me.txtSorte.Text) > 999 Then
            Me.txtSorte.Text = 999
        End If
    Else
        Me.txtSorte.Text = 999
    End If
Else
    Me.txtSorte.Text = 999
End If

End Sub
Private Sub txtVorga_GotFocus()
    Me.txtVorga.SelStart = 0
    Me.txtVorga.SelLength = Len(Me.txtVorga.Text)
End Sub

Private Sub txtZeich_Change()
On Error Resume Next

If Me.txtZeich.Text <> vbNullString Then
    If IsNumeric(Me.txtZeich.Text) = True Then
        If CInt(Me.txtZeich.Text) > 40 Then
            Me.txtZeich.Text = 40
        End If
    Else
        Me.txtZeich.Text = 40
    End If
Else
    Me.txtZeich.Text = 40
End If

End Sub
Private Sub txtZeich_GotFocus()
    Me.txtZeich.SelStart = 0
    Me.txtZeich.SelLength = Len(Me.txtZeich.Text)
End Sub

Private Sub txtZeich_LostFocus()
On Error Resume Next

If Me.txtZeich.Text <> vbNullString Then
    If IsNumeric(Me.txtZeich.Text) = True Then
        If CInt(Me.txtZeich.Text) > 40 Then
            Me.txtZeich.Text = 40
        End If
    Else
        Me.txtZeich.Text = 40
    End If
Else
    Me.txtZeich.Text = 40
End If

End Sub
Private Sub txtZeile_Change()
On Error Resume Next

If Me.txtZeile.Text <> vbNullString Then
    If IsNumeric(Me.txtZeile.Text) = True Then
        If CInt(Me.txtZeile.Text) > 8 Then
            Me.txtZeile.Text = 8
        End If
    Else
        Me.txtZeile.Text = 8
    End If
Else
    Me.txtZeile.Text = 8
End If

End Sub

Private Sub txtZeile_GotFocus()
    Me.txtZeile.SelStart = 0
    Me.txtZeile.SelLength = Len(Me.txtZeile.Text)
End Sub

Private Sub txtZeile_LostFocus()
On Error Resume Next

If Me.txtZeile.Text <> vbNullString Then
    If IsNumeric(Me.txtZeile.Text) = True Then
        If CInt(Me.txtZeile.Text) > 8 Then
            Me.txtZeile.Text = 8
        End If
    Else
        Me.txtZeile.Text = 8
    End If
Else
    Me.txtZeile.Text = 8
End If

End Sub

