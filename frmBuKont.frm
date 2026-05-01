VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmBuKont 
   Caption         =   "Sachkontenplan"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   8085
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1455
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2895
      _Version        =   1048579
      _ExtentX        =   5106
      _ExtentY        =   2566
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3200
      Left            =   600
      TabIndex        =   1
      Top             =   3000
      Width           =   7000
      _Version        =   1048579
      _ExtentX        =   12347
      _ExtentY        =   5644
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkGewEr 
         Height          =   220
         Left            =   1600
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Keine Auswertung bei Gewinnermittlung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkSteKo 
         Height          =   255
         Left            =   3800
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1520
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Steuerkontenzuordnung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkGldKo 
         Height          =   255
         Left            =   3800
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1160
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Geldkontenzuordnung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBezei 
         Height          =   350
         Left            =   1600
         TabIndex        =   5
         Top             =   700
         Width           =   4200
         _Version        =   1048579
         _ExtentX        =   7408
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtKonto 
         Height          =   350
         Left            =   1600
         TabIndex        =   6
         Top             =   300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbGegen 
         Height          =   315
         Left            =   1600
         TabIndex        =   7
         Top             =   1900
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
      Begin XtremeSuiteControls.ComboBox cmbBuStu 
         Height          =   310
         Left            =   1600
         TabIndex        =   8
         Top             =   1500
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbBuTyp 
         Height          =   315
         Left            =   1600
         TabIndex        =   9
         Top             =   1100
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.ComboBox cmbBuTex 
         Height          =   310
         Left            =   1600
         TabIndex        =   10
         Top             =   2310
         Width           =   4200
         _Version        =   1048579
         _ExtentX        =   7408
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
      End
      Begin VB.Label lblLab30 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kontonummer :"
         Height          =   210
         Left            =   300
         TabIndex        =   16
         Top             =   340
         Width           =   1200
      End
      Begin VB.Label lblLab32 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Geldkonto :"
         Height          =   210
         Left            =   300
         TabIndex        =   15
         Top             =   1960
         Width           =   1200
      End
      Begin VB.Label lblLab34 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Kontoart :"
         Height          =   210
         Left            =   300
         TabIndex        =   14
         Top             =   1160
         Width           =   1200
      End
      Begin VB.Label lblLab31 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Bezeichnung :"
         Height          =   210
         Left            =   300
         TabIndex        =   13
         Top             =   740
         Width           =   1200
      End
      Begin VB.Label lblLab33 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Steuersatz :"
         Height          =   210
         Left            =   300
         TabIndex        =   12
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label lblLab35 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Buchungstext :"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   2380
         Width           =   1200
      End
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   500
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBuKont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private ChGew As XtremeSuiteControls.CheckBox
Private ChGel As XtremeSuiteControls.CheckBox
Private ChSte As XtremeSuiteControls.CheckBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private CmBuT As XtremeSuiteControls.ComboBox
Private CmStu As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private TxDa1 As XtremeSuiteControls.FlatEdit
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
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157

Private SuStr As String
Private KtNeu As Boolean
Private FoLad As Boolean

Private clFen As clsFenster

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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
Private Sub FKonf()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim AktZa As Integer
Dim AktKo As Integer
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmBuKont
Set ChGew = FM.chkGewEr
Set ChGel = FM.chkGldKo
Set ChSte = FM.chkSteKo
Set CmGeg = FM.cmbGegen
Set CmBuT = FM.cmbBuTex
Set CmStu = FM.cmbBuStu
Set CmTyp = FM.cmbBuTyp
Set RpCon = FM.repCont1
Set Rahm1 = FM.frmRahm1
Set ImMan = frmMain.imgManag

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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Buchungskonten vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Buchungskonten vorhanden"
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
    .PaintManager.MaxPreviewLines = 5
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.FixedRowHeight = True
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

For AktZa = 1 To UBound(GlStu)
    CmStu.AddItem GlStu(AktZa, 2)
    CmStu.ItemData(CmStu.NewIndex) = GlStu(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlKoA) - 2
    CmTyp.AddItem GlKoA(AktZa, 0)
    CmTyp.ItemData(CmTyp.NewIndex) = CInt(GlKoA(AktZa, 1))
Next AktZa

With CmGeg
    For AktZa = 1 To UBound(GlGeK) 'Geldkonten
        .AddItem GlGeK(AktZa, 3)
        .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
    Next AktZa
End With

If CmBuT.ListCount = 0 Then
    For AktZa = 1 To UBound(GlBTe)
        CmBuT.AddItem GlBTe(AktZa, 1)
        CmBuT.ItemData(AktZa - 1) = GlBTe(AktZa, 0)
    Next AktZa
End If
CmBuT.AutoComplete = True

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
ChGew.BackColor = GlBak
ChGel.BackColor = GlBak
ChSte.BackColor = GlBak

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FMenu()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim ComSu As XtremeCommandBars.CommandBarComboBox
Dim ComTy As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set FM = frmBuKont
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmAcs = CmBrs.Actions
Set CmOpt = CmBrs.Options
Set ImMan = frmMain.imgManag

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
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
    .GlobalSettings.App = App
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
    .ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
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
    .ActiveMenuBar.ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
    .KeyBindings.Add 0, VK_F1, KY_F1
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F8, KY_F8
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
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuFe1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_OP_Uebernahme, vbNullString, vbNullString, vbNullString, vbNullString)
End With

Set CmBar = CmBrs.Add("ID_Toolbar", xtpBarTop)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarTop
    .SetIconSize 24, 24
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
    .ModifyStyle XTP_CBRS_GRIPPER, XTP_CBRS_GRIPPER
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Hinzufügen")
    With CmCon
        .ToolTipText = "Legt ein neues Konto an"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Uebernahme, "Übernehmen")
    With CmCon
        .ToolTipText = "Übernimmt Konto für die Buchung"
        .ShortcutText = "F6"
        .IconId = IC24_Nav_Down_Right
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .ToolTipText = "Speichert die Änderungen"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .ToolTipText = "Löscht die markierten Konten"
        .BeginGroup = True
        .IconId = IC24_Doc_Del
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Drucken, "Drucken")
    With CmCon
        .ToolTipText = "Druckt den Kontenplan"
        .ShortcutText = "F10"
        .IconId = IC24_Printer_Ink
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .ShortcutText = "F1"
        .IconId = IC24_Help
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Abbrechen")
    With CmCon
        .ToolTipText = "Schließt den Dialog"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

Set CmCoS = CmBar.Controls
For Each CmCon In CmCoS
    CmCon.Style = xtpButtonIconAndCaption
Next CmCon

'----------------------------------------------------------------------

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
    Set CmCon = .Add(xtpControlLabel, KA_Capt1, "Kontenrahmen:")
    With CmCon
        .ToolTipText = "Bitte wählen Sie den gewünschten Kontenrahmen"
        .Style = xtpButtonIconAndCaption
    End With
    Set ComSu = .Add(xtpControlComboBox, KA_SuCo1, vbNullString)
    With ComSu
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Wählen Sie hier eine Kontenrahmen"
        .ThemedItems = True
        .Width = 130
        For AktZa = 1 To UBound(GlKoR) 'Standardkontenrahmen
            .AddItem GlKoR(AktZa, 0)
            .ItemData(AktZa) = GlKoR(AktZa, 1)
        Next AktZa
        .ListIndex = GlKtR 'Standardkontenrahmen
    End With
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(2))
    Set CmEdi = .Add(xtpControlEdit, KA_SuFe1, "Suche: ")
    With CmEdi
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Suchkriterium..."
        .IconId = IC16_View
        .Style = xtpButtonIconAndCaption
        .ToolTipText = "Geben Sie bitte hier das Suchkriterium ein und bestätigen mit der ENTER-Taste"
        .Width = 200
    End With
    Set CmCon = .Add(xtpControlLabel, SY_Plac2, Space$(2))
    Set CmCon = .Add(xtpControlLabel, KA_Capt2, "Kontenfilter:")
    With CmCon
        .ToolTipText = "Bitte wählen Sie die gwünschte Kontenart"
        .Style = xtpButtonIconAndCaption
    End With
    Set ComTy = .Add(xtpControlComboBox, KA_SuCo2, vbNullString)
    With ComTy
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Bitte wählen Sie die gwünschte Kontenart"
        .ThemedItems = True
        .Width = 110
        For AktZa = 1 To UBound(GlKoA)
            .AddItem GlKoA(AktZa, 0), GlKoA(AktZa, 1)
        Next AktZa
        .ListIndex = 5
    End With
End With

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMenu " & Err.Number
Resume Next

End Sub
Private Sub FPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmBuKont
Set RpCon = FM.repCont1
Set Rahm1 = FM.frmRahm1
Set CmBrs = FM.comBar02

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    RpCon.Move 0, ClObn, ClBre, ClHoh - 4110
    Rahm1.Move 60, ClHoh - 3260, ClBre - 120
End If

Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPosi " & Err.Number
Resume Next

End Sub
Private Sub FDruk()
On Error GoTo InErr

Dim IdxNr As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmBuKont
Set CmBrs = FM.comBar02

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

IdxNr = CmCom.ListIndex

Unload FM

SDruck "BuKont", True, , , "[IDR] = " & IdxNr

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDruk " & Err.Number
Resume Next

End Sub
Private Sub FNeu()
On Error GoTo InErr

Me.txtKonto.Text = vbNullString
Me.txtBezei.Text = vbNullString
Me.cmbGegen.ListIndex = 0
Me.cmbBuStu.ListIndex = 0
Me.cmbBuTyp.ListIndex = 0

Me.txtKonto.SetFocus

KtNeu = True

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeu " & Err.Number
Resume Next

End Sub
Private Sub FUber()
On Error GoTo InErr

Dim BehNr As Long
Dim BelNr As Long
Dim BelTm As Long
Dim IdxNr As Long
Dim TmpNr As Long
Dim BuJah As Long
Dim KtSt1 As String
Dim KtSt2 As String
Dim KoStu As Single
Dim AktZa As Integer
Dim IdBnk As Integer
Dim Mld1, Tit1 As String
Dim CmMan As XtremeSuiteControls.ComboBox
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CmGlk As XtremeCommandBars.CommandBarComboBox

If WindowLoad("frmBuEdit") = False Then
    Mld1 = "Es ist erforderlich erst eine neue Buchung hinzuzufügen, bevor das Sachkonto übernommen werden kann."
    Tit1 = "Keine neue Buchung"
    SPopu Tit1, Mld1, IC48_Information
    Exit Sub
End If

Set FM = frmBuEdit
Set CmMan = FM.cmbManda
Set TxDa1 = FM.txtDatu1
Set RpCon = Me.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows
Set CmBrs = frmMain.comBar01
Set RpCo1 = frmMain.repCont1

Set CmGlk = CmBrs.FindControl(CmGlk, SY_SuBuh, , True)

IdBnk = CmGlk.ItemData(CmGlk.ListIndex)

If CmMan.Text <> vbNullString Then
    BehNr = CmMan.ItemData(CmMan.ListIndex)
Else
    BehNr = 0
End If

If IsDate(TxDa1.Text) = True Then
    BuJah = Year(TxDa1.Text)
Else
    BuJah = Year(Date)
End If

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Kon_Steuer)
        KoStu = Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1)
        Set RpCol = RpCls.Find(Kon_IDK)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            KtSt1 = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
        End If
        Set RpCol = RpCls.Find(Kon_IDKurz)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            KtSt2 = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
        Else
            KtSt2 = vbNullString
        End If
        FM.txtKonto.Text = KtSt1 & Chr$(32) & KtSt2
        FM.txtBezei.Text = KtSt2
        FM.txtKtoNr.Text = KtSt1
        If IdBnk = 0 Then
            Set RpCol = RpCls.Find(Kon_IDB)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                IdxNr = SCmb(FM.cmbGegen, RpRow.Record(RpCol.ItemIndex).Value)
                FM.cmbGegen.ListIndex = IdxNr
            End If
        Else
            FM.cmbGegen.ListIndex = CmGlk.ListIndex - 1
        End If
        For AktZa = 1 To UBound(GlStu)
            If CSng(GlStu(AktZa, 1)) = KoStu Then
                FM.cmbBuStu.ListIndex = AktZa - 1
                Exit For
            End If
        Next AktZa
        Set RpCol = RpCls.Find(Kon_Selekt)
        If RpRow.Record(RpCol.ItemIndex).Value = -1 Then
            FM.chkGewEr.Value = 1
        Else
            FM.chkGewEr.Value = 0
        End If
        Set RpCol = RpCls.Find(Kon_IDA)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            TmpNr = RpRow.Record(RpCol.ItemIndex).Value
            If TmpNr > 2 Then TmpNr = 1
            IdxNr = SCmb(FM.cmbBuTyp, TmpNr)
            FM.cmbBuTyp.ListIndex = IdxNr
        End If

        FM.txtBuBel.Text = S_BuBel(BehNr, BuJah, IdBnk)
    End If
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set RpCo1 = Nothing

Unload Me

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FUber " & Err.Number
Resume Next

End Sub
Private Sub FOpn()
On Error GoTo InErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmBuKont
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

With clFen
    If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
        .FeLin = (GlxGr - GlFeB) / 2
        .FeObn = (GlyGr - GlFeH) / 2
        .FeBre = 640
        .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
    Else
        .FeLin = 80
        .FeObn = 10
        .FeBre = 640
        .FeHoh = GlyGr - 20
    End If
    .FenMov
End With

CmAcs(SY_OP_Uebernahme).Enabled = WindowLoad("frmBuEdit")

S_KtLa
S_KtPo

Set CmBrs = Nothing

Set clFen = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpn " & Err.Number
Resume Next

End Sub
Private Sub FPlan()
On Error GoTo InErr

S_KtLa
S_KtPo

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPlan " & Err.Number
Resume Next

End Sub
Private Sub FSave()
On Error GoTo MeErr

S_KtSa KtNeu, SuStr
KtNeu = False

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSpla()
On Error GoTo InErr

Dim AktZa As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmBuKont
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns

With RpCls
    Set RpCol = .Add(Kon_IDK, "Sachkonto", 90, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(Kon_IDKurz, "Bezeichnung", 10, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
        .AutoSize = True
    End With
    If RpCon.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Kon_Typ, "Kontoart", 120, False)
    With RpCol
        .Alignment = xtpAlignmentIconLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        For AktZa = 1 To UBound(GlKoA)
            .EditOptions.Constraints.Add GlKoA(AktZa, 0), GlKoA(AktZa, 1)
        Next AktZa
    End With
    If GlBuc = True Then 'Einfache Buchhaltung verwenden
        Set RpCol = .Add(Kon_Bank, "Geldkonto", 90, False)
    Else
        Set RpCol = .Add(Kon_Bank, "Geldkonto", 0, False)
    End If
    With RpCol
        .Alignment = xtpAlignmentIconLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(Kon_IDA, vbNullString, 0, False)
    Set RpCol = .Add(Kon_IDB, vbNullString, 0, False)
    Set RpCol = .Add(Kon_Buchtext, vbNullString, 0, False)
    Set RpCol = .Add(Kon_Steuer, vbNullString, 0, False)
    Set RpCol = .Add(Kon_Privat, vbNullString, 0, False)
    Set RpCol = .Add(Kon_IDArt, vbNullString, 0, False)
    Set RpCol = .Add(Kon_IDBank, vbNullString, 0, False)
    Set RpCol = .Add(Kon_Selekt, vbNullString, 0, False)
    Set RpCol = .Add(Kon_IDSuch, vbNullString, 0, False)
    Set RpCol = .Add(Kon_IDI, "Index", 50, False)
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FSuFe()
On Error GoTo MeErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit

Set FM = frmBuKont
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmEdt = CmBrs.FindControl(CmEdt, KA_SuFe1, , True)

If CmEdt.Text <> vbNullString Then
    SuStr = CmEdt.Text
Else
    SuStr = vbNullString
End If

S_KtFi SuStr
DoEvents

CmEdt.Text = vbNullString

Set CmBrs = Nothing

Exit Sub

MeErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuFe " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3: FNeu
Case KY_F5:
Case KY_F6: If FoLad = False Then FUber
Case KY_F8: If FoLad = False Then FSave
Case KY_F10: FDruk
Case KY_F11: Unload FM
Case SY_OP_Hinzufuegen: FNeu
Case SY_OP_Uebernahme: If FoLad = False Then FUber
Case SY_OP_Speichern: If FoLad = False Then FSave
Case SY_OP_Loeschen: S_KtLo
Case SY_OP_Drucken: FDruk
Case SY_OP_Abbruch: Unload Me
Case SY_OP_Hilfe: FHilfe
Case KA_SuFe1: FSuFe
Case KA_SuCo1: If FoLad = False Then FPlan
Case KA_SuCo2: If FoLad = False Then S_KtTy
End Select

End Sub

Private Sub chkGldKo_Click()
On Error Resume Next

Set FM = frmBuKont
Set ChGel = FM.chkGldKo
Set ChSte = FM.chkSteKo

If ChGel.Value = xtpChecked Then
    ChSte.Value = xtpUnchecked
    ChSte.Enabled = False
Else
    ChSte.Enabled = True
End If

If ChSte.Value = xtpChecked Then
    ChGel.Value = xtpUnchecked
    ChGel.Enabled = False
Else
    ChGel.Enabled = True
End If

End Sub

Private Sub chkSteKo_Click()
On Error Resume Next

Set FM = frmBuKont
Set ChGel = FM.chkGldKo
Set ChSte = FM.chkSteKo

If ChGel.Value = xtpChecked Then
    ChSte.Value = xtpUnchecked
    ChSte.Enabled = False
Else
    ChSte.Enabled = True
End If

If ChSte.Value = xtpChecked Then
    ChGel.Value = xtpUnchecked
    ChGel.Enabled = False
Else
    ChGel.Enabled = True
End If

End Sub
Private Sub cmbBuTex_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Dim RetWe As Long
Dim State As Long

Set CmBuT = Me.cmbBuTex

If KeyCode = vbKeyF2 Then
    CmBuT.SelLength = 0
ElseIf KeyCode = vbKeyDown Then
    State = SendMessage(CmBuT.hwnd, CB_GETDROPPEDSTATE, 0, 0)
    If State = 0 Then State = SendMessage(CmBuT.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End If

End Sub

Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
Private Sub comBar02_Resize()
    FPosi
End Sub

Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 14000
    .ClientMaxWidth = 11000
    .ClientMinHeight = 6000
    .ClientMinWidth = 9500
End With

Screen.MousePointer = vbHourglass

FoLad = True
FKonf
AFont Me
FMenu
FSpla
FOpn
FoLad = False
SFrame 1, Me.hwnd

Set FrmEx = Nothing

Screen.MousePointer = vbNormal

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBuKont = Nothing
End Sub

Private Sub repCont1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        If KeyCode = vbKeyReturn Then
            KeyCode = 0
            FUber
        End If
    End If
End Sub
Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
    If FoLad = False Then
        If Shift = 0 Then
            S_KtPo
        End If
    End If
End Sub
Private Sub repCont1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    If FoLad = False Then
        S_KtPo
    End If
End Sub
Private Sub repCont1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    FUber
End Sub
Private Sub repCont1_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If FoLad = False Then
        S_KtPo
    End If
End Sub
Private Sub txtBezei_GotFocus()
    Me.txtBezei.SelStart = 0
    Me.txtBezei.SelLength = Len(Me.txtBezei.Text)
End Sub
Private Sub txtKonto_GotFocus()
    Me.txtKonto.SelStart = 0
    Me.txtKonto.SelLength = Len(Me.txtKonto.Text)
End Sub

