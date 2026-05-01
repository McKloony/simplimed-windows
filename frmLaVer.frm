VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmLaVer 
   Caption         =   "Ergebnisvergleich"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   ControlBox      =   0   'False
   Icon            =   "frmLaVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   2055
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   3735
      _Version        =   1048579
      _ExtentX        =   6588
      _ExtentY        =   3625
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
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
      Left            =   480
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLaVer"
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

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Sub FClos()
On Error GoTo LiErr

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

If GlIdi = False Then 'Idiotenmodus
    clFen.FenSav
    If clFen.FeSta = 0 Then
        IniSetVal "Laborvergleich", "FenLin", clFen.FeLin
        IniSetVal "Laborvergleich", "FenObe", clFen.FeObn
        IniSetVal "Laborvergleich", "FenBre", clFen.FeBre
        IniSetVal "Laborvergleich", "FenHoh", clFen.FeHoh
    End If
End If

Set clFen = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FClos " & Err.Number
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

Private Sub FMenu()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set FM = frmLaVer
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set ImMan = frmMain.imgManag

With CmBrs
    .EnableActions
    .EnableOffice2007Frame False
    .Icons = ImMan.Icons
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
    CmPan.Width = 140
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    .Visible = True
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
    Set CmCon = .Add(xtpControlButton, SY_OP_Bearbeiten, "Exportieren")
    With CmCon
        .ToolTipText = "Exportiert die angezeigten Ergebnisse"
        .ShortcutText = "F2"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Drucken, "Drucken")
    With CmCon
        .ToolTipText = "Druckt die angezeigten Ergebnisse"
        .ShortcutText = "F1"
        .IconId = IC24_Printer_Ink
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Schließen")
    With CmCon
        .ToolTipText = "Abbrechen"
        .ShortcutText = "F11"
        .IconId = IC24_Exit
        .BeginGroup = True
    End With
End With

Set CmCoS = CmBar.Controls
For Each CmCon In CmCoS
    CmCon.Style = xtpButtonIconAndCaption
Next CmCon

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

Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmLaVer
Set RpCon = FM.repCont1

RpCon.PrintReport 0

Set RpCon = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSet " & Err.Number
Resume Next

End Sub
Private Sub FTool(ByVal TolId As Long)

Select Case TolId
Case KY_F1: FHilfe
Case KY_F2:
Case KY_F3:
Case KY_F10: FPrint
Case KY_F11: Unload Me
Case SY_OP_Oeffnen:
Case SY_OP_Abbruch: Unload Me
Case SY_OP_Drucken: FPrint
End Select

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If GlWLa = False Then FTool Control.id
End Sub
Private Sub comBar02_Resize()
On Error Resume Next

Dim ClRe As RECT

If GlWLa = False Then
    SendMessage Me.hwnd, WM_SETREDRAW, False, 0&
    LVPosi
    SendMessage Me.hwnd, WM_SETREDRAW, True, 0&
    GetClientRect Me.hwnd, ClRe
    RedrawWindow Me.hwnd, ClRe, 0&, RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW
End If

End Sub
Private Sub Form_Load()
    
Set FrmEx = Me.frmExtde

With FrmEx
    .ClientMaxHeight = 14000
    .ClientMaxWidth = 18000
    .ClientMinHeight = 7000
    .ClientMinWidth = 9000
    .TopMost = True
End With

FMenu

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    FClos
    Set frmLaVer = Nothing
End Sub

Private Sub TbBar_SelectedChanged(ByVal Item As XtremeCommandBars.ITabControlItem)
    'FTabu Item.Index
End Sub
