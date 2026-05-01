VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Begin VB.Form frmWebCam 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "WebCam"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picBild2 
      BorderStyle     =   0  'Kein
      Height          =   3600
      Left            =   0
      ScaleHeight     =   3600
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   4800
   End
   Begin XtremeCommandBars.CommandBars comBar02 
      Left            =   240
      Top             =   120
      _Version        =   1048579
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmWebCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PiBi2 As VB.PictureBox
Private TxDe5 As XtremeSuiteControls.FlatEdit
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions

Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WM_USER = &H400
Private Const WM_CAP_START = &H400
Private Const WM_CAP_EDIT_COPY = (WM_CAP_START + 30)
Private Const WM_CAP_DRIVER_CONNECT = (WM_CAP_START + 10)
Private Const WM_CAP_SET_PREVIEWRATE = (WM_CAP_START + 52)
Private Const WM_CAP_SET_OVERLAY = (WM_CAP_START + 51)
Private Const WM_CAP_SET_PREVIEW = (WM_CAP_START + 50)
Private Const WM_CAP_DRIVER_DISCONNECT = (WM_CAP_START + 11)
 
Private ViHwn As Long
 
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal PosLi As Long, ByVal PosOb As Long, ByVal PoWei As Long, ByVal PoHoh As Long, ByVal FoHwn As Long, ByVal nID As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private clFil As clsFile
Private Function SBiCa(FoHwn As Long, Optional PosLi As Long = 0, Optional PosOb As Long = 0, Optional PoWei As Long = 640, Optional PoHoh As Long = 480, Optional CamId As Long = 0) As Long
On Error Resume Next

Dim PrHwn As Long
 
PrHwn = capCreateCaptureWindow("Video", WS_CHILD + WS_VISIBLE, PosLi, PosOb, PoWei, PoHoh, FoHwn, 1)

SendMessage PrHwn, WM_CAP_DRIVER_CONNECT, CamId, 0
SendMessage PrHwn, WM_CAP_SET_PREVIEWRATE, 30, 0
SendMessage PrHwn, WM_CAP_SET_OVERLAY, 1, 0
SendMessage PrHwn, WM_CAP_SET_PREVIEW, 1, 0

SBiCa = PrHwn

End Function
Private Sub SBiCo(PrHwn As Long, Optional CamId = 0)
    SendMessage PrHwn, WM_CAP_DRIVER_DISCONNECT, CamId, 0
End Sub
Private Sub FTool(ByVal TolId As Long)
On Error Resume Next

If GlToo = True Then Exit Sub

GlToo = True

Select Case TolId
Case KY_F1:
Case KY_F8: FBiVi
Case KY_F11: Unload Me
Case SY_OP_Speichern: FBiVi
Case SY_OP_Hilfe:
Case SY_OP_Abbruch: Unload Me
End Select

GlToo = False

End Sub
Private Sub FMenu()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set FM = frmWebCam
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
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
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .ToolTipText = "Das aktuelle Videobild Grabben"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
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
Private Sub FBiVi()
On Error GoTo OpErr

Dim FiNam As String
Dim FiPat As String

Set FM = frmWebCam
Set PiBi2 = FM.picBild2
Set TxDe5 = frmMain.txtDeta5

Set clFil = New clsFile

FiPat = "A" & Format$(GlAdr, "000000") & ".jpg"
FiNam = GlBPf & FiPat

Clipboard.Clear
DoEvents
SendMessage ViHwn, WM_CAP_EDIT_COPY, 0, 0
DoEvents

PiBi2.Picture = Clipboard.GetData
DoEvents

With clFil
    If Not IsNull(FiNam) And Not FiNam = vbNullString Then
        If clFil.FilVor(FiNam) = True Then
            .DaLoe = FiNam & vbNullChar
            .FilLoe
        End If
    End If
End With

Set clFil = Nothing

SavePicture PiBi2.Picture, FiNam
DoEvents
TxDe5.Text = FiPat
GlSav = True
S_KrSv
GlSav = False
DoEvents
If GlABi = True Then
    SPaBi
End If
DoEvents

Unload Me

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FBiVi " & Err.Number
Resume Next

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
Private Sub Form_Load()
On Error Resume Next

FMenu
SFrame 1, Me.hwnd
DoEvents
ViHwn = SBiCa(Me.hwnd, 0, 36)

End Sub
Private Sub Form_Unload(Cancel As Integer)
    SBiCo ViHwn
    Set frmWebCam = Nothing
End Sub
