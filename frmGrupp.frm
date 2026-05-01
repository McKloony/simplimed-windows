VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Begin VB.Form frmGrupp 
   Caption         =   "Gruppen"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.TreeView trvList1 
      Height          =   2055
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1815
      _Version        =   1048579
      _ExtentX        =   3201
      _ExtentY        =   3625
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      ForeColor       =   4473924
   End
   Begin XtremeSuiteControls.FlatEdit txtGrupp 
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   3480
      Width           =   2775
      _Version        =   1048579
      _ExtentX        =   4895
      _ExtentY        =   2566
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      MultiLine       =   -1  'True
      ScrollBars      =   2
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
Attribute VB_Name = "frmGrupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private TxTre As XtremeSuiteControls.FlatEdit
Private TxRas As XtremeSuiteControls.FlatEdit
Private TxGrp As XtremeSuiteControls.FlatEdit
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private CmBar As XtremeCommandBars.CommandBar
Private CmSta As XtremeCommandBars.StatusBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions

Private GrLad As Boolean
Private clFen As clsFenster
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
Private Sub GPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmGrupp
Set TrLi1 = FM.trvList1
Set TxGrp = FM.txtGrupp
Set CmBrs = FM.comBar02

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    If ClBre > 1000 And ClHoh > 1000 Then
        TrLi1.Move 0, 540, ClBre, ClHoh - ClObn - 870
        TxGrp.Move 0, ClHoh - ClObn - 300, ClBre, 800
    End If
End If

Set CmBrs = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "GPosi " & Err.Number
Resume Next

End Sub
Private Sub GOpen()
On Error GoTo AnErr
'Setzt die Checkboxen

Dim AktPo As Long
Dim StaPo As Long
Dim GruNr As Long
Dim PatNr As Long
Dim SuStr As String
Dim GruKy As String
Dim Lange As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmGrupp
Set TrLi1 = FM.trvList1
Set TxTre = frmAdress.txtTreKey
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar

StaPo = 2
SuStr = TxTre.Text
Lange = Len(TxTre.Text)

If Lange > 1 Then
    Do
    AktPo = InStr(StaPo, SuStr, "o", vbTextCompare)
    If AktPo > 0 Then
        If Mid$(SuStr, StaPo, AktPo - StaPo) <> vbNullString Then
            GruNr = CLng(Mid$(SuStr, StaPo, AktPo - StaPo))
            GruKy = "G" & GruNr
            For Each Knote In TrLi1.Nodes
                If Knote.Key = GruKy Then
                    TrLi1.Nodes(GruKy).Checked = True
                    Exit For
                End If
            Next Knote
        End If
    End If
    StaPo = AktPo + 1
    Loop Until AktPo = 0
End If

CmSta.Pane(0).Text = "ID: " & Format$(GlAId, "000000")

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "GOpn " & Err.Number
Resume Next

End Sub
Private Sub GSave()
On Error GoTo AnErr
'Speichert die Gruppen zurück

Dim GruKy As String
Dim GrIdx As String
Dim GrStr As String

Set FM = frmGrupp
Set TrLi1 = FM.trvList1
Set TxTre = frmAdress.txtTreKey
Set TxRas = frmAdress.txtAdrGr

For Each Knote In TrLi1.Nodes
    If Knote.Checked = True Then
        If Knote.Key <> "P801" Then
            GrIdx = Mid$(Knote.Key, 2, Len(Knote.Key) - 1)
            If Len(GruKy) = 0 Then
                GruKy = "o" & GrIdx & "o"
            Else
                GruKy = GruKy & GrIdx & "o"
            End If
            If GrStr <> vbNullString Then
                GrStr = GrStr & "; " & Knote.Text
            Else
                GrStr = Knote.Text
            End If
        End If
    End If
Next Knote

If Len(GruKy) = 0 Then
    GruKy = "o0o"
End If

With TxTre
    .Text = GruKy
    .Tag = 1 & "TreKey"
End With

If Len(GrStr) > 250 Then
    GrStr = Left$(GrStr, 250)
End If

With TxRas
    .Text = GrStr
    .Tag = 1 & "AdrGruppe"
End With

GlAdS = True
Adr_Sav
GlAdS = False

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "GSave " & Err.Number
Resume Next

End Sub
Private Sub GText()
On Error GoTo AnErr
'Fasst Gruppen in Textbox zusammen

Dim TmStr As String

Set FM = frmGrupp
Set TrLi1 = FM.trvList1
Set TxGrp = FM.txtGrupp

TxGrp.Text = vbNullString

For Each Knote In TrLi1.Nodes
    If Knote.Checked = True Then
        If TmStr = vbNullString Then
            TmStr = Knote.Text
        Else
            TmStr = TmStr & "; " & Knote.Text
        End If
    End If
Next Knote

TxGrp.Text = TmStr

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "GText " & Err.Number
Resume Next

End Sub
Private Sub GInit()
On Error GoTo MnErr

Dim RetWe As Long
Dim KeyNa As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Set FM = frmGrupp
Set TrLi1 = FM.trvList1
Set TxGrp = FM.txtGrupp
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
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Width = 70
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
        .BeginGroup = True
        .IconId = IC24_Disk_Norm
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .BeginGroup = True
        .IconId = IC24_Help
        .ShortcutText = "F1"
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Abbrechen")
    With CmCon
        .BeginGroup = True
        .IconId = IC24_Exit
    End With
End With

Set CmCoS = CmBar.Controls
For Each CmCon In CmCoS
    CmCon.Style = xtpButtonIconAndCaption
Next CmCon

With TrLi1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .Checkboxes = True
    .Font.SIZE = GlTFt.SIZE
    .Font.Name = GlTFt.Name
    .ForeColor = -2147483641
    .FullRowSelect = False
    .HideSelection = False
    .HotTracking = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpTreeViewLabelManual
    .Scroll = True
    .ShowLines = xtpTreeViewShowLines
    .ShowPlusMinus = True
    .SingleSel = False
End With

Set Knote = TrLi1.Nodes.Add(, , "P801", "Adressen", IC16_Folder_View)
With Knote
    .Bold = True
    .Checked = False
End With

With TxGrp
    .BackColor = -2147483643
    .Font.SIZE = GlTFt.SIZE
    .Font.Name = GlTFt.Name
End With

FM.BackColor = GlBak

clFen.FenVor

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Set clFen = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "GInit " & Err.Number
Resume Next

End Sub
Private Sub GTool(ByVal TolId As Long)

Select Case TolId
Case SY_OP_Hilfe: FHilfe
Case SY_OP_Speichern: GSave
                      Unload Me
Case SY_OP_Abbruch: Unload Me
End Select

End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    GTool Control.id
End Sub
Private Sub comBar02_Resize()
    GPosi
End Sub
Private Sub Form_Load()
On Error Resume Next

GrLad = True
GInit
AFont Me
AdGru 1
GOpen
GText
GrLad = False
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmGrupp = Nothing
End Sub

Private Sub trvList1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo LaErr

Set TrLi1 = Me.trvList1

If Button = vbRightButton Then
    Set TrLi1.SelectedItem = TrLi1.HitTest(x, y)
End If

LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub trvList1_NodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Set TrLi1 = Me.trvList1

If GrLad = False Then
    For Each Knote In TrLi1.Nodes
        Knote.Image = IC16_Folder_Close
    Next Knote
    
    Node.Image = IC16_Folder_Open
    TrLi1.Nodes(1).Image = IC16_Folder_View
    
    GText
End If

End Sub
Private Sub trvList1_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Set TrLi1 = Me.trvList1
    
For Each Knote In TrLi1.Nodes
    Knote.Image = IC16_Folder_Close
Next Knote

Node.Image = IC16_Folder_Open
TrLi1.Nodes(1).Image = IC16_Folder_View

End Sub
