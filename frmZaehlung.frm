VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#16.3#0"; "Codejock.CommandBars.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmZaehlung 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kassenzählung"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeReportControl.ReportControl repCont1 
      Height          =   1900
      Left            =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1000
      Width           =   6780
      _Version        =   1048579
      _ExtentX        =   11959
      _ExtentY        =   3351
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin XtremeReportControl.ReportControl repCont2 
      Height          =   5940
      Left            =   100
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3020
      Width           =   6780
      _Version        =   1048579
      _ExtentX        =   11959
      _ExtentY        =   10477
      _StockProps     =   64
      FreezeColumnsAbs=   0   'False
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   9800
      Width           =   80
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9800
      Visible         =   0   'False
      Width           =   615
      _Version        =   1048579
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   64
      Show3DBorder    =   2
      VisualTheme     =   0
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
Attribute VB_Name = "frmZaehlung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private ColMa As XtremeCommandBars.ColorManager
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmAct As XtremeCommandBars.CommandBarAction
Private TxDum As XtremeSuiteControls.FlatEdit
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
Private MoKal As XtremeCalendarControl.DatePicker

Dim FoLad As Boolean
Dim FoNeu As Boolean

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private clLis As clsLisLab
Private clFil As clsFile

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASE = &H4
Private Const GWL_WNDPROC = (-4)
Private Const KEYEVENTF_KEYUP = &H2

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim MoKal As XtremeCalendarControl.DatePicker
Dim CmEdi As XtremeCommandBars.CommandBarEdit

Set CmBrs = Me.comBar02
Set MoKal = Me.dtpDatu1

Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    If NeuDa > Date Then
        CmEdi.Text = Date
        SPopu "Zählung liegt in der Zukunft", "Der Termin der Zählung darf nicht in der Zukunft liegen", IC48_Forbidden
    Else
        CmEdi.Text = NeuDa
    End If
End If

Set MoKal = Nothing

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

TeTit = IniGetOpt("Hilfe", 50761)
TeMai = IniGetOpt("Hilfe", 50762)
TeInh = IniGetOpt("Hilfe", 50763)
TeFus = IniGetOpt("Hilfe", 50764)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim StaGe As Long
Dim AktZa As Integer
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmMain
Set MoKal = Me.dtpDatu1
Set RpCo1 = Me.repCont1
Set RpCo2 = Me.repCont2
Set ImMan = FM.imgManag

With RpCo1
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
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderFlat
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = False
    .Icons = ImMan.Icons
    .MultipleSelection = False
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
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Kassenzählungen vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Kassenzählungen vorhanden"
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
    .PreviewMode = False
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With RpCo2
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
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderFlat
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = False
    .Icons = ImMan.Icons
    .MultipleSelection = False
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
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Zählungen vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Zählungen vorhanden"
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
    .PreviewMode = False
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With MoKal
    .AllowNoncontinuousSelection = False
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    If GlSty = 8 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    ElseIf GlSty = 7 Then 'Office 2013
        .BorderStyle = xtpDatePickerBorderStatic
    Else
        .BorderStyle = xtpDatePickerBorderOffice
    End If
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    .MaxSelectionCount = 1
    .RightToLeft = False
    .ShowNoneButton = True
    .ShowNonMonthDays = True
    .ShowTodayButton = True
    .ShowWeekNumbers = False
    .TextNoneButton = "Keine"
    .TextTodayButton = "Heute"
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Behandlungstag"
    .MonthDelta = 1
    .YearsTriangle = False
    Select Case GlSty
    Case 8: .VisualTheme = xtpCalendarThemeResource
    Case 7: .VisualTheme = xtpCalendarThemeResource
    Case Else: .VisualTheme = xtpCalendarThemeResource
    End Select
    .PaintManager.ButtonTextColor = -2147483640
    .PaintManager.ControlBackColor = -2147483643
    .PaintManager.DayBackColor = -2147483643
    .PaintManager.DayTextColor = -2147483640
    .PaintManager.DaysOfWeekBackColor = -2147483643
    .PaintManager.DaysOfWeekTextColor = -2147483640
    .PaintManager.ListControlBackColor = -2147483643
    .PaintManager.ListControlTextColor = -2147483640
    .PaintManager.NonMonthDayBackColor = -2147483643
    .PaintManager.NonMonthDayTextColor = -2147483640
    .PaintManager.SelectedDayBackColor = GlFac
    .PaintManager.SelectedDayTextColor = -2147483640
    .PaintManager.WeekNumbersBackColor = -2147483643
    .PaintManager.WeekNumbersTextColor = -2147483640
    .PaintManager.MonthHeaderBackColor = GlMoB
End With

Me.BackColor = GlBak

Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "Finit " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim KaWei As Long
Dim KaHoh As Long
Dim StaDa As Date
Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim MoKal As XtremeCalendarControl.DatePicker

Set FM = frmZaehlung
Set CmBrs = FM.comBar02
Set MoKal = FM.dtpDatu1

Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)

If CmEdi.Text <> vbNullString Then
    If IsDate(CmEdi.Text) = True Then
        StaDa = CDate(CmEdi.Text)
    Else
        StaDa = Date
    End If
Else
    StaDa = Date
End If

If StaDa > Date Then
    StaDa = Date
End If

With MoKal
    .GetMinReqRect KaWei, KaHoh, 1, 1
    .EnsureVisible StaDa
    .Select StaDa
    .SelectRange StaDa, StaDa
    If .ShowModalEx(51, 88, KaWei, KaHoh, FM.hwnd) = True Then
        If .Selection.BlocksCount > 0 Then
            NeuDa = .Selection.Blocks(0).DateBegin
            CmEdi.Text = Format$(NeuDa, "dd.mm.yyyy")
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKran(Optional ByVal CoIdx As Long)
On Error GoTo PoErr
'Änderungen im Krankneblatt

Dim KrRow As Long
Dim MuWer As Single
Dim GeWer As Single
Dim Fakto As Single
Dim GeSum As Single
Dim RowNr As Integer
Dim Anzal As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmZaehlung
Set CmBrs = FM.comBar02
Set RpCo2 = FM.repCont2
Set RpCls = RpCo2.Columns
Set RpRcs = RpCo2.Records
Set RpSel = RpCo2.SelectedRows
Set CmSta = CmBrs.StatusBar

If CoIdx = 3 Then
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        RowNr = RpRow.Index
        If RpRow.GroupRow = False Then
            If IsNull(RpRow.Record(2).Value) = False Then
                MuWer = CSng(RpRow.Record(2).Value)
            Else
                MuWer = 0
            End If
            If IsNull(RpRow.Record(3).Value) = False Then
                If RpRow.Record(3).Value <> vbNullString Then
                    If IsNumeric(RpRow.Record(3).Value) = True Then
                        Anzal = RpRow.Record(3).Value
                    Else
                        Anzal = 0
                    End If
                Else
                    Anzal = 0
                End If
            Else
                Anzal = 0
            End If
            GeWer = MuWer * Anzal
            RpRow.Record(4).Value = Format$(GeWer, GlWa1)
                        
            For Each RpRec In RpRcs
                If RpRec.Item(4).Value <> vbNullString Then
                    If IsNumeric(RpRec.Item(4).Value) = True Then
                        If CSng(RpRec.Item(4).Value) > 0 Then
                            GeSum = GeSum + CSng(RpRec.Item(4).Value)
                        End If
                    End If
                End If
            Next RpRec

            CmSta.Pane(1).Text = "Gesamt : " & Format$(GeSum, GlWa1) & Space$(3)
        End If
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo2 = Nothing
Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKran " & Err.Number
Resume Next

End Sub
Private Sub FMenu()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim AktPo As Integer
Dim AktKo As Integer
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGeg As XtremeCommandBars.CommandBarComboBox
Dim CmCoZ As XtremeCommandBars.CommandBarComboBox

Set FM = frmZaehlung
Set CmBrs = FM.comBar02
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag

AktPo = 1

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
    CmPan.Alignment = xtpAlignmentRight
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Width = 200
    CmPan.Text = vbNullString
    CmPan.Alignment = xtpAlignmentRight
    .Visible = True
End With

With CmAcs
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_KaBu1, vbNullString, vbNullString, vbNullString, vbNullString)
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
    Set CmCon = .Add(xtpControlButton, SY_OP_Hinzufuegen, "Neue Zählung")
    With CmCon
        .ToolTipText = "Erstellt eine neue Kassenzählung"
        .ShortcutText = "F3"
        .IconId = IC24_Doc_Add
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Speichern")
    With CmCon
        .ToolTipText = "Speichert die aktuelle Kassenzählung"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Loeschen, "Entfernen")
    With CmCon
        .ToolTipText = "Löscht die aktuelle Kassenzählung"
        .IconId = IC24_Doc_Del
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Drucken, "Drucken")
    With CmCon
        .ToolTipText = "Druckt das Kassenzählprotokoll"
        .ShortcutText = "F10"
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
    Set CmEdi = .Add(xtpControlEdit, KA_Kalen, " Datum :")
    With CmEdi
        .ToolTipText = "Wählen Sie hier das Datum aus, zu dem die Kassenzählung durchgeführt wurde"
        .Style = xtpButtonCaption
        .IconId = IC16_Calendar_Year
        .Width = 120
    End With
    Set CmCon = .Add(xtpControlButton, KA_KaBu1, vbNullString)
    With CmCon
        .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
        .Style = xtpButtonIcon
        .IconId = IC16_Calendar_Year
    End With

    Set CmCon = .Add(xtpControlLabel, KA_Capt3, " Uhrzeit :")
    With CmCon
        .ToolTipText = "Wählen Sie hier die Uhrzeit aus, unter dem der Eintrag gespeichetr werden soll"
        .Style = xtpButtonCaption
    End With
    Set CmCoZ = .Add(xtpControlComboBox, KA_Uhrze, vbNullString)
    With CmCoZ
        .ToolTipText = "Wählen Sie hier die Uhrzeit aus, unter dem der Eintrag gespeichetr werden soll"
        .Style = xtpButtonCaption
        .IconId = IC16_Key_Kopf
        .EditStyle = xtpEditStyleCenter
        .DropDownListStyle = True
        .ThemedItems = True
        .Width = 60
        For AktZa = 0 To 23
            .AddItem Format$(AktZa, "00") & ":00"
            .AddItem Format$(AktZa, "00") & ":15"
            .AddItem Format$(AktZa, "00") & ":30"
            .AddItem Format$(AktZa, "00") & ":45"
            .DropDownItemCount = 8
        Next AktZa
    End With

    Set CmCon = .Add(xtpControlLabel, KA_Capt1, "  Geldkonto :")
    With CmCon
        .ToolTipText = "Wählen Sie bitte hier die Kasse aus."
        .Style = xtpButtonIconAndCaption
    End With
    Set CmGeg = .Add(xtpControlComboBox, KA_SuCo1, vbNullString)
    With CmGeg
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ThemedItems = True
        .ToolTipText = "Wählen Sie bitte hier, welchen Eintrag Sie vornehmen möchten"
        .Width = 130
        .DropDownItemCount = UBound(GlGeK)
        If GlBuc = True Then 'einfache Buchhaltung verwenden
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                If CBool(GlGeK(AktZa, 5)) = True Then 'nur Kassen auflisten
                    .AddItem GlGeK(AktZa, 3)
                    .ItemData(AktPo) = GlGeK(AktZa, 0) '[IDB]
                    AktPo = AktPo + 1
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                If CBool(GlGeK(AktZa, 5)) = True Then 'nur Kassen auflisten
                    For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                        If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                            .AddItem GlSaK(AktKo, 3)
                            .ItemData(AktZa - 1) = GlSaK(AktKo, 6) '[IDB]
                            Exit For
                        End If
                    Next AktKo
                    AktPo = AktPo + 1
                End If
            Next AktZa
            If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
                For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                    If CBool(GlGeK(AktZa, 5)) = True Then 'nur Kassen auflisten
                        .AddItem GlGeK(AktZa, 3)
                        .ItemData(AktPo) = GlGeK(AktZa, 0)
                        AktPo = AktPo + 1
                    End If
                Next AktZa
            End If
        End If
    End With
End With

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
Private Sub FNeue()
On Error GoTo KoErr
'Lädt Details der Kassenzählungen

Dim AkZe1 As Date
Dim AkZe2 As Date
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpRws As XtremeReportControl.ReportRows
Dim RpCo2 As XtremeReportControl.ReportControl
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCoZ As XtremeCommandBars.CommandBarComboBox

Set FM = frmZaehlung
Set CmBrs = FM.comBar02
Set RpCo2 = FM.repCont2
Set RpRcs = RpCo2.Records
Set RpRws = RpCo2.Rows
Set CmSta = CmBrs.StatusBar

Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)
Set CmCoZ = CmBrs.FindControl(CmCoZ, KA_Uhrze, , True)

FoNeu = True

CmEdi.Text = Date

AkZe1 = TimeValue(Now)
For AktZa = 1 To CmCoZ.ListCount
    AkZe2 = CmCoZ.List(AktZa)
    If AkZe2 >= AkZe1 Then
        Exit For
    End If
Next AktZa
CmCoZ.ListIndex = AktZa
CmCoZ.Text = Format$(AkZe1, "hh:mm")
DoEvents

With RpCo2
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .AllowEdit = False Then .AllowEdit = True
    If .EditOnClick = False Then .EditOnClick = True
    .Populate
End With

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(0), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(1), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(2), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(3), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(4), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(5), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(6), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Münzen")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(7), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Scheine")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(8), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))
    
Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Scheine")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(9), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Scheine")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(10), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Scheine")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(11), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Scheine")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(12), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Scheine")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(13), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("Scheine")
Set RpItm = RpRec.AddItem(GlWar(GlStW, 1))
Set RpItm = RpRec.AddItem(Format$(GlGel(14), GlWa1))
Set RpItm = RpRec.AddItem("0")
Set RpItm = RpRec.AddItem(Format$(0, GlWa1))

For Each RpRec In RpRcs
    For Each RpItm In RpRec
        If RpItm.Index = 3 Then
            With RpItm
                If .Editable = False Then .Editable = True
                If .Focusable = False Then .Focusable = True
            End With
        End If
    Next RpItm
Next RpRec

With RpCo2
    .GroupsOrder.Add .Columns(0)
    .Populate
    .Navigator.MoveFirstVisibleRow
    .Navigator.MoveToColumn 3, True
    .SetFocus
End With

CmSta.Pane(1).Text = "Gesamt : " & Format$(0, GlWa1) & Space$(3)

Set RpRec = Nothing
Set RpRcs = Nothing
Set RpCo2 = Nothing

Exit Sub

KoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeue " & Err.Number
Resume Next

End Sub
Private Sub FOpen()
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim ManNr As Long
Dim StaGe As Long
Dim AkZe1 As Date
Dim AkZe2 As Date
Dim AktZa As Integer
Dim GesZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCoZ As XtremeCommandBars.CommandBarComboBox

Set FM = frmZaehlung
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)
Set CmCoZ = CmBrs.FindControl(CmCoZ, KA_Uhrze, , True)
Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

ManNr = GlMan(GlSMa, 2)

CmEdi.Text = Date

GesZa = CmCom.ListCount

AkZe1 = TimeValue(Now)
For AktZa = 1 To CmCoZ.ListCount
    AkZe2 = CmCoZ.List(AktZa)
    If AkZe2 >= AkZe1 Then
        Exit For
    End If
Next AktZa
CmCoZ.ListIndex = AktZa
CmCoZ.Text = Format$(AkZe1, "hh:mm")
DoEvents

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            If GlMan(AktZa, 28) <> vbNullString Then
                If GlMan(AktZa, 28) > 0 Then
                    StaGe = GlMan(AktZa, 29) 'Standardgeldkonto Kasse
                Else
                    StaGe = GlGkK
                End If
            Else
                StaGe = CInt(GlSet(2, 27))
            End If
            If StaGe = 0 Then StaGe = 1
            If StaGe > GesZa Then StaGe = 1
            CmCom.ListIndex = StaGe
            Exit For
        End If
    Next AktZa
Else
    CmCom.ListIndex = 1
End If

Set CmSta = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpen " & Err.Number
Resume Next

End Sub
Private Sub FPrint()
On Error GoTo PoErr
'Druckt das Zählprotokoll

Dim IdxNr As Long
Dim ForNa As String
Dim FiNam As String
Dim KopTe As String
Dim Formu As Boolean
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmZaehlung
Set CmBrs = FM.comBar02
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

Set clLis = New clsLisLab
Set clFil = New clsFile

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)

KopTe = CmCom.Text

If KopTe = vbNullString Then
    KopTe = "Kassenbuch"
End If

ForNa = "KasPro"

FiNam = GlFrO & S_FoCh(ForNa)

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        IdxNr = RpRow.Record(0).Value
    End If
End If

If clFil.FilVor(FiNam) = True Then
    Formu = True
Else
    Formu = False
    SMeFr GlMeT, GlMeM, GlMeI, GlMeF, False, 1, True, FM.hwnd
End If

Unload Me

If IdxNr > 0 Then
    If Formu = True Then
        ReDim Preserve GloDr(1)
        GloDr(1) = IdxNr
        With clLis
            .ForNam = ForNa
            .FilNam = FiNam
            .PfaTmp = GlTmp
            .DruVor = GlDrV
            .Gesamt = False
            .KopTex = KopTe
            .MandVo = True
            .MitaVo = GlMiV
            .ArztVo = GlArV
            .LLPrLi
        End With
    End If
End If

ReDim GloDr(0)

Set clLis = Nothing
Set clFil = Nothing

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FPrint " & Err.Number
Resume Next

End Sub

Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FSpla()
On Error GoTo InErr

Dim AktZa As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmZaehlung
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2

Set RpCls = RpCo1.Columns
With RpCls
    Set RpCol = .Add(0, "IDK", 0, False)
    With RpCol
    .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(1, "Datum", 100, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
        .AutoSize = True
    End With
    Set RpCol = .Add(2, "Uhrzeit", 70, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(3, "Gezählt", 90, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(4, "Gebucht", 90, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(5, "Different", 90, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
End With

Set RpCls = RpCo2.Columns
With RpCls
    Set RpCol = .Add(0, "Stückelung", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(1, "Währung", 100, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
        .AutoSize = True
    End With
    Set RpCol = .Add(2, "Wert", 90, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(3, "Anzahl", 90, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        .EditOptions.EditControlStyle = xtpEditStyleNumber
    End With
    Set RpCol = .Add(4, "Summe", 90, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSpla " & Err.Number
Resume Next

End Sub
Private Sub FTaEd(ByVal KeyCode As Integer, Optional ByVal Shift As Integer)
On Error Resume Next

Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmZaehlung
Set RpCo2 = FM.repCont2
Set RpSel = RpCo2.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If Shift = 0 Then
            Select Case KeyCode
            Case vbKeyF2:
                        RpCo2.Navigator.BeginEdit
            Case vbKeyTab:
            Case vbKeyReturn:
                        With RpCo2
                            .Navigator.MoveDown
                            .Navigator.MoveToColumn 3, True
                            .SetFocus
                        End With
            Case vbKeyDown:
            Case vbKeyUp:
            Case vbKeyPageDown:
            Case vbKeyPageUp:
            End Select
        End If
    End If
End If

Set RpSel = Nothing
Set RpCo2 = Nothing

End Sub
Private Sub FTool(ByVal TolId As Long)

Select Case TolId
Case KY_F1: FHilfe
Case KY_F3:
        FNeue
Case KY_F8:
        If FoNeu = True Then
            S_BuKaS
            S_BuKaL
            FoNeu = False
        End If
Case KY_F10:
        FPrint
Case KY_F11:
        Unload Me
Case SY_OP_Hinzufuegen:
        FNeue
Case SY_OP_Speichern:
        If FoNeu = True Then
            S_BuKaS
            S_BuKaL
            FoNeu = False
        End If
Case SY_OP_Loeschen:
        S_BuKaO
        S_BuKaL
Case SY_OP_Drucken:
        FPrint
Case SY_OP_Abbruch:
        Unload Me
Case KA_KaBu1:
        FKale
Case KA_SuCo1:
        S_BuKaL
End Select

End Sub

Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub

Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True
FInit
FMenu
AFont Me
FSpla
FOpen
S_BuKaL
FoLad = False
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmZaehlung = Nothing
End Sub

Private Sub repCont1_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo LaErr

Dim IdxNr As Long
Dim HiTes As XtremeReportControl.ReportHitTestInfo
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set RpCo1 = Me.repCont1
Set HiTes = RpCo1.HitTest(x, y)
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(0)
        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
    End If
End If

If RpSel.Count > 0 Then
    Select Case HiTes.ht
    Case xtpHitTestGroupBox:
    Case xtpHitTestHeader:
    Case xtpHitTestReportArea: S_BuKaD IdxNr
    Case xtpHitTestUnknown:
    End Select
End If

Set RpSel = Nothing
Set RpCo1 = Nothing
    
LaErr:
If GlDbg = True Then
    If Err.Number > 0 Then
        MsgBox Err.Description, 48, "Main " & Err.Number
    End If
End If
Exit Sub

End Sub
Private Sub repCont2_KeyUp(KeyCode As Integer, Shift As Integer)
    FTaEd KeyCode, Shift
End Sub
Private Sub repCont2_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    FKran Column.Index
End Sub
Private Sub comBar02_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    FTool Control.id
End Sub
