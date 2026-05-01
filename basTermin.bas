Attribute VB_Name = "basTermin"
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PiR01 As VB.PictureBox
Private PiR02 As VB.PictureBox
Private Lab17 As XtremeSuiteControls.Label
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private Rahm5 As XtremeSuiteControls.GroupBox
Private Rahm6 As XtremeSuiteControls.GroupBox
Private Rahm7 As XtremeSuiteControls.GroupBox
Private Rahm8 As XtremeSuiteControls.GroupBox
Private Rahm9 As XtremeSuiteControls.GroupBox
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private TxID0 As XtremeSuiteControls.FlatEdit
Private TxID2 As XtremeSuiteControls.FlatEdit
Private TxGui As XtremeSuiteControls.FlatEdit
Private TxFil As XtremeSuiteControls.FlatEdit
Private TxKom As XtremeSuiteControls.FlatEdit
Private TeDum As XtremeSuiteControls.FlatEdit
Private TxNeu As XtremeSuiteControls.FlatEdit
Private TxIdx As XtremeSuiteControls.FlatEdit
Private TxFar As XtremeSuiteControls.FlatEdit
Private TxOrt As XtremeSuiteControls.FlatEdit
Private TxAdr As XtremeSuiteControls.FlatEdit
Private TxIDS As XtremeSuiteControls.FlatEdit
Private VoZei As XtremeSuiteControls.FlatEdit
Private BiZei As XtremeSuiteControls.FlatEdit
Private TxDa4 As XtremeSuiteControls.FlatEdit
Private TxDa5 As XtremeSuiteControls.FlatEdit
Private ZyTag As XtremeSuiteControls.FlatEdit
Private TxAnz As XtremeSuiteControls.FlatEdit
Private TxMul As XtremeSuiteControls.FlatEdit
Private TxEin As XtremeSuiteControls.FlatEdit
Private TxRef As XtremeSuiteControls.FlatEdit
Private TxRzn As XtremeSuiteControls.FlatEdit
Private TxRzA As XtremeSuiteControls.FlatEdit
Private TxZeV As XtremeSuiteControls.FlatEdit
Private TxZeB As XtremeSuiteControls.FlatEdit
Private TxNoS As XtremeSuiteControls.FlatEdit
Private TxNoD As XtremeSuiteControls.FlatEdit
Private TxNoZ As XtremeSuiteControls.FlatEdit
Private CmETy As XtremeSuiteControls.ComboBox
Private CmZif As XtremeSuiteControls.ComboBox
Private CmBez As XtremeSuiteControls.ComboBox
Private CmMar As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmArz As XtremeSuiteControls.ComboBox
Private CmPri As XtremeSuiteControls.ComboBox
Private CmRem As XtremeSuiteControls.ComboBox
Private CmGes As XtremeSuiteControls.ComboBox
Private ZwZei As XtremeSuiteControls.ComboBox
Private ZyWoh As XtremeSuiteControls.ComboBox
Private ZyMo1 As XtremeSuiteControls.ComboBox
Private ZyMo2 As XtremeSuiteControls.ComboBox
Private ZyMo3 As XtremeSuiteControls.ComboBox
Private ZyMo4 As XtremeSuiteControls.ComboBox
Private ZyMoT As XtremeSuiteControls.ComboBox
Private ZyJa1 As XtremeSuiteControls.ComboBox
Private ZyJa2 As XtremeSuiteControls.ComboBox
Private ZyJa3 As XtremeSuiteControls.ComboBox
Private ZyJa4 As XtremeSuiteControls.ComboBox
Private ZyJaT As XtremeSuiteControls.ComboBox
Private ZyEnT As XtremeSuiteControls.ComboBox
Private ZyWho As XtremeSuiteControls.ComboBox
Private ZyMe1 As XtremeSuiteControls.ComboBox
Private ZyMe2 As XtremeSuiteControls.ComboBox
Private ZyMe3 As XtremeSuiteControls.ComboBox
Private ZyJe1 As XtremeSuiteControls.ComboBox
Private ZyTer As XtremeSuiteControls.ComboBox
Private CmBet As XtremeSuiteControls.ComboBox
Private TxSp1 As XtremeSuiteControls.ComboBox
Private TxSp2 As XtremeSuiteControls.ComboBox
Private CmBrf As XtremeSuiteControls.ComboBox
Private CmNot As XtremeSuiteControls.ComboBox
Private CmGan As XtremeSuiteControls.ComboBox
Private CmSpe As XtremeSuiteControls.ComboBox
Private CmAbg As XtremeSuiteControls.ComboBox
Private CmAbr As XtremeSuiteControls.ComboBox
Private CmRmu As XtremeSuiteControls.ComboBox
Private CmOnT As XtremeSuiteControls.ComboBox
Private ChTer As XtremeSuiteControls.CheckBox
Private ChSpl As XtremeSuiteControls.CheckBox
Private ChRau As XtremeSuiteControls.CheckBox
Private ChSpr As XtremeSuiteControls.CheckBox
Private ChMon As XtremeSuiteControls.CheckBox
Private ChDin As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private ChDon As XtremeSuiteControls.CheckBox
Private ChFre As XtremeSuiteControls.CheckBox
Private ChSam As XtremeSuiteControls.CheckBox
Private ChSon As XtremeSuiteControls.CheckBox
Private ChDop As XtremeSuiteControls.CheckBox
Private ChAgh As XtremeSuiteControls.CheckBox
Private FoZy1 As XtremeSuiteControls.RadioButton
Private FoZy2 As XtremeSuiteControls.RadioButton
Private FoZy3 As XtremeSuiteControls.RadioButton
Private FoZy4 As XtremeSuiteControls.RadioButton
Private ZyEn2 As XtremeSuiteControls.RadioButton
Private ZyEn3 As XtremeSuiteControls.RadioButton
Private TaZy1 As XtremeSuiteControls.RadioButton
Private TaZy2 As XtremeSuiteControls.RadioButton
Private MoZy1 As XtremeSuiteControls.RadioButton
Private MoZy2 As XtremeSuiteControls.RadioButton
Private JaZy1 As XtremeSuiteControls.RadioButton
Private JaZy2 As XtremeSuiteControls.RadioButton
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private PuBu4 As XtremeSuiteControls.PushButton
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private CmSta As XtremeCommandBars.StatusBar
Private TbBar As XtremeCommandBars.TabToolBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmBuT As XtremeCommandBars.CommandBarButton
Private CmAct As XtremeCommandBars.CommandBarAction
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions
Private TabCo As XtremeSuiteControls.TabControl
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
Private MoKa1 As XtremeCalendarControl.DatePicker
Private MoKa2 As XtremeCalendarControl.DatePicker
Private MoKa3 As XtremeCalendarControl.DatePicker
Private CaCol As XtremeCalendarControl.CalendarControl
Private ChCon As XtremeChartControl.ChartControl
Private TxCoN As Tx4oleLib.TXTextControl

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private Const olContactItem = 2
Private Const olAppointmentItem = 1
Private Const olFolderCalendar = 9
Private Const olFolderContacts = 10

Private clFen As clsFenster
Private clFil As clsFile
Private clICS As clsICS

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Sub KoAus()
On Error GoTo PoErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmKomment
Set CmBrs = FM.comBar02
Set TxKom = FM.txtKomme

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    TxKom.Move ClLin, ClObn, ClBre, ClHoh
End If

Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " KoAus " & Err.Number
Resume Next

End Sub
Private Sub KoInit()
On Error GoTo InErr

Dim AktZa As Integer
Dim AktPo As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmMta As XtremeCommandBars.CommandBarComboBox

Set FM = frmKomment
Set CmBrs = FM.comBar02
Set TxKom = FM.txtKomme
Set MoKa1 = FM.dtpDatu1
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag

AktPo = 1

With TxKom
    .Font.SIZE = GlTFt.SIZE + 1
    .Font.Name = GlTFt.Name
End With

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
    CmPan.Width = 200
    .Visible = True
End With

With CmAcs
    Set CmAct = .Add(KA_SuCo1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_SuCo3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_KaBu1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(SY_OP_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
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
        .ToolTipText = "Speichert den eintrag ins Krankenblatt"
        .ShortcutText = "F8"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Suchen, "Suchen")
    With CmCon
        .ToolTipText = "Sucht einen Eintrag im Katalog"
        .ShortcutText = "F5"
        .IconId = IC24_View
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Hilfe, "Hilfe")
    With CmCon
        .ToolTipText = "Öffnet die Kurzhilfe"
        .ShortcutText = "F1"
        .BeginGroup = True
        .IconId = IC24_Help
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Drucken, "Drucken")
    With CmCon
        .ShortcutText = "F10"
        .BeginGroup = True
        .IconId = IC24_Printer_Ink
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Abbruch, "Abbrechen")
    With CmCon
        .ShortcutText = "F11"
        .BeginGroup = True
        .IconId = IC24_Exit
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
    Set CmCon = .Add(xtpControlLabel, KA_Capt1, "Eintragstyp :")
    With CmCon
        .ToolTipText = "Wählen Sie bitte hier, welchen Eintrag Sie vornehmen möchten"
        .Style = xtpButtonIconAndCaption
    End With
    Set CmCom = .Add(xtpControlComboBox, KA_SuCo1, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ThemedItems = True
        .ToolTipText = "Wählen Sie bitte hier, welchen Eintrag Sie vornehmen möchten"
        .Width = 130
        .DropDownItemCount = UBound(GlKrA)
        For AktZa = 1 To UBound(GlKrA) 'Krankenblatttypen
            If GlKrA(AktZa, 0) > 9 Then
                .AddItem GlKrA(AktZa, 1) & " - " & GlKrA(AktZa, 2)
                .ItemData(AktPo) = GlKrA(AktZa, 0)
                AktPo = AktPo + 1
            End If
        Next AktZa
    End With
    Set CmCon = .Add(xtpControlLabel, KA_Place, Space$(2))
    Set CmEdi = .Add(xtpControlEdit, KA_Kalen, " Datum :")
    With CmEdi
        .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag gespeichetr werden soll"
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
    Set CmCon = .Add(xtpControlLabel, KA_Place, Space$(2))
    Set CmCon = .Add(xtpControlLabel, KA_Capt3, "Mitarbeiter :")
    With CmCon
        .ToolTipText = "Unter welchem Mitarbeiter soll der Eintrag gespeichert werden?"
        .Style = xtpButtonIconAndCaption
    End With
    Set CmMta = .Add(xtpControlComboBox, KA_SuCo3, vbNullString)
    With CmMta
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Unter welchem Mitarbeiter soll der Eintrag gespeichert werden?"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 130
        If GlMiV = True Then
            For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
                .AddItem GlMiA(AktZa, 3) & ", " & GlMiA(AktZa, 4)
                .ItemData(AktZa) = GlMiA(AktZa, 2)
            Next AktZa
        End If
    End With
End With

Set CmBar = CmBrs.Add("ID_Auswahl", xtpBarTop)
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
    .Visible = False
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlLabel, KA_Capt2, "Suche :")
    With CmCon
        .ToolTipText = "Geben Sie bitte hier den gewünschten Suchbegriff ein"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With
    Set CmCom = .Add(xtpControlComboBox, KA_SuCo2, vbNullString)
    With CmCom
        .AutoComplete = True
        .EditHint = "Hier klicken und mit ENTER bestätigen"
        .EditStyle = xtpEditStyleLeft
        .ThemedItems = True
        .CloseSubMenuOnClick = True
        .DropDownListStyle = True
        .ToolTipText = "Geben Sie bitte hier den gewünschten Suchbegriff ein"
        .Width = 272
    End With
End With

With MoKa1
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

Set CmMta = CmBrs.FindControl(CmMta, KA_SuCo3, , True)

Set MoKa1 = Nothing
Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " KoInit " & Err.Number
Resume Next

End Sub
Public Sub KoMain(Optional ByVal DaNam As String, Optional DaNaO As String, Optional ByVal EiTyp As Integer)
On Error GoTo MeErr

Dim CmBrs As XtremeCommandBars.CommandBars

GlAkK = True

If WindowLoad("frmKomment") = True Then
    Set FM = frmKomment
    frmKomment.ZOrder 0
    Exit Sub
End If

KoReg
DoEvents

Load frmKomment

Set FM = frmKomment
Set TxKom = FM.txtKomme
Set TxFil = FM.txtFiNam
Set CmBrs = frmMain.comBar01
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    .FeLin = IniGetVal("Kommentar", "FenLin")
    .FeObn = IniGetVal("Kommentar", "FenObe")
    .FeBre = IniGetVal("Kommentar", "FenBre")
    .FeHoh = IniGetVal("Kommentar", "FenHoh")
End With

CmAcs(SY_KB_KraBla_Hinzufueg).Enabled = False
CmAcs(SY_KB_KraBla_Loeschen).Enabled = False

If DaNam <> vbNullString Then
    TxFil.Text = DaNam 'WICHTIG vor KoOpn
End If

KoInit
KoOpn

With clFen
    .FenMov
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

frmKomment.Show

If DaNaO <> vbNullString Then
    TxKom.Text = DaNaO
End If

TxKom.SetFocus

GlAkK = False

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " KoMain " & Err.Number
Resume Next

End Sub
Private Sub KoOpn()
On Error GoTo PoErr
'Öffnet das Kommentarfeld

Dim NeuDa As Date
Dim TxFar As Long
Dim LiTyp As Long
Dim IdxNr As Long
Dim MitNr As Long
Dim DaNaO As String
Dim KoStr As String
Dim AktZa As Integer
Dim AktPo As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmMta As XtremeCommandBars.CommandBarComboBox
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmKomment
Set MoKa1 = FM.dtpDatu1
Set TxKom = FM.txtKomme
Set TxFil = FM.txtFiNam
Set TeDum = FM.txtDummy
Set TxNeu = FM.txtNeuEi
Set TxIdx = FM.txtIdxNr
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set RpCo6 = frmMain.repCont6
Set RpCoK = frmMain.repContK

Select Case GlBut
Case RibTab_Krankenbla:
        Set RpSel = RpCoK.SelectedRows
        Set RpCls = RpCoK.Columns
Case RibTab_Abrechnung:
        Set RpSel = RpCo6.SelectedRows
        Set RpCls = RpCo6.Columns
End Select

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)
Set CmMta = CmBrs.FindControl(CmMta, KA_SuCo3, , True)

Select Case GlBut
Case RibTab_Krankenbla:

    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Kra_Typ)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                LiTyp = RpRow.Record(RpCol.ItemIndex).Value
            End If
            Set RpCol = RpCls.Find(Kra_Datum)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                NeuDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                NeuDa = Date
            End If
            Set RpCol = RpCls.Find(Kra_ID2)
            TxIdx.Text = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Kra_IDM)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                    MitNr = RpRow.Record(RpCol.ItemIndex).Value
                Else
                    MitNr = GlMiA(GlSmI, 2)
                End If
            Else
                MitNr = GlMiA(GlSmI, 2)
            End If
            If GlMiV = True Then
                For AktZa = 1 To UBound(GlMiA)
                    If MitNr = GlMiA(AktZa, 2) Then
                        CmMta.ListIndex = AktZa
                        Exit For
                    End If
                Next AktZa
            End If
            Set RpCol = RpCls.Find(Kra_Zusatztext)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                TeDum.Text = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
            End If
            
            Select Case GlBut
            Case RibTab_Krankenbla:
                Set RpCol = RpCls.Find(Kra_Bezeichnung)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    If TxFil.Text <> vbNullString Then
                        Set RpCol = RpCls.Find(Kra_Kommentar)
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            DaNaO = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                        Else
                            DaNaO = vbNullString
                        End If
                    Else
                        DaNaO = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                    End If
                Else
                    Set RpCol = RpCls.Find(Kra_Kommentar)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        DaNaO = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                    Else
                        DaNaO = vbNullString
                    End If
                End If
            Case RibTab_Abrechnung:
                Set RpCol = RpCls.Find(Kra_Kommentar)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    DaNaO = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                Else
                    DaNaO = vbNullString
                End If
            End Select
            
            CmEdi.Text = NeuDa
            TxKom.Text = DaNaO

            For AktZa = 1 To UBound(GlKrA) 'Krankenblatttypen
                If LiTyp = GlKrA(AktZa, 0) Then
                    CmCom.ListIndex = AktZa - 9
                    TxKom.ForeColor = GlKrA(AktZa - 9, 3)
                    Exit For
                End If
            Next AktZa
        End If
    End If

    Select Case CmCom.ListIndex
    Case 1: CmAcs(SY_OP_Suchen).Enabled = True
    Case 13: CmAcs(SY_OP_Suchen).Enabled = True
    Case Else: CmAcs(SY_OP_Suchen).Enabled = False
    End Select
    
Case RibTab_Abrechnung:

    CmCom.Enabled = False
    CmEdi.Enabled = False
    CmAcs(KA_KaBu1).Enabled = False
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Kra_Datum)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                NeuDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                NeuDa = Date
            End If
            Set RpCol = RpCls.Find(Kra_ID2)
            IdxNr = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Kra_IDM)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                    MitNr = RpRow.Record(RpCol.ItemIndex).Value
                Else
                    MitNr = GlMiA(GlSmI, 2)
                End If
            Else
                MitNr = GlMiA(GlSmI, 2)
            End If
            Set RpCol = RpCls.Find(Kra_Zusatztext)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                TeDum.Text = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
            End If
            
            Set RpCol = RpCls.Find(Kra_Kommentar)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                DaNaO = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                DaNaO = vbNullString
            End If

            TxIdx.Text = IdxNr
            CmEdi.Text = NeuDa
            TxKom.Text = DaNaO
                        
            If GlMiV = True Then
                For AktZa = 1 To UBound(GlMiA)
                    If MitNr = GlMiA(AktZa, 2) Then
                        CmMta.ListIndex = AktZa
                        Exit For
                    End If
                Next AktZa
            End If
        End If
    End If
    CmAcs(SY_OP_Suchen).Enabled = False
End Select

Set MoKa1 = Nothing
Set CmBrs = Nothing
Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo6 = Nothing
Set RpCoK = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " KoOpn " & Err.Number
Resume Next

End Sub
Private Sub KoReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "Kommentar") = False Then
    xGro = 600
    yGro = 400
    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)
    
    IniSetSek "Kommentar"
    IniSetVal "Kommentar", "FenLin", xPos
    IniSetVal "Kommentar", "FenObe", yPos
    IniSetVal "Kommentar", "FenBre", xGro
    IniSetVal "Kommentar", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " KoReg " & Err.Number
Resume Next

End Sub
Public Sub KoSav()
On Error GoTo PoErr

Dim IdxNr As Long
Dim KoMit As Long
Dim NeuDa As Date
Dim KoGui As String
Dim KoStr As String
Dim FiNam As String
Dim EiTyp As Integer
Dim NeuEi As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmMta As XtremeCommandBars.CommandBarComboBox

Set FM = frmKomment
Set MoKa1 = FM.dtpDatu1
Set TxKom = FM.txtKomme
Set TxFil = FM.txtFiNam
Set TeDum = FM.txtDummy
Set TxNeu = FM.txtNeuEi
Set TxIdx = FM.txtIdxNr
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmEdi = CmBrs.FindControl(CmEdi, KA_Kalen, , True)
Set CmMta = CmBrs.FindControl(CmMta, KA_SuCo3, , True)

If IsDate(CmEdi.Text) = True Then
    NeuDa = CmEdi.Text
Else
    NeuDa = Date
End If

If GlMiV = True Then
    KoMit = CmMta.ItemData(CmMta.ListIndex)
Else
    KoMit = 0
End If

If CmCom.ItemData(CmCom.ListIndex) > 0 Then
    EiTyp = CmCom.ItemData(CmCom.ListIndex)
Else
    EiTyp = 102
End If

If TxKom.Text <> vbNullString Then
    KoStr = TxKom.Text
Else
    KoStr = vbNullString
End If

If TxFil.Text <> vbNullString Then
    FiNam = TxFil.Text
Else
    FiNam = vbNullString
End If

If TeDum.Text <> vbNullString Then
    KoGui = TeDum.Text
Else
    KoGui = vbNullString
End If

If TxIdx.Text <> vbNullString Then
    IdxNr = TxIdx.Text
Else
    IdxNr = 0
End If

If TxNeu.Text <> vbNullString Then
    NeuEi = True
Else
    NeuEi = False
End If

GlNeK = GlKoX
If GlAdr > 0 Then
    If FiNam <> vbNullString Then
        With GlNeK
            .PatNr = GlAdr
            .IdxNr = IdxNr
            .EiDat = NeuDa
            .EiZei = TimeValue(Now)
            .EiTyp = EiTyp
            .KoStr = FiNam
            .KoGui = KoGui
            .NeuEi = NeuEi
            .TeStr = KoStr
            .Mitar = KoMit
        End With
    Else
        With GlNeK
            .PatNr = GlAdr
            .IdxNr = IdxNr
            .EiDat = NeuDa
            .EiZei = TimeValue(Now)
            .EiTyp = EiTyp
            .KoGui = KoGui
            .NeuEi = NeuEi
            .Mitar = KoMit
            If EiTyp = 106 Then 'Rechnung
                .TeStr = KoStr
            Else
                .KoStr = KoStr
            End If
        End With
    End If
    K_Einf
    If FiNam <> vbNullString Then
        TxNeu.Text = vbNullString
    End If
End If

Set CmBrs = Nothing
Set MoKa1 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " KoSav " & Err.Number
Resume Next

End Sub
Private Sub KrInit()
On Error GoTo InErr

Set FM = frmKraEd
Set MoKa1 = FM.dtpDatu1
Set TxKom = FM.txtKomme

With TxKom
    .Font.SIZE = GlTFt.SIZE + 1
    .Font.Name = GlTFt.Name
    .MaxLength = 8000
End With

With MoKa1
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

Set MoKa1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " KrInit " & Err.Number
Resume Next

End Sub
Public Sub KrMain(Optional ByVal EiTyp As Integer, Optional ByVal EiKop As Boolean = False)
On Error GoTo MeErr

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCom As XtremeCommandBars.CommandBarComboBox

GlAkK = True

If WindowLoad("frmKraEd") = True Then
    Set FM = frmKraEd
    frmKraEd.ZOrder 0
    Exit Sub
End If

KrReg

Load frmKraEd

Set FM = frmKraEd
Set TxKom = FM.txtKomme
Set CmBrs = frmMain.comBar01
Set CmAcs = CmBrs.Actions

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (865 / 2)
        .FeObn = (GlyGr / 2) - (465 / 2)
        .FeBre = 865
        .FeHoh = 465
    Else
        .FeLin = IniGetVal("Krankenblatt", "FenLin")
        .FeObn = IniGetVal("Krankenblatt", "FenObe")
        .FeBre = IniGetVal("Krankenblatt", "FenBre")
        .FeHoh = IniGetVal("Krankenblatt", "FenHoh")
    End If
End With

CmAcs(SY_KB_KraBla_Hinzufueg).Enabled = False
CmAcs(SY_KB_KraBla_Loeschen).Enabled = False

If EiTyp < 0 Then
    EiTyp = GlKrA(GlEiT, 0)
End If

AFont FM
DoEvents
KrInit
KrMen EiTyp
KrOpn EiTyp, EiKop
DoEvents
K_KrEd 1, 3

With clFen
    .FenMov
    Set CmBrs = FM.comBar02
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

frmKraEd.Show

If GlSPh = True Then 'Suchleiste Textphrase
    Set CmBrs = FM.comBar02
    Set CmCom = CmBrs.FindControl(CmCom, SY_SuCm1, , True)
    With CmCom
        If .Enabled = True Then
            .SetFocus
            .Execute
        End If
    End With
Else
    If TxKom.Enabled = True Then
        TxKom.SetFocus
    End If
End If

GlAkK = False

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " KrMain " & Err.Number
Resume Next

End Sub
Private Sub KrMen(Optional ByVal EiTyp As Integer)
On Error GoTo InErr

Dim RetWe As Long
Dim KeyNa As String
Dim AktZa As Integer
Dim AktPo As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdD As XtremeCommandBars.CommandBarEdit
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmMta As XtremeCommandBars.CommandBarComboBox
Dim CmCoZ As XtremeCommandBars.CommandBarComboBox
Dim CbTyp As XtremeCommandBars.CommandBarComboBox
Dim CmBap As XtremeCommandBars.CommandBar
Dim MsBar As XtremeCommandBars.MessageBar
Dim ImMan As XtremeCommandBars.ImageManager
Dim GalDa As XtremeCommandBars.CommandBarGallery
Dim GalFo As XtremeCommandBars.CommandBarGallery
Dim GalGr As XtremeCommandBars.CommandBarGallery
Dim GaItF As XtremeCommandBars.CommandBarGalleryItems
Dim GaItS As XtremeCommandBars.CommandBarGalleryItems
Dim GaItm As XtremeCommandBars.CommandBarGalleryItem
Dim CmCSe As XtremeCommandBars.CommandBarControlColorSelector
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKraEd
Set CmBrs = FM.comBar02
Set ImMan = frmMain.imgManag
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set MsBar = CmBrs.MessageBar
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

AktPo = 1
KeyNa = "ToolTips"

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Width = 100
    CmPan.Style = SBPS_STRETCH
    CmPan.Text = vbNullString
    .Visible = True
End With

With CmAcs
    Set CmAct = .Add(Tex_ForFet, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForKur, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForUnt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForDur, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrLi, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrRe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrZe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FntAu6, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FntGr6, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaVor2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaVor3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaHin1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaHin2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_DatSpe, vbNullString, vbNullString, vbNullString, vbNullString)
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")

Set GaItF = CmBrs.CreateGalleryItems(Tex_FontAu)
With GaItF
    .ItemWidth = 0
    .ItemHeight = 20
    .AddLabel "Häufig verwend. Schriften"
    .AddItem 1, "Arial"
    .AddItem 2, "Tahoma"
    .AddItem 3, "Consolas"
    .AddItem 4, "Freestyle Script"
    .AddLabel "Alle Schriftarten"
    For AktZa = 1 To UBound(FnAry)
        Set GaItm = .AddItem(AktZa, FnAry(AktZa))
    Next AktZa
End With

Set GaItS = CmBrs.CreateGalleryItems(Tex_FontGr)
With GaItS
    .ItemWidth = 0
    .ItemHeight = 18
    .AddItem 6, "6"
    .AddItem 8, "8"
    .AddItem 9, "9"
    .AddItem 10, "10"
    .AddItem 11, "11"
    .AddItem 12, "12"
    .AddItem 14, "14"
    .AddItem 16, "16"
    .AddItem 18, "18"
    .AddItem 20, "20"
    .AddItem 22, "22"
    .AddItem 24, "24"
    .AddItem 26, "26"
    .AddItem 28, "28"
    .AddItem 36, "36"
    .AddItem 48, "48"
    .AddItem 72, "72"
End With

Set CmCon = RbBar.Controls.Add(xtpControlLabel, SY_Cap02, " Datum :")
With CmCon
    .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag gespeichetr werden soll"
    .Style = xtpButtonCaption
End With
Set CmEdD = RbBar.Controls.Add(xtpControlEdit, KA_Kalen, vbNullString)
With CmEdD
    .ToolTipText = "Wählen Sie hier das Datum aus, unter dem der Eintrag gespeichetr werden soll"
    .Style = xtpButtonCaption
    .IconId = IC16_Calendar_Year
    .EditStyle = xtpEditStyleCenter
    .Width = 80
End With
Set CmCon = RbBar.Controls.Add(xtpControlButton, KA_KaBu1, vbNullString)
With CmCon
    .ToolTipText = "Klicken Sie hier, um den Kalender anzuzeigen"
    .Style = xtpButtonIcon
    .IconId = IC16_Calendar_Year
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, KA_Capt3, " Uhrzeit :")
With CmCon
    .ToolTipText = "Wählen Sie hier die Uhrzeit aus, unter dem der Eintrag gespeichetr werden soll"
    .Style = xtpButtonCaption
End With
Set CmCoZ = RbBar.Controls.Add(xtpControlComboBox, KA_Uhrze, vbNullString)
With CmCoZ
    .ToolTipText = "Wählen Sie hier die Uhrzeit aus, unter dem der Eintrag gespeichetr werden soll"
    .Style = xtpButtonCaption
    .IconId = IC16_Key_Kopf
    .EditStyle = xtpEditStyleCenter
    .DropDownListStyle = True
    .ThemedItems = True
    .Width = 70
    For AktZa = 0 To 23
        .AddItem Format$(AktZa, "00") & ":00"
        .AddItem Format$(AktZa, "00") & ":15"
        .AddItem Format$(AktZa, "00") & ":30"
        .AddItem Format$(AktZa, "00") & ":45"
        .DropDownItemCount = 8
    Next AktZa
End With

Set CmCon = RbBar.Controls.Add(xtpControlButton, Tex_EdUndo, "Rückgängig")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Undo
    .flags = xtpFlagRightAlign
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, KA_Hilfe, "Hilfe")
With CmBuT
    .ToolTipText = "Öffnet die Kurzhilfe"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Beenden, "Schließen")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .ToolTipText = "Schließt den Dialog"
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

'-----------------------------------------------------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Tex_Dokumt, "Texteintrag")
With RbTab
    .id = RibTab_Tex_Dokumt
    .ToolTip = IniGetOpt(KeyNa, RbTab.id)
    .Visible = False
    .Selected = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Auswahl", RibGrp_Tex_Patient)
RbGrp.ControlsGrouping = True

Set CmCon = RbGrp.Add(xtpControlLabel, KA_Capt1, vbNullString)
With CmCon
    .ToolTipText = "Wählen Sie bitte hier, welchen Eintrag Sie vornehmen möchten"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Doc_View
End With
Set CbTyp = RbGrp.Add(xtpControlComboBox, KA_SuCo1, vbNullString)
With CbTyp
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .ThemedItems = True
    .ToolTipText = "Wählen Sie bitte hier, welchen Eintrag Sie vornehmen möchten"
    .Width = 135
    .DropDownItemCount = UBound(GlKrA)
    For AktZa = 1 To UBound(GlKrA)
        If GlKrA(AktZa, 0) > 9 Then
            Select Case GlKrA(AktZa, 0)
            Case 24:    'Textdokumente
            Case 101:   'Beleg / Rezept
            Case 102:   'Datei
            Case 104:   'Protokoll
            Case 105:   'Bilddatei
            Case Else:
                .AddItem GlKrA(AktZa, 1) & " - " & GlKrA(AktZa, 2)
                .ItemData(AktPo) = GlKrA(AktZa, 0)
                AktPo = AktPo + 1
            End Select
        End If
    Next AktZa
End With
Set CmCon = RbGrp.Add(xtpControlLabel, KA_Capt2, vbNullString)
With CmCon
    .BeginGroup = True
    .ToolTipText = "Unter welchem Mitarbeiter soll der Eintrag gespeichert werden?"
    .flags = xtpFlagRightAlign
    .IconId = IC16_IDCard_Norm
End With

Set CmMta = RbGrp.Add(xtpControlComboBox, KA_SuCo3, vbNullString)
With CmMta
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .ToolTipText = "Unter welchem Mitarbeiter soll der Eintrag gespeichert werden?"
    .Style = xtpButtonAutomatic
    .ThemedItems = True
    .Width = 135
    If GlMiV = True Then
        For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
            .AddItem GlMiA(AktZa, 3) & ", " & GlMiA(AktZa, 4)
            .ItemData(AktZa) = GlMiA(AktZa, 2)
        Next AktZa
    End If
End With

'-----

Set RbGrp = RbGps.AddGroup("Eintrag", RibGrp_Tex_Dokument)

Set CmCon = RbGrp.Add(xtpControlButton, Tex_Suchen, "Textphrase Suchen")
With CmCon
    .IconId = IC32_Doc_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_DatSpe, "Eintrag Speichern")
With CmCon
    .IconId = IC32_Disk_Document
    .ShortcutText = "F8"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_TexCut, "Ausschneiden")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Cut
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_TexCop, "Kopieren")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Copy
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_TexEin, "Einfügen")
With CmCon
    .IconId = IC16_Paste
    .Width = GlRib
End With

'-----

Set RbGrp = RbGps.AddGroup("Schriftart", RibGrp_Tex_Schrift)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, Tex_FntAu6, vbNullString)
With CmCom
    .DropDownListStyle = True
    .ThemedItems = True
    .Width = 174
    '.Text = KrFnt.Name
    .AutoComplete = True
    .KeyboardTip = "FF"
End With
Set CmBap = CmBrs.Add("ComBar", xtpBarComboBoxGalleryPopup)
Set GalFo = CmBap.Controls.Add(xtpControlGallery, Tex_FontAu, vbNullString)
With GalFo
    .Width = 190
    .Height = 300
    .Resizable = xtpAllowResizeWidth Or xtpAllowResizeHeight
End With
Set GalFo.Items = GaItF
Set CmCom.CommandBar = CmBap
Set CmCom = RbGrp.Add(xtpControlComboBox, Tex_FntGr6, vbNullString)
With CmCom
    .DropDownListStyle = True
    .ThemedItems = True
    .Width = 50
    '.Text = KrFnt.SIZE
    .KeyboardTip = "FS"
End With
Set CmBap = CmBrs.Add("ComBar", xtpBarComboBoxGalleryPopup)
Set GalGr = CmBap.Controls.Add(xtpControlGallery, Tex_FontGr, vbNullString)
With GalGr
    .Width = 50
    .Height = 170
    .Resizable = xtpAllowResizeHeight
    .BeginGroup = True
End With
Set GalGr.Items = GaItS
Set CmCom.CommandBar = CmBap
Set CmCon = RbGrp.Add(xtpControlButton, Tex_AusrLi, "Linksbündig")
With CmCon
    .ToolTipText = "Richtet den markierten Text linksbündig aus"
    .IconId = IC16_Links
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_AusrZe, "Zentriert")
With CmCon
    .ToolTipText = "Richtet den markierten Text zentriert aus"
    .IconId = IC16_Zentr
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_AusrRe, "Rechtsbündig")
With CmCon
    .ToolTipText = "Richtet den markierten Text rechtsbündig aus"
    .IconId = IC16_Rechts
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_ForFet, "Fettdruck")
With CmCon
    .ToolTipText = "Fettdruck"
    .IconId = IC16_Fett
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_ForKur, "Kursiv")
With CmCon
    .ToolTipText = "Kursiv"
    .IconId = IC16_Kursiv
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_ForUnt, "Unterstrichen")
With CmCon
    .ToolTipText = "Unterstrichen"
    .IconId = IC16_Unter
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_ForDur, "Durchgestrichen")
With CmCon
    .ToolTipText = "Durchgestrichen"
    .IconId = IC16_Strike
End With
Set CmCop = CmBrs.CreateCommandBarControl("CXTPControlPopupColor")
With CmCop
    .id = IC16_FarVor
    .Caption = "Fordergrundfarbe"
    .ToolTipText = "Fordergrundfarbe"
    .Style = xtpButtonIcon
    .BeginGroup = True
End With
RbGrp.AddControl CmCop
Set CmBap = CmBrs.Add("ComBar", xtpBarPopup)
CmBap.SetPopupToolBar True
Set CmCop.CommandBar = CmBap
Set CmCSe = CmBrs.CreateCommandBarControl("CXTPControlColorSelector")
Set CmCon = CmBap.Controls.Add(xtpControlButton, Tex_FaVor2, "Standard")
CmCon.Width = 148
With CmCSe
    .id = Tex_FaVor1
    .Caption = "Fordergrundfarbe"
End With
CmBap.Controls.AddControl CmCSe
Set CmCop = CmBrs.CreateCommandBarControl("CXTPControlPopupColor")
With CmCop
    .id = IC16_FarHin
    .Caption = "Hintergrundfarbe"
    .ToolTipText = "Hintergrundfarbe"
    .Style = xtpButtonIcon
End With
RbGrp.AddControl CmCop
Set CmBap = CmBrs.Add("ComBar", xtpBarPopup)
CmBap.SetPopupToolBar True
Set CmCop.CommandBar = CmBap
Set CmCSe = CmBrs.CreateCommandBarControl("CXTPControlColorSelector")
Set CmCon = CmBap.Controls.Add(xtpControlButton, Tex_FaHin2, "Standard")
CmCon.Width = 148
With CmCSe
    .id = Tex_FaHin1
    .Caption = "Hintergrundfarbe"
End With
CmBap.Controls.AddControl CmCSe

'-----

Set RbGrp = RbGps.AddGroup("Ausgabe", RibGrp_Tex_Drucken)
Set CmCon = RbGrp.Add(xtpControlButton, Tex_DatSpV, "Eintrag Exportieren")
With CmCon
    .IconId = IC32_Doc_Export
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_DocDru, "Eintrag Drucken")
With CmCon
    .IconId = IC32_Printer_Ink
    .ShortcutText = "F10"
    .Width = GlRib
    .Enabled = Not GlKrF 'Krankenblattdialog im Vordergund
End With

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
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(23))
    
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Textsuche :")
    With CmCon
        .ToolTipText = "Geben Sie bitte hier Ihre Suchanfrage ein"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_View
    End With

    Set CmCom = .Add(xtpControlComboBox, SY_SuCm1, vbNullString)
    With CmCom
        .DropDownListStyle = True
        .AutoComplete = True
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Textkürzel..."
        .ToolTipText = "Geben Sie bitte hier das Suchkriterium ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 120
    End With
    
    Set CmCon = .Add(xtpControlLabel, SY_Plac2, Space$(1))
    
    Set CmCom = .Add(xtpControlComboBox, SY_SuCm2, vbNullString)
    With CmCom
        .DropDownListStyle = True
        .AutoComplete = True
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Eingabe Suchkriterium..."
        .ToolTipText = "Geben Sie bitte hier das Suchkriterium ein und bestätigen mit der ENTER-Taste"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 220
    End With
End With

'-----------------------------------------------------------------------------------------------------------

Set CmBar = CmBrs.Add("ID_Betrag", xtpBarTop)
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
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(44))

    Set CmCon = .Add(xtpControlLabel, SY_Cap08, "Faktor :")
    With CmCon
        .ToolTipText = "Geben Sie hier den Multiplikator der Leistung ein"
        .Style = xtpButtonIconAndCaption
    End With

    Set CmEdi = .Add(xtpControlEdit, SY_SuMul, vbNullString)
    With CmEdi
        .EditStyle = xtpEditStyleCenter
        .ToolTipText = "Geben Sie hier den Multiplikator der Leistung ein"
        .Style = xtpButtonIconAndCaption
        .Text = "1"
        .Width = 40
    End With

    Set CmCon = .Add(xtpControlLabel, SY_Plac2, Space$(2))

    Set CmCon = .Add(xtpControlLabel, SY_Cap01, "Anzahl :")
    With CmCon
        .ToolTipText = "Geben Sie hier die Anzahl der Leistung ein"
        .Style = xtpButtonIconAndCaption
    End With

    Set CmEdi = .Add(xtpControlEdit, SY_SuAnz, vbNullString)
    With CmEdi
        .EditStyle = xtpEditStyleCenter
        .ToolTipText = "Geben Sie hier die Anzahl der Leistung ein"
        .Style = xtpButtonIconAndCaption
        .Text = "1"
        .Width = 40
    End With
    
    Set CmCon = .Add(xtpControlLabel, SY_Plac2, Space$(2))

    Set CmCon = .Add(xtpControlLabel, SY_Cap04, "Akonto :")
    With CmCon
        .ToolTipText = "Geben Sie hier den Akontobetrag ein"
        .Style = xtpButtonIconAndCaption
    End With

    Set CmEdi = .Add(xtpControlEdit, SY_SuAko, vbNullString)
    With CmEdi
        .EditStyle = xtpEditStyleRight
        .ToolTipText = "Geben Sie hier den Akontobetrag ein"
        .Style = xtpButtonIconAndCaption
        .Text = GlWa2
        .Width = 70
    End With

    Set CmCon = .Add(xtpControlLabel, SY_Plac3, Space$(2))

    Set CmCon = .Add(xtpControlLabel, SY_Cap09, "Minimalbetrag :")
    With CmCon
        .ToolTipText = "Geben Sie hier den minimalen Abrechnungsbetrag ein."
        .Style = xtpButtonIconAndCaption
    End With

    Set CmEdi = .Add(xtpControlEdit, SY_MiBet, vbNullString)
    With CmEdi
        .EditStyle = xtpEditStyleRight
        .ToolTipText = "Geben Sie hier den minimalen Abrechnungsbetrag ein."
        .Style = xtpButtonIconAndCaption
        .Text = GlWa2
        .Width = 70
    End With
    
    Set CmCon = .Add(xtpControlLabel, SY_Plac4, Space$(4))
    
    Set CmCon = .Add(xtpControlLabel, SY_Cap03, "Maximalbetrag :")
    With CmCon
        .ToolTipText = "Geben Sie hier den maximalen Abrechnungsbetrag ein."
        .Style = xtpButtonIconAndCaption
    End With

    Set CmEdi = .Add(xtpControlEdit, SY_MxBet, vbNullString)
    With CmEdi
        .EditStyle = xtpEditStyleRight
        .ToolTipText = "Geben Sie hier den maximalen Abrechnungsbetrag ein."
        .Style = xtpButtonIconAndCaption
        .Text = GlWa2
        .Width = 70
    End With
End With

'-----------------------------------------------------------------------------------------------------------

With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
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
    .SetIconSize True, 32, 32
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
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case 8:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case Else:
        If GlRah = True Then 'Office EnableThemeframe
            .VisualTheme = xtpThemeRibbon
        Else
            If GlFRg = True Then 'farbige Register
                .VisualTheme = xtpThemeResource
            Else
                .VisualTheme = xtpThemeRibbon
            End If
        End If
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End Select
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = True
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
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
    .KeyBindings.Add FCONTROL, Asc("A"), KY_CT_A
    .KeyBindings.Add FCONTROL, Asc("V"), KY_CT_V
    .KeyBindings.Add FCONTROL, VK_BACK, KY_CT_BS
End With

Set RbBar = CmBrs.Item(1)
With RbBar
    .AllowMinimize = False
    .AllowQuickAccessCustomization = False
    .AllowQuickAccessDuplicates = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .EnableAnimation = GlMeA
    .FontHeight = GlToF
    .GroupsVisible = True
    .MinimumVisibleWidth = 100
    .RibbonPaintManager.HotTrackingGroups = True
    .RibbonPaintManager.CaptionFont.SIZE = 8
    .RibbonPaintManager.CaptionFont.Name = GlTFt.Name
    .RibbonPaintManager.WindowCaptionFont.SIZE = 8
    .RibbonPaintManager.WindowCaptionFont.Name = GlTFt.Name
    .ShowQuickAccess = False
    .ShowQuickAccessBelowRibbon = False
    .ShowCaptionAlways = True
    .Position = xtpBarTop
    .SetIconSize 16, 16
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
            .TabPaintManager.Appearance = xtpTabAppearanceVisualStudio2005
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 6
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
        Else
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        End If
    End Select
    .TabPaintManager.Layout = xtpTabLayoutAutoSize
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.MinTabWidth = 100
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ClientFrame = xtpTabFrameNone
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = False
    .TabPaintManager.HotTracking = True
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.Font.SIZE = 8
    .TabPaintManager.Font.Name = GlTFt.Name
    If GlRDP = True Then
        .EnableFrameTheme
    Else
        If GlRah = True Then
            .EnableFrameTheme
        End If
    End If
End With

Set CmCon = Nothing
Set CmPop = Nothing
Set RbBar = Nothing
Set RbTab = Nothing
Set CmBrs = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " KrMen " & Err.Number
Resume Next

End Sub
Private Sub KrOpn(Optional ByVal EiTyp As Integer, Optional ByVal EiKop As Boolean = False)
On Error GoTo PoErr
'Öffnet das Kommentarfeld

Dim AkZe1 As Date
Dim AkZe2 As Date
Dim KrDat As Date
Dim TxFar As Long
Dim MitNr As Long
Dim Anzal As Single
Dim Multi As Single
Dim MiBtr As Single
Dim MxBtr As Single
Dim Akont As Single
Dim TeStr As String
Dim KoStr As String
Dim TmKTF As String
Dim KrTyp As Integer
Dim AktZa As Integer
Dim Lange As Integer
Dim AktPo As Integer
Dim KrLok As Boolean
Dim KrAbg As Boolean
Dim TxFor As XtremeSuiteControls.FlatEdit
Dim RbBar As XtremeCommandBars.RibbonBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit
Dim CmEd3 As XtremeCommandBars.CommandBarEdit
Dim CmEd4 As XtremeCommandBars.CommandBarEdit
Dim CmEd5 As XtremeCommandBars.CommandBarEdit
Dim CmEd6 As XtremeCommandBars.CommandBarEdit
Dim CmCoZ As XtremeCommandBars.CommandBarComboBox
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmFoA As XtremeCommandBars.CommandBarComboBox
Dim CmFoG As XtremeCommandBars.CommandBarComboBox
Dim CmMta As XtremeCommandBars.CommandBarComboBox
Dim CmAu1 As XtremeCommandBars.CommandBarComboBox
Dim CmAu2 As XtremeCommandBars.CommandBarComboBox
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmKraEd
Set TxFor = FM.txtForma
Set MoKa1 = FM.dtpDatu1
Set TxKom = FM.txtKomme
Set TxFil = FM.txtFiNam
Set TeDum = FM.txtDummy
Set TxNeu = FM.txtNeuEi
Set TxIdx = FM.txtIdxNr
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RbBar = CmBrs.Item(1)
Set RpCo6 = frmMain.repCont6
Set RpCoK = frmMain.repContK

Select Case GlBut
Case RibTab_Krankenbla:
        Set RpSel = RpCoK.SelectedRows
        Set RpCls = RpCoK.Columns
Case RibTab_Abrechnung:
        Set RpSel = RpCo6.SelectedRows
        Set RpCls = RpCo6.Columns
End Select

Set CmEd1 = CmBrs.FindControl(CmEd1, KA_Kalen, , True)
Set CmEd2 = CmBrs.FindControl(CmEd2, SY_SuAnz, , True)
Set CmEd3 = CmBrs.FindControl(CmEd3, SY_MiBet, , True)
Set CmEd6 = CmBrs.FindControl(CmEd6, SY_MxBet, , True)
Set CmEd4 = CmBrs.FindControl(CmEd4, SY_SuAko, , True)
Set CmEd5 = CmBrs.FindControl(CmEd5, SY_SuMul, , True)
Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmCoZ = CmBrs.FindControl(CmCoZ, KA_Uhrze, , True)
Set CmMta = CmBrs.FindControl(CmMta, KA_SuCo3, , True)
Set CmAu1 = CmBrs.FindControl(CmCom, SY_SuCm1, , True)
Set CmAu2 = CmBrs.FindControl(CmCom, SY_SuCm2, , True)
Set CmFoA = CmBrs.FindControl(CmFoA, Tex_FntAu6, , True)
Set CmFoG = CmBrs.FindControl(CmFoG, Tex_FntGr6, , True)

CmAcs(Tex_Suchen).Checked = GlSPh 'Suchleiste Textphrase
CmBrs.Item(2).Visible = GlSPh
CmBrs.Item(3).Visible = False

If EiTyp > 0 Then
    TxNeu.Text = EiTyp 'Wenn Inhalt vorhanden, dann neuer Eintrag
    TeDum.Text = CreateID("K")
    CmEd1.Text = Date
    CmMta.ListIndex = GlMiA(GlSmI, 0)
    For AktZa = 1 To UBound(GlKrA)
        If GlKrA(AktZa, 0) > 9 Then
            Select Case GlKrA(AktZa, 0)
            Case 24:
            Case 101:
            Case 102:
            Case 104:
            Case 105:
            Case Else:
                AktPo = AktPo + 1
                If EiTyp = GlKrA(AktZa, 0) Then
                    CmCom.ListIndex = AktPo
                    Exit For
                End If
            End Select
        End If
    Next AktZa
    If GlKFt.SIZE > 9 Then
        TmKTF = "0000L" & Format$(GlKrA(AktZa, 3), "00000000") & "16777215" & Format$(GlKFt.SIZE, "00") & GlKFt.Name
    Else
        TmKTF = "0000L" & Format$(GlKrA(AktZa, 3), "00000000") & "1677721510" & GlKFt.Name
    End If
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
    For AktZa = 1 To UBound(GlKrA)
        If EiTyp = GlKrA(AktZa, 0) Then
            If GlKrA(AktZa, 5) <> vbNullString Then
                CmBrs.Item(3).Visible = CBool(GlKrA(AktZa, 5))
            End If
            Exit For
        End If
    Next AktZa
Else
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            Set RpCol = RpCls.Find(Kra_Typ)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                KrTyp = RpRow.Record(RpCol.ItemIndex).Value
            End If
            Set RpCol = RpCls.Find(Kra_Datum)
            KrDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Kra_ID2)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                TxIdx.Text = Format$(RpRow.Record(RpCol.ItemIndex).Value, "000000")
            Else
                TxIdx.Text = "000000"
            End If
            Set RpCol = RpCls.Find(Kra_Uhrzeit)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                AkZe1 = TimeValue(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                AkZe1 = TimeValue(Now)
            End If
            Set RpCol = RpCls.Find(Kra_Lock)
            If RpRow.Record(RpCol.ItemIndex).Checked = True Then
                KrLok = True
            Else
                KrLok = False
            End If
            Set RpCol = RpCls.Find(Kra_Gedruckt)
            If RpRow.Record(RpCol.ItemIndex).Checked = True Then
                KrAbg = True
            Else
                KrAbg = False
            End If
            Set RpCol = RpCls.Find(Kra_IDM)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                MitNr = RpRow.Record(RpCol.ItemIndex).Value
            Else
                MitNr = 0
            End If
            If GlMiV = True Then
                For AktZa = 1 To UBound(GlMiA)
                    If MitNr = GlMiA(AktZa, 2) Then
                        CmMta.ListIndex = AktZa
                        Exit For
                    End If
                Next AktZa
            End If

            If EiKop = True Then
                TxNeu.Text = KrTyp 'Wenn Inhalt vorhanden, dann neuer Eintrag
                CmEd1.Text = Date
                TeDum.Text = CreateID("K")
                AkZe1 = TimeValue(Now)
            Else
                CmEd1.Text = KrDat
                Set RpCol = RpCls.Find(Kra_Zusatztext)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    TeDum.Text = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                End If
            End If
            
            For AktZa = 1 To CmCoZ.ListCount
                AkZe2 = CmCoZ.List(AktZa)
                If AkZe2 >= AkZe1 Then
                    Exit For
                End If
            Next AktZa
            CmCoZ.ListIndex = AktZa
            CmCoZ.Text = Format$(AkZe1, "hh:mm")
            For AktZa = 1 To UBound(GlKrA)
                If GlKrA(AktZa, 0) > 9 Then
                    Select Case GlKrA(AktZa, 0)
                    Case 24:
                    Case 101:
                    Case 102:
                    Case 104:
                    Case 105:
                    Case Else:
                        AktPo = AktPo + 1
                        If KrTyp = GlKrA(AktZa, 0) Then
                            CmCom.ListIndex = AktPo
                            Exit For
                        End If
                    End Select
                End If
            Next AktZa
            For AktZa = 1 To UBound(GlKrA) 'Krankenblatttypen
                If KrTyp = GlKrA(AktZa, 0) Then
                    If GlKrA(AktZa, 5) <> vbNullString Then
                        CmBrs.Item(3).Visible = CBool(GlKrA(AktZa, 5)) 'Betragsspalte aktivieren
                    End If
                    Exit For
                End If
            Next AktZa

            Set RpCol = RpCls.Find(Kra_Provision)
            If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                TmKTF = RpRow.Record(RpCol.ItemIndex).Value
            Else
                TmKTF = "0000L" & Format$(GlKrA(AktZa, 3), "00000000") & "1677721510Arial"
            End If
            Select Case GlBut
            Case RibTab_Krankenbla:
                Set RpCol = RpCls.Find(Kra_Bezeichnung)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    If TxFil.Text <> vbNullString Then
                        Set RpCol = RpCls.Find(Kra_Kommentar)
                        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                            TeStr = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                        Else
                            TeStr = vbNullString
                        End If
                    Else
                        TeStr = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                    End If
                Else
                    Set RpCol = RpCls.Find(Kra_Kommentar)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TeStr = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                    Else
                        TeStr = vbNullString
                    End If
                End If
                Set RpCol = RpCls.Find(Kra_Faktor)
                If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        Multi = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                    Else
                        Multi = 1
                    End If
                Else
                    Multi = 1
                End If
                Set RpCol = RpCls.Find(Kra_Anz)
                If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        Anzal = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                    Else
                        Anzal = 1
                    End If
                Else
                    Anzal = 1
                End If
                Set RpCol = RpCls.Find(Kra_Betrag)
                If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MiBtr = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                    Else
                        MiBtr = 0
                    End If
                Else
                    MiBtr = 0
                End If
                Set RpCol = RpCls.Find(Kra_GesBetrag)
                If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        MxBtr = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                    Else
                        MxBtr = 0
                    End If
                Else
                    MxBtr = 0
                End If
                Set RpCol = RpCls.Find(Kra_WVBetrag)
                If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                    If RpRow.Record(RpCol.ItemIndex).Value > 0 Then
                        Akont = CSng(RpRow.Record(RpCol.ItemIndex).Value)
                    Else
                        Akont = 0
                    End If
                Else
                    Akont = 0
                End If
            Case RibTab_Abrechnung:
                Set RpCol = RpCls.Find(Kra_Kommentar)
                If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                    TeStr = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
                Else
                    TeStr = vbNullString
                End If
            End Select
        End If
    End If
End If

If Mid$(TmKTF, 1, 1) = "1" Then
    TxKom.Font.Bold = True
    CmAcs(Tex_ForFet).Checked = True
End If
If Mid$(TmKTF, 2, 1) = "1" Then
    TxKom.Font.Italic = True
    CmAcs(Tex_ForKur).Checked = True
End If
If Mid$(TmKTF, 3, 1) = "1" Then
    TxKom.Font.Underline = True
    CmAcs(Tex_ForUnt).Checked = True
End If
If Mid$(TmKTF, 4, 1) = "1" Then
    TxKom.Font.Strikethrough = True
    CmAcs(Tex_ForDur).Checked = True
End If
If Mid$(TmKTF, 5, 1) = "L" Then
    TxKom.Alignment = xtpEditAlignLeft
    CmAcs(Tex_AusrLi).Checked = True
ElseIf Mid$(TmKTF, 5, 1) = "R" Then
    TxKom.Alignment = xtpEditAlignRight
    CmAcs(Tex_AusrRe).Checked = True
ElseIf Mid$(TmKTF, 5, 1) = "Z" Then
    TxKom.Alignment = xtpEditAlignCenter
    CmAcs(Tex_AusrZe).Checked = True
Else
    TxKom.Alignment = xtpEditAlignLeft
    CmAcs(Tex_AusrLi).Checked = True
End If
If Mid$(TmKTF, 6, 8) <> vbNullString Then
    TxKom.ForeColor = CLng(Mid$(TmKTF, 6, 8))
Else
    TxKom.ForeColor = 0
End If
If Mid$(TmKTF, 14, 8) <> vbNullString Then
    TxKom.BackColor = CLng(Mid$(TmKTF, 14, 8))
Else
    TxKom.BackColor = 16777215
End If
If Mid$(TmKTF, 22, 2) <> vbNullString Then
    TxKom.Font.SIZE = CLng(Mid$(TmKTF, 22, 2))
    CmFoG.Text = Mid$(TmKTF, 22, 2)
Else
    TxKom.Font.SIZE = GlKFt.SIZE
    CmFoG.Text = GlKFt.SIZE
End If
If Mid$(TmKTF, 24, Len(TmKTF) - 23) <> vbNullString Then
    TxKom.Font.Name = Mid$(TmKTF, 24, Len(TmKTF) - 23)
    CmFoA.Text = Mid$(TmKTF, 24, Len(TmKTF) - 23)
Else
    TxKom.Font.Name = "Arial"
    CmFoA.Text = "Arial"
End If
If TmKTF <> vbNullString Then
    TxFor.Text = TmKTF
Else
    TxFor = "0000L000000001677721510Arial"
End If

If TxFil.Text <> vbNullString Then
    CmCom.Enabled = False
End If

If TeStr <> vbNullString Then
    TxKom.Text = TeStr
End If

If Anzal = 0 Then
    Anzal = 1
End If

If Multi = 0 Then
    Multi = 1
End If

Lange = Len(TxKom.Text)

TxKom.SelStart = Lange

CmEd2.Text = Anzal
CmEd5.Text = Format$(Multi, GlWa1)
CmEd3.Text = Format$(MiBtr, GlWa1)
CmEd6.Text = Format$(MxBtr, GlWa1)
CmEd4.Text = Format$(Akont, GlWa1)

If KrLok = True Then
    TxKom.Enabled = False
    CmEd1.Enabled = False
    CmEd2.Enabled = False
    CmEd3.Enabled = False
    CmEd4.Enabled = False
    CmEd5.Enabled = False
    CmEd6.Enabled = False
    CmCoZ.Enabled = False
    CmCom.Enabled = False
    CmAu1.Enabled = False
    CmAu2.Enabled = False
    CmAcs(Tex_DatSpe).Enabled = False
    CmAcs(KA_KaBu1).Enabled = False
End If

If KrAbg = True Then
    TxKom.Enabled = False
    CmEd1.Enabled = False
    CmEd2.Enabled = False
    CmEd3.Enabled = False
    CmEd4.Enabled = False
    CmEd5.Enabled = False
    CmEd6.Enabled = False
    CmCoZ.Enabled = False
    CmCom.Enabled = False
    CmAu1.Enabled = False
    CmAu2.Enabled = False
    CmAcs(Tex_DatSpe).Enabled = False
    CmAcs(KA_KaBu1).Enabled = False
End If

CmSta.Pane(0).Text = "Anzahl Zeichen : " & Lange

Set MoKa1 = Nothing
Set CmBrs = Nothing
Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo6 = Nothing
Set RpCoK = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " KrOpn " & Err.Number
Resume Next

End Sub
Public Sub KrPos()
On Error GoTo PoErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBr2 As XtremeCommandBars.CommandBars

Set FM = frmKraEd
Set CmBr2 = FM.comBar02
Set TxKom = FM.txtKomme

If FM.WindowState <> vbMinimized Then
    CmBr2.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    TxKom.Move ClLin, ClObn, ClBre, ClHoh
End If

Set CmBr2 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " KrPos " & Err.Number
Resume Next

End Sub
Private Sub KrReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "Krankenblatt") = False Then
    xGro = 865
    yGro = 465
    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)
    
    IniSetSek "Krankenblatt"
    IniSetVal "Krankenblatt", "FenLin", xPos
    IniSetVal "Krankenblatt", "FenObe", yPos
    IniSetVal "Krankenblatt", "FenBre", xGro
    IniSetVal "Krankenblatt", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " KrReg " & Err.Number
Resume Next

End Sub
Public Sub KrSav(Optional ByVal SaFra As Boolean = False)
On Error GoTo PoErr

Dim IdxNr As Long
Dim KoMit As Long
Dim NeuDa As Date
Dim NeuZe As Date
Dim Anzal As Single
Dim Multi As Single
Dim MiBtr As Single
Dim MxBtr As Single
Dim Akont As Single
Dim KoGui As String
Dim TmKTF As String
Dim KoStr As String
Dim FiNam As String
Dim EiTyp As Integer
Dim NeuEi As Boolean
Dim Frage As Integer
Dim Mld1, Tit1 As String
Dim TxFor As XtremeSuiteControls.FlatEdit
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit
Dim CmEd3 As XtremeCommandBars.CommandBarEdit
Dim CmEd4 As XtremeCommandBars.CommandBarEdit
Dim CmEd5 As XtremeCommandBars.CommandBarEdit
Dim CmEd6 As XtremeCommandBars.CommandBarEdit
Dim CmCoZ As XtremeCommandBars.CommandBarComboBox
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmMta As XtremeCommandBars.CommandBarComboBox

Set FM = frmKraEd
Set TxFor = FM.txtForma
Set MoKa1 = FM.dtpDatu1
Set TxFil = FM.txtFiNam
Set TxKom = FM.txtKomme
Set TeDum = FM.txtDummy
Set TxNeu = FM.txtNeuEi
Set TxIdx = FM.txtIdxNr
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, KA_SuCo1, , True)
Set CmEd1 = CmBrs.FindControl(CmEd1, KA_Kalen, , True)
Set CmEd2 = CmBrs.FindControl(CmEd2, SY_SuAnz, , True)
Set CmEd3 = CmBrs.FindControl(CmEd3, SY_MiBet, , True)
Set CmEd6 = CmBrs.FindControl(CmEd6, SY_MxBet, , True)
Set CmEd4 = CmBrs.FindControl(CmEd4, SY_SuAko, , True)
Set CmEd5 = CmBrs.FindControl(CmEd5, SY_SuMul, , True)
Set CmCoZ = CmBrs.FindControl(CmCoZ, KA_Uhrze, , True)
Set CmMta = CmBrs.FindControl(CmMta, KA_SuCo3, , True)

Tit1 = "Eintrag Speichern"
Mld1 = "Soll dieser Eintrag gespeichert werden?"

GlNeK = GlKoX 'Neuer Kommentareintrag

EiTyp = CmCom.ItemData(CmCom.ListIndex)

If TxFor.Text <> vbNullString Then
    TmKTF = TxFor.Text
Else
    TmKTF = "0000L000000001677721510Arial"
End If

If IsDate(CmEd1.Text) = True Then
    NeuDa = CmEd1.Text
Else
    NeuDa = Date
End If

If CmEd5.Text <> vbNullString Then
    If IsNumeric(CmEd5.Text) = True Then
        Multi = CSng(CmEd5.Text)
    Else
        Multi = 1
    End If
Else
    Multi = 1
End If

If CmEd2.Text <> vbNullString Then
    If IsNumeric(CmEd2.Text) = True Then
        Anzal = CSng(CmEd2.Text)
    Else
        Anzal = 1
    End If
Else
    Anzal = 1
End If

If CmEd3.Text <> vbNullString Then
    If IsNumeric(CmEd3.Text) = True Then
        MiBtr = CSng(CmEd3.Text)
    Else
        MiBtr = 0
    End If
Else
    MiBtr = 0
End If

If CmEd6.Text <> vbNullString Then
    If IsNumeric(CmEd6.Text) = True Then
        MxBtr = CSng(CmEd6.Text)
    Else
        MxBtr = 0
    End If
Else
    MxBtr = 0
End If

If CmEd4.Text <> vbNullString Then
    If IsNumeric(CmEd4.Text) = True Then
        Akont = CSng(CmEd4.Text)
    Else
        Akont = 0
    End If
Else
    Akont = 0
End If

If TxNeu.Text <> vbNullString Then
    If IsNumeric(TxNeu.Text) = True Then
        NeuEi = True
    Else
        NeuEi = False
    End If
Else
    NeuEi = False
End If

If IsDate(CmCoZ.Text) = True Then
    NeuZe = TimeValue(CmCoZ.Text)
Else
    NeuZe = TimeValue(Now)
End If

If GlMiV = True Then
    KoMit = CmMta.ItemData(CmMta.ListIndex)
Else
    KoMit = 0
End If

If TxKom.Text <> vbNullString Then
    KoStr = TxKom.Text
Else
    KoStr = vbNullString
End If

If TxFil.Text <> vbNullString Then
    FiNam = TxKom.Text
Else
    FiNam = vbNullString
End If

If TeDum.Text <> vbNullString Then
    KoGui = TeDum.Text
Else
    KoGui = vbNullString
End If

If TxIdx.Text <> vbNullString Then
    IdxNr = TxIdx.Text
Else
    IdxNr = 0
End If

If SaFra = True Then
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
Else
    Frage = 6
End If

If Frage = 6 Then
    GlNeK = GlKoX
    If GlAdr > 0 Then
        If FiNam <> vbNullString Then
            With GlNeK
                .PatNr = GlAdr
                .IdxNr = IdxNr
                .EiDat = NeuDa
                .EiZei = NeuZe
                .EiTyp = EiTyp
                .KoStr = FiNam
                .KoGui = KoGui
                .NeuEi = NeuEi
                .TeStr = KoStr
                .FoStr = TmKTF
                .Mitar = KoMit
                .Anzal = Anzal
                .Multi = Multi
                .MiBet = MiBtr
                .MxBet = MxBtr
                .Akont = Akont
            End With
        Else
            With GlNeK
                .PatNr = GlAdr
                .IdxNr = IdxNr
                .EiDat = NeuDa
                .EiZei = NeuZe
                .EiTyp = EiTyp
                .KoStr = KoStr
                .KoGui = KoGui
                .NeuEi = NeuEi
                .FoStr = TmKTF
                .Mitar = KoMit
                .Anzal = Anzal
                .Multi = Multi
                .MiBet = MiBtr
                .MxBet = MxBtr
                .Akont = Akont
            End With
        End If
        K_Einf
        If FiNam <> vbNullString Then
            TxNeu.Text = vbNullString
        End If
    End If
End If

GlKaS = False 'Krankenblatteintrag Sepichern

Unload FM
DoEvents

SKrVo

Set CmBrs = Nothing
Set MoKa1 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " KrSav " & Err.Number
Resume Next

End Sub
Private Sub LVInit()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmLaVer
Set RpCon = FM.repCont1
Set ImMan = frmMain.imgManag

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = False 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips False
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = True
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False 'Gruppierungen
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
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.FixedRowHeight = False
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlKZe
    .PaintManager.UseAlternativeBackground = GlZeK
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

FM.BackColor = GlBak

Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " LVInit " & Err.Number
Resume Next

End Sub
Public Sub LVMain()
On Error GoTo LaErr
'Wrtevergleich

If WindowLoad("frmLaVer") = True Then
    frmLaVer.ZOrder 0
    Exit Sub
End If

GlWLa = True

LVReg

Load frmLaVer

Set FM = frmLaVer

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    .FeLin = IniGetVal("Laborvergleich", "FenLin")
    .FeObn = IniGetVal("Laborvergleich", "FenObe")
    .FeBre = IniGetVal("Laborvergleich", "FenBre")
    .FeHoh = IniGetVal("Laborvergleich", "FenHoh")
    .FenVor
End With

LVInit
LVOpn
DoEvents

With clFen
    .FenMov
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

DoEvents
LVPosi

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

frmLaVer.Show
DoEvents
GlWLa = False

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " LVMain " & Err.Number
Resume Next

End Sub
Public Sub LVOpn()
On Error GoTo GrErr

Dim PatNr As Long
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo5 = FM.repCont5
Set RpCls = RpCo5.Columns

Select Case GlBut
Case RibTab_LabBericht: Set RpSel = RpCo5.SelectedRows
Case RibTab_LabBerichte: Set RpSel = RpCo1.SelectedRows
End Select

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Lab_IDP)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            PatNr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            PatNr = 0
        End If
        S_LaVe PatNr
    End If
End If

Set RpSel = Nothing
Set RpRow = Nothing
Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo5 = Nothing

Exit Sub

GrErr:
If GlDbg = True Then SErLog Err.Description & " LVOpn " & Err.Number
Resume Next

End Sub
Public Sub LVPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmLaVer
Set CmBrs = FM.comBar02
Set RpCon = FM.repCont1

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    If ClBre > 10 Then
        If ClHoh > 100 Then
            RpCon.Move 10, ClObn + 10, ClBre - 20, ClHoh - 20
        End If
    End If
End If

Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " LVPosi " & Err.Number
Resume Next

End Sub
Private Sub LVReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "Laborvergleich") = False Then
    xGro = 880
    yGro = 740
    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)

    IniSetSek "Laborvergleich"
    IniSetVal "Laborvergleich", "FenLin", xPos
    IniSetVal "Laborvergleich", "FenObe", yPos
    IniSetVal "Laborvergleich", "FenBre", xGro
    IniSetVal "Laborvergleich", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " LVReg " & Err.Number
Resume Next

End Sub
Private Sub OuInit()
On Error GoTo InErr

Dim RetWe As Long
Dim OuAbg As Integer
Dim AkMon As Integer
Dim AkQua As Integer
Dim IdxZa As Integer
Dim BuJah As Integer
Dim AktZa As Integer
Dim Rahm0 As XtremeSuiteControls.GroupBox
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim MoKal As XtremeCalendarControl.DatePicker
Dim OpMon As XtremeSuiteControls.RadioButton
Dim OpQua As XtremeSuiteControls.RadioButton
Dim OpJah As XtremeSuiteControls.RadioButton
Dim OpZei As XtremeSuiteControls.RadioButton
Dim ChOp1 As XtremeSuiteControls.RadioButton
Dim ChOp2 As XtremeSuiteControls.RadioButton
Dim CmMon As XtremeSuiteControls.ComboBox
Dim CmQua As XtremeSuiteControls.ComboBox
Dim CmJah As XtremeSuiteControls.ComboBox
Dim ChSy1 As XtremeSuiteControls.CheckBox
Dim ChSy2 As XtremeSuiteControls.CheckBox
Dim ChFi1 As XtremeSuiteControls.CheckBox
Dim ChFi2 As XtremeSuiteControls.CheckBox
Dim ChVer As XtremeSuiteControls.CheckBox
Dim ChLoe As XtremeSuiteControls.CheckBox
Dim ChMap As XtremeSuiteControls.CheckBox
Dim PrBr1 As XtremeSuiteControls.ProgressBar
Dim PrBr2 As XtremeSuiteControls.ProgressBar

Set FM = frmOutlook
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set OpMon = FM.optZeit1
Set OpQua = FM.optZeit2
Set OpJah = FM.optZeit3
Set OpZei = FM.optZeit4
Set ChSy1 = FM.chkSync1
Set ChSy2 = FM.chkSync2
Set ChFi1 = FM.chkFilt1
Set ChFi2 = FM.chkFilt2
Set ChOp1 = FM.optOpti1
Set ChOp2 = FM.optOpti2
Set ChVer = FM.chkVergl
Set ChLoe = FM.chkLoWei
Set ChMap = FM.chkMapFo
Set CmMon = FM.cmbMonat
Set CmQua = FM.cmbQurta
Set CmJah = FM.cmbJahre
Set CmAbg = FM.cbmAbgle
Set CmMan = FM.cmbBehan
Set MoKal = FM.dtpDatu1
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnDatu2
Set TrLi1 = FM.trvList1
Set PrBr1 = FM.prbStat1
Set PrBr2 = FM.prbStat2
Set ImMan = frmMain.imgManag

OuAbg = IniGetVal("System", "OutAbg")
If CBool(IniGetVal("System", "OutVgl")) = True Then ChVer.Value = xtpChecked
If CBool(IniGetVal("System", "OutLoe")) = True Then ChLoe.Value = xtpChecked

AkMon = Month(Date)

If AkMon <= 3 Then
    AkQua = 1
ElseIf AkMon <= 6 Then
    AkQua = 2
ElseIf AkMon <= 9 Then
    AkQua = 3
ElseIf AkMon <= 12 Then
    AkQua = 4
End If

With PrBr1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .FlatStyle = False
    .Scrolling = xtpProgressBarStandard
    .UseVisualStyle = False
End With

With PrBr2
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .FlatStyle = False
    .Scrolling = xtpProgressBarStandard
    .UseVisualStyle = False
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
    .ToolTipText = "Markieren Sie bitte hier den gwünschten Rechnungstag"
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

With RpCo1
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .FocusSubItems = False
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .MultiSelectionMode = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    '.SetCustomDraw xtpCustomBeforeDrawRow
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
    .ShowHeader = True
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
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FocusSubItems = False
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .MultiSelectionMode = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    '.SetCustomDraw xtpCustomBeforeDrawRow
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
    .ShowHeader = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

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

With TrLi1
    .Nodes.Clear
    Set Knote = .Nodes.Add(, , "P800", "Adressen", IC16_Desktop)
    Knote.Bold = True
    Set Knote = .Nodes.Add("P800", 4, "P801", "Alle Adressen", IC16_Folder_View)
    Set Knote = .Nodes.Add("P800", 4, "P802", "Serienbriefadressen", IC16_Folder_Check)
    Set Knote = .Nodes.Add("P800", 4, "P803", "Onlineabgleich", IC16_Folder_Paper)
End With

With TrLi1
    .Nodes("P800").Expanded = True
    .Nodes("P801").Expanded = True
    .Nodes("P802").Expanded = True
    .Nodes("P803").Expanded = True
    .Nodes("P801").Checked = True
End With

With CmMon
    .DropDownItemCount = 12
    For IdxZa = 1 To 12
        .AddItem MonthName(IdxZa)
        .ItemData(.NewIndex) = IdxZa
    Next IdxZa
End With

With CmQua
    .AddItem "1. Quartal"
    .ItemData(.NewIndex) = 1
    .AddItem "2. Quartal"
    .ItemData(.NewIndex) = 2
    .AddItem "3. Quartal"
    .ItemData(.NewIndex) = 3
    .AddItem "4. Quartal"
    .ItemData(.NewIndex) = 4
End With

With CmJah
    .DropDownItemCount = 12
    For BuJah = Year(Date) - 15 To Year(Date) + 1
        .AddItem BuJah
        .ItemData(.NewIndex) = IdxZa
        IdxZa = IdxZa + 1
    Next BuJah
    .Text = Year(Date)
End With

With CmAbg
    .AddItem "Eigene Adressdaten und Outlook Kontakte"
    .ItemData(.NewIndex) = 1
    .AddItem "Eigene Termindaten und Outlook Termine"
    .ItemData(.NewIndex) = 2
    .AddItem "Alle Outlook Kontakte Entfernen"
    .ItemData(.NewIndex) = 3
    .AddItem "Alle Outlook Termine Entfernen"
    .ItemData(.NewIndex) = 4
End With

For AktZa = 1 To UBound(GlMan)
    With CmMan
        .AddItem GlMan(AktZa, 1)
        .ItemData(AktZa - 1) = GlMan(AktZa, 2)
    End With
Next AktZa
CmMan.ListIndex = GlMan(GlSMa, 0) - 1

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date) - 3
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date) + 1
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

RetWe = SendMessage(CmMon.hwnd, CB_SETCURSEL, AkMon - 1, ByVal 0&)
RetWe = SendMessage(CmQua.hwnd, CB_SETCURSEL, AkQua - 1, ByVal 0&)
RetWe = SendMessage(CmAbg.hwnd, CB_SETCURSEL, OuAbg, ByVal 0&)

Select Case OuAbg
Case 0: Rahm3.Visible = False
        Rahm4.Visible = True
        ChSy2.Enabled = True
        ChVer.Enabled = True
        ChLoe.Enabled = True
        ChMap.Enabled = True
Case 1: Rahm3.Visible = True
        Rahm4.Visible = False
        ChSy2.Enabled = True
        ChVer.Enabled = True
        ChLoe.Enabled = True
        ChMap.Enabled = True
Case 2: Rahm3.Visible = False
        Rahm4.Visible = True
        ChSy2.Enabled = False
        ChSy2.Value = xtpUnchecked
        ChVer.Enabled = False
        ChLoe.Enabled = False
        ChMap.Enabled = False
Case 3: Rahm3.Visible = True
        Rahm4.Visible = False
        ChSy2.Enabled = True
        ChVer.Enabled = False
        ChLoe.Enabled = False
        ChMap.Enabled = False
End Select

If GlMaF = True Then ChMap.Value = xtpChecked

OpZei.Value = True

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

FM.BackColor = GlBak
OpMon.BackColor = GlBak
OpQua.BackColor = GlBak
OpJah.BackColor = GlBak
OpZei.BackColor = GlBak
ChSy1.BackColor = GlBak
ChSy2.BackColor = GlBak
ChFi1.BackColor = GlBak
ChFi2.BackColor = GlBak
ChOp1.BackColor = GlBak
ChOp2.BackColor = GlBak
ChVer.BackColor = GlBak
ChLoe.BackColor = GlBak
ChMap.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak

Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " OuInit " & Err.Number
Resume Next

End Sub
Public Sub OuMain(Optional ByVal EiTyp As Integer, Optional ByVal DaNam As String)
On Error GoTo MeErr

If WindowLoad("frmOutlook") = True Then
    Set FM = frmOutlook
    frmOutlook.ZOrder 0
    Exit Sub
End If

OuReg

Load frmOutlook

Set FM = frmOutlook

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    .FeLin = IniGetVal("OutAbgl", "FenLin")
    .FeObn = IniGetVal("OutAbgl", "FenObe")
    .FeBre = IniGetVal("OutAbgl", "FenBre")
    .FeHoh = IniGetVal("OutAbgl", "FenHoh")
End With

OuInit
AdGru 6, True
AFont FM

With clFen
    .FenMov
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

frmOutlook.Show

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " OuMain " & Err.Number
Resume Next

End Sub
Private Sub OuReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "OutAbgl") = False Then
    xGro = 682
    yGro = 395
    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)
    
    IniSetSek "OutAbgl"
    IniSetVal "OutAbgl", "FenLin", xPos
    IniSetVal "OutAbgl", "FenObe", yPos
    IniSetVal "OutAbgl", "FenBre", xGro
    IniSetVal "OutAbgl", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " OuReg " & Err.Number
Resume Next

End Sub
Private Sub SeBuIn()
On Error GoTo InErr

Dim RetWe As Long
Dim Lbl05 As XtremeSuiteControls.Label
Dim Lbl06 As XtremeSuiteControls.Label
Dim ChAsw As XtremeSuiteControls.CheckBox
Dim CmBTy As XtremeSuiteControls.ComboBox
Dim CmBuT As XtremeSuiteControls.ComboBox
Dim FeWar As XtremeSuiteControls.ComboBox
Dim FeGeg As XtremeSuiteControls.ComboBox
Dim CmBuS As XtremeSuiteControls.ComboBox
Dim CmRam As XtremeSuiteControls.ComboBox
Dim TxKto As XtremeSuiteControls.FlatEdit
Dim TxHab As XtremeSuiteControls.FlatEdit
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmBuSer
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set Rahm8 = FM.frmRahm8
Set Rahm9 = FM.frmRahm9
Set Lbl05 = FM.lblLab05
Set Lbl06 = FM.lblLab06
Set TxDa1 = FM.txtDatu1
Set TxDa4 = FM.txtDatu4
Set ZyWoh = FM.cmbWoche
Set ZyMo1 = FM.cmoMona1
Set ZyMo2 = FM.cmoMona2
Set ZyMo3 = FM.cmbMona3
Set ZyMo4 = FM.cmbMonat
Set ZyJa1 = FM.cmoJahr1
Set ZyJa2 = FM.cmoJahr2
Set ZyJa3 = FM.cmoJahr3
Set ZyJa4 = FM.cmoJahr4
Set ZyMoT = FM.cmbMona1
Set ZyJaT = FM.cmbJahr1
Set ZyEnT = FM.cmbZyEn1
Set CmBTy = FM.cmbBuTyp
Set CmMan = FM.cmbBeha4
Set CmBuT = FM.cmbBuTex
Set CmBuS = FM.cmbBuStu
Set CmMit = FM.cmbMitar
Set CmRam = FM.cmbKtoRa
Set FeWar = FM.cmbWarun
Set FeGeg = FM.cmbGegen
Set CmRam = FM.cmbKtoRa
Set TxKto = FM.txtKonto
Set TxHab = FM.txtHaben
Set MoKa1 = FM.dtpDatu3
Set ZyEn2 = FM.optZyEn2
Set ZyEn3 = FM.optZyEn3
Set FoZy1 = FM.optZykl1
Set FoZy2 = FM.optZykl2
Set FoZy3 = FM.optZykl3
Set FoZy4 = FM.optZykl4
Set TaZy1 = FM.optZyTa1
Set TaZy2 = FM.optZyTa2
Set MoZy1 = FM.optZyMo1
Set MoZy2 = FM.optZyMo2
Set JaZy1 = FM.optZyJa1
Set JaZy2 = FM.optZyJa2
Set ChAsw = FM.chkGewEr
Set MoKa1 = FM.dtpDatu1
Set MoKa2 = FM.dtpDatu2
Set MoKa3 = FM.dtpDatu3
Set RpCo6 = FM.repCont6
Set ImMan = frmMain.imgManag

With MoKa1
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

With MoKa2
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

With MoKa3
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

With RpCo6
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FocusSubItems = True 'WICHTIG!
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Terminvorschläge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Terminvorschläge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = True
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
    If CBool(IniGetVal("Layout", "LinTyp")) = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.FixedRowHeight = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    .ShowHeader = GlSpU
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxDa4
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

If GlBuc = True Then 'Einfache Buchhaltung verwenden
    TxHab.Visible = False
    FeGeg.Visible = True
    Lbl05.Caption = "Sachkonto :"
    Lbl06.Caption = "Geldkonto :"
Else
    TxHab.Visible = True
    FeGeg.Visible = False
    Lbl05.Caption = "Soll-Konto :"
    Lbl06.Caption = "Haben-Konto :"
End If

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Rahm6.BackColor = GlBak
Rahm7.BackColor = GlBak
Rahm8.BackColor = GlBak
Rahm9.BackColor = GlBak
FoZy1.BackColor = GlBak
FoZy2.BackColor = GlBak
FoZy3.BackColor = GlBak
FoZy4.BackColor = GlBak
TaZy1.BackColor = GlBak
TaZy2.BackColor = GlBak
MoZy1.BackColor = GlBak
MoZy2.BackColor = GlBak
JaZy1.BackColor = GlBak
JaZy2.BackColor = GlBak
ZyEn2.BackColor = GlBak
ZyEn3.BackColor = GlBak
ChAsw.BackColor = GlBak
FM.choTaMon.BackColor = GlBak
FM.choTaDin.BackColor = GlBak
FM.choTaMit.BackColor = GlBak
FM.choTaDon.BackColor = GlBak
FM.choTaFre.BackColor = GlBak
FM.choTaSam.BackColor = GlBak
FM.choTaSon.BackColor = GlBak

Set RpCo6 = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " SeBuin " & Err.Number
Resume Next

End Sub
Private Sub SeBuLa()
On Error GoTo ReErr

Dim RetWe As Long
Dim AktZa As Integer
Dim AktKo As Integer
Dim WoTag As Integer
Dim JaMon As Integer
Dim AnzVo As Integer
Dim AnzSp As Integer
Dim PauSp As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim TxKto As XtremeSuiteControls.FlatEdit
Dim TxBuB As XtremeSuiteControls.FlatEdit
Dim ChAsw As XtremeSuiteControls.CheckBox
Dim CmBTy As XtremeSuiteControls.ComboBox
Dim CmBuT As XtremeSuiteControls.ComboBox
Dim FeWar As XtremeSuiteControls.ComboBox
Dim CmGeg As XtremeSuiteControls.ComboBox
Dim CmBuS As XtremeSuiteControls.ComboBox
Dim CmRam As XtremeSuiteControls.ComboBox

Set FM = frmBuSer
Set TxDa1 = FM.txtDatu1
Set TxDa4 = FM.txtDatu4
Set TxBuB = FM.txtBuBet
Set ZyWoh = FM.cmbWoche
Set ZyMo1 = FM.cmoMona1
Set ZyMo2 = FM.cmoMona2
Set ZyMo3 = FM.cmbMona3
Set ZyMo4 = FM.cmbMonat
Set ZyJa1 = FM.cmoJahr1
Set ZyJa2 = FM.cmoJahr2
Set ZyJa3 = FM.cmoJahr3
Set ZyJa4 = FM.cmoJahr4
Set ZyMoT = FM.cmbMona1
Set ZyJaT = FM.cmbJahr1
Set ZyEnT = FM.cmbZyEn1
Set CmBTy = FM.cmbBuTyp
Set CmMan = FM.cmbBeha4
Set CmBuT = FM.cmbBuTex
Set CmBuS = FM.cmbBuStu
Set CmRam = FM.cmbKtoRa
Set FeWar = FM.cmbWarun
Set CmGeg = FM.cmbGegen
Set TxKto = FM.txtKonto
Set CmMit = FM.cmbMitar
Set ZyEn2 = FM.optZyEn2
Set ZyEn3 = FM.optZyEn3
Set FoZy1 = FM.optZykl1
Set FoZy2 = FM.optZykl2
Set FoZy3 = FM.optZykl3
Set FoZy4 = FM.optZykl4
Set TaZy1 = FM.optZyTa1
Set TaZy2 = FM.optZyTa2
Set MoZy1 = FM.optZyMo1
Set MoZy2 = FM.optZyMo2
Set JaZy1 = FM.optZyJa1
Set JaZy2 = FM.optZyJa2
Set ChAsw = FM.chkGewEr

AnzVo = IniGetVal("TerSys", "AnzVor")
AnzSp = IniGetVal("TerSys", "TeSpAn")
PauSp = IniGetVal("TerSys", "TeSpPa")

With CmBTy
    .AddItem "Ausgabe"
    .ItemData(0) = 1
    .AddItem "Einnahme"
    .ItemData(1) = 2
    .ListIndex = 0
End With

For AktZa = 1 To UBound(GlWar)
    FeWar.AddItem GlWar(AktZa, 1)
    FeWar.ItemData(AktZa - 1) = GlWar(AktZa, 0)
Next AktZa
FeWar.ListIndex = GlStW - 1

With CmGeg
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            .AddItem GlGeK(AktZa, 3)
            .ItemData(.NewIndex) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    .AddItem GlSaK(AktKo, 3)
                    .ItemData(.NewIndex) = GlSaK(AktKo, 6) '[IDB]
                End If
            Next AktKo
        Next AktZa
        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(.NewIndex) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
        End If
    End If
End With
CmGeg.ListIndex = 0

For AktZa = 1 To UBound(GlBTe)
    CmBuT.AddItem GlBTe(AktZa, 1)
    CmBuT.ItemData(CmBuT.NewIndex) = GlBTe(AktZa, 0)
Next AktZa
CmBuT.AutoComplete = True

For AktZa = 1 To UBound(GlStu)
    CmBuS.AddItem GlStu(AktZa, 2)
    CmBuS.ItemData(CmBuS.NewIndex) = GlStu(AktZa, 0)
Next AktZa
CmBuS.ListIndex = GlStS - 1

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(CmMan.NewIndex) = GlThe(AktZa, 0)
Next AktZa
CmMan.ListIndex = GlMan(GlSMa, 0) - 1

If GlMiV = True Then
    For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
        With CmMit
            .AddItem GlMiK(AktZa, 1)
            .ItemData(.NewIndex) = GlMiK(AktZa, 2)
        End With
    Next AktZa
    For AktZa = 1 To UBound(GlMiK)
        If GlMiA(GlSmI, 2) = GlMiK(AktZa, 2) Then
            CmMit.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
End If

With CmRam
    For AktZa = 1 To UBound(GlKoR) 'Standardkontenrahmen
        .AddItem GlKoR(AktZa, 0)
        .ItemData(AktZa - 1) = GlKoR(AktZa, 1)
    Next AktZa
End With

With ZyWoh
    .AddItem "Jede Woche"
    .ItemData(0) = 1
    .AddItem "Jede zweite Woche"
    .ItemData(1) = 2
    .AddItem "Jede dritte Woche"
    .ItemData(2) = 3
    .AddItem "Jede vierte Woche"
    .ItemData(3) = 4
End With

With ZyMo1
    .AddItem "ersten"
    .ItemData(0) = 1
    .AddItem "zweiten"
    .ItemData(1) = 2
    .AddItem "dritten"
    .ItemData(2) = 3
    .AddItem "vierten"
    .ItemData(3) = 4
    .AddItem "letzten"
    .ItemData(4) = 5
End With

With ZyMo2
    .AddItem "Sonntag"
    .ItemData(0) = 1
    .AddItem "Montag"
    .ItemData(1) = 2
    .AddItem "Dienstag"
    .ItemData(2) = 3
    .AddItem "Mittwoch"
    .ItemData(3) = 4
    .AddItem "Donnerstag"
    .ItemData(4) = 5
    .AddItem "Freitag"
    .ItemData(5) = 6
    .AddItem "Samstag"
    .ItemData(6) = 7
End With

With ZyMo3
    .AddItem "jeden Monats"
    .ItemData(0) = 1
    .AddItem "jedes zweiten Monats"
    .ItemData(1) = 2
    .AddItem "jedes dritten Monats"
    .ItemData(2) = 3
    .AddItem "jedes vierten Monats"
    .ItemData(3) = 4
End With

With ZyMo4
    .AddItem "jeden Monats"
    .ItemData(0) = 1
    .AddItem "jedes zweiten Monats"
    .ItemData(1) = 2
    .AddItem "jedes dritten Monats"
    .ItemData(2) = 3
    .AddItem "jedes vierten Monats"
    .ItemData(3) = 4
End With

With ZyJa1
    .AddItem "Januar"
    .ItemData(0) = 1
    .AddItem "Februar"
    .ItemData(1) = 2
    .AddItem "März"
    .ItemData(2) = 3
    .AddItem "April"
    .ItemData(3) = 4
    .AddItem "Mai"
    .ItemData(4) = 5
    .AddItem "Juni"
    .ItemData(5) = 6
    .AddItem "Juli"
    .ItemData(6) = 7
    .AddItem "August"
    .ItemData(7) = 8
    .AddItem "September"
    .ItemData(8) = 9
    .AddItem "Oktober"
    .ItemData(9) = 10
    .AddItem "November"
    .ItemData(10) = 11
    .AddItem "Dezember"
    .ItemData(11) = 12
End With

With ZyJa4
    .AddItem "Januar"
    .ItemData(0) = 1
    .AddItem "Februar"
    .ItemData(1) = 2
    .AddItem "März"
    .ItemData(2) = 3
    .AddItem "April"
    .ItemData(3) = 4
    .AddItem "Mai"
    .ItemData(4) = 5
    .AddItem "Juni"
    .ItemData(5) = 6
    .AddItem "Juli"
    .ItemData(6) = 7
    .AddItem "August"
    .ItemData(7) = 8
    .AddItem "September"
    .ItemData(8) = 9
    .AddItem "Oktober"
    .ItemData(9) = 10
    .AddItem "November"
    .ItemData(10) = 11
    .AddItem "Dezember"
    .ItemData(11) = 12
End With

With ZyJa2
    .AddItem "ersten"
    .ItemData(0) = 1
    .AddItem "zweiten"
    .ItemData(1) = 2
    .AddItem "dritten"
    .ItemData(2) = 3
    .AddItem "vierten"
    .ItemData(3) = 4
    .AddItem "letzten"
    .ItemData(4) = 5
End With

With ZyJa3
     .AddItem "Sonntag"
    .ItemData(0) = 1
    .AddItem "Montag"
    .ItemData(1) = 2
    .AddItem "Dienstag"
    .ItemData(2) = 3
    .AddItem "Mittwoch"
    .ItemData(3) = 4
    .AddItem "Donnerstag"
    .ItemData(4) = 5
    .AddItem "Freitag"
    .ItemData(5) = 6
    .AddItem "Samstag"
    .ItemData(6) = 7
End With

With ZyMoT
    For AktZa = 0 To 31 - 1
        .AddItem Format$(AktZa + 1, "00") & "."
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With ZyJaT
    For AktZa = 0 To 31 - 1
        .AddItem Format$(AktZa + 1, "00") & "."
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With ZyEnT
    For AktZa = 2 To 99
        .AddItem AktZa & " Termine"
        .ItemData(AktZa - 2) = AktZa
    Next AktZa
End With

RetWe = SendMessage(ZyWoh.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyMo1.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyMo3.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyMo4.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyJa2.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyEnT.hwnd, CB_SETCURSEL, AnzVo, ByVal 0&)

TxDa4.Text = Date + ZyEnT.ItemData(ZyEnT.ListIndex)

WoTag = Weekday(Date)
JaMon = Month(Date)

If WoTag = 7 Then
    ZyMo2.ListIndex = 0
Else
    ZyMo2.ListIndex = WoTag - 1
End If

If (GlKtR - 1) <= (CmRam.ListCount) - 1 Then
    CmRam.ListIndex = GlKtR - 1 'Standardkontenrahmen
Else
    CmRam.ListIndex = 0
End If

TxBuB.Text = GlWa2

ZyJa3.ListIndex = WoTag - 1
ZyJa1.ListIndex = JaMon - 1
ZyJa4.ListIndex = JaMon - 1
ZyMoT.ListIndex = 0
ZyJaT.ListIndex = 0

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " SeBuLa " & Err.Number
Resume Next

End Sub
Public Sub SeBuMa(Optional ByVal PatNr As Long, Optional ByVal PatNa As String)
On Error GoTo LaErr

Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmBuSer") = True Then
    Set FM = frmBuSer
    frmBuSer.ZOrder 0
    Exit Sub
End If

GlTeF = True 'Formular wird geladen
GlSeF = True

SeBuRe

Load frmBuSer

Set FM = frmBuSer

Screen.MousePointer = vbHourglass

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    .FeLin = IniGetVal("SeBuVo", "FenLin")
    .FeObn = IniGetVal("SeBuVo", "FenObe")
    .FeBre = IniGetVal("SeBuVo", "FenBre")
    .FeHoh = IniGetVal("SeBuVo", "FenHoh")
End With

SeBuIn
AFont FM
SeBuMe
SeBuLa
SeBuSp
DoEvents
SeBuSt

With clFen
    .FenMov
    Set CmBrs = FM.comBar02
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    SeBuPo
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

frmBuSer.Show
DoEvents

Screen.MousePointer = vbNormal

GlTeF = False 'Formular wird geladen
GlSeF = False
GlVrt = False 'virtuelle Leistungen vorhanden

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " SeBuMa " & Err.Number
Resume Next

End Sub
Private Sub SeBuMe()
On Error GoTo InErr
'Menue erstellen

Dim RetWe As Long
Dim KeyNa As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim RbTem As XtremeCommandBars.RibbonTab
Dim MsBar As XtremeCommandBars.MessageBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmBuSer
Set CmBrs = FM.comBar02
Set PuBu1 = FM.btnDatu1
Set PuBu4 = FM.btnDatu4
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

KeyNa = "ToolTips"

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(AD_Termin_Vorschau, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Save, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Reset, vbNullString, vbNullString, vbNullString, vbNullString)
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Width = 120
    CmPan.Alignment = xtpAlignmentCenter
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Style = SBPS_STRETCH
    Set CmPan = .AddPane(3)
    CmPan.Width = 120
    CmPan.Text = vbNullString
    CmPan.Alignment = xtpAlignmentLeft
    .Visible = True
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")

Set CmBuT = RbBar.Controls.Add(xtpControlButton, TE_Termin_Hilfe, "Hilfe")
With CmBuT
    .ToolTipText = "Öffnet die Kurzhilfe"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F1"
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, TE_Termin_Beenden, "Schließen")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .ToolTipText = "Schließt den Dialog"
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

'---

Set RbTab = RbBar.InsertTab(RibTab_Ter_Haupt, "Termindaten")
With RbTab
    .id = RibTab_Ter_Haupt
    .ToolTip = "Zeigt die Hauptdaten des Termins"
    .Visible = True
    .Selected = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Vorschau, "Buchungen vorschlagen")
With CmCon
    .IconId = IC32_Calendar_Light
    .ShortcutText = "F3"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Save, "Vorschläge Speichern")
With CmCon
    .IconId = IC32_Disk_Calendar
    .ShortcutText = "F8"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Reset, "Vorschläge Zurücksetzen")
With CmCon
    .IconId = IC32_Calendar_Undo
    .ShortcutText = "F7"
    .Width = GlRib
End With

'---
With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
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
    .SetIconSize True, 32, 32
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
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case 8:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case Else:
        If GlRah = True Then 'Office EnableThemeframe
            .VisualTheme = xtpThemeRibbon
        Else
            If GlFRg = True Then 'farbige Register
                .VisualTheme = xtpThemeResource
            Else
                .VisualTheme = xtpThemeRibbon
            End If
        End If
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End Select
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = True
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

Set RbBar = CmBrs.Item(1)
With RbBar
    .AllowMinimize = False
    .AllowQuickAccessCustomization = False
    .AllowQuickAccessDuplicates = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .EnableAnimation = GlMeA
    .FontHeight = GlToF
    .GroupsVisible = True
    .MinimumVisibleWidth = 100
    .RibbonPaintManager.HotTrackingGroups = True
    .RibbonPaintManager.CaptionFont.SIZE = 8
    .RibbonPaintManager.CaptionFont.Name = GlTFt.Name
    .RibbonPaintManager.WindowCaptionFont.SIZE = 8
    .RibbonPaintManager.WindowCaptionFont.Name = GlTFt.Name
    .ShowQuickAccess = False
    .ShowQuickAccessBelowRibbon = False
    .ShowCaptionAlways = True
    .Position = xtpBarTop
    .SetIconSize 16, 16
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
            .TabPaintManager.Appearance = xtpTabAppearanceVisualStudio2005
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 6
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
        Else
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        End If
    End Select
    .TabPaintManager.Layout = xtpTabLayoutAutoSize
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.MinTabWidth = 100
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ClientFrame = xtpTabFrameNone
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = False
    .TabPaintManager.HotTracking = True
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.Font.SIZE = 8
    .TabPaintManager.Font.Name = GlTFt.Name
    If GlRDP = True Then
        .EnableFrameTheme
    Else
        If GlRah = True Then
            .EnableFrameTheme
        End If
    End If
End With

CmAcs(AD_Termin_Save).Enabled = False
CmAcs(AD_Termin_Reset).Enabled = False

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu4.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Set CmPan = Nothing
Set CmSta = Nothing
Set CmPop = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set RbBar = Nothing
Set RbTab = Nothing
Set RbGrp = Nothing
Set RbGps = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " SeBuMe " & Err.Number
Resume Next

End Sub
Private Sub SeBuSp()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmBuSer
Set RpCo6 = FM.repCont6

With RpCo6
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Set RpCls = RpCo6.Columns
With RpCls
    Set RpCol = .Add(Buh_ID0, "ID0", 0, False)
    Set RpCol = .Add(Buh_Datum, "Datum", 0, True)
    RpCol.Groupable = False
    Set RpCol = .Add(Buh_Buchtext, "Buchungstext", 0, True)
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        Set RpCol = .Add(Buh_Einnahme, "Einnahme", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Ausgabe, "Ausgabe", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Sachkonto, "Sachkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Geldkonto", 0, True)
    Else
        Set RpCol = .Add(Buh_Einnahme, "Betrag", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Ausgabe, "Brutto", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Buh_Sachkonto, "Sollkonto", 0, True)
        Set RpCol = .Add(Buh_Gegenkonto, "Habenkonto", 0, True)
    End If
    Set RpCol = .Add(Buh_RechNr, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Calendar_Disk
        .EditOptions.AllowEdit = True
        .Editable = True
    End With
    Set RpCol = .Add(Buh_IDR, "IDR", 0, False)
    Set RpCol = .Add(Buh_Beleg, "Nummer", 0, True)
    RpCol.Alignment = xtpAlignmentRight
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(Buh_Sachkontenbez, "Sachkontenbezeichnung", 0, True)
    Set RpCol = .Add(Buh_Geldkontenbez, "Geldkontenbezeichnung", 0, True)
    Set RpCol = .Add(Buh_Steuer, "Steuer", 0, True)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
    End With
    Set RpCol = .Add(Buh_W, "W", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
    End With
    Set RpCol = .Add(Buh_Privat, "Privat", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Abziehbar, "Abziehbar", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDB, "IDB", 0, False)
    Set RpCol = .Add(Buh_IDA, "IDA", 0, False)
    Set RpCol = .Add(Buh_Währung, "Währung", 0, False)
    Set RpCol = .Add(Buh_Ermittlung, "KE", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Dokument, "DK", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_Paperclip
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_IDP, "IDP", 0, False)
    Set RpCol = .Add(Buh_IDArt, "IDArt", 0, False)
    Set RpCol = .Add(Buh_IDBank, "IDBank", 0, False)
    Set RpCol = .Add(Buh_Kommentar, "Kommentar", 0, False)
    Set RpCol = .Add(Buh_IDT, "Mandant", 0, False)
    Set RpCol = .Add(Buh_Berichtdatum, "Bericht", 0, True)
    Set RpCol = .Add(Buh_GuiID, "GuiID", 0, False)
    Set RpCol = .Add(Buh_Monat, "Monat", 0, False)
    Set RpCol = .Add(Buh_Storniert, "Storniert", 0, False)
    RpCol.Tag = 1
    Set RpCol = .Add(Buh_IDM, "Mitarbeiter", 0, False)
    Set RpCol = .Add(Buh_Zuordnung, "ZU", 0, False)
    With RpCol
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_User_Norm
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Lock, "Lock", 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconLeft
        .Icon = IC16_Lock
        .Tag = 1
    End With
    Set RpCol = .Add(Buh_Datei, "Datei", 0, False)
    Set RpCol = .Add(Buh_Doppelt, "DO", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Tag = 1
    End With
End With

For Each RpCol In RpCls
    RpCol.Groupable = True
    RpCol.Sortable = False
    RpCol.AutoSize = False
    RpCol.Resizable = False
    RpCol.AutoSortWhenGrouped = False
Next RpCol

If GlTFt.SIZE > 10 Then
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 140
    RpCls(Buh_Buchtext).Width = 250
    RpCls(Buh_Einnahme).Width = 100
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        RpCls(Buh_Ausgabe).Width = 100
    Else
        RpCls(Buh_Ausgabe).Width = 0
    End If
Else
    RpCls(Buh_ID0).Width = 0
    RpCls(Buh_Datum).Width = 110
    RpCls(Buh_Buchtext).Width = 220
    RpCls(Buh_Einnahme).Width = 80
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        RpCls(Buh_Ausgabe).Width = 80
    Else
        RpCls(Buh_Ausgabe).Width = 0
    End If
End If
RpCls(Buh_Sachkonto).Width = 80
RpCls(Buh_Gegenkonto).Width = 80
RpCls(Buh_RechNr).Width = 25
RpCls(Buh_IDR).Width = 0
RpCls(Buh_Beleg).Width = 0
If GlBuc = True Then 'einfache Buchhaltung verwenden
    RpCls(Buh_Sachkontenbez).Width = 180
    RpCls(Buh_Geldkontenbez).Width = 160
Else
    RpCls(Buh_Sachkontenbez).Width = 0
    RpCls(Buh_Geldkontenbez).Width = 0
End If
RpCls(Buh_Steuer).Width = 75
RpCls(Buh_W).Width = 40
RpCls(Buh_Privat).Width = 0
RpCls(Buh_Abziehbar).Width = 0
RpCls(Buh_IDB).Width = 0
RpCls(Buh_IDA).Width = 0
RpCls(Buh_Währung).Width = 0
RpCls(Buh_Ermittlung).Width = 25
RpCls(Buh_Dokument).Width = 0
RpCls(Buh_IDP).Width = 0
RpCls(Buh_IDArt).Width = 0
RpCls(Buh_IDBank).Width = 0
RpCls(Buh_Kommentar).Width = 0
RpCls(Buh_IDT).Width = 180
RpCls(Buh_Berichtdatum).Width = 0
RpCls(Buh_GuiID).Width = 0
RpCls(Buh_Monat).Width = 0
RpCls(Buh_Storniert).Width = 0
RpCls(Buh_IDM).Width = 180
RpCls(Buh_Zuordnung).Width = 18
RpCls(Buh_Lock).Width = 18
RpCls(Buh_Datei).Width = 0
RpCls(Buh_Doppelt).Width = 0

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo6 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " SeBuSp " & Err.Number
Resume Next

End Sub
Public Sub SeBuSt()
On Error GoTo KoErr
'Errechnet das wirkliche Startdatum

Dim AnwDa As Date
Dim StaDa As Date
Dim TmpDa As Date
Dim AkDat As String
Dim WoTag As Integer
Dim DifTa As Integer
Dim DifWo As Integer
Dim DifMo As Integer
Dim StaTa As Integer
Dim TagNr As Integer
Dim Monat As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmBuSer
Set CmBrs = FM.comBar02
Set TxDa1 = FM.txtDatu1
Set TxDa4 = FM.txtDatu4
Set TxDa5 = FM.txtDatu5
Set ZyTag = FM.txoTage1
Set FoZy1 = FM.optZykl1
Set FoZy2 = FM.optZykl2
Set FoZy3 = FM.optZykl3
Set FoZy4 = FM.optZykl4
Set TaZy1 = FM.optZyTa1
Set TaZy2 = FM.optZyTa2
Set MoZy1 = FM.optZyMo1
Set MoZy2 = FM.optZyMo2
Set JaZy1 = FM.optZyJa1
Set JaZy2 = FM.optZyJa2
Set ZyEn2 = FM.optZyEn2
Set ZyEn3 = FM.optZyEn3
Set ZyWho = FM.cmbWoche
Set ZyMo1 = FM.cmoMona1
Set ZyMo2 = FM.cmoMona2
Set ZyMe1 = FM.cmbMona1
Set ZyMe2 = FM.cmbMonat
Set ZyMe3 = FM.cmbMona3
Set ZyJa1 = FM.cmoJahr1
Set ZyJa2 = FM.cmoJahr2
Set ZyJa3 = FM.cmoJahr3
Set ZyJa4 = FM.cmoJahr4
Set ZyJe1 = FM.cmbJahr1
Set ZyTer = FM.cmbZyEn1
Set ChMon = FM.choTaMon
Set ChDin = FM.choTaDin
Set ChMit = FM.choTaMit
Set ChDon = FM.choTaDon
Set ChFre = FM.choTaFre
Set ChSam = FM.choTaSam
Set ChSon = FM.choTaSon
Set CmSta = CmBrs.StatusBar

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        StaDa = CDate(TxDa1.Text)
    Else
        StaDa = Date
    End If
Else
    StaDa = Date
End If

AnwDa = S_TeDx(GlTVo, "VonDat")

If FoZy1.Value = True Then
    If TaZy1.Value = True Then
        If StaDa = AnwDa Then
            StaDa = DateAdd("d", 1, StaDa)
        Else
            StaDa = StaDa
        End If
    ElseIf TaZy2.Value = True Then
        If Weekday(StaDa) = vbSaturday Then
            StaDa = DateAdd("d", 2, StaDa)
        ElseIf Weekday(StaDa) = vbSunday Then
            StaDa = DateAdd("d", 1, StaDa)
        ElseIf StaDa = AnwDa Then
            StaDa = DateAdd("d", 1, StaDa)
        End If
    End If
ElseIf FoZy2.Value = True Then
    WoTag = Weekday(StaDa, vbMonday)
    DifWo = ZyWho.ItemData(ZyWho.ListIndex)
    Select Case WoTag
    Case 1: 'Montag
        If ChSon.Value = xtpChecked Then DifTa = 6
        If ChSam.Value = xtpChecked Then DifTa = 5
        If ChFre.Value = xtpChecked Then DifTa = 4
        If ChDon.Value = xtpChecked Then DifTa = 3
        If ChMit.Value = xtpChecked Then DifTa = 2
        If ChDin.Value = xtpChecked Then DifTa = 1
        If ChMon.Value = xtpChecked Then
            If ChSon.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 1
            Else
                DifTa = 7
            End If
        End If
    Case 2: 'Dinestag
        If ChMon.Value = xtpChecked Then DifTa = 6
        If ChSon.Value = xtpChecked Then DifTa = 5
        If ChSam.Value = xtpChecked Then DifTa = 4
        If ChFre.Value = xtpChecked Then DifTa = 3
        If ChDon.Value = xtpChecked Then DifTa = 2
        If ChMit.Value = xtpChecked Then DifTa = 1
        If ChDin.Value = xtpChecked Then
            If ChMon.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 1
            Else
                DifTa = 7
            End If
        End If
    Case 3: 'Mittwoch
        If ChDin.Value = xtpChecked Then DifTa = 6
        If ChMon.Value = xtpChecked Then DifTa = 5
        If ChSon.Value = xtpChecked Then DifTa = 4
        If ChSam.Value = xtpChecked Then DifTa = 3
        If ChFre.Value = xtpChecked Then DifTa = 2
        If ChDon.Value = xtpChecked Then DifTa = 1
        If ChMit.Value = xtpChecked Then
            If ChDin.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 1
            Else
                DifTa = 7
            End If
        End If
    Case 4: 'Donnerstag
        If ChMit.Value = xtpChecked Then DifTa = 6
        If ChDin.Value = xtpChecked Then DifTa = 5
        If ChMon.Value = xtpChecked Then DifTa = 4
        If ChSon.Value = xtpChecked Then DifTa = 3
        If ChSam.Value = xtpChecked Then DifTa = 2
        If ChFre.Value = xtpChecked Then DifTa = 1
        If ChDon.Value = xtpChecked Then
            If ChMit.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 1
            Else
                DifTa = 7
            End If
        End If
    Case 5: 'Freitag
        If ChDon.Value = xtpChecked Then DifTa = 6
        If ChMit.Value = xtpChecked Then DifTa = 5
        If ChDin.Value = xtpChecked Then DifTa = 4
        If ChMon.Value = xtpChecked Then DifTa = 3
        If ChSon.Value = xtpChecked Then DifTa = 2
        If ChSam.Value = xtpChecked Then DifTa = 1
        If ChFre.Value = xtpChecked Then
            If ChDon.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 1
            Else
                DifTa = 7
            End If
        End If
    Case 6: 'Samstag
        If ChFre.Value = xtpChecked Then DifTa = 6
        If ChDon.Value = xtpChecked Then DifTa = 5
        If ChMit.Value = xtpChecked Then DifTa = 4
        If ChDin.Value = xtpChecked Then DifTa = 3
        If ChMon.Value = xtpChecked Then DifTa = 2
        If ChSon.Value = xtpChecked Then DifTa = 1
        If ChSam.Value = xtpChecked Then
            If ChFre.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 1
            Else
                DifTa = 7
            End If
        End If
    Case 7: 'Sonntag
        If ChSam.Value = xtpChecked Then DifTa = 6
        If ChFre.Value = xtpChecked Then DifTa = 5
        If ChDon.Value = xtpChecked Then DifTa = 4
        If ChMit.Value = xtpChecked Then DifTa = 3
        If ChDin.Value = xtpChecked Then DifTa = 2
        If ChMon.Value = xtpChecked Then DifTa = 1
        If ChSon.Value = xtpChecked Then
            If ChSam.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 1
            Else
                DifTa = 7
            End If
        End If
    End Select
    Select Case DifWo
    Case 1: StaDa = DateAdd("d", DifTa, StaDa)
    Case 2: StaDa = DateAdd("d", DifTa + 7, StaDa)
    Case 3: StaDa = DateAdd("d", DifTa + 14, StaDa)
    Case 4: StaDa = DateAdd("d", DifTa + 21, StaDa)
    End Select
ElseIf FoZy3.Value = True Then
    If MoZy1.Value = True Then
        DifMo = ZyMe2.ItemData(ZyMe2.ListIndex)
        StaTa = ZyMe1.ItemData(ZyMe1.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        TmpDa = CDate(StaTa & "." & Format$(StaDa, "mm") & "." & Format$(StaDa, "yyyy"))
        DifTa = Abs(DateDiff("d", TmpDa, StaDa))
        If StaDa > TmpDa Then
            StaDa = DateAdd("m", 1, TmpDa)
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
            StaDa = DateAdd("m", 1, TmpDa)
        Else
            StaDa = StaDa
        End If
    ElseIf MoZy2.Value = True Then
        TagNr = ZyMo1.ItemData(ZyMo1.ListIndex)
        StaTa = ZyMo2.ItemData(ZyMo2.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        DifMo = ZyMe3.ItemData(ZyMe3.ListIndex)
        TmpDa = WoTaMo(StaTa, Month(StaDa), Year(StaDa), TagNr)
        If StaDa > TmpDa Then
            If Month(StaDa) + 1 > 12 Then
                StaDa = WoTaMo(StaTa, 1, Year(StaDa) + 1, TagNr)
            Else
                StaDa = WoTaMo(StaTa, Month(StaDa) + 1, Year(StaDa), TagNr)
            End If
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
             StaDa = WoTaMo(StaTa, Month(StaDa) + 1, Year(StaDa), TagNr)
        Else
            StaDa = StaDa
        End If
    End If
ElseIf FoZy4.Value = True Then
    If JaZy1.Value = True Then
        StaTa = ZyJe1.ItemData(ZyJe1.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        Monat = ZyJa1.ItemData(ZyJa1.ListIndex)
        TmpDa = CDate(StaTa & "." & Format$(Monat, "00") & "." & Format$(StaDa, "yyyy"))
        If StaDa > TmpDa Then
            StaDa = DateAdd("yyyy", 1, TmpDa)
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
            StaDa = DateAdd("yyyy", 1, TmpDa)
        Else
            StaDa = StaDa
        End If
    ElseIf JaZy2.Value = True Then
        TagNr = ZyJa2.ItemData(ZyJa2.ListIndex)
        StaTa = ZyJa3.ItemData(ZyJa3.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        Monat = ZyJa4.ItemData(ZyJa4.ListIndex)
        TmpDa = WoTaMo(StaTa, Monat, Year(StaDa), TagNr)
        If StaDa > TmpDa Then
            StaDa = WoTaMo(StaTa, Monat, Year(TmpDa) + 1, TagNr)
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
            StaDa = WoTaMo(StaTa, Monat, Year(TmpDa) + 1, TagNr)
        Else
            StaDa = StaDa
        End If
    End If
End If

TxDa5.Text = StaDa

CmSta.Pane(1).Text = "Neue Serienbuchung ab: " & Format$(StaDa, "dddd" & ", " & "dd" & ". " & "mmmm" & Chr$(32) & "yyyy")

Exit Sub

KoErr:
If GlDbg = True Then SErLog Err.Description & " SeBuSt " & Err.Number
Resume Next

End Sub
Public Sub SeBuPo()
On Error GoTo ReErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmBuSer
Set Rahm8 = FM.frmRahm8
Set Rahm9 = FM.frmRahm9
Set CmBrs = FM.comBar02
Set RpCo6 = FM.repCont6

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    Rahm9.Move ClLin, ClObn, ClBre - 7200, 4900
    Rahm8.Move ClBre - 7200, ClObn, 7100, 4900
    RpCo6.Move ClLin, ClObn + 4900, ClBre, ClHoh - 4900
End If

Set CmBrs = Nothing
Set RpCo6 = Nothing

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " SeBuPo " & Err.Number
Resume Next

End Sub
Private Sub SeBuRe()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If GlFnt = True Then
    xGro = 930
    yGro = 680
Else
    xGro = 930
    yGro = 720
End If

xPos = (GlxGr / 2) - (xGro / 2)
yPos = (GlyGr / 2) - (yGro / 2)

If IniGetSek(GlINI, "SeBuVo") = False Then IniSetSek "SeBuVo"
If IniGetVal("SeBuVo", "FenLin") = vbNullString Then IniSetVal "SeBuVo", "FenLin", xPos
If IniGetVal("SeBuVo", "FenObe") = vbNullString Then IniSetVal "SeBuVo", "FenObe", yPos
If IniGetVal("SeBuVo", "FenBre") = vbNullString Then IniSetVal "SeBuVo", "FenBre", xGro
If IniGetVal("SeBuVo", "FenHoh") = vbNullString Then IniSetVal "SeBuVo", "FenHoh", yGro

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " SeBuRe " & Err.Number
Resume Next

End Sub
Public Sub TeAdr(ByVal PatNr As Long, ByVal IdStr As String, Optional ByVal TbSel As Boolean = False)
On Error GoTo SeErr
'Füllt die Felder der Terminadresse mit Daten

Dim MitNr As Long
Dim ManNr As Long
Dim TmStr As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Dim LiIdx As Long
Dim TagWe As String
Dim Telef As String
Dim BrStr As String
Dim TeWer As Variant
Dim BeVor As Boolean
Dim AktZa As Integer

If WindowLoad("frmTermin") = True Then
    Set FM = frmTermin
    Set TxOrt = FM.txtRaum1
    If TxOrt.Text <> vbNullString Then
        TmStr = LCase(TxOrt.Text)
    End If
ElseIf WindowLoad("frmTermVo") = True Then
    Set FM = frmTermVo
    Set TxOrt = FM.txtRaum1
    If TxOrt.Text <> vbNullString Then
        TmStr = LCase(TxOrt.Text)
    End If
ElseIf WindowLoad("frmKatTE") = True Then
    Set FM = frmKatTE
End If

Set CmBrs = FM.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

Select Case RbTab.id
Case RibTab_Ter_Haupt:

    FM.txtID0.Text = PatNr
    TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
    FM.txtID0.Tag = 1 & TagWe

    FM.txtAdres.Text = IdStr
    TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
    FM.txtAdres.Tag = 1 & TagWe
    
    If Left$(TmStr, 6) <> "online" And Left$(TmStr, 6) <> "storno" Then
        S_AdDe PatNr 'Adressendetails
        With GlADt
            If .AdTe1 <> vbNullString Then
                Telef = .AdTe1
            ElseIf .AdTe2 <> vbNullString Then
                Telef = .AdTe2
            Else
                Telef = vbNullString
            End If
            
            FM.txtS4F01.Text = .AdFir
            TagWe = Mid$(FM.txtS4F01.Tag, 2, Len(FM.txtS4F01.Tag) - 1)
            FM.txtS4F01.Tag = 1 & TagWe

            FM.txtS4F02.Text = .AdAnr
            TagWe = Mid$(FM.txtS4F02.Tag, 2, Len(FM.txtS4F02.Tag) - 1)
            FM.txtS4F02.Tag = 1 & TagWe
                
            FM.txtS4F03.Text = .AdTit
            TagWe = Mid$(FM.txtS4F03.Tag, 2, Len(FM.txtS4F03.Tag) - 1)
            FM.txtS4F03.Tag = 1 & TagWe
            
            FM.txtS4F04.Text = .AdVor
            TagWe = Mid$(FM.txtS4F04.Tag, 2, Len(FM.txtS4F04.Tag) - 1)
            FM.txtS4F04.Tag = 1 & TagWe
                
            FM.txtS4F05.Text = .AdNam
            TagWe = Mid$(FM.txtS4F05.Tag, 2, Len(FM.txtS4F05.Tag) - 1)
            FM.txtS4F05.Tag = 1 & TagWe
                
            FM.txtS4F06.Text = .AdStr
            TagWe = Mid$(FM.txtS4F06.Tag, 2, Len(FM.txtS4F06.Tag) - 1)
            FM.txtS4F06.Tag = 1 & TagWe
                
            FM.txtS4F08.Text = .AdPLZ
            TagWe = Mid$(FM.txtS4F08.Tag, 2, Len(FM.txtS4F08.Tag) - 1)
            FM.txtS4F08.Tag = 1 & TagWe
                
            FM.txtS4F09.Text = .AdOrt
            TagWe = Mid$(FM.txtS4F09.Tag, 2, Len(FM.txtS4F09.Tag) - 1)
            FM.txtS4F09.Tag = 1 & TagWe
        
            FM.cmbS4F12.Text = .AdLan
            TagWe = Mid$(FM.cmbS4F12.Tag, 2, Len(FM.cmbS4F12.Tag) - 1)
            FM.cmbS4F12.Tag = 1 & TagWe
        
            FM.txtS4F18.Text = .AdGeb
            TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
            FM.txtS4F18.Tag = 1 & TagWe
        
            FM.txtS4F15.Text = Telef
            TagWe = Mid$(FM.txtS4F15.Tag, 2, Len(FM.txtS4F15.Tag) - 1)
            FM.txtS4F15.Tag = 1 & TagWe
        
            FM.txtS4F16.Text = .AdTe5
            TagWe = Mid$(FM.txtS4F16.Tag, 2, Len(FM.txtS4F16.Tag) - 1)
            FM.txtS4F16.Tag = 1 & TagWe
        
            BrStr = .AdBrf
            Ter_Brz BrStr
            TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
            FM.txtS4F18.Tag = 1 & TagWe
        End With
    End If
    
    GlTSa = True

Case RibTab_Ter_Adres:

    FM.txtID0.Text = PatNr
    TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
    FM.txtID0.Tag = 1 & TagWe
    
    FM.txtAdres.Text = IdStr
    TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
    FM.txtAdres.Tag = 1 & TagWe
    
    If Left$(TmStr, 6) <> "online" And Left$(TmStr, 6) <> "storno" Then
        S_AdDe PatNr 'Adressendetails
        With GlADt
            If .AdTe1 <> vbNullString Then
                Telef = .AdTe1
            ElseIf .AdTe2 <> vbNullString Then
                Telef = .AdTe2
            Else
                Telef = vbNullString
            End If
        
            FM.txtS4F01.Text = .AdFir
            TagWe = Mid$(FM.txtS4F01.Tag, 2, Len(FM.txtS4F01.Tag) - 1)
            FM.txtS4F01.Tag = 1 & TagWe
            
            FM.txtS4F02.Text = .AdAnr
            TagWe = Mid$(FM.txtS4F02.Tag, 2, Len(FM.txtS4F02.Tag) - 1)
            FM.txtS4F02.Tag = 1 & TagWe
            
            FM.txtS4F03.Text = .AdTit
            TagWe = Mid$(FM.txtS4F03.Tag, 2, Len(FM.txtS4F03.Tag) - 1)
            FM.txtS4F03.Tag = 1 & TagWe
            
            FM.txtS4F04.Text = .AdVor
            TagWe = Mid$(FM.txtS4F04.Tag, 2, Len(FM.txtS4F04.Tag) - 1)
            FM.txtS4F04.Tag = 1 & TagWe
            
            FM.txtS4F05.Text = .AdNam
            TagWe = Mid$(FM.txtS4F05.Tag, 2, Len(FM.txtS4F05.Tag) - 1)
            FM.txtS4F05.Tag = 1 & TagWe
            
            FM.txtS4F06.Text = .AdStr
            TagWe = Mid$(FM.txtS4F06.Tag, 2, Len(FM.txtS4F06.Tag) - 1)
            FM.txtS4F06.Tag = 1 & TagWe
            
            FM.txtS4F08.Text = .AdPLZ
            TagWe = Mid$(FM.txtS4F08.Tag, 2, Len(FM.txtS4F08.Tag) - 1)
            FM.txtS4F08.Tag = 1 & TagWe
            
            FM.txtS4F09.Text = .AdOrt
            TagWe = Mid$(FM.txtS4F09.Tag, 2, Len(FM.txtS4F09.Tag) - 1)
            FM.txtS4F09.Tag = 1 & TagWe
            
            FM.cmbS4F12.Text = .AdLan
            TagWe = Mid$(FM.cmbS4F12.Tag, 2, Len(FM.cmbS4F12.Tag) - 1)
            FM.cmbS4F12.Tag = 1 & TagWe
            
            FM.txtS4F18.Text = .AdGeb
            TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
            FM.txtS4F18.Tag = 1 & TagWe
            
            FM.txtS4F15.Text = Telef
            TagWe = Mid$(FM.txtS4F15.Tag, 2, Len(FM.txtS4F15.Tag) - 1)
            FM.txtS4F15.Tag = 1 & TagWe
            
            FM.txtS4F16.Text = .AdTe5
            TagWe = Mid$(FM.txtS4F16.Tag, 2, Len(FM.txtS4F16.Tag) - 1)
            FM.txtS4F16.Tag = 1 & TagWe
            
            BrStr = .AdBrf
            Ter_Brz BrStr
            TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
            FM.txtS4F18.Tag = 1 & TagWe
        End With
    End If

    GlTSa = True
    
Case RibTab_Ter_WarZi:

    If TbSel = True Then
        FM.txtID0.Text = PatNr
        FM.txtAdres.Text = IdStr
        TagWe = Mid$(FM.txtID0.Tag, 2, Len(FM.txtID0.Tag) - 1)
        FM.txtID0.Tag = 1 & TagWe
        TagWe = Mid$(FM.txtAdres.Tag, 2, Len(FM.txtAdres.Tag) - 1)
        FM.txtAdres.Tag = 1 & TagWe
        GlTSa = True
    Else
        Set CmMan = FM.cmbBehan
        Set CmMit = FM.cmbMitar
        ManNr = CmMan.ItemData(CmMan.ListIndex)
        MitNr = CmMit.ItemData(CmMit.ListIndex)
        Ter_Edi PatNr, True, MitNr, ManNr, 2 'in Warteliste aufnehmen
        DoEvents
        Ter_WaL
    End If
    
Case RibTab_Kat_EinTer:

    Unload frmAdrSuch
    frmWaKom.PatNr = PatNr
    frmWaKom.Show vbModal

End Select

Exit Sub

SeErr:
If GlDbg = True Then SErLog Err.Description & " TeAdr " & Err.Number
Resume Next

End Sub
Public Sub TeAkt()
On Error GoTo NeErr
'Bereitet die Neueingabe eines Termins vor

Dim EndZe As Date
Dim RetWe As Long
Dim MiNum As Long
Dim MaNum As Long
Dim MitNr As Long
Dim ManNr As Long
Dim NotDa As String
Dim NotZe As String
Dim NotSt As String
Dim ZeStr As String
Dim TagWe As String
Dim AkDat As String
Dim StaZe As String
Dim MiDif As Integer
Dim SelDf As Integer
Dim ZeiRa As Integer
Dim FltTy As Integer 'Filtertyp
Dim FltId As Integer
Dim MiIdx As Integer
Dim MaIdx As Integer
Dim AktZa As Integer
Dim NotVa As Integer
Dim mAnza As Integer
Dim MitOK As Boolean
Dim ManOK As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCo1 As XtremeCommandBars.CommandBarComboBox
Dim CmCo2 As XtremeCommandBars.CommandBarComboBox

If GlTeN = False Then 'neuen Termin hinzufügen
    Exit Sub
End If

If GlTeF = True Then 'Formular wird geladen
    Exit Sub
End If

If WindowLoad("frmTermin") = False Then
    Exit Sub
End If

Set FM = frmTermin
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set TxDa3 = FM.txtRzDat
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set TxNoS = FM.txtNoSta
Set TxNoD = FM.txtNoDat
Set TxNoZ = FM.txtNoTim
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmCo1 = frmMain.comBar01.FindControl(CmCo1, SY_TE_Termin_FiltTyp, , True)
Set CmCo2 = frmMain.comBar01.FindControl(CmCo2, SY_TE_Termin_FiltIdx, , True)

FltTy = CmCo1.ListIndex 'Filtertyp 1=Standard 2=Raum 3=Mitarbeiter
FltId = CmCo2.ListIndex - 1

Screen.MousePointer = vbHourglass

If Format$(GlSel.DaSta, "hh:mm:ss") = "00:00:00" Then 'Markierte Celle im Kalender
    If Not IsDate(Format$(GlSel.DaSta, "dd.mm.yyyy")) Then
        GlSel.DaSta = Format$(Now, "dd.mm.yyyy") & Chr$(32) & Format$(Now, "hh:mm:ss")
        AkDat = Format$(Now, "dd.mm.yyyy")
    Else
        If GlSel.DaSta = "00:00:00" Then
            AkDat = Date
        Else
            AkDat = Format$(GlSel.DaSta, "dd.mm.yyyy")
        End If
    End If
Else
    AkDat = Format$(GlSel.DaSta, "dd.mm.yyyy")
End If

TxDa1.Text = AkDat
TxDa2.Text = AkDat
TxDa3.Text = AkDat

Select Case GlBut
Case RibTab_Ter_Kalend:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        If FltTy = 3 Then 'WICHTIG Filtertyp Mitarbeiter
            CmMit.ListIndex = FltId
            MiIdx = FltId + 1
        Else
            If MiNum > 0 Then
                For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                    If MiNum = GlMiT(AktZa, 2) Then
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        MitOK = True
                        Exit For
                    End If
                Next AktZa
            Else
                For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                    If GlTBn = GlMiT(AktZa, 2) Then 'Termin Behandlernummer
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        MitOK = True
                        Exit For
                    End If
                Next AktZa
            End If
            If MitOK = True Then
                CmMit.ListIndex = GlTBx
                MiIdx = GlTBx + 1
            Else
                CmMit.ListIndex = 0
                MiIdx = 1
            End If
        End If
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(MiIdx, 39)
    Else
        If FltTy = 3 Then 'WICHTIG Filtertyp Mandant
            CmMan.ListIndex = FltId
            MaIdx = FltId + 1
        Else
            If MaNum > 0 Then
                For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                    If MaNum = GlMaT(AktZa, 2) Then
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        ManOK = True
                        Exit For
                    End If
                Next AktZa
            Else
                For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                    If GlTBn = GlMaT(AktZa, 2) Then 'Termin Behandlernummer
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        ManOK = True
                        Exit For
                    End If
                Next AktZa
            End If
            If ManOK = True Then
                CmMan.ListIndex = GlTBx
                MaIdx = GlTBx + 1
            Else
                CmMan.ListIndex = 0
                MaIdx = 1
            End If
        End If
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(MaIdx, 25)
    End If
    
Case RibTab_Ter_Raeume:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        If MiNum > 0 Then
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If MiNum = GlMiT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If GlTBn = GlMiT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If MitOK = False Then
            GlTBx = 0
        End If

        CmMit.ListIndex = GlTBx
        If GlTRa = True Then 'Mitarbeiter / Mandanten in Raumbelegung aktivieren
            If GlMiT(GlSmI + 1, 8) > 0 Then
                ZeiRa = GlMiT(GlSmI + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        Else
            If GlMiT(GlTBx + 1, 8) > 0 Then
                ZeiRa = GlMiT(GlTBx + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(AktZa, 39)
    Else
        If MaNum > 0 Then
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If MaNum = GlMaT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If GlTBn = GlMaT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If ManOK = False Then
            GlTBx = 1
        End If
        CmMan.ListIndex = GlTBx
        If GlTRa = True Then 'Mitarbeiter / Mandanten in Raumbelegung aktivieren
            If GlMaT(GlSMa + 1, 8) > 0 Then
                ZeiRa = GlMaT(GlSMa + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        Else
            If GlMaT(GlTBx + 1, 8) > 0 Then
                ZeiRa = GlMaT(GlTBx + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(AktZa, 25)
    End If
    
Case RibTab_Ter_Mitarb:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        MitNr = GlMiT(GlTBx + 1, 2)
        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
            If MitNr = GlMiT(AktZa, 2) Then
                Exit For
            End If
        Next AktZa
        CmMit.ListIndex = AktZa - 1
        If GlMiT(GlTBx + 1, 8) > 0 Then
            ZeiRa = GlMiT(GlTBx + 1, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(AktZa, 39)
    Else
        ManNr = GlMaT(GlTBx + 1, 2)
        For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
            If ManNr = GlMaT(AktZa, 2) Then
                Exit For
            End If
        Next AktZa
        CmMan.ListIndex = AktZa - 1
        If GlMaT(GlTBx + 1, 8) > 0 Then
            ZeiRa = GlMaT(GlTBx + 1, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(AktZa, 25)
    End If
    
Case Else:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        If MiNum > 0 Then
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If MiNum = GlMiT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If GlTBn = GlMiT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If MitOK = True Then
            CmMit.ListIndex = GlTBx
            MiIdx = GlTBx + 1
        Else
            CmMit.ListIndex = GlSmI - 1
            MiIdx = 1
        End If
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(AktZa, 39)
    Else
        If MaNum > 0 Then
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If MaNum = GlMaT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If GlTBn = GlMaT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If ManOK = True Then
            CmMan.ListIndex = GlTBx
            MaIdx = GlTBx + 1
        Else
            CmMan.ListIndex = GlSMa - 1
            MaIdx = 1
        End If
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(AktZa, 25)
    End If

End Select

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMan.ListCount > 1 Then
        ManNr = 0
        MitNr = CmMit.ItemData(CmMit.ListIndex)

        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
            If MitNr = CLng(GlMiT(AktZa, 2)) Then
                ManNr = GlMiT(AktZa, 7) 'zugeordnete Mandantennummer
                Exit For
            End If
        Next AktZa

        If ManNr > 0 Then
            For AktZa = 1 To UBound(GlMan)  'Aktive Mandanten
                If CBool(GlMan(AktZa, 5)) = False Then 'Passiv / Aktiv
                    If ManNr = CLng(GlMan(AktZa, 2)) Then
                        mAnza = mAnza + 1
                        Exit For
                    End If
                End If
            Next AktZa
            CmMan.ListIndex = AktZa - 1 'ManZa - 1
        Else
            CmMan.ListIndex = GlSMa - 1
        End If
    Else
        CmMan.ListIndex = GlSMa - 1
    End If
    TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
    CmMan.Tag = 1 & TagWe
Else
    If CmMit.ListCount > 1 Then
        CmMit.ListIndex = GlSmI - 1
    Else
        CmMit.ListIndex = GlSmI - 1
    End If
    TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
    CmMit.Tag = 1 & TagWe
End If

CmArz.ListIndex = 0
TagWe = Mid$(CmArz.Tag, 2, Len(CmArz.Tag) - 1)
CmArz.Tag = 1 & TagWe

MiDif = GlTku(ZeiRa, 2)
DoEvents
SRast ZeiRa
DoEvents

If GlSel.DaSta > 0 Then
    If Format$(GlSel.DaSta, "hh:mm") = "00:00" Then
        If MiDif = 0 Then MiDif = 15
        VoZei.Text = "08:00"
        EndZe = DateAdd("n", MiDif, "08:00:00")
        BiZei.Text = Format$(EndZe, "hh:mm")
    Else
        ZeStr = Format$(GlSel.DaSta, "hh:mm")
        For AktZa = 1 To UBound(GlRas) 'Zeitrasterstartzeiten
            If TimeValue(GlRas(AktZa)) <= TimeValue(ZeStr) Then
                StaZe = TimeValue(GlRas(AktZa))
            End If
        Next AktZa
        If GlSSt = True Then 'Starre Termintaktung
            EndZe = DateAdd("n", MiDif, StaZe)
            If TimeValue(GlSel.DaEnd) > TimeValue(EndZe) Then
                EndZe = Format$(GlSel.DaEnd, "hh:mm")
            End If
        Else
            SelDf = DateDiff("n", Format$(StaZe, "hh:mm"), Format$(GlSel.DaEnd, "hh:mm"))
            If SelDf < MiDif Then
                EndZe = DateAdd("n", MiDif, StaZe)
            Else
                EndZe = Format$(GlSel.DaEnd, "hh:mm")
            End If
        End If
        VoZei.Text = Format$(StaZe, "hh:mm")
        BiZei.Text = Format$(EndZe, "hh:mm")
    End If
Else
    If MiDif = 0 Then MiDif = 15
    VoZei.Text = "08:00"
    EndZe = DateAdd("n", MiDif, "08:00:00")
    BiZei.Text = Format$(EndZe, "hh:mm")
End If

If NotVa = 0 Then
    NotVa = 24
End If

If GlTeE = True Then 'Email-Termin-Erinnerung
    CmNot.ListIndex = NotVa

    NotDa = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "dd.mm.yyyy")
    NotZe = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "hh:mm")
    NotSt = NotDa & Chr$(32) & NotZe

    TxNoD.Text = NotDa
    TxNoZ.Text = NotZe

    TagWe = Mid$(TxNoD.Tag, 2, Len(TxNoD.Tag) - 1)
    TxNoD.Tag = "1" & TagWe
    
    TagWe = Mid$(TxNoZ.Tag, 2, Len(TxNoZ.Tag) - 1)
    TxNoZ.Tag = "1" & TagWe
    
    TagWe = Mid$(TxNoS.Tag, 2, Len(TxNoS.Tag) - 1)
    TxNoS.Tag = "1" & TagWe
Else
    RetWe = SendMessage(CmNot.hwnd, CB_SETCURSEL, NotVa, ByVal 0&)
End If

If GlTeE = True Then 'Email-Termin-Erinnerung
    If NotVa > 0 Then
        If CDate(NotSt) > Now Then
            TxNoS.Text = 3 'Senden
            CmAcs(AD_Termin_Notify).Checked = GlTeE 'Email-Termin-Erinnerung
        Else
            TxNoS.Text = 1 'Gesendet
            CmAcs(AD_Termin_Notify).Enabled = False
        End If
    Else
        TxNoS.Text = 0 'Nicht Senden
    End If
Else
    TxNoS.Text = 0 'Nicht Senden
End If

CmSta.Pane(1).Text = "Neuer Termin am: " & Format$(AkDat, "dddd" & ", " & "dd" & ". " & "mmmm" & Chr$(32) & "yyyy") & " um: " & Format$(VoZei.Text, "hh:mm") & " Uhr"

Screen.MousePointer = vbNormal

Exit Sub

NeErr:
If GlDbg = True Then SErLog Err.Description & " TeAkt " & Err.Number
Resume Next

End Sub
Public Function TeAna(ByVal TiStr As String, ByRef AnMin As Long) As Boolean
On Error Resume Next

AnMin = 0
TeAna = False
    
Dim AktZa As Long
Dim AkLen As Long
Dim AkMas As Long
Dim IdxNr As Long
Dim Teile As Long
Dim AkTim As Double
Dim TmpSt As String
Dim StrNr As String
Dim StMas As String
Dim StMon As String
    
TiStr = Trim(TiStr)
AkLen = Len(TiStr)

If AkLen = 0 Then Exit Function

AkMas = -1

For AktZa = 1 To AkLen
    TmpSt = Mid(TiStr, AktZa, 1)
    IdxNr = InStr(1, "-+.,0123456789", TmpSt)
    If IdxNr <= 0 Then
        AkMas = AktZa
        Exit For
    End If
Next
        
If AkMas > 0 Then
    StrNr = Left(TiStr, AkMas - 1)
    StMas = Mid(TiStr, AkMas)
    StMas = Trim(StMas)
Else
    StrNr = TiStr
End If

If Len(StrNr) = 0 Then Exit Function

StMon = Left(StMas, 1)

Teile = 1

If StMon = "m" Or StMon = "M" Then
    Teile = 1
ElseIf StMon = "s" Or StMon = "S" Then
    Teile = 60
ElseIf StMon = "t" Or StMon = "T" Then
    Teile = 60 * 24
ElseIf StMon = "w" Or StMon = "W" Then
    Teile = 60 * 24 * 7
End If

AkTim = Val(StrNr)
AnMin = AkTim * Teile

TeAna = True

End Function
Public Sub TerAd()
On Error GoTo ReErr
'Erstellt die Anschrift im Anschriftenfeld

Dim RAnsh As String
Dim KFirm As Variant
Dim KAnre As Variant
Dim KTite As Variant
Dim KName As Variant
Dim KVorn As Variant
Dim KStra As Variant
Dim KPost As Variant
Dim KOrte As Variant
Dim KLand As Variant

Set FM = frmTermin
KFirm = FM.txtS4F01.Text
KAnre = FM.txtS4F02.Text
KTite = FM.txtS4F03.Text
KName = FM.txtS4F05.Text
KVorn = FM.txtS4F04.Text
KStra = FM.txtS4F06.Text
KPost = FM.txtS4F08.Text
KOrte = FM.txtS4F09.Text
KLand = FM.cmbS4F12.Text

'Rechnunmgsanschrift
If GlFZe = False Then
    If Len(KFirm) > 1 Then
        RAnsh = Trim$(KFirm)
    End If
End If

If GlAno = False Then
    If KAnre <> vbNullString Then
        If KFirm <> vbNullString Then
            If Not KAnre = "Firma" Then
                RAnsh = RAnsh & vbCrLf & Trim$(KAnre)
            Else
                RAnsh = RAnsh & vbCrLf
            End If
        Else
            If KAnre <> "Firma" Then
                RAnsh = RAnsh & Trim$(KAnre)
            Else
                RAnsh = RAnsh
            End If
        End If
    End If
End If

If Len(KFirm) > 1 Then
    If KTite <> vbNullString Then
        If GlAno = False Then
            RAnsh = RAnsh & Chr$(32) & Trim$(KTite)
        Else
            RAnsh = RAnsh & Trim$(KTite)
        End If
        If KVorn <> vbNullString Then
            RAnsh = RAnsh & Chr$(32) & Trim$(KVorn)
            If KName <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(KName)
            End If
        Else
            If KName <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(KName)
            End If
        End If
    Else
        If KVorn <> vbNullString Then
            If KAnre <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(KVorn)
            Else
                RAnsh = RAnsh & vbCrLf & Trim$(KVorn)
            End If
            If KName <> vbNullString Then
                RAnsh = RAnsh & Chr$(32) & Trim$(KName)
            End If
        Else
            If KName <> vbNullString Then
                If GlAno = False Then
                    RAnsh = RAnsh & Chr$(32) & Trim$(KName)
                Else
                    RAnsh = RAnsh & Trim$(KName)
                End If
            End If
        End If
    End If
Else
    If KTite <> vbNullString Then
        RAnsh = RAnsh & vbCrLf & Trim$(KTite)
    End If
    If KVorn <> vbNullString Then
        If Not KTite = vbNullString Then
            RAnsh = RAnsh & Chr$(32) & Trim$(KVorn)
        Else
            RAnsh = RAnsh & vbCrLf & Trim$(KVorn)
        End If
    Else
        RAnsh = RAnsh & vbCrLf
    End If
    If KName <> vbNullString Then
        If Not KTite = vbNullString Or Not KVorn = vbNullString Then
            RAnsh = RAnsh & Chr$(32) & Trim$(KName)
        Else
            RAnsh = RAnsh & Trim$(KName)
        End If
    End If
End If

If GlFZe = True Then
    If Len(KFirm) > 1 Then
        RAnsh = RAnsh & vbCrLf & Trim$(KFirm)
    End If
End If

If KStra <> vbNullString Then RAnsh = RAnsh & vbCrLf & Trim$(KStra)
If KPost <> vbNullString Then RAnsh = RAnsh & vbCrLf & Trim$(KPost)
If KOrte <> vbNullString Then RAnsh = RAnsh & Chr$(32) & Trim$(KOrte)
If KLand <> vbNullString Then RAnsh = RAnsh & vbCrLf & UCase(Trim$(KLand))

If RAnsh <> FM.txtS3F01.Text Then
    FM.txtS3F01.Text = RAnsh
End If

Exit Sub

ReErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TerAd " & Err.Number
Resume Next

End Sub
Public Sub TeFarb(ByVal FaIdx As Integer, ByVal FoTyp As Integer)
On Error Resume Next
'Setzt die Farbe in den Termindialogen

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmPop As XtremeCommandBars.CommandBarPopup
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmCoS As XtremeCommandBars.CommandBarControls

If FaIdx <= 0 Then
    Exit Sub
End If

Select Case FoTyp
Case 1:
    Set FM = frmTermin
    Set CmBrs = FM.comBar02
    Set CmPop = CmBrs.FindControl(CmPop, TE_Farbe, , True)
Case 4:
    Set FM = frmTermVo
    Set CmBrs = FM.comBar02
    Set CmPop = CmBrs.FindControl(CmPop, TE_Farbe, , True)
Case 2:
    Set FM = frmKatTE
    Set CmBrs = FM.comBar02
    Set CmPop = CmBrs.FindControl(CmPop, SY_SuFar, , True)
Case 3:
    Set FM = frmMain
    Set CmBrs = FM.comBar01
    Select Case GlBut
    Case RibTab_Ter_Kalend: Set CmPop = CmBrs.FindControl(CmPop, SY_TE_Termin_FarTe, , True)
    Case RibTab_Ter_Raeume: Set CmPop = CmBrs.FindControl(CmPop, SY_TE_Termin_FarRa, , True)
    Case RibTab_Ter_Mitarb: Set CmPop = CmBrs.FindControl(CmPop, SY_TE_Termin_FarBe, , True)
    End Select
End Select

Set CmCoS = CmPop.CommandBar.Controls

For Each CmCon In CmCoS
    CmCon.Checked = False
Next CmCon

CmCoS(FaIdx).Checked = True

Select Case FoTyp
Case 1:
    Set FM = frmTermin
    Set CmBet = FM.txtBetre
    CmBet.BackColor = GlTmF(FaIdx, 1)
Case 4:
    Set FM = frmTermVo
    Set CmBet = FM.txtBetre
    CmBet.BackColor = GlTmF(FaIdx, 1)
End Select

Set CmBrs = Nothing
Set CmPop = Nothing

End Sub
Public Function TeFor(ByVal MinAn As Long, ByVal Ungef As Boolean) As String
On Error Resume Next
'Formatiert einen Termin Zeitstring
    
Dim AnzWo As Long
Dim AnzTa As Long
Dim AnzSt As Long
Dim TmStr As String

AnzWo = MinAn / (7 * 24 * 60)
AnzTa = MinAn / (24 * 60)
AnzSt = MinAn / 60

If (Ungef Or (MinAn Mod (7 * 24 * 60)) = 0) And AnzWo > 0 Then
    TmStr = AnzWo & " Woche" & IIf(AnzWo > 1, "n", "")
ElseIf (Ungef Or (MinAn Mod (24 * 60)) = 0) And AnzTa > 0 Then
    TmStr = AnzTa & " Tag" & IIf(AnzTa > 1, "e", "")
ElseIf (Ungef Or (MinAn Mod 60) = 0) And AnzSt > 0 Then
    TmStr = AnzSt & " Stunde" & IIf(AnzSt > 1, "n", "")
Else
    TmStr = MinAn & " Minute" & IIf(MinAn > 1, "n", "")
End If

TeFor = TmStr
    
End Function
Private Sub TeInit()
On Error GoTo InErr

Dim RetWe As Long
Dim ZeiUm As Boolean
Dim AktZa As Integer
Dim TxPIN As XtremeSuiteControls.FlatEdit
Dim TxZGe As XtremeSuiteControls.FlatEdit
Dim FeAn3 As XtremeSuiteControls.ComboBox
Dim FeLa3 As XtremeSuiteControls.ComboBox
Dim PuPo3 As XtremeSuiteControls.PushButton
Dim PuBu8 As XtremeSuiteControls.PushButton
Dim PuBu9 As XtremeSuiteControls.PushButton
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmTermin
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set TxDa3 = FM.txtRzDat
Set CmETy = FM.cmbTypen
Set CmZif = FM.cmbZiffe
Set CmBez = FM.cmbBezei
Set CmAbr = FM.cmbAbger
Set CmNot = FM.cmbNotVa
Set CmGan = FM.cmbGanzt
Set CmAbg = FM.cmbAbgeh
Set CmOnT = FM.cmbOnlTe
Set TxPIN = FM.txtS4F20
Set TxAnz = FM.txtAnzal
Set TxMul = FM.txtMulti
Set TxEin = FM.txtEinze
Set TxRef = FM.txtRefNr
Set TxZGe = FM.txtS4F18
Set TxRzn = FM.txtRzNum
Set TxRzA = FM.txtRzAnz
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set FeAn3 = FM.txtS4F02
Set TxOrt = FM.txtRaum1
Set FeLa3 = FM.cmbS4F12
Set MoKa1 = FM.dtpDatu1
Set PuBu2 = FM.btnDatu2
Set PuPo3 = FM.btnPost3
Set PuBu8 = FM.btnTele8
Set PuBu9 = FM.btnTele9
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2
Set ImMan = frmMain.imgManag

ZeiUm = False

With MoKa1
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

With RpCo1
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .AllowSelectionCheck = False
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .EnableMarkup = False 'XAML
    .FocusSubItems = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = True
    .LockExpand = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.AllowMergeCells = True
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.ColumnShadowGradient = -2147483643
    .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind keine Leistungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Leistungen vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = True
    .PaintManager.GroupRowTextBold = True
    .PaintManager.ShadeSortColumn = True
    .PaintManager.TreeStructureStyle = xtpTreeStructureDots
    .PaintManager.UseColumnTextAlignment = True
    .PaintManager.UseEditTextAlignment = True
    .PaintManager.TextFont.Name = GlTFt.Name
    .PaintManager.TextFont.SIZE = GlTFt.SIZE
    .PaintManager.ForeColor = GlFoF
    .PaintManager.MaxPreviewLines = GlAnZ
    .PaintManager.ShowNonActiveInPlaceButton = True
    .PaintManager.ThemedInplaceButtons = True
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = Not ZeiUm 'Zeilenumbruch der Kataloge
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 192, -2, 20, 6
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.ColumnWidthWYSIWYG = False
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlKZe
    .PaintManager.UseAlternativeBackground = GlZeK
    .MultiSelectionMode = False
    .ShowGroupBox = False
    .ShowIconWhenEditing False
    .PreviewMode = GlVoA
    .ShowHeader = True
    .SortedDragDrop = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    .OLEDropMode = xtpOLEDropNone
    RetWe = .EnableDragDrop("Katalog", xtpReportAllowDrop)
End With

With RpCo2
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = False
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Einträge vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Einträge vorhanden"
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
    .PaintManager.FixedRowHeight = Not ZeiUm 'Zeilenumbruch der Kataloge
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .MultiSelectionMode = False
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With CmETy
    .AutoComplete = False
    .DropDownWidth = 1900
End With

With CmZif
    .AutoComplete = False
    .DropDownItemCount = 20
    .DropDownWidth = 4400
End With

With CmBez
    .AutoComplete = True
    .DropDownItemCount = 20
End With

With CmAbr
    .AddItem "keine Leistungen"
    .ItemData(0) = 1
    .AddItem "Leistungen vorhanden"
    .ItemData(1) = 2
    .AddItem "Leistungen abgerechnet"
    .ItemData(2) = 3
End With

With FeAn3 'Anreden
    For AktZa = 0 To UBound(GlAnr) - 1
        .AddItem GlAnr(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With CmGan
    .AddItem "Ja"
    .ItemData(0) = -1
    .AddItem "Nein"
    .ItemData(1) = 0
End With

With CmAbg
    .AddItem "Ja"
    .ItemData(0) = -1
    .AddItem "Nein"
    .ItemData(1) = 0
End With

With CmOnT
    .AddItem "Ja"
    .ItemData(0) = -1
    .AddItem "Nein"
    .ItemData(1) = 0
End With

For AktZa = 1 To UBound(GlLan)
    With FeLa3
        .AddItem GlLan(AktZa, 1)
        .ItemData(AktZa - 1) = GlLan(AktZa, 0)
    End With
Next AktZa

With CmNot
    For AktZa = 0 To 48
        .AddItem AktZa & " Std."
        .ItemData(AktZa) = AktZa
    Next AktZa
End With

VoZei.SetMask "00:00", "__:__"
BiZei.SetMask "00:00", "__:__"
TxDa1.SetMask "00.00.0000", "__.__.____"
TxDa2.SetMask "00.00.0000", "__.__.____"
TxDa3.SetMask "00.00.0000", "__.__.____"

With TxRef
    .Pattern = "\d*"
    .SetMask "000000", "______"
    .Text = "000001"
End With

With TxRzn
    .Pattern = "\d*"
    .Text = "000000"
End With

With TxRzA
    .Pattern = "\d*"
    .Text = "000"
End With

TxPIN.Pattern = "\d*"

TxZGe.SetMask "00.00.0000", "__.__.____"

TxAnz.Text = "1"
TxMul.Text = GlWa3
TxEin.Text = GlWa2

PuPo3.Icon = ImMan.Icons.GetImage(IC16_Mailbox, 16)
PuBu8.Icon = ImMan.Icons.GetImage(IC16_Telephone, 16)
PuBu9.Icon = ImMan.Icons.GetImage(IC16_Earth_Mail, 16)
If GlRDP = True Then PuBu9.Enabled = False

If GlOTS = True Then 'Online-Terminbuchungs Sytem
    If GlOTK = False Then 'Online-Terminbuchungs System autom. Aktualisierung
        TxOrt.Enabled = False
    End If
End If

TxDa2.Enabled = Not GlOTS 'Online-Terminbuchungs Sytem
PuBu2.Enabled = Not GlOTS 'Online-Terminbuchungs Sytem

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak

Set RpCo1 = Nothing
Set RpCo2 = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " TeInit " & Err.Number
Resume Next

End Sub
Private Sub TeLoa(Optional ByVal IdxNr As Long)
On Error GoTo ReErr

Dim RetWe As Long
Dim ManNr As Long
Dim RmuNr As Long
Dim AktZa As Integer
Dim GesZa As Integer
Dim mAnza As Integer

Set FM = frmTermin
Set CmBet = FM.txtBetre
Set CmMar = FM.cmbTeTyp
Set CmTyp = FM.cmbStatu
Set CmRmu = FM.cmbRaum1
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set CmArz = FM.cmbArzNr
Set CmGes = FM.cmbGesch
Set CmETy = FM.cmbTypen
Set CmBez = FM.cmbBezei
Set TxAnz = FM.txtAnzal
Set TxMul = FM.txtMulti
Set TxEin = FM.txtEinze
Set TxKom = FM.txtKomme
Set TxRzn = FM.txtRzNum
Set TxRzA = FM.txtRzAnz
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set TxDa2 = FM.txtDatu2
Set PuBu2 = FM.btnDatu2
Set CmRem = FM.cmbRemin
Set CmPri = FM.cmbPrior
Set Lab17 = FM.lblLab17

For AktZa = 1 To UBound(GlBtr)
    With CmBet
        .AddItem GlBtr(AktZa, 1)
        .ItemData(AktZa - 1) = GlBtr(AktZa, 0)
    End With
Next AktZa

For AktZa = 1 To UBound(GlTep) 'Kalendermarker
    With CmMar
        .AddItem GlTep(AktZa, 1)
        .ItemData(AktZa - 1) = GlTep(AktZa, 0)
    End With
Next AktZa

For AktZa = 1 To UBound(GlTeS)
    With CmTyp
        .AddItem GlTeS(AktZa, 1)
        .ItemData(AktZa - 1) = GlTeS(AktZa, 0)
    End With
Next AktZa

If GlArV = True Then
   For AktZa = 1 To UBound(GlArz) 'Verordner
        CmArz.AddItem GlArz(AktZa, 8)
        CmArz.ItemData(AktZa - 1) = GlArz(AktZa, 0)
    Next AktZa
End If

AktZa = 1
GesZa = UBound(GlRmu)
RmuNr = GlRmu(GesZa, 2) + 1

If GlTRa = True Then 'Mitarbeiter / Mandanten in Raumbelegung aktivieren
    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        ManNr = GlMiA(GlSmI, 2)
    Else
        ManNr = GlMiA(GlSmI, 7)
    End If
    For AktZa = 1 To GesZa
        If GlRmu(AktZa, 4) = ManNr Then
            With CmRmu
                .AddItem GlRmu(AktZa, 1)
                .ItemData(AktZa - 1) = GlRmu(AktZa, 2)
            End With
        End If
    Next AktZa
Else
    For AktZa = 1 To GesZa
        With CmRmu
            .AddItem GlRmu(AktZa, 1)
            .ItemData(AktZa - 1) = GlRmu(AktZa, 2)
        End With
    Next AktZa
End If

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    For AktZa = 1 To UBound(GlMan)
        With CmMan
            If IdxNr = 0 Then
                If CBool(GlMan(AktZa, 5)) = False Then 'Passiv / Aktiv
                    mAnza = mAnza + 1
                    .AddItem GlMan(AktZa, 1)
                    .ItemData(mAnza - 1) = GlMan(AktZa, 2)
                End If
            Else
                mAnza = mAnza + 1
                .AddItem GlMan(AktZa, 1)
                .ItemData(mAnza - 1) = GlMan(AktZa, 2)
            End If
        End With
    Next AktZa
Else
    For AktZa = 1 To UBound(GlMaT)
        With CmMan
            If IdxNr = 0 Then
                If CBool(GlMaT(AktZa, 5)) = False Then 'Passiv / Aktiv
                    mAnza = mAnza + 1
                    .AddItem GlMaT(AktZa, 1)
                    .ItemData(mAnza - 1) = GlMaT(AktZa, 2)
                End If
            Else
                mAnza = mAnza + 1
                .AddItem GlMaT(AktZa, 1)
                .ItemData(mAnza - 1) = GlMaT(AktZa, 2)
            End If
        End With
    Next AktZa
End If

If GlMiV = True Then
    For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminliste
        With CmMit
            .AddItem GlMiT(AktZa, 1)
            .ItemData(AktZa - 1) = GlMiT(AktZa, 2)
        End With
    Next AktZa
End If

With CmRem
    .AddItem "0 Min."
    .ItemData(0) = 1
    .AddItem "1 Min."
    .ItemData(1) = 2
    .AddItem "2 Min."
    .ItemData(2) = 3
    .AddItem "5 Min."
    .ItemData(3) = 4
    .AddItem "10 Min."
    .ItemData(4) = 5
    .AddItem "15 Min."
    .ItemData(5) = 6
    .AddItem "30 Min."
    .ItemData(6) = 7
    .AddItem "1 Std."
    .ItemData(7) = 8
    .AddItem "2 Std."
    .ItemData(8) = 9
    .AddItem "5 Std."
    .ItemData(9) = 10
    .AddItem "10 Std."
    .ItemData(10) = 11
    .AddItem "1 Tag"
    .ItemData(11) = 12
    .AddItem "2 Tage"
    .ItemData(12) = 13
    .AddItem "5 Tage"
    .ItemData(13) = 14
End With

With CmPri
    .AddItem "Hoch"
    .ItemData(0) = 1
    .AddItem "Normal"
    .ItemData(1) = 2
    .AddItem "Niedrig"
    .ItemData(2) = 3
End With

With CmGes
    For AktZa = 0 To UBound(GlGes) - 1
        .AddItem GlGes(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

If CmETy.ListCount = 0 Then
    With CmETy
        For AktZa = 1 To UBound(GlKrA)
            If GlKrA(AktZa, 0) < 10 Then
                .AddItem GlKrA(AktZa, 1) & " - " & GlKrA(AktZa, 2)
                .ItemData(AktZa - 1) = GlKrA(AktZa, 0)
            End If
        Next AktZa
        If GlStS > 1 Then
            .ListIndex = 8
        Else
            .ListIndex = 1
        End If
    End With
End If

If GlOTS = True Then
    Lab17.Caption = "Bearbeitet :" 'Online-Terminbuchungs Sytem
End If

RetWe = SendMessage(CmRem.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(CmPri.hwnd, CB_SETCURSEL, 1, ByVal 0&)
RetWe = SendMessage(CmTyp.hwnd, CB_SETCURSEL, 2, ByVal 0&)
RetWe = SendMessage(CmGes.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(CmArz.hwnd, CB_SETCURSEL, 0, ByVal 0&)

If IdxNr > 0 Then
    If GlOTS = True Then 'Online-Terminbuchungs Sytem
        If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
            CmMit.Enabled = False
        Else
            CmMan.Enabled = False
        End If
    End If
    TxDa2.Enabled = False
    PuBu2.Enabled = False
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " TeLoa " & Err.Number
Resume Next

End Sub
Public Sub TeMain(ByVal TerNr As Long, Optional ByVal PatNr As Long, Optional ByVal PatNa As String, Optional ByVal TeBet As String, Optional ByVal TeKom As String, Optional ByVal MitNr As Long = 0, Optional ByVal ManNr As Long = 0)
On Error GoTo LaErr

Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmTermin") = True Then
    Set FM = frmTermin
    frmTermin.ZOrder 0
    Exit Sub
End If

GlTeF = True 'Formular wird geladen

TeReg

Load frmTermin

Set FM = frmTermin

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (780 / 2)
        .FeObn = (GlyGr / 2) - (560 / 2)
        .FeBre = 775
        .FeHoh = 560
    Else
        .FeLin = IniGetVal("Termin", "FenLin")
        .FeObn = IniGetVal("Termin", "FenObe")
        .FeBre = IniGetVal("Termin", "FenBre")
        .FeHoh = IniGetVal("Termin", "FenHoh")
    End If
End With

AFont FM
TeInit
TeMen
TeLoa TerNr
TeSpl

If TerNr = 0 Then
    TeNew PatNr, PatNa, TeBet, TeKom, MitNr, ManNr
ElseIf TerNr > 0 Then
    Ter_Lad TerNr
    Ter_Lei TerNr
    Ter_Com
    GlTSa = False
End If
DoEvents

With clFen
    .FenMov
    DoEvents
    Set CmBrs = FM.comBar02
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    TePos
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

frmTermin.Show

DoEvents
TMeAc False

DoEvents
GlTeF = False 'Formular wird geladen

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " TeMain " & Err.Number
Resume Next

End Sub
Private Sub TeMen()
On Error GoTo InErr
'Menue erstellen

Dim RetWe As Long
Dim KeyNa As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim RbTem As XtremeCommandBars.RibbonTab
Dim MsBar As XtremeCommandBars.MessageBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmTermin
Set TxID2 = FM.txtID2
Set CmBrs = FM.comBar02
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnDatu2
Set PuBu3 = FM.btnDatu3
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

KeyNa = "ToolTips"

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(AD_Termin_Close, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Delete, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Remind, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Notify, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Ketten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_StaKett, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_StaKet2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Abrechnen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_EintLoe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Clip1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Clip2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TE_Adresse_Ubertrag, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TE_Adresse_Bearbeit, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei01, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei02, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei03, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei04, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei05, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei06, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei07, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei08, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei09, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei10, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei11, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei12, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei13, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei14, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei15, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei16, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei17, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei18, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei19, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei20, vbNullString, vbNullString, vbNullString, vbNullString)
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Width = 120
    CmPan.Alignment = xtpAlignmentCenter
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Style = SBPS_STRETCH
    Set CmPan = .AddPane(3)
    CmPan.Width = 120
    CmPan.Text = "Terminfolge:"
    CmPan.Alignment = xtpAlignmentLeft
    .Visible = True
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmBuT = RbBar.Controls.Add(xtpControlButton, TE_Termin_Hilfe, "Hilfe")
With CmBuT
    .ToolTipText = "Öffnet die Kurzhilfe"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, TE_Termin_Beenden, "Schließen")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .ToolTipText = "Schließt den Dialog"
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

'________________________________________________________________________________

Set RbTab = RbBar.InsertTab(RibTab_Ter_Haupt, "Termindaten")
With RbTab
    .id = RibTab_Ter_Haupt
    .ToolTip = "Zeigt die Hauptdaten des Termins"
    .Visible = True
    .Selected = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Add, "Termin Hinzufügen")
With CmCon
    .IconId = IC32_Calendar_Add
    .ShortcutText = "F3"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Delete, "Termin Entfernen")
With CmCon
    .IconId = IC32_Calendar_Del
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Save, "Termin Speichern")
With CmCon
    .IconId = IC32_Disk_Calendar
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Close, "Termin Schließen")
With CmCon
    .IconId = IC32_Calendar_Copy
    .ShortcutText = "F8"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Adresse", RibGrp_TeR_Kopieren)
Set CmCon = RbGrp.Add(xtpControlButton, TE_Adresse_Hinzufu, "Adresse Hinzufügen")
With CmCon
    .Width = GlRib
    .IconId = IC32_Patient_Add
End With
Set CmCon = RbGrp.Add(xtpControlButton, TE_Adresse_Bearbeit, "Adresse Bearbeiten")
With CmCon
    .Width = GlRib
    .IconId = IC32_Patient_Edit
End With
Set CmCon = RbGrp.Add(xtpControlButton, TE_Adresse_Suchen, "Adresse Suchen")
With CmCon
    .Width = GlRib
    .IconId = IC32_Patient_View
    .ShortcutText = "F5"
End With

Set RbGrp = RbGps.AddGroup("Eigenschaften", RibGrp_Ter_Ansicht)
Set CmCon = RbGrp.Add(xtpControlCheckBox, AD_Termin_Remind, "Terminerinnerung")
Set CmCon = RbGrp.Add(xtpControlCheckBox, AD_Termin_Notify, "Emailerinnerung")
Set CmPop = RbGrp.Add(xtpControlPopup, TE_Farbe, "Terminfarbe")
With CmPop
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Paint 'IC16_Check
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, FaLei01, GlTmF(1, 0))
    CmCon.Checked = True
    Set CmCon = .Add(xtpControlButton, FaLei02, GlTmF(2, 0))
    Set CmCon = .Add(xtpControlButton, FaLei03, GlTmF(3, 0))
    Set CmCon = .Add(xtpControlButton, FaLei04, GlTmF(4, 0))
    Set CmCon = .Add(xtpControlButton, FaLei05, GlTmF(5, 0))
    Set CmCon = .Add(xtpControlButton, FaLei06, GlTmF(6, 0))
    Set CmCon = .Add(xtpControlButton, FaLei07, GlTmF(7, 0))
    Set CmCon = .Add(xtpControlButton, FaLei08, GlTmF(8, 0))
    Set CmCon = .Add(xtpControlButton, FaLei09, GlTmF(9, 0))
    Set CmCon = .Add(xtpControlButton, FaLei10, GlTmF(10, 0))
    Set CmCon = .Add(xtpControlButton, FaLei11, GlTmF(11, 0))
    Set CmCon = .Add(xtpControlButton, FaLei12, GlTmF(12, 0))
    Set CmCon = .Add(xtpControlButton, FaLei13, GlTmF(13, 0))
    Set CmCon = .Add(xtpControlButton, FaLei14, GlTmF(14, 0))
    Set CmCon = .Add(xtpControlButton, FaLei15, GlTmF(15, 0))
    Set CmCon = .Add(xtpControlButton, FaLei16, GlTmF(16, 0))
    Set CmCon = .Add(xtpControlButton, FaLei17, GlTmF(17, 0))
    Set CmCon = .Add(xtpControlButton, FaLei18, GlTmF(18, 0))
    Set CmCon = .Add(xtpControlButton, FaLei19, GlTmF(19, 0))
    Set CmCon = .Add(xtpControlButton, FaLei20, GlTmF(20, 0))
End With

'________________________________________________________________________________

Set RbTab = RbBar.InsertTab(RibTab_Ter_Adres, "Terminadresse")
With RbTab
    .id = RibTab_Ter_Adres
    .ToolTip = "Zeigt die im Termin eingebettete Adresse"
    .Visible = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButtonPopup, SY_TE_Termin_Docume, "Termin Nachricht")
With CmCon
    .IconId = IC32_Calendar_Phone
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_EmlBes, "Email-Bestätigung")
    CmCon.IconId = IC16_Earth_Mail
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_EmlEri, "Email-Erinnerung")
    CmCon.IconId = IC16_Earth_Mail
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_EmlVrs, "Email-Vorschlag")
    CmCon.IconId = IC16_Earth_Mail
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_EmlAbs, "Email-Absage")
    CmCon.IconId = IC16_Earth_Mail
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_EmlSto, "Email-Storno")
    CmCon.IconId = IC16_Earth_Mail
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_SMSBes, "SMS-Bestätigung")
    CmCon.IconId = IC16_Phone_Mobil
    CmCon.BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_SMSEri, "SMS-Erinnerung")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_SMSVrs, "SMS-Vorschlag")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_SMSAbs, "SMS-Absage")
    CmCon.IconId = IC16_Phone_Mobil
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_TE_Termin_SMSSto, "SMS-Storno")
    CmCon.IconId = IC16_Phone_Mobil
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Delete, "Termin Entfernen")
With CmCon
    .IconId = IC32_Calendar_Del
    .Width = GlRib
End With

Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Save, "Termin Speichern")
With CmCon
    .IconId = IC32_Disk_Calendar
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Close, "Termin Schließen")
With CmCon
    .IconId = IC32_Calendar_Copy
    .ShortcutText = "F8"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Adresse", RibGrp_TeR_Kopieren)
Set CmCon = RbGrp.Add(xtpControlButton, TE_Adresse_Ubertrag, "Adresse Importieren")
With CmCon
    .Width = GlRib
    .IconId = IC32_Patient_Copy
End With
Set CmCon = RbGrp.Add(xtpControlButton, TE_Adresse_Bearbeit, "Adresse Bearbeiten")
With CmCon
    .Width = GlRib
    .IconId = IC32_Patient_Edit
End With
Set CmCon = RbGrp.Add(xtpControlButton, TE_Adresse_Suchen, "Adresse Suchen")
With CmCon
    .Width = GlRib
    .IconId = IC32_Patient_View
    .ShortcutText = "F5"
End With

'________________________________________________________________________________

Set RbTab = RbBar.InsertTab(RibTab_Ter_Leist, "Leistungsziffern")
With RbTab
    .id = RibTab_Ter_Leist
    .ToolTip = "Zeigt die zugeordneten Leistungen"
    .Visible = True
    If GlRch(0, 9) = 0 Then .Enabled = False
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Abrechnen, "Rechnung Erstellen")
With CmCon
    .IconId = IC32_Mail_Export
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_EintLoe, "Einträge Entfernen")
With CmCon
    .IconId = IC32_Doc_Del
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Save, "Termin Speichern")
With CmCon
    .IconId = IC32_Disk_Calendar
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Close, "Termin Schließen")
With CmCon
    .IconId = IC32_Calendar_Copy
    .ShortcutText = "F8"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Leistungen", RibGrp_TeR_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Ketten, "Leistungen Auswählen")
With CmCon
    .IconId = IC32_Folder_View
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt1, "Terminbetrag :")
CmCon.flags = xtpFlagRightAlign
Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt2, "Serienbetrag :")
CmCon.flags = xtpFlagRightAlign
Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt3, "Beglichen :")
CmCon.flags = xtpFlagRightAlign
Set CmEdi = RbGrp.Add(xtpControlEdit, AD_Termin_Betrag1, vbNullString)
With CmEdi
    .EditStyle = xtpEditStyleRight
    .Style = xtpButtonIconAndCaption
    .Width = 70
    .Text = GlWa2
End With
Set CmEdi = RbGrp.Add(xtpControlEdit, AD_Termin_Betrag2, vbNullString)
With CmEdi
    .EditStyle = xtpEditStyleRight
    .Style = xtpButtonIconAndCaption
    .Width = 70
    .Text = GlWa2
End With
Set CmEdi = RbGrp.Add(xtpControlEdit, AD_Termin_Betrag3, vbNullString)
With CmEdi
    .EditStyle = xtpEditStyleRight
    .Style = xtpButtonIconAndCaption
    .Width = 70
    .Text = GlWa2
End With

Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, AD_Termin_StaKett, "Standardkette Einfügen")
With CmCon
    .IconId = IC32_Link_Down
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Termin_StaKett, "Standardkette 1 Einfügen")
    CmCon.IconId = IC16_Link_Norm
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, AD_Termin_StaKet2, "Standardkette 2 Einfügen")
    CmCon.IconId = IC16_Link_Norm
    .Width = GlRib
    .BeginGroup = True
End With

'________________________________________________________________________________

Set RbTab = RbBar.InsertTab(RibTab_Ter_WarZi, "Terminwarteliste")
With RbTab
    .id = RibTab_Ter_WarZi
    .ToolTip = "Zeigt die Patienten, die auf einen freien Termin warten"
    .Visible = True
    .Selected = False
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_WarSet, "Wartenden Übernehmen")
With CmCon
    .IconId = IC32_Calendar_Phone
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_WarNeu, "Wartenden Hinzufügen")
With CmCon
    .IconId = IC32_Patient_Add
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_WarDel, "Wartenden Entfernen")
With CmCon
    .IconId = IC32_Patient_Del
    .Width = GlRib
End With

'________________________________________________________________________________

Set RbTab = RbBar.InsertTab(RibTab_Ter_Proto, "Terminprotokoll")
With RbTab
    .id = RibTab_Ter_Proto
    .ToolTip = "Zeigt ein Protokoll über alle Eriegnisse zu diesem Termin"
    .Visible = True
    .Selected = False
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_TermID, "Protokoll Exportieren")
With CmCon
    .IconId = IC32_Book_Export
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_ProEmail, "Protokoll Emailversand")
With CmCon
    .IconId = IC32_Book_Copy
    .Width = GlRib
End With

'________________________________________________________________________________

Set CmCoS = RbBar.Controls
For Each CmCon In CmCoS
    CmCon.ToolTipText = IniGetOpt(KeyNa, CmCon.id)
Next CmCon

'________________________________________________________________________________

With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
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
    .SetIconSize True, 32, 32
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
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case 8:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case Else:
        If GlRah = True Then 'Office EnableThemeframe
            .VisualTheme = xtpThemeRibbon
        Else
            If GlFRg = True Then 'farbige Register
                .VisualTheme = xtpThemeResource
            Else
                .VisualTheme = xtpThemeRibbon
            End If
        End If
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End Select
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = True
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
    .KeyBindings.Add 0, VK_F3, KY_F3
    .KeyBindings.Add 0, VK_F4, KY_F4
    .KeyBindings.Add 0, VK_F5, KY_F5
    .KeyBindings.Add 0, VK_F6, KY_F6
    .KeyBindings.Add 0, VK_F7, KY_F7
    .KeyBindings.Add 0, VK_F8, KY_F8
    .KeyBindings.Add 0, VK_F9, KY_F9
    .KeyBindings.Add 0, VK_F10, KY_F10
    .KeyBindings.Add 0, VK_F11, KY_F11
    .KeyBindings.Add FCONTROL, Asc("Z"), AD_Termin_Clip1
    .KeyBindings.Add FCONTROL, Asc("R"), AD_Termin_Clip2
End With

Set RbBar = CmBrs.Item(1)
With RbBar
    .AllowMinimize = False
    .AllowQuickAccessCustomization = False
    .AllowQuickAccessDuplicates = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .EnableAnimation = GlMeA
    .FontHeight = GlToF
    .GroupsVisible = True
    .MinimumVisibleWidth = 100
    .RibbonPaintManager.HotTrackingGroups = True
    .RibbonPaintManager.CaptionFont.SIZE = 8
    .RibbonPaintManager.CaptionFont.Name = GlTFt.Name
    .RibbonPaintManager.WindowCaptionFont.SIZE = 8
    .RibbonPaintManager.WindowCaptionFont.Name = GlTFt.Name
    .ShowQuickAccess = False
    .ShowQuickAccessBelowRibbon = False
    .ShowCaptionAlways = True
    .Position = xtpBarTop
    .SetIconSize 16, 16
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
            .TabPaintManager.Appearance = xtpTabAppearanceVisualStudio2005
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 6
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
        Else
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        End If
    End Select
    .TabPaintManager.Layout = xtpTabLayoutAutoSize
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.MinTabWidth = 100
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ClientFrame = xtpTabFrameNone
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = False
    .TabPaintManager.HotTracking = True
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.Font.SIZE = 8
    .TabPaintManager.Font.Name = GlTFt.Name
    If GlRDP = True Then
        .EnableFrameTheme
    Else
        If GlRah = True Then
            .EnableFrameTheme
        End If
    End If
End With

If GlTem = 0 Then
    CmAcs(AD_Termin_Ketten).Enabled = False
    CmAcs(AD_Termin_StaKett).Enabled = False
End If

CmAcs(AD_Termin_Abrechnen).Enabled = False
CmAcs(AD_Termin_EintLoe).Enabled = False

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu3.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Set CmPan = Nothing
Set CmSta = Nothing
Set CmPop = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set RbBar = Nothing
Set RbTab = Nothing
Set RbGrp = Nothing
Set RbGps = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " TeMen " & Err.Number
Resume Next

End Sub
Public Sub TeNach(ByVal DoTyp As Integer)
On Error GoTo AdErr
'Erstellt ein Dokuemnt für Terminbestätigungen

Dim DaSta As Date
Dim DaEnd As Date
Dim DaAdd As Date
Dim ZeSta As Date
Dim ZeEnd As Date
Dim TerNr As Long
Dim TeMit As Long
Dim TeMan As Long
Dim PatNr As Long
Dim EmEmp As String
Dim EmTex As String
Dim EmBCC As String
Dim TeGui As String
Dim TelMo As String
Dim TelPr As String
Dim TeStr As String
Dim EmBet As String
Dim TeBet As String
Dim TeOrt As String
Dim MaStr As String
Dim MiNam As String
Dim PaStr As String
Dim DaStL As String
Dim DaStK As String
Dim DaStN As String
Dim ZeStS As String
Dim ZeStE As String
Dim TmStr As String
Dim FiNam As String
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim GnzTa As Boolean
Dim RetWe As Boolean
Dim TePri As Integer
Dim AktZa As Integer
Dim DoGrp As Integer
Dim GesZa As Integer
Dim Mld1, Tit1 As String

Set FM = frmTermin
Set TxID2 = FM.txtID2
Set TxID0 = FM.txtID0
Set TxGui = FM.txtGuiID
Set TxAdr = FM.txtAdres
Set CmBet = FM.txtBetre
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set CmPri = FM.cmbPrior
Set CmGan = FM.cmbGanzt
Set TxOrt = FM.txtRaum1
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set TxDa3 = FM.txtDatum
Set TxZeV = FM.txtVonZe
Set TxZeB = FM.txtBisZe
Set CmBrf = FM.cmbS4F11

Set clICS = New clsICS
Set clFil = New clsFile
clFil.hwnd = FM.hwnd

GlTrM = False 'Terminbearbeitung direkt im Kalender
GlTrB = False 'Terminbearbeitung direkt im Kalender

If TxID2.Text <> vbNullString Then
    If IsNumeric(TxID2.Text) Then
        TerNr = CLng(TxID2.Text)
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If

If TerNr > 0 Then
    TeMan = CmMan.ItemData(CmMan.ListIndex)
    TeMit = CmMit.ItemData(CmMit.ListIndex)
    TePri = CmPri.ListIndex + 1
    TeGui = TxGui.Text
    If IsDate(TxDa1.Text) = True Then
        DaSta = CDate(TxDa1.Text)
    Else
        DaSta = Date
    End If
    If IsDate(TxDa2.Text) = True Then
        DaEnd = CDate(TxDa2.Text)
    Else
        DaEnd = Date
    End If
    If IsDate(TxDa3.Text) = True Then
        DaAdd = CDate(TxDa3.Text)
    Else
        DaAdd = Date
    End If
    ZeSta = TimeValue(TxZeV.Text)
    ZeEnd = TimeValue(TxZeB.Text)
    If TxAdr.Text <> vbNullString Then
        PaStr = TxAdr.Text
    End If
    If TxOrt.Text <> vbNullString Then
        TeOrt = TxOrt.Text
    End If
    If CmBet.Text <> vbNullString Then
        TeBet = CmBet.Text
    End If
    If CmGan.ListIndex = 0 Then
        GnzTa = True
    End If
    If CmMan.Text <> vbNullString Then
        MaStr = CmMan.Text
    End If
    
    If TxID0.Text <> vbNullString Then
        If IsNumeric(TxID0.Text) Then
            PatNr = CLng(TxID0.Text)
        Else
            PatNr = 0
        End If
    Else
        PatNr = 0
    End If

    DaStL = Format$(DaSta, "dddd" & ", " & "dd" & ". " & "mmmm" & Chr$(32) & "yyyy")
    DaStK = Format$(DaSta, "ddd" & ", " & "dd" & ". " & "mmm" & Chr$(32) & "yyyy")
    DaStN = Format$(DaSta, "dd.mm.yyyy")
    ZeStS = Format$(ZeSta, "hh:mm")
    ZeStE = Format$(ZeEnd, "hh:mm")
    
    For AktZa = 1 To UBound(GlMiK)
        If TeMit = GlMiK(AktZa, 2) Then
            If GlMiK(AktZa, 27) <> vbNullString Then
                MiNam = GlMiK(AktZa, 27) 'OTS Name
            Else
                If GlMiK(AktZa, 4) <> vbNullString Then 'Vorname
                    If GlMiK(AktZa, 3) <> vbNullString Then 'Name
                        MiNam = GlMiK(AktZa, 4) & " " & GlMiK(AktZa, 3)
                    Else
                        MiNam = GlMiK(AktZa, 1) 'Kurzbezeichnung
                    End If
                Else
                    MiNam = GlMiK(AktZa, 1) 'Kurzbezeichnung
                End If
            End If
            Exit For
        End If
    Next AktZa

    With GlTxV
        If GlEmN(DoTyp, 1) <> vbNullString Then
            .TxStr = GlEmN(DoTyp, 1)
        Else
            .TxStr = "Termin"
        End If
         If DoTyp < 6 Then
            .DaStr = DaStL
        ElseIf DoTyp < 10 Then
            .DaStr = DaStK
        Else
            .DaStr = DaStN
        End If
        .Datum = DaSta
        .MitNr = TeMit
        .ManNr = TeMan
        .PaStr = PaStr
        .ZeiSt = ZeStS
        .ZeiEn = ZeStE
        .TerID = TeGui
        .PatNr = PatNr
    End With
    TeStr = SEmTx()
    
    If DoTyp > 5 Then
        With GlTxV
            If GlEmN(DoTyp, 3) <> vbNullString Then
                .TxStr = GlEmN(DoTyp, 3)
            Else
                .TxStr = "Termin " & MiNam
            End If
            .MitNr = TeMit
            .ManNr = TeMan
            .PatNr = PatNr
        End With
        EmBet = SEmTx()
    End If
    
    Select Case DoTyp
    Case 6:
        DoGrp = 2
    Case 7:
        DoGrp = 2
    Case 8:
        DoGrp = 2
    Case 9:
        DoGrp = 2
    Case 10:
        DoGrp = 3
        TeStr = SNaFi(TeStr)
    Case 11:
        DoGrp = 3
        TeStr = SNaFi(TeStr)
    Case 12:
        DoGrp = 3
        TeStr = SNaFi(TeStr)
    Case 13:
        DoGrp = 3
        TeStr = SNaFi(TeStr)
    Case 15:
        DoGrp = 2
    Case 16:
        DoGrp = 3
        TeStr = SNaFi(TeStr)
    End Select

    Select Case DoGrp
    Case 2: 'Email
        If GlEKV = False Then 'Emailkonten vorhanden
            TeTit = "E-Mail-Versand"
            TeMai = "Es ist kein E-Mail-Konto vorhanden"
            TeInh = "Um E-Mails versenden zu können, ist es notwendig mind. ein E-Mail-Konto hinzuzufügen."
            TeFus = "Um ein E-Mail-Konto hinzuzufügen, wechseln Sie in das Modul: Textverarbeitung und dann oben auf Emails. Dort klicken Sie auf die Schaltfläche Emailkonten."
            SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, True, FM.hwnd
            Exit Sub
        End If
    
        EmEmp = FM.txtS4F16.Text
        EmTex = CmBrf.Text & vbCrLf & vbCrLf & TeStr
        
        If GlTeB = True Then 'Terminnachricht auch an BCC
            If GlMkt(1, 13) <> vbNullString Then
                EmBCC = GlMkt(1, 13)
            Else
                If GlThe(GlSMa, 16) <> vbNullString Then
                    EmBCC = GlThe(GlSMa, 16)
                End If
            End If
        End If

        If GlICS = True Then 'Terminnachricht mit ICS Dateiversand
            DoEvents
            With clICS
                .moKaNa = MiNam
                .moTeID = TeGui
                .moDaAd = DaAdd
                .moDaSt = DaSta
                .moDaEn = DaEnd
                .moZeSt = ZeSta
                .moZeEn = ZeEnd
                .moTeBe = EmBet
                .moTeGa = GnzTa
                .moTeKo = Left$(TeStr, 62)
                .moTePr = TePri
                TmStr = .ICSHead
                TmStr = TmStr & .ICSBody
                TmStr = TmStr & .ICSFoot
            End With

            FiNam = GlTEx & "Termin_" & Format$(Now, "YYYYMMDD_HHMMSS") & ".ics" 'Termineordner
            With clFil
                If .FilVor(FiNam) = True Then
                    .DaLoe = FiNam & vbNullChar
                    .FilLoe
                End If
                Call .FilCnWr(FiNam, TmStr)
            End With

            DoEvents
            SMaNe PatNr, EmEmp, EmBCC, EmTex, EmBet, FiNam
        Else
            DoEvents
            SMaNe PatNr, EmEmp, EmBCC, EmTex, EmBet
        End If
    
    Case 3: 'SMS
                
        If PatNr > 0 Then
            TelMo = S_AdIdx(PatNr, "Telefon4")
        Else
            TelMo = FM.txtS4F15.Text
        End If
        If TelMo <> vbNullString Then
            frmSMS.NaTex = TeStr
            frmSMS.NaNum = TelMo
            frmSMS.DoTyp = DoTyp
            frmSMS.Show vbModal
        Else
            SPopu "Keine Rufnummer", "Für den Patienten muss eine Rufnummer vorhanden sein", IC48_Forbidden
        End If
    
    End Select
    
    If TerNr > 0 Then
        DBCmEx2 "qryTerOnTe", "@OnlTe", "@IdxNr", -1, TerNr
    End If
    
Else
    SPopu "Kein Termin ausgewählt", "Sie müssen erst einen Termine auswählen", IC48_Forbidden
End If

Set clICS = Nothing
Set clFil = Nothing

Exit Sub

AdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TeNach " & Err.Number
Resume Next

End Sub

Public Sub TeNew(Optional ByVal PatNr As Long, Optional ByVal PatNa As String, Optional ByVal TeBet As String, Optional ByVal TeKom As String, Optional MiNum As Long = 0, Optional MaNum As Long = 0)
On Error GoTo NeErr
'Bereitet die Neueingabe eines Termines vor

Dim RetWe As Long
Dim MitNr As Long
Dim ManNr As Long
Dim AdPIN As Long
Dim MasTe As Long
Dim EndZe As Date
Dim ZeStr As String
Dim TagWe As String
Dim AkDat As String
Dim BrStr As String
Dim Telef As String
Dim NotDa As String
Dim NotZe As String
Dim NotSt As String
Dim StaZe As String
Dim AktZa As Integer
Dim FltTy As Integer 'Filtertyp
Dim FltId As Integer
Dim MiDif As Integer
Dim SelDf As Integer
Dim ZeiRa As Integer
Dim MiIdx As Integer
Dim MaIdx As Integer
Dim NotVa As Integer
Dim MitOK As Boolean
Dim ManOK As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCo1 As XtremeCommandBars.CommandBarComboBox
Dim CmCo2 As XtremeCommandBars.CommandBarComboBox
Dim DaPi1 As XtremeCalendarControl.DatePicker

Set FM = frmTermin
Set TxID0 = FM.txtID0
Set TxID2 = FM.txtID2
Set TxIDS = FM.txtIdSer
Set CmBet = FM.txtBetre
Set TxOrt = FM.txtRaum1
Set TxAdr = FM.txtAdres
Set TxFar = FM.txtFarbe
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set TxDa3 = FM.txtRzDat
Set TxRzn = FM.txtRzNum
Set TxRzA = FM.txtRzAnz
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set TxNoS = FM.txtNoSta
Set TxNoD = FM.txtNoDat
Set TxNoZ = FM.txtNoTim
Set TxRef = FM.txtRefNr
Set CmRmu = FM.cmbRaum1
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set CmArz = FM.cmbArzNr
Set CmAbr = FM.cmbAbger
Set CmPri = FM.cmbPrior
Set CmAbg = FM.cmbAbgeh
Set CmGan = FM.cmbGanzt
Set CmNot = FM.cmbNotVa
Set CmOnT = FM.cmbOnlTe
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set DaPi1 = frmMain.dtpDatu7

GlTeN = True
GlTSa = False

CmAcs(AD_Termin_Ketten).Enabled = False
CmAcs(AD_Termin_StaKett).Enabled = False
CmAcs(AD_Termin_Abrechnen).Enabled = False
CmAcs(AD_Termin_EintLoe).Enabled = False
CmAcs(TE_Adresse_Ubertrag).Enabled = False
CmAcs(TE_Adresse_Bearbeit).Enabled = False

Set CmCo1 = frmMain.comBar01.FindControl(CmCo1, SY_TE_Termin_FiltTyp, , True)
Set CmCo2 = frmMain.comBar01.FindControl(CmCo2, SY_TE_Termin_FiltIdx, , True)

FltTy = CmCo1.ListIndex 'Filtertyp 1=Standard 2=Raum 3=Mitarbeiter
FltId = CmCo2.ListIndex - 1

Select Case GlBut
Case RibTab_Startseite:
        AkDat = Date
Case RibTab_Adressen:
        AkDat = Date
Case RibTab_Mandanten:
        AkDat = Date
Case RibTab_Mitarbeit:
        AkDat = Date
Case RibTab_Verordner:
        AkDat = Date
Case RibTab_Ter_Listen:
        AkDat = Date
Case RibTab_Ter_Akont:
        AkDat = Date
Case RibTab_Ter_Warte:
        AkDat = Date
Case Else:
        If Format$(GlSel.DaSta, "hh:mm:ss") = "00:00:00" Then 'Markierte Celle im Kalender
            If Not IsDate(Format$(GlSel.DaSta, "dd.mm.yyyy")) Then
                GlSel.DaSta = Format$(Now, "dd.mm.yyyy") & Chr$(32) & Format$(Now, "hh:mm:ss")
                AkDat = Format$(Now, "dd.mm.yyyy")
            Else
                If GlSel.DaSta = "00:00:00" Then
                    AkDat = Date
                Else
                    AkDat = Format$(GlSel.DaSta, "dd.mm.yyyy")
                End If
            End If
        Else
            AkDat = Format$(GlSel.DaSta, "dd.mm.yyyy")
        End If
End Select

MasTe = Ter_VoT()

If PatNr > 0 Then
    TxID0.Text = PatNr
    TagWe = Mid$(TxID0.Tag, 2, Len(TxID0.Tag) - 1)
    TxID0.Tag = 1 & TagWe
    
    TxAdr.Text = PatNa
    TagWe = Mid$(TxAdr.Tag, 2, Len(TxAdr.Tag) - 1)
    TxAdr.Tag = 1 & TagWe

    S_AdDe PatNr 'Adressendetails
    With GlADt
        If .AdTe1 <> vbNullString Then
            Telef = .AdTe1
        ElseIf .AdTe2 <> vbNullString Then
            Telef = .AdTe2
        ElseIf .AdTe4 <> vbNullString Then
            Telef = .AdTe4
        End If

        FM.txtS4F01.Text = .AdFir
        TagWe = Mid$(FM.txtS4F01.Tag, 2, Len(FM.txtS4F01.Tag) - 1)
        FM.txtS4F01.Tag = 1 & TagWe

        FM.txtS4F02.Text = .AdAnr
        TagWe = Mid$(FM.txtS4F02.Tag, 2, Len(FM.txtS4F02.Tag) - 1)
        FM.txtS4F02.Tag = 1 & TagWe
            
        FM.txtS4F03.Text = .AdTit
        TagWe = Mid$(FM.txtS4F03.Tag, 2, Len(FM.txtS4F03.Tag) - 1)
        FM.txtS4F03.Tag = 1 & TagWe
        
        FM.txtS4F04.Text = .AdVor
        TagWe = Mid$(FM.txtS4F04.Tag, 2, Len(FM.txtS4F04.Tag) - 1)
        FM.txtS4F04.Tag = 1 & TagWe
            
        FM.txtS4F05.Text = .AdNam
        TagWe = Mid$(FM.txtS4F05.Tag, 2, Len(FM.txtS4F05.Tag) - 1)
        FM.txtS4F05.Tag = 1 & TagWe
            
        FM.txtS4F06.Text = .AdStr
        TagWe = Mid$(FM.txtS4F06.Tag, 2, Len(FM.txtS4F06.Tag) - 1)
        FM.txtS4F06.Tag = 1 & TagWe
            
        FM.txtS4F08.Text = .AdPLZ
        TagWe = Mid$(FM.txtS4F08.Tag, 2, Len(FM.txtS4F08.Tag) - 1)
        FM.txtS4F08.Tag = 1 & TagWe
            
        FM.txtS4F09.Text = .AdOrt
        TagWe = Mid$(FM.txtS4F09.Tag, 2, Len(FM.txtS4F09.Tag) - 1)
        FM.txtS4F09.Tag = 1 & TagWe
    
        FM.cmbS4F12.Text = .AdLan
        TagWe = Mid$(FM.cmbS4F12.Tag, 2, Len(FM.cmbS4F12.Tag) - 1)
        FM.cmbS4F12.Tag = 1 & TagWe
    
        FM.txtS4F18.Text = .AdGeb
        TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
        FM.txtS4F18.Tag = 1 & TagWe

        FM.txtS4F15.Text = Telef
        TagWe = Mid$(FM.txtS4F15.Tag, 2, Len(FM.txtS4F15.Tag) - 1)
        FM.txtS4F15.Tag = 1 & TagWe
    
        FM.txtS4F16.Text = .AdTe5
        TagWe = Mid$(FM.txtS4F16.Tag, 2, Len(FM.txtS4F16.Tag) - 1)
        FM.txtS4F16.Tag = 1 & TagWe
    
        BrStr = .AdBrf
        Ter_Brz BrStr
        TagWe = Mid$(FM.txtS4F18.Tag, 2, Len(FM.txtS4F18.Tag) - 1)
        FM.txtS4F18.Tag = 1 & TagWe
        
        AdPIN = .AdPIN
        If AdPIN = 0 Then AdPIN = Adr_Let()
        FM.txtS4F20.Text = Format$(AdPIN, "000000")
    End With

    GlTSa = True
Else
    TxID0.Text = 0
    CmBet.Text = vbNullString
    TxAdr.Text = vbNullString
    AdPIN = Adr_Let()
    FM.txtS4F20.Text = Format$(AdPIN, "000000")
End If

If TeBet <> vbNullString Then
    CmBet.Text = TeBet
    TagWe = Mid$(CmBet.Tag, 2, Len(CmBet.Tag) - 1)
    CmBet.Tag = 1 & TagWe
    GlTSa = True
End If

If TeKom <> vbNullString Then
    FM.txtKomme.Text = TeKom
    TagWe = Mid$(FM.txtKomme.Tag, 2, Len(FM.txtKomme.Tag) - 1)
    FM.txtKomme.Tag = 1 & TagWe
    GlTSa = True
End If

TxIDS.Text = 5
TxDa1.Text = AkDat
TxDa2.Text = AkDat
TxDa3.Text = AkDat

If GlTeO = True Then 'Mitarbeitername in Terminort
    TxOrt.Text = GlMiA(GlSmI, 1)
    TagWe = Mid$(TxOrt.Tag, 2, Len(TxOrt.Tag) - 1)
    TxOrt.Tag = 1 & TagWe
    TxOrt.Enabled = False
End If

TxFar.Text = 1
TagWe = Mid$(TxFar.Tag, 2, Len(TxFar.Tag) - 1)
TxFar.Tag = 1 & TagWe

If FltTy = 2 Then 'Filtertyp Raum
    CmRmu.ListIndex = FltId
Else
    CmRmu.ListIndex = GlTRx
End If

TagWe = Mid$(CmRmu.Tag, 2, Len(CmRmu.Tag) - 1)
CmRmu.Tag = 1 & TagWe

Select Case GlBut
Case RibTab_Ter_Kalend:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        If FltTy = 3 Then 'WICHTIG Filtertyp Mitarbeiter
            CmMit.ListIndex = FltId
            MiIdx = FltId + 1
        Else
            If MiNum > 0 Then
                For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                    If MiNum = GlMiT(AktZa, 2) Then
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        MitOK = True
                        Exit For
                    End If
                Next AktZa
            Else
                For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                    If GlTBn = GlMiT(AktZa, 2) Then 'Termin Behandlernummer
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        MitOK = True
                        Exit For
                    End If
                Next AktZa
            End If
            If MitOK = True Then
                CmMit.ListIndex = GlTBx
                MiIdx = GlTBx + 1
            Else
                CmMit.ListIndex = 0
                MiIdx = 1
            End If
        End If
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(MiIdx, 39)
    Else
        If FltTy = 3 Then 'WICHTIG Filtertyp Mandant
            CmMan.ListIndex = FltId
            MaIdx = FltId + 1
        Else
            If MaNum > 0 Then
                For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                    If MaNum = GlMaT(AktZa, 2) Then
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        ManOK = True
                        Exit For
                    End If
                Next AktZa
            Else
                For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                    If GlTBn = GlMaT(AktZa, 2) Then 'Termin Behandlernummer
                        GlTBx = AktZa - 1 'Termin Behandlerindex
                        ManOK = True
                        Exit For
                    End If
                Next AktZa
            End If
            If ManOK = True Then
                CmMan.ListIndex = GlTBx
                MaIdx = GlTBx + 1
            Else
                CmMan.ListIndex = 0
                MaIdx = 1
            End If
        End If
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(MaIdx, 25)
    End If
    
Case RibTab_Ter_Raeume:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        If MiNum > 0 Then
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If MiNum = GlMiT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If GlTBn = GlMiT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If MitOK = False Then
            GlTBx = 0
        End If

        CmMit.ListIndex = GlTBx
        If GlTRa = True Then 'Mitarbeiter / Mandanten in Raumbelegung aktivieren
            If GlMiT(GlSmI + 1, 8) > 0 Then
                ZeiRa = GlMiT(GlSmI + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        Else
            If GlMiT(GlTBx + 1, 8) > 0 Then
                ZeiRa = GlMiT(GlTBx + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(AktZa, 39)
    Else
        If MaNum > 0 Then
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If MaNum = GlMaT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If GlTBn = GlMaT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If ManOK = False Then
            GlTBx = 1
        End If
        CmMan.ListIndex = GlTBx
        If GlTRa = True Then 'Mitarbeiter / Mandanten in Raumbelegung aktivieren
            If GlMaT(GlSMa + 1, 8) > 0 Then
                ZeiRa = GlMaT(GlSMa + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        Else
            If GlMaT(GlTBx + 1, 8) > 0 Then
                ZeiRa = GlMaT(GlTBx + 1, 8)
            Else
                ZeiRa = GlZeR 'Zeitrasterindex
            End If
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(AktZa, 25)
    End If
    
Case RibTab_Ter_Mitarb:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        MitNr = GlMiT(GlTBx + 1, 2)
        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
            If MitNr = GlMiT(AktZa, 2) Then
                Exit For
            End If
        Next AktZa
        CmMit.ListIndex = AktZa - 1
        If GlMiT(GlTBx + 1, 8) > 0 Then
            ZeiRa = GlMiT(GlTBx + 1, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(AktZa, 39)
    Else
        ManNr = GlMaT(GlTBx + 1, 2)
        For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
            If ManNr = GlMaT(AktZa, 2) Then
                Exit For
            End If
        Next AktZa
        CmMan.ListIndex = AktZa - 1
        If GlMaT(GlTBx + 1, 8) > 0 Then
            ZeiRa = GlMaT(GlTBx + 1, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(AktZa, 25)
    End If
    
Case Else:

    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        If MiNum > 0 Then
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If MiNum = GlMiT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
                If GlTBn = GlMiT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    MitOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If MitOK = True Then
            CmMit.ListIndex = GlTBx
            MiIdx = GlTBx + 1
        Else
            CmMit.ListIndex = GlSmI - 1
            MiIdx = 1
        End If
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe
        NotVa = GlMiT(AktZa, 39)
    Else
        If MaNum > 0 Then
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If MaNum = GlMaT(AktZa, 2) Then
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
                If GlTBn = GlMaT(AktZa, 2) Then 'Termin Behandlernummer
                    GlTBx = AktZa - 1 'Termin Behandlerindex
                    ManOK = True
                    Exit For
                End If
            Next AktZa
        End If
        If ManOK = True Then
            CmMan.ListIndex = GlTBx
            MaIdx = GlTBx + 1
        Else
            CmMan.ListIndex = GlSMa - 1
            MaIdx = 1
        End If
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
        NotVa = GlMaT(AktZa, 25)
    End If

End Select

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMan.ListCount > 1 Then
        ManNr = 0
        MitNr = CmMit.ItemData(CmMit.ListIndex)

        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
            If MitNr = CLng(GlMiT(AktZa, 2)) Then
                ManNr = GlMiT(AktZa, 7) 'zugeordnete Mandantennummer
                Exit For
            End If
        Next AktZa

        If ManNr > 0 Then
            For AktZa = 1 To UBound(GlMaA)  'Aktive Mandanten
                If ManNr = CLng(GlMaA(AktZa, 2)) Then
                    Exit For
                End If
            Next AktZa
            CmMan.ListIndex = AktZa - 1
        Else
            CmMan.ListIndex = GlSMa - 1
        End If
    Else
        CmMan.ListIndex = GlSMa - 1
    End If
    TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
    CmMan.Tag = 1 & TagWe
Else
    If CmMit.ListCount > 1 Then
        CmMit.ListIndex = GlSmI - 1
    Else
        CmMit.ListIndex = GlSmI - 1
    End If
    TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
    CmMit.Tag = 1 & TagWe
End If

CmArz.ListIndex = 0
TagWe = Mid$(CmArz.Tag, 2, Len(CmArz.Tag) - 1)
CmArz.Tag = 1 & TagWe

MiDif = GlTku(ZeiRa, 2)
DoEvents
SRast ZeiRa
DoEvents

If GlSel.DaSta > 0 Then
    If Format$(GlSel.DaSta, "hh:mm") = "00:00" Then
        If MiDif = 0 Then MiDif = 15
        VoZei.Text = "08:00"
        EndZe = DateAdd("n", MiDif, "08:00:00")
        BiZei.Text = Format$(EndZe, "hh:mm")
    Else
        ZeStr = Format$(GlSel.DaSta, "hh:mm")
        For AktZa = 1 To UBound(GlRas) 'Zeitrasterstartzeiten
            If TimeValue(GlRas(AktZa)) <= TimeValue(ZeStr) Then
                StaZe = TimeValue(GlRas(AktZa))
            End If
        Next AktZa
        If GlSSt = True Then 'Starre Termintaktung
            EndZe = DateAdd("n", MiDif, StaZe)
            If TimeValue(GlSel.DaEnd) > TimeValue(EndZe) Then
                EndZe = Format$(GlSel.DaEnd, "hh:mm")
            End If
        Else
            SelDf = DateDiff("n", Format$(StaZe, "hh:mm"), Format$(GlSel.DaEnd, "hh:mm"))
            If SelDf < MiDif Then
                EndZe = DateAdd("n", MiDif, StaZe)
            Else
                EndZe = Format$(GlSel.DaEnd, "hh:mm")
            End If
        End If
        VoZei.Text = Format$(StaZe, "hh:mm")
        BiZei.Text = Format$(EndZe, "hh:mm")
    End If
Else
    If MiDif = 0 Then MiDif = 15
    VoZei.Text = "08:00"
    EndZe = DateAdd("n", MiDif, "08:00:00")
    BiZei.Text = Format$(EndZe, "hh:mm")
End If

CmNot.Enabled = GlTeE 'Email-Termin-Erinnerung
CmAcs(AD_Termin_Notify).Enabled = GlTeE 'Email-Termin-Erinnerung

RetWe = SendMessage(CmAbr.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(CmAbg.hwnd, CB_SETCURSEL, 1, ByVal 0&)
RetWe = SendMessage(CmGan.hwnd, CB_SETCURSEL, 1, ByVal 0&)
RetWe = SendMessage(CmOnT.hwnd, CB_SETCURSEL, 1, ByVal 0&)

If NotVa = 0 Then
    NotVa = 24
End If

If GlTeE = True Then 'Email-Termin-Erinnerung
    CmNot.ListIndex = NotVa

    NotDa = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "dd.mm.yyyy")
    NotZe = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "hh:mm")
    NotSt = NotDa & Chr$(32) & NotZe

    TxNoD.Text = NotDa
    TxNoZ.Text = NotZe

    TagWe = Mid$(TxNoD.Tag, 2, Len(TxNoD.Tag) - 1)
    TxNoD.Tag = "1" & TagWe
    
    TagWe = Mid$(TxNoZ.Tag, 2, Len(TxNoZ.Tag) - 1)
    TxNoZ.Tag = "1" & TagWe
    
    TagWe = Mid$(TxNoS.Tag, 2, Len(TxNoS.Tag) - 1)
    TxNoS.Tag = "1" & TagWe
Else
    RetWe = SendMessage(CmNot.hwnd, CB_SETCURSEL, NotVa, ByVal 0&)
End If

If GlTeE = True Then 'Email-Termin-Erinnerung
    If NotVa > 0 Then
        If CDate(NotSt) > Now Then
            TxNoS.Text = 3 'Senden
            CmAcs(AD_Termin_Notify).Checked = GlTeE 'Email-Termin-Erinnerung
        Else
            TxNoS.Text = 1 'Gesendet
            CmAcs(AD_Termin_Notify).Enabled = False
        End If
    Else
        TxNoS.Text = 0 'Nicht Senden
    End If
Else
    TxNoS.Text = 0 'Nicht Senden
End If

CmSta.Pane(1).Text = "Neuer Termin am: " & Format$(AkDat, "dddd" & ", " & "dd" & ". " & "mmmm" & Chr$(32) & "yyyy") & " um: " & Format$(VoZei.Text, "hh:mm") & " Uhr"

TxID2.Text = -1
GlTem = -1

TxRef.Text = Format$(MasTe, "000000")
DoEvents

DoEvents
Ter_Lei GlTem, True

Set DaPi1 = Nothing
Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

NeErr:
If GlDbg = True Then SErLog Err.Description & " TeNew " & Err.Number
Resume Next

End Sub
Public Sub TePos()
On Error GoTo ReErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl

Set FM = frmTermin
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set TxOrt = FM.txtRaum1
Set CmBrs = FM.comBar02
Set CmBet = FM.txtBetre
Set TxAdr = FM.txtAdres
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set CmArz = FM.cmbArzNr
Set CmETy = FM.cmbTypen
Set CmBez = FM.cmbBezei
Set TxAnz = FM.txtAnzal
Set TxMul = FM.txtMulti
Set TxEin = FM.txtEinze
Set TxKom = FM.txtKomme
Set CmRmu = FM.cmbRaum1
Set CmMar = FM.cmbTeTyp
Set RpCo1 = FM.repCont1
Set RpCo2 = FM.repCont2

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    Rahm1.Move ClLin, ClObn, ClBre, ClHoh
    Rahm2.Move ClLin, ClObn, ClBre, 440
    Rahm3.Move ClLin, ClObn, ClBre, ClHoh
    RpCo1.Move ClLin, ClObn + 500, ClBre, ClHoh - 500
    RpCo2.Move ClLin, ClObn, ClBre, ClHoh
    TxKom.Height = ClHoh - 4600
    CmBez.Move 1880, 60, ClBre - 3960, 315
    TxAnz.Move ClBre - 2040, 60, 500, 340
    TxMul.Move ClBre - 1500, 60, 600, 340
    TxEin.Move ClBre - 860, 60, 800, 340
End If

Set CmBrs = Nothing
Set RpCo1 = Nothing
Set RpCo2 = Nothing

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " TePos " & Err.Number
Resume Next

End Sub
Public Sub TeOut()
On Error GoTo SpErr

Dim TxKom As XtremeSuiteControls.FlatEdit
Dim TxDa1 As XtremeSuiteControls.FlatEdit
Dim TxDa2 As XtremeSuiteControls.FlatEdit
Dim TxOrt As XtremeSuiteControls.FlatEdit
Dim TxAdr As XtremeSuiteControls.FlatEdit
Dim VoZei As XtremeSuiteControls.FlatEdit
Dim BiZei As XtremeSuiteControls.FlatEdit
Dim TxGui As XtremeSuiteControls.FlatEdit
Dim CmBet As XtremeSuiteControls.ComboBox
Dim CmRem As XtremeSuiteControls.ComboBox

Dim Datu1 As Date
Dim Datu2 As Date
Dim OutRe As Boolean
Dim Mld1, Tit1 As String
Dim Frage As Integer

Dim NaSpa As Object
Dim MapFo As Object
Dim TeIts As Object
Dim TeItm As Object

If WindowLoad("frmTermin") = True Then
    Set FM = frmTermin
    Set TxID2 = FM.txtID2
Else
    Set FM = frmTermVo
End If

Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set TxAdr = FM.txtAdres
Set TxOrt = FM.txtRaum1
Set TxKom = FM.txtKomme
Set TxGui = FM.txtGuiID
Set CmTyp = FM.cmbStatu
Set CmMar = FM.cmbTeTyp
Set CmPri = FM.cmbPrior
Set CmRem = FM.cmbRemin
Set CmBet = FM.txtBetre
Set CmGan = FM.cmbGanzt

Tit1 = "Outlookübergabe"
Mld1 = "Dieser Termin wurde bereits einmal an Outlook übergeben. Möchten Sie diesen erneut an Outlook übergeben?"

If IsDate(TxDa1.Text) Then
    Datu1 = TxDa1.Text
Else
    Datu1 = Date
End If

If IsDate(TxDa2.Text) Then
    Datu2 = TxDa2.Text
Else
    Datu2 = Date
End If

If GlTem < 1 Then
    WindowMess "Sie müssen die aktuellen Daten erst speichern", Dial2, Tit1, FM.hwnd
    Exit Sub
Else
    OutRe = S_TeDx(GlTem, "Replicated")
End If

If OutRe = True Then
    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage <> 6 Then
        Exit Sub
    End If
End If

SOuOp 'Outlook

Set NaSpa = OutOb.GetNamespace("MAPI") 'NameSpace

If GlMaF = True Then
    Set MapFo = NaSpa.GetDefaultFolder(olFolderCalendar)
Else
    Set MapFo = NaSpa.PickFolder
End If

If TypeName(MapFo) = "Nothing" Then
    Exit Sub
ElseIf MapFo.DefaultItemType <> olAppointmentItem Then
    Exit Sub
End If

Set TeIts = MapFo.Items

Set TeItm = TeIts.Add(olAppointmentItem)
With TeItm
    .start = Format$(Datu1, "dd.mm.yyyy") & Chr$(32) & Format$(VoZei.Text, "hh:mm:ss")
    .End = Format$(Datu2, "dd.mm.yyyy") & Chr$(32) & Format$(BiZei.Text, "hh:mm:ss")
    .BusyStatus = CmTyp.ListIndex
    If CmGan.ListIndex = 0 Then .AllDayEvent = True
    If CmRem.Enabled = True Then .ReminderSet = True
    If TxKom.Text <> vbNullString Then .Body = TxKom.Text
    If TxOrt.Text <> vbNullString Then .Location = TxOrt.Text
    If TxGui.Text <> vbNullString Then
        .BillingInformation = TxGui.Text
    Else
        .BillingInformation = CreateID("T")
    End If
    If CmBet.Text = vbNullString Then
        If TxAdr.Text = vbNullString Then
            .Subject = "Termin"
        Else
            .Subject = TxAdr.Text
        End If
    Else
        If TxAdr.Text = vbNullString Then
            .Subject = CmBet.Text
        Else
            .Subject = TxAdr.Text & Chr$(32) & CmBet.Text
        End If
    End If
    Select Case CmPri.ListIndex + 1
    Case 1: .Importance = xtpCalendarImportanceHigh
    Case 2: .Importance = xtpCalendarImportanceNormal
    Case 3: .Importance = xtpCalendarImportanceLow
    End Select
    If CmRem.Enabled = True Then
        Select Case CmRem.ListIndex + 1
        Case 1: .ReminderMinutesBeforeStart = 0
        Case 2: .ReminderMinutesBeforeStart = 1
        Case 3: .ReminderMinutesBeforeStart = 2
        Case 4: .ReminderMinutesBeforeStart = 5
        Case 5: .ReminderMinutesBeforeStart = 10
        Case 6: .ReminderMinutesBeforeStart = 15
        Case 7: .ReminderMinutesBeforeStart = 30
        Case 8: .ReminderMinutesBeforeStart = 60
        Case 9: .ReminderMinutesBeforeStart = 120
        Case 10: .ReminderMinutesBeforeStart = 300
        Case 11: .ReminderMinutesBeforeStart = 600
        Case 12: .ReminderMinutesBeforeStart = 1440
        Case 13: .ReminderMinutesBeforeStart = 2880
        Case 14: .ReminderMinutesBeforeStart = 7200
        Case 15: .ReminderMinutesBeforeStart = 10080
        Case 16: .ReminderMinutesBeforeStart = 20160
        End Select
    End If
    DoEvents
    .Save
End With

DoEvents
DBCmEx1 "qryTerRep1", "@IdxNr", GlTem

WindowMess "Der Termin wurden erfolgreich an Outlook übergeben", Dial2, Tit1, FM.hwnd

Set TeItm = Nothing
Set NaSpa = Nothing
Set MapFo = Nothing
Set TeIts = Nothing
Set OutOb = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " TeOut " & Err.Number
Resume Next

End Sub
Public Sub TerAs(Optional ByVal Ulaub As Boolean = False)
On Error GoTo SpErr

Set FM = frmTerKop

FM.Urlaub = Ulaub
FM.Show vbModal

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " TeOut " & Err.Number
Resume Next

End Sub
Private Sub TeReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "Termin") = False Then
    If GlFnt = True Then
        xGro = 785
        yGro = 590
    Else
        xGro = 885
        yGro = 690
    End If

    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)

    IniSetSek "Termin"
    IniSetVal "Termin", "FenLin", xPos
    IniSetVal "Termin", "FenObe", yPos
    IniSetVal "Termin", "FenBre", xGro
    IniSetVal "Termin", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " TeReg " & Err.Number
Resume Next

End Sub
Private Function TeRut(AkDat As Date) As Integer
On Error Resume Next

Dim LauDa As Date
Dim LauTa As Integer
Dim LaWTa As Integer
Dim VerTa As Integer
Dim VeWTa As Integer
Dim AnzTa As Integer

AnzTa = 1
VerTa = DatePart("d", AkDat, vbMonday, vbFirstFullWeek)
VeWTa = DatePart("w", AkDat, vbMonday, vbFirstFullWeek)

For LauTa = 1 To 31
    If LauTa = VerTa Then
        Exit For
    End If
    LauDa = LauTa & "." & DatePart("m", AkDat, vbMonday, vbFirstFullWeek) & "." & DatePart("yyyy", AkDat, vbMonday, vbFirstFullWeek)
    LaWTa = DatePart("w", LauDa, vbMonday, vbFirstFullWeek)
    If LaWTa = VeWTa Then
        AnzTa = AnzTa + 1
    End If
Next LauTa
TeRut = AnzTa

End Function
Public Sub TeSpa(ByVal TbSel As Long)
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermin
Set RpCo2 = FM.repCont2
Set RpCls = RpCo2.Columns

With RpCo2
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Select Case TbSel
Case RibTab_Ter_WarZi:

    With RpCls
        Set RpCol = .Add(Adr_ID0, "ID0", 0, False)
        Set RpCol = .Add(Adr_ID3, "ID3", 0, False)
        Set RpCol = .Add(Adr_IDKurz, "Suchbegriff", 0, True)
        Set RpCol = .Add(Adr_Geboren, "Geboren", 0, True)
        Set RpCol = .Add(Adr_Name, "Name", 0, True)
        Set RpCol = .Add(Adr_Vorname, "Vorname", 0, True)
        Set RpCol = .Add(Adr_Straße, "Straße", 0, True)
        Set RpCol = .Add(Adr_PLZ, "PLZ", 0, True)
        Set RpCol = .Add(Adr_Ort, "Ort", 0, True)
        Set RpCol = .Add(Adr_Firma1, "Firma", 0, True)
        Set RpCol = .Add(Adr_Telefon1, "Privat", 0, True)
        Set RpCol = .Add(Adr_Telefon2, "Büro", 0, True)
        Set RpCol = .Add(Adr_Telefon3, "Telefax", 0, True)
        Set RpCol = .Add(Adr_Telefon4, "Mobil", 0, True)
        Set RpCol = .Add(Adr_Telefon5, "Email", 0, True)
        Set RpCol = .Add(Adr_Geschlecht, "Geschlecht", 0, True)
        Set RpCol = .Add(Adr_Datum, "Datun", 0, False)
        Set RpCol = .Add(Adr_Briefanrede, "Briefanrede", 0, False)
        Set RpCol = .Add(Adr_Anschrift, "Anschrift", 0, False)
        Set RpCol = .Add(Adr_TreKey, "TreKey", 0, False)
        Set RpCol = .Add(Adr_Grafik, "Grafik", 0, False)
        Set RpCol = .Add(Adr_GuiID, "GuiID", 0, False)
        Set RpCol = .Add(Adr_Objekt, "Objekt", 0, False)
        Set RpCol = .Add(Adr_IDP, "Mandant", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Adr_Mandant, "Nr.", 0, True)
        Set RpCol = .Add(Adr_VIP, "VIP", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Adr_Titel, "Titel", 0, False)
        Set RpCol = .Add(Adr_Land, "Land", 0, False)
        Set RpCol = .Add(Adr_Behindert, "Behindert", 0, False)
        Set RpCol = .Add(Adr_Passiv, "Passiv", 0, False)
        Set RpCol = .Add(Adr_Gruppen, "Gruppen", 0, True)
        Set RpCol = .Add(Adr_Versand, "V", 0, True)
    End With
    
    For Each RpCol In RpCls
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = True
            .Sortable = True
            .AutoSize = False
            .AutoSortWhenGrouped = False
        End With
    Next RpCol
        
    RpCls(Adr_ID0).Width = 0
    RpCls(Adr_ID3).Width = 0
    RpCls(Adr_IDKurz).Width = 220
    If GlTFt.SIZE > 10 Then
        RpCls(Adr_Geboren).Width = 110
    Else
        RpCls(Adr_Geboren).Width = 80
    End If
    RpCls(Adr_Name).Width = 100
    RpCls(Adr_Vorname).Width = 100
    RpCls(Adr_Straße).Width = 120
    RpCls(Adr_PLZ).Width = 60
    RpCls(Adr_Ort).Width = 100
    RpCls(Adr_Firma1).Width = 150
    RpCls(Adr_Telefon1).Width = 90
    RpCls(Adr_Telefon2).Width = 90
    RpCls(Adr_Telefon3).Width = 90
    RpCls(Adr_Telefon4).Width = 90
    RpCls(Adr_Telefon5).Width = 120
    RpCls(Adr_Geschlecht).Width = 80
    RpCls(Adr_Datum).Width = 0
    RpCls(Adr_Briefanrede).Width = 0
    RpCls(Adr_Anschrift).Width = 0
    RpCls(Adr_TreKey).Width = 0
    RpCls(Adr_Grafik).Width = 0
    RpCls(Adr_GuiID).Width = 0
    RpCls(Adr_Objekt).Width = 0
    RpCls(Adr_IDP).Width = 0
    RpCls(Adr_Mandant).Width = 50
    RpCls(Adr_VIP).Width = 0
    RpCls(Adr_Titel).Width = 0
    RpCls(Adr_Land).Width = 0
    RpCls(Adr_Behindert).Width = 0
    RpCls(Adr_Passiv).Width = 0
    RpCls(Adr_Gruppen).Width = 150
    RpCls(Adr_Versand).Width = 20
    
    With RpCo2.PaintManager
        .NoFieldsAvailableText = "Es sind keine Wartenden vorhanden"
        .NoItemsText = "Es sind keine Wartenden vorhanden"
    End With

Case RibTab_Ter_Proto:

    With RpCls
        Set RpCol = .Add(TeP_IDA, "IDA", 0, False)
        Set RpCol = .Add(TeP_ID2, "ID2", 0, False)
        Set RpCol = .Add(TeP_GuiID, "GuiID", 0, False)
        Set RpCol = .Add(TeP_TerID, "TerID", 0, False)
        Set RpCol = .Add(TeP_Datum, "Datum", 80, False)
        Set RpCol = .Add(TeP_Zeit, "Uhrzeit", 55, False)
        Set RpCol = .Add(TeP_IDKurz, "Protokolltext", 0, True)
        RpCol.AutoSize = True
        Set RpCol = .Add(TeP_Kommen, "Kommentar", 160, False)
    End With
        
    For Each RpCol In RpCls
        RpCol.Alignment = xtpAlignmentLeft
        RpCol.Editable = False
        RpCol.Groupable = True
        RpCol.Sortable = False
    Next RpCol

    With RpCo2.PaintManager
        .NoFieldsAvailableText = "Es sind keine Protokolleinträge vorhanden"
        .NoItemsText = "Es sind keine Protokolleinträge vorhanden"
    End With
    
End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo2 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " TeSpa " & Err.Number
Resume Next

End Sub
Private Sub TeSpl()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermin
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns

With RpCls
    Set RpCol = .Add(TeL_ID1, "ID1", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(TeL_ID0, "ID0", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(TeL_ID2, "ID2", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(TeL_Typ, "Typ", 40, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        If GlTyV = True Then
            For AktZa = 1 To UBound(GlKrA)
                RpCol.EditOptions.Constraints.Add GlKrA(AktZa, 1), GlKrA(AktZa, 0)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(TeL_GOID, "Ziffer", 70, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(TeL_IDKurz, "Bezeichnung", 10, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = True
        .Groupable = False
        .Resizable = True
        .Sortable = False
        .AutoSize = True
    End With
    If RpCo1.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(TeL_Anz, "x", 30, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Multi, "Faktor", 40, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Betrag, "Einzel", 60, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Gesamt, "Gesamt", 60, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Selekt, vbNullString, 20, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Printer_Ink
        .Tag = 1
    End With
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " TeSpl " & Err.Number
Resume Next

End Sub
Public Function TeVoBe(ByVal StaDa As Date, ByVal StaZe As Date, ByVal EndZe As Date, Optional ByVal ManNr As Long) As Boolean
On Error GoTo ReErr
'Überprüft, ob der gewünschte Terminzeitraum innerhalb oder außerhalb der Sprechzeiten liegt

Dim SpZe1 As Date
Dim SpZe2 As Date
Dim SpZe3 As Date
Dim SpZe4 As Date
Dim TmSpr As String
Dim AktZa As Integer
Dim WoTag As Integer
Dim SpEn1 As Boolean
Dim SpEn2 As Boolean
Dim SpPru As Boolean

WoTag = Weekday(StaDa, vbMonday)

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
        If ManNr = GlMiT(AktZa, 2) Then
            TmSpr = GlMiT(AktZa, 6)
            SpPru = True
            Exit For
        End If
    Next AktZa
Else
    For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
        If ManNr = GlMaT(AktZa, 2) Then
            TmSpr = GlMaT(AktZa, 6)
            SpPru = True
            Exit For
        End If
    Next AktZa
End If

If SpPru = False Then Exit Function

Select Case WoTag
Case 1:
    SpZe1 = TimeValue(Mid$(TmSpr, 2, 5))
    SpZe2 = TimeValue(Mid$(TmSpr, 8, 5))
    SpZe3 = TimeValue(Mid$(TmSpr, 14, 5))
    SpZe4 = TimeValue(Mid$(TmSpr, 20, 5))
    If Mid$(TmSpr, 1, 1) = "A" Then SpEn1 = True
    If Mid$(TmSpr, 13, 1) = "A" Then SpEn2 = True
Case 2:
    SpZe1 = TimeValue(Mid$(TmSpr, 26, 5))
    SpZe2 = TimeValue(Mid$(TmSpr, 32, 5))
    SpZe3 = TimeValue(Mid$(TmSpr, 38, 5))
    SpZe4 = TimeValue(Mid$(TmSpr, 44, 5))
    If Mid$(TmSpr, 25, 1) = "A" Then SpEn1 = True
    If Mid$(TmSpr, 37, 1) = "A" Then SpEn2 = True
Case 3:
    SpZe1 = TimeValue(Mid$(TmSpr, 50, 5))
    SpZe2 = TimeValue(Mid$(TmSpr, 56, 5))
    SpZe3 = TimeValue(Mid$(TmSpr, 62, 5))
    SpZe4 = TimeValue(Mid$(TmSpr, 68, 5))
    If Mid$(TmSpr, 49, 1) = "A" Then SpEn1 = True
    If Mid$(TmSpr, 61, 1) = "A" Then SpEn2 = True
Case 4:
    SpZe1 = TimeValue(Mid$(TmSpr, 74, 5))
    SpZe2 = TimeValue(Mid$(TmSpr, 80, 5))
    SpZe3 = TimeValue(Mid$(TmSpr, 86, 5))
    SpZe4 = TimeValue(Mid$(TmSpr, 92, 5))
    If Mid$(TmSpr, 73, 1) = "A" Then SpEn1 = True
    If Mid$(TmSpr, 85, 1) = "A" Then SpEn2 = True
Case 5:
    SpZe1 = TimeValue(Mid$(TmSpr, 98, 5))
    SpZe2 = TimeValue(Mid$(TmSpr, 104, 5))
    SpZe3 = TimeValue(Mid$(TmSpr, 110, 5))
    SpZe4 = TimeValue(Mid$(TmSpr, 116, 5))
    If Mid$(TmSpr, 97, 1) = "A" Then SpEn1 = True
    If Mid$(TmSpr, 109, 1) = "A" Then SpEn2 = True
Case 6:
    SpZe1 = TimeValue(Mid$(TmSpr, 122, 5))
    SpZe2 = TimeValue(Mid$(TmSpr, 128, 5))
    SpZe3 = TimeValue(Mid$(TmSpr, 134, 5))
    SpZe4 = TimeValue(Mid$(TmSpr, 140, 5))
    If Mid$(TmSpr, 121, 1) = "A" Then SpEn1 = True
    If Mid$(TmSpr, 133, 1) = "A" Then SpEn2 = True
Case 7:
    SpZe1 = TimeValue(Mid$(TmSpr, 146, 5))
    SpZe2 = TimeValue(Mid$(TmSpr, 152, 5))
    SpZe3 = TimeValue(Mid$(TmSpr, 158, 5))
    SpZe4 = TimeValue(Mid$(TmSpr, 164, 5))
    If Mid$(TmSpr, 145, 1) = "A" Then SpEn1 = True
    If Mid$(TmSpr, 157, 1) = "A" Then SpEn2 = True
End Select
 
If SpEn1 = True Then 'Vormittag = Ein
    If TimeValue(StaZe) < SpZe1 Then
        TeVoBe = True
    ElseIf TimeValue(StaZe) >= SpZe2 Then
        If TimeValue(StaZe) < SpZe3 Then
            TeVoBe = True
        End If
    ElseIf TimeValue(EndZe) > SpZe2 Then
        TeVoBe = True
    End If
Else 'Vormittag = Aus
    If TimeValue(StaZe) < SpZe3 Then
        TeVoBe = True
    End If
End If

If SpEn2 = True Then 'Nachnittag = Ein
    If TimeValue(StaZe) > SpZe4 Then
        TeVoBe = True
    ElseIf TimeValue(StaZe) < SpZe3 Then
        If TimeValue(StaZe) >= SpZe2 Then
            TeVoBe = True
        End If
    ElseIf TimeValue(EndZe) > SpZe4 Then
        TeVoBe = True
    End If
Else 'Nachnittag = Aus
    If TimeValue(StaZe) >= SpZe2 Then
        TeVoBe = True
    End If
End If

If SpEn1 = True Then 'Vormittag = Ein und Nachnittag = Ein
    If SpEn2 = True Then
        If TimeValue(StaZe) >= SpZe2 Then 'Mittagspause
            If TimeValue(StaZe) < SpZe3 Then
                TeVoBe = True
            End If
        End If
    End If
End If

Exit Function

ReErr:
If GlDbg = True Then SErLog Err.Description & " TeVoBe " & Err.Number
Resume Next

End Function
Private Sub TeVoIn()
On Error GoTo InErr

Dim RetWe As Long
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmTermVo
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set Rahm5 = FM.frmRahm5
Set Rahm6 = FM.frmRahm6
Set Rahm7 = FM.frmRahm7
Set Rahm8 = FM.frmRahm8
Set Rahm9 = FM.frmRahm9
Set TxDa1 = FM.txtDatu1
Set TxDa4 = FM.txtDatu4
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set ChSpl = FM.chkTeSpl
Set ChTer = FM.chkFreTe
Set ChRau = FM.chkRauZu
Set ZwZei = FM.txoZwZei
Set ZyWoh = FM.cmbWoche
Set ZyMo1 = FM.cmoMona1
Set ZyMo2 = FM.cmoMona2
Set ZyMo3 = FM.cmbMona3
Set ZyMo4 = FM.cmbMonat
Set ZyJa1 = FM.cmoJahr1
Set ZyJa2 = FM.cmoJahr2
Set ZyJa3 = FM.cmoJahr3
Set ZyJa4 = FM.cmoJahr4
Set ZyMoT = FM.cmbMona1
Set ZyJaT = FM.cmbJahr1
Set ZyEnT = FM.cmbZyEn1
Set CmRmu = FM.cmbRaum1
Set CmGan = FM.cmbGanzt
Set CmAbr = FM.cmbAbger
Set CmNot = FM.cmbNotVa
Set MoKa1 = FM.dtpDatu3
Set ZyEn2 = FM.optZyEn2
Set ZyEn3 = FM.optZyEn3
Set FoZy1 = FM.optZykl1
Set FoZy2 = FM.optZykl2
Set FoZy3 = FM.optZykl3
Set FoZy4 = FM.optZykl4
Set TaZy1 = FM.optZyTa1
Set TaZy2 = FM.optZyTa2
Set MoZy1 = FM.optZyMo1
Set MoZy2 = FM.optZyMo2
Set JaZy1 = FM.optZyJa1
Set JaZy2 = FM.optZyJa2
Set MoKa1 = FM.dtpDatu1
Set MoKa2 = FM.dtpDatu2
Set MoKa3 = FM.dtpDatu3
Set RpCo1 = FM.repCont1
Set RpCo6 = FM.repCont6
Set ImMan = frmMain.imgManag

With MoKa1
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

With MoKa2
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

With MoKa3
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
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = GlKrE
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FocusSubItems = True
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
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.FixedRowHeight = Not GlGZe
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlKZe
    .PaintManager.UseAlternativeBackground = GlZeK
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    .OLEDropMode = xtpOLEDropNone
    RetWe = .EnableDragDrop("Katalog", xtpReportAllowDrop)
End With

With RpCo6
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = False
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips True
    .EnsureFocusedRowVisible = True
    .FocusSubItems = True 'WICHTIG!
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False
    .SkipGroupsFocus = True
    .SetCustomDraw xtpCustomBeforeDrawRow 'vor FixedRowHeight initialisieren
    .PaintManager.CaptionForeColor = -2147483641
    .PaintManager.GroupForeColor = -2147483641
    .PaintManager.NoGroupByText = "Ziehen Sie Spaltenköpfe in dieses Feld, um nach diesen Spalten zu gruppieren"
    .PaintManager.DrawSortTriangleAlways = True
    .PaintManager.HideSelection = False
    .PaintManager.HotTracking = True
    .PaintManager.NoFieldsAvailableText = "Es sind noch keine Terminvorschläge vorhanden"
    .PaintManager.NoItemsText = "Es sind noch keine Terminvorschläge vorhanden"
    .PaintManager.RevertAlignment = False
    .PaintManager.ShadeGroupHeadings = True
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
    If CBool(IniGetVal("Layout", "LinTyp")) = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.FixedRowHeight = True
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.InvertColumnOnClick = True
    .ShowGroupBox = False
    .PreviewMode = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    .ShowHeader = GlSpU
End With

VoZei.SetMask "00:00", "__:__"
BiZei.SetMask "00:00", "__:__"

TxDa1.SetMask "00.00.0000", "__.__.____"
TxDa4.SetMask "00.00.0000", "__.__.____"

CmNot.Enabled = GlTeE 'Email-Termin-Erinnerung

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
Rahm5.BackColor = GlBak
Rahm6.BackColor = GlBak
Rahm7.BackColor = GlBak
Rahm8.BackColor = GlBak
Rahm9.BackColor = GlBak
FoZy1.BackColor = GlBak
FoZy2.BackColor = GlBak
FoZy3.BackColor = GlBak
FoZy4.BackColor = GlBak
TaZy1.BackColor = GlBak
TaZy2.BackColor = GlBak
MoZy1.BackColor = GlBak
MoZy2.BackColor = GlBak
JaZy1.BackColor = GlBak
JaZy2.BackColor = GlBak
ZyEn2.BackColor = GlBak
ZyEn3.BackColor = GlBak
ChRau.BackColor = GlBak
ChSpl.BackColor = GlBak
FM.choTaMon.BackColor = GlBak
FM.choTaDin.BackColor = GlBak
FM.choTaMit.BackColor = GlBak
FM.choTaDon.BackColor = GlBak
FM.choTaFre.BackColor = GlBak
FM.choTaSam.BackColor = GlBak
FM.choTaSon.BackColor = GlBak
FM.chkDopTe.BackColor = GlBak
FM.chkSprZe.BackColor = GlBak
FM.chkFreTe.BackColor = GlBak

Set RpCo1 = Nothing
Set RpCo6 = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " TeVoIn " & Err.Number
Resume Next

End Sub
Private Sub TeVoLa()
On Error GoTo ReErr

Dim RetWe As Long
Dim ZweZe As String
Dim AkDat As String
Dim AktZa As Integer
Dim WoTag As Integer
Dim JaMon As Integer
Dim AnzVo As Integer
Dim AnzSp As Integer
Dim PauSp As Integer
Dim mAnza As Integer
Dim NotVa As Integer

Set FM = frmTermVo
Set TxDa1 = FM.txtDatu1
Set TxDa4 = FM.txtDatu4
Set CmBet = FM.txtBetre
Set CmTyp = FM.cmbStatu
Set CmMar = FM.cmbTeTyp
Set CmGes = FM.cmbGesch
Set CmRmu = FM.cmbRaum1
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set TxKom = FM.txtKomme
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set CmRem = FM.cmbRemin
Set CmGan = FM.cmbGanzt
Set CmAbr = FM.cmbAbger
Set CmNot = FM.cmbNotVa
Set ChTer = FM.chkFreTe
Set ChRau = FM.chkRauZu
Set ChSpr = FM.chkSprZe
Set ZwZei = FM.txoZwZei
Set TxSp1 = FM.txoSplAn
Set TxSp2 = FM.txoSplPa
Set ZyWoh = FM.cmbWoche
Set ZyMo1 = FM.cmoMona1
Set ZyMo2 = FM.cmoMona2
Set ZyMo3 = FM.cmbMona3
Set ZyMo4 = FM.cmbMonat
Set ZyJa1 = FM.cmoJahr1
Set ZyJa2 = FM.cmoJahr2
Set ZyJa3 = FM.cmoJahr3
Set ZyJa4 = FM.cmoJahr4
Set ZyMoT = FM.cmbMona1
Set ZyJaT = FM.cmbJahr1
Set ZyEnT = FM.cmbZyEn1
Set CmRmu = FM.cmbRaum1

AnzVo = IniGetVal("TerSys", "AnzVor")
AnzSp = IniGetVal("TerSys", "TeSpAn")
PauSp = IniGetVal("TerSys", "TeSpPa")
ZweZe = Format$(IniGetVal("TerSys", "ZwZeit"), "hh:mm")

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    NotVa = GlMiT(GlSMo, 39) 'Standardmitarbeiter Online-Terminbuchungs Sytem
Else
    NotVa = GlMaT(GlSMa, 25)
End If

If NotVa = 0 Then
    NotVa = 24
End If

For AktZa = 1 To UBound(GlBtr)
    With CmBet
        .AddItem GlBtr(AktZa, 1)
        .ItemData(AktZa - 1) = GlBtr(AktZa, 0)
    End With
Next AktZa

For AktZa = 1 To UBound(GlTep) 'Kalendermarker
    With CmMar
        .AddItem GlTep(AktZa, 1)
        .ItemData(AktZa - 1) = GlTep(AktZa, 0)
    End With
Next AktZa

For AktZa = 1 To UBound(GlTeS)
    With CmTyp
        .AddItem GlTeS(AktZa, 1)
        .ItemData(AktZa - 1) = GlTeS(AktZa, 0)
    End With
Next AktZa

For AktZa = 1 To UBound(GlRmu)
    With CmRmu
        .AddItem GlRmu(AktZa, 1)
        .ItemData(AktZa - 1) = GlRmu(AktZa, 2)
    End With
Next AktZa

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    For AktZa = 1 To UBound(GlMan) 'Aktive Mandanten
        With CmMan
            If CBool(GlMan(AktZa, 5)) = False Then
                mAnza = mAnza + 1
                .AddItem GlMaT(AktZa, 1)
                .ItemData(mAnza - 1) = GlMan(AktZa, 2)
            End If
        End With
    Next AktZa
Else
    For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
        With CmMan
            If CBool(GlMaT(AktZa, 5)) = False Then
                mAnza = mAnza + 1
                .AddItem GlMaT(AktZa, 1)
                .ItemData(mAnza - 1) = GlMaT(AktZa, 2)
            End If
        End With
    Next AktZa
End If

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If GlMiV = True Then
        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
            With CmMit
                .AddItem GlMiT(AktZa, 1)
                .ItemData(AktZa - 1) = GlMiT(AktZa, 2)
            End With
        Next AktZa
    End If
Else
    If GlMiV = True Then
        For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
            With CmMit
                .AddItem GlMiA(AktZa, 1)
                .ItemData(AktZa - 1) = GlMiA(AktZa, 2)
            End With
        Next AktZa
    End If
End If

With CmGes
    For AktZa = 0 To UBound(GlGes) - 1
        .AddItem GlGes(AktZa)
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With CmRem
    .AddItem "0 Min."
    .ItemData(0) = 1
    .AddItem "1 Min."
    .ItemData(1) = 2
    .AddItem "2 Min."
    .ItemData(2) = 3
    .AddItem "5 Min."
    .ItemData(3) = 4
    .AddItem "10 Min."
    .ItemData(4) = 5
    .AddItem "15 Min."
    .ItemData(5) = 6
    .AddItem "30 Min."
    .ItemData(6) = 7
    .AddItem "1 Std."
    .ItemData(7) = 8
    .AddItem "2 Std."
    .ItemData(8) = 9
    .AddItem "5 Std."
    .ItemData(9) = 10
    .AddItem "10 Std."
    .ItemData(10) = 11
    .AddItem "1 Tag"
    .ItemData(11) = 12
    .AddItem "2 Tage"
    .ItemData(12) = 13
    .AddItem "5 Tage"
    .ItemData(13) = 14
End With

With TxSp1
    For AktZa = 2 To 20
        .AddItem AktZa & " Teile"
        .ItemData(AktZa - 2) = AktZa
    Next AktZa
End With
    
With TxSp2
    .AddItem "00 Min."
    .ItemData(0) = 0
    .AddItem "05 Min."
    .ItemData(1) = 5
    .AddItem "10 Min."
    .ItemData(2) = 10
    .AddItem "15 Min."
    .ItemData(3) = 15
    .AddItem "20 Min."
    .ItemData(4) = 20
    .AddItem "30 Min."
    .ItemData(5) = 30
    .AddItem "45 Min."
    .ItemData(6) = 45
    .AddItem "60 Min."
    .ItemData(7) = 60
End With

With CmGan
    .AddItem "Ja"
    .ItemData(0) = -1
    .AddItem "Nein"
    .ItemData(1) = 0
End With

With ZwZei
    For AktZa = 6 To 22
        .AddItem Format$(AktZa, "00") & ":" & "00"
        .AddItem Format$(AktZa, "00") & ":" & "15"
        .AddItem Format$(AktZa, "00") & ":" & "30"
        .AddItem Format$(AktZa, "00") & ":" & "45"
    Next AktZa
End With

With ZyWoh
    .AddItem "Jede Woche"
    .ItemData(0) = 1
    .AddItem "Jede zweite Woche"
    .ItemData(1) = 2
    .AddItem "Jede dritte Woche"
    .ItemData(2) = 3
    .AddItem "Jede vierte Woche"
    .ItemData(3) = 4
End With

With ZyMo1
    .AddItem "ersten"
    .ItemData(0) = 1
    .AddItem "zweiten"
    .ItemData(1) = 2
    .AddItem "dritten"
    .ItemData(2) = 3
    .AddItem "vierten"
    .ItemData(3) = 4
    .AddItem "letzten"
    .ItemData(4) = 5
End With

With ZyMo2
    .AddItem "Sonntag"
    .ItemData(0) = 1
    .AddItem "Montag"
    .ItemData(1) = 2
    .AddItem "Dienstag"
    .ItemData(2) = 3
    .AddItem "Mittwoch"
    .ItemData(3) = 4
    .AddItem "Donnerstag"
    .ItemData(4) = 5
    .AddItem "Freitag"
    .ItemData(5) = 6
    .AddItem "Samstag"
    .ItemData(6) = 7
End With

With ZyMo3
    .AddItem "jedes Monats"
    .ItemData(0) = 1
    .AddItem "jedes zweiten Monats"
    .ItemData(1) = 2
    .AddItem "jedes dritten Monats"
    .ItemData(2) = 3
    .AddItem "jedes vierten Monats"
    .ItemData(3) = 4
End With

With ZyMo4
    .AddItem "jedes Monats"
    .ItemData(0) = 1
    .AddItem "jedes zweiten Monats"
    .ItemData(1) = 2
    .AddItem "jedes dritten Monats"
    .ItemData(2) = 3
    .AddItem "jedes vierten Monats"
    .ItemData(3) = 4
End With

With ZyJa1
    .AddItem "Januar"
    .ItemData(0) = 1
    .AddItem "Februar"
    .ItemData(1) = 2
    .AddItem "März"
    .ItemData(2) = 3
    .AddItem "April"
    .ItemData(3) = 4
    .AddItem "Mai"
    .ItemData(4) = 5
    .AddItem "Juni"
    .ItemData(5) = 6
    .AddItem "Juli"
    .ItemData(6) = 7
    .AddItem "August"
    .ItemData(7) = 8
    .AddItem "September"
    .ItemData(8) = 9
    .AddItem "Oktober"
    .ItemData(9) = 10
    .AddItem "November"
    .ItemData(10) = 11
    .AddItem "Dezember"
    .ItemData(11) = 12
End With

With ZyJa4
    .AddItem "Januar"
    .ItemData(0) = 1
    .AddItem "Februar"
    .ItemData(1) = 2
    .AddItem "März"
    .ItemData(2) = 3
    .AddItem "April"
    .ItemData(3) = 4
    .AddItem "Mai"
    .ItemData(4) = 5
    .AddItem "Juni"
    .ItemData(5) = 6
    .AddItem "Juli"
    .ItemData(6) = 7
    .AddItem "August"
    .ItemData(7) = 8
    .AddItem "September"
    .ItemData(8) = 9
    .AddItem "Oktober"
    .ItemData(9) = 10
    .AddItem "November"
    .ItemData(10) = 11
    .AddItem "Dezember"
    .ItemData(11) = 12
End With

With ZyJa2
    .AddItem "ersten"
    .ItemData(0) = 1
    .AddItem "zweiten"
    .ItemData(1) = 2
    .AddItem "dritten"
    .ItemData(2) = 3
    .AddItem "vierten"
    .ItemData(3) = 4
    .AddItem "letzten"
    .ItemData(4) = 5
End With

With ZyJa3
     .AddItem "Sonntag"
    .ItemData(0) = 1
    .AddItem "Montag"
    .ItemData(1) = 2
    .AddItem "Dienstag"
    .ItemData(2) = 3
    .AddItem "Mittwoch"
    .ItemData(3) = 4
    .AddItem "Donnerstag"
    .ItemData(4) = 5
    .AddItem "Freitag"
    .ItemData(5) = 6
    .AddItem "Samstag"
    .ItemData(6) = 7
End With

With ZyMoT
    For AktZa = 0 To 31 - 1
        .AddItem Format$(AktZa + 1, "00") & "."
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With ZyJaT
    For AktZa = 0 To 31 - 1
        .AddItem Format$(AktZa + 1, "00") & "."
        .ItemData(AktZa) = AktZa + 1
    Next AktZa
End With

With ZyEnT
    For AktZa = 2 To 99
        .AddItem AktZa & " Termine"
        .ItemData(AktZa - 2) = AktZa
    Next AktZa
End With

With CmNot
    For AktZa = 0 To 48
        .AddItem AktZa & " Std."
        .ItemData(AktZa) = AktZa
    Next AktZa
End With

With CmAbr
    .AddItem "keine Leistungen"
    .ItemData(0) = 1
    .AddItem "Leistungen vorhanden"
    .ItemData(1) = 2
    .AddItem "Leistungen abgerechnet"
    .ItemData(2) = 3
End With

CmAcs(AD_Termin_Notify).Checked = GlTeE 'Email-Termin-Erinnerung

ChTer.Value = CBool(IniGetVal("TerSys", "TerSer"))
ChRau.Value = CBool(IniGetVal("TerSys", "RmuBer"))
ChSpr.Value = GlSpP 'Überprüfung der Sprechzeiten

ZwZei.Text = ZweZe

RetWe = SendMessage(CmRem.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(CmGan.hwnd, CB_SETCURSEL, 1, ByVal 0&)
RetWe = SendMessage(CmTyp.hwnd, CB_SETCURSEL, 2, ByVal 0&)
RetWe = SendMessage(ZyWoh.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyMo1.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyMo3.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyMo4.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(ZyJa2.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(CmAbr.hwnd, CB_SETCURSEL, 0, ByVal 0&)
RetWe = SendMessage(TxSp1.hwnd, CB_SETCURSEL, AnzSp, ByVal 0&)
RetWe = SendMessage(TxSp2.hwnd, CB_SETCURSEL, PauSp, ByVal 0&)
RetWe = SendMessage(ZyEnT.hwnd, CB_SETCURSEL, AnzVo, ByVal 0&)

TxDa4.Text = Format$(DateAdd("d", 30, Date), "dd.mm.yyyy")

WoTag = Weekday(Date)
JaMon = Month(Date)

If WoTag = 7 Then
    ZyMo2.ListIndex = 0
Else
    ZyMo2.ListIndex = WoTag - 1
End If

ZyJa3.ListIndex = WoTag - 1

ZyJa1.ListIndex = JaMon - 1
ZyJa4.ListIndex = JaMon - 1
ZyMoT.ListIndex = 0
ZyJaT.ListIndex = 0

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " TeVoLa " & Err.Number
Resume Next

End Sub
Public Sub TeVoMa(Optional ByVal PatNr As Long, Optional ByVal PatNa As String)
On Error GoTo LaErr

Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmTermVo") = True Then
    Set FM = frmTermVo
    frmTermVo.ZOrder 0
    Exit Sub
End If

GlTeF = True 'Formular wird geladen
GlSeF = True

TeVoRe
DoEvents

Load frmTermVo

Set FM = frmTermVo

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    If GlIdi = True Then 'Idiotenmodus
        .FeLin = (GlxGr / 2) - (930 / 2)
        .FeObn = (GlyGr / 2) - (680 / 2)
        .FeBre = 930
        .FeHoh = 680
    Else
        .FeLin = IniGetVal("TermVo", "FenLin")
        .FeObn = IniGetVal("TermVo", "FenObe")
        .FeBre = IniGetVal("TermVo", "FenBre")
        .FeHoh = IniGetVal("TermVo", "FenHoh")
    End If
End With

TeVoIn
AFont FM
TeVoMe
TeVoLa
TeVoSp

If PatNr > 0 Then
    TeVoNe PatNr, PatNa
Else
    TeVoNe
End If

DoEvents
TeVoSt
TeVoOp

With clFen
    .FenMov
    Set CmBrs = FM.comBar02
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    TeVoPo
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

frmTermVo.Show
DoEvents

GlTeF = False 'Formular wird geladen
GlSeF = False
GlVrt = False 'virtuelle Leistungen vorhanden

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " TeVoMa " & Err.Number
Resume Next

End Sub
Private Sub TeVoMe()
On Error GoTo InErr
'Menue erstellen

Dim RetWe As Long
Dim KeyNa As String
Dim ZaZil As Integer
Dim AktZa As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim RbTem As XtremeCommandBars.RibbonTab
Dim MsBar As XtremeCommandBars.MessageBar
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmTermVo
Set CmBrs = FM.comBar02
Set PuBu1 = FM.btnDatu1
Set PuBu4 = FM.btnDatu4
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set ImMan = frmMain.imgManag
Set CmGlo = XtremeCommandBars.CommandBarsGlobalSettings

KeyNa = "ToolTips"

For AktZa = 1 To UBound(GlZah) 'Standardzahlungsziel
    If GlZah(AktZa, 0) = GlStZ Then
        ZaZil = GlZah(AktZa, 2)
        Exit For
    End If
Next AktZa

With CmBrs
    .EnableActions
    .Icons = ImMan.Icons
End With

With CmAcs
    Set CmAct = .Add(AD_Termin_Remind, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Notify, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Vorschau, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Save, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Reset, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Freie, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Ketten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_StaKett, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_Abrechnen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(AD_Termin_EintLoe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(RibGrp_TeR_Ausgabe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei01, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei02, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei03, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei04, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei05, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei06, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei07, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei08, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei09, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei10, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei11, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei12, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei13, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei14, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei15, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei16, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei17, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei18, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei19, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(FaLei20, vbNullString, vbNullString, vbNullString, vbNullString)
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Width = 120
    CmPan.Alignment = xtpAlignmentCenter
    CmPan.Text = vbNullString
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Style = SBPS_STRETCH
    Set CmPan = .AddPane(3)
    CmPan.Width = 120
    CmPan.Text = vbNullString
    CmPan.Alignment = xtpAlignmentLeft
    .Visible = True
End With

'--------------------------------------------------

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmPop = RbBar.Controls.Add(xtpControlPopup, TE_Farbe, "Terminfarbe")
With CmPop
    .flags = xtpFlagRightAlign
    .Style = xtpButtonCaption
    '.IconId = FaLei01
End With
Set CmCoS = CmPop.CommandBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, FaLei01, GlTmF(1, 0))
    CmCon.Checked = True
    Set CmCon = .Add(xtpControlButton, FaLei02, GlTmF(2, 0))
    Set CmCon = .Add(xtpControlButton, FaLei03, GlTmF(3, 0))
    Set CmCon = .Add(xtpControlButton, FaLei04, GlTmF(4, 0))
    Set CmCon = .Add(xtpControlButton, FaLei05, GlTmF(5, 0))
    Set CmCon = .Add(xtpControlButton, FaLei06, GlTmF(6, 0))
    Set CmCon = .Add(xtpControlButton, FaLei07, GlTmF(7, 0))
    Set CmCon = .Add(xtpControlButton, FaLei08, GlTmF(8, 0))
    Set CmCon = .Add(xtpControlButton, FaLei09, GlTmF(9, 0))
    Set CmCon = .Add(xtpControlButton, FaLei10, GlTmF(10, 0))
    Set CmCon = .Add(xtpControlButton, FaLei11, GlTmF(11, 0))
    Set CmCon = .Add(xtpControlButton, FaLei12, GlTmF(12, 0))
    Set CmCon = .Add(xtpControlButton, FaLei13, GlTmF(13, 0))
    Set CmCon = .Add(xtpControlButton, FaLei14, GlTmF(14, 0))
    Set CmCon = .Add(xtpControlButton, FaLei15, GlTmF(15, 0))
    Set CmCon = .Add(xtpControlButton, FaLei16, GlTmF(16, 0))
    Set CmCon = .Add(xtpControlButton, FaLei17, GlTmF(17, 0))
    Set CmCon = .Add(xtpControlButton, FaLei18, GlTmF(18, 0))
    Set CmCon = .Add(xtpControlButton, FaLei19, GlTmF(19, 0))
    Set CmCon = .Add(xtpControlButton, FaLei20, GlTmF(20, 0))
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, TE_Termin_Hilfe, "Hilfe")
With CmBuT
    .ToolTipText = "Öffnet die Kurzhilfe"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Sign_Help
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmBuT = RbBar.Controls.Add(xtpControlButton, TE_Termin_Beenden, "Schließen")
With CmBuT
    .flags = xtpFlagRightAlign
    .IconId = IC16_Exit
    .ToolTipText = "Schließt den Dialog"
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F11"
End With
Set CmCon = RbBar.Controls.Add(xtpControlLabel, RibCon_Caption, Space$(1))
With CmCon
    .flags = xtpFlagRightAlign
    .Style = xtpButtonIconAndCaption
End With

'--------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Ter_Haupt, "Termindaten")
With RbTab
    .id = RibTab_Ter_Haupt
    .ToolTip = "Zeigt die Hauptdaten des Termins"
    .Visible = True
    .Selected = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Vorschau, "Termine Vorschlagen")
With CmCon
    .IconId = IC32_Calendar_Light
    .ShortcutText = "F3"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Save, "Termine Speichern")
With CmCon
    .IconId = IC32_Disk_Calendar
    .ShortcutText = "F8"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Reset, "Termine Zurücksetzen")
With CmCon
    .IconId = IC32_Calendar_Undo
    .ShortcutText = "F7"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("", RibGrp_Ter_Ansicht)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Freie, "Nächster freier Termin")
With CmCon
    .IconId = IC32_Calendar_Phone
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Ausgabe", RibGrp_TeR_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlCheckBox, AD_Termin_Remind, "Terminerinnerung")
Set CmCon = RbGrp.Add(xtpControlCheckBox, AD_Termin_Notify, "Emailerinnerung")
CmCon.Enabled = GlTeE 'Email-Termin-Erinnerung
Set CmCon = RbGrp.Add(xtpControlButton, TE_Termin_Drucken, "Termine Drucken")
With CmCon
    .flags = xtpFlagRightAlign
    .IconId = IC16_Printer_Ink
    .Style = xtpButtonIconAndCaption
    .ShortcutText = "F10"
End With
'---

Set RbTab = RbBar.InsertTab(RibTab_Ter_Leist, "Leistungsziffern")
With RbTab
    .id = RibTab_Ter_Leist
    .ToolTip = "Zeigt die zugeordneten Leistungen"
    .Visible = True
    If GlRch(0, 9) = 0 Then .Enabled = False
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ter_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Abrechnen, "Rechnung Erstellen")
With CmCon
    .IconId = IC32_Mail_Export
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_EintLoe, "Einträge Entfernen")
With CmCon
    .IconId = IC32_Doc_Del
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Leistungen", RibGrp_TeR_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_Ketten, "Leistungen Auswählen")
With CmCon
    .IconId = IC32_Folder_View
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt1, "Terminbetrag :")
CmCon.flags = xtpFlagRightAlign
Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt2, "Serienbetrag :")
CmCon.flags = xtpFlagRightAlign
Set CmCon = RbGrp.Add(xtpControlLabel, AD_Termin_Capt3, "Fälligkeit :")
CmCon.flags = xtpFlagRightAlign
Set CmEdi = RbGrp.Add(xtpControlEdit, AD_Termin_Betrag1, vbNullString)
With CmEdi
    .EditStyle = xtpEditStyleRight
    .Style = xtpButtonIconAndCaption
    .Width = 70
    .Text = GlWa2
End With
Set CmEdi = RbGrp.Add(xtpControlEdit, AD_Termin_Betrag2, vbNullString)
With CmEdi
    .EditStyle = xtpEditStyleRight
    .Style = xtpButtonIconAndCaption
    .Width = 70
    .Text = GlWa2
End With
Set CmEdi = RbGrp.Add(xtpControlEdit, AD_Termin_Betrag3, vbNullString)
With CmEdi
    .EditStyle = xtpEditStyleRight
    .Style = xtpButtonIconAndCaption
    .Width = 70
    .Text = Format$(DateAdd("d", ZaZil, Date), "dd.mm.yyyy")
End With
Set CmCon = RbGrp.Add(xtpControlButton, AD_Termin_StaKett, "Standardkette Einfügen")
With CmCon
    .IconId = IC32_Link_Down
    .Width = GlRib
    .BeginGroup = True
End With

'---

Set CmCoS = RbBar.Controls
For Each CmCon In CmCoS
    CmCon.ToolTipText = IniGetOpt(KeyNa, CmCon.id)
Next CmCon

'---

With CmGlo
    Select Case GlSty
    Case 1: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Blue.ini"
    Case 2: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Black.ini"
    Case 3: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Silver.ini"
    Case 4: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Aqua.ini"
    Case 5: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2010.dll", "Office2010Silver.ini"
    Case 6: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2007.dll", "Office2007Blue.ini"
    Case 7: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    Case 8: .ResourceImages.LoadFromFile App.Path & "\Styles\Office2013.dll", "Office2013White.ini"
    End Select
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
    .SetIconSize True, 32, 32
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
    .Font.Name = GlTFt.Name
    .ComboBoxFont.SIZE = 8
    .ComboBoxFont.Name = GlTFt.Name
End With

With CmBrs
    .EnableOffice2007Frame False
    Select Case GlSty
    Case 7:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case 8:
        .VisualTheme = xtpThemeOffice2013
        .AllowFrameTransparency False
        .SetAllCaps True
        .StatusBar.SetAllCaps True
    Case Else:
        If GlRah = True Then 'Office EnableThemeframe
            .VisualTheme = xtpThemeRibbon
        Else
            If GlFRg = True Then 'farbige Register
                .VisualTheme = xtpThemeResource
            Else
                .VisualTheme = xtpThemeRibbon
            End If
        End If
        .AllowFrameTransparency True
        .SetAllCaps False
        .StatusBar.SetAllCaps False
    End Select
    .EnableCustomization False
    .ActiveMenuBar.Closeable = False
    .ActiveMenuBar.Customizable = False
    .ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .ActiveMenuBar.Position = xtpBarTop
    .ActiveMenuBar.ShowExpandButton = False
    .ActiveMenuBar.ShowTextBelowIcons = False
    .ActiveMenuBar.Visible = True
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

Set RbBar = CmBrs.Item(1)
With RbBar
    .AllowMinimize = False
    .AllowQuickAccessCustomization = False
    .AllowQuickAccessDuplicates = False
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .EnableAnimation = GlMeA
    .FontHeight = GlToF
    .GroupsVisible = True
    .MinimumVisibleWidth = 100
    .RibbonPaintManager.HotTrackingGroups = True
    .RibbonPaintManager.CaptionFont.SIZE = 8
    .RibbonPaintManager.CaptionFont.Name = GlTFt.Name
    .RibbonPaintManager.WindowCaptionFont.SIZE = 8
    .RibbonPaintManager.WindowCaptionFont.Name = GlTFt.Name
    .ShowQuickAccess = False
    .ShowQuickAccessBelowRibbon = False
    .ShowCaptionAlways = True
    .Position = xtpBarTop
    .SetIconSize 16, 16
    Select Case GlSty
    Case 8:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case 7:
        .TabPaintManager.Color = xtpTabColorOffice2013
        .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
    Case Else:
        If GlFRg = True Then 'Farbige Register
            .TabPaintManager.Appearance = xtpTabAppearanceVisualStudio2005
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.ButtonMargin.Top = 6
            .TabPaintManager.ButtonMargin.Bottom = 0
            .TabPaintManager.HeaderMargin.Top = 0
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter
        Else
            .TabPaintManager.Color = xtpTabColorResource
            .TabPaintManager.DrawTextFormat = xtpTabDrawTextCenter + xtpTabDrawTextVCenter
        End If
    End Select
    .TabPaintManager.Layout = xtpTabLayoutAutoSize
    .TabPaintManager.HeaderMargin.Left = 7
    .TabPaintManager.MinTabWidth = 100
    .TabPaintManager.ClearTypeTextQuality = GlCle
    .TabPaintManager.BoldSelected = False
    .TabPaintManager.ClientFrame = xtpTabFrameNone
    .TabPaintManager.DisableLunaColors = False
    .TabPaintManager.DrawTextNoPrefix = False
    .TabPaintManager.DrawTextPathEllipsis = False
    .TabPaintManager.FillBackground = False
    .TabPaintManager.HotTracking = True
    .TabPaintManager.MultiRowFixedSelection = True
    .TabPaintManager.MultiRowJustified = False
    .TabPaintManager.OneNoteColors = True
    .TabPaintManager.SelectTabOnDragOver = True
    .TabPaintManager.ShowIcons = False
    .TabPaintManager.StaticFrame = False
    .TabPaintManager.ToolTipBehaviour = xtpTabToolTipAlways
    .TabPaintManager.Font.SIZE = 8
    .TabPaintManager.Font.Name = GlTFt.Name
    If GlRDP = True Then
        .EnableFrameTheme
    Else
        If GlRah = True Then
            .EnableFrameTheme
        End If
    End If
End With

CmAcs(AD_Termin_Save).Enabled = False
CmAcs(AD_Termin_Reset).Enabled = False
CmAcs(AD_Termin_Abrechnen).Enabled = False
CmAcs(AD_Termin_EintLoe).Enabled = False
CmAcs(AD_Termin_Ketten).Enabled = False
CmAcs(AD_Termin_StaKett).Enabled = False

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu4.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Set CmPan = Nothing
Set CmSta = Nothing
Set CmPop = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set RbBar = Nothing
Set RbTab = Nothing
Set RbGrp = Nothing
Set RbGps = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " TeVoMe " & Err.Number
Resume Next

End Sub
Private Function TeVoMo(ByVal StaTa As Integer, ByVal StaMo As Integer) As Integer
On Error GoTo KoErr

Dim TmpTa As Integer

Select Case StaMo
Case 2:
    If StaTa > 28 Then
        TmpTa = 28
    Else
        TmpTa = StaTa
    End If
Case 4:
    If StaTa > 30 Then
        TmpTa = 30
    Else
        TmpTa = StaTa
    End If
Case 6:
    If StaTa > 30 Then
        TmpTa = 30
    Else
        TmpTa = StaTa
    End If
Case 9:
    If StaTa > 30 Then
        TmpTa = 30
    Else
        TmpTa = StaTa
    End If
Case 11:
    If StaTa > 30 Then
        TmpTa = 30
    Else
        TmpTa = StaTa
    End If
Case Else:
    TmpTa = StaTa
End Select

TeVoMo = TmpTa

Exit Function

KoErr:
If GlDbg = True Then SErLog Err.Description & " TeVoMo " & Err.Number
Resume Next

End Function
Public Sub TeVoNe(Optional ByVal PatNr As Long, Optional ByVal PatNa As String)
On Error GoTo NeErr
'Bereitet die Neueingabe eines Termines vor

Dim RetWe As Long
Dim MitNr As Long
Dim ManNr As Long
Dim TerNr As Long
Dim MasNr As Long
Dim AkTim As Date
Dim StaZe As Date
Dim EndZe As Date
Dim TeBet As Double
Dim TagWe As String
Dim AkDat As String
Dim ZeStr As String
Dim AktZa As Integer
Dim AnzVo As Integer
Dim MiDif As Integer
Dim SelDf As Integer
Dim ZeiRa As Integer
Dim FltTy As Integer
Dim FltId As Integer
Dim NotVa As Integer
Dim SerZa As Integer
Dim MiGef As Boolean
Dim MaGef As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCo1 As XtremeCommandBars.CommandBarComboBox
Dim CmCo2 As XtremeCommandBars.CommandBarComboBox
Dim DaPi1 As XtremeCalendarControl.DatePicker
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmTermVo
Set TxID0 = FM.txtID0
Set TxIDS = FM.txtIdSer
Set CmBet = FM.txtBetre
Set TxOrt = FM.txtRaum1
Set TxAdr = FM.txtAdres
Set CmRmu = FM.cmbRaum1
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set TxFar = FM.txtFarbe
Set TxDa1 = FM.txtDatu1
Set TxDa4 = FM.txtDatu4
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set CmNot = FM.cmbNotVa
Set ZyEnT = FM.cmbZyEn1
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set DaPi1 = frmMain.dtpDatu7
Set RpCo3 = frmMain.repCont3
Set RpCo4 = frmMain.repCont4

CmAcs(AD_Termin_Ketten).Enabled = False
CmAcs(AD_Termin_StaKett).Enabled = False
CmAcs(AD_Termin_Abrechnen).Enabled = False
CmAcs(AD_Termin_EintLoe).Enabled = False

Set CmCo1 = frmMain.comBar01.FindControl(CmCo1, SY_TE_Termin_FiltTyp, , True)
Set CmCo2 = frmMain.comBar01.FindControl(CmCo2, SY_TE_Termin_FiltIdx, , True)

FltTy = CmCo1.ListIndex
FltId = CmCo2.ListIndex - 1

AnzVo = ZyEnT.ItemData(ZyEnT.ListIndex)

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    NotVa = GlMiT(GlSMo, 39) 'Standardmitarbeiter Online-Terminbuchungs Sytem
Else
    NotVa = GlMaT(GlSMa, 25)
End If

If NotVa = 0 Then
    NotVa = 24
End If

If PatNr > 0 Then
    AkDat = Date
Else
    If GlBut = RibTab_Ter_Listen Then
        If DaPi1.Selection.BlocksCount > 0 Then
            AkDat = DaPi1.Selection.Blocks(0).DateBegin
        Else
            AkDat = Date
        End If
    ElseIf GlBut = RibTab_Ter_Akont Then
        If DaPi1.Selection.BlocksCount > 0 Then
            AkDat = DaPi1.Selection.Blocks(0).DateBegin
        Else
            AkDat = Date
        End If
    ElseIf GlBut = RibTab_Ter_WarZi Then
        If DaPi1.Selection.BlocksCount > 0 Then
            AkDat = DaPi1.Selection.Blocks(0).DateBegin
        Else
            AkDat = Date
        End If
    ElseIf GlBut = RibTab_Rechnungen Then
        Set RpCo7 = frmKatRC.repCont7
        Set RpRws = RpCo7.Rows
        SerZa = RpRws.Count
        If SerZa > 0 Then
            Set RpCls = RpCo7.Columns
            Set RpSel = RpCo7.SelectedRows
        Else
            Set RpCls = RpCo4.Columns
            Set RpRws = RpCo4.Rows
            Set RpSel = RpCo4.SelectedRows
        End If
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                If SerZa > 0 Then
                    Set RpCol = RpCls.Find(Ter_ID0)
                    PatNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_ID2)
                    TerNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_IDP)
                    ManNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_IDM)
                    MitNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_MasTer)
                    MasNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Ter_TerBet)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        TeBet = Round(RpRow.Record(RpCol.ItemIndex).Value, 2)
                    Else
                        TeBet = 0
                    End If
                    Set RpCol = RpCls.Find(Ter_Patient)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        PatNa = RpRow.Record(RpCol.ItemIndex).Value
                    Else
                        PatNa = vbNullString
                    End If
                    DoEvents
                    Tr_VoRe TerNr
                    DoEvents
                    Ter_VoL
                Else
                    Set RpCol = RpCls.Find(Rec_ID0)
                    PatNr = RpRow.Record(RpCol.ItemIndex).Value
                    Set RpCol = RpCls.Find(Rec_IDKurz)
                    PatNa = RpRow.Record(RpCol.ItemIndex).Value
                End If
                CmAcs(AD_Termin_Ketten).Enabled = True
                CmAcs(AD_Termin_StaKett).Enabled = True
                CmAcs(AD_Termin_Abrechnen).Enabled = True
                CmAcs(AD_Termin_EintLoe).Enabled = True
            End If
        End If
        AkDat = Date
    ElseIf GlBut = RibTab_Abrechnung Then
        Set RpCls = RpCo3.Columns
        Set RpRws = RpCo3.Rows
        Set RpSel = RpCo3.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            If RpRow.GroupRow = False Then
                Set RpCol = RpCls.Find(Rec_ID0)
                PatNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_IDKurz)
                PatNa = RpRow.Record(RpCol.ItemIndex).Value
                CmAcs(AD_Termin_Ketten).Enabled = True
                CmAcs(AD_Termin_StaKett).Enabled = True
                CmAcs(AD_Termin_Abrechnen).Enabled = True
                CmAcs(AD_Termin_EintLoe).Enabled = True
            End If
        End If
        AkDat = Date
    Else
        CmRmu.ListIndex = GlTRx
        TagWe = Mid$(CmRmu.Tag, 2, Len(CmRmu.Tag) - 1)
        CmRmu.Tag = 1 & TagWe
            
        ManNr = GlMan(GlSMa, 2) 'Aktive Mandanten
        For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
            If ManNr = GlMaT(AktZa, 2) Then
                MaGef = True
                Exit For
            End If
        Next AktZa
        If MaGef = True Then
            CmMan.ListIndex = AktZa - 1
        Else
            CmMan.ListIndex = GlSMa - 1
        End If
        TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
        CmMan.Tag = 1 & TagWe
                
        MitNr = GlMiA(GlSmI, 2)
        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
            If MitNr = GlMiT(AktZa, 2) Then
                MiGef = True
                Exit For
            End If
        Next AktZa
        If MiGef = True Then
            CmMit.ListIndex = AktZa - 1
        Else
            CmMit.ListIndex = GlSmI - 1
        End If
        TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
        CmMit.Tag = 1 & TagWe

        If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
            If MiGef = True Then
                If GlMiT(AktZa, 8) > 0 Then
                    ZeiRa = GlMiT(AktZa, 8)
                Else
                    ZeiRa = GlZeR 'Zeitrasterindex
                End If
            Else
                If GlMiT(0, 8) <> vbNullString Then
                    ZeiRa = GlMiT(0, 8)
                Else
                    ZeiRa = GlZeR 'Zeitrasterindex
                End If
            End If
        Else
            If MaGef = True Then
                If GlMaT(AktZa, 8) > 0 Then
                    ZeiRa = GlMaT(AktZa, 8)
                Else
                    ZeiRa = GlZeR 'Zeitrasterindex
                End If
            Else
                If GlMaT(0, 8) <> vbNullString Then
                    ZeiRa = GlMaT(0, 8)
                Else
                    ZeiRa = GlZeR 'Zeitrasterindex
                End If
            End If
        End If

        If Format$(GlSel.DaSta, "hh:mm:ss") = "00:00:00" Then 'Markierte Celle im Kalender
            If Not IsDate(Format$(GlSel.DaSta, "dd.mm.yyyy")) Then
                GlSel.DaSta = Format$(Now, "dd.mm.yyyy") & Chr$(32) & Format$(Now, "hh:mm:ss")
                AkDat = Format$(Now, "dd.mm.yyyy")
            Else
                AkDat = Format$(GlSel.DaSta, "dd.mm.yyyy")
            End If
        Else
            AkDat = Format$(GlSel.DaSta, "dd.mm.yyyy")
        End If
    End If
End If

If AkDat < Date Then
    AkDat = Date
End If

If PatNr > 0 Then
    TagWe = Mid$(TxID0.Tag, 2, Len(TxID0.Tag) - 1)
    TxID0.Tag = 1 & TagWe
    TxID0.Text = PatNr
    TagWe = Mid$(TxAdr.Tag, 2, Len(TxAdr.Tag) - 1)
    TxAdr.Tag = 1 & TagWe
    TxAdr.Text = PatNa
Else
    TxID0.Text = vbNullString
    TxAdr.Text = vbNullString
End If
CmBet.Text = vbNullString
TxOrt.Text = vbNullString
TxIDS.Text = 5
TxDa1.Text = AkDat

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If GlMiA(GlTBx + 1, 8) > 0 Then
        ZeiRa = GlMiA(GlTBx + 1, 8)
    Else
        ZeiRa = GlMiA(GlSmI, 8)
    End If
Else
    If GlMan(GlTBx + 1, 8) > 0 Then
        ZeiRa = GlMan(GlTBx + 1, 8)
    Else
        ZeiRa = GlMan(GlSMa, 8)
    End If
End If

MiDif = GlTku(ZeiRa, 2)
DoEvents
SRast ZeiRa
DoEvents

If GlSel.DaSta > 0 Then
    If Format$(GlSel.DaSta, "hh:mm") = "00:00" Then
        VoZei.Text = "08:00"
        BiZei.Text = Format$(DateAdd("n", MiDif, "08:00"), "hh:mm")
    Else
        ZeStr = Format$(GlSel.DaSta, "hh:mm")
        For AktZa = 1 To UBound(GlRas)
            If TimeValue(GlRas(AktZa)) <= TimeValue(ZeStr) Then
                StaZe = TimeValue(GlRas(AktZa))
            End If
        Next AktZa
        If GlSSt = True Then 'Starre Termintaktung
            EndZe = DateAdd("n", MiDif, StaZe)
        Else
            SelDf = DateDiff("n", Format$(StaZe, "hh:mm"), Format$(GlSel.DaEnd, "hh:mm"))
            If SelDf < MiDif Then
                EndZe = DateAdd("n", MiDif, StaZe)
            Else
                EndZe = Format$(GlSel.DaEnd, "hh:mm")
            End If
        End If
        VoZei.Text = Format$(StaZe, "hh:mm")
        BiZei.Text = Format$(EndZe, "hh:mm")
    End If
Else
    VoZei.Text = "08:00"
    BiZei.Text = Format$(DateAdd("n", MiDif, "08:00"), "hh:mm")
End If

TxFar.Text = 1
TagWe = Mid$(TxFar.Tag, 2, Len(TxFar.Tag) - 1)
TxFar.Tag = 1 & TagWe

If GlTeE = True Then 'Email-Termin-Erinnerung
    CmNot.ListIndex = NotVa
Else
    RetWe = SendMessage(CmNot.hwnd, CB_SETCURSEL, NotVa, ByVal 0&)
End If

If GlBut = RibTab_Abrechnung Or GlBut = RibTab_Rechnungen Then
    CmRmu.ListIndex = GlTRx
    TagWe = Mid$(CmRmu.Tag, 2, Len(CmRmu.Tag) - 1)
    CmRmu.Tag = 1 & TagWe

    CmMan.ListIndex = GlSMa - 1
    ManNr = GlMan(GlSMa, 2) 'Aktive Mandanten
    For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
        If ManNr = GlMaT(AktZa, 2) Then
            CmMan.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
    TagWe = Mid$(CmMan.Tag, 2, Len(CmMan.Tag) - 1)
    CmMan.Tag = 1 & TagWe

    MitNr = GlMiA(GlSmI, 2)
    For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
        If MitNr = GlMiT(AktZa, 2) Then
            CmMit.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
    TagWe = Mid$(CmMit.Tag, 2, Len(CmMit.Tag) - 1)
    CmMit.Tag = 1 & TagWe
End If

CmSta.Pane(1).Text = "Neuer Terminvorschlag ab: " & Format$(AkDat, "dddd" & ", " & "dd" & ". " & "mmmm" & Chr$(32) & "yyyy") & " um: " & Format$(VoZei.Text, "hh:mm") & " Uhr"
GlTem = -1

Set DaPi1 = Nothing
Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

NeErr:
If GlDbg = True Then SErLog Err.Description & " TeVoNe " & Err.Number
Resume Next

End Sub
Private Sub TeVoOp()
On Error GoTo NeErr

Dim OpZy1 As XtremeSuiteControls.RadioButton
Dim OpZy2 As XtremeSuiteControls.RadioButton
Dim OpZy3 As XtremeSuiteControls.RadioButton
Dim OpZy4 As XtremeSuiteControls.RadioButton

Set FM = frmTermVo
Set ChDop = FM.chkDopTe
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set OpZy1 = FM.optZykl1
Set OpZy2 = FM.optZykl2
Set OpZy3 = FM.optZykl3
Set OpZy4 = FM.optZykl4

Select Case GlTvM
Case "M1": Rahm1.Visible = True
           Rahm2.Visible = False
           Rahm3.Visible = False
           Rahm4.Visible = False
           ChDop.Enabled = True
Case "M2": Rahm1.Visible = False
           Rahm2.Visible = True
           Rahm3.Visible = False
           Rahm4.Visible = False
           ChDop.Enabled = True
Case "M3": Rahm1.Visible = False
           Rahm2.Visible = False
           Rahm3.Visible = True
           Rahm4.Visible = False
           ChDop.Enabled = False
           ChDop.Value = xtpUnchecked
Case "M4": Rahm1.Visible = False
           Rahm2.Visible = False
           Rahm3.Visible = False
           Rahm4.Visible = True
           ChDop.Enabled = False
           ChDop.Value = xtpUnchecked
End Select

Select Case GlTvM
Case "M1": OpZy1.Value = True
           OpZy2.Value = False
           OpZy3.Value = False
           OpZy4.Value = False
Case "M2": OpZy1.Value = False
           OpZy2.Value = True
           OpZy3.Value = False
           OpZy4.Value = False
Case "M3": OpZy1.Value = False
           OpZy2.Value = False
           OpZy3.Value = True
           OpZy4.Value = False
Case "M4": OpZy1.Value = False
           OpZy2.Value = False
           OpZy3.Value = False
           OpZy4.Value = True
End Select

Exit Sub

NeErr:
If GlDbg = True Then SErLog Err.Description & " TeVoOp " & Err.Number
Resume Next

End Sub
Public Sub TeVoPo()
On Error GoTo ReErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmTermVo
Set Rahm8 = FM.frmRahm8
Set Rahm9 = FM.frmRahm9
Set TxOrt = FM.txtRaum1
Set CmBrs = FM.comBar02
Set CmBet = FM.txtBetre
Set CmAbr = FM.cmbAbger
Set TxAdr = FM.txtAdres
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set TxKom = FM.txtKomme
Set CmRmu = FM.cmbRaum1
Set CmMar = FM.cmbTeTyp
Set CmRmu = FM.cmbRaum1
Set RpCo1 = FM.repCont1
Set RpCo6 = FM.repCont6

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    Rahm9.Move ClLin, ClObn, ClBre - 7200, 4700
    Rahm8.Move ClBre - 7200, ClObn, 7200, 4700
    TxAdr.Width = ClBre - 8500
    CmBet.Width = ClBre - 8500
    CmMan.Width = ClBre - 11620
    CmMit.Width = ClBre - 11620
    CmMar.Width = ClBre - 11620
    CmRmu.Width = ClBre - 11620
    CmAbr.Width = ClBre - 11620
    TxKom.Width = ClBre - 8500
    RpCo1.Move ClLin, ClObn, ClBre, ClHoh
    RpCo6.Move ClLin, ClObn + 4800, ClBre, ClHoh - 4800
End If

Set CmBrs = Nothing
Set RpCo1 = Nothing

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " TeVoPo " & Err.Number
Resume Next

End Sub
Private Sub TeVoRe()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "TermVo") = False Then
    If GlFnt = True Then
        xGro = 980
        yGro = 680
    Else
        xGro = 1100
        yGro = 720
    End If

    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)

    IniSetSek "TermVo"
    IniSetVal "TermVo", "FenLin", xPos
    IniSetVal "TermVo", "FenObe", yPos
    IniSetVal "TermVo", "FenBre", xGro
    IniSetVal "TermVo", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " TeVoRe " & Err.Number
Resume Next

End Sub
Private Sub TeVoSp()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTermVo
Set RpCo1 = FM.repCont1
Set RpCo6 = FM.repCont6

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Set RpCls = RpCo1.Columns
With RpCls
    Set RpCol = .Add(TeL_ID1, "ID1", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(TeL_ID0, "ID0", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(TeL_ID2, "ID2", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(TeL_Typ, "Typ", 40, False)
    With RpCol
        .EditOptions.AllowEdit = True
        .EditOptions.AddComboButton
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        If GlTyV = True Then
            For AktZa = 1 To UBound(GlKrA)
                RpCol.EditOptions.Constraints.Add GlKrA(AktZa, 1), GlKrA(AktZa, 0)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(TeL_GOID, "Ziffer", 70, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(TeL_IDKurz, "Bezeichnung", 10, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = True
        .Groupable = False
        .Resizable = True
        .Sortable = False
        .AutoSize = True
    End With
    If RpCo1.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(TeL_Anz, "x", 30, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Multi, "Faktor", 40, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Betrag, "Einzel", 60, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Gesamt, "Gesamt", 60, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
        If Left$(GlVar, 1) = "M" Then .Visible = False
    End With
    Set RpCol = .Add(TeL_Selekt, vbNullString, 20, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Printer_Ink
        .Tag = 1
    End With
End With

With RpCo6
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Set RpCls = RpCo6.Columns
With RpCls
    Set RpCol = .Add(0, vbNullString, 0, False)
    RpCol.Icon = IC16_Calendar_Day
    Set RpCol = .Add(1, "Wochentag", 0, False)
    Set RpCol = .Add(2, "Datum", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .EditOptions.AddComboButton
        .EditOptions.AllowEdit = True
    End With
    Set RpCol = .Add(3, "Von", 0, True)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(4, "Bis", 0, True)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    Set RpCol = .Add(5, vbNullString, 0, False)
    With RpCol
        .Alignment = xtpAlignmentIconCenter
        .HeaderAlignment = xtpAlignmentCenter
        .Icon = IC16_Calendar_Disk
    End With
    Set RpCol = .Add(6, "Farbe", 0, True)
    Set RpCol = .Add(7, "Betreff", 0, True)
    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        Set RpCol = .Add(8, "Mitarbeiter", 0, True)
    Else
        Set RpCol = .Add(8, "Mandant", 0, True)
    End If
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.AllowEdit = True
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleReadOnly
        If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
            For AktZa = 1 To UBound(GlMiK)
                .EditOptions.Constraints.Add GlMiK(AktZa, 1), GlMiK(AktZa, 2)
            Next AktZa
        Else
            For AktZa = 1 To UBound(GlMan)
                .EditOptions.Constraints.Add GlMan(AktZa, 1), GlMan(AktZa, 2)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(9, "Raumplan", 0, True)
    With RpCol
        .EditOptions.AddComboButton
        .EditOptions.AllowEdit = True
        .EditOptions.ConstraintEdit = True
        .EditOptions.EditControlStyle = xtpEditStyleReadOnly
        If GlRaV = True Then
            For AktZa = 1 To UBound(GlRmu)
                .EditOptions.Constraints.Add GlRmu(AktZa, 1), GlRmu(AktZa, 2)
            Next AktZa
        End If
    End With
    Set RpCol = .Add(10, "Nummer", 0, True)
    RpCol.HeaderAlignment = xtpAlignmentCenter
    If GlTeE = True Then 'Email-Termin-Erinnerung
        Set RpCol = .Add(11, "Benachrichtigung", 0, False)
    End If
End With

For Each RpCol In RpCls
    RpCol.Groupable = True
    RpCol.Sortable = False
    RpCol.AutoSize = False
    RpCol.Resizable = False
    RpCol.AutoSortWhenGrouped = False
    RpCol.Editable = False
Next RpCol

RpCls(0).Width = 20
RpCls(1).Width = 80
RpCls(2).Width = 90
RpCls(2).Editable = True
RpCls(3).Width = 50
RpCls(3).Editable = True
RpCls(4).Width = 50
RpCls(5).Width = 40
RpCls(5).Editable = True
RpCls(6).Width = 0
RpCls(7).Width = 140
RpCls(7).Editable = True
RpCls(8).Width = 140
RpCls(9).Width = 90
RpCls(9).Editable = True
RpCls(10).Width = 80
If GlTeE = True Then 'Email-Termin-Erinnerung
    RpCls(11).Width = 120
End If

RpCls(7).AutoSize = True

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo6 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " TeVoSp " & Err.Number
Resume Next

End Sub
Public Sub TeVoSt()
On Error GoTo KoErr
'Errechnet das wirkliche Startdatum

Dim AnwDa As Date
Dim StaDa As Date
Dim TmpDa As Date
Dim WoTag As Integer
Dim DifTa As Integer
Dim DifWo As Integer
Dim DifMo As Integer
Dim StaTa As Integer
Dim TagNr As Integer
Dim Monat As Integer

Set FM = frmTermVo
Set ChDop = FM.chkDopTe
Set TxDa1 = FM.txtDatu1
Set TxDa4 = FM.txtDatu4
Set TxDa5 = FM.txtDatu5
Set ZyTag = FM.txoTage1
Set FoZy1 = FM.optZykl1
Set FoZy2 = FM.optZykl2
Set FoZy3 = FM.optZykl3
Set FoZy4 = FM.optZykl4
Set TaZy1 = FM.optZyTa1
Set TaZy2 = FM.optZyTa2
Set MoZy1 = FM.optZyMo1
Set MoZy2 = FM.optZyMo2
Set JaZy1 = FM.optZyJa1
Set JaZy2 = FM.optZyJa2
Set ZyEn2 = FM.optZyEn2
Set ZyEn3 = FM.optZyEn3
Set ZyWho = FM.cmbWoche
Set ZyMo1 = FM.cmoMona1
Set ZyMo2 = FM.cmoMona2
Set ZyMe1 = FM.cmbMona1
Set ZyMe2 = FM.cmbMonat
Set ZyMe3 = FM.cmbMona3
Set ZyJa1 = FM.cmoJahr1
Set ZyJa2 = FM.cmoJahr2
Set ZyJa3 = FM.cmoJahr3
Set ZyJa4 = FM.cmoJahr4
Set ZyJe1 = FM.cmbJahr1
Set ZyTer = FM.cmbZyEn1
Set ChMon = FM.choTaMon
Set ChDin = FM.choTaDin
Set ChMit = FM.choTaMit
Set ChDon = FM.choTaDon
Set ChFre = FM.choTaFre
Set ChSam = FM.choTaSam
Set ChSon = FM.choTaSon

StaDa = CDate(TxDa1.Text)
AnwDa = S_TeDx(GlTVo, "VonDat")

If FoZy1.Value = True Then
    If TaZy1.Value = True Then
        If StaDa = AnwDa Then
            If ChDop.Value = xtpChecked Then
                StaDa = StaDa
            Else
                StaDa = DateAdd("d", 1, StaDa)
            End If
        Else
            StaDa = StaDa
        End If
    ElseIf TaZy2.Value = True Then
        If Weekday(StaDa) = vbSaturday Then
            StaDa = DateAdd("d", 2, StaDa)
        ElseIf Weekday(StaDa) = vbSunday Then
            StaDa = DateAdd("d", 1, StaDa)
        ElseIf StaDa = AnwDa Then
            If ChDop.Value = xtpChecked Then
                StaDa = StaDa
            Else
                StaDa = DateAdd("d", 1, StaDa)
            End If
        End If
    End If
ElseIf FoZy2.Value = True Then
    WoTag = Weekday(StaDa, vbMonday)
    DifWo = ZyWho.ItemData(ZyWho.ListIndex)
    Select Case WoTag
    Case 1: 'Montag
        If ChSon.Value = xtpChecked Then DifTa = 6
        If ChSam.Value = xtpChecked Then DifTa = 5
        If ChFre.Value = xtpChecked Then DifTa = 4
        If ChDon.Value = xtpChecked Then DifTa = 3
        If ChMit.Value = xtpChecked Then DifTa = 2
        If ChDin.Value = xtpChecked Then DifTa = 1
        If ChMon.Value = xtpChecked Then
            If ChSon.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 1
            Else
                If ChDop.Value = xtpChecked Then
                    DifTa = 0
                Else
                    DifTa = 7
                End If
            End If
        End If
    Case 2: 'Dinestag
        If ChMon.Value = xtpChecked Then DifTa = 6
        If ChSon.Value = xtpChecked Then DifTa = 5
        If ChSam.Value = xtpChecked Then DifTa = 4
        If ChFre.Value = xtpChecked Then DifTa = 3
        If ChDon.Value = xtpChecked Then DifTa = 2
        If ChMit.Value = xtpChecked Then DifTa = 1
        If ChDin.Value = xtpChecked Then
            If ChMon.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 1
            Else
                If ChDop.Value = xtpChecked Then
                    DifTa = 0
                Else
                    DifTa = 7
                End If
            End If
        End If
    Case 3: 'Mittwoch
        If ChDin.Value = xtpChecked Then DifTa = 6
        If ChMon.Value = xtpChecked Then DifTa = 5
        If ChSon.Value = xtpChecked Then DifTa = 4
        If ChSam.Value = xtpChecked Then DifTa = 3
        If ChFre.Value = xtpChecked Then DifTa = 2
        If ChDon.Value = xtpChecked Then DifTa = 1
        If ChMit.Value = xtpChecked Then
            If ChDin.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 1
            Else
                If ChDop.Value = xtpChecked Then
                    DifTa = 0
                Else
                    DifTa = 7
                End If
            End If
        End If
    Case 4: 'Donnerstag
        If ChMit.Value = xtpChecked Then DifTa = 6
        If ChDin.Value = xtpChecked Then DifTa = 5
        If ChMon.Value = xtpChecked Then DifTa = 4
        If ChSon.Value = xtpChecked Then DifTa = 3
        If ChSam.Value = xtpChecked Then DifTa = 2
        If ChFre.Value = xtpChecked Then DifTa = 1
        If ChDon.Value = xtpChecked Then
            If ChMit.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 1
            Else
                If ChDop.Value = xtpChecked Then
                    DifTa = 0
                Else
                    DifTa = 7
                End If
            End If
        End If
    Case 5: 'Freitag
        If ChDon.Value = xtpChecked Then DifTa = 6
        If ChMit.Value = xtpChecked Then DifTa = 5
        If ChDin.Value = xtpChecked Then DifTa = 4
        If ChMon.Value = xtpChecked Then DifTa = 3
        If ChSon.Value = xtpChecked Then DifTa = 2
        If ChSam.Value = xtpChecked Then DifTa = 1
        If ChFre.Value = xtpChecked Then
            If ChDon.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChSam.Value = xtpChecked Then
                DifTa = 1
            Else
                If ChDop.Value = xtpChecked Then
                    DifTa = 0
                Else
                    DifTa = 7
                End If
            End If
        End If
    Case 6: 'Samstag
        If ChFre.Value = xtpChecked Then DifTa = 6
        If ChDon.Value = xtpChecked Then DifTa = 5
        If ChMit.Value = xtpChecked Then DifTa = 4
        If ChDin.Value = xtpChecked Then DifTa = 3
        If ChMon.Value = xtpChecked Then DifTa = 2
        If ChSon.Value = xtpChecked Then DifTa = 1
        If ChSam.Value = xtpChecked Then
            If ChFre.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChSon.Value = xtpChecked Then
                DifTa = 1
            Else
                If ChDop.Value = xtpChecked Then
                    DifTa = 0
                Else
                    DifTa = 7
                End If
            End If
        End If
    Case 7: 'Sonntag
        If ChSam.Value = xtpChecked Then DifTa = 6
        If ChFre.Value = xtpChecked Then DifTa = 5
        If ChDon.Value = xtpChecked Then DifTa = 4
        If ChMit.Value = xtpChecked Then DifTa = 3
        If ChDin.Value = xtpChecked Then DifTa = 2
        If ChMon.Value = xtpChecked Then DifTa = 1
        If ChSon.Value = xtpChecked Then
            If ChSam.Value = xtpChecked Then
                DifTa = 6
            ElseIf ChFre.Value = xtpChecked Then
                DifTa = 5
            ElseIf ChDon.Value = xtpChecked Then
                DifTa = 4
            ElseIf ChMit.Value = xtpChecked Then
                DifTa = 3
            ElseIf ChDin.Value = xtpChecked Then
                DifTa = 2
            ElseIf ChMon.Value = xtpChecked Then
                DifTa = 1
            Else
                If ChDop.Value = xtpChecked Then
                    DifTa = 0
                Else
                    DifTa = 7
                End If
            End If
        End If
    End Select
    Select Case DifWo
    Case 1: StaDa = DateAdd("d", DifTa, StaDa)
    Case 2: StaDa = DateAdd("d", DifTa + 7, StaDa)
    Case 3: StaDa = DateAdd("d", DifTa + 14, StaDa)
    Case 4: StaDa = DateAdd("d", DifTa + 21, StaDa)
    End Select
ElseIf FoZy3.Value = True Then
    If MoZy1.Value = True Then
        DifMo = ZyMe2.ItemData(ZyMe2.ListIndex)
        StaTa = ZyMe1.ItemData(ZyMe1.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        TmpDa = CDate(StaTa & "." & Format$(StaDa, "mm") & "." & Format$(StaDa, "yyyy"))
        DifTa = Abs(DateDiff("d", TmpDa, StaDa))
        If StaDa > TmpDa Then
            StaDa = DateAdd("m", 1, TmpDa)
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
            StaDa = DateAdd("m", 1, TmpDa)
        Else
            StaDa = StaDa
        End If
    ElseIf MoZy2.Value = True Then
        TagNr = ZyMo1.ItemData(ZyMo1.ListIndex)
        StaTa = ZyMo2.ItemData(ZyMo2.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        DifMo = ZyMe3.ItemData(ZyMe3.ListIndex)
        TmpDa = WoTaMo(StaTa, Month(StaDa), Year(StaDa), TagNr)
        If StaDa > TmpDa Then
            If Month(StaDa) + 1 > 12 Then
                StaDa = WoTaMo(StaTa, 1, Year(StaDa) + 1, TagNr)
            Else
                StaDa = WoTaMo(StaTa, Month(StaDa) + 1, Year(StaDa), TagNr)
            End If
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
             StaDa = WoTaMo(StaTa, Month(StaDa) + 1, Year(StaDa), TagNr)
        Else
            StaDa = StaDa
        End If
    End If
ElseIf FoZy4.Value = True Then
    If JaZy1.Value = True Then
        StaTa = ZyJe1.ItemData(ZyJe1.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        Monat = ZyJa1.ItemData(ZyJa1.ListIndex)
        TmpDa = CDate(StaTa & "." & Format$(Monat, "00") & "." & Format$(StaDa, "yyyy"))
        If StaDa > TmpDa Then
            StaDa = DateAdd("yyyy", 1, TmpDa)
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
            StaDa = DateAdd("yyyy", 1, TmpDa)
        Else
            StaDa = StaDa
        End If
    ElseIf JaZy2.Value = True Then
        TagNr = ZyJa2.ItemData(ZyJa2.ListIndex)
        StaTa = ZyJa3.ItemData(ZyJa3.ListIndex)
        StaTa = TeVoMo(StaTa, Month(StaDa))
        Monat = ZyJa4.ItemData(ZyJa4.ListIndex)
        TmpDa = WoTaMo(StaTa, Monat, Year(StaDa), TagNr)
        If StaDa > TmpDa Then
            StaDa = WoTaMo(StaTa, Monat, Year(TmpDa) + 1, TagNr)
        ElseIf StaDa < TmpDa Then
            StaDa = TmpDa
        ElseIf TmpDa = AnwDa Then
            StaDa = WoTaMo(StaTa, Monat, Year(TmpDa) + 1, TagNr)
        Else
            StaDa = StaDa
        End If
    End If
End If

TxDa5.Text = StaDa

Exit Sub

KoErr:
If GlDbg = True Then SErLog Err.Description & " TeVoSt " & Err.Number
Resume Next

End Sub
Public Sub TeVoSu()
On Error GoTo KoErr
'Erstellt Terminvorschläge

Dim GeEin As Single
Dim GeSum As Single
Dim AnzTe As Integer

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEd1 As XtremeCommandBars.CommandBarEdit
Dim CmEd2 As XtremeCommandBars.CommandBarEdit
Dim CmEd3 As XtremeCommandBars.CommandBarEdit
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl

Set FM = frmTermVo
Set RpCo6 = FM.repCont6
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set RpRcs = RpCo6.Records

Set CmEd1 = CmBrs.FindControl(CmEd1, AD_Termin_Betrag1, , True)
Set CmEd2 = CmBrs.FindControl(CmEd2, AD_Termin_Betrag2, , True)
Set CmEd3 = CmBrs.FindControl(CmEd3, AD_Termin_Betrag3, , True)

GeSum = 0
GeEin = CSng(CmEd1.Text)
AnzTe = RpRcs.Count
GeSum = GeEin * AnzTe

If GeSum = 0 Then
    CmEd2.Text = CmEd1.Text
Else
    CmEd2.Text = Format$(GeSum, GlWa1)
End If

Set RpRcs = Nothing
Set RpCo6 = Nothing

Exit Sub

KoErr:
If GlDbg = True Then SErLog Err.Description & " TeVoSu " & Err.Number
Resume Next

End Sub
Public Sub TMeAc(ByVal EnAbl As Boolean)
On Error GoTo LaErr
'Schaltet das Menu ein / aus

Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmMain
Set CmBrs = FM.comBar01
Set CmAcs = CmBrs.Actions

CmAcs(ME_Termin_Hinzufuegen).Enabled = EnAbl
CmAcs(ME_Termin_Bearbeiten).Enabled = EnAbl
CmAcs(ME_Termin_Loeschen).Enabled = EnAbl
CmAcs(ME_Termin_Kopieren).Enabled = EnAbl
CmAcs(ME_Termin_Ausschneiden).Enabled = EnAbl

CmAcs(ME_Terminliste_Hinzufuegen).Enabled = EnAbl
CmAcs(ME_Terminliste_Bearbeiten).Enabled = EnAbl
CmAcs(ME_Terminliste_Kopieren).Enabled = EnAbl
CmAcs(ME_Terminliste_Loeschen).Enabled = EnAbl

CmAcs(ME_Raumtermin_Hinzufuegen).Enabled = EnAbl
CmAcs(ME_Raumtermin_Bearbeiten).Enabled = EnAbl
CmAcs(ME_Raumtermin_Kopieren).Enabled = EnAbl
CmAcs(ME_Raumtermin_Loeschen).Enabled = EnAbl

CmAcs(SY_TE_Termin_Hinzufu).Enabled = EnAbl
CmAcs(SY_TE_Termin_Bearbeiten).Enabled = EnAbl
CmAcs(SY_TE_Termin_Loeschen).Enabled = EnAbl
CmAcs(SY_TE_Termin_Kopiere).Enabled = EnAbl
CmAcs(SY_TE_Termin_Duplizieren).Enabled = EnAbl
CmAcs(SY_TE_Termin_Ausschn).Enabled = EnAbl

CmAcs(SY_TL_Terminliste_Hinzufuegen).Enabled = EnAbl
CmAcs(SY_TL_Terminliste_Bearbeiten).Enabled = EnAbl
CmAcs(SY_TL_Terminliste_Loeschen).Enabled = EnAbl
CmAcs(SY_TL_Terminliste_Kopieren).Enabled = EnAbl

Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " TMeAc " & Err.Number
Resume Next

End Sub
Public Sub TUpAb(Optional ByVal KrRow As Long, Optional ByVal TerNr As Long)
On Error GoTo LaErr
'Aktualisiert die Listenansicht

Dim RowFi As Long
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmTermin
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If KrRow > 0 Then
    Set RpRws = RpCon.Rows
    If RpRws.Count > 0 Then
        RowFi = RpCon.TopRowIndex
    End If
End If

DoEvents
If TerNr > 0 Then
     Ter_Lei TerNr, True
Else
    Ter_Lei GlTem, True
End If
DoEvents

If KrRow > 0 Then
    Set RpRws = RpCon.Rows
    If RpRws.Count > 0 Then
        If KrRow <= RpRws.Count Then
            RpCon.TopRowIndex = RowFi
            RpRws.Row(0).Selected = False
            'RpRws.Row(KrRow).EnsureVisible
            'RpRws.Row(KrRow).Selected = True
            'If GlFoc = True Then
            '    Set RpCon.FocusedRow = RpRws.Row(KrRow)
            'End If
        End If
    End If
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RpRws = Nothing
Set RpCls = Nothing
Set RpSel = Nothing
Set RpCon = Nothing

Set clFen = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " TUpAb " & Err.Number
Resume Next

End Sub
Public Sub TVoUp(Optional ByVal KrRow As Long)
On Error GoTo LaErr
'Aktualisiert die Listenansicht

Dim RowFi As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmTermVo
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns
Set RpSel = RpCon.SelectedRows

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If KrRow > 0 Then
    Set RpRws = RpCon.Rows
    If RpRws.Count > 0 Then
        RowFi = RpCon.TopRowIndex
    End If
End If

DoEvents
Ter_VoL
DoEvents

If KrRow > 0 Then
    Set RpRws = RpCon.Rows
    If RpRws.Count > 0 Then
        If KrRow <= RpRws.Count Then
            RpCon.TopRowIndex = RowFi
            RpRws.Row(0).Selected = False
            'RpRws.Row(KrRow).EnsureVisible
            'RpRws.Row(KrRow).Selected = True
            'If GlFoc = True Then
            '    Set RpCon.FocusedRow = RpRws.Row(KrRow)
            'End If
        End If
    End If
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RpRws = Nothing
Set RpCls = Nothing
Set RpSel = Nothing
Set RpCon = Nothing

Set clFen = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " TVoUp " & Err.Number
Resume Next

End Sub
Private Sub WaInit()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmAufga
Set TxCoN = FM.TexCont4
Set RpCon = FM.repCont1
Set MoKa1 = FM.dtpDatu1
Set ImMan = frmMain.imgManag

With TxCoN
    .ViewMode = 3 'Floating Text
    .Alignment = 0
    .AllowDrop = True
    .AllowUndo = True
    .Enabled = True
    .DataTextFormat = 0
    .AutoExpand = False
    .ClipChildren = False
    .ClipSiblings = False
    .ControlChars = False
    .ColumnLineColor = 0
    .BackColor = -2147483643 '16777215
    .BackStyle = 1
    .BaseLine = 2
    .BorderStyle = 0
    .EditMode = 0
    .FontBold = GlXFt.Bold
    .FontItalic = GlXFt.Italic
    .FontUnderline = GlXFt.Underline
    .FontStrikethru = GlXFt.Strikethrough
    .FontName = GlEFt.Name
    .FontSize = GlEFt.SIZE
    .ForeColor = GlFoF
    .FormatSelection = True
    .HeaderFooterStyle = txNoDblClk
    .HideSelection = False
    .InsertionMode = True
    .Language = 49
    .LineSpacing = 110
    .PageViewStyle = txGradientColors
    .PageHeight = 13000
    .PageWidth = 11000
    .PageMarginL = 400
    .PageMarginR = 400
    .PageMarginT = 400
    .PageMarginB = 400
    .PageOrientation = 0
    .PrintColors = True
    .ScrollBars = 3
    .SizeMode = 0
    .SelectionViewMode = 1
    .TabKey = True
    .TextBkColor = 16777215
    .TextFrameMarkerLines = True
    .TableGridLines = True
    .EnableHyperlinks = True
    .ZoomFactor = 100
    .WordWrapMode = 1
End With

With RpCon
    .PaintManager.ColumnStyle = xtpColumnResource
    Select Case GlSty
    Case 8: .VisualTheme = xtpReportThemeOffice2013
    Case 7: .VisualTheme = xtpReportThemeOffice2013
    Case Else: .VisualTheme = xtpReportThemeResource
    End Select
    .AllowColumnRemove = False
    .AllowColumnReorder = False
    .AllowColumnResize = True
    .AllowColumnSort = GlSPS
    .AllowEdit = True
    .AllowEditPreview = False
    .AutoColumnSizing = True 'WICHTIG!
    .Behavior.Scheme = xtpReportBehaviorCodejockDefault
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = False
    .EnableToolTips False
    .EnsureFocusedRowVisible = True
    .FastDeselectMode = False
    .FreezeColumnsCount = 0
    .FocusSubItems = True
    .Icons = ImMan.Icons
    .MultipleSelection = False
    .ShowItemsInGroups = False 'Gruppierungen
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
    If GlGrL = True Then
        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
    Else
        .PaintManager.HorizontalGridStyle = xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridNoLines
    End If
    .PaintManager.GridLineColor = GlGrC
    .PaintManager.FixedRowHeight = False
    .PaintManager.CaptionFont.SIZE = 8
    .PaintManager.CaptionFont.Name = GlTFt.Name
    .PaintManager.PreviewTextFont.SIZE = 8
    .PaintManager.PreviewTextFont.Name = GlTFt.Name
    .PaintManager.SortByText = "Sortieren nach : "
    .PaintManager.SetPreviewIndent 20, -2, 20, 4
    .PaintManager.DrawGridForEmptySpace = True
    .PaintManager.InvertColumnOnClick = True
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .ShowGroupBox = False
    .PreviewMode = False
    .ShowHeader = True
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With MoKa1
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
    .ToolTipText = "Markieren Sie bitte hier die Behandlungstage des Patienten"
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

FM.BackColor = GlBak

Set MoKa1 = Nothing
Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " WaInit " & Err.Number
Resume Next

End Sub
Public Sub WaMain(Optional ByVal SelTa As Long)
On Error GoTo LaErr

If WindowLoad("frmAufga") = True Then
    frmAufga.ZOrder 0
    Exit Sub
End If

GlWLa = True

WaReg

frmAufga.SelTa = SelTa
Load frmAufga

Set FM = frmAufga

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

With clFen
    If GlIdi = True Then 'Idiotenmodus
        If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
            .FeLin = ((GlxGr - GlFeB) / 2) + (GlFeB - 463)
            .FeObn = (GlyGr - GlFeH) / 2
            .FeBre = 470
            .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
        Else
            .FeLin = GlxGr - 503
            .FeObn = 10
            .FeBre = 470
            .FeHoh = GlyGr - 100
        End If
    Else
        If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
            .FeLin = ((GlxGr - GlFeB) / 2) + (GlFeB - 463)
            .FeObn = (GlyGr - GlFeH) / 2
            .FeBre = 455
            .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
        Else
            .FeLin = IniGetVal("Aufgaben", "FenLin")
            .FeObn = IniGetVal("Aufgaben", "FenObe")
            .FeBre = IniGetVal("Aufgaben", "FenBre")
            .FeHoh = IniGetVal("Aufgaben", "FenHoh")
        End If
    End If
End With

WaInit
WaOpn
WaSpl SelTa
S_WaLa SelTa
DoEvents

With clFen
    .FenMov
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

frmAufga.Show
DoEvents
GlWLa = False

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " WaMain " & Err.Number
Resume Next

End Sub
Private Sub WaOpn()
On Error GoTo PoErr

Dim NeuDa As Date
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmDat As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmThe As XtremeCommandBars.CommandBarComboBox

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

Set CmCom = CmBrs.FindControl(CmCom, SY_SuWi1, , True)
Set CmEdi = CmBrs.FindControl(CmEdi, SY_SuWi2, , True)
Set CmThe = CmBrs.FindControl(CmThe, SY_SuWi3, , True)
Set CmDat = CmBrs.FindControl(CmDat, SY_SuDat, , True)

CmDat.Text = Date

CmCom.ListIndex = 1

CmThe.ListIndex = GlMan(GlSMa, 0) - 1

Select Case CmCom.ListIndex
Case 1:
    CmAcs(SY_SuWi2).Visible = False
    CmAcs(SY_SuDat).Visible = True
    CmAcs(SY_SuBut).Visible = True
    CmAcs(SY_SuWi3).Visible = False
Case 2:
    CmAcs(SY_SuWi2).Visible = True
    CmAcs(SY_SuDat).Visible = False
    CmAcs(SY_SuBut).Visible = False
    CmAcs(SY_SuWi3).Visible = False
Case 3:
    CmAcs(SY_SuWi2).Visible = True
    CmAcs(SY_SuDat).Visible = False
    CmAcs(SY_SuBut).Visible = False
    CmAcs(SY_SuWi3).Visible = False
Case 4:
    CmAcs(SY_SuWi2).Visible = False
    CmAcs(SY_SuDat).Visible = False
    CmAcs(SY_SuBut).Visible = False
    CmAcs(SY_SuWi3).Visible = True
End Select

Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " WaOpn " & Err.Number
Resume Next

End Sub
Public Sub WaPosi()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmAufga
Set CmBrs = FM.comBar02
Set RpCon = FM.repCont1
Set TxCoN = FM.TexCont4

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    If ClBre > 10 Then
        If ClHoh > 100 Then
            RpCon.Move 10, ClObn + 10, ClBre - 20, ClHoh - 20
            TxCoN.Move 100, ClObn + 10, ClBre - 110, ClHoh - 20
        End If
    End If
End If

Set CmBrs = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " WaPosi " & Err.Number
Resume Next

End Sub
Private Sub WaReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "Aufgaben") = False Then
    xPos = GlxGr - 500
    yPos = 10
    xGro = 450
    yGro = GlyGr - 100
     
    IniSetSek "Aufgaben"
    IniSetVal "Aufgaben", "FenLin", xPos
    IniSetVal "Aufgaben", "FenObe", yPos
    IniSetVal "Aufgaben", "FenBre", xGro
    IniSetVal "Aufgaben", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " WaReg " & Err.Number
Resume Next

End Sub
Public Sub WaSpl(ByVal SelTa As Long)
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmAufga
Set RpCon = FM.repCont1
Set RpCls = RpCon.Columns

With RpCon
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

Select Case SelTa
Case RibTab_Wart_Wied:
    With RpCls
        Set RpCol = .Add(0, "ID0", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(1, "ID2", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
        Set RpCol = .Add(2, "Datum", 75, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(3, "Uhrzeit", 45, False)
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = True
        End With
        Set RpCol = .Add(4, "Betreff", 100, False)
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
        Set RpCol = .Add(5, "Mitarbeiter", 60, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            If GlMiV = True Then
                For AktZa = 1 To UBound(GlMiK)
                    RpCol.EditOptions.Constraints.Add GlMiK(AktZa, 1), GlMiK(AktZa, 2)
                Next AktZa
            End If
        End With
        Set RpCol = .Add(6, "Erledigt", 50, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentIconCenter
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End With
Case RibTab_Wart_Beha:
    With RpCls
        Set RpCol = .Add(Rec_ID1, "ID1", 0, False)
        Set RpCol = .Add(Rec_ID0, "ID0", 0, False)
        Set RpCol = .Add(Rec_RechNr, "Rechnung", 0, True)
        Set RpCol = .Add(Rec_Datum, "Datum", 0, True)
        RpCol.Groupable = False
        Set RpCol = .Add(Rec_Selekt, "Abgeschlossen", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Type, "T", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Versand, "V", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Betrag, "Betrag", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Bezahlt, "Bezahlt", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Differe, "Differenz", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_IDKurz, "Patient", 0, True)
        Set RpCol = .Add(Rec_Offen, "B", 0, False)
        With RpCol
            .Alignment = xtpAlignmentIconCenter
            .HeaderAlignment = xtpAlignmentCenter
            .Tag = 1
        End With
        Set RpCol = .Add(Rec_Extrageb, "Extrageb.", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        Set RpCol = .Add(Rec_Fallig, "Fälligkeit", 0, True)
        Set RpCol = .Add(Rec_Wahrung, "Währung", 0, False)
        Set RpCol = .Add(Rec_IDR, "Zähler", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_ID3, "ID3", 0, False)
        Set RpCol = .Add(Rec_IDZ, "IDZ", 0, False)
        Set RpCol = .Add(Rec_Versicherer, "Katalog", 0, True)
        Set RpCol = .Add(Rec_Zahlziel, "Zahlungsziel", 0, True)
        Set RpCol = .Add(Rec_Drucken, "Drucken", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_IDW, "IDW", 0, False)
        Set RpCol = .Add(Rec_Symbol, "W", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Faktor, "Faktor", 0, False)
        Set RpCol = .Add(Rec_Ziel, "Ziel", 0, False)
        Set RpCol = .Add(Rec_Kommentar, "Kommentar", 0, False)
        Set RpCol = .Add(Rec_IDP, "Mandant", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Rec_Druckdatum, "Gedruckt", 0, True)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Kopie, "Kopie", 0, False)
        Set RpCol = .Add(Rec_Steuer, "Steuer", 0, False)
        With RpCol
            .Alignment = xtpAlignmentRight
            .HeaderAlignment = xtpAlignmentCenter
        End With
        Set RpCol = .Add(Rec_Monat, "Monat", 0, True)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            .EditOptions.Constraints.Add "Januar", 1
            .EditOptions.Constraints.Add "Februar", 2
            .EditOptions.Constraints.Add "März", 3
            .EditOptions.Constraints.Add "April", 4
            .EditOptions.Constraints.Add "Mai", 5
            .EditOptions.Constraints.Add "Juni", 6
            .EditOptions.Constraints.Add "Juli", 7
            .EditOptions.Constraints.Add "August", 8
            .EditOptions.Constraints.Add "September", 9
            .EditOptions.Constraints.Add "Oktober", 10
            .EditOptions.Constraints.Add "November", 11
            .EditOptions.Constraints.Add "Dezember", 12
        End With
        Set RpCol = .Add(Rec_Termin, "Termins.", 0, True)
        RpCol.Alignment = xtpAlignmentRight
        RpCol.HeaderAlignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Storniert, "Storniert", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_PKU, "PKU", 0, False)
        RpCol.Alignment = xtpAlignmentCenter
        Set RpCol = .Add(Rec_Gruppe, "G", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Beendet, "E", 0, False)
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Rabatt, "Rabatt", 0, False)
        Set RpCol = .Add(Rec_IDM, "Mitarbeiter", 0, False)
        With RpCol
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
        End With
        Set RpCol = .Add(Rec_GuStr, "Gutschrift", 0, False)
        Set RpCol = .Add(Rec_GutNr, "GutNr", 0, False)
        Set RpCol = .Add(Rec_GuiID, "GuiID", 0, False)
        Set RpCol = .Add(Rec_AufNr, "AufNr", 0, False)
        Set RpCol = .Add(Rec_AuStr, "Auftrag", 0, False)
        Set RpCol = .Add(Rec_Formu, "Formular", 0, False)
        Set RpCol = .Add(Rec_OPLoe, "OPL", 0, False)
        RpCol.Alignment = xtpAlignmentIconLeft
        RpCol.Icon = IC16_Pin_Green
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_Lock, "Lock", 0, False)
        RpCol.Alignment = xtpAlignmentIconLeft
        RpCol.Icon = IC16_Lock
        RpCol.Tag = 1
        Set RpCol = .Add(Rec_IDO, "IDO", 0, False)
        Set RpCol = .Add(Rec_RzDat, "RzDat", 0, False)
        Set RpCol = .Add(Rec_RzNum, "RzNum", 0, False)
        Set RpCol = .Add(Rec_RzTex, "RzTex", 0, False)
        Set RpCol = .Add(Rec_Grund, "Grund", 0, False)
        Set RpCol = .Add(Rec_ForID, "FID", 0, False)
        RpCol.Tag = 1
    End With
    
    For Each RpCol In RpCls
        With RpCol
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .Groupable = True
            .Sortable = True
            .AutoSize = False
            .AutoSortWhenGrouped = False
        End With
    Next RpCol
    
    RpCls(Rec_ID1).Width = 0
    RpCls(Rec_ID0).Width = 0
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_RechNr).Width = 140
        RpCls(Rec_Datum).Width = 110
    Else
        RpCls(Rec_RechNr).Width = 110
        RpCls(Rec_Datum).Width = 80
    End If
    RpCls(Rec_Selekt).Width = 0
    RpCls(Rec_Type).Width = 20
    RpCls(Rec_Versand).Width = 20
    RpCls(Rec_Betrag).Width = 75
    RpCls(Rec_Bezahlt).Width = 75
    RpCls(Rec_Differe).Width = 75
    RpCls(Rec_IDKurz).Width = 220
    RpCls(Rec_Offen).Width = 0
    RpCls(Rec_Extrageb).Width = 75
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_Fallig).Width = 110
    Else
        RpCls(Rec_Fallig).Width = 80
    End If
    RpCls(Rec_Wahrung).Width = 0
    RpCls(Rec_IDR).Width = 60
    RpCls(Rec_ID3).Width = 0
    RpCls(Rec_IDZ).Width = 0
    RpCls(Rec_Versicherer).Width = 140
    RpCls(Rec_Zahlziel).Width = 140
    RpCls(Rec_Drucken).Width = 0
    RpCls(Rec_IDW).Width = 0
    RpCls(Rec_Symbol).Width = 30
    RpCls(Rec_Faktor).Width = 0
    RpCls(Rec_Ziel).Width = 0
    RpCls(Rec_Kommentar).Width = 0
    RpCls(Rec_IDP).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_Druckdatum).Width = 110
    Else
        RpCls(Rec_Druckdatum).Width = 80
    End If
    RpCls(Rec_Kopie).Width = 0
    RpCls(Rec_Steuer).Width = 60
    RpCls(Rec_Monat).Width = 0
    RpCls(Rec_Termin).Width = 75
    RpCls(Rec_Storniert).Width = 0
    RpCls(Rec_PKU).Width = 50
    RpCls(Rec_Beendet).Width = 0
    RpCls(Rec_Rabatt).Width = 0
    RpCls(Rec_IDM).Width = 180
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_GuStr).Width = 110
    Else
        RpCls(Rec_GuStr).Width = 80
    End If
    RpCls(Rec_GutNr).Width = 0
    RpCls(Rec_AufNr).Width = 0
    If GlTFt.SIZE > 10 Then
        RpCls(Rec_AuStr).Width = 110
    Else
        RpCls(Rec_AuStr).Width = 80
    End If
    RpCls(Rec_Formu).Width = 120
    RpCls(Rec_OPLoe).Width = 18
    RpCls(Rec_Lock).Width = 18
Case RibTab_Wart_Noti:

End Select

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " WaSpl " & Err.Number
Resume Next

End Sub
Public Function WoTaJa(ByVal WoTag As VbDayOfWeek, ByVal JahZa As Integer, ByVal TagNr As Integer, ByVal Intvl As Integer) As Date()
On Error Resume Next

Dim TeVor() As Date
Dim Monat As Integer
Dim AktZa As Integer
Dim TeGes As Integer

If JahZa = 0 Then JahZa = VBA.Year(Now)

Monat = 1

Select Case Intvl
Case 1: ReDim TeVor(1 To 12)
Case 2: ReDim TeVor(1 To 6)
Case 3: ReDim TeVor(1 To 4)
Case 4: ReDim TeVor(1 To 3)
End Select

TeGes = UBound(TeVor)

For AktZa = 1 To TeGes
    TeVor(AktZa) = WoTaMo(WoTag, Monat, JahZa, TagNr)
    Monat = Monat + Intvl
Next

WoTaJa = TeVor()
    
End Function
Public Function WoTaMo(ByVal WoTag As VbDayOfWeek, ByVal Monat As Integer, ByVal TeJah As Integer, ByVal TagNr As Integer) As Date
On Error Resume Next
    
Dim AktTa As Date
    
If Monat = 0 Then Monat = VBA.Month(Now)

If TeJah = 0 Then TeJah = VBA.Year(Now)

AktTa = DateSerial(TeJah, Monat, 1)

If VBA.Weekday(AktTa) <> WoTag Then
    AktTa = DateAdd("d", (WoTag - VBA.Weekday(AktTa) + 7) Mod 7, AktTa)
End If

Select Case TagNr
Case Is < 0: TagNr = 5
Case 1 To 5:
Case Is > 5: TagNr = 5
End Select

AktTa = DateAdd("ww", TagNr - 1, AktTa)

If VBA.Month(AktTa) <> Monat Then AktTa = DateAdd("ww", -1, AktTa)

WoTaMo = AktTa
    
End Function

