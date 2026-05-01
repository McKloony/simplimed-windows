Attribute VB_Name = "basKetten"
Option Explicit

Private FM As Form
Private TxIdx As VB.TextBox
Private TxDro As VB.TextBox
Private CmSta As XtremeCommandBars.StatusBar
Private TbBar As XtremeCommandBars.TabToolBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmSwi As XtremeCommandBars.StatusBarSwitchPane
Private CmPgs As XtremeCommandBars.StatusBarProgressPane
Private CmPop As XtremeCommandBars.CommandBarPopup
Private CmCop As XtremeCommandBars.CommandBarPopupColor
Private CmBuT As XtremeCommandBars.CommandBarButton
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
Private Rahm2 As XtremeSuiteControls.GroupBox
Private TxBet As XtremeSuiteControls.FlatEdit
Private TxCCM As XtremeSuiteControls.FlatEdit
Private TxFil As XtremeSuiteControls.FlatEdit
Private TxMai As XtremeSuiteControls.FlatEdit
Private CmEmp As XtremeSuiteControls.ComboBox
Private CmBCC As XtremeSuiteControls.ComboBox
Private WeBr1 As XtremeSuiteControls.WebBrowser
Private CoDia As XtremeSuiteControls.CommonDialog
Private PrtPr As XtremeCommandBars.PrintPreview
Private ChCon As XtremeChartControl.ChartControl
Private CaCol As XtremeCalendarControl.CalendarControl
Private LiFld As FolderViewControl.FolderView
Private LiFi4 As FileViewControl.FileView
Private LiFit As FileViewControl.ListItem
Private LiVw4 As XtremeSuiteControls.ListView
Private TrLi5 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private LiNod As FolderViewControl.TreeNode
Private TxCoN As Tx4oleLib.TXTextControl
Private TxRu1 As Tx4oleLib.TXRuler
Private TxRu2 As Tx4oleLib.TXRuler
Private PrtVo As LlViewCtrl

Private MaiMail As EAGetMailObjLib.Mail
Private MaiGeTo As EAGetMailObjLib.Tools
Private MaiAtta As EAGetMailObjLib.Attachment
Private MaiAdCo As EAGetMailObjLib.AddressCollection
Private MaiAdCc As EAGetMailObjLib.AddressCollection
Private MaiAtCo As EAGetMailObjLib.AttachmentCollection

Private clFil As clsFile
Private clFen As clsFenster
Private clLis As clsLisLab
Public Sub EButt(ByVal BuKey As String)
On Error GoTo OpErr
'Lädt eine andere Auswahl in das ListView

Dim IdxNr As Long
Dim KaStr As String
Dim RpCo4 As XtremeReportControl.ReportControl

Set FM = frmKetten
Set RpCo4 = FM.repCont4

GlKeL = True

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

Select Case BuKey
Case "A00": IdxNr = GlGKa(1, 0)
            KaStr = "A" & IdxNr
Case "C00": IdxNr = GlDia(1, 0)
            KaStr = "C" & IdxNr
Case "D00": IdxNr = GlGKa(1, 0)
            KaStr = "D" & IdxNr
Case "F00": KaStr = "F1"
Case "G00": IdxNr = GlLab(1, 0)
            KaStr = "G" & IdxNr
Case "H00": IdxNr = GlLab(1, 0)
            KaStr = "H" & IdxNr
Case "I00": IdxNr = GlMed(1, 0)
            KaStr = "I" & IdxNr
Case "J00": KaStr = "J1"
Case "K00": KaStr = "K3"
Case "L00": IdxNr = GlAnG(1, 0)
            KaStr = "L" & IdxNr
Case "M00": KaStr = "G5"
Case "N00": IdxNr = GlFrB(1, 0)
            KaStr = "N" & IdxNr
Case "P00": IdxNr = GlArt(1, 0)
            KaStr = "p" & IdxNr
Case "Q00": KaStr = "Q1"
End Select

E_Pos

clFen.FenDsk 3
Screen.MousePointer = vbNormal

RpCo4.SetFocus

Set RpCo4 = Nothing

Set clFen = Nothing

GlKeL = False

Exit Sub

OpErr:
If GlDbg = True Then SErLog Err.Description & " EButt " & Err.Number
Resume Next

End Sub
Public Sub EFilt(ByVal Flag As Integer, Optional ByVal SuStr As String, Optional ByVal SuPar As String)
On Error GoTo OrErr
'Filtert bestimmte Einträge heraus

Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

GlKeL = True

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If Flag < 5 Then
    CmAcs(KA_Eint_Favoriten).Checked = False
    GlFav = False
End If

If Flag > 0 Then
    CmAcs(KA_Eint_Vollst).Enabled = True
Else
    CmAcs(KA_Eint_Vollst).Enabled = False
End If

CmAcs(142).Checked = False
CmAcs(153).Checked = False
CmAcs(154).Checked = False
For AktZa = 65 To 90
    CmAcs(AktZa).Checked = False
Next AktZa

Select Case Flag
Case 0: E_Filt 0
Case 1: E_Filt 1, SuStr
Case 2: E_Filt 2, SuStr
Case 3: E_Filt 3, SuStr, SuPar
Case 4: E_Filt 4, SuStr
        Select Case SuStr
        Case "Ä": CmAcs(142).Checked = True
        Case "Ö": CmAcs(153).Checked = True
        Case "Ü": CmAcs(154).Checked = True
        Case Else: CmAcs(Asc(SuStr)).Checked = True
        End Select
Case 5:
        If CmAcs(KA_Eint_Favoriten).Checked = False Then
            E_Filt 5
            CmAcs(KA_Eint_Favoriten).Checked = True
            GlFav = True
        Else
            E_Filt 0
            CmAcs(KA_Eint_Favoriten).Checked = False
            GlFav = False
        End If
End Select

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set CmSta = Nothing
Set CmBrs = Nothing
Set CmOpt = Nothing

Set clFen = Nothing

GlKeL = False

Exit Sub

OrErr:
If GlDbg = True Then SErLog Err.Description & " EFilt " & Err.Number
Resume Next

End Sub
Private Sub EInit()
On Error GoTo InErr
'Initialisiert alle Objekte

Dim RetWe As Long
Dim ZeiUm As Boolean
Dim LiTip As Boolean
Dim ImMan As XtremeCommandBars.ImageManager
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmKetten
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set ImMan = frmMain.imgManag

LiTip = CBool(IniGetVal("Layout", "GrdTip"))
ZeiUm = False

With RpCo4
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
    .BorderStyle = xtpBorderThemedFrame
    .EditOnClick = True
    .EnableToolTips LiTip
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Leistungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Leistungen vorhanden"
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
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .SortedDragDrop = True
    .UnrestrictedDragDrop = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    RetWe = .EnableDragDrop("Ketten", xtpReportAllowDragCopy + xtpReportAllowDrop)
End With

With RpCo5
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
    .PaintManager.NoFieldsAvailableText = "Es sind keine Leistungen vorhanden"
    .PaintManager.NoItemsText = "Es sind keine Leistungen vorhanden"
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
    .PaintManager.AlternativeBackgroundColor = GlZeF
    .PaintManager.UseAlternativeBackground = GlZei
    .ShowGroupBox = False
    .PreviewMode = GlGrV
    .ShowHeader = GlGKo
    .SortedDragDrop = True
    .UnrestrictedDragDrop = False
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
    RetWe = .EnableDragDrop("Ketten", xtpReportAllowDragCopy + xtpReportAllowDrop)
End With

Set RpCo4 = Nothing
Set RpCo5 = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " Einit " & Err.Number
Resume Next

End Sub
Public Sub EMain(ByVal KetNr As Long, Optional ByVal KetNa As String, Optional ByVal KetKu As String, Optional ByVal DroFe As Integer)
On Error GoTo MeErr

Dim TreKy As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

If WindowLoad("frmKetten") = True Then
    Set FM = frmKetten
    FM.ZOrder 0
    Exit Sub
End If

TreKy = Left$(GlNod, 1)

Screen.MousePointer = vbHourglass
DoEvents

EReg

Load frmKetten

Set FM = frmKetten
Set RpCo5 = FM.repCont5
Set TxIdx = FM.txtIdxNr
Set TxDro = FM.txtDopFe

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    .FenDsk 2
    If GlIdi = True Then 'Idiotenmodus
        If GlKeE = True Then 'Kette einfügen
            .FeObn = (GlyGr / 2) - (652 / 2)
            .FeLin = (GlxGr / 2) - (700 / 2)
            .FeBre = 700
            .FeHoh = 652
        Else
            .FeObn = (GlyGr / 2) - (652 / 2)
            .FeLin = (GlxGr / 2) - (900 / 2)
            .FeBre = 900
            .FeHoh = 652
        End If
    Else
        .FeLin = IniGetVal("Ketten", "FenLin")
        .FeObn = IniGetVal("Ketten", "FenObe")
        .FeBre = IniGetVal("Ketten", "FenBre")
        .FeHoh = IniGetVal("Ketten", "FenHoh")
    End If
End With

EInit
EMenu
EOpg KetNa, KetKu
ESpa1
ESpa2
E_Pos
E_Ket KetNr

TxIdx.Text = KetNr
TxDro.Text = DroFe

Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set RbBar = CmBrs.Item(1)
Set RpCls = RpCo5.Columns

If GlKeE = True Then 'Kette einfügen
    RpCls(Ket_Selekt).Width = 40
    If TreKy = "R" Then
        RpCls(Ket_Anz).Width = 120
        RpCls(Ket_Fakto).Width = 180
        RpCls(Ket_Zeit).Width = 80
    End If
    CmBrs.Item(2).Visible = False
    CmBrs.Item(3).Visible = False
    Set RbTab = RbBar.FindTab(RibTab_Ket_Anwe)
    RbTab.Selected = True
    CmAcs(KA_Kett_Ubernehmen).Enabled = True
    CmAcs(KA_Kett_Selekt).Enabled = True
    CmAcs(KA_Kett_Deselekt).Enabled = True
    CmAcs(KA_Kett_Auswahl).Enabled = True
Else
    CmAcs(KA_Kett_Ubernehmen).Enabled = False
    CmAcs(KA_Kett_Selekt).Enabled = False
    CmAcs(KA_Kett_Deselekt).Enabled = False
    CmAcs(KA_Kett_Auswahl).Enabled = False
End If

With clFen
    .FenMov
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    EPosi
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

frmKetten.Show
DoEvents

DoEvents
Screen.MousePointer = vbNormal

GlKeL = False

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " EMain " & Err.Number
Resume Next

End Sub
Private Sub EMenu()
On Error GoTo MeErr
'Erstellt alle Menü- und Tolleisten

Dim RetWe As Long
Dim KeyNa As String
Dim AktZa As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim RbGrp As XtremeCommandBars.RibbonGroup
Dim RbGps As XtremeCommandBars.RibbonGroups
Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmGlo As XtremeCommandBars.CommandBarsGlobalSettings

Set FM = frmKetten
Set CmBrs = FM.comBar02
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
    Set CmAct = .Add(KA_Eint_Favoriten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_KeKur, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_KeNam, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Eint_Vollst, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Edit_Einfuegen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Edit_Entfernen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Edit_NachOben, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Edit_NachUnten, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Ubernehmen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Selekt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Deselekt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Auswahl, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(KA_Kett_Patient, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(142, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(153, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(154, vbNullString, vbNullString, vbNullString, vbNullString)
    For AktZa = 65 To 90
        Set CmAct = .Add(AktZa, vbNullString, vbNullString, vbNullString, vbNullString)
    Next AktZa
End With

With CmSta
    .Font.SIZE = 8
    .Font.Name = GlTFt.Name
    Set CmPan = .AddPane(1)
    CmPan.Text = vbNullString
    CmPan.Style = SBPS_STRETCH
    Set CmPan = .AddPane(2)
    CmPan.Text = vbNullString
    CmPan.Width = 300
    Set CmPan = .AddPane(3)
    CmPan.Text = vbNullString
    CmPan.Width = 100
    .Visible = True
End With

Set RbBar = CmBrs.AddRibbonBar("ToolBar")
Set CmBuT = RbBar.Controls.Add(xtpControlButton, KA_Edit_Hilfe, "Hilfe")
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

'--------------------------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Ket_Edit, "Bearbeiten")
With RbTab
    .id = RibTab_Ket_Edit
    .Visible = True
    .Selected = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Suchen", RibGrp_Ket_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Favoriten, "Favoriten Anzeigen")
With CmCon
    .IconId = IC32_Doc_Check
    .ShortcutText = "F4"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Suchen, "Eintrag Suchen")
With CmCon
    .IconId = IC32_Doc_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Eint_Vollst, "Alle Einträge zeigen")
With CmCon
    .IconId = IC32_Doc_Eye
    .ShortcutText = "F6"
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Ket_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Edit_Einfuegen, "Eintrag Einfügen")
With CmCon
    .IconId = IC32_Doc_Export
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Edit_Entfernen, "Eintrag Entfernen")
With CmCon
    .IconId = IC32_Link_Left
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Edit_NachOben, "Nach Oben verschieben")
With CmCon
    .IconId = IC32_Link_Up
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Edit_NachUnten, "Nach Unten verschieben")
With CmCon
    .IconId = IC32_Link_Down
    .Width = GlRib
End With

Set RbGrp = RbGps.AddGroup("Ausgabe", RibGrp_Ket_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Speichern, "Kette Speichern")
With CmCon
    .IconId = IC32_Disk_Link
    .ShortcutText = "F8"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Drucken, "Kette Drucken")
With CmCon
    .IconId = IC32_Printer_Ink
    .ShortcutText = "F10"
    .Width = GlRib
End With

'--------------------------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Ket_Anwe, "Einfügen")
With RbTab
    .id = RibTab_Ket_Anwe
    .Visible = True
    .Selected = False
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Kette Anwenden", RibGrp_Ket_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Ubernehmen, "Kette Einfügen")
With CmCon
    .IconId = IC32_Nav_Down
    .ShortcutText = "F7"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Selekt, "Alle Einträge Markieren")
With CmCon
    .IconId = IC32_Doc_Check
    .Width = GlRib
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Deselekt, "Keine Einträge Markieren")
With CmCon
    .IconId = IC32_Doc_Uncheck
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlButton, KA_Kett_Patient, "Patient Suchen")
With CmCon
    .IconId = IC32_Patient_View
    .Width = GlRib
    .BeginGroup = True
    If GlBut = RibTab_Abrechnung Or GlBut = RibTab_Kat_Ketten Then
        .Visible = False
    End If
End With

Set RbGrp = RbGps.AddGroup("Kettenkennzeichnung", RibGrp_Ket_Ansicht)
RbGrp.ControlsGrouping = True
If GlKSt <> "GbEi" Then
    RbGrp.Visible = False
End If
Set CmCon = RbGrp.Add(xtpControlLabel, KA_Kett_Caption, vbNullString)
With CmCon
    .ToolTipText = "Bitte wählen Sie eine Kettenkennzeichnung aus"
    .flags = xtpFlagRightAlign
    .IconId = IC16_Link_Norm
End With
Set CmCom = RbGrp.Add(xtpControlComboBox, KA_Kett_Auswahl, vbNullString)
With CmCom
    .CloseSubMenuOnClick = True
    .DropDownListStyle = False
    .DropDownItemCount = 5
    .EditHint = "Kettenkennzeichnung wählen..."
    .ToolTipText = "Bitte wählen Sie eine Kettenkennzeichnung aus"
    .Style = xtpButtonAutomatic
    .ThemedItems = True
    .Width = 300
End With

'--------------------------------------------------------------------------------

Set CmCoS = RbBar.Controls
For Each CmCon In CmCoS
    CmCon.ToolTipText = IniGetVal(KeyNa, CmCon.id)
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
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(10))
    Set CmEdt = .Add(xtpControlEdit, KA_KeKur, "Kettenkürzel:")
    With CmEdt
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Kettenkürzel"
        .ShowLabel = True
        .ToolTipText = "Geben Sie bitte hier bitte das Eingabekürzel für die Kette ein"
        .Width = 170
        .IconId = IC16_Link_Norm
        .Style = xtpButtonIconAndCaption
    End With
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(10))
    Set CmEdt = .Add(xtpControlEdit, KA_KeNam, "Kettenname:")
    With CmEdt
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Ausführlicher Kettenname"
        .ShowLabel = True
        .ToolTipText = "Geben Sie bitte hier bitte das den Namen der Kette ein"
        .Width = 310
        .IconId = IC16_Link_View
        .Style = xtpButtonIconAndCaption
    End With
End With

Set CmBar = CmBrs.Add("ID_ABC", xtpBarBottom)
With CmBar
    .Closeable = False
    .ContextMenuPresent = False
    .Customizable = False
    .CustomizeDialogPresent = False
    .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    .Position = xtpBarBottom
    .SetIconSize 16, 16
    .ShowExpandButton = True
    .ShowTextBelowIcons = False
End With
Set CmCoS = CmBar.Controls
With CmCoS
    Set CmCon = .Add(xtpControlButton, 42, Chr$(42))
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge in der Auswahl"
    End With
    For AktZa = 65 To 90
        Set CmCon = .Add(xtpControlButton, AktZa, Chr$(AktZa))
        With CmCon
            .Style = xtpButtonCaption
            .ToolTipText = "Zeigt alle Einträge, die mit " & Chr$(AktZa) & " beginnen"
        End With
    Next AktZa
    Set CmCon = .Add(xtpControlButton, 142, "Ä")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ä beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 153, "Ö")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ö beginnen"
    End With
    Set CmCon = .Add(xtpControlButton, 154, "Ü")
    With CmCon
        .Style = xtpButtonCaption
        .ToolTipText = "Zeigt alle Einträge, die mit Ü beginnen"
    End With
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

CmAcs(KA_Eint_Vollst).Enabled = False

Set CmPan = Nothing
Set CmSta = Nothing
Set CmOpt = Nothing
Set CmAct = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " EMenu " & Err.Number
Resume Next

End Sub
Public Sub EMov(ByVal Flag As Boolean)
On Error GoTo OrErr
'Verschiebt den Eintrag in der Kette

Dim KetNr As Long
Dim RowNr As Long
Dim AnzPo As Long
Dim GesZa As Long
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpNav As XtremeReportControl.ReportNavigator

Set FM = frmKetten
Set TxIdx = FM.txtIdxNr
Set RpCo5 = FM.repCont5
Set RpSel = RpCo5.SelectedRows
Set RpNav = RpCo5.Navigator

If TxIdx.Text <> vbNullString Then
    If IsNumeric(TxIdx.Text) Then
        If CLng(TxIdx.Text) > 0 Then
            KetNr = CLng(TxIdx.Text)
        Else
            KetNr = 0
        End If
    Else
        KetNr = 0
    End If
Else
    KetNr = 0
End If

AnzPo = RpSel.Count

If AnzPo = 0 Then Exit Sub

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If Flag = True Then
            RowNr = RpRow.Index - 1
            E_Mov KetNr, Flag
            E_Ket KetNr
            If RowNr > 0 Then
                RpNav.MoveToRow RowNr
            Else
                RpNav.MoveFirstRow
            End If
        Else
            RowNr = RpRow.Index + 1
            GesZa = RpCo5.Records.Count
            E_Mov KetNr, Flag
            E_Ket KetNr
            If RowNr <= GesZa Then
                RpNav.MoveToRow RowNr
            Else
                RpNav.MoveLastRow
            End If
        End If
    End If
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RpSel = Nothing
Set RpRow = Nothing
Set RpNav = Nothing
Set RpCo5 = Nothing

Set clFen = Nothing

Exit Sub

OrErr:
If GlDbg = True Then SErLog Err.Description & " EMov " & Err.Number
Resume Next

End Sub
Private Sub EOpg(Optional ByVal KetNa As String, Optional ByVal KetKu As String)
On Error GoTo OpErr

Dim TreKy As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmEdt As XtremeCommandBars.CommandBarEdit
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

TreKy = Left$(GlNod, 1)

If KetNa <> vbNullString Then
    Set CmEdt = CmBrs.FindControl(CmEdt, KA_KeKur, , True)
    CmEdt.Text = KetKu
End If
If KetKu <> vbNullString Then
    Set CmEdt = CmBrs.FindControl(CmEdt, KA_KeNam, , True)
    CmEdt.Text = KetNa
End If

Set CmCom = CmBrs.FindControl(CmCom, KA_Kett_Auswahl, , True)
With CmCom
    .AddItem "Keine Kettenkennzeichnung hinzufügen", 1
    .AddItem "Summation einer Abrechnungsfolge", 2
    .AddItem "Summation einer analogen Abrechnungsfolge", 3
    .AddItem "Summation der Abrechnungsfolge: " & KetNa, 4
    .AddItem "Summation der analogen Abrechnungsfolge: " & KetNa, 5
    .ListIndex = GlKeK
End With

CmSta.Pane(0).Text = KetNa

If GlKeE = True Then 'Kette einfügen
    If TreKy = "R" Then
        CmSta.Pane(1).Text = "Kein Patient gewählt !"
    End If
End If

Set CmSta = Nothing
Set CmBrs = Nothing

Exit Sub

OpErr:
If GlDbg = True Then SErLog Err.Description & " EOpg " & Err.Number
Resume Next

End Sub
Public Sub EPosi()
On Error GoTo PoErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim ClBrH As Long
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set RpCo4 = FM.repCont4
Set RpCo5 = FM.repCont5
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    ClBrH = ClBre / 2
    Select Case RbTab.id
    Case RibTab_Ket_Edit:
        RpCo4.Move ClLin, ClObn, ClBrH, ClHoh
        RpCo5.Move ClBrH, ClObn, ClBrH, ClHoh
    Case RibTab_Ket_Anwe:
        RpCo4.Move ClLin, ClObn, ClBrH, ClHoh
        RpCo5.Move ClLin, ClObn, ClBre, ClHoh
    End Select
End If

Set CmBrs = Nothing
Set RpCo4 = Nothing
Set RpCo5 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " EPosi " & Err.Number
Resume Next

End Sub
Public Sub EPrint(ByVal ForNa As String, ByVal DruVo As Boolean)
On Error GoTo LaErr
'Druckeinleitung

Dim KetNr As Long
Dim FiNam As String
Dim LoNam As String
Dim Formu As Boolean

Set FM = frmKetten
Set TxIdx = FM.txtIdxNr

Set clLis = New clsLisLab
Set clFil = New clsFile

FiNam = GlFrO & S_FoCh(ForNa) 'Formulardaten auslesen

If clFil.FilVor(FiNam) = True Then
    Formu = True
Else
    Formu = False
    SMeFr GlMeT, GlMeM, GlMeI, GlMeF, False, 1, True, FM.hwnd
End If

If Formu = True Then
    KetNr = TxIdx.Text
    If KetNr = 0 Then Exit Sub
    With clLis
        .ForNam = ForNa
        .FilNam = FiNam
        .PfaTmp = GlTmp
        .IndxNr = KetNr
        .DruDia = False
        .DruVor = GlDrV
        .LLPrKa
    End With
End If

Set clFil = Nothing
Set clLis = Nothing

Exit Sub

LaErr:
If GlDbg = True Then SErLog Err.Description & " EPrint " & Err.Number
Resume Next

End Sub
Private Sub EReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "Ketten") = False Then
    xGro = 800
    yGro = 652
        
    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)

    IniSetSek "Ketten"
    IniSetVal "Ketten", "FenLin", xPos
    IniSetVal "Ketten", "FenObe", yPos
    IniSetVal "Ketten", "FenBre", xGro
    IniSetVal "Ketten", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " EReg " & Err.Number
Resume Next

End Sub
Private Sub ESpa1()
On Error GoTo OpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim TreKy As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set RpCo4 = FM.repCont4
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCls = RpCo4.Columns

TreKy = Left$(GlNod, 1)

With RpCls
    Set RpCol = .Add(Ket_Selekt, vbNullString, 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentIconCenter
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_Check
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        .Tag = 1
    End With
    Set RpCol = .Add(Ket_ID0, "ID0", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If TreKy = "R" Then
        Set RpCol = .Add(Ket_GOID, vbNullString, 0, False)
    Else
        Set RpCol = .Add(Ket_GOID, "Ziffer", 80, False)
    End If
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    Set RpCol = .Add(Ket_IDKurz, "Bezeichnung", 10, False)
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
        .AutoSize = True
    End With
    If RpCo4.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Set RpCol = .Add(Ket_Anz, "Anz", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Ket_Fakto, "Fakt.", 0, False)
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        Select Case GlFri
        Case 1: .Visible = True  'Arzt (GOÄ)
        Case 2: .Visible = False 'Heilpraktiker (GebüH)
        Case 3: .Visible = True  'Zahnarzt (GOZ)
        Case 4: .Visible = False 'Veterinär (GOT)
        Case 5: .Visible = False 'Naturheilpraktiker (Tarif 590)
        Case 6: .Visible = False 'Physiotherapeut
        Case 7: .Visible = False 'Wahlarzt (AT)
        End Select
    End With
    Select Case TreKy
    Case "F": Set RpCol = .Add(Ket_Preis, vbNullString, 0, False)
    Case "R": Set RpCol = .Add(Ket_Preis, "Zeit", 80, False)
    Case Else: Set RpCol = .Add(Ket_Preis, "Preis", 60, False)
    End Select
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Ket_ID2, "ID2", 0, False) 'Sortierung
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If TreKy = "D" Then 'Gebühren
        Set RpCol = .Add(Ket_IDA, "Typ", 0, False) 'Eintragstyp
    End If
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo4 = Nothing

Exit Sub

OpErr:
If GlDbg = True Then SErLog Err.Description & " ESpa1 " & Err.Number
Resume Next

End Sub
Private Sub ESpa2()
On Error GoTo OpErr
'Stellt Spaltenbreiten und Spaltenposition des GridEx ein

Dim TreKy As String
Dim AktZa As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set RpCo5 = FM.repCont5
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab
Set RpCls = RpCo5.Columns

TreKy = Left$(GlNod, 1)

With RpCls
    Select Case RbTab.id
    Case RibTab_Ket_Edit: Set RpCol = .Add(Ket_Selekt, vbNullString, 0, False)
    Case RibTab_Ket_Anwe: Set RpCol = .Add(Ket_Selekt, vbNullString, 40, False)
    End Select
    With RpCol
        .HeaderAlignment = xtpAlignmentIconCenter
        .Alignment = xtpAlignmentCenter
        .Icon = IC16_Check
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
        .Tag = 1
    End With
    Set RpCol = .Add(Ket_ID0, "ID0", 0, False)
    With RpCol
        .Alignment = xtpAlignmentRight
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If TreKy = "R" Then
        Set RpCol = .Add(Ket_GOID, vbNullString, 0, False)
    Else
        Set RpCol = .Add(Ket_GOID, "Ziffer", 80, False)
    End If
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = True
    End With
    If TreKy = "R" Then
        Set RpCol = .Add(Ket_IDKurz, "Terminbetreff", 10, False)
    Else
        Set RpCol = .Add(Ket_IDKurz, "Bezeichnung", 10, False)
    End If
    With RpCol
        .Alignment = xtpAlignmentLeft
        .Editable = False
        .Groupable = False
        .Resizable = True
        .Sortable = False
        .AutoSize = True
    End With
    If RpCo5.PaintManager.FixedRowHeight = False Then
        RpCol.Alignment = xtpAlignmentLeft Or xtpAlignmentWordBreak
        If RpCol.Editable = True Then
            RpCol.EditOptions.EditControlStyle = xtpEditStyleMultiline Or xtpEditStyleAutoVScroll
        End If
    End If
    Select Case TreKy
    Case "F":
        Set RpCol = .Add(Ket_Anz, vbNullString, 0, False)
    Case "R":
        Set RpCol = .Add(Ket_Anz, "Raum", 0, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentLeft
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            If GlRaV = True Then
                For AktZa = 1 To UBound(GlRmu)
                    .EditOptions.Constraints.Add GlRmu(AktZa, 1), GlRmu(AktZa, 2)
                Next AktZa
            End If
        End With
    Case Else:
        Set RpCol = .Add(Ket_Anz, "Anz", 60, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentCenter
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End Select
    Select Case TreKy
    Case "D":
        Set RpCol = .Add(Ket_Fakto, "Fakt.", 50, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentCenter
            .Editable = True
            .Groupable = False
            .Resizable = False
            .Sortable = False
            Select Case GlFri
            Case 1: .Visible = True 'Arzt (GOÄ)
            Case 2: .Visible = True 'Heilpraktiker (GebüH)
            Case 3: .Visible = True 'Zahnarzt (GOZ)
            Case 4: .Visible = False 'Veterinär (GOT)
            Case 5: .Visible = False 'Naturheilpraktiker (Tarif 590)
            Case 6: .Visible = False 'Physiotherapeut
            Case 7: .Visible = False 'Wahlarzt (AT)
            End Select
        End With
    Case "R":
        Set RpCol = .Add(Ket_Fakto, "Mitarbeiter", 0, False)
         With RpCol
            .HeaderAlignment = xtpAlignmentLeft
            .Alignment = xtpAlignmentLeft
            .Editable = False
            .EditOptions.AllowEdit = True
            .EditOptions.AddComboButton
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleAutoVScroll
            If GlMiV = True Then 'Mitarbeiter vorhanden + Terminspalte
                For AktZa = 1 To UBound(GlMiT)
                    .EditOptions.Constraints.Add GlMiT(AktZa, 1), GlMiT(AktZa, 2)
                Next AktZa
            End If
        End With
    Case Else:
        Set RpCol = .Add(Ket_Fakto, vbNullString, 0, False)
    End Select
    Select Case TreKy
    Case "F": Set RpCol = .Add(Ket_Preis, vbNullString, 0, False)
    Case "R": Set RpCol = .Add(Ket_Preis, "Zeit", 80, False)
    Case Else: Set RpCol = .Add(Ket_Preis, "Preis", 60, False)
    End Select
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentRight
        .Editable = True
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Ket_ID2, "ID2", 0, False) 'Sortierung
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    Set RpCol = .Add(Ket_IDA, "IDA", 0, False) 'Index
    With RpCol
        .HeaderAlignment = xtpAlignmentCenter
        .Alignment = xtpAlignmentCenter
        .Editable = False
        .Groupable = False
        .Resizable = False
        .Sortable = False
    End With
    If TreKy = "R" Then
        Set RpCol = .Add(Ket_Zeit, "Nachlauf", 0, False)
        With RpCol
            .HeaderAlignment = xtpAlignmentCenter
            .Alignment = xtpAlignmentCenter
            .Editable = False
            .Groupable = False
            .Resizable = False
            .Sortable = False
        End With
    End If
    If TreKy = "D" Then 'Gebühren
        Set RpCol = .Add(Ket_Zeit, "Typ", 0, False) 'Eintragstyp
    End If
End With

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo5 = Nothing

Exit Sub

OpErr:
If GlDbg = True Then SErLog Err.Description & " ESpa2 " & Err.Number
Resume Next

End Sub
Public Sub MaAdr()
On Error GoTo PoErr
'Öffnet die Adresseingabemaske

Dim PatNr As Long
Dim AktZa As Long
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo0 = FM.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If MaAry(Mai_ID0, RpRow.Index) <> vbNullString Then
            PatNr = MaAry(Mai_ID0, RpRow.Index)
            Select Case GlAdO
            Case 0: SReZe PatNr
            Case 1: SKrZe PatNr
            Case 2: AMain PatNr
            End Select
        Else
            SPopu "Keine Adreszuordnung", "Dieser Email wurde noch keine Adresse zugeordnet", IC48_Forbidden
        End If
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo0 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaAdr " & Err.Number
Resume Next

End Sub
Private Function MaChE(ByVal ExTen As String) As Boolean
On Error Resume Next
'Prüfung verbotener Dateien

Dim AusZa As Integer

For AusZa = 1 To UBound(GlAus)
    If LCase(ExTen) = LCase(GlAus(0, AusZa)) Then
        If GlAus(1, AusZa) = 0 Then
            MaChE = True
        Else
            MaChE = False
        End If
        Exit For
    End If
Next AusZa

End Function
Public Sub MaDet()
On Error GoTo PoErr
'Zeigt die Emaildetails

Dim IdxNr As Long
Dim PatNr As Long
Dim AktZa As Long
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo0 = FM.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then

        If MaAry(Mai_IDA, RpRow.Index) <> vbNullString Then
            IdxNr = MaAry(Mai_IDA, RpRow.Index)
        Else
            IdxNr = 0
        End If
        If MaAry(Mai_ID0, RpRow.Index) <> vbNullString Then
            PatNr = MaAry(Mai_ID0, RpRow.Index)
        Else
            PatNr = 0
        End If
        
        Screen.MousePointer = vbHourglass
        DoEvents
        
        S_MaDet IdxNr, PatNr
        
        DoEvents
        Screen.MousePointer = vbNormal
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo0 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaDet " & Err.Number
Resume Next

End Sub
Private Function MaFor(ByRef MaiAdCo As EAGetMailObjLib.AddressCollection, ByVal PrStr As String) As String
On Error Resume Next

Dim TmStr As String
Dim AdAkt As Integer
Dim AdGes As Integer

AdGes = MaiAdCo.Count

If AdGes > 0 Then
    TmStr = "<b>" & PrStr & ":</b> " ' To or Cc
    For AdAkt = 0 To MaiAdCo.Count - 1
        TmStr = TmStr & MaTag(MaiAdCo.Item(AdAkt).Name & " <" & MaiAdCo.Item(AdAkt).Address & ">")
        If (AdAkt < MaiAdCo.Count - 1) Then
            TmStr = TmStr & "; "
        End If
    Next
    MaFor = TmStr & "<br>"
End If

End Function

Private Sub MaInit()
On Error GoTo InErr

Dim FoCol As Long

Set FM = frmMaiView
Set Rahm2 = FM.frmRahm2
Set WeBr1 = FM.WebBrow1
Set TxCoN = FM.TexCont3
Set TxMai = FM.txtMaiTx

FoCol = CInt(IniGetVal("Layout", "MaiFar"))

With WeBr1
    .ScrollBarStyle = xtpScrollBarStandard
    .Silent = True
    .StaticText = False
    .WebBrowserContextMenu = False
    .RegisterAsDropTarget = False
    .Font.SIZE = GlTFt.SIZE
    .Font.Name = GlTFt.Name
End With

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

With TxMai
    .FlatStyle = True
    .Locked = True
    .ShowBorder = False
End With

FM.BackColor = -2147483643
Rahm2.BackColor = -2147483643

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " MaInit " & Err.Number
Resume Next

End Sub
Public Sub MaKat()
On Error GoTo PoErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmKat As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01

Set CmKat = CmBrs.FindControl(CmKat, KA_Mail_KatCombo, , True)

GlMKa = CmKat.ListIndex 'Mailkatalog (1=Posteingang 2=Postaisgang)

IniSetVal "Layout", "MaiKat", GlMKa

SBuLa True

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaKat " & Err.Number
Resume Next

End Sub
Private Sub MaMen()
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
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmPrg As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox
Dim CmEdD As XtremeCommandBars.CommandBarEdit
Dim CmCoZ As XtremeCommandBars.CommandBarComboBox
Dim CmMit As XtremeCommandBars.CommandBarComboBox
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

Set FM = frmMaiView
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

With CmAcs
    Set CmAct = .Add(Tex_ForFet, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForKur, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForUnt, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_ForDur, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrLi, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrRe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_AusrZe, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaVor2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaVor3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaHin2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_FaHin3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(Tex_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Clip1, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Clip2, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Clip3, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Suchen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Prioritaet, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_Notific, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_NoHTML, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_AttOpen, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_AttSave, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_AttExpo, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_AttView, vbNullString, vbNullString, vbNullString, vbNullString)
    Set CmAct = .Add(TX_Mail_AttImpo, vbNullString, vbNullString, vbNullString, vbNullString)
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

Set CmBuT = RbBar.Controls.Add(xtpControlButton, AM_Hilfe, "Hilfe")
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

'--------------------------------------------------------------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Tex_Dokumt, "Vorschau")
With RbTab
    .id = RibTab_Tex_Dokumt
    .Visible = False
    .Selected = True
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Emails", RibGrp_Kat_Bearbeit)
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Antworten, "Email Antworten")
With CmCon
    .IconId = IC32_Mail_Edit
    .Width = GlRib
    .ShortcutText = "F3"
    If GlBut = RibTab_Krankenbla Then .Enabled = False
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Weiterleiten, "Email Weiterleiten")
With CmCon
    .IconId = IC32_Mail_Export
    .Width = GlRib
    .ShortcutText = "F4"
    If GlBut = RibTab_Krankenbla Then .Enabled = False
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Loeschen, "Email Entfernen")
With CmCon
    .IconId = IC32_Mail_Del
    .Width = GlRib
    If GlBut = RibTab_Krankenbla Then .Enabled = False
End With

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Kat_Ansicht)
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_PatSuch, "Patient Zuordnen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_IDCard_View
    If GlBut = RibTab_Krankenbla Then .Enabled = False
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_PatEdit, "Patient Anzeigen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_IDCard_Edit
    If GlBut = RibTab_Krankenbla Then .Enabled = False
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Rechnun, "Rechnungsimport")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Paperclip
    If GlBut = RibTab_Krankenbla Then .Enabled = False
End With

Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Ungelesen, "Email Gelesen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Mail_Check
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Junkmail, "Email Junkmail")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Mail_Lock
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Markieren, "Email Markieren")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Pin_Norm
End With

Set RbGrp = RbGps.AddGroup("Dateianhang", RibGrp_Ket_Suchen)
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_AttOpen, "Anhang Öffnen")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Folder_Paper
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_AttSave, "Anhang Speichern")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Folder_View
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_AttExpo, "Anhang Exportieren")
With CmCon
    .Style = xtpButtonIconAndCaption
    .IconId = IC16_Folder_Export
End With

Set RbGrp = RbGps.AddGroup("Ausgabe", RibGrp_Kat_Ausgabe)
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Erneut, "Erneut Senden")
With CmCon
    .IconId = IC32_Mail_Check
    .Width = GlRib
    If GlMKa = 1 Then .Enabled = False 'Mailkatalog (1=Posteingang 2=Postaisgang)
End With
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Drucken, "Email Drucken")
With CmCon
    .IconId = IC32_Printer_Ink
    .Width = GlRib
    .ShortcutText = "F10"
End With

'--------------------------------------------------------------------------------------------------------------------

Set RbTab = RbBar.InsertTab(RibTab_Tex_Vorlag, "Bearbeiten")
With RbTab
    .id = RibTab_Tex_Vorlag
    .ToolTip = IniGetOpt(KeyNa, RbTab.id)
    .Visible = False
    .Selected = False
End With
Set RbGps = RbTab.Groups

Set RbGrp = RbGps.AddGroup("Bearbeiten", RibGrp_Tex_Dokument)
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Suchen, "Textphrase Suchen")
With CmCon
    .IconId = IC32_Doc_View
    .ShortcutText = "F5"
    .Width = GlRib
End With
Set CmCon = RbGrp.Add(xtpControlSplitButtonPopup, TX_Mail_Anhang, "Dateianhang Anfügen")
With CmCon
    .IconId = IC32_Paperclip
    .Width = GlRib
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, TX_Mail_Anhang, "Dateianhang Anfügen")
    CmCon.IconId = IC16_Paperclip
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, TX_Mail_Vorlage, "Newsvorlage Öffnen")
    CmCon.IconId = IC16_Folder_Paper
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
Set CmCon = RbGrp.Add(xtpControlButton, Tex_ClpInh, "Einfügen")
With CmCon
    .IconId = IC16_Paste
    .Width = GlRib
End With

'-------------------- RibTab_Tex_Vorlag --------------------

Set RbGrp = RbGps.AddGroup("Schriftart", RibGrp_Tex_Schrift)
RbGrp.ControlsGrouping = True
Set CmCom = RbGrp.Add(xtpControlComboBox, Tex_FntAu4, vbNullString)
With CmCom
    .DropDownListStyle = True
    .ThemedItems = True
    .Width = 150
    .Text = GlTFt.Name
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
Set CmCom = RbGrp.Add(xtpControlComboBox, Tex_FntGr4, vbNullString)
With CmCom
    .DropDownListStyle = True
    .ThemedItems = True
    .Width = 50
    .Text = "10"
    .KeyboardTip = "FS"
End With
Set CmBap = CmBrs.Add("ComBar", xtpBarComboBoxGalleryPopup)
Set GalGr = CmBap.Controls.Add(xtpControlGallery, Tex_FontGr, vbNullString)
With GalGr
    .Width = 50
    .Height = 170
    .Resizable = xtpAllowResizeHeight
End With
Set GalGr.Items = GaItS
Set CmCom.CommandBar = CmBap
Set CmCon = RbGrp.Add(xtpControlButton, Tex_ForFet, "Fettdruck")
With CmCon
    .ToolTipText = "Fettdruck"
    .IconId = IC16_Fett
    .BeginGroup = True
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
Set CmCon = RbGrp.Add(xtpControlButton, Tex_FntTif, "Tiefgestellt")
With CmCon
    .ToolTipText = "Tiefgestellt"
    .IconId = IC16_Tief
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_FntHoh, "Hochgestellt")
With CmCon
    .ToolTipText = "Hochgestellt"
    .IconId = IC16_Hoch
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
Set CmCon = CmBap.Controls.Add(xtpControlButton, Tex_FaVor3, "Weitere Farben")
CmCon.Width = 148
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
Set CmCon = CmBap.Controls.Add(xtpControlButton, Tex_FaHin3, "Weitere Farben")
CmCon.Width = 148

'-----

Set RbGrp = RbGps.AddGroup("Absatz", RibGrp_Tex_Absatz)
RbGrp.ControlsGrouping = True

Set CmCon = RbGrp.Add(xtpControlButton, Tex_Aufzah, "Aufzählung")
With CmCon
    .ToolTipText = "Fügt eine Aufzählung ein"
    .IconId = IC16_Aufza
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_Numeri, "Numerierung")
With CmCon
    .ToolTipText = "Fügt eine Numerierung ein"
    .IconId = IC16_Numer
End With

Set CmCon = RbGrp.Add(xtpControlButton, Tex_EinzLi, "Einzug Vergrößern")
With CmCon
    .ToolTipText = "Einzug Vergrößern"
    .IconId = IC16_AbsRe
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_EinzRe, "Einzug Verkleinern")
With CmCon
    .ToolTipText = "Einzug Verkleinern"
    .IconId = IC16_AbsLi
End With

Set CmCon = RbGrp.Add(xtpControlButtonPopup, Tex_Abstan, "Zeilenabstand")
With CmCon
    .ToolTipText = "Zeilenabstand"
    .IconId = IC16_Zeilen
    .BeginGroup = True
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, Tex_ZeiAb1, "1.0 Zeilenabstand")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, Tex_ZeiAb2, "1.2 Zeilenabstand")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, Tex_ZeiAb3, "1.3 Zeilenabstand")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, Tex_ZeiAb4, "1.5 Zeilenabstand")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, Tex_ZeiAb5, "2.0 Zeilenabstand")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, Tex_ZeiAb6, "2.5 Zeilenabstand")
    Set CmCon = .CommandBar.Controls.Add(xtpControlButton, Tex_ZeiAb7, "3.0 Zeilenabstand")
End With

Set CmCon = RbGrp.Add(xtpControlButton, Tex_TexMar, "Absatzmarke")
With CmCon
    .ToolTipText = "Schaltet den die Absatzmarkierung ein/aus"
    .IconId = IC16_Marke
    .BeginGroup = True
End With

Set CmCon = RbGrp.Add(xtpControlButton, Tex_AusrLi, "Linksbündig")
With CmCon
    .ToolTipText = "Richtet den markierten Text linksbündig aus"
    .IconId = IC16_Links
    .BeginGroup = True
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_AusrRe, "Rechtsbündig")
With CmCon
    .ToolTipText = "Richtet den markierten Text rechtsbündig aus"
    .IconId = IC16_Rechts
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_AusrZe, "Zentriert")
With CmCon
    .ToolTipText = "Richtet den markierten Text zentriert aus"
    .IconId = IC16_Zentr
End With
Set CmCon = RbGrp.Add(xtpControlButton, Tex_AusrBl, "Blocksatz")
With CmCon
    .ToolTipText = "Richtet den markierten Text im Blocksatz aus"
    .IconId = IC16_Block
End With

Set CmCon = RbGrp.Add(xtpControlButton, Tex_Absatz, "Absatz")
With CmCon
    .ToolTipText = "Die Absatzeigenschaften Ändern"
    .IconId = IC16_Absatz
    .Style = xtpButtonIconAndCaption
    .BeginGroup = True
End With

'-----

Set RbGrp = RbGps.AddGroup("Ausgabe", RibGrp_Tex_Drucken)
Set CmCon = RbGrp.Add(xtpControlCheckBox, TX_Mail_Prioritaet, "Priorität")
CmCon.ToolTipText = "Diese Email hat eine hohe Priorität"
Set CmCon = RbGrp.Add(xtpControlCheckBox, TX_Mail_Notific, "Bestätigung")
CmCon.ToolTipText = "Für die Email eine Eingangsbestätigung anfordern"
Set CmCon = RbGrp.Add(xtpControlCheckBox, TX_Mail_NoHTML, "Ascii-Textmail")
CmCon.ToolTipText = "Mögliche Dateianhänge sollen automatisch im ZIP Format komprimiert werden"
Set CmCon = RbGrp.Add(xtpControlButton, TX_Mail_Senden, "Email Senden")
With CmCon
    .IconId = IC32_Mail_Earth
    .ShortcutText = "F10"
    .Width = GlRib
    .BeginGroup = True
End With

'--------------------------------------------------------------------------------------------------------------------

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
    .Visible = False
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

'--------------------------------------------------------------------------------------------------------------------

Set CmBar = CmBrs.Add("ID_Anhang", xtpBarTop)
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
    Set CmCon = .Add(xtpControlLabel, SY_Plac1, Space$(1))
    
    Set CmCon = .Add(xtpControlLabel, SY_Cap02, "Dateianhang :")
    With CmCon
        .ToolTipText = "Wählen Sie den gewünschten Dateianhang"
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Paperclip
    End With
    
    Set CmCom = .Add(xtpControlComboBox, SY_SuCm3, vbNullString)
    With CmCom
        .DropDownListStyle = False
        .AutoComplete = False
        .EditStyle = xtpEditStyleLeft
        .EditHint = "Dateianhänge..."
        .ToolTipText = "Bitte wählen Sie den gewünschten Dateianhang"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .Width = 590
    End With
    
    Set CmCon = .Add(xtpControlLabel, SY_Plac2, Space$(1))
    
    Set CmCon = .Add(xtpControlButton, TX_Mail_AttView, "Vorschau")
    With CmCon
        .Style = xtpButtonIconAndCaption
        .IconId = IC16_Folder_Paper
    End With
End With

'--------------------------------------------------------------------------------------------------------------------

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
    .KeyBindings.Add FCONTROL, Asc("Z"), TX_Mail_Clip1
    .KeyBindings.Add FCONTROL, Asc("R"), TX_Mail_Clip2
    .KeyBindings.Add FCONTROL, Asc("W"), TX_Mail_Clip3
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
If GlDbg = True Then SErLog Err.Description & " MaMen " & Err.Number
Resume Next

End Sub
Public Sub MaNav(ByVal NavDi As Integer)
On Error GoTo PoErr
'Öffnet das Emailformular

Dim RowNr As Long
Dim IdxNr As Long
Dim FiNam As String
Dim DaNam As String
Dim TmPfa As String
Dim TmMes As String
Dim TmStr As String
Dim TmHtm As String
Dim TmpNa As String
Dim MaiDa As String
Dim MaiTi As String
Dim MaAbs As String
Dim MaSub As String
Dim TmpEm As String
Dim MaHtm As String
Dim MaHed As String
Dim AtNam As String
Dim AtKom As String
Dim GesZa As Integer
Dim RowZa As Integer
Dim MaIdx As Integer
Dim AdAkt As Integer
Dim AdGes As Integer
Dim AtGes As Integer
Dim AtAkt As Integer
Dim RetWe As Boolean
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmAtt As XtremeCommandBars.CommandBarComboBox
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMaiView
Set WeBr1 = FM.WebBrow1
Set TxBet = FM.txtEmBet
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RbBar = CmBrs.Item(1)
Set RpCo0 = frmMain.repCont0
Set RpRws = RpCo0.Rows
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

Set clFil = New clsFile

Set MaiGeTo = New EAGetMailObjLib.Tools
Set MaiMail = New EAGetMailObjLib.Mail

Set CmAtt = CmBrs.FindControl(CmAtt, SY_SuCm3, , True)

RowZa = RpRws.Count

MaiMail.LicenseCode = "EG-C1653719494-01199-8BVCAA171B372F1D-F797V24BDU55166A"

Screen.MousePointer = vbHourglass
DoEvents

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

Erase GlAtt 'Array Reset

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
    If RpRow.GroupRow = False Then
        Select Case NavDi
        Case 661: 'Zurück
            If RowNr > 0 Then
                MaIdx = RowNr - 1
            End If
        Case 589: 'Hoch
                MaIdx = 0
        Case 662: 'Vor
            If RowNr < RowZa - 1 Then
                MaIdx = RowNr + 1
            End If
        End Select
        RpRws.Row(RowNr).Selected = False
        RpRws.Row(MaIdx).EnsureVisible
        RpRws.Row(MaIdx).Selected = True
        DoEvents

        If MaAry(Mai_MailFile, MaIdx) <> vbNullString Then
        
            IdxNr = MaAry(Mai_IDA, MaIdx)
            MaSub = MaAry(Mai_Subject, RowNr)
            MaiDa = MaAry(Mai_Maildate, MaIdx)
            MaiTi = MaAry(Mai_Mailtime, MaIdx)
            DaNam = MaAry(Mai_MailFile, MaIdx)
            MaAbs = MaAry(Mai_SenderEmail, MaIdx)
            FiNam = GlDpf & "Emails\" & DaNam
            TmPfa = GlDpf & "Emails\Temp\"
            TmpNa = TmPfa & Left$(DaNam, Len(DaNam) - 3) & "htm"
            
            If clFil.FilVor(FiNam) = True Then
                MaiMail.LoadFile FiNam, False
                
                If Err.Number = 0 Then
                    If MaiMail.IsEncrypted = True Then
                        Set MaiMail = MaiMail.Decrypt(Nothing)
                        If Err.Number <> 0 Then
                            SPopu "Emailverschlüsselung", Err.Description, IC48_Forbidden
                        End If
                    End If
                    
                    If MaiMail.IsSigned = True Then
                        MaiMail.VerifySignature
                        If Err.Number <> 0 Then
                            SPopu "Emailsignatur", Err.Description, IC48_Forbidden
                        End If
                    End If
                    
                    Set MaiAdCo = MaiMail.ToList
                    Set MaiAdCc = MaiMail.CcList
                    AdGes = MaiAdCo.Count
                    
                    Set MaiAtCo = MaiMail.AttachmentList
                    AtGes = MaiAtCo.Count
                    
                    For AdAkt = 0 To AdGes - 1
                        If TmpEm = vbNullString Then
                            TmpEm = MaiAdCo.Item(AdAkt).Address
                        Else
                            TmpEm = TmpEm & ";" & MaiAdCo.Item(AdAkt).Address
                        End If
                    Next AdAkt
                    
                    CmSta.Pane(0).Text = Format$(MaiDa, "dd.mm.yyyy") & " - " & Format$(MaiTi, "hh:mm")
                    CmSta.Pane(1).Text = "von : " & MaAbs
                    CmSta.Pane(2).Text = "an : " & TmpEm
                    
                    MaiMail.DecodeTNEF
                    
                    Select Case GlMTx
                    Case 1: 'HTML-Text
                        TmHtm = MaiMail.HtmlBody
                        TmMes = MaiMail.TextBody
                        If TmMes <> vbNullString Then
                            TmStr = Replace(TmMes, vbCrLf, "<br>", 1)
                            MaHtm = "<!doctype html> <html> <head> <meta charset='utf-8'> <title>" & MaSub & "</title> <meta name='generator' content='SimpliMed'> </head> <body> "
                            MaHtm = MaHtm & "<span style='color:#000000;font-family:" & GlEFt.Name & ";font-size:" & Round(GlEFt.SIZE) + 3 & "px;'>" & TmStr & "</span></div> </body> </html>"
                        End If
                        MaHed = MaHed & "<div style=""font-family: 'Consolas', 'Courier New', 'Arial'; font-size: 14px; background-color: #fff;"">"
                        MaHed = MaHed & "<b> Absender: </b> " + MaTag(MaiMail.From.Name & " <" & MaiMail.From.Address & ">") + "<br>"
                        MaHed = MaHed & MaFor(MaiAdCo, " Empfänger")
                        MaHed = MaHed & MaFor(MaiAdCc, " Kopie")
                        MaHed = MaHed & "<b> Betreff: </b>" & MaTag(MaiMail.Subject) & "<br>" & vbCrLf
                    Case 2: 'Ascii-Text
                        TmMes = MaiMail.TextBody
                        If TmMes = vbNullString Then
                            TmMes = S_MaIdx(MaIdx, "Kommentar")
                        End If
                    Case 3: 'Daten-Text
                        TmMes = S_MaIdx(MaIdx, "Kommentar")
                    End Select
                    DoEvents
                    
                    If clFil.FilDir(TmPfa) = False Then
                        MkDir TmPfa
                    End If

                    If AtGes > 0 Then
                        CmAtt.Clear
                        CmBrs.Item(3).Visible = True
                        CmAcs(TX_Mail_AttView).Visible = False
                        MaHed = MaHed & "<b>Anhänge: </b>"
                        For AtAkt = 0 To AtGes - 1
                            Set MaiAtta = MaiAtCo.Item(AtAkt)
                            AtNam = TmPfa & MaiAtta.Name
                            MaiAtta.SaveAs AtNam, True
                            CmAtt.AddItem MaiAtta.Name
                            If AtKom = vbNullString Then
                                AtKom = MaiAtta.Name
                            Else
                                AtKom = AtKom & ";" & MaiAtta.Name
                            End If
                            If Len(MaiAtta.ContentID) > 0 And InStr(TmHtm, MaiAtta.ContentID) > 0 Then
                                TmHtm = Replace$(TmHtm, "cid:" & MaiAtta.ContentID, AtNam)
                            End If
                        Next AtAkt
                        MaHed = MaHed & "<b>Anhänge: </b>" & AtKom & "<br>" & vbCrLf
                        CmAtt.ListIndex = 1
                    Else
                        CmAtt.Clear
                        CmAcs(TX_Mail_AttOpen).Enabled = False
                        CmAcs(TX_Mail_AttSave).Enabled = False
                        CmAcs(TX_Mail_AttExpo).Enabled = False
                        CmBrs.Item(3).Visible = False
                    End If

                    MaHed = MaHed & "</div>"
                    MaHed = "<meta HTTP-EQUIV=""Content-Type"" Content=""text/html; charset=utf-8"">" & MaHed
                    TmHtm = MaHed & "<hr>" & TmHtm
                    
                    If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
                        If TmHtm <> vbNullString Then
                            MaiGeTo.WriteTextFile TmpNa, TmHtm, 65001
                        Else
                            MaiGeTo.WriteTextFile TmpNa, MaHtm, 65001
                        End If
                    End If

                    If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
                        If TmHtm <> vbNullString Then
                            If clFil.FilVor(TmpNa) = True Then
                                WeBr1.Navigate TmpNa
                            Else
                                WeBr1.Navigate "about:" & TmHtm
                            End If
                        ElseIf MaHtm <> vbNullString Then
                            If clFil.FilVor(TmpNa) = True Then
                                WeBr1.Navigate TmpNa
                            Else
                                WeBr1.Navigate "about:" & MaHtm
                            End If
                        ElseIf TmMes <> vbNullString Then
                            WeBr1.Navigate "about:" & TmMes
                        Else
                            WeBr1.Navigate FiNam
                        End If
                    Else
                        If TmMes <> vbNullString Then
                            TxMai.Text = TmMes
                        End If
                    End If
                    DoEvents
                    MaiMail.Clear
                Else
                    WeBr1.Navigate vbNullString
                    SPopu "Dateifehler", Err.Description, IC48_Forbidden
                End If
                DoEvents
                                
                If MaAry(Mai_Gelesen, MaIdx) <> vbNullString Then
                    If CBool(MaAry(Mai_Gelesen, MaIdx)) = False Then
                        Select Case GlMKa 'Mailkatalog (1=Posteingang 2=Postaisgang)
                        Case 1: DBCmEx2 "qryMailInGel", "@IdGel", "@IdxNr", -1, IdxNr
                        Case 2: DBCmEx2 "qryMailOutGel", "@IdGel", "@IdxNr", -1, IdxNr
                        Case 3: DBCmEx2 "qryMailInGel", "@IdGel", "@IdxNr", -1, IdxNr
                        Case 4: DBCmEx2 "qryMailOutGel", "@IdGel", "@IdxNr", -1, IdxNr
                        End Select
                    End If
                Else
                    Select Case GlMKa 'Mailkatalog (1=Posteingang 2=Postaisgang)
                    Case 1: DBCmEx2 "qryMailInGel", "@IdGel", "@IdxNr", -1, IdxNr
                    Case 2: DBCmEx2 "qryMailOutGel", "@IdGel", "@IdxNr", -1, IdxNr
                    Case 3: DBCmEx2 "qryMailInGel", "@IdGel", "@IdxNr", -1, IdxNr
                    Case 4: DBCmEx2 "qryMailOutGel", "@IdGel", "@IdxNr", -1, IdxNr
                    End Select
                End If
            Else
                WeBr1.Navigate vbNullString
                SPopu "Keine Emaildatei", "Die Emaildatei wurde nicht gefunden", IC48_Forbidden
            End If
        End If
    End If
End If

Set clFil = Nothing

DoEvents
Screen.MousePointer = vbNormal

Set MaiAdCo = Nothing
Set MaiAtCo = Nothing
Set MaiGeTo = Nothing
Set MaiMail = Nothing

Set CmBrs = Nothing
Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo0 = Nothing

Set clFil = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaNav " & Err.Number
Resume Next

End Sub
Public Sub MaMain(Optional ByVal MaiNr As Long, Optional ByVal PatNr As Long = 0, Optional ByVal EmEmp As String, Optional ByVal EmBCC As String, Optional ByVal EmTex As String, Optional ByVal EmBet As String, Optional ByVal FiNam As String)
On Error GoTo MeErr

Dim RetWe As Boolean
Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmMaiView") = True Then
    Set FM = frmMaiView
    Unload FM
End If

GlAkK = True

Screen.MousePointer = vbHourglass
DoEvents

MaReg
DoEvents
Load frmMaiView

Set FM = frmMaiView

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    If GlBiA = False Then 'Bildschirmaktualisierung
        clFen.FenDsk 1
    Else
        clFen.FenDsk 2
    End If
    DoEvents
    If GlIdi = True Then 'Idiotenmodus
        If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
            .FeLin = (GlxGr / 2) - (950 / 2)
            .FeObn = (GlyGr - GlFeH) / 2
            .FeBre = 950
            .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
        Else
            .FeLin = (GlxGr / 2) - (950 / 2)
            .FeObn = (GlyGr / 2) - ((GlyGr - (GlyGr / 7)) / 2)
            .FeBre = 950
            .FeHoh = GlyGr - (GlyGr / 7)
        End If
    Else
        If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
            .FeLin = (GlxGr / 2) - (950 / 2)
            .FeObn = (GlyGr - GlFeH) / 2
            .FeBre = 950
            .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
        Else
            .FeLin = IniGetVal("MailView", "FenLin")
            .FeObn = IniGetVal("MailView", "FenObe")
            .FeBre = IniGetVal("MailView", "FenBre")
            .FeHoh = IniGetVal("MailView", "FenHoh")
        End If
    End If
End With

AFont FM
DoEvents
MaInit
MaMen
RetWe = MaOpn(MaiNr, PatNr, EmEmp, EmBCC, EmTex, EmBet, FiNam)
MTxFo
M_MaEd 3

With clFen
    .FenMov
    Set CmBrs = FM.comBar02
    DoEvents
    CmBrs.RecalcLayout
    DoEvents
    CmBrs.PaintManager.RefreshMetrics
    DoEvents
    MaPos
    DoEvents
    .FenDsk 3
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

If RetWe = True Then
    frmMaiView.Show
    GlAkK = False
Else
    GlAkK = False
    Unload FM
End If

DoEvents
Screen.MousePointer = vbNormal

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " MaMain " & Err.Number
Resume Next

End Sub
Private Function MaOpn(Optional ByVal MaiNr As Long = 0, Optional ByVal PatNr As Long = 0, Optional ByVal EmEmp As String, Optional ByVal EmBCC As String, Optional ByVal EmTex As String, Optional ByVal EmBet As String, Optional ByVal AttNa As String) As Boolean
On Error GoTo PoErr
'Öffnet das Emailformular

Dim PaNum As Long
Dim Posi1 As Long
Dim Posi4 As Long
Dim MaNum As Long
Dim ZugZa As Long
Dim RowNr As Long
Dim FiNam As String
Dim DaNam As String
Dim SigNa As String
Dim SigDa As String
Dim TmMes As String
Dim TmStr As String
Dim TmpBe As String
Dim TmpEm As String
Dim TmpBa As String
Dim TmpEi As String
Dim TmPfa As String
Dim TmHtm As String
Dim MaHtm As String
Dim TmpNa As String
Dim TmpDa As String
Dim FilNa As String
Dim SelTx As String
Dim TmpVo As String
Dim MaGui As String
Dim MaiDa As String
Dim MaiTi As String
Dim MaAbs As String
Dim MaSub As String
Dim MaHed As String
Dim AtNam As String
Dim AtKom As String
Dim TmPat As Variant
Dim AktZa As Integer
Dim AtAkt As Integer
Dim AtGes As Integer
Dim AdAkt As Integer
Dim AdGes As Integer
Dim EmkGe As Integer
Dim EmkAk As Integer
Dim Posit As Integer
Dim DaVor As Boolean
Dim SigVo As Boolean
Dim AttVo As Boolean
Dim SeiPo As Variant
Dim AryIt() As String
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmMaT As XtremeCommandBars.CommandBarComboBox
Dim CmAtt As XtremeCommandBars.CommandBarComboBox
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMaiView
Set Rahm2 = FM.frmRahm2
Set WeBr1 = FM.WebBrow1
Set TxCoN = FM.TexCont3
Set TxBet = FM.txtEmBet
Set TxCCM = FM.txtEmCCM
Set TxMai = FM.txtMaiTx
Set CmEmp = FM.cmbEmEmp
Set CmBCC = FM.cmbEmBCC
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RbBar = CmBrs.Item(1)
Set RpCo0 = frmMain.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

Set clFil = New clsFile

Set MaiGeTo = New EAGetMailObjLib.Tools
Set MaiMail = New EAGetMailObjLib.Mail

Set CmAtt = CmBrs.FindControl(CmAtt, SY_SuCm3, , True)
Set CmSwi = CmSta.FindPane(7)
Set CmPgs = CmSta.FindPane(5)

MaiMail.LicenseCode = "EG-C1653719494-01199-8BVCAA171B372F1D-F797V24BDU55166A"

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

GlAnt = 0 'Emilantwort
GlAtV = False
PaNum = PatNr

Erase GlAtt 'Array zurücksetzen

If MaiNr > 0 Then
    MaNum = MaiNr
    If PatNr = 0 Then
        TmPat = S_MaIdx(MaNum, "ID0")
        If TmPat <> vbNullString Then
            PaNum = CLng(TmPat)
        Else
            PaNum = 0
        End If
    Else
        PaNum = PatNr
    End If
    MaGui = S_MaIdx(MaNum, "GuiID")
    MaSub = S_MaIdx(MaNum, "Subject")
    MaiDa = S_MaIdx(MaNum, "Maildate")
    MaiTi = S_MaIdx(MaNum, "Mailtime")
    DaNam = S_MaIdx(MaNum, "MailFile")
    MaAbs = S_MaIdx(MaNum, "SenderEMail")
    FiNam = GlDpf & "Emails\" & DaNam
    TmPfa = GlDpf & "Emails\Temp\"
    If DaNam <> vbNullString Then
        TmpNa = TmPfa & Left$(DaNam, Len(DaNam) - 3) & "htm"
    End If
Else
    If RpSel.Count > 0 Then
        Set RpRow = RpSel(0)
        If RpRow.GroupRow = False Then
            RowNr = RpRow.Index
            MaNum = MaAry(Mai_IDA, RowNr)
            MaGui = MaAry(Mai_GuiID, RowNr)
            MaSub = MaAry(Mai_Subject, RowNr)
            MaiDa = MaAry(Mai_Maildate, RowNr)
            MaiTi = MaAry(Mai_Mailtime, RowNr)
            DaNam = MaAry(Mai_MailFile, RowNr)
            MaAbs = MaAry(Mai_SenderEmail, RowNr)
            FiNam = GlDpf & "Emails\" & DaNam
            TmPfa = GlDpf & "Emails\Temp\"
            TmpNa = TmPfa & Left$(DaNam, Len(DaNam) - 3) & "htm"
            If GlNaT <> 2 Then 'View the Email
                If MaAry(Mai_ID0, RowNr) <> vbNullString Then
                    PaNum = CLng(MaAry(Mai_ID0, RowNr))
                End If
            End If
        End If
    End If
End If

If FiNam <> vbNullString Then
    If clFil.FilVor(FiNam) = True Then
        DaVor = True
    End If
End If

DoEvents
FM.mPaNr = PaNum

Select Case GlNaT 'Mailtyp
Case 1: 'View the Email
    
    CmSwi.Visible = True
    CmPgs.Visible = False

    If DaVor = True Then
        MaiMail.LoadFile FiNam, False
        DoEvents

        If Err.Number = 0 Then
            If MaiMail.IsEncrypted = True Then
                Set MaiMail = MaiMail.Decrypt(Nothing)
                If Err.Number <> 0 Then
                    SPopu "Emailverschlüsselung", Err.Description, IC48_Forbidden
                End If
            End If
            
            If MaiMail.IsSigned = True Then
                MaiMail.VerifySignature
                If Err.Number <> 0 Then
                    SPopu "Emailsignatur", Err.Description, IC48_Forbidden
                End If
            End If
            
            Set MaiAdCo = MaiMail.ToList
            Set MaiAdCc = MaiMail.CcList
            AdGes = MaiAdCo.Count
            
            Set MaiAtCo = MaiMail.AttachmentList
            AtGes = MaiAtCo.Count
            
            For AdAkt = 0 To AdGes - 1
                If TmpEm = vbNullString Then
                    TmpEm = MaiAdCo.Item(AdAkt).Address
                Else
                    TmpEm = TmpEm & ";" & MaiAdCo.Item(AdAkt).Address
                End If
            Next AdAkt
            
            CmSta.Pane(0).Text = Format$(MaiDa, "dd.mm.yyyy") & " - " & Format$(MaiTi, "hh:mm")
            CmSta.Pane(1).Text = "von : " & MaAbs
            CmSta.Pane(2).Text = "an : " & TmpEm
            
            MaiMail.DecodeTNEF

            Select Case GlMTx
            Case 1: 'HTML-Text
                TmHtm = MaiMail.HtmlBody
                TmMes = MaiMail.TextBody
                If TmMes <> vbNullString Then
                    TmStr = Replace(TmMes, vbCrLf, "<br>", 1)
                    MaHtm = "<!doctype html> <html> <head> <meta charset='utf-8'> <title>" & MaSub & "</title> <meta name='generator' content='SimpliMed'> </head> <body> "
                    MaHtm = MaHtm & "<span style='color:#000000;font-family:" & GlEFt.Name & ";font-size:" & Round(GlEFt.SIZE) + 3 & "px;'>" & TmStr & "</span></div> </body> </html>"
                End If
                MaHed = MaHed & "<div style=""font-family: 'Consolas', 'Courier New', 'Arial'; font-size: 14px; background-color: #fff;"">"
                MaHed = MaHed & "<b>Absender: </b> " + MaTag(MaiMail.From.Name & " <" & MaiMail.From.Address & ">") + "<br>"
                MaHed = MaHed & MaFor(MaiAdCo, "Empfänger")
                MaHed = MaHed & MaFor(MaiAdCc, "Kopie")
                MaHed = MaHed & "<b>Betreff: </b>" & MaTag(MaiMail.Subject) & "<br>" & vbCrLf
            Case 2: 'Ascii-Text
                TmMes = MaiMail.TextBody
                If TmMes = vbNullString Then
                    TmMes = S_MaIdx(MaNum, "Kommentar")
                End If
            Case 3: 'Daten-Text
                TmMes = S_MaIdx(MaNum, "Kommentar")
            End Select
            DoEvents

            If clFil.FilDir(TmPfa) = False Then
                MkDir TmPfa
            End If
            
            If AtGes > 0 Then
                CmAtt.Clear
                CmBrs.Item(3).Visible = True
                CmAcs(TX_Mail_AttView).Visible = False
                For AtAkt = 0 To AtGes - 1
                    Set MaiAtta = MaiAtCo.Item(AtAkt)
                    AtNam = TmPfa & MaiAtta.Name
                    MaiAtta.SaveAs AtNam, True
                    CmAtt.AddItem MaiAtta.Name
                    If AtKom = vbNullString Then
                        AtKom = MaiAtta.Name
                    Else
                        AtKom = AtKom & ";" & MaiAtta.Name
                    End If
                    If Len(MaiAtta.ContentID) > 0 And InStr(TmHtm, MaiAtta.ContentID) > 0 Then
                        TmHtm = Replace$(TmHtm, "cid:" & MaiAtta.ContentID, AtNam)
                    End If
                Next AtAkt
                MaHed = MaHed & "<b>Anhänge: </b>" & AtKom & "<br>" & vbCrLf
                CmAtt.ListIndex = 1
            Else
                CmAtt.Clear
                CmAcs(TX_Mail_AttOpen).Enabled = False
                CmAcs(TX_Mail_AttSave).Enabled = False
                CmAcs(TX_Mail_AttExpo).Enabled = False
                CmBrs.Item(3).Visible = False
            End If
            
            MaHed = MaHed & "</div>"
            MaHed = "<meta HTTP-EQUIV=""Content-Type"" Content=""text/html; charset=utf-8"">" & MaHed
            TmHtm = MaHed & "<hr>" & TmHtm

            If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
                If TmpNa <> vbNullString Then
                    If TmHtm <> vbNullString Then
                        MaiGeTo.WriteTextFile TmpNa, TmHtm, 65001
                    Else
                        MaiGeTo.WriteTextFile TmpNa, MaHtm, 65001
                    End If
                End If
            End If

            If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
                If TmHtm <> vbNullString Then
                    If clFil.FilVor(TmpNa) = True Then
                        WeBr1.Navigate TmpNa
                    Else
                        WeBr1.Navigate "about:" & TmHtm
                    End If
                ElseIf MaHtm <> vbNullString Then
                    If clFil.FilVor(TmpNa) = True Then
                        WeBr1.Navigate TmpNa
                    Else
                        WeBr1.Navigate "about:" & MaHtm
                    End If
                ElseIf TmMes <> vbNullString Then
                    WeBr1.Navigate "about:" & TmMes
                Else
                    WeBr1.Navigate FiNam
                End If
            Else
                If TmMes <> vbNullString Then
                    TxMai.Text = TmMes
                End If
            End If
            DoEvents
            MaiMail.Clear
        Else
            WeBr1.InnerHTML = vbNullString
            SPopu "Dateifehler", Err.Description, IC48_Forbidden
        End If

        Select Case GlMKa 'Mailkatalog (1=Posteingang 2=Postaisgang)
        Case 1: DBCmEx2 "qryMailInGel", "@IdGel", "@IdxNr", -1, MaNum
        Case 2: DBCmEx2 "qryMailOutGel", "@IdGel", "@IdxNr", -1, MaNum
        End Select
        DoEvents
        
        If GlSuI.SuIdx < 1 Then
            SUpMa RowNr
        End If
        MaOpn = True
    Else
        CmAcs(TX_Mail_AttOpen).Enabled = False
        CmAcs(TX_Mail_AttSave).Enabled = False
        CmAcs(TX_Mail_AttExpo).Enabled = False
        CmSta.Pane(0).Text = Format$(MaiDa, "dd.mm.yyyy") & " - " & Format$(MaiTi, "hh:mm")
        CmSta.Pane(1).Text = "von : " & MaAbs
        TmMes = S_MaIdx(MaNum, "Kommentar")
        If TmMes <> vbNullString Then
            If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
                WeBr1.Navigate "about:" & TmMes
            Else
                TxMai.Text = TmMes
            End If
        End If
        DoEvents
         If GlSuI.SuIdx < 1 Then
            SUpMa RowNr
        End If
        MaOpn = True
    End If

Case 2: 'New Email

    CmSwi.Visible = False
    CmPgs.Visible = True
    Set RbTab = RbBar.FindTab(RibTab_Tex_Vorlag)
    CmAcs(Tex_Suchen).Checked = GlSMv
    If GlSMv = True Then 'Suchleiste Mailviewer
        CmBrs.Item(2).Visible = True
        CmAcs(TX_Mail_Suchen).Checked = True
    End If
    Rahm2.Visible = True
    RbTab.Selected = True
    
    For AktZa = 1 To UBound(GlZAd) 'Zugehörige Adressen Emailadressen
        CmEmp.AddItem GlZAd(AktZa, 2)
        CmBCC.AddItem GlZAd(AktZa, 2)
    Next AktZa
    DoEvents
    
    If GlMiA(GlSmI, 25) <> vbNullString Then 'Signaturen
        SigDa = GlVor & GlMiA(GlSmI, 25)
        If clFil.FilVor(SigDa) = True Then
            If Right$(LCase(SigDa), 3) = "txn" Then
                SigVo = True
            End If
        End If
        If SigVo = False Then
            If GlMiA(GlSmI, 11) <> vbNullString Then 'Emaisignatur
                SigNa = GlMiA(GlSmI, 11)
            Else
                SigNa = GlMiA(GlSmI, 1) 'Mitarbeitername
            End If
        End If
    ElseIf GlMiA(GlSmI, 11) <> vbNullString Then 'Emaisignatur
        If Len(GlMiA(GlSmI, 11)) > 8 Then
            SigNa = GlMiA(GlSmI, 11)
        Else
            SigNa = GlMiA(GlSmI, 1)
        End If
    Else
        SigNa = GlMiA(GlSmI, 1)
    End If
    
    If EmEmp <> vbNullString Then
        CmEmp.Text = EmEmp
    End If
    If EmBCC <> vbNullString Then
        CmBCC.Text = EmBCC
    End If
    If EmBet <> vbNullString Then
        TxBet.Text = EmBet
    Else
        TxBet.Text = "kein Betreff"
    End If

    If EmTex <> vbNullString Then
        SelTx = vbCrLf & EmTex & vbCrLf & vbCrLf
    Else
        SelTx = vbCrLf & vbCrLf
    End If

    If SigVo = True Then
        With TxCoN
            .ForeColor = vbBlack
            .LoadFromMemory SelTx, 1, True
            If Right$(LCase(SigDa), 3) = "txn" Then
                .Append SigDa, 0, 3
            End If
        End With
    Else
        With TxCoN
            .ForeColor = vbBlack
            .LoadFromMemory SelTx, 1, True
            SelTx = vbCrLf & SigNa & vbCrLf & vbCrLf
            .FontItalic = 1
            .ForeColor = 8404992
            .LoadFromMemory SelTx, 1, True
        End With
    End If

    If AttNa <> vbNullString Then
        Posi1 = InStr(1, AttNa, ";", 1) 'Mehrere Dateien vorhanden?
        If Posi1 > 0 Then
            AryIt = Split(AttNa, ";")
            AtGes = UBound(AryIt)
            ReDim Preserve GlAtt(AtGes)
            For AtAkt = 0 To AtGes - 1
                GlAtt(AtAkt + 1) = AryIt(AtAkt)
                With clFil
                    AttVo = clFil.FilVor(AryIt(AtAkt))
                    .FilPfa AryIt(AtAkt)
                    DaNam = .DaNam
                End With
                If AttVo = True Then
                    With CmAtt
                        .AddItem DaNam, AtAkt + 1
                        .ListIndex = AtAkt + 1
                    End With
                End If
            Next AtAkt
        Else
            ReDim Preserve GlAtt(1)
            GlAtt(1) = AttNa
            With clFil
                AttVo = clFil.FilVor(AttNa)
                .FilPfa AttNa
                DaNam = .DaNam
            End With
            If AttVo = True Then
                With CmAtt
                    .AddItem DaNam, 1
                    .ListIndex = 1
                End With
            End If
        End If
        GlAtV = True
        CmBrs.Item(3).Visible = True
    End If
    MaOpn = True
    
Case 3: 'Antwort

    CmSwi.Visible = False
    CmPgs.Visible = True
    
    For AktZa = 1 To UBound(GlZAd) 'zugehörige Adressen Emailadressen
        CmBCC.AddItem GlZAd(AktZa, 2)
    Next AktZa
    DoEvents

    If DaVor = True Then
        MaiMail.LoadFile FiNam, False
        DoEvents

        If Err.Number = 0 Then
            If MaiMail.IsEncrypted = True Then
                Set MaiMail = MaiMail.Decrypt(Nothing)
                If Err.Number <> 0 Then
                    SPopu "Emailverschlüsselung", Err.Description, IC48_Forbidden
                End If
            End If
            
            If MaiMail.IsSigned = True Then
                MaiMail.VerifySignature
                If Err.Number <> 0 Then
                    SPopu "Emailsignatur", Err.Description, IC48_Forbidden
                End If
            End If
        End If
        
        MaiMail.DecodeTNEF
    End If
        
    EmkGe = UBound(GlMkt)
    If EmkGe > 0 Then
        For EmkAk = 1 To EmkGe
            If CLng(GlMkt(EmkAk, 1)) = CLng(GlMiA(GlSmI, 2)) Then
                Exit For
            End If
        Next EmkAk
    End If
    
    CmSta.Pane(0).Text = Format$(MaiDa, "dd.mm.yyyy") & " - " & Format$(MaiTi, "hh:mm")
    CmSta.Pane(1).Text = "von : " & GlMkt(EmkAk, 13)
    CmSta.Pane(2).Text = "an : " & MaAbs

    TmpEm = MaAbs
    TmpBe = MaSub
    
    If DaVor = True Then
        TmMes = MaiMail.TextBody
        If TmMes = vbNullString Then
            TmMes = S_MaIdx(MaNum, "Kommentar")
        End If
    Else
        TmMes = S_MaIdx(MaNum, "Kommentar")
    End If

    If Len(TmMes) > 800 Then
        TmMes = Left$(TmMes, 800)
    End If

    If TmpEm <> vbNullString Then
        If PaNum > 0 Then
            ZugZa = S_ZuIdx(PaNum)

            For AktZa = 1 To ZugZa 'zugehötige Adressen
                If Trim$(LCase(TmpEm)) = LCase(GlZug(AktZa, 1)) Then
                    TmpBa = GlZug(AktZa, 2)
                    TmpVo = GlZug(AktZa, 3)
                    Exit For
                End If
            Next AktZa
            Erase GlZug
            
            DoEvents
            If TmpBa = vbNullString Then
                S_AdDe PaNum 'Adressendetails
                With GlADt
                    TmpBa = .AdRBr
                    TmpVo = .AdRVo
                End With
            End If
        End If
    End If

    If GlEmD = True Then 'automatischer Email Einleitungssatz
        Posi4 = InStr(1, TmpBa, TmpVo, vbTextCompare)
        If Posi4 > 0 Then
            TmpEi = "vielen Dank für Deine Email vom: " & Format$(MaiDa, "dd.mm.yyyy") & "."
        Else
            TmpEi = "vielen Dank für Ihre Email vom: " & Format$(MaiDa, "dd.mm.yyyy") & "."
        End If
    End If

    If GlMiA(GlSmI, 25) <> vbNullString Then 'Signaturen
        SigDa = GlVor & GlMiA(GlSmI, 25)
        If clFil.FilVor(SigDa) = True Then
            If Right$(LCase(SigDa), 3) = "txn" Then
                SigVo = True
            End If
        End If
        If SigVo = False Then
            If GlMiA(GlSmI, 11) <> vbNullString Then
                SigNa = GlMiA(GlSmI, 11)
            Else
                SigNa = GlMiA(GlSmI, 1)
            End If
        End If
    ElseIf GlMiA(GlSmI, 11) <> vbNullString Then
        SigNa = GlMiA(GlSmI, 11)
    Else
        SigNa = GlMiA(GlSmI, 1)
    End If

    If TmpEm <> vbNullString Then
        CmEmp.Text = TmpEm
    End If

    If TmpBe <> vbNullString Then
        If LCase(Left$(TmpBe, 2)) <> "re" Then
            TxBet.Text = "Re: " & TmpBe
        Else
            TxBet.Text = TmpBe
        End If
    Else
        TxBet.Text = "Ihre Email vom: " & Format$(MaiDa, "dd.mm.yyyy") & " - " & Format$(MaiTi, "hh:mm")
    End If

    If SigVo = True Then
        With TxCoN
            SelTx = vbCrLf & TmpBa & vbCrLf & vbCrLf & TmpEi & vbCrLf & vbCrLf & vbCrLf
            .ForeColor = vbBlack
            .LoadFromMemory SelTx, 1, True
            If Right$(LCase(SigDa), 3) = "txn" Then
                .Append SigDa, 0, 3
            End If
        End With
    Else
        With TxCoN
            SelTx = vbCrLf & TmpBa & vbCrLf & vbCrLf & TmpEi & vbCrLf & vbCrLf & vbCrLf
            .ForeColor = vbBlack
            .LoadFromMemory SelTx, 1, True
            SelTx = vbCrLf & SigNa & vbCrLf & vbCrLf
            .FontItalic = 1
            .ForeColor = 8404992
            .LoadFromMemory SelTx, 1, True
        End With
    End If

    With TxCoN
        .ForeColor = 9868950
        .FontItalic = 0
        .FontSize = 8
        TmMes = Replace(TmMes, vbCrLf & vbCrLf, vbCrLf, 1)
        .LoadFromMemory TmMes, 1, True
    End With
            
    Set RbTab = RbBar.FindTab(RibTab_Tex_Vorlag)
    CmAcs(Tex_Suchen).Checked = GlSMv
    If GlSMv = True Then
        CmBrs.Item(2).Visible = True
        CmAcs(TX_Mail_Suchen).Checked = True
    End If
    Rahm2.Visible = True
    RbTab.Selected = True

    DoEvents
    S_MaMa 7
    
    MaOpn = True
    
Case 4: 'Weiterleiten

    CmSwi.Visible = False
    CmPgs.Visible = True
    
    For AktZa = 1 To UBound(GlZAd) 'Zugehörige Adressen Emailadressen
        CmEmp.AddItem GlZAd(AktZa, 2)
        CmBCC.AddItem GlZAd(AktZa, 2)
    Next AktZa
    DoEvents
    
    If DaVor = True Then
        MaiMail.LoadFile FiNam, False
        DoEvents

        If Err.Number = 0 Then
            If MaiMail.IsEncrypted = True Then
                Set MaiMail = MaiMail.Decrypt(Nothing)
                If Err.Number <> 0 Then
                    SPopu "Emailverschlüsselung", Err.Description, IC48_Forbidden
                End If
            End If
            
            If MaiMail.IsSigned = True Then
                MaiMail.VerifySignature
                If Err.Number <> 0 Then
                    SPopu "Emailsignatur", Err.Description, IC48_Forbidden
                End If
            End If
            
            EmkGe = UBound(GlMkt)
            If EmkGe > 0 Then
                For EmkAk = 1 To EmkGe
                    If CLng(GlMkt(EmkAk, 1)) = CLng(GlMiA(GlSmI, 2)) Then
                        Exit For
                    End If
                Next EmkAk
            End If

            CmSta.Pane(0).Text = Format$(MaiDa, "dd.mm.yyyy") & " - " & Format$(MaiTi, "hh:mm")
            CmSta.Pane(1).Text = "von : " & GlMkt(EmkAk, 13)
            CmSta.Pane(2).Text = DaNam

            TmpBe = MaSub
            
            MaiMail.DecodeTNEF
            
            Select Case GlMTx
            Case 1: 'HTML-Text
                TmHtm = MaiMail.HtmlBody
                TmMes = MaiMail.TextBody
                If TmMes <> vbNullString Then
                    TmStr = Replace(TmMes, vbCrLf, "<br>", 1)
                    MaHtm = "<!doctype html> <html> <head> <meta charset='utf-8'> <title>" & MaSub & "</title> <meta name='generator' content='SimpliMed'> </head> <body> <span style='color:#000000;font-family:" & GlEFt.Name & ";font-size:" & Round(GlEFt.SIZE) + 3 & "px;'>" & TmStr & "</span></div> </body> </html>"
                End If
            Case 2: 'Ascii-Text
                TmMes = MaiMail.TextBody
                If TmMes = vbNullString Then
                    TmMes = S_MaIdx(MaNum, "Kommentar")
                End If
            Case 3: 'Daten-Text
                TmMes = S_MaIdx(MaNum, "Kommentar")
            End Select
            DoEvents
            
            For AktZa = 1 To UBound(GlZAd) 'Zugehörige Adressen Emailadressen
                CmEmp.AddItem GlZAd(AktZa, 2)
                CmBCC.AddItem GlZAd(AktZa, 2)
            Next AktZa
            DoEvents

            If clFil.FilDir(TmPfa) = False Then
                MkDir TmPfa
            End If

            If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
                If TmHtm <> vbNullString Then
                    MaiGeTo.WriteTextFile TmpNa, TmHtm, 65001
                Else
                    MaiGeTo.WriteTextFile TmpNa, MaHtm, 65001
                End If
            End If
            
            Set MaiAtCo = MaiMail.AttachmentList
            AtGes = MaiAtCo.Count
            
            ReDim Preserve GlAtt(AtGes)
            GlAtV = True

            If AtGes > 0 Then
                For AtAkt = 0 To AtGes - 1
                    Set MaiAtta = MaiAtCo.Item(AtAkt)
                    FilNa = TmPfa & "\" & MaiAtta.Name
                    MaiAtta.SaveAs FilNa, True
                    CmAtt.AddItem MaiAtta.Name
                    GlAtt(AtAkt + 1) = FilNa
                Next AtAkt
                CmAtt.ListIndex = AtGes
                CmBrs.Item(3).Visible = True
                DoEvents
            End If

            If TmpBe <> vbNullString Then
                TxBet.Text = "Fw: " & TmpBe
            Else
                TxBet.Text = "kein Betreff"
            End If
            
            TxCoN.SelText = vbCrLf & "Nachricht vom: " & MaiDa & " " & MaiTi & vbCrLf & vbCrLf
            If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
                If TmHtm <> vbNullString Then
                    If clFil.FilVor(TmpNa) = True Then
                        TxCoN.Load TmpNa, , 4, True
                    Else
                         TxCoN.LoadFromMemory TmHtm, 4, True
                    End If
                ElseIf MaHtm <> vbNullString Then
                    TxCoN.LoadFromMemory MaHtm, 4, True
                ElseIf TmMes <> vbNullString Then
                    TxCoN.LoadFromMemory TmMes, 1, True
                End If
            Else
                TxCoN.LoadFromMemory TmMes, 1, True
            End If

            Set RbTab = RbBar.FindTab(RibTab_Tex_Vorlag)
            CmAcs(Tex_Suchen).Checked = GlSMv
            If GlSMv = True Then
                CmBrs.Item(2).Visible = True
                CmAcs(TX_Mail_Suchen).Checked = True
            End If
            Rahm2.Visible = True
            RbTab.Selected = True
        End If
        DoEvents
        MaiMail.Clear
        
        S_MaMa 7
        MaOpn = True
    End If

Case 5: 'Resend Email

    CmSwi.Visible = False
    CmPgs.Visible = True
    
    For AktZa = 1 To UBound(GlZAd) 'Zugehörige Adressen Emailadressen
        CmBCC.AddItem GlZAd(AktZa, 2)
    Next AktZa
    DoEvents
    
    If DaVor = True Then
        MaiMail.LoadFile FiNam, False
        DoEvents

        If Err.Number = 0 Then
            If MaiMail.IsEncrypted = True Then
                Set MaiMail = MaiMail.Decrypt(Nothing)
                If Err.Number <> 0 Then
                    SPopu "Emailverschlüsselung", Err.Description, IC48_Forbidden
                End If
            End If
            
            If MaiMail.IsSigned = True Then
                MaiMail.VerifySignature
                If Err.Number <> 0 Then
                    SPopu "Emailsignatur", Err.Description, IC48_Forbidden
                End If
            End If

            CmSta.Pane(0).Text = Format$(MaiDa, "dd.mm.yyyy") & " - " & Format$(MaiTi, "hh:mm")
            CmSta.Pane(1).Text = "von : " & MaAbs
            CmSta.Pane(2).Text = DaNam

            TmpBe = MaSub
            TmpEm = MaiMail.ToAddr
            
            If Left$(TmpBe, 3) = "Fw:" Then
                TmpBe = Mid$(TmpBe, 4, Len(TmpBe) - 4)
            End If
            
            MaiMail.DecodeTNEF
            
            Select Case GlMTx
            Case 1: 'HTML-Text
                TmHtm = MaiMail.HtmlBody
                TmMes = MaiMail.TextBody
                If TmMes <> vbNullString Then
                    TmStr = Replace(TmMes, vbCrLf, "<br>", 1)
                    MaHtm = "<!doctype html> <html> <head> <meta charset='utf-8'> <title>" & MaSub & "</title> <meta name='generator' content='SimpliMed'> </head> <body> <span style='color:#000000;font-family:" & GlEFt.Name & ";font-size:" & Round(GlEFt.SIZE) + 3 & "px;'>" & TmStr & "</span></div> </body> </html>"
                End If
            Case 2: 'Ascii-Text
                TmMes = MaiMail.TextBody
                If TmMes = vbNullString Then
                    TmMes = S_MaIdx(MaNum, "Kommentar")
                End If
            Case 3: 'Daten-Text
                TmMes = S_MaIdx(MaNum, "Kommentar")
            End Select
            DoEvents
            
            If clFil.FilDir(TmPfa) = False Then
                MkDir TmPfa
            End If

            If TmHtm <> vbNullString Then
                MaiGeTo.WriteTextFile TmpNa, TmHtm, 65001
            Else
                MaiGeTo.WriteTextFile TmpNa, MaHtm, 65001
            End If
              
            Set MaiAtCo = MaiMail.AttachmentList
            AtGes = MaiAtCo.Count

            ReDim Preserve GlAtt(AtGes)
            GlAtV = True
            
            If AtGes > 0 Then
                For AtAkt = 0 To AtGes - 1
                    Set MaiAtta = MaiAtCo.Item(AtAkt)
                    FilNa = TmPfa & "\" & MaiAtta.Name
                    MaiAtta.SaveAs FilNa, True
                    CmAtt.AddItem MaiAtta.Name
                    GlAtt(AtAkt + 1) = FilNa
                Next AtAkt
                CmAtt.ListIndex = AtGes
                CmBrs.Item(3).Visible = True
                DoEvents
            End If
            
            Set RbTab = RbBar.FindTab(RibTab_Tex_Vorlag)
            CmAcs(Tex_Suchen).Checked = GlSMv
            If GlSMv = True Then
                CmBrs.Item(2).Visible = True
                CmAcs(TX_Mail_Suchen).Checked = True
            End If
            Rahm2.Visible = True
            RbTab.Selected = True

            If TmpEm <> vbNullString Then
                CmEmp.Text = TmpEm
            End If
            
            If TmpBe <> vbNullString Then
                TxBet.Text = TmpBe
            Else
                TxBet.Text = "kein Betreff"
            End If
            
            If TmMes <> vbNullString Then
                TxCoN.Load TmpNa, , 4, True
            ElseIf TmHtm <> vbNullString Then
                TxCoN.LoadFromMemory TmHtm, 4
            ElseIf TmMes <> vbNullString Then
                TxCoN.LoadFromMemory TmMes, 1
            End If
            DoEvents
            
            MaiMail.Clear
        End If
        
        DoEvents
        S_MaMa 7
        MaOpn = True
    End If

End Select

Set MaiAtCo = Nothing
Set MaiGeTo = Nothing
Set MaiMail = Nothing
Set MaiAdCo = Nothing

Set clFil = Nothing

Set CmBrs = Nothing
Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo0 = Nothing

Set clFil = Nothing

Set FM = Nothing

Exit Function

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaOpn " & Err.Number
Resume Next

End Function
Public Sub MaPos()
On Error GoTo PoErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBr2 As XtremeCommandBars.CommandBars

Set FM = frmMaiView
Set Rahm2 = FM.frmRahm2
Set CmBr2 = FM.comBar02
Set WeBr1 = FM.WebBrow1
Set TxCoN = FM.TexCont3
Set CmEmp = FM.cmbEmEmp
Set CmBCC = FM.cmbEmBCC
Set TxBet = FM.txtEmBet
Set TxCCM = FM.txtEmCCM
Set TxMai = FM.txtMaiTx

If FM.WindowState <> vbMinimized Then
    CmBr2.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    If GlNaT = 1 Then 'Mailtyp (1=View 2=Neu 3=Antwort)
        If GlMTx = 1 Then 'Mailtextanzeige (1=HTML 2=ASCII)
            TxCoN.Move ClLin, ClHoh + 4000, 300, 300
            TxMai.Move ClLin, ClHoh + 4000, 300, 300
            WeBr1.Move ClLin + 100, ClObn, ClBre - 100, ClHoh
        Else
            TxCoN.Move ClLin, ClHoh + 4000, 300, 300
            TxMai.Move ClLin + 100, ClObn + 510, ClBre - 100, ClHoh - 520
            WeBr1.Move ClLin, ClHoh + 4000, 300, 300
        End If
    Else
        WeBr1.Move ClLin, ClHoh + 4000, 300, 300
        TxCoN.Move ClLin + 100, ClObn + 1900, ClBre - 100, ClHoh - 1900
        TxMai.Move ClLin, ClHoh + 4000, 300, 300
        Rahm2.Move ClLin, ClObn, ClBre, 1900
        CmEmp.Width = ClBre - 1140
        TxCCM.Width = ClBre - 1140
        CmBCC.Width = ClBre - 1140
        TxBet.Width = ClBre - 1140
    End If
End If

Set CmBr2 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaPos " & Err.Number
Resume Next

End Sub
Private Sub MaReg()
On Error GoTo ReErr
'Legt benötigte Einträge in der Registry an

Dim xPos As Long
Dim yPos As Long
Dim xGro As Long
Dim yGro As Long

If IniGetSek(GlINI, "MailView") = False Then
    xGro = 950
    yGro = GlyGr - (GlyGr / 7)
    
    xPos = (GlxGr / 2) - (xGro / 2)
    yPos = (GlyGr / 2) - (yGro / 2)
    
    IniSetSek "MailView"
    IniSetVal "MailView", "FenLin", xPos
    IniSetVal "MailView", "FenObe", yPos
    IniSetVal "MailView", "FenBre", xGro
    IniSetVal "MailView", "FenHoh", yGro
End If

Exit Sub

ReErr:
If GlDbg = True Then SErLog Err.Description & " MaReg " & Err.Number
Resume Next

End Sub
Public Sub MaSav(ByVal SaTyp As Integer)
On Error GoTo PoErr
'üffnet den Dateianhang

Dim PatNr As Long
Dim FiNam As String
Dim DaNam As String
Dim DaExt As String
Dim NeNam As String
Dim TmpPf As String
Dim TmpDa As String
Dim TmpNa As String
Dim AtGes As Integer
Dim AtAkt As Integer
Dim AusZa As Integer
Dim RowNr As Integer
Dim RetWe As Boolean
Dim AtSav As Boolean
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmAtt As XtremeCommandBars.CommandBarComboBox
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set MaiGeTo = New EAGetMailObjLib.Tools
Set MaiMail = New EAGetMailObjLib.Mail

If SaTyp < 5 Then
    Set FM = frmMaiView
    Set CmBrs = FM.comBar02
    Set CmAtt = CmBrs.FindControl(CmAtt, SY_SuCm3, , True)
End If

Set CoDia = frmMain.comDialo
Set RpCo0 = frmMain.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

MaiMail.LicenseCode = "EG-C1653719494-01199-8BVCAA171B372F1D-F797V24BDU55166A"
   
Set clFil = New clsFile
clFil.hwnd = frmMain.hwnd

Screen.MousePointer = vbHourglass
DoEvents

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        RowNr = RpRow.Index
        If MaAry(Mai_MailFile, RowNr) <> vbNullString Then
            If GlMaY = True Then 'Emailflyoutfenster Mailindex
                DaNam = S_MaIdx(GlMaY, "MailFile")
            Else
                DaNam = MaAry(Mai_MailFile, RowNr)
            End If
            FiNam = GlDpf & "Emails\" & DaNam
            TmpPf = GlDpf & "Emails\Temp\" & GlMiA(GlSmI, 20) & "\"

            If clFil.FilVor(FiNam) = True Then
                MaiMail.LoadFile FiNam, True

                If Err.Number = 0 Then
                    Set MaiAtCo = MaiMail.AttachmentList
                    AtGes = MaiAtCo.Count

                    If AtGes > 0 Then
                        Select Case SaTyp
                        Case 1: 'öfnen aus Emailfenster
                        
                            If clFil.FilVor(TmpPf) = False Then
                                MkDir TmpPf
                            End If

                            For AtAkt = 0 To AtGes - 1
                                Set MaiAtta = MaiAtCo.Item(AtAkt)

                                If LCase(MaiAtta.Name) = LCase(CmAtt.Text) Then
                                    TmpDa = CmAtt.Text
                                    TmpNa = TmpPf & TmpDa

                                    With clFil
                                        .FilPfa TmpNa
                                        DaExt = .DaExt
                                    End With

                                    AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                                    If AtSav = False Then
                                        MaiAtta.SaveAs TmpNa, True
                                        DoEvents
                                        Select Case LCase(DaExt)
                                        Case "pdf": SImage TmpNa
                                        Case "jpg": SImage TmpNa
                                        Case "png": SImage TmpNa
                                        Case "bmp": SImage TmpNa
                                        Case "tif": SImage TmpNa
                                        Case "gif": SImage TmpNa
                                        Case "wmf": SImage TmpNa
                                        Case "emf": SImage TmpNa
                                        Case "jpeg": SImage TmpNa
                                        Case "tiff": SImage TmpNa
                                        Case "doc": VoTxMa TmpNa, TmpDa, 9
                                        Case "dot": VoTxMa TmpNa, TmpDa, 9
                                        Case "rtf": VoTxMa TmpNa, TmpDa, 5
                                        Case "txt": VoTxMa TmpNa, TmpDa, 1
                                        Case "csv": VoTxMa TmpNa, TmpDa, 1
                                        Case "docx": VoTxMa TmpNa, TmpDa, 13
                                        Case Else: SPopu "Ungültiger Dateityp", "Dieser Dateityp darf nicht geöffnet werden", IC48_Warning
                                        End Select
                                    Else
                                        SPopu "Sicherheitshinweis", "Das ?fnen von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                                    End If
                                    Exit For
                                Else
                                    SPopu "Falscher Dateiname", "Der Dateiname aus dem Anhang stimmt nicht mit dem dargestellten Dateinamen überein!", IC48_Forbidden
                                End If
                            Next AtAkt
                            
                        Case 2: 'Speichern mit Speicherdialog
                        
                            For AtAkt = 0 To AtGes - 1
                                Set MaiAtta = MaiAtCo.Item(AtAkt)

                                If LCase(MaiAtta.Name) = LCase(CmAtt.Text) Then
                                    TmpNa = GlTmp & CmAtt.Text

                                    With clFil
                                        If .FilVor(TmpNa) = True Then
                                            .DaLoe = TmpNa & vbNullChar
                                            .FilLoe
                                        End If
                                        .FilPfa TmpNa
                                        DaExt = .DaExt
                                    End With

                                    AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                                    If AtSav = False Then
                                        With CoDia
                                            .CancelError = True
                                            .DialogStyle = 1
                                            .DefaultExt = "*." & DaExt
                                            .Filter = "Alle Dateien (*.*)|*.*"
                                            .DialogTitle = "Bitte Name und Ordner der Datei angeben"
                                            .FileName = GlEPf & CmAtt.Text
                                            .InitDir = GlEPf
                                            .ShowSave
                                            NeNam = .FileName
                                            If .FileTitle = vbNullString Then
                                                Set clFil = Nothing
                                                Set CoDia = Nothing
                                                Set CmBrs = Nothing
                                                Set RpCls = Nothing
                                                Set RpSel = Nothing
                                                Set RpCo0 = Nothing
                                                Exit Sub
                                            End If
                                        End With
                                        
                                        If NeNam <> vbNullString Then
                                            With clFil
                                                If .FilVor(TmpNa) = True Then
                                                    .DaLoe = TmpNa & vbNullChar
                                                    .FilLoe
                                                End If
                                            End With
                                            MaiAtta.SaveAs NeNam, True
                                        End If
                                        Exit For
                                    Else
                                        SPopu "Sicherheitshinweis", "Das Speichern von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                                    End If
                                Else
                                    SPopu "Falscher Dateiname", "Der Dateiname aus dem Anhang stimmt nicht mit dem dargestellten Dateinamen überein!", IC48_Forbidden
                                End If
                            Next AtAkt
                            
                        Case 3: 'Exportieren aus Emailfenster

                            With CoDia
                                .CancelError = True
                                .DialogStyle = 1
                                .DialogTitle = "Geben Sie bitte den gewünschten Ordner an"
                                .FileName = GlEPf
                                RetWe = .ShowBrowseFolder
                                NeNam = .FileName
                                If RetWe = 0 Then
                                    Set clFil = Nothing
                                    Set CoDia = Nothing
                                    Set CmBrs = Nothing
                                    Set RpCls = Nothing
                                    Set RpSel = Nothing
                                    Set RpCo0 = Nothing
                                    Exit Sub
                                End If
                            End With
                            If NeNam <> vbNullString Then
                                NeNam = NeNam & "\"
                                For AtAkt = 0 To AtGes - 1
                                    Set MaiAtta = MaiAtCo.Item(AtAkt)

                                    TmpNa = NeNam & MaiAtta.Name

                                    With clFil
                                        .FilPfa TmpNa
                                        DaExt = .DaExt
                                    End With

                                    AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                                    
                                    With clFil
                                        If .FilVor(TmpNa) = True Then
                                            .DaLoe = TmpNa & vbNullChar
                                            .FilLoe
                                        End If
                                    End With
                                    
                                    If AtSav = False Then
                                        MaiAtta.SaveAs TmpNa, True
                                    Else
                                        SPopu "Sicherheitshinweis", "Das Exportieren von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                                    End If
                                Next AtAkt
                            End If
                            
                        Case 4: 'Rechnungsimport aus Emailfenster
                        
                            For AtAkt = 0 To AtGes - 1
                                Set MaiAtta = MaiAtCo.Item(AtAkt)

                                TmpNa = GlIPf & MaiAtta.Name

                                With clFil
                                    .FilPfa TmpNa
                                    DaExt = .DaExt
                                End With

                                AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                                
                                With clFil
                                    If .FilVor(TmpNa) = True Then
                                        .DaLoe = TmpNa & vbNullChar
                                        .FilLoe
                                    End If
                                End With
                                
                                If AtSav = False Then
                                    MaiAtta.SaveAs TmpNa, True
                                Else
                                    SPopu "Sicherheitshinweis", "Das ?fnen von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                                End If
                            Next AtAkt
                            If LCase(DaExt) = "zip" Then
                                frmImport.Show vbModal
                            ElseIf LCase(DaExt) = "smp" Then
                                frmImport.Show vbModal
                            ElseIf LCase(DaExt) = "xml" Then
                                frmImport.Show vbModal
                            End If
                            
                        Case 5: 'öfnen aus der Übersicht
                            
                            If clFil.FilVor(TmpPf) = False Then
                                MkDir TmpPf
                            End If
                            
                            Set MaiAtta = MaiAtCo.Item(AtGes - 1)
                                                    
                            TmpDa = MaiAtta.Name
                            TmpNa = TmpPf & TmpDa

                            With clFil
                                .FilPfa TmpNa
                                DaExt = .DaExt
                            End With
                            
                            AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                            If AtSav = False Then
                                MaiAtta.SaveAs TmpNa, True
                                DoEvents
                                Select Case LCase(DaExt)
                                Case "pdf": SImage TmpNa
                                Case "jpg": SImage TmpNa
                                Case "png": SImage TmpNa
                                Case "bmp": SImage TmpNa
                                Case "tif": SImage TmpNa
                                Case "gif": SImage TmpNa
                                Case "wmf": SImage TmpNa
                                Case "emf": SImage TmpNa
                                Case "jpeg": SImage TmpNa
                                Case "tiff": SImage TmpNa
                                Case "doc": VoTxMa TmpNa, TmpDa, 9
                                Case "dot": VoTxMa TmpNa, TmpDa, 9
                                Case "rtf": VoTxMa TmpNa, TmpDa, 5
                                Case "txt": VoTxMa TmpNa, TmpDa, 1
                                Case "csv": VoTxMa TmpNa, TmpDa, 1
                                Case "docx": VoTxMa TmpNa, TmpDa, 13
                                Case Else: SPopu "Ungültiger Dateityp", "Dieser Dateityp darf nicht geöffnet werden", IC48_Warning
                                End Select
                            Else
                                SPopu "Sicherheitshinweis", "Das ?fnen von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                            End If

                        Case 6: 'Speichern aus der üersiicht

                            Set MaiAtta = MaiAtCo.Item(0)

                            TmpDa = MaiAtta.Name
                            TmpNa = GlTmp & TmpDa

                            With clFil
                                If .FilVor(TmpNa) = True Then
                                    .DaLoe = TmpNa & vbNullChar
                                    .FilLoe
                                End If
                                .FilPfa TmpNa
                                DaExt = .DaExt
                            End With

                            AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                            If AtSav = False Then
                                With CoDia
                                    .CancelError = True
                                    .DialogStyle = 1
                                    .DefaultExt = "*." & DaExt
                                    .Filter = "Alle Dateien (*.*)|*.*"
                                    .DialogTitle = "Bitte Name und Ordner der Datei angeben"
                                    .FileName = GlEPf & TmpDa
                                    .InitDir = GlEPf
                                    .ShowSave
                                    NeNam = .FileName
                                    If .FileTitle = vbNullString Then
                                        Set clFil = Nothing
                                        Set CoDia = Nothing
                                        Set CmBrs = Nothing
                                        Set RpCls = Nothing
                                        Set RpSel = Nothing
                                        Set RpCo0 = Nothing
                                        Exit Sub
                                    End If
                                End With
                                If NeNam <> vbNullString Then
                                     With clFil
                                        If .FilVor(TmpNa) = True Then
                                            .DaLoe = TmpNa & vbNullChar
                                            .FilLoe
                                        End If
                                    End With
                                    MaiAtta.SaveAs NeNam, True
                                End If
                            Else
                                SPopu "Sicherheitshinweis", "Das ?fnen von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                            End If
                            
                        Case 7: 'Exportieren auis der ?ersicht
                        
                            With CoDia
                                .CancelError = True
                                .DialogStyle = 1
                                .DialogTitle = "Geben Sie bitte den gewünschten Ordner an"
                                .FileName = GlEPf
                                RetWe = .ShowBrowseFolder
                                NeNam = .FileName
                                If RetWe = 0 Then
                                    Set clFil = Nothing
                                    Set CoDia = Nothing
                                    Set CmBrs = Nothing
                                    Set RpCls = Nothing
                                    Set RpSel = Nothing
                                    Set RpCo0 = Nothing
                                    Exit Sub
                                End If
                            End With
                            If NeNam <> vbNullString Then
                                NeNam = NeNam & "\"

                                For AtAkt = 0 To AtGes - 1
                                    Set MaiAtta = MaiAtCo.Item(AtAkt)

                                    TmpNa = NeNam & MaiAtta.Name

                                    With clFil
                                        .FilPfa TmpNa
                                        DaExt = .DaExt
                                        If .FilVor(TmpNa) = True Then
                                            .DaLoe = TmpNa & vbNullChar
                                            .FilLoe
                                        End If
                                    End With

                                    AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                                    If AtSav = False Then
                                        MaiAtta.SaveAs TmpNa, True
                                    Else
                                        SPopu "Sicherheitshinweis", "Das ?fnen von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                                    End If
                                Next AtAkt
                            End If
                            
                        Case 8: 'Rechnungsimport aus üersicht

                            For AtAkt = 0 To AtGes - 1
                                Set MaiAtta = MaiAtCo.Item(AtAkt)
                                TmpNa = MaiAtta.Name
                                                                
                                With clFil
                                    .FilPfa TmpNa
                                    DaExt = .DaExt
                                End With

                                AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien
                                
                                With clFil
                                    If .FilVor(TmpNa) = True Then
                                        .DaLoe = TmpNa & vbNullChar
                                        .FilLoe
                                    End If
                                End With
                                
                                If AtSav = False Then
                                    MaiAtta.SaveAs TmpNa, True
                                Else
                                    SPopu "Sicherheitshinweis", "Das ?fnen von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                                End If
                            Next AtAkt

                            If LCase(DaExt) = "zip" Then
                                frmImport.Show vbModal
                            ElseIf LCase(DaExt) = "smp" Then
                                frmImport.Show vbModal
                            ElseIf LCase(DaExt) = "xml" Then
                                frmImport.Show vbModal
                            End If
                            
                        Case 9: 'Anhang Importieren

                            If MaAry(Mai_ID0, RpRow.Index) <> vbNullString Then
                                PatNr = MaAry(Mai_ID0, RpRow.Index)
                                If PatNr > 0 Then
                                    
                                    If clFil.FilVor(TmpPf) = False Then
                                        MkDir TmpPf
                                    End If

                                    For AtAkt = 0 To AtGes - 1
                                        Set MaiAtta = MaiAtCo.Item(AtAkt)

                                        TmpDa = MaiAtta.Name
                                        TmpNa = TmpPf & TmpDa
                                        
                                        With clFil
                                            .FilPfa TmpNa
                                            DaExt = .DaExt
                                        End With
                                        
                                        AtSav = MaChE(DaExt) 'Prüfung verbotener Dateien

                                        If AtSav = False Then
                                            MaiAtta.SaveAs TmpNa, True
                                            DoEvents

                                            Select Case LCase(DaExt)
                                            Case "jpg": SFilIm TmpNa, PatNr
                                            Case "peg": SFilIm TmpNa, PatNr
                                            Case "png": SFilIm TmpNa, PatNr
                                            Case "bmp": SFilIm TmpNa, PatNr
                                            Case "tif": SFilIm TmpNa, PatNr
                                            Case "gif": SFilIm TmpNa, PatNr
                                            Case "wmf": SFilIm TmpNa, PatNr
                                            Case "emf": SFilIm TmpNa, PatNr
                                            Case "pdf": SFilIm TmpNa, PatNr
                                            Case "doc": SFilIm TmpNa, PatNr
                                            Case "dot": SFilIm TmpNa, PatNr
                                            Case "rtf": SFilIm TmpNa, PatNr
                                            Case "txt": SFilIm TmpNa, PatNr
                                            Case "csv": SFilIm TmpNa, PatNr
                                            Case "ocx": SFilIm TmpNa, PatNr
                                            Case Else: SPopu "Ungültiger Dateityp", "Dieser Dateityp darf nicht zugeordnet werden", IC48_Warning
                                            End Select
                                        Else
                                            SPopu "Sicherheitshinweis", "Das Zuordnen von *." & DaExt & " Dateien ist nicht zul?ig!", IC48_Forbidden
                                        End If
                                    Next AtAkt
                                Else
                                    SPopu "Keine Adresse zugeordnet", "Diese Email muss erst einem Patienten zugeordnet werden", IC48_Forbidden
                                End If
                            Else
                                SPopu "Keine Adresse zugeordnet", "Diese Email muss erst einem Patienten zugeordnet werden", IC48_Forbidden
                            End If
                        End Select
                        
                        If CBool(MaAry(Mai_Gelesen, RpRow.Index)) = False Then
                            If SaTyp > 3 Then
                                S_MaMa 7
                            End If
                        End If
                    Else
                        SPopu "Kein Dateianhang", "Es ist kein Dateianhang vorhanden", IC48_Forbidden
                    End If
                Else
                    SPopu "Dateianhangfehler", "Beim ?fnen der Mail ist ein Fehler aufgetreten", IC48_Forbidden
                End If
                DoEvents
                MaiMail.Clear
            Else
                SPopu "Keine Emaildatei", "Die Emaildatei wurde nicht gefunden", IC48_Forbidden
            End If
        End If
    End If
End If

Set MaiAtCo = Nothing
Set MaiGeTo = Nothing
Set MaiMail = Nothing

Set clFil = Nothing

DoEvents
Screen.MousePointer = vbNormal

Set CoDia = Nothing
Set CmBrs = Nothing
Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo0 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaSav " & Err.Number
Resume Next

End Sub
Private Function MaTag(ByVal TmStr As String) As String
On Error Resume Next

TmStr = Replace(TmStr, ">", "&gt;")
TmStr = Replace(TmStr, "<", "&lt;")

MaTag = TmStr

End Function

Public Sub MaTex()
On Error GoTo PoErr

Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmMaT As XtremeCommandBars.CommandBarComboBox

Set FM = frmMain
Set CmBrs = FM.comBar01

Set CmMaT = CmBrs.FindControl(CmMaT, KA_Mail_TexCombo, , True)

GlMTx = CmMaT.ListIndex 'Mailtextanzeige (1=HTML 2=ASCII)

IniSetVal "Layout", "MaiTex", GlMTx

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaTex " & Err.Number
Resume Next

End Sub
Public Sub MaThr()
On Error GoTo PoErr
'Mail Thread

Dim PatNr As Long
Dim RpCo0 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo0 = FM.repCont0
Set RpCls = RpCo0.Columns
Set RpSel = RpCo0.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If MaAry(Mai_ID0, RpRow.Index) <> vbNullString Then
            PatNr = MaAry(Mai_ID0, RpRow.Index)
            
            SSuAu 'Hebt die markierten Suchbuchstaben wieder auf
            DoEvents

            With GlSuI
                .SuIdx = 7
                .SuPat = PatNr
            End With
            
            DoEvents
            SSuch
            DoEvents
            
            If GlEmV > 0 Then
                RpCo0.SetFocus
            End If
            
        Else
            SPopu "Keine Adreszuordnung", "Dieser Email wurde noch keine Adresse zugeordnet", IC48_Forbidden
        End If
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo0 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaThr " & Err.Number
Resume Next

End Sub
Public Sub MaViw()
On Error GoTo PoErr
'Öffnet den E-Mail-Anhang

Dim TmpDa As String
Dim TmpNa As String
Dim TmpPf As String
Dim DaExt As String
Dim AktZa As Integer
Dim GesZa As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmAtt As XtremeCommandBars.CommandBarComboBox

Set FM = frmMaiView
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar
Set RbBar = CmBrs.Item(1)
Set CmAtt = CmBrs.FindControl(CmAtt, SY_SuCm3, , True)

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

GesZa = CmAtt.ListCount

If GesZa > 0 Then
    If GlAtV = True Then
        For AktZa = 1 To UBound(GlAtt)
            TmpNa = GlAtt(AktZa)
            With clFil
                If .FilVor(TmpNa) = True Then
                    .FilPfa TmpNa
                    TmpDa = .DaNam
                    TmpPf = .DaPfa
                    DaExt = .DaExt
                    If LCase(CmAtt.Text) = LCase(TmpDa) Then
                        CmSta.Pane(1).Text = TmpDa
                        Select Case LCase(DaExt)
                        Case "pdf": SImage TmpNa
                        Case "jpg": SImage TmpNa
                        Case "png": SImage TmpNa
                        Case "bmp": SImage TmpNa
                        Case "tif": SImage TmpNa
                        Case "gif": SImage TmpNa
                        Case "wmf": SImage TmpNa
                        Case "emf": SImage TmpNa
                        Case "jpeg": SImage TmpNa
                        Case "tiff": SImage TmpNa
                        Case "doc": VoTxMa TmpNa, TmpDa, 9
                        Case "dot": VoTxMa TmpNa, TmpDa, 9
                        Case "rtf": VoTxMa TmpNa, TmpDa, 5
                        Case "txt": VoTxMa TmpNa, TmpDa, 1
                        Case "csv": VoTxMa TmpNa, TmpDa, 1
                        Case "docx": VoTxMa TmpNa, TmpDa, 13
                        Case Else: SPopu "Ungültiger Dateityp", "Dieser Dateityp darf nicht geöffnet werden", IC48_Warning
                        End Select
                        Exit For
                    End If
                End If
            End With
        Next AktZa
    End If
End If
                        
Set CmBrs = Nothing
Set clFil = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MaViw " & Err.Number
Resume Next

End Sub
Public Sub MaVor()
On Error GoTo InErr
'Öffnet eine Newslettervorlage

Dim FiNam As String
Dim SuFix As String
Dim Posit As Integer

Set FM = frmMaiView
Set TxCoN = FM.TexCont3
Set CoDia = frmMain.comDialo

With CoDia
    .CancelError = True
    .DialogStyle = 1
    .DialogTitle = "Bitte Name und Ordner der Datei angeben"
    .DefaultExt = "*.txn"
    .Filter = "Newslettervorlage (.txn)|*.txn|Microsoft Word 2002/2003 (.doc)|*.doc|Rich Text Format (.rtf)|*.rtf|ANSI-Textdatei (.txt)|*.txt|Hypertext Markup Language (.htm)|*.htm|Alle Dateien (*.*)|*.*"
    .InitDir = GlVor
    .FileName = vbNullString
    .ShowOpen
    FiNam = .FileName
    If .FileTitle = vbNullString Then
        Set CoDia = Nothing
        Set clFil = Nothing
        Exit Sub
    End If
End With
If Not IsNull(FiNam) And Not FiNam = vbNullString Then
    Posit = InStrRev(FiNam, ".", Len(FiNam), 1)
    If Posit > 0 Then
        SuFix = Mid$(FiNam, Posit + 1, Len(FiNam) - Posit)
    Else
        SuFix = vbNullString
    End If

    Select Case LCase(SuFix)
    Case "txm":
        With TxCoN
            .ResetContents
            .Load FiNam, , 3
        End With
    Case "txn":
        With TxCoN
            .ResetContents
            .Load FiNam, , 3
        End With
    Case "txr":
        With TxCoN
            .ResetContents
            .Load FiNam, , 3
        End With
    Case "htm":
        With TxCoN
            .ResetContents
            .Load FiNam, , 4
        End With
    Case "css":
        With TxCoN
            .ResetContents
            .Load FiNam, , 11
        End With
    Case "doc":
        With TxCoN
            .ResetContents
            .Load FiNam, , 9
        End With
    Case "docx":
        With TxCoN
            .ResetContents
            .Load FiNam, , 13
        End With
    Case Else:
        With TxCoN
            .ResetContents
           .Text = vbNullString
        End With
    End Select
End If

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " MaVor " & Err.Number
Resume Next

End Sub

Public Sub MTxFo()
On Error GoTo PoErr

Dim SeiZa As Variant
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmCo1 As XtremeCommandBars.CommandBarComboBox
Dim CmCo2 As XtremeCommandBars.CommandBarComboBox
Dim CmCo3 As XtremeCommandBars.CommandBarComboBox
Dim CmPa1 As XtremeCommandBars.StatusBarPane

Set FM = frmMaiView
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont3
Set CmAcs = CmBrs.Actions
Set CmSta = CmBrs.StatusBar

Set CmCo1 = CmBrs.FindControl(CmCo1, Tex_FntAu4, , True)
Set CmCo2 = CmBrs.FindControl(CmCo2, Tex_FntGr4, , True)
Set CmCo3 = CmBrs.FindControl(CmCo3, Tex_DaFeAd, , True)

SeiZa = TxCoN.CurrentInputPosition

If TxCoN.Alignment = 0 Then
    CmAcs(Tex_AusrRe).Checked = False
    CmAcs(Tex_AusrZe).Checked = False
    CmAcs(Tex_AusrBl).Checked = False
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrLi).Checked = True
ElseIf TxCoN.Alignment = 1 Then
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrZe).Checked = False
    CmAcs(Tex_AusrBl).Checked = False
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrRe).Checked = True
ElseIf TxCoN.Alignment = 2 Then
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrRe).Checked = False
    CmAcs(Tex_AusrBl).Checked = False
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrZe).Checked = True
ElseIf TxCoN.Alignment = 3 Then
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrRe).Checked = False
    CmAcs(Tex_AusrZe).Checked = False
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrBl).Checked = True
Else
    CmAcs(Tex_AusrRe).Checked = False
    CmAcs(Tex_AusrZe).Checked = False
    CmAcs(Tex_AusrBl).Checked = False
    CmAcs(Tex_AusrLi).Checked = False
    CmAcs(Tex_AusrLi).Checked = True
End If

If TxCoN.FontBold = 0 Then
    CmAcs(Tex_ForFet).Checked = False
Else
    CmAcs(Tex_ForFet).Checked = True
End If

If TxCoN.FontItalic = 0 Then
    CmAcs(Tex_ForKur).Checked = False
Else
    CmAcs(Tex_ForKur).Checked = True
End If

If TxCoN.FontUnderline = 0 Then
    CmAcs(Tex_ForUnt).Checked = False
Else
    CmAcs(Tex_ForUnt).Checked = True
End If

If TxCoN.FontStrikethru = 0 Then
    CmAcs(Tex_ForDur).Checked = False
Else
    CmAcs(Tex_ForDur).Checked = True
End If

If TxCoN.BulletAttribute(txBulletLevel) <> vbNullString Then
    CmAcs(Tex_Aufzah).Checked = True
Else
    CmAcs(Tex_Aufzah).Checked = False
End If

If TxCoN.NumberingAttribute(txNumberingLevel) <> vbNullString Then
    CmAcs(Tex_Numeri).Checked = True
Else
    CmAcs(Tex_Numeri).Checked = False
End If

If TxCoN.BaseLine = 100 Then
    CmAcs(Tex_FntHoh).Checked = True
    CmAcs(Tex_FntTif).Checked = False
ElseIf TxCoN.BaseLine = -100 Then
    CmAcs(Tex_FntHoh).Checked = False
    CmAcs(Tex_FntTif).Checked = True
Else
    CmAcs(Tex_FntHoh).Checked = False
    CmAcs(Tex_FntTif).Checked = False
End If

If TxCoN.IndentL > 0 Then
    CmAcs(Tex_EinzLi).Checked = True
    CmAcs(Tex_EinzRe).Checked = False
ElseIf TxCoN.IndentR > 0 Then
    CmAcs(Tex_EinzLi).Checked = False
    CmAcs(Tex_EinzRe).Checked = True
Else
    CmAcs(Tex_EinzLi).Checked = False
    CmAcs(Tex_EinzRe).Checked = False
End If

If TxCoN.LineSpacing <> 100 Then
    CmAcs(Tex_Abstan).Checked = True
Else
    CmAcs(Tex_Abstan).Checked = False
End If

CmCo1.Text = TxCoN.FontName
CmCo2.Text = TxCoN.FontSize

Set CmSta = Nothing
Set CmAcs = Nothing
Set CmBrs = Nothing

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " MTxFo " & Err.Number
Resume Next

End Sub
Private Sub PrtInit(ByVal PrCon As Integer)
On Error GoTo InErr

Dim PrtHo As Boolean
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo2 As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCoK As XtremeReportControl.ReportControl
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCo9 As XtremeReportControl.ReportControl

Set FM = frmPrintPrev
Set PrtPr = FM.prtPrev1
Set ChCon = frmMain.chrCont1
Set CaCol = frmMain.calCont1
Set RpCo1 = frmMain.repCont1
Set RpCo2 = frmMain.repCont2
Set RpCo3 = frmMain.repCont3
Set RpCo4 = frmMain.repCont4
Set RpCo5 = frmMain.repCont5
Set RpCo6 = frmMain.repCont6
Set RpCoK = frmMain.repContK
Set RpCo8 = frmMain.repCont8
Set RpCo9 = frmMain.repCont9

With PrtPr
    Select Case GlBut:
    Case RibTab_Adressen:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Mandanten:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Verordner:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Mitarbeit:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Fragebogen:
            .Orientation = xtpOrientationPortrait
    Case RibTab_Krankenbla:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Abrechnung:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Tagesproto:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Vorbereit:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Rezeptmodul:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Belegmodul:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Bildmodul:
            .Orientation = xtpOrientationPortrait
    Case RibTab_Rechnungen:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Mahnwesen:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Buchungen:
            .Orientation = xtpOrientationLandscape
    Case RibTab_HomeBanki:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Statistik:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Ter_Kalend:
            PrtHo = CBool(IniGetVal("TerSys", "TiPrHo"))
            If GlCal = 1 Then
                .Orientation = xtpOrientationPortrait
            Else
                If PrtHo = True Then
                    .Orientation = xtpOrientationLandscape
                Else
                    .Orientation = xtpOrientationPortrait
                End If
            End If
    Case RibTab_Ter_Listen:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Ter_Akont:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Ter_Warte:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Ter_Mitarb:
            .Orientation = xtpOrientationLandscape
    Case RibTab_Ter_Raeume:
            .Orientation = xtpOrientationLandscape
    Case RibTab_LabBericht:
            .Orientation = xtpOrientationPortrait
    Case RibTab_LabAuftrag:
            .Orientation = xtpOrientationPortrait
    Case RibTab_LabBerichte:
            .Orientation = xtpOrientationPortrait
    Case RibTab_LabAuftrage:
            .Orientation = xtpOrientationPortrait
    Case RibTab_Kat_Eintrg:
            .Orientation = xtpOrientationPortrait
    Case RibTab_Kat_Ketten:
            .Orientation = xtpOrientationPortrait
    Case RibTab_Tex_Dokumt:
    Case RibTab_Tex_Vorlag:
    Case RibTab_Tex_Rezept:
    Case RibTab_Tex_NewsLe:
    Case RibTab_Kat_Explor:
    Case RibTab_Kat_Frage:
            .Orientation = xtpOrientationPortrait
    Case RibTab_Tex_Email:
    Case Else:
            .Orientation = xtpOrientationPortrait
    End Select
    
    Select Case GlSty
    Case 7: .VisualTheme = xtpThemeOffice2013
    Case 8: .VisualTheme = xtpThemeOffice2013
    Case Else: .VisualTheme = xtpThemeRibbon
    End Select
    .Title = "Vorschau"
    .ZoomState = 100
    Select Case PrCon
    Case 1:
        .PrintView = ChCon.CreatePrintView()
    Case 2:
        .PrintView = CaCol.CreatePrintView()
    Case 3:
        Select Case GlBut:
        Case RibTab_Adressen: .PrintView = RpCo2.CreatePrintView()
        Case RibTab_Mandanten: .PrintView = RpCo2.CreatePrintView()
        Case RibTab_Verordner: .PrintView = RpCo2.CreatePrintView()
        Case RibTab_Mitarbeit: .PrintView = RpCo2.CreatePrintView()
        Case RibTab_Fragebogen: .PrintView = RpCo5.CreatePrintView()
        Case RibTab_Krankenbla: .PrintView = RpCoK.CreatePrintView()
        Case RibTab_Abrechnung: .PrintView = RpCo6.CreatePrintView()
        Case RibTab_Tagesproto: .PrintView = RpCo6.CreatePrintView()
        Case RibTab_Vorbereit: .PrintView = RpCo6.CreatePrintView()
        Case RibTab_Rezeptmodul: .PrintView = RpCo3.CreatePrintView()
        Case RibTab_Belegmodul: .PrintView = RpCo3.CreatePrintView()
        Case RibTab_Rechnungen: .PrintView = RpCo4.CreatePrintView()
        Case RibTab_Mahnwesen: .PrintView = RpCo1.CreatePrintView()
        Case RibTab_Buchungen: .PrintView = RpCo1.CreatePrintView()
        Case RibTab_HomeBanki: .PrintView = RpCo1.CreatePrintView()
        Case RibTab_Ter_Listen: .PrintView = RpCo1.CreatePrintView()
        Case RibTab_Ter_Akont: .PrintView = RpCo1.CreatePrintView()
        Case RibTab_Ter_Warte: .PrintView = RpCo1.CreatePrintView()
        Case RibTab_LabBericht: .PrintView = RpCo5.CreatePrintView()
        Case RibTab_LabAuftrag: .PrintView = RpCo5.CreatePrintView()
        Case RibTab_LabBerichte: .PrintView = RpCo5.CreatePrintView()
        Case RibTab_LabAuftrage: .PrintView = RpCo5.CreatePrintView()
        Case RibTab_Kat_Eintrg: .PrintView = RpCo8.CreatePrintView()
        Case RibTab_Kat_Ketten: .PrintView = RpCo8.CreatePrintView()
        Case RibTab_Tex_Dokumt:
        Case RibTab_Tex_Vorlag:
        Case RibTab_Tex_Rezept:
        Case RibTab_Tex_NewsLe:
        Case RibTab_Kat_Explor:
        Case RibTab_Kat_Frage:
        Case RibTab_Tex_Email:
        End Select
    End Select
End With

FM.BackColor = GlBak

Set CaCol = Nothing
Set ChCon = Nothing
Set PrtPr = Nothing

Exit Sub

InErr:
If GlDbg = True Then SErLog Err.Description & " PrtInit " & Err.Number
Resume Next

End Sub
Public Sub PrtMain(ByVal PrCon As Integer)
On Error GoTo MeErr

If GlBut = RibTab_Startseite Then
    Exit Sub
End If

If WindowLoad("frmPrintPrev") = True Then
    Set FM = frmPrintPrev
    frmPrintPrev.ZOrder 0
    Exit Sub
End If

GlKeL = True

Load frmPrintPrev

Set FM = frmPrintPrev

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    .FenDsk 2
    If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
        .FeLin = (GlxGr - GlFeB) / 2
        .FeObn = (GlyGr - GlFeH) / 2
        .FeBre = IIf(GlxGr >= GlFeB, GlFeB, GlxGr)
        .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
    Else
        .FeLin = 0
        .FeObn = 0
        .FeBre = GlxGr
        .FeHoh = GlyGr
    End If
End With

PrtInit PrCon

With clFen
    .FenMov
    .FenDsk 3
    .FenVor
    Screen.MousePointer = vbNormal
End With

Set clFen = Nothing

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

frmPrintPrev.Show
DoEvents
GlKeL = False

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " PrtMain " & Err.Number
Resume Next

End Sub
Public Sub PrtRez()
On Error GoTo PoErr

Dim FenBr As Long
Dim FenHo As Long

Set FM = frmPrintPrev
Set PrtPr = FM.prtPrev1

FenBr = FM.ScaleWidth
FenHo = FM.ScaleHeight

If FM.WindowState <> vbMinimized Then
    PrtPr.Move 0, 0, FenBr, FenHo
End If

Exit Sub

PoErr:
If GlDbg = True Then SErLog Err.Description & " PrtRez " & Err.Number
Resume Next

End Sub
Private Sub VoDrIn(ByVal FiNan As String)
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim SeAkt As Integer
Dim SeGes As Integer

Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmDruVo
Set PrtVo = FM.LLDruVo
Set CmBrs = FM.comBar02
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag

With PrtVo
    .AsyncDownload = True
    .BackColor = -2147483643
    .CurrentPage = 1
    .Language = CMBTLANG_DEFAULT
    If GlRDP = True Then
        .SaveAsFilePath = GlIPf
    Else
        .SaveAsFilePath = GlEPf
    End If
    .ShowExitButton = False
    .ShowThumbnails = GlLiD
    .ShowUnprintableArea = True
    .SlideshowMode = False
    .ToolbarEnabled = False
    .FileURL = FiNan
    DoEvents
    .SetZoom GlZoD
    SeGes = .PaGes
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
    Set CmCon = .Add(xtpControlButton, SY_OP_Nav_Zuru, "Seite Zurück")
    With CmCon
        .ToolTipText = "Navigiert zur vorherigen Seite der Vorschau"
        .IconId = IC24_Nav_Left
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Nav_Vor, "Seite Vorw.")
    With CmCon
        .ToolTipText = "Navigiert zur nächsten Seite der Vorschau"
        .IconId = IC24_Nav_Right
        .BeginGroup = True
    End With
    Set CmCom = .Add(xtpControlComboBox, SY_OP_Ansicht, vbNullString)
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Welche Seite soll angezeigt werden?"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .BeginGroup = True
        .Width = 100
        For SeAkt = 1 To SeGes
            .AddItem "Seite " & Format$(SeAkt, "000")
        Next SeAkt
        .ListIndex = 1
    End With
    
    Set CmEdi = .Add(xtpControlEdit, SY_OP_SubDe1, "Suche :")
    With CmEdi
        .ToolTipText = "Eingabe des Suchbegriffes"
        .EditHint = "Suchbegriff eingeben"
        .Width = 160
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_SubDe2, "Zurück")
    With CmCon
        .ToolTipText = "Vorherigen Suchen"
        .IconId = IC24_Find_Prev
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_SubDe3, "Weiter")
    With CmCon
        .ToolTipText = "Nächsten Suchen"
        .IconId = IC24_Find_Next
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Exportieren")
    With CmCon
        .ToolTipText = "Speichert die Vorschau in einem anderen Format"
        .IconId = IC24_Disk_Norm
        .BeginGroup = True
        If GlDru.GoBDk = True Then
            .Enabled = GlDru.ReAbs 'WICHTIG GoBD
        End If
    End With
    Set CmCon = .Add(xtpControlSplitButtonPopup, SY_OP_Drucken, "Alle Seiten Drucken")
    With CmCon
        .ToolTipText = "Druckt alle angezeigten Seiten aus"
        .ShortcutText = "F10"
        .IconId = IC24_Printer_Ink
        .BeginGroup = True
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_UeberEinz, "Drucken mit Druckerauswahl")
        CmCon.IconId = IC16_Printer_Ink
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_UeberNetz, "Aktuelle Seite Drucken")
        CmCon.IconId = IC16_Printer_Ink
        Set CmCon = .CommandBar.Controls.Add(xtpControlButton, SY_OP_Nav_Erst, "Erste Seite Drucken")
        CmCon.IconId = IC16_Printer_Ink
        If GlDru.GoBDk = True Then
            .Enabled = GlDru.ReAbs 'WICHTIG GoBD
        End If
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

CmSta.Pane(3).Text = "Seiten: " & SeGes

FM.BackColor = GlBak

Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Set PrtVo = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "VoDrIn " & Err.Number
Resume Next

End Sub
Public Sub VoDrMa(ByVal FiNam As String)
On Error GoTo MeErr

Dim CmBrs As XtremeCommandBars.CommandBars

If WindowLoad("frmDruVo") = True Then
    Set FM = frmDruVo
    frmDruVo.ZOrder 0
    Exit Sub
End If

GlKeL = True

frmDruVo.DaNam = FiNam
DoEvents
Load frmDruVo

Set FM = frmDruVo

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    .FenDsk 2
    If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
        .FeLin = (GlxGr - GlFeB) / 2
        .FeObn = (GlyGr - GlFeH) / 2
        .FeBre = IIf(GlxGr >= GlFeB, GlFeB, GlxGr)
        .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
    Else
        .FeLin = 0
        .FeObn = 0
        .FeBre = GlxGr
        .FeHoh = GlyGr
    End If
End With

AFont FM
VoDrIn FiNam
DoEvents

With clFen
    .FenMov
    DoEvents
    VoDrPo
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

If GlRah = True Then
    SFrame 1, FM.hwnd
End If

VoDrPo
DoEvents

frmDruVo.Show

GlKeL = False

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " VoDrMa " & Err.Number
Resume Next

End Sub
Public Sub VoDrPo()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmDruVo
Set CmBrs = FM.comBar02
Set PrtVo = FM.LLDruVo

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    If ClBre > 1000 And ClHoh > 1000 Then
        PrtVo.Move ClLin, ClObn, ClBre, ClHoh
    End If
End If

Set CmBrs = Nothing
Set PrtVo = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "VoDrPo " & Err.Number
Resume Next

End Sub
Private Sub VoTxIni(ByVal FiNam As String, ByVal DaNam As String, Optional ByVal TxTyp As Integer)
On Error GoTo MnErr
'Legt alle Menüs und Toolleisten an

Dim SuFix As String
Dim SeGes As Integer
Dim SeAkt As Integer
Dim Posit As Integer

Dim CmBrs As XtremeCommandBars.CommandBars
Dim ImMan As XtremeCommandBars.ImageManager
Dim CmEdi As XtremeCommandBars.CommandBarEdit
Dim CmCon As XtremeCommandBars.CommandBarControl
Dim CmCoS As XtremeCommandBars.CommandBarControls
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmTxVor
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont3
Set TxRu1 = FM.TexRule1
Set TxRu2 = FM.TexRule2
Set CmSta = CmBrs.StatusBar
Set CmOpt = CmBrs.Options
Set CmAcs = CmBrs.Actions
Set ImMan = frmMain.imgManag

Posit = InStrRev(FiNam, ".", Len(FiNam), 1)
If Posit > 0 Then
    SuFix = LCase$(Mid$(FiNam, Posit + 1, Len(FiNam) - Posit))
Else
    SuFix = vbNullString
End If

With TxRu1
    .Appearance = txColorScheme
    .direction = txHorizontal
    .EnablePageMargins = True
    .Language = 49
    .ScaleUnits = 0
    .DisplayColor(rlBackColor) = GlBkk
    .DisplayColor(rlGradientBackColor) = GlBkk
End With

With TxRu2
    .Appearance = txColorScheme
    .direction = txVertical
    .EnablePageMargins = True
    .Language = 49
    .ScaleUnits = 0
    .DisplayColor(rlBackColor) = GlBkk
    .DisplayColor(rlGradientBackColor) = GlBkk
End With

With TxCoN
    .ViewMode = GlViW 'ViewMode Textvorschau
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
    .EditMode = 1 'Readonly
    .FontBold = GlXFt.Bold
    .FontItalic = GlXFt.Italic
    .FontUnderline = GlXFt.Underline
    .FontStrikethru = GlXFt.Strikethrough
    If GlViW <> 2 Then .FontName = GlXFt.Name
    .FontSize = GlXFt.SIZE
    .ForeColor = vbBlack
    .FormatSelection = True
    .HeaderFooterStyle = txDividingLine + txMouseClick
    .HideSelection = False
    .InsertionMode = True
    .Language = 49
    .LineSpacing = 110
    .PageViewStyle = txGradientColors
    .PageHeight = (297 / 10) * 567
    .PageWidth = (210 / 10) * 567
    .PageMarginL = (10 / 10) * 567
    .PageMarginR = (10 / 10) * 567
    .PageMarginT = (5 / 10) * 567
    .PageMarginB = (5 / 10) * 567
    .PageOrientation = 0
    .PrintColors = True
    .ScrollBars = 3
    .SizeMode = 0
    .SelectionViewMode = 1
    .TabKey = True
    .TextBkColor = 16777215
    .TextFrameMarkerLines = True
    .RulerHandle = TxRu1.hwnd
    .VerticalRulerHandle = TxRu2.hwnd
    .TableGridLines = True
    .EnableHyperlinks = True
    .ZoomFactor = GlZoW
    .WordWrapMode = 1
    .DisplayColor(txDesktopColor) = GlBkk
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

'---

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
    Set CmCom = .Add(xtpControlComboBox, SY_OP_Ansicht, " Seite:")
    With CmCom
        .CloseSubMenuOnClick = True
        .DropDownListStyle = False
        .ToolTipText = "Welche Seite soll angezeigt werden?"
        .Style = xtpButtonAutomatic
        .ThemedItems = True
        .BeginGroup = True
        .Width = 130
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Kopieren, "Kopieren")
    With CmCon
        .ToolTipText = "Kopiert den markierten Text in die Zwischenablage"
        .IconId = IC24_Copy
        .BeginGroup = True
    End With
    Set CmEdi = .Add(xtpControlEdit, SY_OP_SubDe1, "Suche :")
    With CmEdi
        .ToolTipText = "Eingabe des Suchbegriffes"
        .EditHint = "Suchbegriff eingeben"
        .Width = 160
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_SubDe2, "Zurück")
    With CmCon
        .ToolTipText = "Vorherigen Suchen"
        .IconId = IC24_Find_Prev
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_SubDe3, "Weiter")
    With CmCon
        .ToolTipText = "Nächsten Suchen"
        .IconId = IC24_Find_Next
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Speichern, "Exportieren")
    With CmCon
        .ToolTipText = "Speichert die Vorschau in einem anderen Format"
        .IconId = IC24_Doc_Out
        .BeginGroup = True
    End With
    Set CmCon = .Add(xtpControlButton, SY_OP_Drucken, " Drucken")
    With CmCon
        .ToolTipText = "Druckt die angezeigten Seiten aus"
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

FM.BackColor = GlBak

Set CmCom = CmBrs.FindControl(CmCom, SY_OP_Ansicht, , True)

GlTxF = FiNam 'Filname für Textcontrol Error
GlTxU = LCase(SuFix)
With TxCoN
    .ResetContents
    .Load FiNam, , TxTyp
    SeGes = .CurrentPages
End With

With CmCom
    For SeAkt = 1 To SeGes
        .AddItem "Seite " & Format$(SeAkt, "000")
    Next SeAkt
    .ListIndex = 1
End With

CmSta.Pane(3).Text = "Seiten: " & SeGes

Set CmSta = Nothing
Set CmOpt = Nothing
Set CmBar = Nothing
Set CmBrs = Nothing
Set ImMan = Nothing

Set TxCoN = Nothing

Exit Sub

MnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "VoTxIni " & Err.Number
Resume Next

End Sub
Public Sub VoTxMa(ByVal FiNam As String, ByVal DaNam As String, Optional ByVal TxTyp As Integer)
On Error GoTo MeErr

If WindowLoad("frmTxVor") = True Then
    Set FM = frmTxVor
    frmTxVor.ZOrder 0
    Exit Sub
End If

GlKeL = True

frmTxVor.DaNam = DaNam
DoEvents
Load frmTxVor

Set FM = frmTxVor

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

With clFen
    Screen.MousePointer = vbHourglass
    .FenDsk 2
    If Right$(GlFeG, 1) = 3 Then 'Fenstergröße Programmstart
        .FeLin = (GlxGr - GlFeB) / 2
        .FeObn = (GlyGr - GlFeH) / 2
        .FeBre = IIf(GlxGr >= GlFeB, GlFeB, GlxGr)
        .FeHoh = IIf(GlyGr >= GlFeH, GlFeH, GlyGr)
    Else
        .FeLin = 0
        .FeObn = 0
        .FeBre = GlxGr
        .FeHoh = GlyGr
    End If
End With

AFont FM
VoTxIni FiNam, DaNam, TxTyp
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

VoTxPo
DoEvents

frmTxVor.Show

GlKeL = False

Exit Sub

MeErr:
If GlDbg = True Then SErLog Err.Description & " VoTxMa " & Err.Number
Resume Next

End Sub

Public Sub VoTxPo()
On Error GoTo InErr

Dim ClLin As Long
Dim ClObn As Long
Dim ClBre As Long
Dim ClHoh As Long
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmTxVor
Set CmBrs = FM.comBar02
Set TxCoN = FM.TexCont3
Set TxRu1 = FM.TexRule1
Set TxRu2 = FM.TexRule2

If FM.WindowState <> vbMinimized Then
    CmBrs.GetClientRect ClLin, ClObn, ClBre, ClHoh
    ClHoh = ClHoh - ClObn
    If ClBre > 1000 And ClHoh > 1000 Then
        If GlLiW = True Then 'Lineal Textvorschau
            TxCoN.Move ClLin + 380, ClObn + 380, ClBre - 380, ClHoh - 380
            TxRu1.Move ClLin, ClObn, ClBre - ClLin, 380
            TxRu2.Move ClLin, ClObn + 380, 380, ClHoh - 380
        Else
            TxCoN.Move ClLin, ClObn, ClBre - ClLin, ClHoh
        End If
    End If
End If

Set CmBrs = Nothing
Set TxCoN = Nothing
Set TxRu1 = Nothing
Set TxRu2 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "VoTxPo " & Err.Number
Resume Next

End Sub

