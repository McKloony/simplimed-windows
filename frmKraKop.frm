VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmKraKop 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kopierassistent"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   18
      Top             =   3500
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   6000
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schließen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   4600
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Weiter"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnZurück 
         Height          =   400
         Left            =   3200
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zurück"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   1900
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hilfe"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3400
      Left            =   3
      TabIndex        =   16
      Top             =   0
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   5997
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkExpMo 
         Height          =   220
         Left            =   2300
         TabIndex        =   4
         Top             =   2500
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "mehrere Tage markieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optAnder 
         Height          =   220
         Left            =   2300
         TabIndex        =   3
         Top             =   1700
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "in eine andere Rechnung kopieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optSelbe 
         Height          =   220
         Left            =   2300
         TabIndex        =   2
         Top             =   1300
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "in diese Rechnung kopieren"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin VB.Label lblLabl1 
         BackStyle       =   0  'Transparent
         Caption         =   "In welche Rechnung möchten Sie die ausgewählten Positionen kopieren? Wählen Sie eine Option und klicken auf Weiter."
         Height          =   400
         Left            =   1400
         TabIndex        =   17
         Top             =   200
         Width           =   5000
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   8000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3400
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   5997
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtReNum 
         Height          =   350
         Left            =   1800
         TabIndex        =   11
         Top             =   2000
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtKurz 
         Height          =   350
         Left            =   1800
         TabIndex        =   10
         Top             =   1300
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin VB.Label lblLabl3 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Rechnungsnummer"
         Height          =   195
         Left            =   1810
         TabIndex        =   15
         Top             =   1740
         Width           =   3000
      End
      Begin VB.Label lblLabl2 
         BackStyle       =   0  'Transparent
         Caption         =   "Suche nach Patientenname"
         Height          =   195
         Left            =   1810
         TabIndex        =   14
         Top             =   1040
         Width           =   3000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte geben Sie jetzt entweder den Patientennamen oder direkt die gewünschte Rechnungsnummer ein und klicken auf Weiter."
         Height          =   400
         Left            =   1200
         TabIndex        =   13
         Top             =   100
         Width           =   5500
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   3400
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   5997
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   2800
         Left            =   500
         TabIndex        =   9
         Top             =   500
         Width           =   6800
         _Version        =   1048579
         _ExtentX        =   11994
         _ExtentY        =   4939
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin VB.Label lblLabl5 
         BackStyle       =   0  'Transparent
         Caption         =   "In welche Rechnung sollen die Einträge kopiert werden?"
         Height          =   200
         Left            =   510
         TabIndex        =   8
         Top             =   100
         Width           =   6500
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   3400
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   5997
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeCalendarControl.DatePicker dtpDatu1 
         Height          =   2600
         Left            =   0
         TabIndex        =   6
         Top             =   700
         Width           =   8000
         _Version        =   1048579
         _ExtentX        =   14111
         _ExtentY        =   4586
         _StockProps     =   64
         Show3DBorder    =   0
         ColumnCount     =   3
      End
      Begin VB.Label lblLabl6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte markieren Sie die Tage, für die Sie die markierten Einträge kopieren möchten."
         Height          =   400
         Left            =   500
         TabIndex        =   5
         Top             =   100
         Width           =   6500
      End
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   80
   End
End
Attribute VB_Name = "frmKraKop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private ChExp As XtremeSuiteControls.CheckBox
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private RpRow As XtremeReportControl.ReportRow
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ImMan As XtremeCommandBars.ImageManager
Private DaPi1 As XtremeCalendarControl.DatePicker

Private PatNr As Long
Private ReNum As Long
Private ReBet As Double
Private GeBet As Double
Private NeStr As String
Private AbExp As Boolean
Private Sub FDatu()
On Error GoTo OrErr

Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim AnzTa As Integer
Dim AnzBl As Integer
Dim AktBl As Integer
Dim AktTa As Integer
Dim BloTa As Integer

Set DaPi1 = Me.dtpDatu1
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ

AktTa = 0
AnzBl = DaPi1.Selection.BlocksCount

If AnzBl = 1 Then
    DaBeg = DaPi1.Selection(0).DateBegin
    DaEnd = DaPi1.Selection(0).DateEnd
    If DaEnd > DaBeg Then
        Do
        DaAkt = DaBeg + AktTa
        AktTa = AktTa + 1
        ReDim Preserve GlTag(AktTa)
        GlTag(AktTa) = DaAkt
        Loop Until DaAkt >= DaEnd
    Else
        ReDim Preserve GlTag(1)
        GlTag(1) = DaBeg
    End If
ElseIf AnzBl > 1 Then
    For AktBl = 0 To AnzBl - 1
        DaBeg = DaPi1.Selection.Blocks(AktBl).DateBegin
        DaEnd = DaPi1.Selection.Blocks(AktBl).DateEnd
        If DaEnd > DaBeg Then
            BloTa = 0
            Do
            DaAkt = DaBeg + BloTa
            AktTa = AktTa + 1
            BloTa = BloTa + 1
            ReDim Preserve GlTag(AktTa)
            GlTag(AktTa) = DaAkt
            Loop Until DaAkt >= DaEnd
        Else
            AktTa = AktTa + 1
            ReDim Preserve GlTag(AktTa)
            GlTag(AktTa) = DaBeg
        End If
    Next AktBl
End If

AnzTa = UBound(GlTag)

Me.Caption = "Kopierassistent (" & Format$((GeBet * AnzTa) + ReBet, GlWa1) & ")"

PuBu1.Enabled = True

If GlPop = True Then
    S_AbTa
    S_AbDo
End If

Set DaPi1 = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FExp()
On Error GoTo InErr

Set ChExp = Me.chkExpMo
Set DaPi1 = Me.dtpDatu1

If ChExp.Value = 1 Then
    IniSetVal "System", "KopExp", -1
    AbExp = True
Else
    IniSetVal "System", "KopExp", 0
    AbExp = False
End If

DaPi1.MultiSelectionMode = AbExp

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FExp " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo SuErr

Dim NeuDa As Date
Dim DayFi As Date
Dim DayLa As Date
Dim GesZa As Long
Dim ZeiUm As Boolean
Dim AktPo As Integer
Dim AnzPo As Integer
Dim TeWer As Variant
Dim Mld1, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set ImMan = FM.imgManag
Set RpCon = Me.repCont1
Set RpCo3 = FM.repCont3
Set RpCo6 = FM.repCont6
Set ChExp = Me.chkExpMo
Set DaPi1 = Me.dtpDatu1
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück
Set RpCls = RpCon.Columns

AbExp = CBool(IniGetVal("System", "KopExp"))
ZeiUm = False

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
    If ZeiUm = True Then
        .PaintManager.FixedRowHeight = False
    Else
        .PaintManager.FixedRowHeight = True
    End If
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
    .PreviewMode = False
    .ShowHeader = GlGKo
    .ScrollModeH = xtpReportScrollModeSmooth
    .ScrollModeV = xtpReportScrollModeBlock
End With

With DaPi1
    .AllowNoncontinuousSelection = True
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    .BorderStyle = 0
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    .MaxSelectionCount = GlMxK 'Maximal slektierbare Kalendertage
    .MultiSelectionMode = AbExp
    .RightToLeft = False
    .ShowNoneButton = True
    .ShowNonMonthDays = True
    .ShowTodayButton = True
    .ShowWeekNumbers = False
    .TextNoneButton = "Markiere Keine"
    .TextTodayButton = "Markiere Heute"
    .ToolTipText = vbNullString
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

With RpCls
    Set RpCol = .Add(0, "Nummer", 0, False)
    Set RpCol = .Add(1, "Rechnung", 100, True)
    Set RpCol = .Add(2, "Datum", 80, True)
    Set RpCol = .Add(3, "Patient", 120, True)
    RpCol.AutoSize = True
    Set RpCol = .Add(4, "Betrag", 60, True)
    RpCol.Alignment = xtpAlignmentRight
    Set RpCol = .Add(5, "Selekt", 0, True)
    Set RpCol = .Add(6, "ID0", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Editable = False
    RpCol.Groupable = True
    RpCol.Sortable = False
Next RpCol

Set RpCls = RpCo3.Columns
Set RpSel = RpCo3.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Rec_ID0)
        PatNr = RpRow.Record(RpCol.ItemIndex).Value
        Set RpCol = RpCls.Find(Rec_Betrag)
        ReBet = CSng(RpRow.Record(RpCol.ItemIndex).Value)
    Else
        Me.btnWeiter.Enabled = False
        Mld1 = "Es wurden keine Einträge markiert"
        Tit1 = "Keine Einträge"
        WindowMess Mld1, Dial2, Tit1, FM.hwnd
    End If
Else
    Me.btnWeiter.Enabled = False
    Mld1 = "Es wurden keine Einträge markiert"
    Tit1 = "Keine Einträge"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If
        
Set RpCls = RpCo6.Columns
Set RpSel = RpCo6.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Kra_Datum)
        NeuDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
        With DaPi1
            .EnsureVisible NeuDa
            DayFi = .FirstDayOfWeek
            DayLa = .LastVisibleDay
        End With
        S_AbTe DayFi, DayLa
    End If
Else
    Me.btnWeiter.Enabled = False
    Mld1 = "Es wurden keine Einträge markiert"
    Tit1 = "Keine Einträge"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

If ClKop = True Then
    AnzPo = UBound(GlClp)
    If GlClp(AktPo, 2) <> 6 Then
        For AktPo = 1 To AnzPo
            If GlClp(AktPo, 9) <> vbNullString Then
                GeBet = GeBet + Replace(GlClp(AktPo, 9), ".", ",", 1)
            End If
        Next AktPo
    End If
End If

If AbExp = True Then
    ChExp.Value = 1
End If

PuBu3.Enabled = False

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak
ChExp.BackColor = GlBak
Opti1.BackColor = GlBak
Opti2.BackColor = GlBak

PatNr = GlAdr

Set DaPi1 = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set RpCo3 = Nothing
Set RpCo6 = Nothing
Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FSuda()
On Error GoTo SuErr

Dim GesZa As Long
Dim Mld1, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtReNum
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set RpCon = Me.repCont1
Set RpRcs = RpCon.Records

If FTex1.Text <> vbNullString Then
    S_KrFi FTex1.Text, 5, 5
ElseIf FTex2.Text <> vbNullString Then
    S_KrFi FTex2.Text, 6, 5
End If

GesZa = RpRcs.Count

If GesZa > 0 Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
Else
    Mld1 = "Das von Ihnen eingegebene Suchkriterium brachte leider keine Suchergebnisse"
    Tit1 = "Adressuche"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

Set RpCon = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSuda " & Err.Number
Resume Next

End Sub
Private Sub FZur()
On Error Resume Next

Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück

If Rahm2.Visible = True Then
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = False
    Rahm1.Visible = True
    PuBu1.Enabled = True
    PuBu3.Enabled = False
ElseIf Rahm3.Visible = True Then
    Rahm4.Visible = False
    Rahm3.Visible = False
    Rahm2.Visible = True
    Rahm1.Visible = False
    PuBu1.Enabled = True
ElseIf Rahm4.Visible = True Then
    If Opti1.Value = True Then
        Rahm4.Visible = False
        Rahm3.Visible = False
        Rahm2.Visible = False
        Rahm1.Visible = True
        PuBu1.Enabled = True
        PuBu3.Enabled = False
    Else
        Rahm4.Visible = False
        Rahm3.Visible = True
        Rahm2.Visible = False
        Rahm1.Visible = False
        PuBu1.Enabled = True
    End If
End If

End Sub
Private Sub TWeit()
On Error GoTo SuErr

Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim DayFi As Date
Dim DayLa As Date
Dim RowNr As Long
Dim KrRow As Long
Dim AnzBl As Long
Dim AktBl As Long
Dim AktTa As Long
Dim BloTa As Long
Dim RecNr As Long
Dim Frage As Long
Dim FoWär As Integer
Dim ZaZil As Integer
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim ReStr As String
Dim KrGui As String
Dim GeSum As Double
Dim ReSum As Double
Dim ReRab As Double
Dim NeReB As Double
Dim BeBet As Double
Dim RetWe As Boolean
Dim ReAbg As Boolean
Dim Mld1, Mld2, Tit1 As String
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCo6 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4
Set RpCo6 = FM.repCont6
Set RpSel = RpCo3.SelectedRows
Set RpCls = RpCo3.Columns
Set RpRws = RpCo3.Rows

Set FTex1 = Me.txtKurz
Set FTex2 = Me.txtReNum
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set DaPi1 = Me.dtpDatu1
Set RpCon = Me.repCont1
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück

TeTit = "Einträge Einfügen"
TeMai = "Diese Rechnung wurde bereits abgeschlossen!"
TeInh = "Diese kann nicht mehr geändert werden, da diese verriegelt ist. Um diese Rechnung zu verändern, muss diese zuerst entriegelt werden."
TeFus = "Um eine Rechnung zu entriegeln markieren Sie diese bitte mit der rechten Maustaste und wählen die gleichnamige Funktion."

If RpSel.Count > 0 Then
    If Opti1.Value = True Then
        Set RpRow = RpSel(0)
    Else
        Set RpRow = RpRws(0)
    End If
    If RpRow.GroupRow = False Then
        RowNr = RpRow.Index
        If Rahm1.Visible = True Then
            Set RpCol = RpCls.Find(Rec_ID1)
            ReNum = RpRow.Record(RpCol.ItemIndex).Value
            Set RpCol = RpCls.Find(Rec_Betrag)
            ReBet = CSng(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Rec_RechNr)
            ReStr = Trim$(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Rec_Selekt)
            ReAbg = CBool(RpRow.Record(RpCol.ItemIndex).Value)
            Set RpCol = RpCls.Find(Rec_Rabatt)
            ReRab = Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1)
            Set RpCol = RpCls.Find(Rec_Bezahlt)
            BeBet = RpRow.Record(RpCol.ItemIndex).Value

            If Opti1.Value = True Then
                If ReAbg = True Then
                    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, frmKraKop.hwnd
                    Set DaPi1 = Nothing
                    Set RpSel = Nothing
                    Set RpRws = Nothing
                    Set RpCls = Nothing
                    Set RpCon = Nothing
                    Set RpCo3 = Nothing
                    Set RpCo4 = Nothing
                    Set RpCo6 = Nothing
                    Exit Sub
                Else
                    Rahm1.Visible = False
                    Rahm2.Visible = False
                    Rahm3.Visible = False
                    Rahm4.Visible = True
                    PuBu1.Enabled = False
                End If
            Else
                Rahm1.Visible = False
                Rahm2.Visible = True
                Rahm3.Visible = False
                Rahm4.Visible = False
                FTex2.Text = ReStr
            End If
            PuBu3.Enabled = True
        ElseIf Rahm2.Visible = True Then
            FSuda
        ElseIf Rahm3.Visible = True Then
            Set RpSel = RpCon.SelectedRows
            Set RpCls = RpCon.Columns
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    ReNum = RpRow.Record(0).Value
                    ReBet = RpRow.Record(4).Value
                    PatNr = RpRow.Record(6).Value
                    If S_KrAb(ReNum) = True Then
                        SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, frmKraKop.hwnd
                        Set DaPi1 = Nothing
                        Set RpSel = Nothing
                        Set RpRws = Nothing
                        Set RpCls = Nothing
                        Set RpCon = Nothing
                        Set RpCo3 = Nothing
                        Set RpCo4 = Nothing
                        Set RpCo6 = Nothing
                        Exit Sub
                    Else
                        Rahm1.Visible = False
                        Rahm2.Visible = False
                        Rahm3.Visible = False
                        Rahm4.Visible = True
                        PuBu1.Enabled = False
                    End If
                End If
            End If
        ElseIf Rahm4.Visible = True Then
            AktTa = 0
            AnzBl = DaPi1.Selection.BlocksCount
            With DaPi1
                DayFi = .FirstDayOfWeek
                DayLa = .LastVisibleDay
            End With
            If AnzBl = 1 Then
                DaBeg = DaPi1.Selection(0).DateBegin
                DaEnd = DaPi1.Selection(0).DateEnd
                If DaEnd > DaBeg Then
                    Do
                    DaAkt = DaBeg + AktTa
                    AktTa = AktTa + 1
                    ReDim Preserve GlTag(AktTa)
                    GlTag(AktTa) = DaAkt
                    Loop Until DaAkt >= DaEnd
                Else
                    ReDim Preserve GlTag(1)
                    GlTag(1) = DaBeg
                End If
            ElseIf AnzBl > 1 Then
                For AktBl = 0 To AnzBl - 1
                    DaBeg = DaPi1.Selection.Blocks(AktBl).DateBegin
                    DaEnd = DaPi1.Selection.Blocks(AktBl).DateEnd
                    If DaEnd > DaBeg Then
                        BloTa = 0
                        Do
                        DaAkt = DaBeg + BloTa
                        AktTa = AktTa + 1
                        BloTa = BloTa + 1
                        ReDim Preserve GlTag(AktTa)
                        GlTag(AktTa) = DaAkt
                        Loop Until DaAkt >= DaEnd
                    Else
                        AktTa = AktTa + 1
                        ReDim Preserve GlTag(AktTa)
                        GlTag(AktTa) = DaBeg
                    End If
                Next AktBl
            End If
            
            Set RpSel = RpCo6.SelectedRows
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                If RpRow.GroupRow = False Then
                    Set RpCol = RpCls.Find(Rec_Wahrung)
                    If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
                        If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) Then
                            FoWär = CInt(RpRow.Record(RpCol.ItemIndex).Value)
                        Else
                            FoWär = 2
                        End If
                    Else
                        FoWär = 2
                    End If
                End If
            End If
            
            Set RpCls = RpCo3.Columns
            Set RpSel = RpCo3.SelectedRows
            If RpSel.Count > 0 Then
                Set RpRow = RpSel(0)
                KrRow = RpRow.Index
                Set RpCol = RpCls.Find(Rec_ID1)
                RecNr = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_IDZ)
                ZaZil = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Betrag)
                ReSum = RpRow.Record(RpCol.ItemIndex).Value
                Set RpCol = RpCls.Find(Rec_Rabatt)
                ReRab = Format$(RpRow.Record(RpCol.ItemIndex).Value, GlWa1)
                Set RpCol = RpCls.Find(Rec_Bezahlt)
                BeBet = RpRow.Record(RpCol.ItemIndex).Value
            End If
            KrGui = S_KrKo(ReNum, PatNr) 'Einfügen
            S_AbTe DayFi, DayLa
            NeReB = S_ReBet(ReNum, Round(ReBet, 2), ReAbg, ReRab)
            S_KrLi
            RetWe = S_KrBe(RecNr, ZaZil, ReSum, BeBet)
            DoEvents
            
            ReDim Preserve GlTag(1)
            GlTag(1) = Date
            DoEvents

            SUpAb RowNr, , KrGui
            DoEvents
            
            If GlVrz = False Then
                Set RpSel = RpCo4.SelectedRows
                If RpSel.Count > 0 Then
                    Set RpRow = RpSel(0)
                    RowNr = RpRow.Index
                    SUpRe RowNr
                End If
            Else
                GlVzA = True
            End If
                        
            If GlPop = True Then 'Rechnungsobergrenze
                If NeReB > GlObG Then
                    SReMe
                End If
            End If
            Unload Me
            DoEvents
            RpCo6.SetFocus
        End If
    End If
Else
    Mld1 = "Sie müssen erst eine neue Rechnung anlegen"
    Tit1 = "Keine Rechnung"
    WindowMess Mld1, Dial2, Tit1, FM.hwnd
End If

Set DaPi1 = Nothing
Set RpSel = Nothing
Set RpRws = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing
Set RpCo6 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TWeit " & Err.Number
Resume Next

End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 51101)
TeMai = IniGetOpt("Hilfe", 51102)
TeInh = IniGetOpt("Hilfe", 51103)
TeFus = IniGetOpt("Hilfe", 51104)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Set frmKraKop = Nothing
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    TWeit
End Sub
Private Sub btnZurück_Click()
    FZur
End Sub
Private Sub chkExpMo_Click()
    FExp
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
On Error Resume Next

Dim AktTa As Long

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTaV > 0 Then
    For AktTa = 1 To GlTaV
        If Day = GlBet(AktTa) Then
            Metrics.BackColor = GlMkr
        End If
    Next AktTa
End If

End Sub
Private Sub dtpDatu1_MonthChanged()
On Error Resume Next

Dim DayFi As Date
Dim DayLa As Date

Set DaPi1 = Me.dtpDatu1

With DaPi1
    DayFi = .FirstDayOfWeek
    DayLa = .LastVisibleDay
End With

S_AbTe DayFi, DayLa

Set DaPi1 = Nothing

End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF11: Unload Me
    End Select
End Sub
Private Sub Form_Load()
On Error Resume Next

AFont Me
SClip
FKonf
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKraKop = Nothing
End Sub
Private Sub txtKurz_GotFocus()
On Error Resume Next

Set FTex1 = Me.txtKurz

FTex1.Text = vbNullString

End Sub
Private Sub txtKurz_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
Private Sub txtReNum_GotFocus()
On Error Resume Next

Set FTex2 = Me.txtReNum

FTex2.SelStart = 0
FTex2.SelLength = Len(FTex2.Text)

End Sub
Private Sub txtReNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        FSuda
    End If
End Sub
