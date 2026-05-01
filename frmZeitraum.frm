VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmZeitraum 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Auswertungen & Berichte"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   14
      Top             =   5200
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4000
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Abbrechen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   2600
         TabIndex        =   16
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
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   1300
         TabIndex        =   15
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
      Height          =   4000
      Left            =   200
      TabIndex        =   3
      Top             =   0
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeReportControl.ReportControl repCont1 
         Height          =   2220
         Left            =   200
         TabIndex        =   1
         Top             =   800
         Width           =   5160
         _Version        =   1048579
         _ExtentX        =   9102
         _ExtentY        =   3916
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte wählen Sie die gewünschte Auswertung und klicken dann auf Weiter. Wähen Sie dann den gewünschten Zeitraum"
         Height          =   600
         Left            =   300
         TabIndex        =   5
         Top             =   150
         Width           =   5000
      End
      Begin VB.Label lblLab04 
         BackStyle       =   0  'Transparent
         Height          =   800
         Left            =   300
         TabIndex        =   4
         Top             =   3200
         Width           =   5000
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4000
      Left            =   200
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.RadioButton optWoche 
         Height          =   220
         Left            =   1000
         TabIndex        =   21
         Top             =   1000
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Woche :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   310
         Left            =   3220
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3360
         Width           =   310
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   547
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optDatum 
         Height          =   220
         Left            =   1000
         TabIndex        =   29
         Top             =   3400
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Datum :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optJahre 
         Height          =   220
         Left            =   1000
         TabIndex        =   27
         Top             =   2800
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Jahr :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optQuart 
         Height          =   220
         Left            =   1000
         TabIndex        =   25
         Top             =   2200
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Quartal :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optMonat 
         Height          =   220
         Left            =   1000
         TabIndex        =   23
         Top             =   1600
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Monat :"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMonat 
         Height          =   310
         Left            =   2000
         TabIndex        =   24
         Top             =   1560
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbQuart 
         Height          =   310
         Left            =   2000
         TabIndex        =   26
         Top             =   2160
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatum 
         Height          =   310
         Left            =   2000
         TabIndex        =   30
         Top             =   3360
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu1 
         Height          =   400
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   5000
         Visible         =   0   'False
         Width           =   400
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   706
         _StockProps     =   64
         Show3DBorder    =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbJahre 
         Height          =   310
         Left            =   2000
         TabIndex        =   28
         Top             =   2760
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbWoche 
         Height          =   310
         Left            =   2000
         TabIndex        =   22
         Top             =   960
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte stellen Sie den gewünschten Zeitraum der Auswertung ein und klicken dann auf Weiter."
         Height          =   400
         Left            =   800
         TabIndex        =   32
         Top             =   150
         Width           =   3700
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   4000
      Left            =   200
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.RadioButton optEinGe 
         Height          =   220
         Left            =   1100
         TabIndex        =   10
         Top             =   2900
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Für ein bestimmtes Geldkonto :"
         Appearance      =   12
      End
      Begin XtremeSuiteControls.RadioButton optAllGe 
         Height          =   220
         Left            =   1100
         TabIndex        =   9
         Top             =   2500
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Für alle Geldkonten"
         Appearance      =   12
      End
      Begin XtremeSuiteControls.RadioButton optEinKo 
         Height          =   220
         Left            =   1100
         TabIndex        =   7
         Top             =   1300
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Für ein bestimmtes Sachkonto :"
         Appearance      =   12
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optAllKo 
         Height          =   220
         Left            =   1100
         TabIndex        =   6
         Top             =   900
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Für alle Sachkonten"
         Appearance      =   12
      End
      Begin XtremeSuiteControls.ComboBox cmbKonto 
         Height          =   310
         Left            =   1100
         TabIndex        =   8
         Top             =   1640
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Appearance      =   12
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGegen 
         Height          =   310
         Left            =   1100
         TabIndex        =   11
         Top             =   3240
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Appearance      =   12
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox2"
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   8
         X1              =   130
         X2              =   5200
         Y1              =   2200
         Y2              =   2200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   9
         X1              =   130
         X2              =   5200
         Y1              =   2220
         Y2              =   2220
      End
      Begin VB.Label lblLab05 
         BackStyle       =   0  'Transparent
         Caption         =   "Soll die Auswertung für ein bestimmtes Konto oder für alle Konten ausgegeben werden? Bitte wählen Sie eine Option."
         Height          =   600
         Left            =   600
         TabIndex        =   19
         Top             =   150
         Width           =   4400
      End
   End
   Begin XtremeSuiteControls.ComboBox cmbBehan 
      Height          =   315
      Left            =   1400
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4140
      Width           =   3200
      _Version        =   1048579
      _ExtentX        =   5662
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbMitar 
      Height          =   315
      Left            =   1400
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4700
      Width           =   3200
      _Version        =   1048579
      _ExtentX        =   5662
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Mitarbeiter :"
      Height          =   200
      Left            =   420
      TabIndex        =   34
      Top             =   4760
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "Mandant :"
      Height          =   200
      Left            =   420
      TabIndex        =   2
      Top             =   4200
      Width           =   900
   End
End
Attribute VB_Name = "frmZeitraum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Labe1 As VB.Label
Private List4 As XtremeSuiteControls.ListBox
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private TxDum As XtremeSuiteControls.FlatEdit
Private CmWoc As XtremeSuiteControls.ComboBox
Private CmMon As XtremeSuiteControls.ComboBox
Private CmQua As XtremeSuiteControls.ComboBox
Private CmJah As XtremeSuiteControls.ComboBox
Private CmEiK As XtremeSuiteControls.ComboBox
Private CmEiG As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private OpWoc As XtremeSuiteControls.RadioButton
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private OpDat As XtremeSuiteControls.RadioButton
Private OpAlK As XtremeSuiteControls.RadioButton
Private OpEiK As XtremeSuiteControls.RadioButton
Private OpAlG As XtremeSuiteControls.RadioButton
Private OpEiG As XtremeSuiteControls.RadioButton
Private TxDat As XtremeSuiteControls.FlatEdit
Private MoKal As XtremeCalendarControl.DatePicker
Private PuBu1 As XtremeSuiteControls.PushButton
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ImMan As XtremeCommandBars.ImageManager

Private AnSal As Double
Private FrLad As Boolean
Private MitWa As Boolean
Private ZeiWa As Boolean

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date
Dim AkJah As Long
Dim SeJah As Long

Set TxDat = Me.txtDatum
Set MoKal = Me.dtpDatu1
Set CmJah = Me.cmbJahre

If IsDate(TxDat.Text) Then
    NeuDa = CDate(TxDat.Text)
    TxDat.Text = Format$(NeuDa, "dd.mm.yyyy")
End If

SeJah = Year(NeuDa)
AkJah = CLng(CmJah.Text)

With MoKal
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

If AkJah <> SeJah Then
    CmJah.Text = SeJah
End If

If NeuDa > Date Then
    SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDat = Me.txtDatum
Set MoKal = Me.dtpDatu1
Set OpDat = Me.optDatum

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = CDate(MoKal.Selection.Blocks(0).DateBegin)
    TxDat.Text = Format$(NeuDa, "dd.mm.yyyy")
    TxDat.SetFocus
End If

OpDat.Value = True

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date
Dim TmDat As Date

Set TxDat = Me.txtDatum
Set MoKal = Me.dtpDatu1

If IsDate(TxDat.Text) = True Then
    NeuDa = CDate(TxDat.Text)
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = TxDat.Top + TxDat.Height
    .Left = TxDat.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TmDat = CDate(.Selection.Blocks(0).DateBegin)
            TxDat.Text = Format$(TmDat, "dd.mm.yyyy")
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FKonf()
On Error GoTo InErr

Dim RetWe As Long
Dim ManNr As Long
Dim AkWoc As Integer
Dim AkMon As Integer
Dim AkQua As Integer
Dim IdxZa As Integer
Dim BuJah As Integer
Dim AktZa As Integer
Dim ZeiUm As Boolean
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCon As XtremeReportControl.ReportControl

Set FM = frmMain
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set RpCon = Me.repCont1
Set OpWoc = Me.optWoche
Set OpMon = Me.optMonat
Set OpQua = Me.optQuart
Set OpJah = Me.optJahre
Set OpDat = Me.optDatum
Set OpAlK = Me.optAllKo
Set OpEiK = Me.optEinKo
Set OpAlG = Me.optAllGe
Set OpEiG = Me.optEinGe
Set CmWoc = Me.cmbWoche
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQuart
Set CmJah = Me.cmbJahre
Set CmEiK = Me.cmbKonto
Set CmEiG = Me.cmbGegen
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set TxDat = Me.txtDatum
Set MoKal = Me.dtpDatu1
Set PuBu1 = Me.btnDatu1
Set ImMan = FM.imgManag
Set RpCls = RpCon.Columns
Set RpRcs = RpCon.Records

ZeiUm = False

ManNr = GlMiA(GlSmI, 7)
AkMon = Month(Date)
AkWoc = Format$(Date, "ww", vbMonday, vbFirstFourDays)

If AkMon <= 3 Then
    AkQua = 1
ElseIf AkMon <= 6 Then
    AkQua = 2
ElseIf AkMon <= 9 Then
    AkQua = 3
ElseIf AkMon <= 12 Then
    AkQua = 4
End If

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

With CmWoc
    .DropDownItemCount = 10
    For IdxZa = 1 To 53
        .AddItem Format$(IdxZa, "00") & " KW"
        .ItemData(.NewIndex) = IdxZa
    Next IdxZa
End With

With CmMon
    .DropDownItemCount = 12
    For IdxZa = 1 To 12
        .AddItem MonthName(IdxZa)
        .ItemData(.NewIndex) = IdxZa
    Next IdxZa
End With

With CmQua
    .DropDownItemCount = 4
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

With RpCls
    Set RpCol = .Add(0, "Formulare", 0, False)
    Set RpCol = .Add(1, "Bericht", 200, True)
    RpCol.AutoSize = True
    Set RpCol = .Add(2, "Kommentar", 0, False)
End With

For Each RpCol In RpCls
    RpCol.Alignment = xtpAlignmentIconLeft
    RpCol.Editable = False
    RpCol.Groupable = True
    RpCol.Sortable = False
Next RpCol

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("BuJour")
Set RpItm = RpRec.AddItem("Buchungsjournal")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt alle eingetragenen Buchungen in chronologischer Sortierung für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("KonAus")
Set RpItm = RpRec.AddItem("Einzelkontenauswertung")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt alle eingetragenen Buchungen für ein auszuwählendes Buchungs- oder Geldkonto in chronologischer Sortierung für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("KonGru")
Set RpItm = RpRec.AddItem("Einzelkontenauswertung (Gruppiert Sachkonten)")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt alle eingetragenen Buchungen für ein auszuwählendes Sachkonto in gruppierter Sichtweise für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("KonGel")
Set RpItm = RpRec.AddItem("Einzelkontenauswertung (Gruppiert Geldkonten)")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt alle eingetragenen Buchungen für ein auszuwählendes Geldkonto in gruppierter Sichtweise für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("EiUbMo")
Set RpItm = RpRec.AddItem("Einnahmenüberschussrechnung")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt die Einnahme- und Ausgabesummen des jeweiligen Monats für den eingestellten Auswertungszeitraum. Nutzen Sie diese Auswertung für Ihre Gewinnermittlung")
    
Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("EiUbGr")
Set RpItm = RpRec.AddItem("Einnahmenüberschussrechnung (Gruppiert)")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt die Einnahme- und Ausgabesummen des jeweiligen Jahres für den eingestellten Auswertungszeitraum gruppiert nach Monaten. Nutzen Sie diese Auswertung für Ihre Gewinnermittlung")
    
Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("SaKoMo")
Set RpItm = RpRec.AddItem("Saldenliste der Sachkonten")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt die Summen der einzelnen Sachkonten für den eingestellten Auswertungszeitraum. Nutzen Sie diese Auswertung für eine Analyse der jeweiligen Sachkonten")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("SaGeMo")
Set RpItm = RpRec.AddItem("Saldenliste der Geldkonten")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt die Summen der einzelnen Geldkonten für den Eingestellten Auswertungszeitraum. Nutzen Sie diese Auswertung für eine Analyse des Geldkontos")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("UmStMo")
Set RpItm = RpRec.AddItem("Umsatzsteuerübersicht")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Auswertung der in den jeweiligen Buchungen angefallenen Umsatzsteuer für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("ReUmsa")
Set RpItm = RpRec.AddItem("Rechnungsumsatzauswertung")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt die Umsätze der einzelnen Monate für den eingestellten Auswertungszeitraum auf Basis der gestellten Rechnungen")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("KomAus")
Set RpItm = RpRec.AddItem("Kompaktauswertung")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Kreuztabellenauswertung der Buchungskonten summiert für die einzelnen Monate des ausgewählten Jahres")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("PatUms")
Set RpItm = RpRec.AddItem("Patientenumsatzliste")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Liste der Patienten, die an einem bestimmten Tag in der Praxis waren und dessen erwirtschafteter Umsatz für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("TagUms")
Set RpItm = RpRec.AddItem("Tagesumsatzliste")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Liste der erwirtschafteten Umsätze an einem bestimmten Tag für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("LabUms")
Set RpItm = RpRec.AddItem("Steuermixumsatzliste")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Liste der erwirtschafteten Umsätze getrännt nach zugewiesener UmSt. an einem bestimmten Tag für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("TagSum")
Set RpItm = RpRec.AddItem("Tagessummen")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Liste der summierten Umsätze an einem bestimmten Tag für den eingestellten Auswertungszeitraum")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("StaDia")
Set RpItm = RpRec.AddItem("Diagnosestatistik")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Übersicht aller erstellter ICD-10 Diagnosen und deren Häufigkeit")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("StaGeb")
Set RpItm = RpRec.AddItem("Gebührenstatistik")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Übersicht aller verwendeten Gebührenpositionen und deren Häufigkeit")

Set RpRec = RpRcs.Add()
Set RpItm = RpRec.AddItem("RecSum")
Set RpItm = RpRec.AddItem("Rechnungssummen")
RpItm.Icon = IC16_Printer_Ink
Set RpItm = RpRec.AddItem("Zeigt eine Liste der summierten Umsätze an einem bestimmten Rechnungstag für den eingestellten Auswertungszeitraum")

With CmMan
    For AktZa = 1 To UBound(GlThe)
        .AddItem GlThe(AktZa, 13)
        .ItemData(AktZa - 1) = GlThe(AktZa, 0)
    Next AktZa
    .AddItem "für alle Mandanten"
    .ItemData(AktZa - 1) = 0
    .ListIndex = AktZa - 1
    .Enabled = True
    If GlRst = True Then 'Mandantenbezogene Datenbegrenzung
        For AktZa = 1 To UBound(GlMan)
            If ManNr = GlMan(AktZa, 2) Then 'Mandantennummer
                .ListIndex = AktZa - 1
                .Enabled = False
            End If
        Next AktZa
    End If
End With

With CmMit
    For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
        .AddItem GlMiK(AktZa, 1)
        .ItemData(AktZa - 1) = GlMiK(AktZa, 2)
    Next AktZa
    .AddItem "für alle Mitarbeiter"
    .ItemData(AktZa - 1) = 0
    .ListIndex = AktZa - 1
    .Enabled = GlMiV
End With

With TxDat
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Date, "dd.mm.yyyy")
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
OpJah.BackColor = GlBak
OpWoc.BackColor = GlBak
OpMon.BackColor = GlBak
OpQua.BackColor = GlBak
OpJah.BackColor = GlBak
OpDat.BackColor = GlBak
OpAlK.BackColor = GlBak
OpEiK.BackColor = GlBak
OpAlG.BackColor = GlBak
OpEiG.BackColor = GlBak

RetWe = SendMessage(CmMon.hwnd, CB_SETCURSEL, AkMon - 1, ByVal 0&)
RetWe = SendMessage(CmQua.hwnd, CB_SETCURSEL, AkQua - 1, ByVal 0&)
RetWe = SendMessage(CmWoc.hwnd, CB_SETCURSEL, AkWoc - 1, ByVal 0&)

RpCon.Populate

FText

Set RpCol = Nothing
Set RpRcs = Nothing
Set RpCls = Nothing
Set RpCon = Nothing
Set ImMan = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FText()
On Error GoTo InErr

Dim LiKey As String
Dim LiKom As String
Dim WocWa As Boolean
Dim QuaWa As Boolean
Dim MonWa As Boolean
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set OpWoc = Me.optWoche
Set OpMon = Me.optMonat
Set OpQua = Me.optQuart
Set OpJah = Me.optJahre
Set OpDat = Me.optDatum
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set CmWoc = Me.cmbWoche
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQuart
Set CmJah = Me.cmbJahre
Set Labe1 = Me.lblLab04
Set TxDat = Me.txtDatum
Set RpCon = Me.repCont1
Set RpCls = RpCon.Columns
Set RpRcs = RpCon.Records
Set RpSel = RpCon.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        LiKey = Trim$(RpRow.Record(0).Value)
        LiKom = Trim$(RpRow.Record(2).Value)
    End If
End If

Labe1.Caption = LiKom

MitWa = True
QuaWa = True
MonWa = True
ZeiWa = True
WocWa = True

Select Case LiKey
Case "BuJour": ZeiWa = True
               MitWa = False
Case "BuKass": ZeiWa = True
               MitWa = False
Case "EiUbMo": ZeiWa = False
               WocWa = False
               MitWa = False
Case "EiUbGr": ZeiWa = False
               QuaWa = False
               MonWa = False
               WocWa = False
               MitWa = False
Case "SaKoMo": ZeiWa = False
               WocWa = False
               MitWa = False
Case "SaGeMo": ZeiWa = False
               WocWa = False
               MitWa = False
Case "UmStMo": ZeiWa = False
               WocWa = False
               MitWa = False
Case "KomAus": ZeiWa = False
               WocWa = False
               QuaWa = False
               MonWa = False
               MitWa = False
Case "KonAus": ZeiWa = True
               MitWa = False
Case "KonGru": ZeiWa = True
               MitWa = False
Case "KonGel": ZeiWa = True
               MitWa = False
Case "ReUmsa": ZeiWa = False
               MitWa = False
Case "PatUms": ZeiWa = True
               MitWa = False
Case "TagUms": ZeiWa = True
               QuaWa = False
               MitWa = True
Case "LabUms": ZeiWa = True
               QuaWa = False
               MitWa = True
Case "TagSum": ZeiWa = True
               QuaWa = False
               MitWa = True
Case "RecSum": ZeiWa = True
               WocWa = False
               MitWa = True
Case "StaDia": ZeiWa = False
               WocWa = False
               QuaWa = False
               MonWa = False
               MitWa = False
Case "StaGeb": ZeiWa = False
               WocWa = False
               MitWa = False
End Select

CmMit.Enabled = MitWa
OpQua.Enabled = QuaWa
CmQua.Enabled = QuaWa
OpMon.Enabled = MonWa
CmMon.Enabled = MonWa
OpDat.Enabled = ZeiWa
TxDat.Enabled = ZeiWa
PuBu1.Enabled = ZeiWa
OpWoc.Enabled = WocWa
CmWoc.Enabled = WocWa

Set RpRcs = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FText " & Err.Number
Resume Next

End Sub
Private Function FStar() As String
On Error GoTo InErr

Dim BisDa As Date
Dim DaSta As Date
Dim IdxNr As Long
Dim ManNr As Long
Dim MitNr As Long
Dim KtoNr As Long
Dim GldNr As Long
Dim LiKey As String
Dim Krit1 As String
Dim Krit2 As String
Dim Krit3 As String
Dim Datu1 As String
Dim AkWoc As Integer
Dim AkMon As Integer
Dim AkJah As Integer
Dim AkQua As Integer
Dim IdBnk As Integer
Dim Mld1, Tit1 As String
Dim CmBrs As XtremeCommandBars.CommandBars
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim CmGlk As XtremeCommandBars.CommandBarComboBox

Set TxDum = Me.txtDummy
Set OpWoc = Me.optWoche
Set OpMon = Me.optMonat
Set OpQua = Me.optQuart
Set OpJah = Me.optJahre
Set OpDat = Me.optDatum
Set OpAlK = Me.optAllKo
Set OpEiK = Me.optEinKo
Set OpAlG = Me.optAllGe
Set OpEiG = Me.optEinGe
Set CmWoc = Me.cmbWoche
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQuart
Set CmJah = Me.cmbJahre
Set CmEiK = Me.cmbKonto
Set CmEiG = Me.cmbGegen
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set TxDat = Me.txtDatum
Set RpCon = Me.repCont1
Set RpCls = RpCon.Columns
Set RpRcs = RpCon.Records
Set RpSel = RpCon.SelectedRows

Set CmBrs = frmMain.comBar01

Set CmGlk = CmBrs.FindControl(CmGlk, SY_SuBuh, , True)

IdBnk = CmGlk.ItemData(CmGlk.ListIndex)

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        LiKey = Trim$(RpRow.Record(0).Value)
    End If
End If

If IsDate(TxDat.Text) Then
    DaSta = CDate(TxDat.Text)
Else
    DaSta = Date
End If

AkJah = CInt(CmJah.Text)
AkWoc = CmWoc.ItemData(CmWoc.ListIndex)
AkMon = CmMon.ItemData(CmMon.ListIndex)
AkQua = CmQua.ItemData(CmQua.ListIndex)
KtoNr = CmEiK.ItemData(CmEiK.ListIndex)
GldNr = CmEiG.ItemData(CmEiG.ListIndex)
ManNr = CmMan.ItemData(CmMan.ListIndex)
MitNr = CmMit.ItemData(CmMit.ListIndex)

Datu1 = DatePart("m", DaSta) & "/" & DatePart("d", DaSta) & "/" & DatePart("yyyy", DaSta)
Mld1 = "Sie haben keinen Auswertungszeitraum gewählt"
Tit1 = "Bericht drucken"

If GlMaV = True Then 'Mandanten vorhanden
    If ManNr > 0 Then
        If GlTyp < 2 Then
            Krit3 = " AND (IDT = " & ManNr & ")"
        Else
            Krit3 = " AND ([IDT] = " & ManNr & ")"
        End If
    End If
End If

If MitWa = True Then
    If GlMiV = True Then
        If MitNr > 0 Then
            If GlTyp < 2 Then
                If Krit3 <> vbNullString Then
                    Krit3 = Krit3 & " AND (IDM = " & MitNr & ")"
                Else
                    Krit3 = " AND (IDM = " & MitNr & ")"
                End If
            Else
                If Krit3 <> vbNullString Then
                    Krit3 = Krit3 & " AND ([IDM] = " & MitNr & ")"
                Else
                    Krit3 = " AND ([IDM] = " & MitNr & ")"
                End If
            End If
        End If
    End If
End If

If TxDum.Text = "KonAus" Then
    If OpEiK.Value = True Then
        If GlTyp < 2 Then
            Krit2 = " AND (Konto = " & KtoNr & ")"
        Else
            Krit2 = " AND ([Konto] = " & KtoNr & ")"
        End If
    ElseIf OpEiG.Value = True Then
        If GlTyp < 2 Then
            Krit2 = " AND (IDB = " & GldNr & ")"
        Else
            Krit2 = " AND ([IDB] = " & GldNr & ")"
        End If
    End If
End If

If TxDum.Text = "KonGru" Then
    If OpEiK.Value = True Then
        Krit2 = " AND ([Konto]=" & KtoNr & ")"
    End If
End If

If TxDum.Text = "KonGel" Then
    If OpEiG.Value = True Then
        If GlTyp < 2 Then
            Krit2 = " AND (IDB = " & GldNr & ")"
        Else
            Krit2 = " AND ([IDB] = " & GldNr & ")"
        End If
    End If
End If

If LiKey = "BuJour" Then
    If IdBnk > 0 Then
        If GlTyp < 2 Then
            Krit2 = " AND (IDB = " & IdBnk & ")"
        Else
            Krit2 = " AND ([IDB] = " & IdBnk & ")"
        End If
    End If
End If

If OpWoc.Value = True Then 'Wochenauswertung

    If GlTyp < 2 Then
        Krit1 = "(((DATEPART(wk, Datum)) = " & AkWoc & ") AND ((YEAR(Datum)) = " & AkJah & "))"
    Else
        Krit1 = "(((Year([Datum])) = " & AkJah & ") AND ((DatePart(" & Chr$(34) & "ww" & Chr$(34) & ",[Datum])) = " & AkWoc & "))"
    End If
    TxDum.Text = CmWoc.Text & " / " & CmJah.Text
    BisDa = DateAdd("d", ((AkWoc * 7) - 7), "01.01." & AkJah)

ElseIf OpMon.Value = True Then 'Monatsauswertung

    If GlTyp < 2 Then
        If ZeiWa = True Then
            Krit1 = "(((MONTH(Datum)) = " & AkMon & ") AND ((YEAR(Datum)) = " & AkJah & "))"
        Else
            Krit1 = "((Jahr = " & AkJah & ") AND (Monat = " & AkMon & "))"
        End If
    Else
        If ZeiWa = True Then
            Krit1 = "(((Month([Datum])) = " & AkMon & ") AND ((Year([Datum])) = " & AkJah & "))"
        Else
            Krit1 = "(([Jahr] = " & AkJah & ") AND ([Monat] = " & AkMon & "))"
        End If
    End If
    TxDum.Text = CmMon.Text & " / " & CmJah.Text
    If AkMon = 1 Then
        BisDa = CDate("01." & AkMon & "." & AkJah)
    Else
        BisDa = CDate("01." & AkMon & "." & AkJah) - 1
    End If
    
ElseIf OpQua.Value = True Then 'Quartalsauswertung

    If ZeiWa = True Then
        If GlTyp < 2 Then
            Select Case AkQua
            Case 1: Krit1 = "((Datum >= '01.01." & AkJah & "') AND (Datum <= '31.03." & AkJah & "'))"
            Case 2: Krit1 = "((Datum >= '01.04." & AkJah & "') AND (Datum <= '30.06." & AkJah & "'))"
            Case 3: Krit1 = "((Datum >= '01.07." & AkJah & "') AND (Datum <= '30.09." & AkJah & "'))"
            Case 4: Krit1 = "((Datum >= '01.10." & AkJah & "') AND (Datum <= '31.12." & AkJah & "'))"
            End Select
        Else
            Select Case AkQua
            Case 1: Krit1 = "(([Datum] Between #01/01/" & AkJah & "# AND #03/31/" & AkJah & "#))"
            Case 2: Krit1 = "(([Datum] Between #04/01/" & AkJah & "# AND #06/30/" & AkJah & "#))"
            Case 3: Krit1 = "(([Datum] Between #07/01/" & AkJah & "# AND #09/30/" & AkJah & "#))"
            Case 4: Krit1 = "(([Datum] Between #10/01/" & AkJah & "# AND #12/31/" & AkJah & "#))"
            End Select
        End If
    Else
        If GlTyp < 2 Then
            Krit1 = "((Jahr = " & AkJah & ") AND (Quartal = " & AkQua & "))"
        Else
            Krit1 = "(([Jahr] = " & AkJah & ") AND ([Quartal] = " & AkQua & "))"
        End If
    End If
    TxDum.Text = CmQua.Text & " / " & CmJah.Text
    Select Case AkQua
    Case 1: BisDa = CDate("01.01." & AkJah)
    Case 2: BisDa = CDate("01.04." & AkJah) - 1
    Case 3: BisDa = CDate("01.07." & AkJah) - 1
    Case 4: BisDa = CDate("01.10." & AkJah) - 1
    End Select
    
ElseIf OpJah.Value = True Then 'Jahresauswertung

    If ZeiWa = True Then
        If GlTyp < 2 Then
            Krit1 = "((YEAR(Datum) = " & AkJah & "))"
        Else
            Krit1 = "((Year([Datum]) = " & AkJah & "))"
        End If
    Else
        If GlTyp < 2 Then
            Krit1 = "(Jahr = " & AkJah & ")"
        Else
            Krit1 = "([Jahr] = " & AkJah & ")"
        End If
    End If
    TxDum.Text = "Jahr: " & CmJah.Text
    BisDa = CDate("01.01." & AkJah)
    
ElseIf OpDat.Value = True Then 'Zeitraumsauswertung

    If GlTyp < 2 Then
        Krit1 = "(Datum = '" & DaSta & "')"
    Else
        Krit1 = "([Datum] = #" & Datu1 & "#)"
    End If
    TxDum.Text = DaSta
    BisDa = DaSta
    
Else

    WindowMess Mld1, Dial2, Tit1, Me.hwnd
    
End If

If ManNr > 0 Then
    AnSal = S_BuSa(AkJah, BisDa, ManNr)
Else
    AnSal = S_BuSa(AkJah, BisDa)
End If

If Krit1 <> vbNullString Then
    If Krit2 <> vbNullString Then
        If Krit3 <> vbNullString Then
            FStar = Krit1 & Krit2 & Krit3
        Else
            FStar = Krit1 & Krit2
        End If
    ElseIf Krit3 <> vbNullString Then
        FStar = Krit1 & Krit3
    Else
        FStar = Krit1
    End If
Else
    FStar = vbNullString
End If

Set RpRcs = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStar " & Err.Number
Resume Next

End Function
Private Sub FWeit()
On Error GoTo InErr

Dim ThIdx As Long
Dim LiKey As String
Dim LiKom As String
Dim Krit1 As String
Dim ForNa As String
Dim KopTe As String
Dim ZeRau As Integer
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set OpWoc = Me.optWoche
Set OpMon = Me.optMonat
Set OpQua = Me.optQuart
Set OpJah = Me.optJahre
Set OpDat = Me.optDatum
Set CmMan = Me.cmbBehan
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set TxDum = Me.txtDummy
Set OpAlK = Me.optAllKo
Set OpEiK = Me.optEinKo
Set OpAlG = Me.optAllGe
Set OpEiG = Me.optEinGe
Set CmEiK = Me.cmbKonto
Set CmEiG = Me.cmbGegen
Set RpCon = Me.repCont1
Set RpCls = RpCon.Columns
Set RpRcs = RpCon.Records
Set RpSel = RpCon.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        LiKey = Trim$(RpRow.Record(0).Value)
        LiKom = Trim$(RpRow.Record(2).Value)
    End If
End If

If OpWoc.Value = True Then
    ZeRau = 1
ElseIf OpMon.Value = True Then
    ZeRau = 1
ElseIf OpJah.Value = True Then
    ZeRau = 2
ElseIf OpQua.Value = True Then
    ZeRau = 3
Else
    ZeRau = 4
End If

ForNa = LiKey

If Rahm1.Visible = True Then

    Rahm1.Visible = False
    Select Case ForNa
    Case "KonAus":
            Rahm3.Visible = True
    Case "KonGru":
            Rahm3.Visible = True
            OpEiG.Enabled = False
            OpAlG.Enabled = False
            OpEiG.Enabled = False
            CmEiG.Enabled = False
            OpAlK.Value = True
    Case "KonGel":
            Rahm3.Visible = True
            CmEiK.Enabled = False
            OpAlK.Enabled = False
            OpEiK.Enabled = False
            OpAlG.Value = True
    Case Else:
            Rahm2.Visible = True
    End Select
    
ElseIf Rahm2.Visible = True Then

    Krit1 = FStar
    KopTe = TxDum.Text

    If CmMan.ItemData(CmMan.ListIndex) = 0 Then
        ThIdx = 0
    Else
        ThIdx = CmMan.ListIndex + 1
    End If

    If Krit1 <> vbNullString Then
        With GlBuD 'Buchhaltungsdruck
            .Krit1 = Krit1
            .AnSal = AnSal
            .ZeRau = ZeRau
            If ThIdx > 0 Then
                .ManNr = ThIdx
            End If
        End With
        SDruck ForNa, True, KopTe, False
    Else
        Unload Me
    End If
    
ElseIf Rahm3.Visible = True Then

    TxDum.Text = ForNa
    Rahm3.Visible = False
    Rahm2.Visible = True

End If

Set RpRcs = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    FKale
End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50691)
TeMai = IniGetOpt("Hilfe", 50692)
TeInh = IniGetOpt("Hilfe", 50693)
TeFus = IniGetOpt("Hilfe", 50694)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub
Private Sub cmbJahre_Click()
    Me.optJahre.Value = True
End Sub
Private Sub cmbJahre_DropDown()
    Me.optJahre.Value = True
End Sub
Private Sub cmbMonat_Click()
    Me.optMonat.Value = True
End Sub
Private Sub cmbMonat_DropDown()
    Me.optMonat.Value = True
End Sub

Private Sub cmbQuart_Click()
    Me.optQuart.Value = True
End Sub

Private Sub cmbQuart_DropDown()
    Me.optQuart.Value = True
End Sub

Private Sub cmbWoche_Click()
    Me.optWoche.Value = True
End Sub

Private Sub cmbWoche_DropDown()
    Me.optWoche.Value = True
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

FrLad = True
FKonf
AFont Me
S_ZeKo
FrLad = False
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmZeitraum = Nothing
End Sub

Private Sub repCont1_KeyUp(KeyCode As Integer, Shift As Integer)
    FText
End Sub
Private Sub repCont1_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    FText
End Sub

Private Sub txtDatum_GotFocus()
    Me.txtDatum.SelStart = 0
    Me.txtDatum.SelLength = Len(Me.txtDatum.Text)
End Sub
Private Sub txtDatum_LostFocus()
    FDaKo
End Sub
