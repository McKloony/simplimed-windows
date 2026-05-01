VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReExpo 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnungsexport"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3900
      Left            =   300
      TabIndex        =   5
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   6879
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.RadioButton optAnder 
         Height          =   220
         Left            =   1600
         TabIndex        =   6
         Top             =   1700
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungen eines bestimmten Zeitraums"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optSelbe 
         Height          =   220
         Left            =   1600
         TabIndex        =   8
         Top             =   1300
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Nur die markierten Rechnungen exportieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   1600
         TabIndex        =   7
         Top             =   2800
         Width           =   2700
         _Version        =   1048579
         _ExtentX        =   4763
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   220
         Left            =   1600
         TabIndex        =   34
         Top             =   2560
         Width           =   855
         _Version        =   1048579
         _ExtentX        =   1508
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandant:"
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   "Mˆchten Sie nur die markierten Rechnungen exportieren oder alle Rechnungen eines bestimmten Zeitraums?"
         Height          =   400
         Left            =   900
         TabIndex        =   9
         Top             =   200
         Width           =   4500
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   6010
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3900
      Left            =   300
      TabIndex        =   10
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   6879
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkBlgEx 
         Height          =   240
         Left            =   1600
         TabIndex        =   15
         Top             =   3100
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Belege exportieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkExEml 
         Height          =   240
         Left            =   1600
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3500
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Emailversand"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbForma 
         Height          =   310
         Left            =   1600
         TabIndex        =   11
         Top             =   1550
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkReAbs 
         Height          =   220
         Left            =   1600
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2300
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungen verriegeln"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkReBez 
         Height          =   220
         Left            =   1600
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2700
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4410
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Offene Posten generieren"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblLab06 
         BackStyle       =   0  'Transparent
         Caption         =   "In welchem Format sollen die Rechnungen exportiert werden? Nicht alle Formate beinhalten auch die Belege."
         Height          =   600
         Left            =   900
         TabIndex        =   16
         Top             =   200
         Width           =   4500
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Undurchsichtig
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   6010
      End
      Begin VB.Label lblLab07 
         BackStyle       =   0  'Transparent
         Caption         =   "Exportformat :"
         Height          =   210
         Left            =   1600
         TabIndex        =   14
         Top             =   1300
         Width           =   1500
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   3900
      Left            =   300
      TabIndex        =   19
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   6879
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   350
         Left            =   3820
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3360
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   3820
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2860
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit4 
         Height          =   225
         Left            =   1500
         TabIndex        =   22
         Top             =   2900
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zeitraum"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit3 
         Height          =   225
         Left            =   1500
         TabIndex        =   23
         Top             =   2300
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Jahr"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit2 
         Height          =   225
         Left            =   1500
         TabIndex        =   24
         Top             =   1700
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Quartal"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optZeit1 
         Height          =   225
         Left            =   1500
         TabIndex        =   25
         Top             =   1100
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Monat"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbMonat 
         Height          =   315
         Left            =   2600
         TabIndex        =   26
         Top             =   1060
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
      Begin XtremeSuiteControls.ComboBox cmbQurta 
         Height          =   315
         Left            =   2600
         TabIndex        =   27
         Top             =   1660
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   2600
         TabIndex        =   28
         Top             =   2860
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   350
         Left            =   2600
         TabIndex        =   29
         Top             =   3360
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbJahre 
         Height          =   315
         Left            =   2600
         TabIndex        =   30
         Top             =   2260
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
      End
      Begin XtremeCalendarControl.DatePicker dtpDatu1 
         Height          =   405
         Left            =   4300
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   405
         _Version        =   1048579
         _ExtentX        =   706
         _ExtentY        =   706
         _StockProps     =   64
         Show3DBorder    =   2
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte legen Sie den Zeitraum fest, f¸r den Sie die Rechnungen exportieren mˆchten."
         Height          =   450
         Left            =   900
         TabIndex        =   33
         Top             =   200
         Width           =   3500
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Undurchsichtig
         BorderStyle     =   0  'Transparent
         Height          =   800
         Left            =   0
         Top             =   0
         Width           =   6010
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "bis"
         Height          =   195
         Left            =   1500
         TabIndex        =   32
         Top             =   3400
         Width           =   850
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   0
      Top             =   4000
      Width           =   6700
      _Version        =   1048579
      _ExtentX        =   11818
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   4700
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schlieﬂen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   3300
         TabIndex        =   2
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
      Begin XtremeSuiteControls.PushButton btnZuruk 
         Height          =   400
         Left            =   1900
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zur¸ck"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   600
         TabIndex        =   4
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
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   18
      Top             =   9000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   800
      Left            =   0
      Top             =   0
      Width           =   6700
   End
End
Attribute VB_Name = "frmReExpo"
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
Private TxDum As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private CmMon As XtremeSuiteControls.ComboBox
Private CmQua As XtremeSuiteControls.ComboBox
Private CmJah As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmFma As XtremeSuiteControls.ComboBox
Private ChEml As XtremeSuiteControls.CheckBox
Private ChBlg As XtremeSuiteControls.CheckBox
Private ChReA As XtremeSuiteControls.CheckBox
Private ChReB As XtremeSuiteControls.CheckBox
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private OpMon As XtremeSuiteControls.RadioButton
Private OpQua As XtremeSuiteControls.RadioButton
Private OpJah As XtremeSuiteControls.RadioButton
Private OpZei As XtremeSuiteControls.RadioButton
Private MoKal As XtremeCalendarControl.DatePicker
Private ImMan As XtremeCommandBars.ImageManager
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private RpRow As XtremeReportControl.ReportRow
Private RpSel As XtremeReportControl.ReportSelectedRows

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_SETCURSEL = &H14E
Private Const LB_SETCURSEL = &H186

Private ManNr As Long
Private Krite As String
Private KalWa As Integer
Private OptWe As Integer
Private GeKto As Integer
Private FoLad As Boolean

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
        TxDa1.Text = NeuDa
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
        TxDa2.Text = NeuDa
    End If
End Select

With MoKal
    .EnsureVisible NeuDa - 30
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
End With

If NeuDa > Date Then SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1
Set OpZei = Me.optZeit4

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    Select Case KalWa
    Case 1: TxDa1.Text = NeuDa
            TxDa2.Text = NeuDa
            TxDa1.SetFocus
    Case 2: TxDa2.Text = NeuDa
            TxDa2.SetFocus
    End Select
End If

OpZei.Value = True

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo LaErr

Dim RetWe As Long
Dim AktZa As Integer
Dim AktKo As Integer
Dim IdxZa As Integer
Dim AkMon As Integer
Dim AkQua As Integer
Dim BuJah As Integer
Dim AnzPo As Integer
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmReExpo
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set CmFma = FM.cmbForma
Set ChEml = FM.chkExEml
Set ChBlg = FM.chkBlgEx
Set ChReA = Me.chkReAbs
Set ChReB = Me.chkReBez
Set CmMan = FM.cmbManda
Set Opti1 = FM.optSelbe
Set Opti2 = FM.optAnder
Set OpMon = FM.optZeit1
Set OpQua = FM.optZeit2
Set OpJah = FM.optZeit3
Set OpZei = FM.optZeit4
Set CmMon = FM.cmbMonat
Set CmQua = FM.cmbQurta
Set CmJah = FM.cmbJahre
Set MoKal = FM.dtpDatu1
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnDatu2
Set PuBu3 = FM.btnZuruk
Set ImMan = frmMain.imgManag
Set RpCo3 = frmMain.repCont3
Set RpCo4 = frmMain.repCont4

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select
AnzPo = RpSel.Count

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
    .ToolTipText = "Markieren Sie bitte hier den gw¸nschten Rechnungstag"
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

With CmFma
    .AddItem "SMP Abrechnung (SMP)"
    .ItemData(.NewIndex) = 1
    .AddItem "PAD Abrechnung (PAD)"
    .ItemData(.NewIndex) = 2
    .AddItem "PADNext Abrechnung (XML)"
    .ItemData(.NewIndex) = 3
    .AddItem "Adobe Acrobat (PDF)"
    .ItemData(.NewIndex) = 4
    .AddItem "Rechnungsliste (CSV)"
    .ItemData(.NewIndex) = 5
    .AddItem "Rechnungsexportliste (PDF)"
    .ItemData(.NewIndex) = 6
    .AddItem "Rechnungsschnell¸bers. (PDF)"
    .ItemData(.NewIndex) = 7
    .AddItem "DATEV 6.0 Belegsatzdaten"
    .ItemData(.NewIndex) = 8
    .AddItem "DATEV 6.0 Belegarchivierung"
    .ItemData(.NewIndex) = 9
    .AddItem "X-Rechnung Dateien (XML)"
    .ItemData(.NewIndex) = 10
    .AddItem "ZUGFeRD Dateien (PDF)"
    .ItemData(.NewIndex) = 11
    .DropDownItemCount = 11
    .AutoComplete = False
    .ListIndex = 1
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

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa
CmMan.AddItem "f¸r alle Mandanten"
CmMan.ItemData(AktZa - 1) = 0
CmMan.ListIndex = AktZa - 1

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

If GlBlE = True Then 'DATEV Belegexport
    ChBlg.Value = xtpChecked
    ChEml.Enabled = False
End If
DoEvents

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date) - 1
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Format$(Day(Date), "00") & "." & Format$(Month(Date), "00") & "." & Year(Date) + 1
End With

If AnzPo > 1 Then
    Opti1.Value = True
Else
    Opti2.Value = True
End If

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

RetWe = SendMessage(CmMon.hwnd, CB_SETCURSEL, AkMon - 1, ByVal 0&)
RetWe = SendMessage(CmQua.hwnd, CB_SETCURSEL, AkQua - 1, ByVal 0&)

If CBool(IniGetVal("System", "ReExAb")) = True Then ChReA.Value = xtpChecked
If CBool(IniGetVal("System", "ReExBe")) = True Then ChReB.Value = xtpChecked

DoEvents
FTyp

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
OpMon.BackColor = GlBak
OpQua.BackColor = GlBak
OpJah.BackColor = GlBak
OpZei.BackColor = GlBak
Opti1.BackColor = GlBak
Opti2.BackColor = GlBak
ChEml.BackColor = GlBak
ChBlg.BackColor = GlBak
ChReA.BackColor = GlBak
ChReB.BackColor = GlBak

Set ImMan = Nothing
Set RpSel = Nothing
Set RpCo3 = Nothing
Set RpCo4 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub

Private Sub FKale()
On Error GoTo LaErr
'L‰ﬂt den Kalender aufklappen

Dim NeuDa As Date
Dim Datu1 As Date
Dim Datu2 As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set MoKal = Me.dtpDatu1
Set Rahm1 = Me.frmRahm1

Select Case KalWa
Case 1:
    If IsDate(TxDa1.Text) Then
        NeuDa = TxDa1.Text
    Else
        NeuDa = Date
    End If
Case 2:
    If IsDate(TxDa2.Text) Then
        NeuDa = TxDa2.Text
    Else
        NeuDa = Date
    End If
End Select

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    Select Case KalWa
    Case 1: .Top = TxDa1.Top + TxDa1.Height
            .Left = TxDa1.Left + Rahm1.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa1.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left + Rahm1.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    End Select
End With

Datu1 = TxDa1.Text
Datu2 = TxDa2.Text

If Datu2 < Datu1 Then TxDa1.Text = Datu2

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Function FStar() As String
On Error GoTo InErr

Dim DaSta As Date
Dim DaEnd As Date
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim Krit1 As String
Dim Datu1 As String
Dim Datu2 As String
Dim AkMon As Integer
Dim AkJha As Integer
Dim AkQua As Integer
Dim Mld1, Tit1 As String

Set OpMon = Me.optZeit1
Set OpQua = Me.optZeit2
Set OpJah = Me.optZeit3
Set OpZei = Me.optZeit4
Set CmMon = Me.cmbMonat
Set CmQua = Me.cmbQurta
Set CmJah = Me.cmbJahre
Set CmMan = Me.cmbManda
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDum = Me.txtDummy

If IsDate(TxDa1.Text) Then
    DaSta = TxDa1.Text
Else
    DaSta = Date
End If

If IsDate(TxDa2.Text) Then
    DaEnd = TxDa2.Text
Else
    DaEnd = Date
End If

AkJha = CInt(CmJah.Text)
AkMon = CmMon.ItemData(CmMon.ListIndex)
AkQua = CmQua.ItemData(CmQua.ListIndex)

ManNr = CmMan.ItemData(CmMan.ListIndex)

Datu1 = DatePart("m", DaSta) & "/" & DatePart("d", DaSta) & "/" & DatePart("yyyy", DaSta)
Datu2 = DatePart("m", DaEnd) & "/" & DatePart("d", DaEnd) & "/" & DatePart("yyyy", DaEnd)
Mld1 = "Sie haben keinen Auswertungszeitraum gew‰hlt"
Tit1 = "Rechnungs¸bersicht"

If OpMon.Value = True Then
    If GlTyp < 2 Then
        Krit1 = "(((MONTH(Datum))=" & AkMon & ") AND ((YEAR(Datum))=" & AkJha & "))"
    Else
        Krit1 = "(((Month([Datum]))=" & AkMon & ") AND ((Year([Datum]))=" & AkJha & "))"
    End If
    TxDum.Text = CmMon.Text & " / " & CmJah.Text
ElseIf OpQua.Value = True Then
    If GlTyp < 2 Then
        Select Case AkQua
        Case 1: Krit1 = "((Datum >= '01.01." & AkJha & "') AND (Datum <= '31.03." & AkJha & "'))"
        Case 2: Krit1 = "((Datum >= '01.04." & AkJha & "') AND (Datum <= '30.06." & AkJha & "'))"
        Case 3: Krit1 = "((Datum >= '01.07." & AkJha & "') AND (Datum <= '30.09." & AkJha & "'))"
        Case 4: Krit1 = "((Datum >= '01.10." & AkJha & "') AND (Datum <= '31.12." & AkJha & "'))"
        End Select
    Else
        Select Case AkQua
        Case 1: Krit1 = "(([Datum] Between #01/01/" & AkJha & "# AND #03/31/" & AkJha & "#))"
        Case 2: Krit1 = "(([Datum] Between #04/01/" & AkJha & "# AND #06/30/" & AkJha & "#))"
        Case 3: Krit1 = "(([Datum] Between #07/01/" & AkJha & "# AND #09/30/" & AkJha & "#))"
        Case 4: Krit1 = "(([Datum] Between #10/01/" & AkJha & "# AND #12/31/" & AkJha & "#))"
        End Select
    End If
    TxDum.Text = CmQua.Text & " / " & CmJah.Text
ElseIf OpJah.Value = True Then
    If GlTyp < 2 Then
        Krit1 = "((YEAR(Datum) = " & AkJha & "))"
    Else
        Krit1 = "((Year([Datum]) = " & AkJha & "))"
    End If
    TxDum.Text = "Jahr: " & CmJah.Text
ElseIf OpZei.Value = True Then
    Select Case GlTyp
    Case 0: Krit1 = "((Datum >= '" & DaSta & "') AND (Datum <= '" & DaEnd & "'))"
    Case 1: Krit1 = "((Datum >= '" & DaSta & "') AND (Datum <= '" & DaEnd & "'))"
    Case 2: Krit1 = "(([Datum] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
    Case 3: Krit1 = "(([Datum] Between #" & Datu1 & "# AND #" & Datu2 & "#))"
    End Select
    TxDum.Text = DaSta & " - " & DaEnd
Else
    WindowMess Mld1, Dial2, Tit1, Me.hwnd
End If

If GlMaV = True Then 'Mandanten vorhanden
    If ManNr > 0 Then
        If GlTyp < 2 Then
            Krit1 = Krit1 & " AND (IDP = " & ManNr & " )"
        Else
            Krit1 = Krit1 & " AND ([IDP] = " & ManNr & " )"
        End If
    End If
End If

If Krit1 <> vbNullString Then
    FStar = Krit1
Else
    FStar = vbNullString
End If

Exit Function

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FStar " & Err.Number
Resume Next

End Function
Private Sub FTyp()
On Error GoTo LdErr

Dim EmSen As Boolean
Dim ReAbs As Boolean
Dim BeExp As Boolean
Dim LiIdx As Integer

Set CmFma = Me.cmbForma
Set ChReA = Me.chkReAbs
Set ChReB = Me.chkReBez
Set ChBlg = Me.chkBlgEx
Set ChEml = Me.chkExEml

LiIdx = CmFma.ListIndex

Select Case LiIdx
Case 0: BeExp = False 'SMP Abrechnung
        ReAbs = True
        EmSen = True
Case 1: BeExp = False 'PAD Abrechnung
        ReAbs = True
        EmSen = True
Case 2: BeExp = False 'PADNext Abrechnung
        ReAbs = True
        EmSen = True
Case 3: BeExp = False 'Adobe Acrobat (PDF)
        ReAbs = True
        EmSen = False
Case 4: BeExp = False 'Rechnungsliste (CSV)
        ReAbs = True
        EmSen = True
Case 5: BeExp = False 'Rechnungsexportliste(PDF)
        ReAbs = True
        EmSen = True
Case 6: BeExp = False 'Rechnungsschnell¸bers. (PDF)
        ReAbs = True
        EmSen = True
Case 7: BeExp = True 'DATEV 6.0 Belegsatzdaten"
        ReAbs = False
        EmSen = True
Case 8: BeExp = True 'DATEV 6.0 Belegarchivierung"
        ReAbs = False
        EmSen = True
Case 9: BeExp = False 'X-Rechnung Dateien (XML)
        ReAbs = True
        EmSen = True
Case 10: BeExp = False 'ZUGFeRD Dateien (PDF)"
         ReAbs = True
         EmSen = True
End Select

FoLad = True
DoEvents

ChEml.Enabled = EmSen

If BeExp = False Then
    ChBlg.Enabled = False
    ChBlg.Value = xtpUnchecked
Else
    ChBlg.Enabled = True
End If

If ReAbs = True Then
    ChReA.Enabled = True
    ChReB.Enabled = True
Else
    ChReA.Enabled = False
    ChReB.Enabled = False
End If
ChReA.Value = xtpUnchecked
ChReB.Value = xtpUnchecked

If LiIdx = 3 Then
    ChBlg.Value = xtpGrayed
ElseIf LiIdx = 10 Then
    ChBlg.Value = xtpGrayed
Else
    ChBlg.Value = xtpUnchecked
End If

FoLad = False

Exit Sub

LdErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FTyp " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo InErr

Dim RowNr As Long
Dim AbrNr As Long
Dim SQL1 As String
Dim SQL2 As String
Dim SQL3 As String
Dim ForNa As String
Dim ExFmt As String
Dim KopTe As String
Dim AnzPo As Integer
Dim ZeRau As Integer
Dim LiIdx As Integer
Dim EmlVe As Integer
Dim AbgRe As Integer
Dim ExVer As Boolean
Dim ExKom As Boolean
Dim BelEx As Boolean
Dim ReAbs As Boolean
Dim ZahZi As Boolean
Dim RetWe As Boolean
Dim RpCo3 As XtremeReportControl.ReportControl
Dim RpCo4 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set RpCo3 = FM.repCont3
Set RpCo4 = FM.repCont4

Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set CmMan = Me.cmbManda
Set ChReA = Me.chkReAbs
Set ChReB = Me.chkReBez
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set OpMon = Me.optZeit1
Set OpQua = Me.optZeit2
Set OpJah = Me.optZeit3
Set OpZei = Me.optZeit4
Set ChBlg = Me.chkBlgEx
Set CmFma = Me.cmbForma
Set ChEml = Me.chkExEml
Set PuBu3 = Me.btnZuruk

Select Case GlBut
Case RibTab_Abrechnung:
    Set RpCls = RpCo3.Columns
    Set RpSel = RpCo3.SelectedRows
Case RibTab_Rechnungen:
    Set RpCls = RpCo4.Columns
    Set RpSel = RpCo4.SelectedRows
End Select
AnzPo = RpSel.Count

Set RpSel = RpCo4.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    RowNr = RpRow.Index
End If

Set RpSel = RpCo3.SelectedRows
If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    AbrNr = RpRow.Index
End If

ForNa = GlFrm(2, 0)
GlDru.ForNa = ForNa

LiIdx = CmFma.ListIndex
If ChEml.Value = xtpChecked Then EmlVe = 1
If ChBlg.Value = xtpChecked Then BelEx = True
If ChReA.Value = xtpChecked Then ReAbs = True
If ChReB.Value = xtpChecked Then ZahZi = True

Select Case LiIdx
Case 0: ExFmt = "SMP" 'SMP SimpliMed (SMP)
Case 1: ExFmt = "PAD" 'PAD Abrechnung (PAD)
Case 2: ExFmt = "MAD" 'PADNext Abrechnung (XML)
Case 3: ExFmt = "PDF" 'Adobe Acrobat (PDF)
Case 4: ExFmt = "CSV" 'Rechnungsliste (CSV)
Case 5: ExFmt = "LIS" 'Rechnungsexportliste (PDF)
Case 6: ExFmt = "LIS" 'Rechnungsschnell¸bers. (PDF)
Case 7: ExFmt = "DAV" 'DATEV 6.0 Belegsatzdaten"
Case 8: ExFmt = "DAV" 'DATEV 6.0 Belegarchivierung"
Case 9: ExFmt = "XML" 'X-Rechnung Dateien (XML)
Case 10: ExFmt = "ZUG" 'ZUGFeRD Dateien (PDF)"
End Select

If Rahm1.Visible = True Then

    If Opti1.Value = True Then
        OptWe = 1
    ElseIf Opti2.Value = True Then
        OptWe = 2
    End If
    If OptWe = 1 Then
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
    ElseIf OptWe = 2 Then
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = True
    End If
    PuBu3.Enabled = True
    
ElseIf Rahm2.Visible = True Then

    Screen.MousePointer = vbHourglass
    DoEvents
    
    Unload Me
    DoEvents

    If ExFmt = "LIS" Then

        Select Case LiIdx
        Case 5: S_BeDat "ReList", EmlVe, False, 0, False, True
        Case 6: S_BeDat "ResUbe", EmlVe, False, 0, False, True
        End Select

    ElseIf ExFmt = "CSV" Then

        S_ReExC EmlVe, ReAbs
        DoEvents
        RetWe = S_ReAn(True, ReAbs)
        If RetWe = True Then
            Exit Sub
        End If
        DoEvents
        If ZahZi = True Then
            AbgRe = S_OPAn(Date)
        End If
        DoEvents
        Select Case GlBut
        Case RibTab_Abrechnung:
                SUpAb AbrNr
                SUpRe RowNr
        Case RibTab_Rechnungen:
                SUpRe RowNr
        End Select
    
    ElseIf ExFmt = "SMP" Then

        If GlVar = "PS3" Then
            Unload Me
            DoEvents
            RetWe = S_ReExT(EmlVe, ReAbs)
            DoEvents
            If RetWe = True Then
                RetWe = S_ReAn(False, ReAbs, Date)
                If RetWe = True Then
                    Exit Sub
                End If
                DoEvents
                If ZahZi = True Then
                    AbgRe = S_OPAn(Date)
                End If
                DoEvents
                Select Case GlBut
                Case RibTab_Abrechnung:
                        SUpAb AbrNr
                        SUpRe RowNr
                Case RibTab_Rechnungen:
                        SUpRe RowNr
                End Select
            End If
        Else
            SPopu "SMP Export", "Die SMP Schnittstelle wurde noch nicht freigeschaltet bzw. lizenziert!", IC48_Forbidden
        End If
    
    ElseIf ExFmt = "PAD" Then

        If GlVar = "PS3" Then
            Unload Me
            DoEvents
            RetWe = S_ReExP()
            DoEvents
            If RetWe = True Then
                RetWe = S_ReAn(False, ReAbs, Date)
                If RetWe = True Then
                    Exit Sub
                End If
                DoEvents
                If ZahZi = True Then
                    AbgRe = S_OPAn(Date)
                End If
                DoEvents
                Select Case GlBut
                Case RibTab_Abrechnung:
                        SUpAb AbrNr
                        SUpRe RowNr
                Case RibTab_Rechnungen:
                        SUpRe RowNr
                End Select
            End If
        Else
            SPopu "PAD Export", "Die PAD Schnittstelle wurde noch nicht freigeschaltet bzw. lizenziert!", IC48_Forbidden
        End If
    
    ElseIf ExFmt = "MAD" Then

        If GlVar = "PS3" Then
            Unload Me
            DoEvents
            RetWe = S_ReExN(True)
            DoEvents
            If RetWe = True Then
                RetWe = S_ReAn(False, ReAbs, Date)
                If RetWe = True Then
                    Exit Sub
                End If
                DoEvents
                If ZahZi = True Then
                    AbgRe = S_OPAn(Date)
                End If
                DoEvents
                Select Case GlBut
                Case RibTab_Abrechnung:
                        SUpAb AbrNr
                        SUpRe RowNr
                Case RibTab_Rechnungen:
                        SUpRe RowNr
                End Select
            End If
        Else
            SPopu "PADNext Export", "Die PADNext Schnittstelle wurde noch nicht freigeschaltet bzw. lizenziert!", IC48_Forbidden
        End If
        
    ElseIf ExFmt = "PDF" Then

        SExpo ForNa, ExFmt, EmlVe, False, 0, True
        DoEvents
        If ExFmt = "PDF" Then
            RetWe = S_ReAn(True, ReAbs) 'Passt das Rechnungsdatum an und legt Rechnungsnummer an
            If RetWe = True Then
                Exit Sub
            End If
            DoEvents
            If ZahZi = True Then
                AbgRe = S_OPAn(Date)
            End If
            DoEvents
            Select Case GlBut
            Case RibTab_Abrechnung:
                    SUpAb AbrNr
                    SUpRe RowNr
            Case RibTab_Rechnungen:
                    SUpRe RowNr
            End Select
        End If

    ElseIf ExFmt = "XML" Then

        If GlVar = "PS3" Then
            Unload Me
            DoEvents
            If AnzPo > 1 Then
                S_ReExX True
            Else
                S_ReExX True, EmlVe
            End If
            DoEvents
            RetWe = S_ReAn(False, ReAbs, Date)
            If RetWe = True Then
                Exit Sub
            End If
            DoEvents
            If ZahZi = True Then
                AbgRe = S_OPAn(Date)
            End If
            DoEvents
            Select Case GlBut
            Case RibTab_Abrechnung:
                    SUpAb AbrNr
                    SUpRe RowNr
            Case RibTab_Rechnungen:
                    SUpRe RowNr
            End Select
        Else
            SPopu "X-Rechnung Export", "Die X-Rechnung Schnittstelle wurde noch nicht freigeschaltet bzw. lizenziert!", IC48_Forbidden
        End If
        
    ElseIf ExFmt = "ZUG" Then

        If GlVar = "PS3" Then
            Unload Me
            DoEvents
            SExpo ForNa, "PDF", EmlVe, False, 0, True
            DoEvents
            RetWe = S_ReAn(True, ReAbs) 'Passt das Rechnungsdatum an und legt Rechnungsnummer an
            If RetWe = True Then
                Exit Sub
            End If
            DoEvents
            If ZahZi = True Then
                AbgRe = S_OPAn(Date)
            End If
            DoEvents
            Select Case GlBut
            Case RibTab_Abrechnung:
                    SUpAb AbrNr
                    SUpRe RowNr
            Case RibTab_Rechnungen:
                    SUpRe RowNr
            End Select
        Else
            SPopu "ZUGFeRD Export", "Die ZUGFeRD Schnittstelle wurde noch nicht freigeschaltet bzw. lizenziert!", IC48_Forbidden
        End If
    
    ElseIf ExFmt = "DAV" Then

        If OptWe = 1 Then
            Select Case LiIdx
            Case 7: DATEV_Expor "B", EmlVe, BelEx
            Case 8: DATEV_Expor "A", EmlVe, BelEx
            End Select
        ElseIf OptWe = 2 Then
            Select Case LiIdx
            Case 7: DATEV_BuEx "B", EmlVe, Krite, BelEx
            Case 8: DATEV_BuEx "A", EmlVe, Krite, BelEx
            End Select
        End If

    End If

    DoEvents
    Screen.MousePointer = vbNormal
    
ElseIf Rahm3.Visible = True Then

    If OpMon.Value = True Then
        ZeRau = 1
    ElseIf OpJah.Value = True Then
        ZeRau = 2
    ElseIf OpQua.Value = True Then
        ZeRau = 3
    Else
        ZeRau = 4
    End If

    SQL1 = FStar

    If SQL1 <> vbNullString Then
        Krite = SQL1
    End If
    
    If Right$(Krite, 1) = ")" Then
        If SQL2 <> vbNullString Then
            Krite = Krite & " AND " & SQL2
        End If
    Else
        Krite = Krite & SQL2
    End If
    
    If Right$(Krite, 1) = ")" Then
        If SQL3 <> vbNullString Then
            Krite = Krite & " AND " & SQL3
        End If
    Else
        Krite = Krite & SQL3
    End If

    Rahm1.Visible = False
    Rahm2.Visible = True
    Rahm3.Visible = False
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub
Private Sub btnDatu2_Click()
    KalWa = 2
    FKale
End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50131)
TeMai = IniGetOpt("Hilfe", 50132)
TeInh = IniGetOpt("Hilfe", 50133)
TeFus = IniGetOpt("Hilfe", 50134)

'SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub
Private Sub btnZuruk_Click()
On Error Resume Next

Set PuBu3 = Me.btnZuruk
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3

Rahm1.Visible = True
Rahm2.Visible = False
Rahm3.Visible = False
PuBu3.Enabled = False

End Sub
Private Sub chkBlgEx_Click()
On Error Resume Next

Set ChBlg = Me.chkBlgEx
Set ChEml = Me.chkExEml

If FoLad = False Then
    If ChBlg.Value = xtpChecked Then
        GlBlE = True 'DATEV Belegexport
        ChEml.Enabled = False
        ChEml.Value = xtpUnchecked
    Else
        GlBlE = False 'DATEV Belegexport
        ChEml.Enabled = True
    End If
End If

End Sub

Private Sub chkReAbs_Click()
On Error Resume Next

Set CmFma = Me.cmbForma
Set ChReA = Me.chkReAbs
Set ChReB = Me.chkReBez

If FoLad = False Then
    If ChReA.Value = xtpChecked Then
        ChReB.Enabled = True
    Else
        ChReB.Enabled = False
        ChReB.Value = xtpUnchecked
    End If
    
    If ChReA.Value = xtpChecked Then
        IniSetVal "System", "ReExAb", -1
    Else
        IniSetVal "System", "ReExAb", 0
        IniSetVal "System", "ReExBe", 0
    End If
End If

End Sub
Private Sub chkReBez_Click()
On Error Resume Next

Set ChReA = Me.chkReAbs
Set ChReB = Me.chkReBez

If FoLad = False Then
    If ChReB.Value = xtpChecked Then
        IniSetVal "System", "ReExBe", -1
    Else
        ChReA.Value = xtpUnchecked
        IniSetVal "System", "ReExBe", 0
        IniSetVal "System", "ReExAb", 0
    End If
End If

End Sub

Private Sub cmbForma_Click()
On Error Resume Next

If FoLad = False Then
    FTyp
End If

End Sub
Private Sub cmbJahre_Click()
    Me.optZeit3.Value = True
End Sub
Private Sub cmbMonat_Click()
    Me.optZeit1.Value = True
End Sub
Private Sub cmbQurta_Click()
    Me.optZeit2.Value = True
End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True
FInit
FoLad = False
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub optAnder_Click()
    
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set CmMan = Me.cmbManda

If Opti1.Value = True Then
    CmMan.Enabled = False
Else
    CmMan.Enabled = True
End If
    
End Sub
Private Sub optSelbe_Click()
    
Set Opti1 = Me.optSelbe
Set Opti2 = Me.optAnder
Set CmMan = Me.cmbManda

If Opti1.Value = True Then
    CmMan.Enabled = False
Else
    CmMan.Enabled = True
End If
    
End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub
Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If Weekday(Day, vbMonday) = vbSaturday Then
        Metrics.ForeColor = vbRed
    End If
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu2_GotFocus()
    Me.txtDatu2.SelStart = 0
    Me.txtDatu2.SelLength = Len(Me.txtDatu2.Text)
End Sub

