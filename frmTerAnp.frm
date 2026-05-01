VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmTerAnp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Termine ƒndern"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   20
      Top             =   6900
      Width           =   6800
      _Version        =   1048579
      _ExtentX        =   11994
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   4800
         TabIndex        =   23
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
         Height          =   400
         Left            =   3400
         TabIndex        =   22
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
         Left            =   2100
         TabIndex        =   21
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
      Height          =   6900
      Left            =   300
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   12171
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtKennu 
         Height          =   350
         Left            =   3900
         TabIndex        =   37
         Top             =   6400
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkKennu 
         Height          =   220
         Left            =   3900
         TabIndex        =   36
         Top             =   6100
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Terminkennung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkNotif 
         Height          =   220
         Left            =   3900
         TabIndex        =   34
         Top             =   5200
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Emailerinnerung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   350
         Left            =   1830
         TabIndex        =   19
         Top             =   6400
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   1830
         TabIndex        =   16
         Top             =   5500
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   350
         Left            =   600
         TabIndex        =   18
         Top             =   6400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   600
         TabIndex        =   15
         Top             =   5500
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkFall2 
         Height          =   220
         Left            =   600
         TabIndex        =   17
         Top             =   6100
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "F‰lligkeit 2"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFall1 
         Height          =   220
         Left            =   600
         TabIndex        =   14
         Top             =   5200
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "F‰lligkeit 1"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkPassi 
         Height          =   220
         Left            =   3900
         TabIndex        =   32
         Top             =   4300
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Entfernt"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkMitar 
         Height          =   220
         Left            =   600
         TabIndex        =   6
         Top             =   2500
         Width           =   1095
         _Version        =   1048579
         _ExtentX        =   1931
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mitarbeiter"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkManda 
         Height          =   220
         Left            =   600
         TabIndex        =   4
         Top             =   1600
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandant"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkTeTyp 
         Height          =   220
         Left            =   600
         TabIndex        =   2
         Top             =   800
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Termintyp"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbTeTyp 
         Height          =   310
         Left            =   600
         TabIndex        =   3
         Top             =   1100
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   315
         Left            =   600
         TabIndex        =   5
         Top             =   1900
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.CheckBox chkRefNr 
         Height          =   220
         Left            =   3900
         TabIndex        =   24
         Top             =   800
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Terminserie"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtRefNr 
         Height          =   350
         Left            =   3900
         TabIndex        =   25
         Top             =   1100
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbRaum1 
         Height          =   310
         Left            =   3900
         TabIndex        =   27
         Top             =   1900
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkRaum1 
         Height          =   220
         Left            =   3900
         TabIndex        =   26
         Top             =   1600
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Raum"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkRepli 
         Height          =   220
         Left            =   3900
         TabIndex        =   28
         Top             =   2500
         Width           =   1605
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Synchronisierung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbRepli 
         Height          =   315
         Left            =   3900
         TabIndex        =   29
         Top             =   2800
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkZeSta 
         Height          =   220
         Left            =   600
         TabIndex        =   8
         Top             =   3400
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Startzeit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   340
         Left            =   1810
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3690
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu1"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtVonZe 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   3700
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkZeEnd 
         Height          =   220
         Left            =   600
         TabIndex        =   11
         Top             =   4300
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Endzeit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.UpDown updCont3 
         Height          =   340
         Left            =   1810
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4590
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         Enabled         =   0   'False
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDatu1"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.FlatEdit txtBisZe 
         Height          =   315
         Left            =   600
         TabIndex        =   12
         Top             =   4600
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.CheckBox chkAbger 
         Height          =   220
         Left            =   3900
         TabIndex        =   30
         Top             =   3400
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Abgerechnet"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbAbger 
         Height          =   315
         Left            =   3900
         TabIndex        =   31
         Top             =   3700
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   600
         TabIndex        =   7
         Top             =   2800
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbPassi 
         Height          =   315
         Left            =   3900
         TabIndex        =   33
         Top             =   4600
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbNotVa 
         Height          =   315
         Left            =   3900
         TabIndex        =   35
         Top             =   5500
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTerAnp.frx":0000
         Height          =   585
         Left            =   400
         TabIndex        =   38
         Top             =   100
         Width           =   5500
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   8400
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   300
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   8400
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   0
   End
End
Attribute VB_Name = "frmTerAnp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private TxRef As XtremeSuiteControls.FlatEdit
Private VoZei As XtremeSuiteControls.FlatEdit
Private BiZei As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxKen As XtremeSuiteControls.FlatEdit
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmRmu As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmRep As XtremeSuiteControls.ComboBox
Private CmAbg As XtremeSuiteControls.ComboBox
Private CmPas As XtremeSuiteControls.ComboBox
Private CmNot As XtremeSuiteControls.ComboBox
Private ChRef As XtremeSuiteControls.CheckBox
Private ChTyp As XtremeSuiteControls.CheckBox
Private ChRmu As XtremeSuiteControls.CheckBox
Private ChMan As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private ChRep As XtremeSuiteControls.CheckBox
Private ChZeS As XtremeSuiteControls.CheckBox
Private ChZeE As XtremeSuiteControls.CheckBox
Private ChAbg As XtremeSuiteControls.CheckBox
Private ChPas As XtremeSuiteControls.CheckBox
Private ChFa1 As XtremeSuiteControls.CheckBox
Private ChFa2 As XtremeSuiteControls.CheckBox
Private ChNot As XtremeSuiteControls.CheckBox
Private ChKen As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private UpCo2 As XtremeSuiteControls.UpDown
Private UpCo3 As XtremeSuiteControls.UpDown
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private MoKal As XtremeCalendarControl.DatePicker
Private ImMan As XtremeCommandBars.ImageManager

Private KalWa As Integer
Private Sub TAbs()
On Error GoTo OpErr
'Anpassen der Termine

Dim RowNr As Long
Dim AnzPo As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

AnzPo = RpSel.Count

Screen.MousePointer = vbHourglass
DoEvents

If AnzPo > 0 Then
    S_TeAnp
    DoEvents
    If AnzPo > 1 Then
        SUpTe
    Else
        Set RpSel = RpCo1.SelectedRows
        If RpSel.Count > 0 Then
            Set RpRow = RpSel(0)
            RowNr = RpRow.Index
            SUpTe RowNr
        End If
    End If
    DoEvents
    S_TeLi
End If

DoEvents
Screen.MousePointer = vbNormal

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TAbs " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim TmGui As String
Dim AktZa As Integer
Dim NotVa As Integer

Set FM = frmTerAnp
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set MoKal = FM.dtpDatu1
Set TxRef = FM.txtRefNr
Set CmTyp = FM.cmbTeTyp
Set CmRmu = FM.cmbRaum1
Set CmMan = FM.cmbBehan
Set CmMit = FM.cmbMitar
Set CmRep = FM.cmbRepli
Set CmAbg = FM.cmbAbger
Set CmNot = FM.cmbNotVa
Set CmPas = FM.cmbPassi
Set ChTyp = FM.chkTeTyp
Set ChRef = FM.chkRefNr
Set ChRmu = FM.chkRaum1
Set ChMan = FM.chkManda
Set ChMit = FM.chkMitar
Set ChRep = FM.chkRepli
Set ChZeS = FM.chkZeSta
Set ChZeE = FM.chkZeEnd
Set ChAbg = FM.chkAbger
Set ChPas = FM.chkPassi
Set ChFa1 = FM.chkFall1
Set ChFa2 = FM.chkFall2
Set ChNot = FM.chkNotif
Set ChKen = FM.chkKennu
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtDatu2
Set TxKen = FM.txtKennu
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnDatu2
Set ImMan = frmMain.imgManag

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

For AktZa = 1 To UBound(GlTep) 'Kalendermarker
    CmTyp.AddItem GlTep(AktZa, 1)
    CmTyp.ItemData(AktZa - 1) = GlTep(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlRmu)
    CmRmu.AddItem GlRmu(AktZa, 1)
    CmRmu.ItemData(AktZa - 1) = GlRmu(AktZa, 2)
Next AktZa

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa

With CmMit
    For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
        .AddItem GlMiK(AktZa, 1)
        .ItemData(AktZa - 1) = GlMiK(AktZa, 2)
    Next AktZa
    .ListIndex = GlSmI - 1
End With

With CmRep
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmAbg
    .AddItem "keine Leistungen"
    .ItemData(0) = 1
    .AddItem "Leistungen vorhanden"
    .ItemData(1) = 2
    .AddItem "Leistungen abgerechnet"
    .ItemData(2) = 3
End With

With CmNot
    For AktZa = 0 To 48
        .AddItem AktZa & " Std."
        .ItemData(AktZa) = AktZa
    Next AktZa
End With

With CmPas
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

CmTyp.ListIndex = 0
CmRmu.ListIndex = 0
CmRep.ListIndex = 0
CmAbg.ListIndex = 0
CmMan.ListIndex = GlSMa - 1
CmMit.ListIndex = GlSmI - 1
CmPas.ListIndex = 1

With TxRef
    .Pattern = "\d*"
    .SetMask "000000", "______"
    .Text = "000001"
End With

VoZei.SetMask "00:00", "__:__"
BiZei.SetMask "00:00", "__:__"

VoZei.Text = "08:00"
BiZei.Text = "09:00"

If GlMiV = False Then
    CmMit.Enabled = False
    ChMit.Enabled = False
End If

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    NotVa = GlMiT(1, 39)
Else
    NotVa = GlMaT(1, 25)
End If

If NotVa = 0 Then
    NotVa = 24
End If

If GlTeE = True Then 'Email-Termin-Erinnerung
    CmNot.ListIndex = NotVa
Else
    CmNot.ListIndex = 0
End If

If GlOTS = True Then 'Online-Terminbuchungs Sytem
    If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
        CmMit.Enabled = False
        ChMit.Enabled = False
    Else
        CmMan.Enabled = False
        ChMan.Enabled = False
    End If
    'ChPas.Enabled = False
    ChZeS.Enabled = False
    ChZeE.Enabled = False
End If

TmGui = CreateID("T")
TxKen.Text = TmGui

ChNot.Enabled = GlTeE 'Email-Termin-Erinnerung

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
ChRef.BackColor = GlBak
ChTyp.BackColor = GlBak
ChRmu.BackColor = GlBak
ChMan.BackColor = GlBak
ChMit.BackColor = GlBak
ChRep.BackColor = GlBak
ChZeS.BackColor = GlBak
ChZeE.BackColor = GlBak
ChAbg.BackColor = GlBak
ChPas.BackColor = GlBak
ChFa1.BackColor = GlBak
ChFa2.BackColor = GlBak
ChNot.BackColor = GlBak
ChKen.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
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

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    Select Case KalWa
    Case 1: TxDa1.Text = NeuDa
    Case 2: TxDa2.Text = NeuDa
    End Select
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
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
            .Left = TxDa1.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa1.Text = .Selection.Blocks(0).DateBegin
                    TxDa2.Text = .Selection.Blocks(0).DateBegin + 14
                End If
            End If
    Case 2: .Top = TxDa2.Top + TxDa2.Height
            .Left = TxDa2.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    End Select
End With

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        Datu1 = TxDa1.Text
    End If
End If
If TxDa2.Text <> vbNullString Then
    If IsDate(TxDa2.Text) = True Then
        Datu2 = TxDa2.Text
    End If
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub

Private Sub chkAbger_Click()
On Error Resume Next

Set ChAbg = Me.chkAbger
Set CmAbg = Me.cmbAbger

If ChAbg.Value = xtpChecked Then
    CmAbg.Enabled = True
Else
    CmAbg.Enabled = False
End If

End Sub

Private Sub chkFall1_Click()
On Error Resume Next

Set ChFa1 = Me.chkFall1
Set TxDa1 = Me.txtDatu1
Set PuBu1 = Me.btnDatu1

If ChFa1.Value = xtpChecked Then
    TxDa1.Enabled = True
    PuBu1.Enabled = True
Else
    TxDa1.Enabled = False
    PuBu1.Enabled = False
End If

End Sub

Private Sub chkFall2_Click()
On Error Resume Next

Set ChFa2 = Me.chkFall2
Set TxDa2 = Me.txtDatu2
Set PuBu2 = Me.btnDatu2

If ChFa2.Value = xtpChecked Then
    TxDa2.Enabled = True
    PuBu2.Enabled = True
Else
    TxDa2.Enabled = False
    PuBu2.Enabled = False
End If

End Sub

Private Sub chkKennu_Click()
On Error Resume Next

Set ChKen = Me.chkKennu
Set TxKen = Me.txtKennu

If ChKen.Value = xtpChecked Then
    TxKen.Enabled = True
Else
    TxKen.Enabled = False
End If

End Sub
Private Sub chkManda_Click()
On Error Resume Next

Set ChMan = Me.chkManda
Set CmMan = Me.cmbBehan

If ChMan.Value = xtpChecked Then
    CmMan.Enabled = True
Else
    CmMan.Enabled = False
End If

End Sub

Private Sub chkMitar_Click()
On Error Resume Next

Set ChMit = Me.chkMitar
Set CmMit = Me.cmbMitar

If ChMit.Value = xtpChecked Then
    CmMit.Enabled = True
Else
    CmMit.Enabled = False
End If

End Sub

Private Sub chkNotif_Click()
On Error Resume Next

Set ChNot = Me.chkNotif
Set CmNot = Me.cmbNotVa

If ChNot.Value = xtpChecked Then
    CmNot.Enabled = True
Else
    CmNot.Enabled = False
End If

End Sub
Private Sub chkPassi_Click()
On Error Resume Next

Set ChPas = Me.chkPassi
Set CmPas = Me.cmbPassi

If ChPas.Value = xtpChecked Then
    CmPas.Enabled = True
Else
    CmPas.Enabled = False
End If

End Sub
Private Sub chkRaum1_Click()
On Error Resume Next

Set ChRmu = Me.chkRaum1
Set CmRmu = Me.cmbRaum1

If ChRmu.Value = xtpChecked Then
    CmRmu.Enabled = True
Else
    CmRmu.Enabled = False
End If

End Sub
Private Sub chkRefNr_Click()
On Error Resume Next

Set ChRef = Me.chkRefNr
Set TxRef = Me.txtRefNr

If ChRef.Value = xtpChecked Then
    TxRef.Enabled = True
Else
    TxRef.Enabled = False
End If

End Sub

Private Sub chkRepli_Click()
On Error Resume Next

Set ChRep = Me.chkRepli
Set CmRep = Me.cmbRepli

If ChRep.Value = xtpChecked Then
    CmRep.Enabled = True
Else
    CmRep.Enabled = False
End If

End Sub
Private Sub chkTeTyp_Click()
On Error Resume Next

Set ChTyp = Me.chkTeTyp
Set CmTyp = Me.cmbTeTyp

If ChTyp.Value = xtpChecked Then
    CmTyp.Enabled = True
Else
    CmTyp.Enabled = False
End If

End Sub

Private Sub chkZeEnd_Click()
On Error Resume Next

Set ChZeE = Me.chkZeEnd
Set BiZei = Me.txtBisZe
Set UpCo3 = Me.updCont3

If ChZeE.Value = xtpChecked Then
    BiZei.Enabled = True
    UpCo3.Enabled = True
Else
    BiZei.Enabled = False
    UpCo3.Enabled = False
End If

End Sub
Private Sub chkZeSta_Click()
On Error Resume Next

Set ChZeS = Me.chkZeSta
Set VoZei = Me.txtVonZe
Set UpCo2 = Me.updCont2

If ChZeS.Value = xtpChecked Then
    VoZei.Enabled = True
    UpCo2.Enabled = True
Else
    VoZei.Enabled = False
    UpCo2.Enabled = False
End If

End Sub

Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub btnHilfe_Click()
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
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTerAnp = Nothing
End Sub
Private Sub btnWeiter_Click()
    TAbs
    Unload Me
End Sub

Private Sub txtBisZe_GotFocus()
    Me.txtBisZe.SelStart = 0
    Me.txtBisZe.SelLength = Len(Me.txtBisZe.Text)
End Sub

Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu1_LostFocus()
    KalWa = 1
    FDaKo
End Sub

Private Sub txtDatu2_GotFocus()
    Me.txtDatu2.SelStart = 0
    Me.txtDatu2.SelLength = Len(Me.txtDatu2.Text)
End Sub
Private Sub txtDatu2_LostFocus()
    KalWa = 2
    FDaKo
End Sub

Private Sub txtKennu_GotFocus()
    Me.txtKennu.SelStart = 0
    Me.txtKennu.SelLength = Len(Me.txtKennu.Text)
End Sub
Private Sub txtRefNr_GotFocus()
    Me.txtRefNr.SelStart = 0
    Me.txtRefNr.SelLength = Len(Me.txtRefNr.Text)
End Sub
Private Sub txtVonZe_GotFocus()
    Me.txtVonZe.SelStart = 0
    Me.txtVonZe.SelLength = Len(Me.txtVonZe.Text)
End Sub
Private Sub updCont2_DownClick()
On Error Resume Next

Dim MitNr As Long
Dim ManNr As Long
Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    MitNr = GlMiA(GlSmI, 2)
    For AktZa = 1 To UBound(GlMiT)
        If MitNr = GlMiT(AktZa, 2) Then
            ZeiRa = GlMiT(AktZa, 8)
            Exit For
        End If
    Next AktZa
Else
    ManNr = GlMan(GlSMa, 2)
    For AktZa = 1 To UBound(GlMaT)
        If ManNr = GlMaT(AktZa, 2) Then
            ZeiRa = GlMaT(AktZa, 8)
            Exit For
        End If
    Next AktZa
End If

If ZeiRa = 0 Then
    ZeiRa = GlZeR 'Zeitrasterindex
End If

ZeiVo = 15
MiDif = GlTku(ZeiRa, 2)

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
                
        TmVon = DateAdd("n", -MiDif, AlDa1)
        VoZei.Text = Format$(TmVon, "hh:mm")

        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmBis = DateAdd("n", ZeiVo, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
            End If
        Else
            If TmVon >= AlDa2 Then
                TmBis = DateAdd("n", MiDif, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
            End If
        End If
    End If
End If

End Sub
Private Sub updCont2_UpClick()
On Error Resume Next

Dim MitNr As Long
Dim ManNr As Long
Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    MitNr = GlMiA(GlSmI, 2)
    For AktZa = 1 To UBound(GlMiT)
        If MitNr = GlMiT(AktZa, 2) Then
            ZeiRa = GlMiT(AktZa, 8)
            Exit For
        End If
    Next AktZa
Else
    ManNr = GlMan(GlSMa, 2)
    For AktZa = 1 To UBound(GlMaT)
        If ManNr = GlMaT(AktZa, 2) Then
            ZeiRa = GlMaT(AktZa, 8)
            Exit For
        End If
    Next AktZa
End If

If ZeiRa = 0 Then
    ZeiRa = GlZeR 'Zeitrasterindex
End If

ZeiVo = 15
MiDif = GlTku(ZeiRa, 2)

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
        
        TmVon = DateAdd("n", MiDif, AlDa1)
        VoZei.Text = Format$(TmVon, "hh:mm")

        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmBis = DateAdd("n", ZeiVo, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
            End If
        Else
            If TmVon >= AlDa2 Then
                TmBis = DateAdd("n", MiDif, TmVon)
                BiZei.Text = Format$(TmBis, "hh:mm")
            End If
        End If
    End If
End If

End Sub
Private Sub updCont3_DownClick()
On Error Resume Next

Dim MitNr As Long
Dim ManNr As Long
Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim AktZa As Integer
Dim ZeiVo As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    MitNr = GlMiA(GlSmI, 2)
    For AktZa = 1 To UBound(GlMiT)
        If MitNr = GlMiT(AktZa, 2) Then
            ZeiRa = GlMiT(AktZa, 8)
            Exit For
        End If
    Next AktZa
Else
    ManNr = GlMan(GlSMa, 2)
    For AktZa = 1 To UBound(GlMaT)
        If ManNr = GlMaT(AktZa, 2) Then
            ZeiRa = GlMaT(AktZa, 8)
            Exit For
        End If
    Next AktZa
End If

If ZeiRa = 0 Then
    ZeiRa = GlZeR 'Zeitrasterindex
End If

ZeiVo = 15
MiDif = GlTku(ZeiRa, 2)

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
        
        TmBis = DateAdd("n", -MiDif, AlDa2)
        BiZei.Text = Format$(TmBis, "hh:mm")
        
        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmVon = DateAdd("n", -ZeiVo, TmBis)
                VoZei.Text = Format$(TmVon, "hh:mm")
            End If
        Else
            If TmBis <= AlDa1 Then
                TmVon = DateAdd("n", -MiDif, AlDa1)
                VoZei.Text = Format$(TmVon, "hh:mm")
            End If
        End If
    End If
End If

End Sub
Private Sub updCont3_UpClick()
On Error Resume Next

Dim MitNr As Long
Dim ManNr As Long
Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim AktZa As Integer
Dim ZeiVo As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    MitNr = GlMiA(GlSmI, 2)
    For AktZa = 1 To UBound(GlMiT)
        If MitNr = GlMiT(AktZa, 2) Then
            ZeiRa = GlMiT(AktZa, 8)
            Exit For
        End If
    Next AktZa
Else
    ManNr = GlMan(GlSMa, 2)
    For AktZa = 1 To UBound(GlMaT)
        If ManNr = GlMaT(AktZa, 2) Then
            ZeiRa = GlMaT(AktZa, 8)
            Exit For
        End If
    Next AktZa
End If

If ZeiRa = 0 Then
    ZeiRa = GlZeR 'Zeitrasterindex
End If

ZeiVo = 15
MiDif = GlTku(ZeiRa, 2)

If VoZei.Text <> vbNullString Then
    If BiZei.Text <> vbNullString Then
        AlDa1 = TimeValue(VoZei.Text)
        AlDa2 = TimeValue(BiZei.Text)
        
        TmBis = DateAdd("n", MiDif, AlDa2)
        BiZei.Text = Format$(TmBis, "hh:mm")
        
        If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
            If ZeiVo > 0 Then
                TmVon = DateAdd("n", -ZeiVo, TmBis)
                VoZei.Text = Format$(TmVon, "hh:mm")
            End If
        End If
    End If
End If

End Sub
Private Sub btnDatu1_Click()
    KalWa = 1
    FKale
End Sub
Private Sub btnDatu2_Click()
    KalWa = 2
    FKale
End Sub
