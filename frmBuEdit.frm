VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmBuEdit 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Buchung"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   23
      Top             =   7000
      Width           =   7000
      _Version        =   1048579
      _ExtentX        =   12347
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   900
         TabIndex        =   24
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "Hilfe"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   5000
         TabIndex        =   27
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
      Begin XtremeSuiteControls.PushButton cmdNeuBu 
         Height          =   400
         Left            =   2200
         TabIndex        =   25
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Neu Buchung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdWeite 
         Height          =   400
         Left            =   3600
         TabIndex        =   26
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Speichern [F8]"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtTSELg 
      Height          =   350
      Left            =   2800
      TabIndex        =   16
      Top             =   4630
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Enabled         =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtBezei 
      Height          =   200
      Left            =   750
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtKonto 
      Height          =   350
      Left            =   1200
      TabIndex        =   6
      Top             =   1130
      Width           =   4740
      _Version        =   1048579
      _ExtentX        =   8361
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      MaxLength       =   250
   End
   Begin XtremeSuiteControls.UpDown updCont1 
      Height          =   340
      Left            =   2340
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   440
      Width           =   255
      _Version        =   1048579
      _ExtentX        =   450
      _ExtentY        =   600
      _StockProps     =   64
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtDatu1"
      BuddyProperty   =   ""
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   495
      Left            =   4000
      TabIndex        =   31
      Top             =   8500
      Visible         =   0   'False
      Width           =   495
      _Version        =   1048579
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.PushButton btnDatu1 
      Height          =   350
      Left            =   2610
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Öffnet den Auswahlkalender"
      Top             =   430
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   617
      _ExtentY        =   617
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkGewEr 
      Height          =   220
      Left            =   1220
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Diese Buchung wird bei der Einnahmenüberschussrechnung nicht berücksichtigt"
      Top             =   6600
      Width           =   3300
      _Version        =   1048579
      _ExtentX        =   5821
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Keine Berücksichtigung bei Erlösermittlung"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtBuKom 
      Height          =   350
      Left            =   1200
      TabIndex        =   17
      Top             =   5330
      Width           =   4700
      _Version        =   1048579
      _ExtentX        =   8290
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbBuTyp 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Top             =   2530
      Width           =   1380
      _Version        =   1048579
      _ExtentX        =   2434
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.ComboBox cmbBuStu 
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   3230
      Width           =   1380
      _Version        =   1048579
      _ExtentX        =   2434
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox4"
   End
   Begin XtremeSuiteControls.ComboBox cmbGegen 
      Height          =   315
      Left            =   2800
      TabIndex        =   9
      Top             =   2530
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox5"
   End
   Begin XtremeSuiteControls.ComboBox cmbManda 
      Height          =   315
      Left            =   2800
      TabIndex        =   12
      Top             =   3230
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox6"
   End
   Begin XtremeSuiteControls.ComboBox cmbWarun 
      Height          =   315
      Left            =   4545
      TabIndex        =   5
      Top             =   435
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbBuTex 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   1830
      Width           =   4740
      _Version        =   1048579
      _ExtentX        =   8361
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      MaxLength       =   250
   End
   Begin XtremeSuiteControls.FlatEdit txtDatu1 
      Height          =   350
      Left            =   1200
      TabIndex        =   1
      Top             =   431
      Width           =   1130
      _Version        =   1048579
      _ExtentX        =   1993
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtBuBel 
      Height          =   350
      Left            =   1200
      TabIndex        =   15
      Top             =   4630
      Width           =   1380
      _Version        =   1048579
      _ExtentX        =   2434
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Text            =   "1"
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtBuBet 
      Height          =   350
      Left            =   3100
      TabIndex        =   4
      Top             =   431
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtBaIdx 
      Height          =   200
      Left            =   1200
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.ComboBox cmbMitar 
      Height          =   315
      Left            =   2800
      TabIndex        =   14
      Top             =   3930
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtRecNr 
      Height          =   350
      Left            =   1200
      TabIndex        =   13
      Top             =   3930
      Width           =   1380
      _Version        =   1048579
      _ExtentX        =   2434
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
      MaxLength       =   50
   End
   Begin XtremeSuiteControls.FlatEdit txtDatei 
      Height          =   350
      Left            =   1200
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6030
      Width           =   3970
      _Version        =   1048579
      _ExtentX        =   7003
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.PushButton btnFile1 
      Height          =   350
      Left            =   5200
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Ordnet der Buchung ein Dokument zu"
      Top             =   6030
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   617
      _ExtentY        =   617
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnFile2 
      Height          =   350
      Left            =   5580
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Löscht das zugeordnete Dokument"
      Top             =   6030
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   617
      _ExtentY        =   617
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoNr 
      Height          =   200
      Left            =   360
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtHaben 
      Height          =   350
      Left            =   2800
      TabIndex        =   10
      Top             =   2530
      Width           =   3120
      _Version        =   1048579
      _ExtentX        =   5503
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtKtoHa 
      Height          =   200
      Left            =   1680
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtBezHa 
      Height          =   200
      Left            =   2085
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtIdxNr 
      Height          =   200
      Left            =   2500
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtBuGui 
      Height          =   200
      Left            =   3000
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   8500
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.CheckBox chkGutha 
      Height          =   230
      Left            =   4700
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Dem zugeordneten Patienten wird der Buchungsbetrag als Guthaben hinzugefügt"
      Top             =   6600
      Width           =   1700
      _Version        =   1048579
      _ExtentX        =   2999
      _ExtentY        =   406
      _StockProps     =   79
      Caption         =   "Guthabenbuchung"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab15 
      Height          =   210
      Left            =   2805
      TabIndex        =   48
      Top             =   4380
      Width           =   900
      _Version        =   1048579
      _ExtentX        =   1587
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "TSE-Log :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab06 
      Height          =   210
      Left            =   2850
      TabIndex        =   45
      Top             =   2300
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Geldkonto :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab05 
      Height          =   210
      Left            =   1205
      TabIndex        =   44
      Top             =   870
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Sachkonto :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Label lblLab14 
      BackStyle       =   0  'Transparent
      Caption         =   "Dokument :"
      Height          =   210
      Left            =   1205
      TabIndex        =   43
      Top             =   5800
      Width           =   1000
   End
   Begin VB.Label lblLab12 
      BackStyle       =   0  'Transparent
      Caption         =   "Mitarbeiter :"
      Height          =   210
      Left            =   2805
      TabIndex        =   42
      Top             =   3680
      Width           =   900
   End
   Begin VB.Label lblLab11 
      BackStyle       =   0  'Transparent
      Caption         =   "Belegkennzeichen :"
      Height          =   210
      Left            =   1205
      TabIndex        =   41
      Top             =   3680
      Width           =   1500
   End
   Begin VB.Label lblLab08 
      BackStyle       =   0  'Transparent
      Caption         =   "Steuersatz :"
      Height          =   210
      Left            =   1205
      TabIndex        =   40
      Top             =   2990
      Width           =   1095
   End
   Begin VB.Label lblLab02 
      BackStyle       =   0  'Transparent
      Caption         =   "Betrag :"
      Height          =   210
      Left            =   3105
      TabIndex        =   39
      Top             =   200
      Width           =   900
   End
   Begin VB.Label lblLab09 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungsnummer :"
      Height          =   210
      Left            =   1205
      TabIndex        =   38
      Top             =   4380
      Width           =   1400
   End
   Begin VB.Label lblLab04 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungstext :"
      Height          =   210
      Left            =   1205
      TabIndex        =   37
      Top             =   1590
      Width           =   1200
   End
   Begin VB.Label lblLab07 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungstyp :"
      Height          =   210
      Left            =   1205
      TabIndex        =   36
      Top             =   2300
      Width           =   1100
   End
   Begin VB.Label lblLab03 
      BackStyle       =   0  'Transparent
      Caption         =   "Währung :"
      Height          =   210
      Left            =   4545
      TabIndex        =   35
      Top             =   200
      Width           =   900
   End
   Begin VB.Label lblLab01 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum :"
      Height          =   210
      Left            =   1205
      TabIndex        =   34
      Top             =   200
      Width           =   900
   End
   Begin VB.Label lblLab13 
      BackStyle       =   0  'Transparent
      Caption         =   "Kommentar :"
      Height          =   210
      Left            =   1205
      TabIndex        =   33
      Top             =   5100
      Width           =   1000
   End
   Begin VB.Label lblLab10 
      BackStyle       =   0  'Transparent
      Caption         =   "Mandant :"
      Height          =   210
      Left            =   2805
      TabIndex        =   32
      Top             =   2990
      Width           =   900
   End
End
Attribute VB_Name = "frmBuEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Lbl05 As XtremeSuiteControls.Label
Private Lbl06 As XtremeSuiteControls.Label
Private Rahm0 As XtremeSuiteControls.GroupBox
Private ChAsw As XtremeSuiteControls.CheckBox
Private ChGut As XtremeSuiteControls.CheckBox
Private CmWar As XtremeSuiteControls.ComboBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmBuT As XtremeSuiteControls.ComboBox
Private CmBuS As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMta As XtremeSuiteControls.ComboBox
Private CmBTe As XtremeSuiteControls.ComboBox
Private CmBTy As XtremeSuiteControls.ComboBox
Private CmBar As XtremeCommandBars.CommandBar
Private CmAcs As XtremeCommandBars.CommandBarActions
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRow As XtremeReportControl.ReportRow
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxBel As XtremeSuiteControls.FlatEdit
Private TxFil As XtremeSuiteControls.FlatEdit
Private TxKto As XtremeSuiteControls.FlatEdit
Private TxHab As XtremeSuiteControls.FlatEdit
Private TxBet As XtremeSuiteControls.FlatEdit
Private UpDo1 As XtremeSuiteControls.UpDown
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private PuBu4 As XtremeSuiteControls.PushButton
Private PuBu5 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private MoKal As XtremeCalendarControl.DatePicker
Private CoDia As XtremeSuiteControls.CommonDialog
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private BuBet As Double

Public PatNr As Long
Public BaBet As Double
Public GeBet As Double
Public BaGui As String

Private BuBel As Long
Private KntRa As Integer
Private FoLad As Boolean

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

Private clFil As clsFile

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FBuBe()
On Error GoTo InErr
'Prüft und ermittelt die nächste Belegnummer

Dim BuDat As Date
Dim ManNr As Long
Dim BuJah As Long
Dim BnkNr As Long
Dim BelNr As Long
Dim BlNum As Long
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmBuEdit
Set TxBel = FM.txtBuBel
Set TxDa1 = Me.txtDatu1
Set CmMan = FM.cmbManda
Set CmGeg = FM.cmbGegen

Set RpCo1 = frmMain.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Buh_Datum)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            BuDat = CDate(RpRow.Record(RpCol.ItemIndex).Value)
        Else
            BuDat = Date
        End If
        Set RpCol = RpCls.Find(Buh_Beleg)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            If IsNumeric(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                BlNum = CLng(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                BlNum = 0
            End If
        Else
            BlNum = 0
        End If
    End If
End If

If TxBel.Text <> vbNullString Then
    If IsNumeric(TxBel.Text) = True Then
        BelNr = TxBel.Text
    Else
        BelNr = 0
    End If
Else
    BelNr = 0
End If

If CmGeg.Text <> vbNullString Then
    BnkNr = CmGeg.ItemData(CmGeg.ListIndex)
Else
    BnkNr = GlGkB 'Standardgeldkonto Bankkonto
End If

If CmMan.Text <> vbNullString Then
    ManNr = CmMan.ItemData(CmMan.ListIndex)
Else
    ManNr = 0
End If

If IsDate(TxDa1.Text) = True Then
    BuJah = Year(TxDa1.Text)
Else
    BuJah = Year(Date)
End If

If GlNeB = True Then 'neue Buchung
    TxBel.Text = Format$(S_BuBel(ManNr, BuJah, BnkNr), "000000")
Else
    If Year(BuDat) <> BuJah Then
        TxBel.Text = Format$(S_BuBel(ManNr, BuJah, BnkNr), "000000")
    Else
        TxBel.Text = BlNum
    End If
End If

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FBuBe & Err.Number"
Resume Next

End Sub
Private Sub FBuBp()
On Error GoTo InErr
'Prüft ob die Belegnummer bereits vorhanden ist

Dim ManNr As Long
Dim BuJah As Long
Dim BnkNr As Long
Dim BelNr As Long
Dim BelOr As Long
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmBuEdit
Set TxBel = FM.txtBuBel
Set TxDa1 = Me.txtDatu1
Set CmMan = FM.cmbManda
Set CmGeg = FM.cmbGegen
Set RpCo1 = frmMain.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Buh_Beleg)
        If RpRow.Record(RpCol.ItemIndex).Value <> vbNullString Then
            BelOr = RpRow.Record(RpCol.ItemIndex).Value
        Else
            BelOr = 0
        End If
    End If
End If

If TxBel.Text <> vbNullString Then
    BelNr = TxBel.Text
Else
    BelNr = 0
End If

If CmGeg.Text <> vbNullString Then
    BnkNr = CmGeg.ItemData(CmGeg.ListIndex)
Else
    BnkNr = GlGkB 'Standardgeldkonto Bankkonto
End If

If CmMan.Text <> vbNullString Then
    ManNr = CmMan.ItemData(CmMan.ListIndex)
Else
    ManNr = 0
End If

If IsDate(TxDa1.Text) = True Then
    BuJah = Year(TxDa1.Text)
Else
    BuJah = Year(Date)
End If

If S_BuBeP(BelNr, ManNr, BuJah, BnkNr) = True Then
    If GlNeB = True Then 'neue Buchung
        If GlBGe = True Then 'Getrennter Geldkonten Belegnummernkreis
            TxBel.Text = Format$(S_BuBel(ManNr, BuJah, BnkNr), "000000")
        ElseIf GlBMa = True Then 'Getrennter Mandanten Belegnummernkreis
            TxBel.Text = Format$(S_BuBel(ManNr, BuJah, BnkNr), "000000")
        End If
    Else
        TxBel.Text = Format$(BelOr, "000000")
    End If
End If

Set RpCo1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FBuBp & Err.Number"
Resume Next

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = CDate(TxDa1.Text)
    If Year(NeuDa) > Year(Date) Then
        NeuDa = "31.12." & Year(Date)
        SPopu "Buchungsjahr überschritten", "Das neue Buchungsdatum muss sich im selben Buchungsjahr befinden.", IC48_Information
    End If
    TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
    With MoKal
        .EnsureVisible NeuDa - 30
        .Select NeuDa
        .SelectRange NeuDa, NeuDa
    End With
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

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = CDate(MoKal.Selection.Blocks(0).DateBegin)
    If Year(NeuDa) <= Year(Date) Then
        TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
    Else
        SPopu "Buchungsjahr überschritten", "Das neue Buchungsdatum muss sich im selben Buchungsjahr befinden.", IC48_Information
    End If
    TxDa1.SetFocus
End If

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

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If IsDate(TxDa1.Text) Then
    NeuDa = CDate(TxDa1.Text)
Else
    NeuDa = Date
End If

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .Top = TxDa1.Top + TxDa1.Height
    .Left = TxDa1.Left
    If .ShowModal(1, 1) Then
        If .Selection.BlocksCount > 0 Then
            TmDat = .Selection.Blocks(0).DateBegin
            If Year(TmDat) <= Year(Date) Then
                TxDa1.Text = Format$(TmDat, "dd.mm.yyyy")
            End If
        End If
    End If
End With

Set MoKal = Nothing

DoEvents
FBuBe

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub FLoe()
On Error GoTo LaErr

Dim FiNam As String
Dim DaNam As String
Dim Frage As Integer
Dim Mld1, Tit1 As String

Set FM = frmBuEdit
Set TxFil = FM.txtDatei
Set PuBu5 = FM.cmdWeite

Tit1 = "Dokument Entfernen"
Mld1 = "Möchten Sie das zugeordnete Dokument wirklich entfernen?"

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

If GlRch(0, 19) = 0 Then
    WindowMess "Sie besitzen keine Berechtigung für diesen Vorgang", Dial3, "Entfernen", FM.hwnd
    Exit Sub
End If

If TxFil.Text <> vbNullString Then
    DaNam = TxFil.Text
    FiNam = GlBPf & DaNam

    Frage = WindowMess(Mld1, Dial1, Tit1, FM.hwnd)
    If Frage = 6 Then
    
        If clFil.FilVor(FiNam) = True Then
            clFil.DaLoe = FiNam & vbNullChar
            clFil.FilLoe
        End If
        
        TxFil.Text = vbNullString
        
        If GlStB = False Then 'Stapelbuchung
            PuBu5.Enabled = False
        End If
        
        S_Save
    End If
End If

Set clFil = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoe " & Err.Number
Resume Next

End Sub
Private Sub FMand()
On Error GoTo OrErr

Dim ManNr As Long
Dim StaRa As Integer
Dim AktZa As Integer

Set CmMan = FM.cmbManda

ManNr = CmMan.ItemData(CmMan.ListIndex)

If GlMVo = True Then 'mandantenbezogene Vorgaben verwenden
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            If GlMan(AktZa, 25) <> vbNullString Then
                KntRa = GlMan(AktZa, 25) 'Standardkontenrahmen
            Else
                KntRa = GlKtR
            End If
            Exit For
        End If
    Next AktZa
End If

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMand " & Err.Number
Resume Next

End Sub
Private Sub FNeBu()
On Error GoTo InErr

Dim ManNr As Long
Dim IdxNr As Long
Dim BuJah As Long
Dim BuAus As String
Dim GeKto As Integer
Dim AktZa As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmGlk As XtremeCommandBars.CommandBarComboBox
Dim CmJah As XtremeCommandBars.CommandBarComboBox
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmBuEdit
Set PuBu1 = FM.btnDatu1
Set TxBel = FM.txtBuBel
Set TxDa1 = Me.txtDatu1
Set CmMan = FM.cmbManda
Set CmMta = FM.cmbMitar
Set CmGeg = FM.cmbGegen
Set CmBTy = FM.cmbBuTyp
Set PuBu5 = FM.cmdWeite
Set UpDo1 = Me.updCont1
Set CmBrs = frmMain.comBar01

Set RpCo1 = frmMain.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

Set CmGlk = CmBrs.FindControl(CmGlk, SY_SuBuh, , True)
Set CmJah = CmBrs.FindControl(CmJah, SY_SuJah, , True)

GeKto = CmGlk.ItemData(CmGlk.ListIndex)

If CmMan.Text <> vbNullString Then
    ManNr = CmMan.ItemData(CmMan.ListIndex)
Else
    ManNr = 0
End If

If IsDate(TxDa1.Text) = True Then
    BuJah = CmJah.Text
Else
    BuJah = Year(Date)
End If

GlNeB = True 'neue Buchung

CmBTy.ListIndex = 0

If BaBet > 0 Then
    FM.txtBuBet.Text = Format$(GeBet, GlWa1)
Else
    FM.txtBuBet.Text = GlWa2
    FM.cmbBuTex.Text = vbNullString
End If
If FM.txtDatu1.Text = vbNullString Then
    FM.txtDatu1.Text = Format$(Date, "dd.mm.yyyy")
End If
FM.txtKonto.Text = vbNullString
FM.txtBezei.Text = vbNullString
FM.cmbWarun.ListIndex = GlStW - 1
FM.cmbBuStu.ListIndex = GlStS - 1
If GlBut <> RibTab_HomeBanki Then
    If GeKto = 0 Then
        If CmGeg.ListCount > 0 Then
            CmGeg.ListIndex = 0
        End If
    Else
        CmGeg.ListIndex = CmGlk.ListIndex - 1
    End If
End If

TxBel.Text = Format$(S_BuBel(ManNr, BuJah, GeKto), "000000")

PuBu5.Enabled = True
TxDa1.Enabled = True
UpDo1.Enabled = True
PuBu1.Enabled = True
DoEvents

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeBu & Err.Number"
Resume Next

End Sub
Private Sub FOpn()
On Error GoTo AnErr

Dim BelNr As Long
Dim DaNam As String
Dim DaExt As String
Dim DaNaO As String
Dim FiNam As String
Dim DaPfa As String
Dim ImOrd As String
Dim BeStr As String
Dim TmGui As String
Dim TypNa As String
Dim PfaNa As String
Dim TeFil As String
Dim NeuDa As String
Dim NeuNa As String

Set FM = frmBuEdit
Set TxBel = FM.txtBuBel
Set TxFil = FM.txtDatei
Set PuBu5 = FM.cmdWeite

Set clFil = New clsFile
clFil.hwnd = FM.hwnd

If TxBel.Text <> vbNullString Then
    If IsNumeric(TxBel.Text) = True Then
        BelNr = CLng(TxBel.Text)
    End If
End If

If BelNr > 0 Then
    BeStr = "B" & Format$(BelNr, "000000")
    TmGui = CreateID("B")
    
    If GlRDP = True Then
        If clFil.FilDir(GlIPf) = False Then
            ImOrd = GlDpf & "Import\"
        Else
            ImOrd = GlIPf
        End If
    Else
        If clFil.FilDir(GlImO) = False Then
            ImOrd = GlIPf 'Importordner
        Else
            ImOrd = GlImO
        End If
    End If
    If Right$(ImOrd, 1) <> "\" Then
        ImOrd = ImOrd & "\"
    End If

    With clFil
        .hwnd = FM.hwnd
        .StaVe = GlIPf
        .DaTit = "Bitte Name und Ordner der Datei angeben"
        .DaStr = "Unterstützte Formate (*.pdf;*.jpg;*.bmp;*.png;*.tif;*.wmf;*.zip)" & Chr(0) & "*.pdf;*.jpg;*.bmp;*.png;*.tif;*.wmf;*.zip" & Chr(0) & "Adobe-Acrobat Dokument (*.pdf)" & Chr(0) & "*.pdf" & Chr(0) & "Joint Photographic Experts Group (.jpg)" & Chr(0) & "*.jpg" & Chr(0) & "Windows Bitmap (.bmp)" & Chr(0) & "*.bmp" & Chr(0) & "Portable Network Graphics (.png)" & Chr(0) & "*.png" & Chr(0) & "Tagged Image Format (.tif)" & Chr(0) & "*.tif" & Chr(0) & "Windows-Meta-File (.wmf)" & Chr(0) & "*.wmf" & Chr(0) & "Microsoft Word 2002/2003 (.doc)" & Chr(0) & "*.doc" & Chr(0) & "Alle Dateien (*.*)" & Chr(0) & "*.*" & Chr(0)
        FiNam = .FilOpn
    End With
    
    If FiNam <> vbNullString Then
        If SNaPr(FiNam) = False Then

            With clFil
                .FilPfa FiNam
                DaNam = .DaNam
                DaExt = .DaExt
                DaNaO = .DaNaO
                DaPfa = .DaPfa & "\"
                If LCase(DaPfa) <> LCase(GlImO) Then
                    IniSetVal "System", "ImpOrd", LCase(DaPfa)
                    GlImO = DaPfa
                End If

                Select Case LCase(DaExt)
                Case "pdf":
                    TypNa = "PDF-Dokument"
                    PfaNa = GlBPf
                Case "jpg":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "jpeg":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "png":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "psd":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "bmp":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "tif":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "tiff":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "gif":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "wmf":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "emf":
                    TypNa = "Bilddokument"
                    PfaNa = GlBPf
                Case "doc":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "docx":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "xlsx":
                    TypNa = "Excel-Dokument"
                    PfaNa = GlDox
                Case "xls":
                    TypNa = "Excel-Dokument"
                    PfaNa = GlDox
                Case "rtf":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "txm":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "txn":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "txr":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "txt":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "eml":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "msg":
                    TypNa = "Textdokument"
                    PfaNa = GlDox
                Case "xml"
                    TypNa = "X-Rechnung"
                    PfaNa = GlDox
                Case Else:
                    Set clFil = Nothing
                    SPopu "Ungültiger Dateityp", "Der von Ihnen ausgewählte Dateityp kann nicht geöffnet werden", IC48_Warning
                    Exit Sub
                End Select

                TeFil = PfaNa & DaNam
                If LCase(TeFil) <> LCase(FiNam) Then
                    NeuDa = BeStr & "_" & TmGui & "#_" & DaNaO & "." & LCase(DaExt)
                    NeuNa = PfaNa & NeuDa

                    If .FilVor(NeuNa) = True Then
                        .DaLoe = NeuNa & vbNullChar
                        .FilLoe
                    End If
                    .DaCop = FiNam & ";" & NeuNa & vbNullChar
                    If .FilCop(1) = False Then
                        SPopu "Datei schreibgeschützt", "Die Datei kann nicht kopiert werden", IC48_Warning
                    End If

                    TxFil.Text = NeuDa
                    If GlNeB = False Then 'neue Buchung
                        PuBu5.Enabled = False
                        S_Save
                    End If
                End If
            End With
        Else
            SPopu "Dateiname nicht lesbar", "Der von Ihnen ausgewählte Dateiname kann nicht gelesen werden weil er ggf. Sonderzeichen enthält", IC48_Warning
        End If
    End If
End If

Set clFil = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FOpn " & Err.Number
Resume Next

End Sub


Private Sub FSave()
On Error GoTo AnErr

Dim BuBet As Double
Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim BuTyp As Integer
Dim TabId As Integer
Dim RbBar As XtremeCommandBars.RibbonBar
Dim RbTab As XtremeCommandBars.RibbonTab
Dim CmBrs As XtremeCommandBars.CommandBars

Set FM = frmBuEdit
Set TxBet = FM.txtBuBet
Set PuBu5 = FM.cmdWeite
Set CmBTy = FM.cmbBuTyp
Set ChGut = FM.chkGutha

Set CmBrs = frmKatBU.comBar02
Set RbBar = CmBrs.Item(1)
Set RbTab = RbBar.SelectedTab

If Me.txtBuBet.Text = vbNullString Then
    SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
    Exit Sub
End If

If IsNumeric(Me.txtBuBet.Text) = False Then
    SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
    Exit Sub
End If

If CDbl(Me.txtBuBet.Text) <= 0 Then
    SPopu "Kein Buchungsbetrag", "Es wurde kein Buchungsbetrag eingegeben.", IC48_Forbidden
    Exit Sub
End If

If Me.txtBuBel.Text = vbNullString Then
    SPopu "Keine Buchungsnummer", "Es wurde keine Buchungsnummer eingegeben.", IC48_Forbidden
    Exit Sub
End If

If IsNumeric(Me.txtBuBel.Text) = False Then
    SPopu "Keine Buchungsnummer", "Es wurde keine Buchungsnummer eingegeben.", IC48_Forbidden
    Exit Sub
End If

If CDbl(Me.txtBuBel.Text) <= 0 Then
    SPopu "Keine Buchungsnummer", "Es wurde keine Buchungsnummer eingegeben.", IC48_Forbidden
    Exit Sub
End If

If Me.cmbBuTex.Text = vbNullString Then
    SPopu "Kein Buchungstext", "Es wurde kein Buchungstext eingegeben.", IC48_Forbidden
    Exit Sub
End If

If Me.txtKtoNr.Text = vbNullString Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If IsNumeric(Me.txtKtoNr.Text) = False Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If CLng(Me.txtKtoNr.Text) <= 0 Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If Me.txtKonto.Text = vbNullString Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If Me.txtBezei.Text = vbNullString Then
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        SPopu "Keine Sachkontennummer", "Es wurde kein Sachkontennummer eingegeben.", IC48_Forbidden
    Else
        SPopu "Keine Sollkentennummer", "Es wurde kein Sollkentennummer eingegeben.", IC48_Forbidden
    End If
    Exit Sub
End If

If GlBuc = False Then 'einfache Buchhaltung verwenden
    If Me.txtKtoHa.Text = vbNullString Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
    
    If IsNumeric(Me.txtKtoHa.Text) = False Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
    
    If CLng(Me.txtKtoHa.Text) <= 0 Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
        
    If Me.txtHaben.Text = vbNullString Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
    
    If Me.txtBezHa.Text = vbNullString Then
        SPopu "Keine Habenkontennummer", "Es wurde kein Habenkontennummer eingegeben.", IC48_Forbidden
        Exit Sub
    End If
End If

BuBet = Round(CDbl(Me.txtBuBet.Text), 2)

If GlBut = RibTab_HomeBanki Then
    TeTit = "Falscher Buchungsbetrag"
    TeMai = "Der Buchungsbetrag überschreitet die zulässige Gesamtsumme"
    TeInh = "Bei einer Split Buchung werden die Beträge aller getätigten Einzelbuchungen summiert und geprüft, ob diese den Umsatzbetrag überschreitet oder nicht."
    TeFus = "Die Summe aus der jetzt zu tätigen Buchung zuzüglich der bereits getätigten Einzelbuchungen zu diesem Buchungssplitt, übersteigt den Betrag des Umsatzes im Kontoauszug."
    If BaBet > 0 Then
        If GeBet > 0 Then
            If Round((GeBet - BuBet), 2) < 0 Then
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
                Me.txtBuBet.Text = Format$(GeBet, GlWa1)
                Exit Sub
            End If
        Else
            If Round((BaBet - BuBet), 2) < 0 Then
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, FM.hwnd
                Me.txtBuBet.Text = Format$(GeBet, GlWa1)
                Exit Sub
            End If
        End If
    End If
End If

If GlStB = False Then 'Stapelbuchung
    PuBu5.Enabled = False
End If

S_Save
DoEvents
If TabId = RibTab_Kat_KetBuc Then
    K_BuEi
End If
DoEvents

If ChGut.Value = xtpChecked Then
    If PatNr > 0 Then
        BuBet = TxBet.Text
        BuTyp = CmBTy.ItemData(CmBTy.ListIndex)
        Select Case BuTyp
        Case 1: S_PaGu PatNr, , BuBet
        Case 2: S_PaGu PatNr, BuBet
        End Select
    End If
End If

If GeBet > 0 Then
    GeBet = Round((GeBet - BuBet), 2)
Else
    GeBet = Round((BaBet - BuBet), 2)
End If

If GlNeB = False Then 'neue Buchung
    Unload Me
Else
    GlNeB = False 'neue Buchung
    If GlStB = False Then 'Stapelbuchung
        If BaBet > 0 Then
            If GeBet = 0 Then
                Unload Me
            Else
                FNeBu
            End If
        Else
            Unload Me
        End If
    Else
        If GlBuV = True Then 'Buchungsvorlage einfügen
            GlBuV = False
            Unload Me
        Else
            FNeBu
        End If
    End If
End If

Set RbBar = Nothing
Set CmBrs = Nothing

Exit Sub

AnErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FSave " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim ThIdx As Integer
Dim MiIdx As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim GesZa As Integer
Dim LauZa As Integer
Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmMdt As XtremeCommandBars.CommandBarComboBox
Dim CmMit As XtremeCommandBars.CommandBarComboBox
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmBuEdit
Set Rahm0 = FM.frmRahm0
Set Lbl05 = FM.lblLab05
Set Lbl06 = FM.lblLab06
Set CmWar = FM.cmbWarun
Set CmGeg = FM.cmbGegen
Set CmBuT = FM.cmbBuTex
Set CmBuS = FM.cmbBuStu
Set CmMan = FM.cmbManda
Set CmMta = FM.cmbMitar
Set CmBTy = FM.cmbBuTyp
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.cmdNeuBu
Set PuBu3 = FM.btnFile1
Set PuBu4 = FM.btnFile2
Set TxDa1 = FM.txtDatu1
Set TxBel = FM.txtBuBel
Set TxKto = FM.txtKonto
Set TxHab = FM.txtHaben
Set ChAsw = FM.chkGewEr
Set ChGut = FM.chkGutha
Set MoKal = FM.dtpDatu1
Set ImMan = frmMain.imgManag
Set CmBrs = frmMain.comBar01

Set CmCom = CmBrs.FindControl(CmCom, SY_BU_Buchung_SuchCombo, , True)
LiIdx = CmCom.ListIndex
Set CmMdt = CmBrs.FindControl(CmMdt, SY_SuMan, , True)
ThIdx = CmMdt.ListIndex
Set CmMit = CmBrs.FindControl(CmMit, SY_SuMit, , True)
MiIdx = CmMit.ListIndex

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
    .ToolTipText = "Markieren Sie bitte hier das gwünschte Buchungsdatum"
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

With CmBTy
    .AddItem "Ausgabe"
    .ItemData(0) = 1
    .AddItem "Einnahme"
    .ItemData(1) = 2
    .Enabled = GlNeB 'neue Buchung
End With

For AktZa = 1 To UBound(GlWar)
    CmWar.AddItem GlWar(AktZa, 1)
    CmWar.ItemData(AktZa - 1) = GlWar(AktZa, 0)
Next AktZa

With CmGeg
    If GlBuc = True Then 'einfache Buchhaltung verwenden
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            .AddItem GlGeK(AktZa, 3)
            .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
        Next AktZa
    Else
        For AktZa = 1 To UBound(GlGeK) 'Geldkonten
            For AktKo = 1 To UBound(GlSaK) 'Sachkonten mit Geldkontenzuordnung
                If GlGeK(AktZa, 0) = GlSaK(AktKo, 6) Then
                    .AddItem GlSaK(AktKo, 3)
                    .ItemData(AktZa - 1) = GlSaK(AktKo, 6) '[IDB]
                End If
            Next AktKo
        Next AktZa
        If .ListCount = 0 Then 'füge die Geldkonten aus der einfachen Buchführung hinzu
            For AktZa = 1 To UBound(GlGeK) 'Geldkonten
                .AddItem GlGeK(AktZa, 3)
                .ItemData(AktZa - 1) = GlGeK(AktZa, 0) '[IDB]
            Next AktZa
        End If
    End If
End With

For AktZa = 1 To UBound(GlBTe)
    CmBuT.AddItem GlBTe(AktZa, 1)
    CmBuT.ItemData(CmBuT.NewIndex) = GlBTe(AktZa, 0)
Next AktZa
CmBuT.AutoComplete = True

For AktZa = 1 To UBound(GlStu)
    CmBuS.AddItem GlStu(AktZa, 2)
    CmBuS.ItemData(CmBuS.NewIndex) = GlStu(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlThe)
    If GlNeB = True Then 'neue Buchung
        If CBool(GlThe(AktZa, 25)) = False Then
            LauZa = LauZa + 1
            CmMan.AddItem GlThe(AktZa, 13)
            CmMan.ItemData(CmMan.NewIndex) = GlThe(AktZa, 0)
        End If
    Else
        LauZa = LauZa + 1
        CmMan.AddItem GlThe(AktZa, 13)
        CmMan.ItemData(CmMan.NewIndex) = GlThe(AktZa, 0)
    End If
Next AktZa

If GlNeB = True Then 'neue Buchung
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        CmMta.AddItem GlMiA(AktZa, 1)
        CmMta.ItemData(CmMta.NewIndex) = GlMiA(AktZa, 2)
    Next AktZa

    If LiIdx = 7 Then 'Mandant
        CmMan.ListIndex = ThIdx - 1
    ElseIf LiIdx = 8 Then 'Mitarbeiter
        CmMta.ListIndex = MiIdx - 1
    Else
        CmMan.ListIndex = GlMan(GlSMa, 0) - 1
        CmMta.ListIndex = GlMiA(GlSmI, 0) - 1
    End If
Else
    For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
        CmMta.AddItem GlMiK(AktZa, 1)
        CmMta.ItemData(CmMta.NewIndex) = GlMiK(AktZa, 2)
    Next AktZa

    CmMan.ListIndex = GlMan(GlSMa, 0) - 1
    CmMta.ListIndex = GlMiK(GlSmI, 0) - 1
End If

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxBel
    .Pattern = "\d*"
    .SetMask "000000", "______"
    .Enabled = False
End With

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu3.Icon = ImMan.Icons.GetImage(IC16_Folder_Open, 16)
PuBu4.Icon = ImMan.Icons.GetImage(IC16_Folder_Del, 16)

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

If GlStB = False Then 'Stapelbuchung
    PuBu2.Enabled = False
End If

If GlBuc = True Then 'einfache Buchhaltung verwenden
    TxHab.Visible = False
    CmGeg.Visible = True
    Lbl05.Caption = "Sachkonto :"
    Lbl06.Caption = "Geldkonto :"
Else
    TxHab.Visible = True
    CmGeg.Visible = False
    Lbl05.Caption = "Soll-Konto :"
    Lbl06.Caption = "Haben-Konto :"
End If

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
ChAsw.BackColor = GlBak
ChGut.BackColor = GlBak

Set CmBrs = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub btnDatu1_Click()
    If FoLad = False Then
        FKale
    End If
End Sub

Private Sub btnFile1_Click()
    FOpn
End Sub

Private Sub btnFile2_Click()
    FLoe
End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50111)
TeMai = IniGetOpt("Hilfe", 50112)
TeInh = IniGetOpt("Hilfe", 50113)
TeFus = IniGetOpt("Hilfe", 50114)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    GlBuV = False 'Buchungsvorlage einfügen
    GlNeB = False 'neue Buchung
    Unload Me
End Sub

Private Sub chkGewEr_Click()
On Error Resume Next

Set FM = frmBuEdit

Set ChAsw = FM.chkGewEr
Set ChGut = FM.chkGutha

If ChAsw.Value = xtpChecked Then
    ChGut.Value = xtpUnchecked
End If

End Sub

Private Sub chkGutha_Click()
On Error Resume Next

Set FM = frmBuEdit

Set ChAsw = FM.chkGewEr
Set ChGut = FM.chkGutha

If ChGut.Value = xtpChecked Then
    ChAsw.Value = xtpUnchecked
End If

End Sub
Private Sub cmbBuTex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuStu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuStu_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.cmbBuTex.SetFocus
    Case vbKeyUp: 'Me.cmbBuTyp.SetFocus
    End Select
End Sub

Private Sub cmbBuTex_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF8: 'FSave
    Case vbKeyDown: 'Me.cmbGegen.SetFocus
    Case vbKeyUp: 'Me.txtKonto.SetFocus
    End Select
End Sub

Private Sub cmbBuTyp_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbBuTyp_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.cmbGegen.SetFocus
    Case vbKeyUp: 'Me.cmbBuTex.SetFocus
    End Select
End Sub
Private Sub cmbGegen_Click()
    If FoLad = False Then
        FBuBe
    End If
End Sub
Private Sub cmbGegen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbGegen_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.cmbBuStu.SetFocus
    Case vbKeyUp: 'Me.cmbBuTyp.SetFocus
    End Select
End Sub

Private Sub cmbManda_Click()
    If FoLad = False Then
        If GlNeB = True Then 'neue Buchung
            FBuBe
            FMand
        End If
    End If
End Sub

Private Sub cmbManda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbManda_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.txtBuKom.SetFocus
    Case vbKeyUp: 'Me.txtBuBel.SetFocus
    End Select
End Sub
Private Sub cmbMitar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbMitar_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.txtBuKom.SetFocus
    Case vbKeyUp: 'Me.txtBuBel.SetFocus
    End Select
End Sub
Private Sub cmbWarun_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbWarun_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.cmbBuTex.SetFocus
    Case vbKeyUp: 'Me.txtBuBet.SetFocus
    End Select
End Sub

Private Sub cmdNeuBu_Click()
    FNeBu
    Me.txtDatu1.SetFocus
End Sub
Private Sub cmdWeite_Click()
    FSave
End Sub
Private Sub dtpDatu1_SelectionChanged()
    FDatu
End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

FoLad = True

FInit

If GlBuV = True Then 'Buchungsvorlage einfügen
    FNeBu
Else
    If GlNeB = True Then 'neue Buchung
        FNeBu
    Else
        S_Posi
    End If
End If

FoLad = False

AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

Set FrmEx = Nothing

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBuEdit = Nothing
End Sub

Private Sub txtBezei_GotFocus()
    Me.txtBezei.SelStart = 0
    Me.txtBezei.SelLength = Len(Me.txtBezei.Text)
End Sub
Private Sub txtBezei_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtBezei_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBezei.SelLength = 0
    Case vbKeyDown: Me.cmbGegen.SetFocus
    Case vbKeyUp: Me.txtKonto.SetFocus
    End Select
End Sub
Private Sub txtBuBel_GotFocus()
On Error Resume Next

If FoLad = False Then
    If TxBel.Text <> vbNullString Then
        If IsNumeric(TxBel.Text) = True Then
            BuBel = TxBel.Text
        End If
    End If
End If

TxBel.SelStart = 0
TxBel.SelLength = Len(TxBel.Text)
    
End Sub
Private Sub txtBuBel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBuBel_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBuBel.SelLength = 0
    Case vbKeyF8: FSave
    Case vbKeyDown: Me.cmbManda.SetFocus
    Case vbKeyUp: Me.cmbGegen.SetFocus
    End Select
End Sub

Private Sub txtBuBel_Validate(Cancel As Boolean)
On Error Resume Next

Set FM = frmBuEdit
Set TxBel = FM.txtBuBel

If FoLad = False Then
    If TxBel.Text = vbNullString Then
        TxBel.Text = Format$(BuBel, "000000")
    ElseIf TxBel.isValid = False Then
        TxBel.Text = Format$(BuBel, "000000")
    Else
        FBuBp
    End If
End If
    
End Sub
Private Sub txtBuBet_GotFocus()
    Me.txtBuBet.SelStart = 0
    Me.txtBuBet.SelLength = Len(Me.txtBuBet.Text)
End Sub
Private Sub txtBuBet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBuBet_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBuBet.SelLength = 0
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.cmbWarun.SetFocus
    Case vbKeyUp: 'Me.txtDatu1.SetFocus
    End Select
End Sub
Private Sub txtBuBet_LostFocus()
On Error Resume Next

Dim Betra As Double

If Me.txtBuBet.Text <> vbNullString Then
    If IsNumeric(Me.txtBuBet.Text) = True Then
        Betra = CDbl(Me.txtBuBet.Text)
        If Betra < 0 Then
            Betra = Betra * (-1)
        End If
        If BaBet > 0 Then
            If Betra > BaBet Then
                Betra = BaBet
            End If
        End If
        Me.txtBuBet.Text = Format$(Betra, GlWa1)
    End If
End If

End Sub
Private Sub txtDatu1_GotFocus()
    Me.txtDatu1.SelStart = 0
    Me.txtDatu1.SelLength = Len(Me.txtDatu1.Text)
End Sub
Private Sub txtDatu1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtDatu1_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtDatu1.SelLength = 0
    Case vbKeyF8: FSave
    Case vbKeyDown: 'Me.txtBuBet.SetFocus
    Case vbKeyUp: 'Me.txtBezei.SetFocus
    End Select
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub
Private Sub txtBuKom_GotFocus()
    Me.txtBuKom.SelStart = 0
    Me.txtBuKom.SelLength = Len(Me.txtBuKom.Text)
End Sub
Private Sub txtBuKom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtBuKom_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtBuKom.SelLength = 0
    Case vbKeyF8: FSave
    Case vbKeyDown: Me.chkGewEr.SetFocus
    Case vbKeyUp: Me.cmbManda.SetFocus
    End Select
End Sub

Private Sub txtDatu1_Validate(Cancel As Boolean)
    If FoLad = False Then
        FBuBe
    End If
End Sub

Private Sub txtHaben_GotFocus()
    Me.txtHaben.SelStart = 0
    Me.txtHaben.SelLength = Len(Me.txtHaben.Text)
End Sub

Private Sub txtHaben_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn: KeyAscii = 0
    Case vbKeyTab: KeyAscii = 0
    End Select
End Sub
Private Sub txtHaben_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode
Case vbKeyF2:
        Me.txtKonto.SelLength = 0
Case vbKeyF8:
        FSave
Case vbKeyDown:
        Me.cmbBuStu.SetFocus
Case vbKeyUp:
        Me.cmbBuTyp.SetFocus
Case vbKeyReturn:
        GlBuF = 5 'Buchungsdialog
        S_KtSu "BuVo", KntRa
End Select
    
End Sub
Private Sub txtKonto_GotFocus()
    Me.txtKonto.SelStart = 0
    Me.txtKonto.SelLength = Len(Me.txtKonto.Text)
End Sub
Private Sub txtKonto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn: KeyAscii = 0
    Case vbKeyTab: KeyAscii = 0
    End Select
End Sub

Private Sub txtKonto_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode
Case vbKeyF2:
        Me.txtKonto.SelLength = 0
Case vbKeyF8:
        FSave
Case vbKeyDown:
        Me.cmbBuTex.SetFocus
Case vbKeyUp:
        Me.cmbWarun.SetFocus
Case vbKeyReturn:
        GlBuF = 1 'Buchungsdialog
        S_KtSu "BuVo", KntRa
End Select

End Sub
Private Sub txtRecNr_GotFocus()
    Me.txtRecNr.SelStart = 0
    Me.txtRecNr.SelLength = Len(Me.txtRecNr.Text)
End Sub
Private Sub txtRecNr_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtRecNr_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtRecNr.SelLength = 0
    Case vbKeyF8: FSave
    Case vbKeyDown:
    Case vbKeyUp:
    End Select
End Sub
Private Sub updCont1_DownClick()

Dim AltDa As Date

If FoLad = False Then
    Set TxDa1 = Me.txtDatu1
    
    AltDa = TxDa1.Text
    
    TxDa1.Text = DateAdd("d", -1, AltDa)
        
    FBuBe
End If

End Sub
Private Sub updCont1_UpClick()

Dim AltDa As Date

If FoLad = False Then
    Set TxDa1 = Me.txtDatu1
    
    AltDa = TxDa1.Text
    
    TxDa1.Text = DateAdd("d", 1, AltDa)
    
    FBuBe
End If

End Sub
