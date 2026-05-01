VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#16.3#0"; "Codejock.ReportControl.v16.3.1.ocx"
Begin VB.Form frmTerKop 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Terminassistent"
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
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   1
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   24
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
         TabIndex        =   25
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
      Left            =   0
      TabIndex        =   2
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
         TabIndex        =   9
         Top             =   2500
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "mehrere Tage markieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optTeKop 
         Height          =   220
         Left            =   2300
         TabIndex        =   8
         Top             =   1700
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Den markierten Termin kopieren"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optTeNeu 
         Height          =   220
         Left            =   2300
         TabIndex        =   7
         Top             =   1300
         Width           =   3000
         _Version        =   1048579
         _ExtentX        =   5292
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Einen neuen Termin kopieren"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin VB.Label lblLab01 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTerKop.frx":0000
         Height          =   400
         Left            =   1000
         TabIndex        =   6
         Top             =   200
         Width           =   6100
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
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   5997
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ComboBox cmbGanzt 
         Height          =   315
         Left            =   940
         TabIndex        =   30
         Top             =   3500
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbAbger 
         Height          =   315
         Left            =   4060
         TabIndex        =   31
         Tag             =   "0Aufgabe"
         Top             =   3500
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.UpDown updCont3 
         Height          =   350
         Left            =   6510
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1300
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBisZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.UpDown updCont2 
         Height          =   350
         Left            =   5120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1300
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   600
         _StockProps     =   64
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtVonZe"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.ComboBox txtBetre 
         Height          =   310
         Left            =   1100
         TabIndex        =   10
         Tag             =   "0IDKurz"
         Top             =   800
         Width           =   6000
         _Version        =   1048579
         _ExtentX        =   10583
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         MaxLength       =   250
         AutoComplete    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbStatu 
         Height          =   310
         Left            =   1100
         TabIndex        =   11
         Tag             =   "0Farbtyp"
         Top             =   1300
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3545
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox4"
      End
      Begin XtremeSuiteControls.ComboBox cmbRaum1 
         Height          =   315
         Left            =   4200
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "0IDR"
         Top             =   1800
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox6"
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   315
         Left            =   4200
         TabIndex        =   19
         Tag             =   "0IDP"
         Top             =   2300
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox7"
      End
      Begin XtremeSuiteControls.FlatEdit txtVonZe 
         Height          =   350
         Left            =   4200
         TabIndex        =   12
         Tag             =   "0ZeiVon"
         Top             =   1300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBisZe 
         Height          =   350
         Left            =   5600
         TabIndex        =   14
         Tag             =   "0ZeiBis"
         Top             =   1300
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbTeTyp 
         Height          =   315
         Left            =   1100
         TabIndex        =   18
         Tag             =   "0TerTyp"
         Top             =   2300
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   4200
         TabIndex        =   20
         Tag             =   "0IDM"
         Top             =   2800
         Width           =   2900
         _Version        =   1048579
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox7"
      End
      Begin XtremeSuiteControls.ComboBox cmbNotVa 
         Height          =   315
         Left            =   1100
         TabIndex        =   16
         Tag             =   "0NotifyValue"
         Top             =   1800
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3545
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtNoDat 
         Height          =   350
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "0NotifySetDate"
         Top             =   3500
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtNoTim 
         Height          =   350
         Left            =   1400
         TabIndex        =   43
         TabStop         =   0   'False
         Tag             =   "0NotifySetTime"
         Top             =   3500
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtFarbe 
         Height          =   195
         Left            =   600
         TabIndex        =   44
         TabStop         =   0   'False
         Tag             =   "0Farbe"
         Top             =   4000
         Width           =   195
         _Version        =   1048579
         _ExtentX        =   353
         _ExtentY        =   353
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Appearance      =   6
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbPrior 
         Height          =   315
         Left            =   1100
         TabIndex        =   21
         Tag             =   "0Priorität"
         Top             =   2800
         Width           =   1995
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox5"
      End
      Begin VB.Label lblLab03 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTerKop.frx":00A3
         Height          =   400
         Left            =   1000
         TabIndex        =   46
         Top             =   200
         Width           =   6100
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   240
         Left            =   140
         TabIndex        =   45
         Top             =   2850
         Width           =   920
         _Version        =   1048579
         _ExtentX        =   1623
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Priorität :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab02 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Betreff :"
         Height          =   240
         Left            =   140
         TabIndex        =   41
         Top             =   850
         Width           =   920
      End
      Begin VB.Label lblLab10 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Mitarbeiter :"
         Height          =   240
         Left            =   3200
         TabIndex        =   40
         Top             =   2850
         Width           =   930
      End
      Begin VB.Label lblLab15 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   240
         Left            =   140
         TabIndex        =   39
         Top             =   1350
         Width           =   920
      End
      Begin VB.Label lblLab14 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Marker :"
         Height          =   240
         Left            =   140
         TabIndex        =   38
         Top             =   2350
         Width           =   920
      End
      Begin XtremeSuiteControls.Label lblLab18 
         Height          =   240
         Left            =   3200
         TabIndex        =   37
         Top             =   2350
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Mandant :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   240
         Left            =   3200
         TabIndex        =   36
         Top             =   1850
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Raumplan :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab21 
         Height          =   240
         Left            =   140
         TabIndex        =   35
         Top             =   1850
         Width           =   920
         _Version        =   1048579
         _ExtentX        =   1614
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Emailerinn.: "
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab23 
         Height          =   240
         Left            =   3200
         TabIndex        =   34
         Top             =   1350
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Terminzeit :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblLab09 
         Alignment       =   1  'Rechts
         BackStyle       =   0  'Transparent
         Caption         =   "Ganztägig :"
         Height          =   240
         Left            =   0
         TabIndex        =   33
         Top             =   3550
         Width           =   920
      End
      Begin XtremeSuiteControls.Label lblLab22 
         Height          =   240
         Left            =   3060
         TabIndex        =   32
         Top             =   3550
         Width           =   930
         _Version        =   1048579
         _ExtentX        =   1640
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Leistungen :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   3400
      Left            =   0
      TabIndex        =   5
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
         TabIndex        =   28
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
         Caption         =   "Bitte markieren Sie die Tage, an denen der Termin eingefügt werden soll."
         Height          =   400
         Left            =   1000
         TabIndex        =   29
         Top             =   200
         Width           =   6100
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   3400
      Left            =   0
      TabIndex        =   4
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
         Height          =   2400
         Left            =   60
         TabIndex        =   26
         Top             =   800
         Width           =   7840
         _Version        =   1048579
         _ExtentX        =   13829
         _ExtentY        =   4233
         _StockProps     =   64
         FreezeColumnsAbs=   0   'False
      End
      Begin VB.Label lblLabl5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTerKop.frx":012C
         Height          =   400
         Left            =   1000
         TabIndex        =   27
         Top             =   200
         Width           =   6100
      End
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu2 
      Height          =   400
      Left            =   0
      TabIndex        =   47
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
End
Attribute VB_Name = "frmTerKop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private VoZei As XtremeSuiteControls.FlatEdit
Private BiZei As XtremeSuiteControls.FlatEdit
Private TxNoD As XtremeSuiteControls.FlatEdit
Private TxNoZ As XtremeSuiteControls.FlatEdit
Private TxFar As XtremeSuiteControls.FlatEdit
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private ChExp As XtremeSuiteControls.CheckBox
Private CmBet As XtremeSuiteControls.ComboBox
Private CmMar As XtremeSuiteControls.ComboBox
Private CmRmu As XtremeSuiteControls.ComboBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmNot As XtremeSuiteControls.ComboBox
Private CmPri As XtremeSuiteControls.ComboBox
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private DaPi1 As XtremeCalendarControl.DatePicker
Private DaPi2 As XtremeCalendarControl.DatePicker
Private RpRow As XtremeReportControl.ReportRow
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpCol As XtremeReportControl.ReportColumn
Private RpRec As XtremeReportControl.ReportRecord
Private RpRcs As XtremeReportControl.ReportRecords
Private RpItm As XtremeReportControl.ReportRecordItem
Private ImMan As XtremeCommandBars.ImageManager

Public Urlaub As Boolean

Private VoDat() As Date

Private RetWe As Long
Private AbExp As Boolean

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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub FBetr()
On Error GoTo LiErr

Dim FarWe As Long
Dim TmVon As Date
Dim TmBis As Date
Dim RmuNr As Long
Dim IdxNr As Long
Dim MitNr As Long
Dim ManNr As Long
Dim TmStr As String
Dim RmIdx As Integer
Dim MiIdx As Integer
Dim MaIdx As Integer
Dim AktZa As Integer
Dim ZeiVo As Integer
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim DaPro As XtremeCalendarControl.CalendarDataProvider
Dim CaLbs As XtremeCalendarControl.CalendarEventLabels
Dim CaLbl As XtremeCalendarControl.CalendarEventLabel

Set TxFar = Me.txtFarbe
Set CmBet = Me.txtBetre
Set CmRmu = Me.cmbRaum1
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set PuBu1 = Me.btnWeiter
Set PuBu3 = Me.btnZurück

Set FM = frmMain
Set CaCol = FM.calCont1
Set DaPro = CaCol.DataProvider
Set CaLbs = DaPro.LabelList

RmIdx = 0
MiIdx = 0
IdxNr = CmBet.ListIndex + 1

If CmBet.Text <> vbNullString Then
    TmStr = CmBet.Text
End If

If IdxNr > 0 Then
    If LCase(GlBtr(IdxNr, 1)) = LCase(TmStr) Then
        RmuNr = GlBtr(IdxNr, 2)
        ZeiVo = GlBtr(IdxNr, 3)

        If GlBtr(IdxNr, 4) <> vbNullString Then
            If GlBtr(IdxNr, 4) > 0 Then
                FarWe = GlBtr(IdxNr, 4)
            Else
                FarWe = vbWhite
            End If
        Else
            FarWe = vbWhite
        End If
        
        If GlBtr(IdxNr, 5) <> vbNullString Then
            MitNr = GlBtr(IdxNr, 5)
        Else
            MitNr = 0
        End If
    
        If RmuNr > 0 Then
            For AktZa = 1 To UBound(GlRmu)
                If RmuNr = GlRmu(AktZa, 2) Then
                    RmIdx = AktZa - 1
                    CmRmu.ListIndex = RmIdx
                    Exit For
                End If
            Next AktZa
        End If
    
        If FarWe > 0 Then 'Farbwert aus Terminbetreffs
            For Each CaLbl In CaLbs
                If CaLbl.Color = FarWe Then
                    TxFar.Text = CaLbl.LabelID
                    CmBet.BackColor = FarWe
                    Exit For
                End If
            Next CaLbl
        End If
        
        If MitNr > 0 Then
            For AktZa = 1 To UBound(GlMiT)
                If MitNr = GlMiT(AktZa, 2) Then
                    ManNr = GlMiT(AktZa, 7)
                    MiIdx = AktZa - 1
                    CmMit.ListIndex = MiIdx
                    Exit For
                End If
            Next AktZa
    
            For AktZa = 1 To UBound(GlMaT)
                If ManNr = GlMaT(AktZa, 2) Then
                    MaIdx = AktZa - 1
                    CmMan.ListIndex = MaIdx
                    Exit For
                End If
            Next AktZa
        End If
    Else
        FarWe = vbWhite
        ZeiVo = 0
        If FarWe > 0 Then 'Farbwert aus Terminbetreffs
            For Each CaLbl In CaLbs
                If CaLbl.Color = FarWe Then
                    TxFar.Text = CaLbl.LabelID
                    CmBet.BackColor = FarWe
                    Exit For
                End If
            Next CaLbl
        End If
    End If
Else
    'FarWe = vbWhite
    ZeiVo = 0
    If FarWe > 0 Then 'Farbwert aus Terminbetreffs
        For Each CaLbl In CaLbs
            If CaLbl.Color = FarWe Then
                TxFar.Text = CaLbl.LabelID
                CmBet.BackColor = FarWe
                Exit For
            End If
        Next CaLbl
    End If
End If

If ZeiVo > 0 Then
    If VoZei.Text <> vbNullString Then
        TmVon = TimeValue(VoZei.Text)
        TmBis = DateAdd("n", ZeiVo, TmVon)
        BiZei.Text = Format$(TmBis, "hh:mm")
    End If
End If

If PuBu1.Enabled = False Then
    PuBu1.Enabled = True
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FBetr " & Err.Number
Resume Next

End Sub
Private Sub FDatu()
On Error GoTo OrErr

Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim AnzTa As Long
Dim AnzBl As Long
Dim AktBl As Long
Dim AktTa As Long
Dim BloTa As Long

Set DaPi1 = Me.dtpDatu1
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ

AktTa = 0
AnzBl = DaPi1.Selection.BlocksCount

If AnzBl = 0 Then
    Set DaPi1 = Nothing
    Exit Sub
ElseIf AnzBl = 1 Then
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

If PuBu1.Enabled = False Then
    PuBu1.Enabled = True
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
Private Sub FInit()
On Error GoTo SuErr

Dim AbExp As Boolean
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn

Set FM = frmMain
Set ImMan = FM.imgManag
Set RpCo1 = Me.repCont1
Set ChExp = Me.chkExpMo
Set DaPi1 = Me.dtpDatu1
Set DaPi2 = Me.dtpDatu2
Set Opti1 = Me.optTeNeu
Set Opti2 = Me.optTeKop
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück
Set RpCls = RpCo1.Columns

AbExp = CBool(IniGetVal("System", "KopExp"))

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

With DaPi1
    .AllowNoncontinuousSelection = True
    .AskDayMetrics = True
    .AutoSizeRowCol = True
    .BorderStyle = 0
    .Enabled = True
    .FirstDayOfWeek = 2
    .FirstWeekOfYearDays = 4
    .HighlightToday = True
    If GlBut = RibTab_Rechnungen Then
        .MaxSelectionCount = 25
    Else
        .MaxSelectionCount = GlMxK 'Maximal slektierbare Kalendertage
    End If
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

With DaPi2
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

Set DaPi1 = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FKrCa(ByVal CoIdx As Long)
On Error GoTo PoErr
'Änderungen im Krankneblatt

Dim ZeSta As Date
Dim ZeEnd As Date
Dim AdMin As Long
Dim RmuNr As Long
Dim NeuDa As Date
Dim StaZe As Date
Dim EndZe As Date
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTerKop
Set VoZei = FM.txtVonZe
Set BiZei = FM.txtBisZe
Set CmRmu = FM.cmbRaum1
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

ZeSta = TimeValue(VoZei.Text)
ZeEnd = TimeValue(BiZei.Text)
AdMin = DateDiff("n", TimeValue(ZeSta), TimeValue(ZeEnd))

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Select Case CoIdx
        Case 2:
            Set RpCol = RpCls.Find(2)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = True Then
                NeuDa = CDate(RpRow.Record(RpCol.ItemIndex).Value)
            Else
                NeuDa = Date
            End If
            RpRow.Record(RpCol.ItemIndex).Value = NeuDa
            Set RpCol = RpCls.Find(1)
            RpRow.Record(RpCol.ItemIndex).Value = Format$(NeuDa, "dddd")
        Case 3:
            Set RpCol = RpCls.Find(3)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = False Then
                StaZe = "08:00"
                RpRow.Record(RpCol.ItemIndex).Value = "08:00"
            Else
                StaZe = RpRow.Record(RpCol.ItemIndex).Value
            End If
            Set RpCol = RpCls.Find(4)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = False Then
                EndZe = "20:00"
                RpRow.Record(RpCol.ItemIndex).Value = "20:00"
            Else
                EndZe = RpRow.Record(RpCol.ItemIndex).Value
            End If
            If StaZe >= EndZe Then
                Set RpCol = RpCls.Find(3)
                RpRow.Record(RpCol.ItemIndex).Value = "08:00"
                Set RpCol = RpCls.Find(4)
                RpRow.Record(RpCol.ItemIndex).Value = "20:00"
            End If
        Case 4:
            Set RpCol = RpCls.Find(3)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = False Then
                StaZe = "08:00"
                RpRow.Record(RpCol.ItemIndex).Value = "08:00"
            Else
                StaZe = RpRow.Record(RpCol.ItemIndex).Value
            End If
            Set RpCol = RpCls.Find(4)
            If IsDate(RpRow.Record(RpCol.ItemIndex).Value) = False Then
                EndZe = "20:00"
                RpRow.Record(RpCol.ItemIndex).Value = "20:00"
            Else
                EndZe = RpRow.Record(RpCol.ItemIndex).Value
            End If
            If StaZe >= EndZe Then
                Set RpCol = RpCls.Find(3)
                RpRow.Record(RpCol.ItemIndex).Value = "08:00"
                Set RpCol = RpCls.Find(4)
                RpRow.Record(RpCol.ItemIndex).Value = "20:00"
            End If
        Case 7:
            Set RpCol = RpCls.Find(7)
            If RpRow.Record(RpCol.ItemIndex).Value = vbNullString Then
                RpRow.Record(RpCol.ItemIndex).Value = "kein Betreff"
            End If
        End Select
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKrCa " & Err.Number
Resume Next

End Sub
Private Sub FKaRo(ByVal RpBut As XtremeReportControl.IReportInplaceButton)
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim ItmLi As Long
Dim ItmOb As Long
Dim ItmRe As Long
Dim ItmHo As Long
Dim ItmBr As Long
Dim ItmTo As Long
Dim RmuNr As Long
Dim AltDa As Date
Dim NeuDa As Date
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTerKop
Set RpCo1 = FM.repCont1
Set RpCls = RpCo1.Columns
Set RpSel = RpCo1.SelectedRows

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        If RpRow.Record(2).Value <> vbNullString Then
            AltDa = CDate(RpRow.Record(2).Value)
        Else
            AltDa = Date
        End If
        RpBut.GetRect ItmLi, ItmOb, ItmRe, ItmHo
        ItmBr = ItmRe
        ItmTo = ItmHo + 1
        NeuDa = FKaSh(ItmBr, ItmTo, AltDa, RpCo1.hwnd, True)
        If IsDate(NeuDa) Then
            RpCo1.EditItem Nothing, Nothing
            RpRow.Record(1).Value = Format$(NeuDa, "dddd")
            RpRow.Record(2).Value = CDate(NeuDa)
            RpCo1.Populate
        End If
    End If
End If

Set RpCls = Nothing
Set RpSel = Nothing
Set RpCo1 = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaRo " & Err.Number
Resume Next

End Sub
Private Function FKaSh(ByVal KalLi As Long, ByVal KalOb As Long, ByVal NeuDa As Date, ByVal mHwnd As Long, Optional ByVal Flag As Boolean = False) As Date
On Error GoTo LaErr

Dim RpCo6 As XtremeReportControl.ReportControl

Dim Datu1 As Date
Dim DayFi As Date
Dim DayLa As Date
Dim KaBre As Long
Dim KaHoh As Long
Dim RetWe As Boolean

Set FM = frmTermVo
Set DaPi2 = FM.dtpDatu2

DayFi = NeuDa - 30
DayLa = NeuDa + 30

With DaPi2
    .RedrawControl
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    .GetMinReqRect KaBre, KaHoh, 1, 1
    If Flag = True Then
        RetWe = .ShowModalEx(KalLi - KaBre - 4 - 4, KalOb, KaBre + 4, KaHoh + 4, mHwnd)
    Else
        RetWe = .ShowModalEx(-1, 20, KaBre + 4, KaHoh + 4, mHwnd)
    End If
End With

If RetWe = True Then
    If DaPi2.Selection.BlocksCount > 0 Then
        Datu1 = DaPi2.Selection.Blocks(0).DateBegin()
        If IsDate(Datu1) Then
            FKaSh = CDate(Datu1)
        Else
            FKaSh = NeuDa
        End If
    Else
        FKaSh = NeuDa
    End If
Else
    FKaSh = NeuDa
End If

Set DaPi2 = Nothing

Exit Function

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKaSh " & Err.Number
Resume Next

End Function

Private Sub FLoad()
On Error GoTo SuErr

Dim DayFi As Date
Dim DayLa As Date
Dim StaDa As Date
Dim StaZe As Date
Dim EndZe As Date
Dim RetWe As Long
Dim GesZa As Long
Dim MitNr As Long
Dim ManNr As Long
Dim ZeSta As String
Dim ZeEnd As String
Dim AkDat As String
Dim NotDa As String
Dim NotZe As String
Dim NotSt As String
Dim ZeiUm As Boolean
Dim LiLin As Boolean
Dim LiKop As Boolean
Dim SelDf As Integer
Dim AktPo As Integer
Dim AktZa As Integer
Dim mAnza As Integer
Dim MiIdx As Integer
Dim MaIdx As Integer
Dim ZeiRa As Integer
Dim NotVa As Integer
Dim AdMin As Integer
Dim AnzPo As Integer
Dim MitOK As Boolean
Dim ManOK As Boolean
Dim DayGa As Boolean
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set CaCol = FM.calCont1
Set ImMan = FM.imgManag
Set RpCon = Me.repCont1
Set ChExp = Me.chkExpMo
Set DaPi1 = Me.dtpDatu1
Set Opti1 = Me.optTeNeu
Set Opti2 = Me.optTeKop
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set TxNoD = Me.txtNoDat
Set TxNoZ = Me.txtNoTim
Set CmBet = Me.txtBetre
Set CmMar = Me.cmbTeTyp
Set CmTyp = Me.cmbStatu
Set CmRmu = Me.cmbRaum1
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set CmNot = Me.cmbNotVa
Set CmPri = Me.cmbPrior
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück

AkDat = Date
        
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

With CmPri
    .AddItem "Hoch"
    .ItemData(0) = 1
    .AddItem "Normal"
    .ItemData(1) = 2
    .AddItem "Niedrig"
    .ItemData(2) = 3
End With

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

With CmMit
    .AddItem "Alle Mitarbeiter"
    .ItemData(AktZa) = 0
End With

With CmNot
    For AktZa = 0 To 48
        .AddItem AktZa & " Std."
        .ItemData(AktZa) = AktZa
    Next AktZa
End With

VoZei.SetMask "00:00", "__:__"
BiZei.SetMask "00:00", "__:__"

CmRmu.ListIndex = GlTRx

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
        If GlTBn = GlMiT(AktZa, 2) Then 'Termin Behandlernummer
            GlTBx = AktZa - 1 'Termin Behandlerindex
            MitOK = True
            Exit For
        End If
    Next AktZa
    If MitOK = True Then
        CmMit.ListIndex = GlTBx
        MiIdx = GlTBx + 1
    Else
        CmMit.ListIndex = 0
        MiIdx = 1
    End If
    If GlMiT(MiIdx, 8) > 0 Then
        ZeiRa = GlMiT(MiIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
    NotVa = GlMiT(MiIdx, 39)
    If NotVa = 0 Then
        NotVa = 24
    End If
Else
    For AktZa = 1 To UBound(GlMaT) 'Aktive Mandanten + Terminspalte
        If GlTBn = GlMaT(AktZa, 2) Then 'Termin Behandlernummer
            GlTBx = AktZa - 1 'Termin Behandlerindex
            ManOK = True
            Exit For
        End If
    Next AktZa
    If ManOK = True Then
        CmMan.ListIndex = GlTBx
        MaIdx = GlTBx + 1
    Else
        CmMan.ListIndex = 0
        MaIdx = 1
    End If
    If GlMaT(MaIdx, 8) > 0 Then
        ZeiRa = GlMaT(MaIdx, 8)
    Else
        ZeiRa = GlZeR 'Zeitrasterindex
    End If
    NotVa = GlMaT(MaIdx, 25)
    If NotVa = 0 Then
        NotVa = 24
    End If
End If

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
Else
    If CmMit.ListCount > 1 Then
        CmMit.ListIndex = GlSmI - 1
    Else
        CmMit.ListIndex = GlSmI - 1
    End If
End If

DoEvents
SRast ZeiRa
DoEvents

With CmNot
    .ListIndex = NotVa
    .Enabled = GlTeE 'Email-Termin-Erinnerung
End With

If GlBut = RibTab_Rechnungen Then
    Set RpCo7 = frmKatRC.repCont7
    Set RpSel = RpCo7.SelectedRows
    AnzPo = RpSel.Count
    VoZei.Text = "08:00"
    BiZei.Text = "20:00"
    Opti2.Value = True
    Opti1.Enabled = False
Else
    Set ViEvs = CaCol.ActiveView.GetSelectedEvents
    AnzPo = ViEvs.Count
    If AnzPo > 0 Then 'Kopieren
        Opti2.Value = True
    End If
    
    Set CaHit = CaCol.ActiveView.HitTest
    If CaHit.ViewEvent Is Nothing Then
        CaCol.ActiveView.GetSelection DayFi, DayLa, DayGa
        With GlSel
            .DaSta = DayFi
            .DaEnd = DayLa
            .DaGes = DayGa
        End With
    Else
        With GlSel
            .DaSta = CaHit.ViewEvent.Event.StartTime
            .DaEnd = CaHit.ViewEvent.Event.EndTime
            .DaGes = CaHit.ViewEvent.Event.AllDayEvent
        End With
    End If
    
    If AnzPo > 0 Then 'Kopieren
        For Each ViEvt In ViEvs
            If ViEvt.Selected = True Then
                Set CaEvt = ViEvt.Event
                StaDa = DateValue(CaEvt.StartTime)
                StaZe = TimeValue(CaEvt.StartTime)
                EndZe = TimeValue(CaEvt.EndTime)
                Exit For
            End If
        Next ViEvt
    End If
    
    If GlSel.DaSta > 0 Then
        If Format$(GlSel.DaSta, "hh:mm") = "00:00" Then
            If AnzPo > 0 Then 'Kopieren
                VoZei.Text = Format$(StaZe, "hh:mm")
                BiZei.Text = Format$(EndZe, "hh:mm")
            Else
                VoZei.Text = "08:00"
                If Urlaub = True Then
                    BiZei.Text = "20:00"
                Else
                    BiZei.Text = "09:00"
                End If
            End If
        Else
            If AnzPo > 0 Then 'Kopieren
                VoZei.Text = Format$(StaZe, "hh:mm")
                BiZei.Text = Format$(EndZe, "hh:mm")
            Else
                ZeSta = Format$(GlSel.DaSta, "hh:mm")
                ZeEnd = Format$(GlSel.DaEnd, "hh:mm")
                AdMin = DateDiff("n", TimeValue(ZeSta), TimeValue(ZeEnd))
                For AktZa = 1 To UBound(GlRas) 'Zeitrasterstartzeiten
                    If TimeValue(GlRas(AktZa)) <= TimeValue(ZeSta) Then
                        StaZe = TimeValue(GlRas(AktZa))
                    End If
                    If TimeValue(GlRas(AktZa)) <= TimeValue(ZeEnd) Then
                        EndZe = TimeValue(GlRas(AktZa))
                    End If
                Next AktZa
                If Urlaub = True Then
                    VoZei.Text = "08:00"
                    BiZei.Text = "20:00"
                Else
                    VoZei.Text = Format$(StaZe, "hh:mm")
                    BiZei.Text = Format$(EndZe, "hh:mm")
                End If
            End If
        End If
    Else
       If AnzPo > 0 Then 'Kopieren
            VoZei.Text = Format$(StaZe, "hh:mm")
            BiZei.Text = Format$(EndZe, "hh:mm")
        Else
            VoZei.Text = "08:00"
            If Urlaub = True Then
                BiZei.Text = "20:00"
            Else
                BiZei.Text = "09:00"
            End If
        End If
    End If
    
    NotDa = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "dd.mm.yyyy")
    NotZe = Format$(CDate(DateAdd("h", -NotVa, AkDat & " " & StaZe)), "hh:mm")
    NotSt = NotDa & Chr$(32) & NotZe
    
    TxNoD.Text = NotDa
    TxNoZ.Text = NotZe
            
    If Urlaub = True Then
        CmBet.Text = "Urlaub"
        CmMar.Text = "Urlaub"
    End If
            
    RetWe = SendMessage(CmTyp.hwnd, CB_SETCURSEL, 2, ByVal 0&)
    RetWe = SendMessage(CmPri.hwnd, CB_SETCURSEL, 1, ByVal 0&)
            
    If Urlaub = True Then
        DoEvents
        Rahm1.Visible = False
        Rahm2.Visible = True
        Rahm3.Visible = False
        Rahm4.Visible = False
    End If
End If

If AbExp = True Then
    ChExp.Value = 1
End If

Set DaPi1 = Nothing
Set RpCon = Nothing
Set RpCo7 = Nothing
Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

End Sub
Private Sub FMita()
On Error GoTo LiErr

Dim MitNr As Long
Dim ManNr As Long
Dim MaIdx As Integer
Dim AktZa As Integer

Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMan.ListCount > 1 Then
        MitNr = CmMit.ItemData(CmMit.ListIndex)
        For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter
            If MitNr = GlMiT(AktZa, 2) Then
                ManNr = GlMiT(AktZa, 7) 'zugeordnete Mandantennummer
                Exit For
            End If
        Next AktZa
        For AktZa = 1 To UBound(GlMaT) 'Aktive Mitarbeiter
            If ManNr = GlMaT(AktZa, 2) Then
                MaIdx = AktZa - 1
                CmMan.ListIndex = MaIdx
                Exit For
            End If
        Next AktZa
    End If
End If

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FMita " & Err.Number
Resume Next

End Sub
Private Sub FNoVa()
On Error GoTo LiErr
'Ändert dem Notification Wert

Dim ZeiSt As String
Dim NotDa As String
Dim NotZe As String
Dim NotSt As String
Dim NotVa As Integer
Dim MaIdx As Integer
Dim MiIdx As Integer

Set VoZei = Me.txtVonZe
Set CmMit = Me.cmbMitar
Set CmMan = Me.cmbBehan
Set CmNot = Me.cmbNotVa
Set TxNoD = Me.txtNoDat
Set TxNoZ = Me.txtNoTim

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If MiIdx <= UBound(GlMiT) Then
        If GlMiT(MiIdx, 39) > 0 Then
            NotVa = GlMiT(MiIdx, 39)
        Else
            NotVa = 24
        End If
    Else
        NotVa = GlMiT(GlSmI, 39)
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If MaIdx <= UBound(GlMaT) Then
        If GlMaT(MaIdx, 25) > 0 Then
            NotVa = GlMaT(MaIdx, 25)
        Else
            NotVa = 24
        End If
    Else
        NotVa = GlMaT(GlSMa, 25)
    End If
End If

CmNot.ListIndex = NotVa

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNoVa " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FSpal()
On Error GoTo SpErr
'Formratieren der Spalten

Dim AktZa As Integer
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmTerKop
Set RpCo1 = FM.repCont1

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    If .Columns.Count > 0 Then .Columns.DeleteAll
    .Populate
End With

If GlBut = RibTab_Rechnungen Then
    Set RpCls = RpCo1.Columns
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
        Set RpCol = .Add(3, vbNullString, 0, False)
        With RpCol
            .Alignment = xtpAlignmentIconCenter
            .HeaderAlignment = xtpAlignmentCenter
            .Icon = IC16_Calendar_Disk
        End With
        Set RpCol = .Add(4, "Mandant", 0, True)
        With RpCol
            .EditOptions.AddComboButton
            .EditOptions.AllowEdit = True
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleReadOnly
            For AktZa = 1 To UBound(GlMan)
                .EditOptions.Constraints.Add GlMan(AktZa, 1), GlMan(AktZa, 2)
            Next AktZa
            .EditOptions.Constraints.Add "Alle Mandanten", 0
        End With
        Set RpCol = .Add(5, "Mitarbeiter", 0, True)
        With RpCol
            .EditOptions.AddComboButton
            .EditOptions.AllowEdit = True
            .EditOptions.ConstraintEdit = True
            .EditOptions.EditControlStyle = xtpEditStyleReadOnly
            For AktZa = 1 To UBound(GlMiK)
                .EditOptions.Constraints.Add GlMiK(AktZa, 1), GlMiK(AktZa, 2)
            Next AktZa
            .EditOptions.Constraints.Add "Alle Mitarbeiter", 0
        End With
        Set RpCol = .Add(6, "Nummer", 0, True)
        RpCol.HeaderAlignment = xtpAlignmentCenter
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
    RpCls(2).Width = 80
    RpCls(2).Editable = True
    RpCls(3).Width = 40
    RpCls(3).Editable = True
    RpCls(4).Width = 160
    RpCls(4).Editable = True
    RpCls(5).Width = 140
    RpCls(6).Width = 80
    
Else

    Set RpCls = RpCo1.Columns
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
        RpCol.EditOptions.MaxLength = 250
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
                .EditOptions.Constraints.Add "Alle Mitarbeiter", 0
            Else
                For AktZa = 1 To UBound(GlMan)
                    .EditOptions.Constraints.Add GlMan(AktZa, 1), GlMan(AktZa, 2)
                Next AktZa
                .EditOptions.Constraints.Add "Alle Mandanten", 0
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
    RpCls(2).Width = 80
    RpCls(2).Editable = True
    RpCls(3).Width = 50
    RpCls(3).Editable = True
    RpCls(4).Width = 50
    RpCls(4).Editable = True
    RpCls(5).Width = 40
    RpCls(5).Editable = True
    RpCls(6).Width = 0
    RpCls(7).Width = 160
    RpCls(7).Editable = True
    RpCls(8).Width = 140
    RpCls(9).Width = 90
    RpCls(9).Editable = True
    RpCls(10).Width = 80
    If GlTeE = True Then 'Email-Termin-Erinnerung
        RpCls(11).Width = 120
    End If
End If

Set RpCol = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing

Exit Sub

SpErr:
If GlDbg = True Then SErLog Err.Description & " FSpal " & Err.Number
Resume Next

End Sub
Private Sub FZuru()
On Error Resume Next

Set Opti1 = Me.optTeNeu
Set Opti2 = Me.optTeKop
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück

If Rahm2.Visible = True Then
    Rahm1.Visible = True
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = False
    PuBu1.Enabled = True
    PuBu3.Enabled = False
ElseIf Rahm3.Visible = True Then
    If GlBut = RibTab_Rechnungen Then
        Rahm1.Visible = True
        Rahm2.Visible = False
        Rahm3.Visible = False
        Rahm4.Visible = False
    Else
        If Opti1.Value = True Then
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
        Else
            Rahm1.Visible = True
            Rahm2.Visible = False
            Rahm3.Visible = False
            Rahm4.Visible = False
        End If
    End If
    PuBu1.Enabled = True
ElseIf Rahm4.Visible = True Then
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
    PuBu1.Enabled = True
End If

End Sub
Private Sub FVors()
On Error GoTo LiErr

Dim TmVon As Date
Dim TmBis As Date
Dim DaBeg As Date
Dim DaEnd As Date
Dim DaAkt As Date
Dim NotDa As Date
Dim NotZe As Date
Dim PatNr As Long
Dim ManNr As Long
Dim MitNr As Long
Dim RmuNr As Long
Dim TerNr As Long
Dim MasNr As Long
Dim TeBtr As String
Dim AnzTa As Integer
Dim AnzBl As Integer
Dim AktBl As Integer
Dim AktTa As Integer
Dim BloTa As Integer
Dim AktTe As Integer
Dim NotVa As Integer
Dim AnzPo As Integer
Dim AnzTe As Integer
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo7 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpRws As XtremeReportControl.ReportRows
Dim RpRow As XtremeReportControl.ReportRow

Set FM = frmMain
Set CaCol = FM.calCont1
Set Opti1 = Me.optTeNeu
Set Opti2 = Me.optTeKop
Set DaPi1 = Me.dtpDatu1
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set TxNoD = Me.txtNoDat
Set TxNoZ = Me.txtNoTim
Set CmBet = Me.txtBetre
Set CmMar = Me.cmbTeTyp
Set CmTyp = Me.cmbStatu
Set CmRmu = Me.cmbRaum1
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar
Set CmNot = Me.cmbNotVa
Set CmPri = Me.cmbPrior
Set RpCo1 = Me.repCont1
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück
Set RpCls = RpCo1.Columns
Set RpRcs = RpCo1.Records
Set RpRws = RpCo1.Rows

RmuNr = CmRmu.ItemData(CmRmu.ListIndex)
ManNr = CmMan.ItemData(CmMan.ListIndex)
MitNr = CmMit.ItemData(CmMit.ListIndex)
NotVa = CmNot.ItemData(CmNot.ListIndex)

TmVon = TimeValue(VoZei.Text)
TmBis = TimeValue(BiZei.Text)

If CmBet.Text <> vbNullString Then
    TeBtr = CmBet.Text
End If

AnzBl = DaPi1.Selection.BlocksCount

If Opti2.Value = True Then
    If GlBut = RibTab_Rechnungen Then
        Set RpCo7 = frmKatRC.repCont7
        Set RpCls = RpCo7.Columns
        Set RpSel = RpCo7.SelectedRows
        AnzPo = RpSel.Count
        If AnzPo > 0 Then
            Set RpRow = RpSel(0)
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
            Set RpCol = RpCls.Find(Ter_Patient)
            TeBtr = RpRow.Record(RpCol.ItemIndex).Value
        End If
    Else
        Set ViEvs = CaCol.ActiveView.GetSelectedEvents
        AnzPo = ViEvs.Count
        If AnzPo > 0 Then 'Kopieren
            For Each ViEvt In ViEvs
                If ViEvt.Selected = True Then
                    Set CaEvt = ViEvt.Event
                    TerNr = CaEvt.id
                    If TerNr > 0 Then
                        S_TeDe TerNr
                        With GlTDt
                            MitNr = .TeMit
                            ManNr = .TeMan
                            RmuNr = .TeRau
                            If .TeBet <> vbNullString Then
                                If .PaStr <> vbNullString Then
                                    TeBtr = .PaStr & " - " & .TeBet
                                Else
                                    TeBtr = .TeBet
                                End If
                            ElseIf .PaStr <> vbNullString Then
                                TeBtr = .PaStr
                            End If
                        End With
                    End If
                End If
            Next ViEvt
        Else
            Exit Sub
        End If
    End If
End If

If AnzBl = 0 Then
    SPopu "Keine Tage selektiert", "Es wurde keine Tage selektiert", IC48_Forbidden
    PuBu1.Enabled = False
    Set DaPi1 = Nothing
    Exit Sub
ElseIf AnzBl = 1 Then
    DaBeg = DaPi1.Selection(0).DateBegin
    DaEnd = DaPi1.Selection(0).DateEnd
    If DaBeg < Date Then
        DaBeg = Date
    End If
    If DaEnd > DaBeg Then
        Do
        DaAkt = DaBeg + AktTa
        AktTa = AktTa + 1
        ReDim Preserve VoDat(AktTa)
        VoDat(AktTa) = DaAkt
        Loop Until DaAkt >= DaEnd
    Else
        ReDim Preserve VoDat(1)
        VoDat(1) = DaBeg
    End If
ElseIf AnzBl > 1 Then
    For AktBl = 0 To AnzBl - 1
        DaBeg = DaPi1.Selection.Blocks(AktBl).DateBegin
        DaEnd = DaPi1.Selection.Blocks(AktBl).DateEnd
        If DaBeg < Date Then
            DaBeg = Date
        End If
        If DaEnd > DaBeg Then
            BloTa = 0
            Do
            DaAkt = DaBeg + BloTa
            AktTa = AktTa + 1
            BloTa = BloTa + 1
            ReDim Preserve VoDat(AktTa)
            VoDat(AktTa) = DaAkt
            Loop Until DaAkt >= DaEnd
        Else
            AktTa = AktTa + 1
            ReDim Preserve VoDat(AktTa)
            VoDat(AktTa) = DaBeg
        End If
    Next AktBl
End If

AnzTe = UBound(VoDat)

With RpCo1
    .EditItem Nothing, Nothing
    If .SortOrder.Count > 0 Then .SortOrder.DeleteAll
    If .GroupsOrder.Count > 0 Then .GroupsOrder.DeleteAll
    If .Records.Count > 0 Then .Records.DeleteAll
    .Populate
End With

If GlBut = RibTab_Rechnungen Then
    For AktTe = 1 To AnzTe
        Set RpRec = RpRcs.Add()
        Set RpItm = RpRec.AddItem(vbNullString)
        RpItm.Icon = IC16_Calendar_Clock
        RpItm.Focusable = False
        Set RpItm = RpRec.AddItem(Format$(VoDat(AktTe), "dddd"))
        RpItm.Focusable = False
        Set RpItm = RpRec.AddItem(Format$(VoDat(AktTe), "dd.mm.yyyy"))
        RpItm.Alignment = xtpAlignmentCenter
        Set RpItm = RpRec.AddItem(vbNullString)
        RpItm.HasCheckbox = True
        RpItm.Checked = True
        Set RpItm = RpRec.AddItem(ManNr)
        RpItm.Focusable = False
        Set RpItm = RpRec.AddItem(MitNr)
        RpItm.Focusable = False
        Set RpItm = RpRec.AddItem(Format$(AktTe, "000") & " / " & Format$(AnzTe, "000"))
        RpItm.Alignment = xtpAlignmentCenter
        RpItm.Focusable = False
    Next AktTe
Else
    For AktTe = 1 To AnzTe
        Set RpRec = RpRcs.Add()
        Set RpItm = RpRec.AddItem(vbNullString)
        RpItm.Icon = IC16_Calendar_Clock
        RpItm.Focusable = False
        Set RpItm = RpRec.AddItem(Format$(VoDat(AktTe), "dddd"))
        RpItm.Focusable = False
        Set RpItm = RpRec.AddItem(Format$(VoDat(AktTe), "dd.mm.yyyy"))
        RpItm.Alignment = xtpAlignmentCenter
        Set RpItm = RpRec.AddItem(Format$(TmVon, "hh:mm"))
        RpItm.Alignment = xtpAlignmentCenter
        Set RpItm = RpRec.AddItem(Format$(TmBis, "hh:mm"))
        RpItm.Alignment = xtpAlignmentCenter
        Set RpItm = RpRec.AddItem(vbNullString)
        RpItm.HasCheckbox = True
        RpItm.Checked = True
        Set RpItm = RpRec.AddItem(vbNullString)
        Set RpItm = RpRec.AddItem(TeBtr)
        Set RpItm = RpRec.AddItem(MitNr)
        RpItm.Focusable = False
        Set RpItm = RpRec.AddItem(RmuNr)
        Set RpItm = RpRec.AddItem(Format$(AktTe, "000") & " / " & Format$(AnzTe, "000"))
        RpItm.Alignment = xtpAlignmentCenter
        RpItm.Focusable = False
        If GlTeE = True Then 'Email-Termin-Erinnerung
            NotDa = CDate(DateAdd("h", -NotVa, VoDat(AktTe) & " " & TmVon))
            NotZe = TimeValue(DateAdd("h", -NotVa, VoDat(AktTe) & " " & TmVon))
            Set RpItm = RpRec.AddItem(Format$(NotDa, "dd.mm.yyyy") & " " & Format$(NotZe, "hh:mm"))
        End If
    Next AktTe
End If

RpCo1.Populate
If RpRws.Count > 0 Then
    RpRws.Row(0).Selected = False
End If

Set RpRcs = Nothing
Set RpCo1 = Nothing

Exit Sub

LiErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FVors " & Err.Number
Resume Next

End Sub
Private Sub TWeit()
On Error GoTo SuErr

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim AnzPo As Integer
Dim AnzBl As Integer
Dim CaEvt As XtremeCalendarControl.CalendarEvent
Dim CaEvs As XtremeCalendarControl.CalendarEvents
Dim CaHit As XtremeCalendarControl.CalendarHitTestInfo
Dim ViEvt As XtremeCalendarControl.CalendarViewEvent
Dim ViEvs As XtremeCalendarControl.CalendarViewEvents
Dim CaCol As XtremeCalendarControl.CalendarControl
Dim RpCon As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns
Dim RpCol As XtremeReportControl.ReportColumn
Dim RpRws As XtremeReportControl.ReportRows

Set FM = frmMain
Set CaCol = FM.calCont1
Set Opti1 = Me.optTeNeu
Set Opti2 = Me.optTeKop
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set Rahm3 = Me.frmRahm3
Set Rahm4 = Me.frmRahm4
Set DaPi1 = Me.dtpDatu1
Set CmBet = Me.txtBetre
Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set RpCon = Me.repCont1
Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück

TeTit = "Termin Kopieren"
TeMai = "Es wurde kein Termin markiert"
TeInh = "Um einen Termin kopieren zu können, ist es erforderlich, diesen vorher im Kalender zu markieren."
TeFus = "Der Terminassistent ermöglicht das hinzufügen neuer Termine oder das Kopieren eines Termins auf mehrere Tage."

If Rahm1.Visible = True Then
    If GlBut = RibTab_Rechnungen Then
        Rahm1.Visible = False
        Rahm2.Visible = False
        Rahm3.Visible = True
        Rahm4.Visible = False
        PuBu3.Enabled = True
    Else
        Set ViEvs = CaCol.ActiveView.GetSelectedEvents
        AnzPo = ViEvs.Count
        If Opti2.Value = True Then
            If AnzPo = 0 Then
                SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, frmTerKop.hwnd
                PuBu1.Enabled = False
                Set DaPi1 = Nothing
                Set CaCol = Nothing
                Exit Sub
            Else
                Rahm1.Visible = False
                Rahm2.Visible = False
                Rahm3.Visible = True
                Rahm4.Visible = False
                PuBu3.Enabled = True
            End If
        Else
            Rahm1.Visible = False
            Rahm2.Visible = True
            Rahm3.Visible = False
            Rahm4.Visible = False
            PuBu3.Enabled = True
        End If
    End If
ElseIf Rahm2.Visible = True Then
    If CmBet.Text = vbNullString Then
        SPopu "Kein Terminbetreff", "Es wurde kein Terminbetreff festgelegt", IC48_Forbidden
        PuBu1.Enabled = False
        Set DaPi1 = Nothing
        Set CaCol = Nothing
        Exit Sub
    End If
    If VoZei.Text <> vbNullString Then
        If IsDate(VoZei.Text) = False Then
            PuBu1.Enabled = False
            Set DaPi1 = Nothing
            Set CaCol = Nothing
        End If
    Else
        PuBu1.Enabled = False
        Set DaPi1 = Nothing
        Set CaCol = Nothing
    End If
    If BiZei.Text <> vbNullString Then
        If IsDate(BiZei.Text) = False Then
            PuBu1.Enabled = False
            Set DaPi1 = Nothing
            Set CaCol = Nothing
        End If
    Else
        PuBu1.Enabled = False
        Set DaPi1 = Nothing
        Set CaCol = Nothing
    End If
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = True
    Rahm4.Visible = False
ElseIf Rahm3.Visible = True Then
    AnzBl = DaPi1.Selection.BlocksCount
    If AnzBl = 0 Then
        SPopu "Keine Tage selektiert", "Es wurde keine Tage selektiert", IC48_Forbidden
        PuBu1.Enabled = False
        Set DaPi1 = Nothing
        Set CaCol = Nothing
        Exit Sub
    End If
    FVors
    DoEvents
    Rahm1.Visible = False
    Rahm2.Visible = False
    Rahm3.Visible = False
    Rahm4.Visible = True
ElseIf Rahm4.Visible = True Then
    Screen.MousePointer = vbHourglass
    DoEvents

    If GlBut = RibTab_Rechnungen Then
        Ter_Ser
        DoEvents
        P_List "ReSe", 0, 1
    Else
        If Opti2.Value = True Then
            Ter_VoP True
        Else
            Ter_VoP
        End If
        DoEvents
        S_TeLi
        DoEvents
        S_TePi 'Kalndermarker setzen
        DoEvents
        SUpTe
    End If
    
    DoEvents
    Screen.MousePointer = vbNormal
        
    DoEvents
    Unload Me
End If

Set DaPi1 = Nothing
Set RpSel = Nothing
Set RpRws = Nothing
Set RpCls = Nothing
Set RpCon = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TWeit " & Err.Number
Resume Next

End Sub
Private Sub FZKo1()
On Error GoTo PoErr

Dim TmVon As Date
Dim TmBis As Date
Dim ZeiVo As Integer
Dim AktZa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre

If InStrRev(VoZei.Text, "_", -1, 1) > 0 Then
    VoZei.Text = "00:00"
End If

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
Else
    TmVon = Now
    VoZei.Text = Format$(TimeValue(Now), "hh:mm")
End If

If BiZei.Text <> vbNullString Then
    TmBis = TimeValue(BiZei.Text)
Else
    TmVon = Now
    BiZei.Text = Format$(TimeValue(Now), "hh:mm")
End If

If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
    If ZeiVo > 0 Then
        TmBis = DateAdd("n", ZeiVo, TmVon)
        BiZei.Text = Format$(TmBis, "hh:mm")
    Else
        If TmVon > TmBis Then
            BiZei.Text = VoZei.Text
        End If
    End If
Else
    If TmVon > TmBis Then
        BiZei.Text = VoZei.Text
    End If
End If

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZKo1 " & Err.Number
Resume Next

End Sub
Private Sub FZKo2()
On Error GoTo PoErr

Dim TmVon As Date
Dim TmBis As Date
Dim ZeiVo As Integer
Dim AktZa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre

If InStrRev(BiZei.Text, "_", -1, 1) > 0 Then
    BiZei.Text = "00:00"
End If

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

If VoZei.Text <> vbNullString Then
    TmVon = TimeValue(VoZei.Text)
Else
    VoZei.Text = Format$(TimeValue(Now), "hh:mm")
    TmVon = Now
End If

If BiZei.Text <> vbNullString Then
    TmBis = TimeValue(BiZei.Text)
Else
    BiZei.Text = Format$(TimeValue(Now), "hh:mm")
    TmVon = Now
End If

If GlTeZ = True Then 'Terminzeit aus dem Terminbetreff verwenden
    If ZeiVo > 0 Then
        TmVon = DateAdd("n", -ZeiVo, TmBis)
        VoZei.Text = Format$(TmVon, "hh:mm")
    Else
        If TmVon > TmBis Then
            BiZei.Text = VoZei.Text
        End If
    End If
Else
    If TmVon > TmBis Then
        BiZei.Text = VoZei.Text
    End If
End If

Exit Sub

PoErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FZKo2 " & Err.Number
Resume Next

End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Set Opti1 = Me.optTeNeu
Set Opti2 = Me.optTeKop

If Opti2.Value = True Then  'Kopieren
    TeTit = IniGetOpt("Hilfe", 51111)
    TeMai = IniGetOpt("Hilfe", 51112)
    TeInh = IniGetOpt("Hilfe", 51113)
    TeFus = IniGetOpt("Hilfe", 51114)
Else
    If Urlaub = True Then
        TeTit = IniGetOpt("Hilfe", 51121)
        TeMai = IniGetOpt("Hilfe", 51122)
        TeInh = IniGetOpt("Hilfe", 51123)
        TeFus = IniGetOpt("Hilfe", 51124)
    Else
        TeTit = IniGetOpt("Hilfe", 51131)
        TeMai = IniGetOpt("Hilfe", 51132)
        TeInh = IniGetOpt("Hilfe", 51133)
        TeFus = IniGetOpt("Hilfe", 51134)
    End If
End If

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    TWeit
End Sub
Private Sub btnZurück_Click()
    FZuru
End Sub
Private Sub chkExpMo_Click()
    FExp
End Sub

Private Sub cmbBehan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbMitar_Click()

If GlTeF = False Then
    FMita
    FNoVa
End If

End Sub
Private Sub cmbMitar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbNotVa_GotFocus()
    RetWe = SendMessage(Me.cmbNotVa.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbNotVa_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbPrior_GotFocus()
    RetWe = SendMessage(Me.cmbPrior.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbPrior_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbPrior_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbPrior.SelLength = 0
End Sub
Private Sub cmbRaum1_GotFocus()
    RetWe = SendMessage(Me.cmbRaum1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbRaum1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbStatu_GotFocus()
    RetWe = SendMessage(Me.cmbStatu.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbStatu_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbStatu_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Me.cmbStatu.SelLength = 0
End Sub
Private Sub dtpDatu1_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
On Error Resume Next

Dim AktTa As Long
Dim AktZa As Integer

If Weekday(Day, vbMonday) = vbSaturday Then
    Metrics.ForeColor = vbRed
End If

If GlTeV = True Then 'Termine vorhanden
    If GlTpV = True Then 'Kalendermarker vorhanden
        For AktTa = 0 To GlKMa - 1 'Anzahl Kalendermatker
            If Day = Left$(GlTEr(0, AktTa), 10) Then
                For AktZa = 1 To UBound(GlTep) 'Kalendermarker
                    If GlTep(AktZa, 0) = GlTEr(1, AktTa) Then
                        Metrics.BackColor = GlTep(AktZa, 2)
                        Exit For
                    End If
                Next AktZa
            End If
        Next AktTa
    End If
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
FInit
FSpal
FLoad
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmTerKop = Nothing
End Sub

Private Sub optTeNeu_Click()
On Error Resume Next

Set PuBu1 = Me.btnWeiter
Set PuBu2 = Me.btnSchließ
Set PuBu3 = Me.btnZurück

If PuBu1.Enabled = False Then
    PuBu1.Enabled = True
End If

End Sub

Private Sub repCont1_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error Resume Next

Dim FrbZa As Long

If Row.GroupRow = False Then
    If Row.Record(6).Value <> vbNullString Then
        If IsNumeric(Row.Record(6).Value) = True Then
            FrbZa = Row.Record(6).Value
            If FrbZa > 0 Then
                Metrics.BackColor = FrbZa
            End If
        End If
    End If
End If

End Sub

Private Sub repCont1_InplaceButtonDown(ByVal Button As XtremeReportControl.IReportInplaceButton)
    If Button.Column.ItemIndex = 2 Then
        FKaRo Button
    End If
End Sub
Private Sub repCont1_ValueChanged(ByVal Row As XtremeReportControl.IReportRow, ByVal Column As XtremeReportControl.IReportColumn, ByVal Item As XtremeReportControl.IReportRecordItem)
    FKrCa Column.ItemIndex
End Sub
Private Sub txtBetre_Change()
    If GlTeF = False Then 'Formular wird geladen
        FBetr
    End If
End Sub
Private Sub txtBetre_Click()
    If GlTeF = False Then
        FBetr
    End If
End Sub
Private Sub txtBetre_GotFocus()
    GlTeF = False
End Sub

Private Sub txtBetre_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
    GlTeF = False
End Sub

Private Sub txtBisZe_GotFocus()
    Me.txtBisZe.SelStart = 0
    Me.txtBisZe.SelLength = Len(Me.txtBisZe.Text)
End Sub

Private Sub txtBisZe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtBisZe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtBisZe.SelLength = 0
    ElseIf KeyCode = vbKeyReturn Then
        FZKo2
    End If
End Sub

Private Sub txtBisZe_LostFocus()
    FZKo2
End Sub
Private Sub txtVonZe_GotFocus()
    Me.txtVonZe.SelStart = 0
    Me.txtVonZe.SelLength = Len(Me.txtVonZe.Text)
End Sub

Private Sub txtVonZe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtVonZe_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Me.txtVonZe.SelLength = 0
    ElseIf KeyCode = vbKeyReturn Then
        FZKo1
    End If
End Sub

Private Sub txtVonZe_LostFocus()
    FZKo1
End Sub
Private Sub updCont2_DownClick()
On Error Resume Next

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If MiIdx <= UBound(GlMiT) Then
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMiT(GlSmI, 8)
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If MaIdx <= UBound(GlMaT) Then
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMaT(GlSMa, 8)
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If MiIdx <= UBound(GlMiT) Then
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMiT(GlSmI, 8)
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If MaIdx <= UBound(GlMaT) Then
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMaT(GlSMa, 8)
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If MiIdx <= UBound(GlMiT) Then
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMiT(GlSmI, 8)
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If MaIdx <= UBound(GlMaT) Then
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMaT(GlSMa, 8)
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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

Dim AlDa1 As Date
Dim AlDa2 As Date
Dim TmVon As Date
Dim TmBis As Date
Dim MaIdx As Integer
Dim MiIdx As Integer
Dim ZeiVo As Integer
Dim AktZa As Integer
Dim MiDif As Integer
Dim ZeiRa As Integer

Set VoZei = Me.txtVonZe
Set BiZei = Me.txtBisZe
Set CmBet = Me.txtBetre
Set CmMan = Me.cmbBehan
Set CmMit = Me.cmbMitar

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    If CmMit.ListCount > 0 Then
        If CmMit.ListIndex >= 0 Then
            MiIdx = CmMit.ListIndex + 1
        Else
            MiIdx = GlSmI
        End If
    Else
        MiIdx = GlSmI
    End If
    If MiIdx <= UBound(GlMiT) Then
        If GlMiT(MiIdx, 8) > 0 Then
            ZeiRa = GlMiT(MiIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMiT(GlSmI, 8)
    End If
Else
    If CmMan.ListCount > 0 Then
        If CmMan.ListIndex >= 0 Then
            MaIdx = CmMan.ListIndex + 1
        Else
            MaIdx = GlSMa
        End If
    Else
        MaIdx = GlSMa
    End If
    If MaIdx <= UBound(GlMaT) Then
        If GlMaT(MaIdx, 8) > 0 Then
            ZeiRa = GlMaT(MaIdx, 8)
        Else
            ZeiRa = GlZeR 'Zeitrasterindex
        End If
    Else
        ZeiRa = GlMaT(GlSMa, 8)
    End If
End If

MiDif = GlTku(ZeiRa, 2)

If CmBet.Text <> vbNullString Then
    For AktZa = 1 To UBound(GlBtr)
        If CmBet.Text = GlBtr(AktZa, 1) Then
            ZeiVo = GlBtr(AktZa, 3)
            Exit For
        End If
    Next AktZa
End If

If ZeiVo = 0 Then
    ZeiVo = 15
End If

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

