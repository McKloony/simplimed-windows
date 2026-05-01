VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmBaEdit 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Kontoauszug"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   27
      Top             =   7900
      Width           =   7600
      _Version        =   1048579
      _ExtentX        =   13406
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   5600
         TabIndex        =   30
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
      Begin XtremeSuiteControls.PushButton btnWieter 
         Default         =   -1  'True
         Height          =   400
         Left            =   4200
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schließen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   2900
         TabIndex        =   28
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
   Begin XtremeSuiteControls.CheckBox chkGewEr 
      Height          =   240
      Left            =   700
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7300
      Width           =   3300
      _Version        =   1048579
      _ExtentX        =   5821
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Keine Berücksichtigung bei Erlösermittlung"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtKonto 
      Height          =   350
      Left            =   3840
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6030
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtBuTex 
      Height          =   350
      Left            =   3840
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6730
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtUmsat 
      Height          =   1000
      Left            =   700
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2530
      Width           =   6140
      _Version        =   1048579
      _ExtentX        =   10830
      _ExtentY        =   1764
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      MultiLine       =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkBezal 
      Height          =   240
      Left            =   5500
      TabIndex        =   26
      Tag             =   "0Abhaken"
      Top             =   7300
      Width           =   1100
      _Version        =   1048579
      _ExtentX        =   1940
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Gebucht"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkSelek 
      Height          =   240
      Left            =   4200
      TabIndex        =   25
      Tag             =   "0Erledigt"
      Top             =   7300
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Zugeordnet"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   9200
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbGegen 
      Height          =   310
      Left            =   700
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5330
      Width           =   2000
      _Version        =   1048579
      _ExtentX        =   3519
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox3"
   End
   Begin XtremeSuiteControls.FlatEdit txtOffen 
      Height          =   350
      Left            =   5460
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3930
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cmbManda 
      Height          =   310
      Left            =   700
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4630
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.FlatEdit txtKomme 
      Height          =   350
      Left            =   700
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1830
      Width           =   6140
      _Version        =   1048579
      _ExtentX        =   10830
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtBetra 
      Height          =   350
      Left            =   2300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   430
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtBeleg 
      Height          =   350
      Left            =   5460
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   430
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtDatum 
      Height          =   350
      Left            =   700
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   430
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtBele2 
      Height          =   350
      Left            =   700
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1130
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtBele3 
      Height          =   350
      Left            =   2300
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1130
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtBele4 
      Height          =   350
      Left            =   3840
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1130
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtOffe2 
      Height          =   345
      Left            =   3825
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5330
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtOffe3 
      Height          =   345
      Left            =   5415
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5330
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtOffe4 
      Height          =   350
      Left            =   3840
      TabIndex        =   15
      Top             =   4630
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtBele5 
      Height          =   350
      Left            =   5460
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1130
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtGesam 
      Height          =   350
      Left            =   3840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   430
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtOffe5 
      Height          =   350
      Left            =   5460
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4630
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtAusga 
      Height          =   350
      Left            =   3840
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3930
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox cmbMitar 
      Height          =   315
      Left            =   700
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3930
      Width           =   2920
      _Version        =   1048579
      _ExtentX        =   5133
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbBuStu 
      Height          =   315
      Left            =   700
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6030
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
   Begin XtremeSuiteControls.FlatEdit txtKonNr 
      Height          =   350
      Left            =   700
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6730
      Width           =   2000
      _Version        =   1048579
      _ExtentX        =   3528
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label lblLab09 
      Height          =   210
      Left            =   705
      TabIndex        =   53
      Top             =   6490
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Sachkontennummer :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab08 
      Height          =   210
      Left            =   705
      TabIndex        =   52
      Top             =   4390
      Width           =   900
      _Version        =   1048579
      _ExtentX        =   1587
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Mandant :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab07 
      Height          =   210
      Left            =   3845
      TabIndex        =   51
      Top             =   5790
      Width           =   1900
      _Version        =   1048579
      _ExtentX        =   3351
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Sachkontenbezeichnung :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab06 
      Height          =   210
      Left            =   3845
      TabIndex        =   50
      Top             =   6490
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Buchungstext :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblLab05 
      Height          =   210
      Left            =   705
      TabIndex        =   49
      Top             =   5790
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Steuersatz :"
      Transparent     =   -1  'True
   End
   Begin VB.Label lblLab29 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungsart :"
      Height          =   210
      Left            =   3840
      TabIndex        =   48
      Top             =   3690
      Width           =   1400
   End
   Begin VB.Label lblLab28 
      BackStyle       =   0  'Transparent
      Caption         =   "Offen5 :"
      Height          =   210
      Left            =   5465
      TabIndex        =   47
      Top             =   4390
      Width           =   1400
   End
   Begin VB.Label lblLab27 
      BackStyle       =   0  'Transparent
      Caption         =   "Summe Offen :"
      Height          =   210
      Left            =   3845
      TabIndex        =   46
      Top             =   190
      Width           =   1400
   End
   Begin VB.Label lblLab26 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnung5 :"
      Height          =   210
      Left            =   5465
      TabIndex        =   45
      Top             =   890
      Width           =   1400
   End
   Begin VB.Label lblLab24 
      BackStyle       =   0  'Transparent
      Caption         =   "Offen4 :"
      Height          =   210
      Left            =   3845
      TabIndex        =   44
      Top             =   4390
      Width           =   1400
   End
   Begin VB.Label lblLab23 
      BackStyle       =   0  'Transparent
      Caption         =   "Offen3 :"
      Height          =   210
      Left            =   5465
      TabIndex        =   43
      Top             =   5090
      Width           =   1395
   End
   Begin VB.Label lblLab22 
      BackStyle       =   0  'Transparent
      Caption         =   "Offen2 :"
      Height          =   210
      Left            =   3845
      TabIndex        =   42
      Top             =   5090
      Width           =   1395
   End
   Begin VB.Label lblLab21 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnung4 :"
      Height          =   210
      Left            =   3845
      TabIndex        =   41
      Top             =   890
      Width           =   1400
   End
   Begin VB.Label lblLab19 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnung3 :"
      Height          =   210
      Left            =   2305
      TabIndex        =   40
      Top             =   890
      Width           =   1400
   End
   Begin VB.Label lblLab18 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnung2 :"
      Height          =   210
      Left            =   705
      TabIndex        =   39
      Top             =   890
      Width           =   1400
   End
   Begin VB.Label lblLab04 
      BackStyle       =   0  'Transparent
      Caption         =   "Geldkonto :"
      Height          =   210
      Left            =   705
      TabIndex        =   38
      Top             =   5090
      Width           =   900
   End
   Begin VB.Label lblLab01 
      BackStyle       =   0  'Transparent
      Caption         =   "Mitarbeiter :"
      Height          =   210
      Left            =   705
      TabIndex        =   37
      Top             =   3690
      Width           =   1200
   End
   Begin VB.Label lblLab02 
      BackStyle       =   0  'Transparent
      Caption         =   "Offen1 :"
      Height          =   210
      Left            =   5465
      TabIndex        =   36
      Top             =   3690
      Width           =   1400
   End
   Begin VB.Label lblLab20 
      BackStyle       =   0  'Transparent
      Caption         =   "Umsatztext :"
      Height          =   210
      Left            =   705
      TabIndex        =   35
      Top             =   2300
      Width           =   1200
   End
   Begin VB.Label lblLab25 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient / Kommentar :"
      Height          =   210
      Left            =   705
      TabIndex        =   34
      Top             =   1590
      Width           =   2000
   End
   Begin VB.Label lblLab17 
      BackStyle       =   0  'Transparent
      Caption         =   "Rechnung1 :"
      Height          =   210
      Left            =   5465
      TabIndex        =   33
      Top             =   190
      Width           =   1400
   End
   Begin VB.Label lblLab16 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungsbetrag :"
      Height          =   210
      Left            =   2305
      TabIndex        =   32
      Top             =   190
      Width           =   1400
   End
   Begin VB.Label lblLab15 
      BackStyle       =   0  'Transparent
      Caption         =   "Buchungsdatum :"
      Height          =   210
      Left            =   705
      TabIndex        =   31
      Top             =   190
      Width           =   1400
   End
End
Attribute VB_Name = "frmBaEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private CmGeg As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmBuS As XtremeSuiteControls.ComboBox
Private TxDat As XtremeSuiteControls.FlatEdit
Private TxBet As XtremeSuiteControls.FlatEdit
Private TxBel As XtremeSuiteControls.FlatEdit
Private TxKom As XtremeSuiteControls.FlatEdit
Private TxBuT As XtremeSuiteControls.FlatEdit
Private TxUms As XtremeSuiteControls.FlatEdit
Private TxOff As XtremeSuiteControls.FlatEdit
Private ChSel As XtremeSuiteControls.CheckBox
Private ChGeb As XtremeSuiteControls.CheckBox
Private ChErm As XtremeSuiteControls.CheckBox

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

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim ThIdx As Integer
Dim AktZa As Integer
Dim AktKo As Integer
Dim LiIdx As Integer
Dim CmBrs As XtremeCommandBars.CommandBars
Dim CmThe As XtremeCommandBars.CommandBarComboBox
Dim CmCom As XtremeCommandBars.CommandBarComboBox

Set FM = frmBaEdit
Set Rahm0 = FM.frmRahm0
Set CmGeg = FM.cmbGegen
Set CmMan = FM.cmbManda
Set CmMit = FM.cmbMitar
Set CmBuS = FM.cmbBuStu
Set TxDat = FM.txtDatum
Set TxBel = FM.txtBeleg
Set TxBet = FM.txtBetra
Set TxKom = FM.txtKomme
Set TxUms = FM.txtUmsat
Set TxOff = FM.txtOffen
Set ChSel = FM.chkSelek
Set ChGeb = FM.chkBezal
Set ChErm = FM.chkGewEr
Set CmBrs = frmMain.comBar01

Set CmThe = CmBrs.FindControl(CmThe, SY_SuMan, , True)
ThIdx = CmThe.ListIndex

Set CmCom = CmBrs.FindControl(CmCom, SY_BA_Banking_SuchCombo, , True)
LiIdx = CmCom.ListIndex

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

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(CmMan.NewIndex) = GlThe(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
    CmMit.AddItem GlMiK(AktZa, 1)
    CmMit.ItemData(CmMit.NewIndex) = GlMiK(AktZa, 2)
Next AktZa

For AktZa = 1 To UBound(GlStu)
    CmBuS.AddItem GlStu(AktZa, 2)
    CmBuS.ItemData(CmBuS.NewIndex) = GlStu(AktZa, 0)
Next AktZa

CmMan.ListIndex = GlMan(GlSMa, 0) - 1
CmMit.ListIndex = GlMiK(GlSmI, 0) - 1

With TxDat
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

If CmMan.Enabled = False Then
    CmMan.Enabled = True
End If

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
ChSel.BackColor = GlBak
ChGeb.BackColor = GlBak
ChErm.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50071)
TeMai = IniGetOpt("Hilfe", 50072)
TeInh = IniGetOpt("Hilfe", 50073)
TeFus = IniGetOpt("Hilfe", 50074)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnWieter_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

FInit
AFont Me
S_Posi
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmBaEdit = Nothing
End Sub

Private Sub btnSchließ_Click()
    Unload Me
End Sub

Private Sub txtBele2_GotFocus()
    Me.txtBele2.SelStart = 0
    Me.txtBele2.SelLength = Len(Me.txtBele2.Text)
End Sub

Private Sub txtBele3_GotFocus()
    Me.txtBele3.SelStart = 0
    Me.txtBele3.SelLength = Len(Me.txtBele3.Text)
End Sub

Private Sub txtBele4_GotFocus()
    Me.txtBele4.SelStart = 0
    Me.txtBele4.SelLength = Len(Me.txtBele4.Text)
End Sub

Private Sub txtBele5_GotFocus()
    Me.txtBele5.SelStart = 0
    Me.txtBele5.SelLength = Len(Me.txtBele5.Text)
End Sub
Private Sub txtBeleg_GotFocus()
    Me.txtBeleg.SelStart = 0
    Me.txtBeleg.SelLength = Len(Me.txtBeleg.Text)
End Sub


Private Sub txtBetra_GotFocus()
    Me.txtBetra.SelStart = 0
    Me.txtBetra.SelLength = Len(Me.txtBetra.Text)
End Sub


Private Sub txtBuTex_GotFocus()
    Me.txtBuTex.SelStart = 0
    Me.txtBuTex.SelLength = Len(Me.txtBuTex.Text)
End Sub
Private Sub txtDatum_GotFocus()
    Me.txtDatum.SelStart = 0
    Me.txtDatum.SelLength = Len(Me.txtDatum.Text)
End Sub
Private Sub txtKomme_GotFocus()
    Me.txtKomme.SelStart = 0
    Me.txtKomme.SelLength = Len(Me.txtKomme.Text)
End Sub

Private Sub txtKonNr_GotFocus()
    Me.txtKonNr.SelStart = 0
    Me.txtKonNr.SelLength = Len(Me.txtKonNr.Text)
End Sub
Private Sub txtKonto_Change()
    Me.txtKonto.SelStart = 0
    Me.txtKonto.SelLength = Len(Me.txtKonto.Text)
End Sub


Private Sub txtOffe2_GotFocus()
    Me.txtOffe2.SelStart = 0
    Me.txtOffe2.SelLength = Len(Me.txtOffe2.Text)
End Sub

Private Sub txtOffe3_GotFocus()
    Me.txtOffe3.SelStart = 0
    Me.txtOffe3.SelLength = Len(Me.txtOffe3.Text)
End Sub

Private Sub txtOffe4_GotFocus()
    Me.txtOffe4.SelStart = 0
    Me.txtOffe4.SelLength = Len(Me.txtOffe4.Text)
End Sub

Private Sub txtOffe5_GotFocus()
    Me.txtOffe5.SelStart = 0
    Me.txtOffe5.SelLength = Len(Me.txtOffe5.Text)
End Sub
Private Sub txtOffen_GotFocus()
    Me.txtOffen.SelStart = 0
    Me.txtOffen.SelLength = Len(Me.txtOffen.Text)
End Sub

Private Sub txtUmsat_GotFocus()
    Me.txtUmsat.SelStart = 0
    Me.txtUmsat.SelLength = Len(Me.txtUmsat.Text)
End Sub


