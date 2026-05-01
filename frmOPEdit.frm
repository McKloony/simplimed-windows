VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmOPEdit 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Posten"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   28
      Top             =   6600
      Width           =   7600
      _Version        =   1048579
      _ExtentX        =   13406
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   1500
         TabIndex        =   29
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Hilfe"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   5600
         TabIndex        =   32
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
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Speichern"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnBezal 
         Height          =   400
         Left            =   2800
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Bezahlt"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtBICNr 
      Height          =   350
      Left            =   3840
      TabIndex        =   18
      Top             =   3230
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtIBANr 
      Height          =   350
      Left            =   3840
      TabIndex        =   16
      Top             =   2530
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   450
      Left            =   240
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   8500
      Visible         =   0   'False
      Width           =   450
      _Version        =   1048579
      _ExtentX        =   794
      _ExtentY        =   794
      _StockProps     =   64
      Show3DBorder    =   2
      VisualTheme     =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtOpKom 
      Height          =   350
      Left            =   700
      TabIndex        =   23
      Top             =   5330
      Width           =   6140
      _Version        =   1048579
      _ExtentX        =   10830
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtOPBel 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   8500
      Visible         =   0   'False
      Width           =   315
      _Version        =   1048579
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.PushButton btnPatSu 
      Height          =   350
      Left            =   3330
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Sucht den gewünschten Patienten"
      Top             =   1830
      Width           =   350
      _Version        =   1048579
      _ExtentX        =   617
      _ExtentY        =   617
      _StockProps     =   79
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPatie 
      Height          =   350
      Left            =   700
      TabIndex        =   11
      Top             =   1830
      Width           =   2610
      _Version        =   1048579
      _ExtentX        =   4604
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.PushButton btnDatu1 
      Height          =   350
      Left            =   6600
      TabIndex        =   5
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
   Begin XtremeSuiteControls.FlatEdit txtRechn 
      Height          =   350
      Left            =   700
      TabIndex        =   1
      Top             =   430
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
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
   Begin XtremeSuiteControls.ComboBox cmbStufe 
      Height          =   310
      Left            =   3840
      TabIndex        =   13
      Top             =   1830
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cmbManda 
      Height          =   315
      Left            =   705
      TabIndex        =   15
      Top             =   2535
      Width           =   2895
      _Version        =   1048579
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox2"
   End
   Begin XtremeSuiteControls.ComboBox cmbMahnb 
      Height          =   315
      Left            =   5460
      TabIndex        =   14
      Top             =   1830
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox3"
   End
   Begin XtremeSuiteControls.FlatEdit txtOPDat 
      Height          =   350
      Left            =   2300
      TabIndex        =   2
      Top             =   430
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDatu1 
      Height          =   350
      Left            =   5460
      TabIndex        =   4
      Top             =   430
      Width           =   1120
      _Version        =   1048579
      _ExtentX        =   1976
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtFalli 
      Height          =   350
      Left            =   3840
      TabIndex        =   3
      Top             =   430
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtOPBet 
      Height          =   350
      Left            =   700
      TabIndex        =   6
      Top             =   1130
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
   Begin XtremeSuiteControls.FlatEdit txtOffen 
      Height          =   350
      Left            =   2300
      TabIndex        =   7
      Top             =   1130
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
   Begin XtremeSuiteControls.FlatEdit txtOPBez 
      Height          =   350
      Left            =   3840
      TabIndex        =   9
      Top             =   1130
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtMahn1 
      Height          =   350
      Left            =   720
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6030
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtMahn2 
      Height          =   350
      Left            =   2310
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6030
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtMahn3 
      Height          =   350
      Left            =   3855
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6030
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtStKto 
      Height          =   350
      Left            =   700
      TabIndex        =   19
      Top             =   3930
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5115
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtStBLZ 
      Height          =   350
      Left            =   700
      TabIndex        =   21
      Top             =   4630
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5115
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.ComboBox cmbMitar 
      Height          =   315
      Left            =   700
      TabIndex        =   17
      Top             =   3230
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5106
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtGlaID 
      Height          =   350
      Left            =   3840
      TabIndex        =   22
      Top             =   4630
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.FlatEdit txtManID 
      Height          =   350
      Left            =   3840
      TabIndex        =   20
      Top             =   3930
      Width           =   3000
      _Version        =   1048579
      _ExtentX        =   5292
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      BackColor       =   16777215
   End
   Begin XtremeSuiteControls.ComboBox cmbWarun 
      Height          =   315
      Left            =   5460
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6030
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
   Begin XtremeSuiteControls.FlatEdit txtGebue 
      Height          =   350
      Left            =   5460
      TabIndex        =   10
      Top             =   1130
      Width           =   1395
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin VB.Label lblLab32 
      BackStyle       =   0  'Transparent
      Caption         =   "Währung :"
      Height          =   210
      Left            =   5465
      TabIndex        =   57
      Top             =   5800
      Width           =   900
   End
   Begin XtremeSuiteControls.Label Label12 
      Height          =   210
      Left            =   3860
      TabIndex        =   56
      Top             =   3700
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Mandatsreferenz :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label11 
      Height          =   210
      Left            =   3860
      TabIndex        =   55
      Top             =   4400
      Width           =   1800
      _Version        =   1048579
      _ExtentX        =   3175
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Gläubigeridentifikation :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label10 
      Height          =   210
      Left            =   3860
      TabIndex        =   54
      Top             =   3000
      Width           =   900
      _Version        =   1048579
      _ExtentX        =   1587
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "BIC :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label9 
      Height          =   210
      Left            =   3860
      TabIndex        =   53
      Top             =   2300
      Width           =   900
      _Version        =   1048579
      _ExtentX        =   1587
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "IBAN :"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label8 
      Height          =   210
      Left            =   705
      TabIndex        =   52
      Top             =   3000
      Width           =   900
      _Version        =   1048579
      _ExtentX        =   1587
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Mitarbeiter :"
      Transparent     =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Konto :"
      Height          =   210
      Left            =   705
      TabIndex        =   51
      Top             =   3700
      Width           =   900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BLZ :"
      Height          =   210
      Left            =   705
      TabIndex        =   50
      Top             =   4400
      Width           =   900
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Mahnung"
      Height          =   210
      Left            =   3855
      TabIndex        =   49
      Top             =   5800
      Width           =   900
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Mahnung"
      Height          =   210
      Left            =   2325
      TabIndex        =   48
      Top             =   5800
      Width           =   900
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Mahnung"
      Height          =   210
      Left            =   720
      TabIndex        =   47
      Top             =   5800
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bezahlt :"
      Height          =   210
      Left            =   3860
      TabIndex        =   46
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mahnbar :"
      Height          =   210
      Left            =   5465
      TabIndex        =   44
      Top             =   1600
      Width           =   900
   End
   Begin VB.Label lblLab15 
      BackStyle       =   0  'Transparent
      Caption         =   "Rech-Nr :"
      Height          =   210
      Left            =   705
      TabIndex        =   43
      Top             =   200
      Width           =   900
   End
   Begin VB.Label lblLab16 
      BackStyle       =   0  'Transparent
      Caption         =   "Datum :"
      Height          =   210
      Left            =   2305
      TabIndex        =   42
      Top             =   200
      Width           =   900
   End
   Begin VB.Label lblLab17 
      BackStyle       =   0  'Transparent
      Caption         =   "Fälligkeit :"
      Height          =   210
      Left            =   3860
      TabIndex        =   41
      Top             =   200
      Width           =   900
   End
   Begin VB.Label lblLab20 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient :"
      Height          =   210
      Left            =   705
      TabIndex        =   40
      Top             =   1600
      Width           =   900
   End
   Begin VB.Label lblLab21 
      BackStyle       =   0  'Transparent
      Caption         =   "Mahnstufe :"
      Height          =   210
      Left            =   3860
      TabIndex        =   39
      Top             =   1600
      Width           =   900
   End
   Begin VB.Label lblLab23 
      BackStyle       =   0  'Transparent
      Caption         =   "Mahnfrist :"
      Height          =   210
      Left            =   5465
      TabIndex        =   38
      Top             =   200
      Width           =   900
   End
   Begin VB.Label lblLab25 
      BackStyle       =   0  'Transparent
      Caption         =   "Betrag :"
      Height          =   210
      Left            =   705
      TabIndex        =   37
      Top             =   900
      Width           =   900
   End
   Begin VB.Label lblLab26 
      BackStyle       =   0  'Transparent
      Caption         =   "Offen :"
      Height          =   210
      Left            =   2305
      TabIndex        =   36
      Top             =   900
      Width           =   900
   End
   Begin VB.Label lblLab43 
      BackStyle       =   0  'Transparent
      Caption         =   "Kommentar :"
      Height          =   210
      Left            =   705
      TabIndex        =   35
      Top             =   5100
      Width           =   900
   End
   Begin VB.Label lblLab44 
      BackStyle       =   0  'Transparent
      Caption         =   "Gebühr :"
      Height          =   210
      Left            =   5465
      TabIndex        =   34
      Top             =   900
      Width           =   900
   End
   Begin VB.Label lblLab60 
      BackStyle       =   0  'Transparent
      Caption         =   "Mandant :"
      Height          =   210
      Left            =   705
      TabIndex        =   33
      Top             =   2300
      Width           =   900
   End
End
Attribute VB_Name = "frmOPEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmWar As XtremeSuiteControls.ComboBox
Private TxRen As XtremeSuiteControls.FlatEdit
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private MoKal As XtremeCalendarControl.DatePicker

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
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
    TxDa1.SetFocus
End If

Set MoKal = Nothing

Exit Sub

OrErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDatu " & Err.Number
Resume Next

End Sub
Private Sub FDaKo()
On Error GoTo LaErr
'Kontrolloert und formatiert das Eingabedatum neu

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        NeuDa = CDate(TxDa1.Text)
        TxDa1.Text = Format$(NeuDa, "dd.mm.yyyy")
        With MoKal
            .EnsureVisible NeuDa - 30
            .Select NeuDa
            .SelectRange NeuDa, NeuDa
        End With
        If NeuDa > Date Then
            SPopu NeuDa & " liegt in der Zukunft!", "Der Tag " & NeuDa & " liegt in der Zukunft", IC48_Information
        End If
    End If
End If

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FDaKo " & Err.Number
Resume Next

End Sub
Private Sub FNeu()
On Error GoTo SuErr

Dim ReStr As String

Set FM = frmOPEdit

If FM.txtRechn.Text <> vbNullString Then
    ReStr = FM.txtRechn.Text
    FM.txtRechn.Text = ReStr
End If
FM.txtPatie.Text = vbNullString
FM.txtPatie.Enabled = False
FM.txtOPDat.Text = Date
FM.txtOPBet.Text = GlWa2
FM.txtOPBez.Text = GlWa2
FM.txtFalli.Text = Date + 30
FM.txtOffen.Text = GlWa2
FM.cmbStufe.ListIndex = 0
FM.txtOPBel.Text = "0"
FM.txtOpKom.Text = vbNullString
FM.txtDatu1.Text = Date
FM.txtStKto.Text = vbNullString
FM.txtStBLZ.Text = vbNullString
FM.txtIBANr.Text = vbNullString
FM.txtBICNr.Text = vbNullString
FM.txtManID.Text = vbNullString
FM.cmbMahnb.ListIndex = 1
FM.cmbWarun.ListIndex = GlStW - 1
FM.txtPatie.Enabled = False
FM.btnPatSu.Enabled = True

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FNeu " & Err.Number
Resume Next

End Sub
Private Sub FSeKy(ByVal vkKey As Long)
    keybd_event vkKey, 0, 0, 0
    keybd_event vkKey, 0, KEYEVENTF_KEYUP, 0
End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim AktZa As Integer
Dim GesZa As Integer
Dim LauZa As Integer

Set FM = frmOPEdit
Set Rahm0 = Me.frmRahm0
Set MoKal = FM.dtpDatu1
Set CmMan = FM.cmbManda
Set CmMit = FM.cmbMitar
Set CmWar = FM.cmbWarun
Set PuBu1 = FM.btnDatu1
Set PuBu2 = FM.btnPatSu
Set TxDa1 = FM.txtDatu1
Set TxDa2 = FM.txtOPDat
Set TxDa3 = FM.txtFalli
Set TxRen = FM.txtRechn
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

With CmWar
    For AktZa = 1 To UBound(GlWar)
        .AddItem GlWar(AktZa, 1)
        .ItemData(AktZa - 1) = GlWar(AktZa, 0)
    Next AktZa
End With

With FM.cmbStufe
    .AddItem "Keine Mahnung"
    .ItemData(0) = 0
    .AddItem "1. Mahnung"
    .ItemData(1) = 1
    .AddItem "2. Mahnung"
    .ItemData(2) = 2
    .AddItem "3. Mahnung"
    .ItemData(3) = 3
    .AddItem "4. Mahnung"
    .ItemData(4) = 4
    .AddItem "5. Mahnung"
    .ItemData(5) = 5
    .AddItem "6. Mahnung"
    .ItemData(6) = 6
    .AddItem "7. Mahnung"
    .ItemData(7) = 7
End With

With FM.cmbMahnb
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

For AktZa = 1 To UBound(GlThe)
    If GlNeP = True Then 'neuer Posten
        If CBool(GlThe(AktZa, 25)) = False Then
            LauZa = LauZa + 1
            CmMan.AddItem GlThe(AktZa, 13)
            CmMan.ItemData(LauZa - 1) = GlThe(AktZa, 0)
        End If
    Else
        LauZa = LauZa + 1
        CmMan.AddItem GlThe(AktZa, 13)
        CmMan.ItemData(LauZa - 1) = GlThe(AktZa, 0)
    End If
Next AktZa

For AktZa = 1 To UBound(GlMiK) 'Alle Mitarbeiter
    CmMit.AddItem GlMiK(AktZa, 1)
    CmMit.ItemData(AktZa - 1) = GlMiK(AktZa, 2)
Next AktZa

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

TxDa2.SetMask "00.00.0000", "__.__.____"
TxDa3.SetMask "00.00.0000", "__.__.____"

Select Case GlRFm 'Rechnungsnummernformat
Case 2: TxRen.SetMask "00-00-000000", "__-__-______"
Case 3: TxRen.SetMask "0000-000000", "____-______"
Case 4: TxRen.SetMask "00-000000", "__-______"
End Select

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_IDCard_View, 16)

FM.txtOPBet.Enabled = GlNoM
FM.txtOPBez.Enabled = GlNoM
FM.txtOffen.Enabled = GlNoM
FM.txtGebue.Enabled = GlNoM

CmMan.Enabled = GlMaV 'Mandanten vorhanden
CmMit.Enabled = GlMiV

Me.BackColor = GlBak
Rahm0.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
Resume Next

End Sub
Private Sub FKale()
On Error GoTo LaErr
'Läßt den Kalender aufklappen

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set MoKal = Me.dtpDatu1

If TxDa1.Text <> vbNullString Then
    If IsDate(TxDa1.Text) = True Then
        NeuDa = CDate(TxDa1.Text)
    Else
        NeuDa = Date
    End If
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
            TxDa1.Text = .Selection.Blocks(0).DateBegin
        End If
    End If
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub btnBezal_Click()
    GlNeP = False
    Unload Me
    frmOPAusg.Show vbModal
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

TeTit = IniGetOpt("Hilfe", 50751)
TeMai = IniGetOpt("Hilfe", 50752)
TeInh = IniGetOpt("Hilfe", 50753)
TeFus = IniGetOpt("Hilfe", 50754)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub btnSchließ_Click()
    GlNeP = False
    Unload Me
End Sub
Private Sub btnWieter_Click()
    S_Save
    GlNeP = False
    Unload Me
End Sub

Private Sub cmbMahnb_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbMahnb_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.txtOpKom.SetFocus
    Case vbKeyUp: 'Me.cmbManda.SetFocus
    End Select
End Sub

Private Sub cmbManda_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub cmbManda_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.cmbMahnb.SetFocus
    Case vbKeyUp: 'Me.txtOPBez.SetFocus
    End Select
End Sub
Private Sub cmbStufe_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub cmbStufe_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDown: 'Me.txtOPBel.SetFocus
    Case vbKeyUp: 'Me.txtPatie.SetFocus
    End Select
End Sub
Private Sub btnPatSu_Click()
    frmAdrSuch.Show vbModal
End Sub

Private Sub cmbMitar_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
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

FInit
AFont Me
If GlNeP = True Then 'Neuer Posten
    FNeu
Else
    S_Posi
End If
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmOPEdit = Nothing
End Sub

Private Sub txtBICNr_GotFocus()
    Me.txtBICNr.SelStart = 0
    Me.txtBICNr.SelLength = Len(Me.txtBICNr.Text)
End Sub
Private Sub txtDatu1_LostFocus()
    FDaKo
End Sub

Private Sub txtFalli_GotFocus()
    Me.txtFalli.SelStart = 0
    Me.txtFalli.SelLength = Len(Me.txtFalli.Text)
End Sub
Private Sub txtFalli_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtFalli_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtFalli.SelLength = 0
    Case vbKeyDown: 'Me.txtOPBet.SetFocus
    Case vbKeyUp: 'Me.txtOPDat.SetFocus
    End Select
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
    Case vbKeyDown: 'Me.txtPatie.SetFocus
    Case vbKeyUp: 'Me.txtOffen.SetFocus
    End Select
End Sub

Private Sub txtGebue_GotFocus()
    Me.txtGebue.SelStart = 0
    Me.txtGebue.SelLength = Len(Me.txtGebue.Text)
End Sub

Private Sub txtGebue_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtGebue_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtGebue.SelLength = 0
    Case vbKeyDown: 'Me.txtPatie.SetFocus
    Case vbKeyUp: 'Me.cmbMahnb.SetFocus
    End Select
End Sub
Private Sub txtGlaID_GotFocus()
    Me.txtGlaID.SelStart = 0
    Me.txtGlaID.SelLength = Len(Me.txtGlaID.Text)
End Sub
Private Sub txtIBANr_GotFocus()
    Me.txtIBANr.SelStart = 0
    Me.txtIBANr.SelLength = Len(Me.txtIBANr.Text)
End Sub

Private Sub txtManID_GotFocus()
    Me.txtManID.SelStart = 0
    Me.txtManID.SelLength = Len(Me.txtManID.Text)
End Sub
Private Sub txtOffen_GotFocus()
    Me.txtOffen.SelStart = 0
    Me.txtOffen.SelLength = Len(Me.txtOffen.Text)
End Sub
Private Sub txtOffen_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtOffen_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtOffen.SelLength = 0
    Case vbKeyDown: 'Me.txtDatu1.SetFocus
    Case vbKeyUp: 'Me.txtOPBet.SetFocus
    End Select
End Sub
Private Sub txtOPBel_GotFocus()
    Me.txtOPBel.SelStart = 0
    Me.txtOPBel.SelLength = Len(Me.txtOPBel.Text)
End Sub
Private Sub txtOPBel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtOPBel_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtOPBel.SelLength = 0
    Case vbKeyDown: 'Me.cmbManda.SetFocus
    Case vbKeyUp: 'Me.cmbStufe.SetFocus
    End Select
End Sub
Private Sub txtOPBet_GotFocus()
    Me.txtOPBet.SelStart = 0
    Me.txtOPBet.SelLength = Len(Me.txtOPBet.Text)
End Sub
Private Sub txtOPBet_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtOPBet_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtOPBet.SelLength = 0
    Case vbKeyDown: 'Me.txtOffen.SetFocus
    Case vbKeyUp: 'Me.txtFalli.SetFocus
    End Select
End Sub

Private Sub txtOPBez_GotFocus()
    Me.txtOPBez.SelStart = 0
    Me.txtOPBez.SelLength = Len(Me.txtOPBez.Text)
End Sub

Private Sub txtOPBez_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

Private Sub txtOPBez_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtOPBez.SelLength = 0
    Case vbKeyDown: 'Me.txtOpKom.SetFocus
    Case vbKeyUp: 'Me.cmbMahnb.SetFocus
    End Select
End Sub

Private Sub txtOPDat_GotFocus()
    Me.txtOPDat.SelStart = 0
    Me.txtOPDat.SelLength = Len(Me.txtOPDat.Text)
End Sub
Private Sub txtOPDat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtOPDat_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtOPDat.SelLength = 0
    Case vbKeyDown: 'Me.txtFalli.SetFocus
    Case vbKeyUp: 'Me.txtRechn.SetFocus
    End Select
End Sub
Private Sub txtOpKom_GotFocus()
    Me.txtOpKom.SelStart = 0
    Me.txtOpKom.SelLength = Len(Me.txtOpKom.Text)
End Sub
Private Sub txtOpKom_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtOpKom_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtOpKom.SelLength = 0
    Case vbKeyDown: 'Me.txtDatu1.SetFocus
    Case vbKeyUp: 'Me.cmbMahnb.SetFocus
    End Select
End Sub
Private Sub txtPatie_GotFocus()
    Me.txtPatie.SelStart = 0
    Me.txtPatie.SelLength = Len(Me.txtPatie.Text)
End Sub
Private Sub txtPatie_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtPatie_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtPatie.SelLength = 0
    Case vbKeyDown: 'Me.cmbStufe.SetFocus
    Case vbKeyUp: 'Me.txtDatu1.SetFocus
    End Select
End Sub
Private Sub txtRechn_GotFocus()
    Me.txtRechn.SelStart = 0
    Me.txtRechn.SelLength = Len(Me.txtRechn.Text)
End Sub
Private Sub txtRechn_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub
Private Sub txtRechn_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2: Me.txtRechn.SelLength = 0
    Case vbKeyDown: 'Me.txtOPDat.SetFocus
    End Select
End Sub

Private Sub txtStBLZ_GotFocus()
    Me.txtStBLZ.SelStart = 0
    Me.txtStBLZ.SelLength = Len(Me.txtStBLZ.Text)
End Sub
Private Sub txtStKto_GotFocus()
    Me.txtStKto.SelStart = 0
    Me.txtStKto.SelLength = Len(Me.txtStKto.Text)
End Sub
Private Sub cmbWarun_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        FSeKy vbKeyTab
    End If
End Sub

