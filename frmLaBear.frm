VERSION 5.00
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "Codejock.Calendar.v16.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmLaBear 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Bearbeiten"
   ClientHeight    =   5325
   ClientLeft      =   4050
   ClientTop       =   330
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   32
      Top             =   4200
      Width           =   7600
      _Version        =   1048579
      _ExtentX        =   13406
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   5600
         TabIndex        =   35
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
         TabIndex        =   34
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
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   2900
         TabIndex        =   33
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
   Begin XtremeCalendarControl.DatePicker dtpDatu1 
      Height          =   400
      Left            =   240
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   400
      _Version        =   1048579
      _ExtentX        =   706
      _ExtentY        =   706
      _StockProps     =   64
      Show3DBorder    =   2
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4300
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   7585
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkBefun 
         Height          =   220
         Left            =   2305
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3800
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Befundung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFerti 
         Height          =   220
         Left            =   705
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3800
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Fertiggestellt"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBeKom 
         Height          =   350
         Left            =   700
         TabIndex        =   29
         Top             =   3210
         Width           =   6300
         _Version        =   1048579
         _ExtentX        =   11112
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtPaVor 
         Height          =   350
         Left            =   2300
         TabIndex        =   26
         Top             =   1830
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtPaNam 
         Height          =   350
         Left            =   700
         TabIndex        =   25
         Top             =   1830
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton btnDatu2 
         Height          =   350
         Left            =   3420
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "÷ffnet den Auswahlkalender"
         Top             =   1130
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnDatu1 
         Height          =   350
         Left            =   1810
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "÷ffnet den Auswahlkalender"
         Top             =   1130
         Width           =   350
         _Version        =   1048579
         _ExtentX        =   617
         _ExtentY        =   617
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBefun 
         Height          =   350
         Left            =   5500
         TabIndex        =   24
         Top             =   1130
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtPaGeb 
         Height          =   350
         Left            =   3900
         TabIndex        =   23
         Top             =   1130
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu2 
         Height          =   350
         Left            =   2300
         TabIndex        =   21
         Top             =   1130
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu1 
         Height          =   350
         Left            =   700
         TabIndex        =   19
         Top             =   1130
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtLabNr 
         Height          =   350
         Left            =   2300
         TabIndex        =   17
         Top             =   430
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtAufNr 
         Height          =   350
         Left            =   700
         TabIndex        =   16
         Top             =   430
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbBeArt 
         Height          =   310
         Left            =   3900
         TabIndex        =   18
         Top             =   430
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
      Begin XtremeSuiteControls.ComboBox cbmBehan 
         Height          =   310
         Left            =   3900
         TabIndex        =   27
         Top             =   1830
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5477
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.ComboBox cmbLaKat 
         Height          =   310
         Left            =   3900
         TabIndex        =   28
         Top             =   2530
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5477
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Laborkatalog"
         Height          =   210
         Left            =   3905
         TabIndex        =   60
         Top             =   2300
         Width           =   1500
      End
      Begin VB.Label Lab01 
         BackStyle       =   0  'Transparent
         Caption         =   "Befundart"
         Height          =   210
         Left            =   3905
         TabIndex        =   58
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname Patient"
         Height          =   210
         Left            =   2305
         TabIndex        =   57
         Top             =   1600
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Patient"
         Height          =   210
         Left            =   705
         TabIndex        =   56
         Top             =   1600
         Width           =   1500
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Labornummer"
         Height          =   210
         Left            =   2305
         TabIndex        =   55
         Top             =   200
         Width           =   1500
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ausgang"
         Height          =   210
         Left            =   2305
         TabIndex        =   54
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragsnummer"
         Height          =   210
         Left            =   705
         TabIndex        =   53
         Top             =   200
         Width           =   1500
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Eingang"
         Height          =   210
         Left            =   705
         TabIndex        =   52
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Lab05 
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren"
         Height          =   210
         Left            =   3905
         TabIndex        =   51
         Top             =   900
         Width           =   1200
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant"
         Height          =   210
         Left            =   3905
         TabIndex        =   50
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar"
         Height          =   210
         Left            =   705
         TabIndex        =   49
         Top             =   2980
         Width           =   1200
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Befundet"
         Height          =   210
         Left            =   5505
         TabIndex        =   48
         Top             =   900
         Width           =   1200
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   4300
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      _Version        =   1048579
      _ExtentX        =   13229
      _ExtentY        =   7585
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnDatu3 
         Height          =   310
         Left            =   3400
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   430
         Width           =   310
         _Version        =   1048579
         _ExtentX        =   547
         _ExtentY        =   547
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatu3 
         Height          =   310
         Left            =   2300
         TabIndex        =   2
         Top             =   430
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.CheckBox chkRechn 
         Height          =   220
         Left            =   3905
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3700
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnung erstellt"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkErste 
         Height          =   220
         Left            =   2305
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3700
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Bericht erstellt"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkBefng 
         Height          =   220
         Left            =   705
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3700
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Befundung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAuKom 
         Height          =   310
         Left            =   700
         TabIndex        =   12
         Top             =   3210
         Width           =   6300
         _Version        =   1048579
         _ExtentX        =   11112
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtPatGe 
         Height          =   310
         Left            =   2300
         TabIndex        =   9
         Top             =   1830
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtBefKo 
         Height          =   310
         Left            =   700
         TabIndex        =   8
         Top             =   1830
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtMulti 
         Height          =   310
         Left            =   3900
         TabIndex        =   7
         Top             =   1130
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtVorPa 
         Height          =   310
         Left            =   2300
         TabIndex        =   6
         Top             =   1130
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtNamPa 
         Height          =   310
         Left            =   700
         TabIndex        =   5
         Top             =   1130
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit txtAuNum 
         Height          =   310
         Left            =   700
         TabIndex        =   1
         Top             =   430
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbAuTyp 
         Height          =   310
         Left            =   3900
         TabIndex        =   4
         Top             =   430
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
      Begin XtremeSuiteControls.ComboBox cmbKunde 
         Height          =   310
         Left            =   3900
         TabIndex        =   10
         Top             =   1830
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5477
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.ComboBox cmbLabKa 
         Height          =   310
         Left            =   3900
         TabIndex        =   11
         Top             =   2530
         Width           =   3100
         _Version        =   1048579
         _ExtentX        =   5477
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Laborkatalog"
         Height          =   210
         Left            =   3905
         TabIndex        =   61
         Top             =   2300
         Width           =   1500
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragstyp"
         Height          =   210
         Left            =   3905
         TabIndex        =   46
         Top             =   200
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Vorname Patient"
         Height          =   210
         Left            =   2305
         TabIndex        =   45
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Name Patient"
         Height          =   210
         Left            =   705
         TabIndex        =   44
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragsnummer"
         Height          =   210
         Left            =   705
         TabIndex        =   43
         Top             =   200
         Width           =   1500
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragsdatum"
         Height          =   210
         Left            =   2305
         TabIndex        =   42
         Top             =   200
         Width           =   1500
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Befundkosten"
         Height          =   210
         Left            =   705
         TabIndex        =   41
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant"
         Height          =   210
         Left            =   3905
         TabIndex        =   40
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Geboren"
         Height          =   210
         Left            =   2305
         TabIndex        =   39
         Top             =   1600
         Width           =   1200
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Kommentar"
         Height          =   210
         Left            =   705
         TabIndex        =   38
         Top             =   2980
         Width           =   1200
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Steigerungsfaktor"
         Height          =   255
         Left            =   3905
         TabIndex        =   37
         Top             =   900
         Width           =   1305
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
Attribute VB_Name = "frmLaBear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private TxDa1 As XtremeSuiteControls.FlatEdit
Private TxDa2 As XtremeSuiteControls.FlatEdit
Private TxDa3 As XtremeSuiteControls.FlatEdit
Private TxGe1 As XtremeSuiteControls.FlatEdit
Private TxGe2 As XtremeSuiteControls.FlatEdit
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private CmBer As XtremeSuiteControls.ComboBox
Private CmAuf As XtremeSuiteControls.ComboBox
Private CmBeh As XtremeSuiteControls.ComboBox
Private CmKun As XtremeSuiteControls.ComboBox
Private CmPa1 As XtremeSuiteControls.ComboBox
Private CmPa2 As XtremeSuiteControls.ComboBox
Private CmKa1 As XtremeSuiteControls.ComboBox
Private CmKa2 As XtremeSuiteControls.ComboBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private PuBu2 As XtremeSuiteControls.PushButton
Private PuBu3 As XtremeSuiteControls.PushButton
Private ImMan As XtremeCommandBars.ImageManager
Private MoKal As XtremeCalendarControl.DatePicker
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow
Private RpCol As XtremeReportControl.ReportColumn

Private KalWa As Integer
Private clFen As clsFenster
Private Sub FDatu()
On Error GoTo OrErr

Dim NeuDa As Date

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
Set MoKal = Me.dtpDatu1

If MoKal.Selection.BlocksCount > 0 Then
    NeuDa = MoKal.Selection.Blocks(0).DateBegin
    Select Case KalWa
    Case 1: TxDa1.Text = NeuDa
            TxDa1.SetFocus
    Case 2: TxDa2.Text = NeuDa
            TxDa2.SetFocus
    Case 3: TxDa3.Text = NeuDa
            TxDa3.SetFocus
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

Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
Set MoKal = Me.dtpDatu1
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2

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
Case 3:
    If IsDate(TxDa3.Text) Then
        NeuDa = TxDa3.Text
    Else
        NeuDa = Date
    End If
End Select

With MoKal
    .EnsureVisible NeuDa
    .Select NeuDa
    .SelectRange NeuDa, NeuDa
    Select Case KalWa
    Case 1:
            .Top = Rahm1.Top + TxDa1.Top + TxDa1.Height
            .Left = Rahm1.Left + TxDa1.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa1.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 2:
            .Top = Rahm1.Top + TxDa2.Top + TxDa2.Height
            .Left = Rahm1.Left + TxDa2.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa2.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    Case 3:
            .Top = Rahm2.Top + TxDa3.Top + TxDa3.Height
            .Left = Rahm2.Left + TxDa3.Left
            If .ShowModal(1, 1) Then
                If .Selection.BlocksCount > 0 Then
                    TxDa3.Text = .Selection.Blocks(0).DateBegin
                End If
            End If
    End Select
End With

Set MoKal = Nothing

Exit Sub

LaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKale " & Err.Number
Resume Next

End Sub
Private Sub TInit()
On Error GoTo SuErr

Dim AktZa As Long

Set Rahm0 = Me.frmRahm0
Set TxDa1 = Me.txtDatu1
Set TxDa2 = Me.txtDatu2
Set TxDa3 = Me.txtDatu3
Set TxGe1 = Me.txtPaGeb
Set TxGe2 = Me.txtPatGe
Set CmBer = Me.cmbBeArt
Set CmAuf = Me.cmbAuTyp
Set CmBeh = Me.cbmBehan
Set MoKal = Me.dtpDatu1
Set CmKun = Me.cmbKunde
Set CmKa1 = Me.cmbLaKat
Set CmKa2 = Me.cmbLabKa
Set PuBu1 = Me.btnDatu1
Set PuBu2 = Me.btnDatu2
Set PuBu3 = Me.btnDatu3
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
    .ToolTipText = "Markieren Sie bitte hier den gw¸nschten Behandlungstag"
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

With CmBer
    .AddItem "Eigenbefund"
    .ItemData(0) = 1
    .AddItem "Endbefund"
    .ItemData(1) = 2
    .AddItem "Teilbefund"
    .ItemData(2) = 3
    .AddItem "Vorbefund"
    .ItemData(3) = 4
    .AddItem "Archivbefund"
    .ItemData(4) = 5
    .AddItem "Nachforderung"
    .ItemData(5) = 6
    .AddItem "Urlaub"
    .ItemData(6) = 7
End With

With CmAuf
    .AddItem "A"
    .ItemData(0) = 1
    .AddItem "B"
    .ItemData(1) = 2
    .AddItem "C"
    .ItemData(2) = 3
    .AddItem "D"
    .ItemData(3) = 4
    .AddItem "E"
    .ItemData(4) = 5
End With

For AktZa = 1 To UBound(GlThe)
    CmBeh.AddItem GlThe(AktZa, 13)
    CmBeh.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa
    
For AktZa = 1 To UBound(GlThe)
    CmKun.AddItem GlThe(AktZa, 13)
    CmKun.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlLab)
    CmKa1.AddItem GlLab(AktZa, 1)
    CmKa1.ItemData(AktZa - 1) = GlLab(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlLab)
    CmKa2.AddItem GlLab(AktZa, 1)
    CmKa2.ItemData(AktZa - 1) = GlLab(AktZa, 0)
Next AktZa

If CmBeh.Enabled = False Then
    CmBeh.Enabled = True
End If

If CmKun.Enabled = False Then
    CmKun.Enabled = True
End If

With TxDa1
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxDa2
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxDa3
    .SetMask "00.00.0000", "__.__.____"
    .Text = Date
End With

With TxGe1
    .SetMask "00.00.0000", "__.__.____"
End With

With TxGe2
    .SetMask "00.00.0000", "__.__.____"
End With

Rahm0.BackColor = GlBak
Me.chkFerti.BackColor = GlBak
Me.chkBefun.BackColor = GlBak
Me.chkBefun.BackColor = GlBak
Me.chkErste.BackColor = GlBak
Me.chkBefng.BackColor = GlBak
Me.chkRechn.BackColor = GlBak

PuBu1.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu2.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)
PuBu3.Icon = ImMan.Icons.GetImage(IC16_Calendar_Month, 16)

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TInit " & Err.Number
Resume Next

End Sub
Private Sub TLoad()
On Error GoTo SuErr

Dim IdxNr As Long
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set Rahm1 = Me.frmRahm1
Set Rahm2 = Me.frmRahm2
Set RpCo1 = FM.repCont1

Select Case GlBut
Case RibTab_LabBericht:
        Rahm1.Visible = True
        Set RpCo5 = FM.repCont5
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_LabBerichte:
        Rahm1.Visible = True
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
Case RibTab_LabAuftrag:
        Rahm2.Visible = True
        Set RpCo5 = FM.repCont5
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_LabAuftrage:
        Rahm2.Visible = True
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
End Select

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Lab_ID0)
        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
        S_LaDe False, IdxNr
    End If
End If

Me.BackColor = GlBak
Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo5 = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TLoad " & Err.Number
Resume Next

End Sub
Private Sub TWeit()
On Error GoTo SuErr

Dim IdxNr As Long
Dim RowNr As Long
Dim RpCo1 As XtremeReportControl.ReportControl
Dim RpCo5 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo1 = FM.repCont1
Set RpCo5 = FM.repCont5

Select Case GlBut
Case RibTab_LabBericht:
        Rahm1.Visible = True
        Set RpCo5 = FM.repCont5
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_LabBerichte:
        Rahm1.Visible = True
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
Case RibTab_LabAuftrag:
        Rahm2.Visible = True
        Set RpCo5 = FM.repCont5
        Set RpCls = RpCo5.Columns
        Set RpSel = RpCo5.SelectedRows
Case RibTab_LabAuftrage:
        Rahm2.Visible = True
        Set RpCo1 = FM.repCont1
        Set RpCls = RpCo1.Columns
        Set RpSel = RpCo1.SelectedRows
End Select

Set clFen = New clsFenster
clFen.hwnd = FM.hwnd

Screen.MousePointer = vbHourglass
clFen.FenDsk 2

If RpSel.Count > 0 Then
    Set RpRow = RpSel(0)
    If RpRow.GroupRow = False Then
        Set RpCol = RpCls.Find(Lab_ID0)
        IdxNr = RpRow.Record(RpCol.ItemIndex).Value
        RowNr = RpRow.Index
        S_LaDe True, IdxNr
        SUpLa RowNr
    End If
End If

clFen.FenDsk 3
Screen.MousePointer = vbNormal

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo1 = Nothing
Set RpCo5 = Nothing

Set clFen = Nothing

Unload Me

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "Tweit " & Err.Number
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
Private Sub btnDatu3_Click()
    KalWa = 3
    FKale
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

Private Sub btnWieter_Click()
    TWeit
End Sub

Private Sub cmbPati2_Click()

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

TInit
AFont Me
TLoad
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmLaBear = Nothing
End Sub

Private Sub txtAufNr_GotFocus()
    Me.txtAufNr.SelStart = 0
    Me.txtAufNr.SelLength = Len(Me.txtAufNr.Text)
End Sub

Private Sub txtAuNum_GotFocus()
    Me.txtAuNum.SelStart = 0
    Me.txtAuNum.SelLength = Len(Me.txtAuNum.Text)
End Sub

Private Sub txtDatu3_GotFocus()
    Me.txtDatu3.SelStart = 0
    Me.txtDatu3.SelLength = Len(Me.txtDatu3.Text)
End Sub
Private Sub txtLabNr_GotFocus()
    Me.txtLabNr.SelStart = 0
    Me.txtLabNr.SelLength = Len(Me.txtLabNr.Text)
End Sub
