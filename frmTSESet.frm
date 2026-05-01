VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmTSESet 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "TSE Setup"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm4 
      Height          =   1700
      Left            =   2310
      TabIndex        =   37
      Top             =   5500
      Width           =   1900
      _Version        =   1048579
      _ExtentX        =   3351
      _ExtentY        =   2999
      _StockProps     =   79
      Caption         =   "Selbsttest"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnSeTes 
         Height          =   350
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "Selbsttest"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTesID 
         Height          =   300
         Left            =   240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   600
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLab13 
         Height          =   210
         Left            =   250
         TabIndex        =   38
         Top             =   360
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Test-ID:"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm3 
      Height          =   2500
      Left            =   2310
      TabIndex        =   31
      Top             =   2800
      Width           =   1900
      _Version        =   1048579
      _ExtentX        =   3351
      _ExtentY        =   4410
      _StockProps     =   79
      Caption         =   "Deblockieren:"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtFePIN 
         Height          =   300
         Left            =   240
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   600
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFePUK 
         Height          =   300
         Left            =   240
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnUnblo 
         Height          =   350
         Left            =   240
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1900
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "&PIN Erneuern"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   210
         Left            =   250
         TabIndex        =   35
         Top             =   1060
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Aktuelle PUK:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   210
         Left            =   250
         TabIndex        =   34
         Top             =   360
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Neue PIN:"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4400
      Left            =   150
      TabIndex        =   16
      Top             =   2800
      Width           =   1900
      _Version        =   1048579
      _ExtentX        =   3351
      _ExtentY        =   7761
      _StockProps     =   79
      Caption         =   "Neustart:"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnNeust 
         Height          =   350
         Left            =   240
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "TSE &Neustart"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFW_ID 
         Height          =   300
         Left            =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFWTyp 
         Height          =   300
         Left            =   240
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2000
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCusID 
         Height          =   300
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDrive 
         Height          =   300
         Left            =   240
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton btnReset 
         Height          =   350
         Left            =   240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   3800
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "TSE &Reset"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab12 
         Height          =   210
         Left            =   250
         TabIndex        =   20
         Top             =   2460
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "FW-ID:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab11 
         Height          =   210
         Left            =   250
         TabIndex        =   19
         Top             =   1760
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "FW-Type:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab10 
         Height          =   210
         Left            =   250
         TabIndex        =   18
         Top             =   1060
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Cust-ID:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab09 
         Height          =   210
         Left            =   250
         TabIndex        =   17
         Top             =   360
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Laufwerk:"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm2 
      Height          =   3600
      Left            =   4470
      TabIndex        =   3
      Top             =   2805
      Width           =   3940
      _Version        =   1048579
      _ExtentX        =   6950
      _ExtentY        =   6350
      _StockProps     =   79
      Caption         =   "Einrichtung:"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtTeAdm 
         Height          =   300
         Left            =   2300
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2000
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSePUK 
         Height          =   300
         Left            =   2300
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1300
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSePin 
         Height          =   300
         Left            =   2300
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   600
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtPuKey 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2000
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtSerNu 
         Height          =   300
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1300
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtKasse 
         Height          =   300
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   529
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton btnNeTSE 
         Height          =   345
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2500
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "TSE &Einrichten"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   345
         Left            =   2300
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2500
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "&Lese PIN/PUK"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   345
         Left            =   2300
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3010
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "Enable CTSS"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   345
         Left            =   240
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3010
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   79
         Caption         =   "Init Schreiben"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   210
         Left            =   250
         TabIndex        =   15
         Top             =   360
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Kassenname:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   210
         Left            =   250
         TabIndex        =   14
         Top             =   1060
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Seriennummer:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   210
         Left            =   250
         TabIndex        =   13
         Top             =   1760
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Public-Key:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   210
         Left            =   2310
         TabIndex        =   12
         Top             =   360
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "PIN:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab07 
         Height          =   210
         Left            =   2310
         TabIndex        =   11
         Top             =   1060
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "PUK:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab08 
         Height          =   210
         Left            =   2310
         TabIndex        =   10
         Top             =   1760
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "TimeAdmin:"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
      Height          =   350
      Left            =   6720
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   79
      Caption         =   "&Schlieﬂen"
      UseVisualStyle  =   -1  'True
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
      Appearance      =   6
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtTSEIn 
      Height          =   2505
      Left            =   130
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   8320
      _Version        =   1048579
      _ExtentX        =   14676
      _ExtentY        =   4419
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
End
Attribute VB_Name = "frmTSESet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TSE Einstellungen und Wartung

Private FM As Form
Private AktCo As VB.Control
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Rahm2 As XtremeSuiteControls.GroupBox
Private Rahm3 As XtremeSuiteControls.GroupBox
Private Rahm4 As XtremeSuiteControls.GroupBox
Private TxTSE As XtremeSuiteControls.FlatEdit
Private TxPIN As XtremeSuiteControls.FlatEdit
Private TxPUK As XtremeSuiteControls.FlatEdit
Private TxKas As XtremeSuiteControls.FlatEdit
Private TxSer As XtremeSuiteControls.FlatEdit
Private TxKey As XtremeSuiteControls.FlatEdit
Private TxNPI As XtremeSuiteControls.FlatEdit
Private TxNPU As XtremeSuiteControls.FlatEdit
Private TxTim As XtremeSuiteControls.FlatEdit
Private TxDrv As XtremeSuiteControls.FlatEdit
Private TxCID As XtremeSuiteControls.FlatEdit
Private TxTyp As XtremeSuiteControls.FlatEdit
Private TxFID As XtremeSuiteControls.FlatEdit
Private TxTID As XtremeSuiteControls.FlatEdit

Private clFen As clsFenster

Private Sub TSE_Check()
On Error GoTo SuErr

Set FM = frmTSESet
Set Rahm1 = FM.frmRahm1
Set Rahm2 = FM.frmRahm2
Set Rahm3 = FM.frmRahm3
Set Rahm4 = FM.frmRahm4
Set TxTSE = FM.txtTSEIn
Set TxPIN = FM.txtFePIN
Set TxPUK = FM.txtFePUK
Set TxKas = FM.txtKasse
Set TxSer = FM.txtSerNu
Set TxKey = FM.txtPuKey
Set TxNPI = FM.txtSePin
Set TxNPU = FM.txtSePUK
Set TxTim = FM.txtTeAdm
Set TxDrv = FM.txtDrive
Set TxCID = FM.txtCusID
Set TxTyp = FM.txtFWTyp
Set TxFID = FM.txtFW_ID
Set TxTID = FM.txtTesID

TxTSE.Text = "Hallo1"
TxTSE.Text = TxTSE.Text & vbCrLf & "Hallo2"

TxPIN.Text = "12345"
TxPUK.Text = "123456"

TxKas.Text = "Kasse1"
TxNPI.Text = "12345"
TxNPU.Text = "123456"
TxTim.Text = "789789"

TxTID.Text = "SimpliMed"


Rahm1.BackColor = GlBak
Rahm2.BackColor = GlBak
Rahm3.BackColor = GlBak
Rahm4.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg Then MsgBox Err.Description, 48, "TSE_Check " & Err.Number
Resume Next

End Sub
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

Me.BackColor = GlBak

AFont Me

clFen.FenVor

Set clFen = Nothing

SFrame 1, Me.hwnd

TSE_Check

End Sub

Private Sub txtCusID_GotFocus()
    Me.txtCusID.SelStart = 0
    Me.txtCusID.SelLength = Len(Me.txtCusID.Text)
End Sub
Private Sub txtDrive_GotFocus()
    Me.txtDrive.SelStart = 0
    Me.txtDrive.SelLength = Len(Me.txtFePIN.Text)
End Sub
Private Sub txtFePIN_GotFocus()
    Me.txtFePIN.SelStart = 0
    Me.txtFePIN.SelLength = Len(Me.txtFePIN.Text)
End Sub

Private Sub txtFePUK_GotFocus()
    Me.txtFePUK.SelStart = 0
    Me.txtFePUK.SelLength = Len(Me.txtFePUK.Text)
End Sub

Private Sub txtFW_ID_GotFocus()
    Me.txtFW_ID.SelStart = 0
    Me.txtFW_ID.SelLength = Len(Me.txtFW_ID.Text)
End Sub
Private Sub txtFWTyp_GotFocus()
    Me.txtFWTyp.SelStart = 0
    Me.txtFWTyp.SelLength = Len(Me.txtFWTyp.Text)
End Sub
Private Sub txtKasse_GotFocus()
    Me.txtKasse.SelStart = 0
    Me.txtKasse.SelLength = Len(Me.txtKasse.Text)
End Sub

Private Sub txtPuKey_GotFocus()
    Me.txtPuKey.SelStart = 0
    Me.txtPuKey.SelLength = Len(Me.txtPuKey.Text)
End Sub

Private Sub txtSePin_GotFocus()
    Me.txtSePin.SelStart = 0
    Me.txtSePin.SelLength = Len(Me.txtSePin.Text)
End Sub

Private Sub txtSePUK_GotFocus()
    Me.txtSePUK.SelStart = 0
    Me.txtSePUK.SelLength = Len(Me.txtSePUK.Text)
End Sub
Private Sub txtSerNu_GotFocus()
    Me.txtSerNu.SelStart = 0
    Me.txtSerNu.SelLength = Len(Me.txtSerNu.Text)
End Sub

Private Sub txtTeAdm_GotFocus()
    Me.txtTeAdm.SelStart = 0
    Me.txtTeAdm.SelLength = Len(Me.txtTeAdm.Text)
End Sub

Private Sub txtTesID_GotFocus()
    Me.txtTesID.SelStart = 0
    Me.txtTesID.SelLength = Len(Me.txtTesID.Text)
End Sub
