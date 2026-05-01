VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReErs 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Rechnungsgenerator"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   4
      Top             =   2600
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
         TabIndex        =   7
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
         Left            =   2600
         TabIndex        =   6
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
         TabIndex        =   5
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
      Height          =   2500
      Left            =   0
      TabIndex        =   1
      Top             =   100
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   4410
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.RadioButton optRechn 
         Height          =   220
         Left            =   1700
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1500
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "in neue Rechnung einfügen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optPatie 
         Height          =   220
         Left            =   1700
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1100
         Width           =   3500
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "in vorhand. Rechnung einfügen"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   500
         Left            =   400
         TabIndex        =   8
         Top             =   100
         Width           =   5000
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   882
         _StockProps     =   79
         Caption         =   "Sollen die Leistungen in einer vorhandenen Rechnung oder einer neuen Rechnung eingefügt werden? Bitte wählen Sie eine Option."
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   4000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
End
Attribute VB_Name = "frmReErs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control

Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private Lab01 As XtremeSuiteControls.Label
Private Sub FKonf()
On Error GoTo InErr

Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set Opti1 = Me.optPatie
Set Opti2 = Me.optRechn
Set Lab01 = Me.lblLab01

Select Case GlBut
Case RibTab_LabBericht:
    Opti1.Caption = "in vorhand. Laborrechnung einfügen"
    Opti2.Caption = "in neue Laborrechnung einfügen"
    Lab01.Caption = "Sollen die Laborparameter in eine vorhandene Laborrechnung oder in eine neue Laborrechnung eingefügt werden?"
Case RibTab_LabBerichte:
    Opti1.Caption = "in vorhand. Laborrechnung einfügen"
    Opti2.Caption = "in neue Laborrechnung einfügen"
    Lab01.Caption = "Sollen die Laborparameter in eine vorhandene Laborrechnung oder in eine neue Laborrechnung eingefügt werden?"
Case RibTab_Ter_Listen:
    Opti1.Caption = "in vorhand. Rechnungen einfügen"
    Opti2.Caption = "in neue Rechnungen einfügen"
    Lab01.Caption = "Sollen die Leistungen in eine vorhandene Rechnung oder eine neue Rechnung eingefügt werden? Bitte wählen Sie eine Option."
End Select

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
Opti1.BackColor = GlBak
Opti2.BackColor = GlBak

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub FWeit()
On Error GoTo SuErr

Dim ReNeu As Boolean

Set Opti1 = Me.optPatie
Set Opti2 = Me.optRechn

ReNeu = Opti2.Value

Unload Me
DoEvents

Select Case GlBut
Case RibTab_Rechnungen: S_ReSer ReNeu
Case RibTab_Ter_Kalend: S_TeRec ReNeu
Case RibTab_Ter_Raeume: S_TeRec ReNeu
Case RibTab_Ter_Mitarb: S_TeRec ReNeu
Case RibTab_Ter_Listen: S_TeReL ReNeu
Case RibTab_LabBericht: S_LaReB ReNeu
Case RibTab_LabBerichte: S_LaReB ReNeu
End Select

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWeit " & Err.Number
Resume Next

End Sub

Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

Select Case GlBut
Case RibTab_Rechnungen:
    TeTit = IniGetOpt("Hilfe", 50861)
    TeMai = IniGetOpt("Hilfe", 50862)
    TeInh = IniGetOpt("Hilfe", 50863)
    TeFus = IniGetOpt("Hilfe", 50864)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
Case RibTab_LabBericht:
    TeTit = IniGetOpt("Hilfe", 50871)
    TeMai = IniGetOpt("Hilfe", 50872)
    TeInh = IniGetOpt("Hilfe", 50873)
    TeFus = IniGetOpt("Hilfe", 50874)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
Case RibTab_LabBerichte:
    TeTit = IniGetOpt("Hilfe", 50871)
    TeMai = IniGetOpt("Hilfe", 50872)
    TeInh = IniGetOpt("Hilfe", 50873)
    TeFus = IniGetOpt("Hilfe", 50874)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
Case Else:
    TeTit = IniGetOpt("Hilfe", 50851)
    TeMai = IniGetOpt("Hilfe", 50852)
    TeInh = IniGetOpt("Hilfe", 50853)
    TeFus = IniGetOpt("Hilfe", 50854)
    SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
End Select

End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FWeit
End Sub
Private Sub Form_Load()
On Error Resume Next

FKonf
AFont Me
SFrame 1, Me.hwnd

End Sub

