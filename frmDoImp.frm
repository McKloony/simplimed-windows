VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmDoImp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Dokument Autoimport"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.PushButton btnButt1 
      Height          =   300
      Left            =   500
      TabIndex        =   3
      Top             =   520
      Width           =   300
      _Version        =   1048579
      _ExtentX        =   529
      _ExtentY        =   529
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
      Height          =   400
      Left            =   2900
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1900
      Width           =   1200
      _Version        =   1048579
      _ExtentX        =   2117
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Abbrechen"
      UseVisualStyle  =   -1  'True
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
   Begin XtremeSuiteControls.CheckBox chkEiBil 
      Height          =   220
      Left            =   800
      TabIndex        =   4
      Top             =   1200
      Width           =   3500
      _Version        =   1048579
      _ExtentX        =   6174
      _ExtentY        =   388
      _StockProps     =   79
      Caption         =   "Dialog automatisch schlieﬂen nach Import"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.Label lblLab01 
      Height          =   350
      Left            =   860
      TabIndex        =   1
      Top             =   500
      Width           =   3500
      _Version        =   1048579
      _ExtentX        =   6174
      _ExtentY        =   617
      _StockProps     =   79
      Caption         =   "Warte auf neues Dokument im Importordner..."
   End
End
Attribute VB_Name = "frmDoImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PuBu1 As XtremeSuiteControls.PushButton
Private CmChk As XtremeSuiteControls.CheckBox
Private LbLa1 As XtremeSuiteControls.Label
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Public FmAnz As Boolean
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub

Private Sub chkEiBil_Click()
On Error Resume Next

Set FM = frmDoImp
Set CmChk = FM.chkEiBil

If CmChk.Value = xtpChecked Then
    IniSetVal "System", "BiImCl", -1
    GlBCl = True
Else
    IniSetVal "System", "BiImCl", 0
    GlBCl = False
End If

End Sub
Private Sub Form_Load()
On Error Resume Next

Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmDoImp
Set PuBu1 = FM.btnButt1
Set CmChk = FM.chkEiBil
Set LbLa1 = FM.lblLab01
Set FrmEx = FM.frmExtde
Set ImMan = frmMain.imgManag

AFont FM

With PuBu1
    .Appearance = xtpAppearanceFlat
    .BackColor = GlBak
    .FlatStyle = True
    .Icon = ImMan.Icons.GetImage(IC16_Folder_Close, 16)
End With
    
If GlBCl = True Then
    CmChk.Value = xtpChecked
End If
    
FM.BackColor = GlBak
CmChk.BackColor = GlBak
LbLa1.BackColor = GlBak

FrmEx.TopMost = True

SFrame 1, FM.hwnd

If FmAnz = False Then
    TimInit 6, 1
Else
    TimInit 7, 1
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    TimEnde 6
    TimEnde 7
    Set frmDoImp = Nothing
End Sub
