VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Status"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   ControlBox      =   0   'False
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.PushButton cmdButt1 
      Default         =   -1  'True
      Height          =   400
      Left            =   1700
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1300
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Abbrechen"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   -20
      TabIndex        =   0
      Text            =   "A"
      Top             =   3000
      Width           =   80
   End
   Begin XtremeSuiteControls.ProgressBar prbStat1 
      Height          =   350
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   3640
      _Version        =   1048579
      _ExtentX        =   6421
      _ExtentY        =   617
      _StockProps     =   93
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar prbStat2 
      Height          =   350
      Left            =   480
      TabIndex        =   2
      Top             =   1100
      Width           =   3640
      _Version        =   1048579
      _ExtentX        =   6421
      _ExtentY        =   617
      _StockProps     =   93
      UseVisualStyle  =   0   'False
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
      Height          =   210
      Left            =   500
      TabIndex        =   4
      Top             =   320
      Width           =   2960
      _Version        =   1048579
      _ExtentX        =   5221
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "Bitte warten..."
      Alignment       =   4
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private PrBr1 As XtremeSuiteControls.ProgressBar
Private PrBr2 As XtremeSuiteControls.ProgressBar
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private Sub cmdButt1_Click()
    Me.txtDummy.Text = "B"
    Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next

Set PrBr1 = Me.prbStat1
Set PrBr2 = Me.prbStat2
Set FrmEx = Me.frmExtde

FrmEx.TopMost = True

AFont Me

With PrBr1
    Select Case GlSty
    Case 8:
        .Appearance = xtpAppearanceOffice2013
        .UseVisualStyle = False
    Case 7:
        .Appearance = xtpAppearanceOffice2013
        .UseVisualStyle = False
    Case Else:
        .Appearance = xtpAppearanceResource
        .UseVisualStyle = True
    End Select
    .Scrolling = xtpProgressBarStandard
End With

With PrBr2
    Select Case GlSty
    Case 8:
        .Appearance = xtpAppearanceOffice2013
        .UseVisualStyle = False
    Case 7:
        .Appearance = xtpAppearanceOffice2013
        .UseVisualStyle = False
    Case Else:
        .Appearance = xtpAppearanceResource
        .UseVisualStyle = True
    End Select
    .Scrolling = xtpProgressBarStandard
End With
    
Me.BackColor = GlBak

Me.lblLab01.BackColor = GlBak

SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmStatus = Nothing
End Sub

Private Sub txtDummy_Change()
On Error Resume Next

Dim AktZa As Long

Set PrBr1 = Me.prbStat1
Set PrBr2 = Me.prbStat2

AktZa = Val(Me.txtDummy.Text)

PrBr1.Value = AktZa
PrBr2.Value = AktZa

End Sub
