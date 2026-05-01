VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmKaSuch 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Suchen"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   5
      Top             =   2500
      Width           =   5000
      _Version        =   1048579
      _ExtentX        =   8819
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   3000
         TabIndex        =   8
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
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   1600
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Suchen"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   300
         TabIndex        =   6
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
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H8000000F&
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   4000
      Width           =   80
   End
   Begin XtremeSuiteControls.FlatEdit txtSuch3 
      Height          =   350
      Left            =   2500
      TabIndex        =   4
      Top             =   1830
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2469
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.ComboBox cmbAusw1 
      Height          =   310
      Left            =   1000
      TabIndex        =   3
      Top             =   1830
      Width           =   1400
      _Version        =   1048579
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtSuch1 
      Height          =   350
      Left            =   1000
      TabIndex        =   1
      Top             =   430
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5115
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtSuch2 
      Height          =   350
      Left            =   1000
      TabIndex        =   2
      Top             =   1130
      Width           =   2900
      _Version        =   1048579
      _ExtentX        =   5115
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin VB.Label Lab03 
      BackStyle       =   0  'Transparent
      Caption         =   "Suche nach Preis"
      Height          =   195
      Left            =   1000
      TabIndex        =   11
      Top             =   1600
      Width           =   2900
   End
   Begin VB.Label Lab02 
      BackStyle       =   0  'Transparent
      Caption         =   "Suche nach Ziffer / Kürzel"
      Height          =   195
      Left            =   1000
      TabIndex        =   10
      Top             =   900
      Width           =   2900
   End
   Begin VB.Label Lab01 
      BackStyle       =   0  'Transparent
      Caption         =   "Suche nach Bezeichnung"
      Height          =   195
      Left            =   1000
      TabIndex        =   9
      Top             =   200
      Width           =   2900
   End
End
Attribute VB_Name = "frmKaSuch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private FTex1 As XtremeSuiteControls.FlatEdit
Private FTex2 As XtremeSuiteControls.FlatEdit
Private FTex3 As XtremeSuiteControls.FlatEdit
Private FCom1 As XtremeSuiteControls.ComboBox
Private CmSta As XtremeCommandBars.StatusBar
Private CmBar As XtremeCommandBars.CommandBar
Private CmPan As XtremeCommandBars.StatusBarPane
Private CmAcs As XtremeCommandBars.CommandBarActions
Private CmOpt As XtremeCommandBars.CommandBarsOptions

Private clFen As clsFenster
Private Sub TLoad()
On Error Resume Next

Set FTex1 = Me.txtSuch1
Set FTex2 = Me.txtSuch2
Set FTex3 = Me.txtSuch3
Set FCom1 = Me.cmbAusw1
Set Rahm0 = Me.frmRahm0

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

With FCom1
    .AddItem "ist größer als"
    .ItemData(.NewIndex) = 1
    .AddItem "ist kleiner als"
    .ItemData(.NewIndex) = 2
    .AddItem "ist größer gleich"
    .ItemData(.NewIndex) = 3
    .AddItem "ist kleiner gleich"
    .ItemData(.NewIndex) = 4
    .AddItem "ist gleich"
    .ItemData(.NewIndex) = 5
    .ListIndex = 0
End With

Me.BackColor = GlBak
Rahm0.BackColor = GlBak

clFen.FenVor

Set clFen = Nothing

End Sub
Private Sub TRes(ByVal Falg As Boolean)
On Error Resume Next

Set FTex1 = Me.txtSuch1
Set FTex2 = Me.txtSuch2
Set FTex3 = Me.txtSuch3
Set FCom1 = Me.cmbAusw1

FTex1.Text = vbNullString
FTex2.Text = vbNullString
FTex3.Text = vbNullString

If Falg = True Then
    FCom1.Enabled = True
Else
    FCom1.Enabled = False
End If

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
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    TSuda
End Sub
Private Sub Form_GotFocus()

Set FTex1 = Me.txtSuch1

FTex1.SetFocus

End Sub
Private Sub Form_Load()
On Error Resume Next

TLoad
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKaSuch = Nothing
End Sub
Private Sub txtSuch1_GotFocus()
    TRes False
End Sub
Private Sub txtSuch1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TSuda
End Sub
Private Sub TSuda()
On Error GoTo SuErr

Dim SuPar As String 'Suchparameter
Dim CmBrs As XtremeCommandBars.CommandBars

Set FTex1 = Me.txtSuch1
Set FTex2 = Me.txtSuch2
Set FTex3 = Me.txtSuch3
Set FCom1 = Me.cmbAusw1

Set FM = frmKetten
Set CmBrs = FM.comBar02
Set CmAcs = CmBrs.Actions

If FTex1.Text <> vbNullString Then
    EFilt 1, FTex1.Text
ElseIf FTex2.Text <> vbNullString Then
    EFilt 2, FTex2.Text
ElseIf FTex3.Text <> vbNullString Then
    Select Case FCom1.ListIndex
    Case 0: SuPar = ">"
    Case 1: SuPar = "<"
    Case 2: SuPar = ">="
    Case 3: SuPar = "<="
    Case 4: SuPar = "="
    End Select
    EFilt 3, FTex3.Text, SuPar
Else
    'CmAcs(KA_Erweit_Suchen).Checked = False
    EFilt 0
End If

Set CmBrs = Nothing

Unload Me

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TSuda " & Err.Number
Resume Next

End Sub
Private Sub txtSuch2_GotFocus()
    TRes False
End Sub
Private Sub txtSuch2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TSuda
End Sub
Private Sub txtSuch3_GotFocus()
    TRes True
End Sub
Private Sub txtSuch3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TSuda
End Sub

