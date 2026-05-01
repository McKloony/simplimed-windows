VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmReGen 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Auftragserstellung"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   10
      Top             =   5100
      Width           =   5500
      _Version        =   1048579
      _ExtentX        =   9701
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   3500
         TabIndex        =   13
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
         Left            =   2100
         TabIndex        =   12
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
         Left            =   800
         TabIndex        =   11
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
      Height          =   5100
      Left            =   500
      TabIndex        =   1
      Top             =   0
      Width           =   4500
      _Version        =   1048579
      _ExtentX        =   7937
      _ExtentY        =   8996
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkAuDia 
         Height          =   220
         Left            =   600
         TabIndex        =   3
         Top             =   1400
         Width           =   3400
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Aufnahmediagnosen übernehmen"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkKrAbs 
         Height          =   220
         Left            =   600
         TabIndex        =   2
         Top             =   1000
         Width           =   3400
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Krankenblatteinträge abschließen"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkKraOf 
         Height          =   225
         Left            =   600
         TabIndex        =   7
         Top             =   3000
         Width           =   3405
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Krankenblatteinträge wieder öffnen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbReTyp 
         Height          =   315
         Left            =   600
         TabIndex        =   8
         Top             =   3700
         Width           =   2300
         _Version        =   1048579
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkReDia 
         Height          =   220
         Left            =   600
         TabIndex        =   5
         Top             =   2200
         Width           =   3400
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Rechnungsdiagnosen übernehmen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKrDia 
         Height          =   225
         Left            =   600
         TabIndex        =   4
         Top             =   1800
         Width           =   3405
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Krankenblattdiagnosen übernehmen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbBehan 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   4500
         Width           =   3300
         _Version        =   1048579
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox3"
      End
      Begin XtremeSuiteControls.CheckBox chkReNum 
         Height          =   220
         Left            =   600
         TabIndex        =   6
         Top             =   2600
         Width           =   3400
         _Version        =   1048579
         _ExtentX        =   5997
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Belegnummern jetzt erzeugen"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Mandant :"
         Height          =   225
         Left            =   620
         TabIndex        =   16
         Top             =   4240
         Width           =   1395
      End
      Begin VB.Label lblLab48 
         BackStyle       =   0  'Transparent
         Caption         =   "Belegtyp :"
         Height          =   210
         Left            =   620
         TabIndex        =   15
         Top             =   3440
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bitte legen Sie die gewünschten Optionen fest und klicken auf Weiter, um die Rechnungen erstellen zu lassen."
         Height          =   580
         Left            =   300
         TabIndex        =   14
         Top             =   100
         Width           =   4200
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6800
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
Attribute VB_Name = "frmReGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private CheAb As XtremeSuiteControls.CheckBox
Private ChAuD As XtremeSuiteControls.CheckBox
Private ChKrD As XtremeSuiteControls.CheckBox
Private ChReD As XtremeSuiteControls.CheckBox
Private CheOf As XtremeSuiteControls.CheckBox
Private ChRen As XtremeSuiteControls.CheckBox
Private CmTyp As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox

Private ImMan As XtremeCommandBars.ImageManager
Private RpSel As XtremeReportControl.ReportSelectedRows
Private RpRow As XtremeReportControl.ReportRow

Private clFen As clsFenster

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
Private Sub btnWeiter_Click()
    TAbs
End Sub

Private Sub chkKrAbs_Click()
On Error Resume Next

Set CheAb = Me.chkKrAbs
Set ChAuD = Me.chkAuDia
Set ChKrD = Me.chkKrDia
Set CheOf = Me.chkKraOf
Set ChReD = Me.chkReDia

If CheAb.Value = xtpChecked Then
    CheOf.Value = xtpUnchecked
    CheOf.Enabled = False
Else
    CheOf.Enabled = True
End If

End Sub

Private Sub chkKraOf_Click()
On Error Resume Next

Set CheAb = Me.chkKrAbs
Set ChAuD = Me.chkAuDia
Set ChKrD = Me.chkKrDia
Set CheOf = Me.chkKraOf
Set ChReD = Me.chkReDia

If CheOf.Value = xtpChecked Then
    CheAb.Value = xtpUnchecked
    ChAuD.Value = xtpUnchecked
    ChKrD.Value = xtpUnchecked
    ChReD.Value = xtpUnchecked
    CheAb.Enabled = False
    ChAuD.Enabled = False
    ChKrD.Enabled = False
    ChReD.Enabled = False
Else
    CheAb.Enabled = True
    ChAuD.Enabled = True
    ChKrD.Enabled = True
    ChReD.Enabled = True
End If

End Sub

Private Sub chkKrDia_Click()
On Error Resume Next

Set ChKrD = Me.chkKrDia

If ChKrD.Value = xtpChecked Then
    IniSetVal "System", "ReGeKd", -1
Else
    IniSetVal "System", "ReGeKd", 0
End If

End Sub
Private Sub chkReDia_Click()
On Error Resume Next

Set ChReD = Me.chkReDia

If ChReD.Value = xtpChecked Then
    IniSetVal "System", "ReGeRd", -1
Else
    IniSetVal "System", "ReGeRd", 0
End If

End Sub
Private Sub chkReNum_Click()
On Error Resume Next

Set ChRen = Me.chkReNum

If ChRen.Value = xtpChecked Then
    IniSetVal "System", "ReGeNu", -1
Else
    IniSetVal "System", "ReGeNu", 0
End If

End Sub

Private Sub Form_Load()
On Error Resume Next

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

TInit
AFont Me

clFen.FenVor

Set clFen = Nothing

SFrame 1, Me.hwnd
  
End Sub

Private Sub TInit()
On Error GoTo SuErr

Dim AktZa As Integer
Dim ReNuG As Boolean
Dim ReDiU As Boolean
Dim KrDiU As Boolean

Set FM = frmMain
Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set CheAb = Me.chkKrAbs
Set ChAuD = Me.chkAuDia
Set ChKrD = Me.chkKrDia
Set CheOf = Me.chkKraOf
Set ChRen = Me.chkReNum
Set CmTyp = Me.cmbReTyp
Set CmMan = Me.cmbBehan
Set ChReD = Me.chkReDia

ReNuG = CBool(IniGetVal("System", "ReGeNu"))
ReDiU = CBool(IniGetVal("System", "ReGeRd"))
KrDiU = CBool(IniGetVal("System", "ReGeKd"))

With CmTyp
    .AddItem "R - Standardrechnung"
    .ItemData(0) = 1
    .AddItem "V - Kostenvoranschlag"
    .ItemData(1) = 2
    .AddItem "L - Laborrechnung"
    .ItemData(2) = 3
    .AddItem "A - Abrechnungsstelle"
    .ItemData(3) = 4
    .AddItem "U - Gutschrift"
    .ItemData(4) = 5
    .AddItem "M - Rechnungsauftrag"
    .ItemData(5) = 6
    .AddItem "G - Gewerberechnung"
    .ItemData(6) = 7
    .AddItem "I - Importrechnung"
    .ItemData(7) = 8
    .ListIndex = 5
End With

For AktZa = 1 To UBound(GlMan)
    With CmMan
        .AddItem GlMan(AktZa, 1)
        .ItemData(AktZa - 1) = GlMan(AktZa, 2)
    End With
Next AktZa
CmMan.ListIndex = GlMan(GlSMa, 0) - 1

If ReNuG = True Then
    ChRen.Value = xtpChecked
Else
    ChRen.Value = xtpUnchecked
End If

If ReDiU = True Then
    ChReD.Value = xtpChecked
Else
    ChReD.Value = xtpUnchecked
End If

If KrDiU = True Then
    ChKrD.Value = xtpChecked
Else
    ChKrD.Value = xtpUnchecked
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
CheAb.BackColor = GlBak
ChAuD.BackColor = GlBak
ChKrD.BackColor = GlBak
CheOf.BackColor = GlBak
ChReD.BackColor = GlBak
ChRen.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TInit " & Err.Number
Resume Next

End Sub
Private Sub TAbs()
On Error GoTo SuErr

Dim ManNr As Long
Dim TyStr As String
Dim KrAbs As Boolean
Dim AuDia As Boolean
Dim KrDia As Boolean
Dim KrNur As Boolean
Dim KrOff As Boolean
Dim ReDia As Boolean
Dim ReErz As Boolean
Dim ReTyp As Integer

Set CheAb = Me.chkKrAbs
Set ChAuD = Me.chkAuDia
Set ChKrD = Me.chkKrDia
Set CheOf = Me.chkKraOf
Set CmTyp = Me.cmbReTyp
Set ChReD = Me.chkReDia
Set ChRen = Me.chkReNum
Set CmMan = Me.cmbBehan

If CheAb.Value = xtpChecked Then KrAbs = True
If ChAuD.Value = xtpChecked Then AuDia = True
If ChKrD.Value = xtpChecked Then KrDia = True
If CheOf.Value = xtpChecked Then KrOff = True
If ChReD.Value = xtpChecked Then ReDia = True
If ChRen.Value = xtpChecked Then ReErz = True

ReTyp = CmTyp.ListIndex

Select Case ReTyp
Case 0: TyStr = "R"
Case 1: TyStr = "V"
Case 2: TyStr = "L"
Case 3: TyStr = "A"
Case 4: TyStr = "U"
Case 5: TyStr = "M"
Case 6: TyStr = "K"
Case 7: TyStr = "I"
End Select

ManNr = CmMan.ItemData(CmMan.ListIndex)

Unload Me

If KrOff = True Then 'Rechnungen wieder rückgänfig machen
    S_VoOp
Else
    S_VoRe KrAbs, AuDia, TyStr, ReDia, KrDia, ManNr, ReErz
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TAbs " & Err.Number
Resume Next

End Sub

Private Sub btnSchließ_Click()
    Unload Me
End Sub

