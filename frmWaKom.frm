VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmWaKom 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Terminwartelistenkommentar"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   14
      Top             =   3100
      Width           =   6400
      _Version        =   1048579
      _ExtentX        =   11289
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   4400
         TabIndex        =   11
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
         Left            =   3000
         TabIndex        =   10
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
         Left            =   1600
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Zurück"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnHilfe 
         Height          =   400
         Left            =   300
         TabIndex        =   8
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
      Height          =   3000
      Left            =   300
      TabIndex        =   2
      Top             =   100
      Width           =   5700
      _Version        =   1048579
      _ExtentX        =   10054
      _ExtentY        =   5292
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkMitar 
         Height          =   250
         Left            =   1000
         TabIndex        =   1
         Top             =   760
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mitarbeiter"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkManda 
         Height          =   250
         Left            =   1000
         TabIndex        =   5
         Top             =   1440
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   441
         _StockProps     =   79
         Caption         =   "Mandant"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtDummy 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H8000000F&
         Height          =   200
         Left            =   0
         TabIndex        =   0
         Top             =   4000
         Width           =   80
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   1000
         TabIndex        =   3
         Top             =   1030
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbPrior 
         Height          =   315
         Left            =   1000
         TabIndex        =   4
         Top             =   330
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtKomme 
         Height          =   350
         Left            =   1000
         TabIndex        =   7
         Top             =   2430
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   315
         Left            =   1000
         TabIndex        =   6
         Top             =   1730
         Width           =   3600
         _Version        =   1048579
         _ExtentX        =   6350
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   210
         Left            =   1000
         TabIndex        =   13
         Top             =   2180
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3528
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Kommentar :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   210
         Left            =   1000
         TabIndex        =   12
         Top             =   80
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Priorität :"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FormExtender frmExtde 
      Left            =   0
      Top             =   0
      _Version        =   1048579
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmWaKom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private CmPri As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private ChMan As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private PuBu1 As XtremeSuiteControls.PushButton
Private FTex1 As XtremeSuiteControls.FlatEdit
Private WithEvents FrmEx As XtremeSuiteControls.FormExtender
Attribute FrmEx.VB_VarHelpID = -1

Private clFen As clsFenster

Public PatNr As Long
Private Sub FLoad()
On Error GoTo InErr

Dim MitNr As Long
Dim ManNr As Long
Dim HiStr As String
Dim Prior As Integer
Dim AktZa As Integer
Dim MaSet As Boolean
Dim MiSet As Boolean

Set Rahm0 = Me.frmRahm0
Set Rahm1 = Me.frmRahm1
Set FTex1 = Me.txtKomme
Set CmPri = Me.cmbPrior
Set CmMit = Me.cmbMitar
Set CmMan = Me.cmbManda
Set ChMan = Me.chkManda
Set ChMit = Me.chkMitar

MaSet = CBool(IniGetVal("TerSys", "WaMaSe"))
MiSet = CBool(IniGetVal("TerSys", "WaMiSe"))

S_AdDe PatNr 'Adressendetails
With GlADt
    Prior = .AdPri
    HiStr = .AdHin
    MitNr = .AdMit
    ManNr = .AdMan
End With

If MitNr = 0 Then
    MitNr = GlMiA(GlSmI, 2)
End If

If ManNr = 0 Then
    ManNr = GlMan(GlSMa, 2)
End If

If Prior = 0 Then
    Prior = 2
End If

With CmPri
    .AddItem "Hoch"
    .ItemData(0) = 1
    .AddItem "Normal"
    .ItemData(1) = 2
    .AddItem "Niedrig"
    .ItemData(2) = 3
    .ListIndex = Prior - 1
End With

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
        CmMit.AddItem GlMiT(AktZa, 1)
        CmMit.ItemData(AktZa - 1) = GlMiT(AktZa, 2)
    Next AktZa
    For AktZa = 1 To UBound(GlMan)
        CmMan.AddItem GlMan(AktZa, 1)
        CmMan.ItemData(AktZa - 1) = GlMan(AktZa, 2)
    Next AktZa
Else
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        CmMit.AddItem GlMiA(AktZa, 1)
        CmMit.ItemData(AktZa - 1) = GlMiA(AktZa, 2)
    Next AktZa
    For AktZa = 1 To UBound(GlMaT)
        CmMan.AddItem GlMaT(AktZa, 1)
        CmMan.ItemData(AktZa - 1) = GlMaT(AktZa, 2)
    Next AktZa
End If

If GlMPl = True Then 'Mitarbeiterplan anstelle von Mandantenplan
    For AktZa = 1 To UBound(GlMiT) 'Aktive Mitarbeiter + Terminspalte
        If MitNr = GlMiT(AktZa, 2) Then
            CmMit.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
    For AktZa = 1 To UBound(GlMan)
        If ManNr = GlMan(AktZa, 2) Then
            CmMan.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
Else
    For AktZa = 1 To UBound(GlMiA) 'Aktive Mitarbeiter
        If MitNr = GlMiA(AktZa, 2) Then
            CmMit.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
    For AktZa = 1 To UBound(GlMaT)
        If ManNr = GlMaT(AktZa, 2) Then
            CmMan.ListIndex = AktZa - 1
            Exit For
        End If
    Next AktZa
End If

If CmMit.ListIndex < 0 Then
    CmMit.ListIndex = GlSmI - 1
End If

If CmMan.ListIndex < 0 Then
    CmMan.ListIndex = GlSMa - 1
End If

FTex1.Text = HiStr

If MaSet = True Then
    ChMan.Value = xtpChecked
    CmMan.Enabled = True
End If

If MiSet = True Then
    ChMit.Value = xtpChecked
    CmMit.Enabled = True
End If

Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
ChMan.BackColor = GlBak
ChMit.BackColor = GlBak

Exit Sub

InErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FLoad " & Err.Number
Resume Next

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
On Error GoTo SaErr

Dim MitNr As Long
Dim ManNr As Long
Dim HiStr As String
Dim Prior As Integer

Set FTex1 = Me.txtKomme
Set CmPri = Me.cmbPrior
Set CmMit = Me.cmbMitar
Set CmMan = Me.cmbManda

ManNr = CmMan.ItemData(CmMan.ListIndex)
MitNr = CmMit.ItemData(CmMit.ListIndex)
Prior = CmPri.ItemData(CmPri.ListIndex)

Ter_Edi PatNr, True, MitNr, ManNr, Prior 'in Warteliste aufnehmen
DoEvents

If FTex1.Text <> vbNullString Then
    HiStr = FTex1.Text
    If Len(HiStr) > 200 Then
        HiStr = Left$(HiStr, 200)
    End If
    S_WaKo HiStr, PatNr
    DoEvents
End If

P_List "TeDe", 0, 2
DoEvents

Unload Me

Exit Sub

SaErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FWart " & Err.Number
Resume Next

End Sub

Private Sub chkManda_Click()
On Error Resume Next

Dim MaSet As Boolean

Set CmMan = Me.cmbManda
Set ChMan = Me.chkManda

If ChMan.Value = xtpChecked Then
    MaSet = True
    CmMan.Enabled = True
Else
    MaSet = False
    CmMan.Enabled = False
End If

IniSetVal "TerSys", "WaMaSe", MaSet

End Sub

Private Sub chkMitar_Click()
On Error Resume Next

Dim MiSet As Boolean

Set CmMit = Me.cmbMitar
Set ChMit = Me.chkMitar

If ChMit.Value = xtpChecked Then
    MiSet = True
    CmMit.Enabled = True
Else
    MiSet = False
    CmMit.Enabled = False
End If

IniSetVal "TerSys", "WaMiSe", MiSet

End Sub
Private Sub Form_Load()
On Error Resume Next

Set FrmEx = Me.frmExtde

Me.BackColor = GlBak

AFont Me
SFrame 1, Me.hwnd

FrmEx.TopMost = True

If PatNr = 0 Then
    PatNr = GlAdr
End If

FLoad

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmWaKom = Nothing
End Sub

