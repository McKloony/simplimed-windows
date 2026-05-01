VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmLayout 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Layoutoptionen"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   17
      Top             =   4000
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAbbru 
         Height          =   400
         Left            =   6000
         TabIndex        =   16
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
         Left            =   4600
         TabIndex        =   15
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
      Begin XtremeSuiteControls.PushButton btnZuruck 
         Height          =   400
         Left            =   1900
         TabIndex        =   13
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
      Begin XtremeSuiteControls.PushButton btnLosche 
         Height          =   400
         Left            =   3200
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   400
         Width           =   1300
         _Version        =   1048579
         _ExtentX        =   2293
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Standardwerte"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   3900
      Left            =   0
      TabIndex        =   1
      Top             =   105
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   6879
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkFaMod 
         Height          =   225
         Left            =   4200
         TabIndex        =   10
         Top             =   2400
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Farbige Modulkennzeichnung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFaReg 
         Height          =   225
         Left            =   4200
         TabIndex        =   9
         Top             =   1950
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Farbige Modulregister"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkSysBu 
         Height          =   225
         Left            =   4200
         TabIndex        =   8
         Top             =   1500
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Runder Systembutton"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkOffRa 
         Height          =   225
         Left            =   4200
         TabIndex        =   7
         Top             =   1050
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Designfensterrahmen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkMenAn 
         Height          =   225
         Left            =   4200
         TabIndex        =   6
         Top             =   600
         Width           =   2805
         _Version        =   1048579
         _ExtentX        =   4939
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Menüanimation"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbFarbe 
         Height          =   315
         Left            =   1000
         TabIndex        =   3
         Top             =   1500
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
      Begin XtremeSuiteControls.ComboBox cmbSkins 
         Height          =   315
         Left            =   1000
         TabIndex        =   2
         Top             =   700
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
      Begin XtremeSuiteControls.ComboBox cmbTaskb 
         Height          =   315
         Left            =   1000
         TabIndex        =   4
         Top             =   2300
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
      Begin XtremeSuiteControls.ComboBox cmbFenst 
         Height          =   315
         Left            =   1000
         TabIndex        =   5
         Top             =   3100
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
      Begin XtremeSuiteControls.FlatEdit txtBreit 
         Height          =   350
         Left            =   4200
         TabIndex        =   11
         Top             =   3100
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtHoehe 
         Height          =   350
         Left            =   5200
         TabIndex        =   12
         Top             =   3100
         Width           =   800
         _Version        =   1048579
         _ExtentX        =   1411
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblLab06 
         Height          =   255
         Left            =   5220
         TabIndex        =   23
         Top             =   2850
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Höhe :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab05 
         Height          =   255
         Left            =   4220
         TabIndex        =   22
         Top             =   2850
         Width           =   900
         _Version        =   1048579
         _ExtentX        =   1587
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Breite :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab04 
         Height          =   255
         Left            =   1000
         TabIndex        =   21
         Top             =   2850
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fenstergröße :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab01 
         Height          =   225
         Left            =   1000
         TabIndex        =   20
         Top             =   460
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Skinset :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab02 
         Height          =   225
         Left            =   1000
         TabIndex        =   19
         Top             =   1260
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Farbschema :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLab03 
         Height          =   255
         Left            =   1000
         TabIndex        =   18
         Top             =   2040
         Width           =   2200
         _Version        =   1048579
         _ExtentX        =   3881
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Taskbaranzeige :"
         Transparent     =   -1  'True
      End
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   6200
      Width           =   80
   End
End
Attribute VB_Name = "frmLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private ChMen As XtremeSuiteControls.CheckBox
Private ChRah As XtremeSuiteControls.CheckBox
Private ChBut As XtremeSuiteControls.CheckBox
Private ChReg As XtremeSuiteControls.CheckBox
Private ChMod As XtremeSuiteControls.CheckBox
Private CmFar As XtremeSuiteControls.ComboBox
Private CmSkn As XtremeSuiteControls.ComboBox
Private CmTas As XtremeSuiteControls.ComboBox
Private CmFen As XtremeSuiteControls.ComboBox
Private TxBre As XtremeSuiteControls.FlatEdit
Private TxHoh As XtremeSuiteControls.FlatEdit
Private ImMan As XtremeCommandBars.ImageManager

Private FoLad As Boolean

Private clFil As clsFile
Private Sub ALayo()
On Error Resume Next

If IniGetSek(GlINI, "AdrForm") = True Then IniDelSek GlINI, "AdrForm"
If IniGetSek(GlINI, "ManForm") = True Then IniDelSek GlINI, "ManForm"
If IniGetSek(GlINI, "Termin") = True Then IniDelSek GlINI, "Termin"
If IniGetSek(GlINI, "Aufgaben") = True Then IniDelSek GlINI, "Aufgaben"
If IniGetSek(GlINI, "Ketten") = True Then IniDelSek GlINI, "Ketten"
If IniGetSek(GlINI, "Email") = True Then IniDelSek GlINI, "Email"

End Sub
Private Sub ALoad()
On Error Resume Next

Dim FeGro As Integer

Set FM = frmLayout
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set ChMen = FM.chkMenAn
Set ChRah = FM.chkOffRa
Set ChBut = FM.chkSysBu
Set ChReg = FM.chkFaReg
Set ChMod = FM.chkFaMod
Set CmFar = FM.cmbFarbe
Set CmSkn = FM.cmbSkins
Set CmTas = FM.cmbTaskb
Set CmFen = FM.cmbFenst
Set TxBre = FM.txtBreit
Set TxHoh = FM.txtHoehe

Set clFil = New clsFile

With CmFar
    .AddItem "Schema Blau"
    .ItemData(0) = 1
    .AddItem "Schema Schwarz"
    .ItemData(1) = 2
    .AddItem "Schema Silber"
    .ItemData(2) = 3
    .AddItem "Schema Aqua"
    .ItemData(3) = 4
    .AddItem "Schema Weiß"
    .ItemData(4) = 5
    .AddItem "Schema Ocean"
    .ItemData(5) = 6
    .AddItem "Schema Dark"
    .ItemData(6) = 7
    .AddItem "Schema Ergo"
    .ItemData(7) = 8
    .ListIndex = GlSty - 1
End With

With CmSkn
    .AddItem "Traditionell"
    .ItemData(0) = 1
    .AddItem "Modernstyle"
    .ItemData(1) = 2
    .AddItem "Officestyle"
    .ItemData(2) = 3
    .ListIndex = GlSkn - 1
End With
    
With CmTas
    .AddItem "Taskbar seitlich"
    .ItemData(0) = 1
    .AddItem "Taskbar unten"
    .ItemData(1) = 2
    .ListIndex = GlTkB - 1
End With

With CmFen
    .AddItem "Variable"
    .ItemData(0) = 1
    .AddItem "Maximiert"
    .ItemData(1) = 2
    .AddItem "Vorgegeben"
    .ItemData(2) = 3
    .ListIndex = Right$(GlFeG, 1) - 1
End With

TxBre.Text = CStr(IniGetVal("Layout", "FeVoBr"))
TxHoh.Text = CStr(IniGetVal("Layout", "FeVoHo"))

TxBre.Pattern = "\d*"
TxHoh.Pattern = "\d*"

If Right$(GlFeG, 1) = 3 Then
    TxBre.Enabled = True
    TxHoh.Enabled = True
End If

If GlMeA = True Then ChMen.Value = xtpChecked
If GlRah = True Then ChRah.Value = xtpChecked
If GlBty = True Then ChBut.Value = xtpChecked
If GlFRg = True Then ChReg.Value = xtpChecked
If GlFMo = True Then ChMod.Value = xtpChecked

FM.BackColor = GlBak
ChMen.BackColor = GlBak
ChRah.BackColor = GlBak
ChBut.BackColor = GlBak
ChReg.BackColor = GlBak
ChMod.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak

Set clFil = Nothing

End Sub
Private Sub ASave()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String
Dim StyId As Integer
Dim SknId As Integer
Dim TskBa As Integer
Dim FeGro As Integer
Dim FeBre As Integer
Dim FeHoh As Integer
Dim MeAni As Boolean
Dim Rahme As Boolean
Dim SyBut As Boolean
Dim FaReg As Boolean
Dim FaMod As Boolean

Set FM = frmLayout
Set ChMen = FM.chkMenAn
Set ChRah = FM.chkOffRa
Set ChBut = FM.chkSysBu
Set ChReg = FM.chkFaReg
Set ChMod = FM.chkFaMod
Set CmFar = FM.cmbFarbe
Set CmSkn = FM.cmbSkins
Set CmTas = FM.cmbTaskb
Set CmFen = FM.cmbFenst
Set TxBre = FM.txtBreit
Set TxHoh = FM.txtHoehe

TeTit = "Programmneustart"
TeMai = "Möchten Sie das Programm jetzt neu starten?"
TeInh = "Damit die vorgenommenen Einstellungen wirksam werden, muss das Programm neu gestartet werden."
TeFus = "Bei einem Programmneustart werden alle Daten gespeichert und die Verdingung zur Datenbank neu aufgebaut."

StyId = CmFar.ItemData(CmFar.ListIndex)
SknId = CmSkn.ItemData(CmSkn.ListIndex)
TskBa = CmTas.ItemData(CmTas.ListIndex)
FeGro = CmFen.ItemData(CmFen.ListIndex)

If ChMen.Value = xtpChecked Then
    MeAni = True
Else
    MeAni = False
End If
If ChRah.Value = xtpChecked Then
    Rahme = True
Else
    Rahme = False
End If
If ChBut.Value = xtpChecked Then
    SyBut = True
Else
    SyBut = False
End If
If ChReg.Value = xtpChecked Then
    FaReg = True
Else
    FaReg = False
End If
If ChMod.Value = xtpChecked Then
    FaMod = True
Else
    FaMod = False
End If

If IsNumeric(TxBre.Text) = True Then
    FeBre = TxBre.Text
Else
    FeBre = 1600
End If

If IsNumeric(TxHoh.Text) = True Then
    FeHoh = TxHoh.Text
Else
    FeHoh = 900
End If

IniSetVal "RDPSek", "PrgLay", "P" & StyId
IniSetVal "Layout", "SknSet", "A" & SknId
IniSetVal "Layout", "TskBar", "T" & TskBa
IniSetVal "Layout", "FenGro", "A" & FeGro

IniSetVal "GUI", "Rahmen", Rahme
IniSetVal "Layout", "MenAni", MeAni
IniSetVal "Layout", "SysBut", SyBut
IniSetVal "Layout", "FarReg", FaReg
IniSetVal "Layout", "FarMod", FaMod

IniSetVal "Layout", "FeVoBr", FeBre
IniSetVal "Layout", "FeVoHo", FeHoh

SMeFr TeTit, TeMai, TeInh, TeFus, False, 0, False, FM.hwnd
If GlMes = 33565 Then
    ALayo
    GlRes = True 'Reset der Einstellungen
    Unload Me
    DoEvents
    Unload frmMain
Else
    SNeSt True
    Unload Me
End If

End Sub
Private Sub AStan()
On Error Resume Next

Set FM = frmLayout
Set ChMen = FM.chkMenAn
Set ChRah = FM.chkOffRa
Set ChBut = FM.chkSysBu
Set ChReg = FM.chkFaReg
Set ChMod = FM.chkFaMod
Set CmFar = FM.cmbFarbe
Set CmSkn = FM.cmbSkins
Set CmTas = FM.cmbTaskb
Set TxBre = FM.txtBreit
Set TxHoh = FM.txtHoehe

GlSty = 8
CmFar.ListIndex = GlSty - 1

GlSkn = 2 'Skinset Einstellung
CmSkn.ListIndex = GlSkn - 1

GlTkB = 2 'Taskbar Einstellung
CmTas.ListIndex = GlTkB - 1

GlMeA = True
ChMen = xtpChecked

GlRah = True
ChRah.Value = xtpChecked

GlBty = False
ChBut.Value = xtpUnchecked

GlFRg = False
ChReg.Value = xtpUnchecked

GlFMo = True
ChMod.Value = xtpChecked

TxBre.Text = 1600
TxHoh.Text = 900

End Sub
Private Sub btnAbbru_Click()
    Unload Me
End Sub
Private Sub btnLosche_Click()
    AStan
End Sub
Private Sub btnWeiter_Click()
    ASave
End Sub

Private Sub btnZuruck_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50891)
TeMai = IniGetOpt("Hilfe", 50892)
TeInh = IniGetOpt("Hilfe", 50893)
TeFus = IniGetOpt("Hilfe", 50894)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd

End Sub
Private Sub chkFaReg_Click()
On Error Resume Next

Set FM = frmLayout
Set ChMen = FM.chkMenAn
Set ChRah = FM.chkOffRa
Set ChBut = FM.chkSysBu
Set ChReg = FM.chkFaReg
Set ChMod = FM.chkFaMod

If FoLad = False Then
    If ChReg.Value = xtpChecked Then 'Wenn farbige Register, dann keine Modulfarben
        FoLad = True
        ChMod.Value = xtpUnchecked
        ChMod.Enabled = False
        FoLad = False
    Else
        ChMod.Enabled = True
    End If
End If

End Sub
Private Sub chkOffRa_Click()
On Error Resume Next

Set FM = frmLayout
Set ChMen = FM.chkMenAn
Set ChRah = FM.chkOffRa
Set ChBut = FM.chkSysBu
Set ChReg = FM.chkFaReg
Set ChMod = FM.chkFaMod

If FoLad = False Then
    If ChRah.Value = xtpChecked Then 'Wenn Designerrahmen, dann keine farbigen Register
        FoLad = True
        ChReg.Value = xtpUnchecked
        ChReg.Enabled = False
        FoLad = False
    Else
        ChReg.Enabled = True
    End If
End If

End Sub

Private Sub chkSysBu_Click()
On Error Resume Next

Set FM = frmLayout
Set ChMen = FM.chkMenAn
Set ChRah = FM.chkOffRa
Set ChBut = FM.chkSysBu
Set ChReg = FM.chkFaReg
Set ChMod = FM.chkFaMod

If FoLad = False Then
    If ChBut.Value = xtpChecked Then 'Wenn runder Systembutton, dann nur mit Designerrahmen
        FoLad = True
        ChRah.Value = xtpChecked
        ChRah.Enabled = False
        FoLad = False
    Else
        ChRah.Enabled = True
    End If
End If

End Sub

Private Sub cmbFenst_Click()
On Error Resume Next

Dim FeGro As Integer

Set CmFen = Me.cmbFenst
Set TxBre = Me.txtBreit
Set TxHoh = Me.txtHoehe

FeGro = CmFen.ItemData(CmFen.ListIndex)

If FoLad = False Then
    If FeGro = 3 Then
        TxBre.Enabled = True
        TxHoh.Enabled = True
    Else
        TxBre.Enabled = False
        TxHoh.Enabled = False
    End If
End If

End Sub
Private Sub Form_Load()
On Error Resume Next

FoLad = True
ALoad
FoLad = False
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmLayout = Nothing
End Sub


Private Sub txtBreit_GotFocus()
    Me.txtBreit.SelStart = 0
    Me.txtBreit.SelLength = Len(Me.txtBreit.Text)
End Sub

Private Sub txtBreit_LostFocus()
On Error Resume Next

Set TxBre = Me.txtBreit
Set TxHoh = Me.txtHoehe

If IsNumeric(TxBre.Text) = False Then
    TxBre.Text = 1600
End If

End Sub
Private Sub txtHoehe_GotFocus()
    Me.txtHoehe.SelStart = 0
    Me.txtHoehe.SelLength = Len(Me.txtHoehe.Text)
End Sub

Private Sub txtHoehe_LostFocus()
On Error Resume Next

Set TxBre = Me.txtBreit
Set TxHoh = Me.txtHoehe

If IsNumeric(TxHoh.Text) = False Then
    TxHoh.Text = 900
End If

End Sub
