VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmLaImp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Laborimport"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   9
      Top             =   3600
      Width           =   5800
      _Version        =   1048579
      _ExtentX        =   10231
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchlieﬂ 
         Height          =   400
         Left            =   3800
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   400
         Width           =   1200
         _Version        =   1048579
         _ExtentX        =   2117
         _ExtentY        =   706
         _StockProps     =   79
         Caption         =   "&Schlieﬂen"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton btnWeiter 
         Default         =   -1  'True
         Height          =   400
         Left            =   2400
         TabIndex        =   11
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
         Left            =   1100
         TabIndex        =   10
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
   Begin XtremeSuiteControls.CheckBox chkManda 
      Height          =   255
      Left            =   800
      TabIndex        =   7
      Top             =   2700
      Width           =   1305
      _Version        =   1048579
      _ExtentX        =   2302
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Mandant"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  '2D
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Kein
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   5200
      Width           =   80
   End
   Begin XtremeSuiteControls.ComboBox cmbLaKat 
      Height          =   315
      Left            =   800
      TabIndex        =   2
      Top             =   1200
      Width           =   4200
      _Version        =   1048579
      _ExtentX        =   7408
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
      DropDownItemCount=   14
   End
   Begin XtremeSuiteControls.CheckBox chkMulti 
      Height          =   225
      Left            =   800
      TabIndex        =   3
      Top             =   1800
      Width           =   1305
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Multiplikator"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtMulti 
      Height          =   350
      Left            =   800
      TabIndex        =   4
      Top             =   2100
      Width           =   1100
      _Version        =   1048579
      _ExtentX        =   1940
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Text            =   "1,00"
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkPreis 
      Height          =   225
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1305
      _Version        =   1048579
      _ExtentX        =   2293
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Einheitspreis"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPreis 
      Height          =   350
      Left            =   2400
      TabIndex        =   6
      Top             =   2100
      Width           =   1095
      _Version        =   1048579
      _ExtentX        =   1931
      _ExtentY        =   617
      _StockProps     =   77
      ForeColor       =   4473924
      BackColor       =   16777215
      Text            =   "0,00"
      BackColor       =   16777215
      Alignment       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkLaKat 
      Height          =   225
      Left            =   800
      TabIndex        =   1
      Top             =   900
      Width           =   1500
      _Version        =   1048579
      _ExtentX        =   2646
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Laborkatalog"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cmbManda 
      Height          =   315
      Left            =   800
      TabIndex        =   8
      Top             =   3000
      Width           =   4200
      _Version        =   1048579
      _ExtentX        =   7408
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      Style           =   2
      Text            =   "ComboBox1"
   End
   Begin VB.Label lblLabe1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte w‰hlen Sie, mit welchem Laborkatalog der Laborbericht abgeglichen werden soll, um fehlende Daten zu erg‰nzen."
      Height          =   435
      Left            =   800
      TabIndex        =   13
      Top             =   100
      Width           =   4500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   660
      Left            =   0
      Top             =   0
      Width           =   5800
   End
End
Attribute VB_Name = "frmLaImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private TxPre As XtremeSuiteControls.FlatEdit
Private TxMul As XtremeSuiteControls.FlatEdit
Private CheMu As XtremeSuiteControls.CheckBox
Private ChePr As XtremeSuiteControls.CheckBox
Private CheKa As XtremeSuiteControls.CheckBox
Private CheMa As XtremeSuiteControls.CheckBox
Private CmKat As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox

Private Sub FKonf()
On Error GoTo SuErr

Dim AktZa As Integer
Dim AktPo As Integer

Set Rahm0 = Me.frmRahm0
Set TxMul = Me.txtMulti
Set TxPre = Me.txtPreis
Set CmKat = Me.cmbLaKat
Set CmMan = Me.cmbManda
Set ChePr = Me.chkPreis
Set CheMu = Me.chkMulti
Set CheMa = Me.chkManda
Set CheKa = Me.chkLaKat

For AktZa = 1 To UBound(GlLab)
    CmKat.AddItem GlLab(AktZa, 1)
    CmKat.ItemData(AktZa) = GlLab(AktZa, 0)
Next AktZa
CmKat.ListIndex = GlStL - 1 'Standardlaborkatalog

For AktZa = 1 To UBound(GlMan)
    CmMan.AddItem GlMan(AktZa, 1)
    CmMan.ItemData(AktZa - 1) = GlMan(AktZa, 2)
Next AktZa

For AktZa = 1 To UBound(GlMan)
    If GlMan(GlSMa, 2) = GlMan(AktZa, 2) Then
        CmMan.ListIndex = AktZa - 1
        Exit For
    End If
Next AktZa

If GlLaF <= 0 Then
    TxMul.Text = "1,00"
Else
    TxMul.Text = GlLaF
End If

If GlLaI = True Then
    CheKa.Value = xtpChecked
End If

If GlLiM = True Then
    CheMa.Value = xtpChecked
End If

Me.BackColor = GlBak
Rahm0.BackColor = GlBak
ChePr.BackColor = GlBak
CheMu.BackColor = GlBak
CheKa.BackColor = GlBak
CheMa.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FKonf " & Err.Number
Resume Next

End Sub
Private Sub TWeit()
On Error GoTo SuErr

Dim AktZa As Integer

Set TxMul = Me.txtMulti
Set TxPre = Me.txtPreis
Set CmKat = Me.cmbLaKat
Set CmMan = Me.cmbManda
Set ChePr = Me.chkPreis
Set CheMu = Me.chkMulti
Set CheKa = Me.chkLaKat
Set CheMa = Me.chkManda

If CheMu.Value = xtpChecked Then
    If IsNumeric(TxMul.Text) = True Then
        GlLiF = True
        GlLaF = TxMul.Text
    Else
        GlLiF = False
        GlLaF = 1
    End If
Else
    GlLiF = False
    GlLaF = 1
End If

If ChePr.Value = xtpChecked Then
    If IsNumeric(TxPre.Text) = True Then
        GlLiP = True
        GlLaP = TxPre.Text
    Else
        GlLiP = False
        GlLaP = 0
    End If
Else
    GlLiP = False
    GlLaP = 0
End If

If CheMa.Value = xtpChecked Then
    GlLaM = CmMan.ItemData(CmMan.ListIndex)
End If

If CheKa.Value = xtpChecked Then
    GlStL = CmKat.ListIndex + 1
    For AktZa = 1 To UBound(GlLab)
        If GlLab(AktZa, 1) = CmKat.Text Then
            IniSetVal "Vorgabe", "StaLab", AktZa
            GlStL = AktZa 'Standardlaborkatalog
            Exit For
        End If
    Next AktZa
End If

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TWeit " & Err.Number
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
Private Sub btnSchlieﬂ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    TWeit
    Unload Me
    STran 2
End Sub

Private Sub chkLaKat_Click()
On Error Resume Next

Set CheKa = Me.chkLaKat

If CheKa.Value = xtpChecked Then
    GlLaI = True
Else
    GlLaI = False
End If

IniSetVal "System", "LaImVe", GlLaI

End Sub

Private Sub chkManda_Click()
On Error Resume Next

Set CheMa = Me.chkManda

If CheMa.Value = xtpChecked Then
    GlLiM = True
Else
    GlLiM = False
End If

IniSetVal "System", "LaImMa", GlLiM

End Sub
Private Sub Form_Load()
On Error Resume Next
FKonf
AFont Me
SFrame 1, Me.hwnd

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmLaImp = Nothing
End Sub
Private Sub txtMulti_GotFocus()
    Me.txtMulti.SelStart = 0
    Me.txtMulti.SelLength = Len(Me.txtMulti.Text)
End Sub
Private Sub txtPreis_GotFocus()
    Me.txtPreis.SelStart = 0
    Me.txtPreis.SelLength = Len(Me.txtPreis.Text)
End Sub

