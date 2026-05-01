VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmKaAnp 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Einträge Anpassen"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin XtremeSuiteControls.GroupBox frmRahm1 
      Height          =   4005
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      _Version        =   1048579
      _ExtentX        =   10583
      _ExtentY        =   7056
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkGrupp 
         Height          =   220
         Left            =   2620
         TabIndex        =   19
         Top             =   3000
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Laborgruppe"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkFavor 
         Height          =   220
         Left            =   380
         TabIndex        =   15
         Top             =   3000
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Favorit"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAnalo 
         Height          =   225
         Left            =   2620
         TabIndex        =   6
         Top             =   800
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Analog"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkSteur 
         Height          =   225
         Left            =   1520
         TabIndex        =   4
         Top             =   800
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Steuer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkPreis 
         Height          =   225
         Left            =   3720
         TabIndex        =   8
         Top             =   800
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Preis"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkMulti 
         Height          =   220
         Left            =   380
         TabIndex        =   2
         Top             =   800
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Faktor"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPrei1 
         Height          =   315
         Left            =   3700
         TabIndex        =   9
         Top             =   1100
         Width           =   1395
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   547
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "0,00"
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtSteue 
         Height          =   350
         Left            =   1500
         TabIndex        =   5
         Top             =   1100
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1676
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "0,0"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMulti 
         Height          =   350
         Left            =   360
         TabIndex        =   3
         Top             =   1100
         Width           =   940
         _Version        =   1048579
         _ExtentX        =   1658
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Text            =   "1,0"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRest2 
         Height          =   350
         Left            =   1335
         TabIndex        =   13
         Top             =   2200
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   873
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "1"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtRest1 
         Height          =   350
         Left            =   1335
         TabIndex        =   11
         Top             =   1800
         Width           =   495
         _Version        =   1048579
         _ExtentX        =   873
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "1"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.RadioButton optRest3 
         Height          =   240
         Left            =   360
         TabIndex        =   14
         Top             =   2550
         Width           =   4995
         _Version        =   1048579
         _ExtentX        =   8819
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Darf nur als alleinige Leistung am Tag abgerechnet werden"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optRest2 
         Height          =   240
         Left            =   360
         TabIndex        =   12
         Top             =   2200
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Darf max."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optRest1 
         Height          =   240
         Left            =   360
         TabIndex        =   10
         Top             =   1850
         Width           =   1005
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Darf max."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbAnalo 
         Height          =   315
         Left            =   2600
         TabIndex        =   7
         Top             =   1100
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1693
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbFavor 
         Height          =   315
         Left            =   360
         TabIndex        =   16
         Top             =   3300
         Width           =   950
         _Version        =   1048579
         _ExtentX        =   1667
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkZiffe 
         Height          =   220
         Left            =   1520
         TabIndex        =   17
         Top             =   3000
         Width           =   1000
         _Version        =   1048579
         _ExtentX        =   1764
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Ziffer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtZiffe 
         Height          =   350
         Left            =   1500
         TabIndex        =   18
         Top             =   3300
         Width           =   940
         _Version        =   1048579
         _ExtentX        =   1658
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cmbGrupp 
         Height          =   315
         Left            =   2600
         TabIndex        =   20
         Top             =   3300
         Width           =   2500
         _Version        =   1048579
         _ExtentX        =   4419
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.Label Label19 
         Height          =   240
         Left            =   1905
         TabIndex        =   27
         Top             =   1850
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "mal pro Behandlungstag abgerechnet werden"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label21 
         Height          =   240
         Left            =   1905
         TabIndex        =   26
         Top             =   2200
         Width           =   3495
         _Version        =   1048579
         _ExtentX        =   6174
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "mal pro Rechnung abgerechnet werden"
         Transparent     =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmKaAnp.frx":0000
         Height          =   585
         Left            =   400
         TabIndex        =   24
         Top             =   100
         Width           =   5505
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   25
      Top             =   5000
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnSchließ 
      Height          =   400
      Left            =   4680
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4240
      Width           =   1155
      _Version        =   1048579
      _ExtentX        =   2037
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Schließen"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnWeiter 
      Height          =   400
      Left            =   3240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4240
      Width           =   1350
      _Version        =   1048579
      _ExtentX        =   2381
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Weiter"
      UseVisualStyle  =   -1  'True
      PushButtonStyle =   2
   End
   Begin XtremeSuiteControls.PushButton btnHilfe 
      Height          =   400
      Left            =   2000
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4240
      Width           =   1140
      _Version        =   1048579
      _ExtentX        =   2011
      _ExtentY        =   706
      _StockProps     =   79
      Caption         =   "&Hilfe"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit1 
      Height          =   200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   200
      _Version        =   1048579
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      FlatStyle       =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   130
      X2              =   5800
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   130
      X2              =   5800
      Y1              =   4000
      Y2              =   4000
   End
End
Attribute VB_Name = "frmKaAnp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control

Private Rahm1 As XtremeSuiteControls.GroupBox
Private ChFak As XtremeSuiteControls.CheckBox
Private ChStu As XtremeSuiteControls.CheckBox
Private ChPre As XtremeSuiteControls.CheckBox
Private ChAna As XtremeSuiteControls.CheckBox
Private ChFav As XtremeSuiteControls.CheckBox
Private ChZif As XtremeSuiteControls.CheckBox
Private ChGrp As XtremeSuiteControls.CheckBox
Private Opti1 As XtremeSuiteControls.RadioButton
Private Opti2 As XtremeSuiteControls.RadioButton
Private Opti3 As XtremeSuiteControls.RadioButton
Private TxFak As XtremeSuiteControls.FlatEdit
Private TxStu As XtremeSuiteControls.FlatEdit
Private TxPre As XtremeSuiteControls.FlatEdit
Private TxAn1 As XtremeSuiteControls.FlatEdit
Private TxAn2 As XtremeSuiteControls.FlatEdit
Private CmAna As XtremeSuiteControls.ComboBox
Private CmFav As XtremeSuiteControls.ComboBox
Private CmGrp As XtremeSuiteControls.ComboBox
Private RpSel As XtremeReportControl.ReportSelectedRows

Private clFen As clsFenster
Private Sub TAbs()
On Error GoTo OpErr
'Ändert die Rechnungen

Dim RowNr As Long
Dim AnzPo As Integer
Dim RpCo8 As XtremeReportControl.ReportControl
Dim RpCls As XtremeReportControl.ReportColumns

Set FM = frmMain
Set RpCo8 = FM.repCont8
Set RpCls = RpCo8.Columns
Set RpSel = RpCo8.SelectedRows

AnzPo = RpSel.Count

If AnzPo > 0 Then
    K_Anpa
    DoEvents
    KUpKa
End If

Set RpSel = Nothing
Set RpCls = Nothing
Set RpCo8 = Nothing

Exit Sub

OpErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TAbs " & Err.Number
Resume Next

End Sub
Private Sub FInit()
On Error GoTo SuErr

Dim Katal As String
Dim AktZa As Integer
Dim GesZa As Integer
Dim LauZa As Integer

Set FM = frmKaAnp
Set Rahm1 = FM.frmRahm1
Set ChFak = FM.chkMulti
Set ChStu = FM.chkSteur
Set ChPre = FM.chkPreis
Set ChAna = FM.chkAnalo
Set ChFav = FM.chkFavor
Set ChZif = FM.chkZiffe
Set ChGrp = FM.chkGrupp
Set TxFak = FM.txtMulti
Set TxStu = FM.txtSteue
Set TxPre = FM.txtPrei1
Set TxAn1 = FM.txtRest1
Set TxAn2 = FM.txtRest2
Set CmAna = FM.cmbAnalo
Set CmFav = FM.cmbFavor
Set CmGrp = FM.cmbGrupp
Set Opti1 = FM.optRest1
Set Opti2 = FM.optRest2
Set Opti3 = FM.optRest3

Katal = Left$(GlNod, 1)

Select Case Katal
Case "A": 'Gebührenkataloge
    ChGrp.Enabled = False
    CmGrp.Enabled = False
Case "G": 'Laborparameter
    ChAna.Enabled = False
    ChStu.Enabled = False
    ChGrp.Enabled = True
    TxStu.Enabled = False
    CmAna.Enabled = False
    CmGrp.Enabled = True
    Opti1.Enabled = False
    Opti2.Enabled = False
    Opti3.Enabled = False
Case "I": 'Arzneikataloge
    ChAna.Enabled = False
    ChFak.Enabled = False
    ChGrp.Enabled = False
    TxFak.Enabled = False
    CmAna.Enabled = False
    CmGrp.Enabled = False
    Opti1.Enabled = False
    Opti2.Enabled = False
    Opti3.Enabled = False
Case "P": 'Artikelkataloge
    ChAna.Enabled = False
    ChFak.Enabled = False
    ChGrp.Enabled = False
    TxFak.Enabled = False
    CmAna.Enabled = False
    CmGrp.Enabled = False
    Opti1.Enabled = False
    Opti2.Enabled = False
    Opti3.Enabled = False
Case Else:
    ChFak.Enabled = False
    ChStu.Enabled = False
    ChPre.Enabled = False
    ChGrp.Enabled = False
    CmGrp.Enabled = False
    TxFak.Enabled = False
    TxStu.Enabled = False
    TxPre.Enabled = False
    Opti1.Enabled = False
    Opti2.Enabled = False
    Opti3.Enabled = False
End Select

With CmAna
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
    .ListIndex = 0
End With

With CmFav
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
    .ListIndex = 0
End With

With CmGrp
    For AktZa = 1 To UBound(GlLGr)
        .AddItem GlLGr(AktZa, 1)
        .ItemData(AktZa - 1) = GlLGr(AktZa, 0)
    Next AktZa
    .ListIndex = 0
End With

FM.BackColor = GlBak
Rahm1.BackColor = GlBak
ChFak.BackColor = GlBak
ChZif.BackColor = GlBak
ChStu.BackColor = GlBak
ChPre.BackColor = GlBak
ChAna.BackColor = GlBak
ChGrp.BackColor = GlBak
ChFav.BackColor = GlBak
Opti1.BackColor = GlBak
Opti2.BackColor = GlBak
Opti3.BackColor = GlBak

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FInit " & Err.Number
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
    TAbs
    Unload Me
End Sub

Private Sub txtMulti_GotFocus()
    Me.txtMulti.SelStart = 0
    Me.txtMulti.SelLength = Len(Me.txtMulti.Text)
End Sub
Private Sub txtPrei1_GotFocus()
    Me.txtPrei1.SelStart = 0
    Me.txtPrei1.SelLength = Len(Me.txtPrei1.Text)
End Sub

Private Sub txtRest1_GotFocus()
    Me.txtRest1.SelStart = 0
    Me.txtRest1.SelLength = Len(Me.txtRest1.Text)
End Sub

Private Sub txtRest2_GotFocus()
    Me.txtRest2.SelStart = 0
    Me.txtRest2.SelLength = Len(Me.txtRest2.Text)
End Sub
Private Sub txtSteue_GotFocus()
    Me.txtSteue.SelStart = 0
    Me.txtSteue.SelLength = Len(Me.txtSteue.Text)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmKaAnp = Nothing
End Sub
Private Sub Form_Load()
On Error Resume Next

Set clFen = New clsFenster
clFen.hwnd = Me.hwnd

FInit
AFont Me
SFrame 1, Me.hwnd

clFen.FenVor

Set clFen = Nothing

End Sub
Private Sub optRest1_Click()
On Error Resume Next

Set TxAn1 = Me.txtRest1
Set TxAn2 = Me.txtRest2

TxAn1.Enabled = True
TxAn2.Enabled = False

End Sub

Private Sub optRest2_Click()
On Error Resume Next

Set TxAn1 = Me.txtRest1
Set TxAn2 = Me.txtRest2

TxAn1.Enabled = False
TxAn2.Enabled = True

End Sub

Private Sub optRest3_Click()
On Error Resume Next

Set TxAn1 = Me.txtRest1
Set TxAn2 = Me.txtRest2

TxAn1.Enabled = False
TxAn2.Enabled = False

End Sub
