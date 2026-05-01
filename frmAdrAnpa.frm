VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#16.3#0"; "Codejock.Controls.v16.3.1.ocx"
Begin VB.Form frmAdrAnpa 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Adressen Anpassen"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin XtremeSuiteControls.GroupBox frmRahm0 
      Height          =   1100
      Left            =   0
      TabIndex        =   26
      Top             =   6700
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   1940
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnSchließ 
         Height          =   400
         Left            =   6000
         TabIndex        =   33
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
         Height          =   400
         Left            =   4600
         TabIndex        =   32
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
         Left            =   3200
         TabIndex        =   31
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
         Left            =   1900
         TabIndex        =   29
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
      Height          =   6700
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8000
      _Version        =   1048579
      _ExtentX        =   14111
      _ExtentY        =   11818
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkVersa 
         Height          =   220
         Left            =   600
         TabIndex        =   14
         Top             =   5700
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Versandart"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkMitar 
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   4900
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Status"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.TreeView trvList1 
         Height          =   4600
         Left            =   4200
         TabIndex        =   28
         Top             =   1700
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5644
         _ExtentY        =   8114
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         BackColor       =   16777215
         ForeColor       =   4473924
      End
      Begin XtremeSuiteControls.UpDown updCont1 
         Height          =   360
         Left            =   3510
         TabIndex        =   22
         Top             =   5190
         Width           =   255
         _Version        =   1048579
         _ExtentX        =   450
         _ExtentY        =   635
         _StockProps     =   64
         Enabled         =   0   'False
         Min             =   1
         Value           =   1
         Max             =   10
         SyncBuddy       =   -1  'True
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtKopie"
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.CheckBox chkGrupp 
         Height          =   220
         Left            =   4200
         TabIndex        =   25
         Top             =   900
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Adressengruppe"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKopie 
         Height          =   220
         Left            =   2400
         TabIndex        =   20
         Top             =   4900
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Ausdrucke"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkZahlu 
         Height          =   220
         Left            =   600
         TabIndex        =   4
         Top             =   1700
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Zahlungsziel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkKatal 
         Height          =   220
         Left            =   600
         TabIndex        =   2
         Top             =   900
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Gebührensatz"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbVersi 
         Height          =   310
         Left            =   600
         TabIndex        =   3
         Top             =   1200
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbZaZie 
         Height          =   310
         Left            =   600
         TabIndex        =   5
         Top             =   2000
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox2"
      End
      Begin XtremeSuiteControls.FlatEdit txtKopie 
         Height          =   350
         Left            =   2400
         TabIndex        =   21
         Top             =   5200
         Width           =   1100
         _Version        =   1048579
         _ExtentX        =   1940
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "1"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbManda 
         Height          =   310
         Left            =   600
         TabIndex        =   7
         Top             =   2800
         Width           =   3200
         _Version        =   1048579
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkManda 
         Height          =   220
         Left            =   600
         TabIndex        =   6
         Top             =   2500
         Width           =   1800
         _Version        =   1048579
         _ExtentX        =   3175
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Mandant"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkSerie 
         Height          =   225
         Left            =   600
         TabIndex        =   8
         Top             =   3300
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Serienmailadresse"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbSerie 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   3600
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkOutlo 
         Height          =   225
         Left            =   600
         TabIndex        =   10
         Top             =   4100
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Onlineabgleich"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkRepli 
         Height          =   225
         Left            =   2400
         TabIndex        =   16
         Top             =   3300
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Synchronisierung"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbOutlo 
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   4400
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbRepli 
         Height          =   315
         Left            =   2400
         TabIndex        =   17
         Top             =   3600
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkPassi 
         Height          =   225
         Left            =   2400
         TabIndex        =   18
         Top             =   4100
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Gelöscht"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbPassi 
         Height          =   315
         Left            =   2400
         TabIndex        =   19
         Top             =   4400
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbMitar 
         Height          =   315
         Left            =   600
         TabIndex        =   13
         Top             =   5200
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cmbGrupp 
         Height          =   315
         Left            =   4200
         TabIndex        =   27
         Top             =   1200
         Width           =   2000
         _Version        =   1048579
         _ExtentX        =   3519
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkNumme 
         Height          =   220
         Left            =   2400
         TabIndex        =   23
         Top             =   5700
         Width           =   1600
         _Version        =   1048579
         _ExtentX        =   2822
         _ExtentY        =   388
         _StockProps     =   79
         Caption         =   "Pat.-Nummer"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtNumme 
         Height          =   350
         Left            =   2400
         TabIndex        =   24
         Top             =   6000
         Width           =   1400
         _Version        =   1048579
         _ExtentX        =   2469
         _ExtentY        =   617
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         Enabled         =   0   'False
         Text            =   "1"
         BackColor       =   16777215
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cmbVersa 
         Height          =   315
         Left            =   600
         TabIndex        =   15
         Top             =   6000
         Width           =   1500
         _Version        =   1048579
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         Enabled         =   0   'False
         Style           =   2
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdrAnpa.frx":0000
         Height          =   580
         Left            =   600
         TabIndex        =   30
         Top             =   100
         Width           =   7000
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Undurchsichtig
         BorderStyle     =   0  'Transparent
         Height          =   650
         Left            =   0
         Top             =   0
         Width           =   7900
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtDummy 
      Height          =   200
      Left            =   0
      TabIndex        =   0
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
End
Attribute VB_Name = "frmAdrAnpa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FM As Form
Private AktCo As VB.Control
Private Rahm0 As XtremeSuiteControls.GroupBox
Private Rahm1 As XtremeSuiteControls.GroupBox
Private CmVer As XtremeSuiteControls.ComboBox
Private CmZil As XtremeSuiteControls.ComboBox
Private CmMan As XtremeSuiteControls.ComboBox
Private CmOut As XtremeSuiteControls.ComboBox
Private CmRep As XtremeSuiteControls.ComboBox
Private CmSer As XtremeSuiteControls.ComboBox
Private CmPas As XtremeSuiteControls.ComboBox
Private CmGrp As XtremeSuiteControls.ComboBox
Private CmMit As XtremeSuiteControls.ComboBox
Private CmVrs As XtremeSuiteControls.ComboBox
Private TxKop As XtremeSuiteControls.FlatEdit
Private TxNum As XtremeSuiteControls.FlatEdit
Private ChVer As XtremeSuiteControls.CheckBox
Private ChZil As XtremeSuiteControls.CheckBox
Private ChKop As XtremeSuiteControls.CheckBox
Private ChGrp As XtremeSuiteControls.CheckBox
Private ChMan As XtremeSuiteControls.CheckBox
Private ChSer As XtremeSuiteControls.CheckBox
Private ChOut As XtremeSuiteControls.CheckBox
Private ChRep As XtremeSuiteControls.CheckBox
Private ChPas As XtremeSuiteControls.CheckBox
Private ChMit As XtremeSuiteControls.CheckBox
Private ChNum As XtremeSuiteControls.CheckBox
Private ChVrs As XtremeSuiteControls.CheckBox
Private TrLi1 As XtremeSuiteControls.TreeView
Private Knote As XtremeSuiteControls.TreeViewNode
Private UpCo1 As XtremeSuiteControls.UpDown
Private Sub FInit()
On Error GoTo SuErr

Dim AktZa As Integer
Dim ImMan As XtremeCommandBars.ImageManager

Set FM = frmAdrAnpa
Set Rahm0 = FM.frmRahm0
Set Rahm1 = FM.frmRahm1
Set CmVer = FM.cmbVersi
Set CmZil = FM.cmbZaZie
Set CmMan = FM.cmbManda
Set CmSer = FM.cmbSerie
Set CmOut = FM.cmbOutlo
Set CmRep = FM.cmbRepli
Set CmPas = FM.cmbPassi
Set CmMit = FM.cmbMitar
Set CmGrp = FM.cmbGrupp
Set CmVrs = FM.cmbVersa
Set TxKop = FM.txtKopie
Set TxNum = FM.txtNumme
Set ChVer = FM.chkKatal
Set ChZil = FM.chkZahlu
Set ChKop = FM.chkKopie
Set ChGrp = FM.chkGrupp
Set ChMan = FM.chkManda
Set ChSer = FM.chkSerie
Set ChOut = FM.chkOutlo
Set ChRep = FM.chkRepli
Set ChPas = FM.chkPassi
Set ChMit = FM.chkMitar
Set ChNum = FM.chkNumme
Set ChVrs = FM.chkVersa
Set TrLi1 = FM.trvList1
Set ImMan = frmMain.imgManag

For AktZa = 1 To UBound(GlGKa)
    CmVer.AddItem GlGKa(AktZa, 1)
    CmVer.ItemData(AktZa - 1) = GlGKa(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlZah)
    CmZil.AddItem GlZah(AktZa, 1)
    CmZil.ItemData(AktZa - 1) = GlZah(AktZa, 0)
Next AktZa

For AktZa = 1 To UBound(GlThe)
    CmMan.AddItem GlThe(AktZa, 13)
    CmMan.ItemData(AktZa - 1) = GlThe(AktZa, 0)
Next AktZa
CmMan.ListIndex = GlMan(GlSMa, 0) - 1

With CmSer
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmOut
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmRep
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmPas
    .AddItem "Ja"
    .ItemData(0) = 1
    .AddItem "Nein"
    .ItemData(1) = 2
End With

With CmMit
    .AddItem "Patient"
    .ItemData(0) = 1
    .AddItem "Mitarbeiter"
    .ItemData(1) = 2
    .AddItem "Mandant"
    .ItemData(2) = 3
    .AddItem "Verordner"
    .ItemData(3) = 4
End With

With CmGrp
    .AddItem "Hinzufügen"
    .ItemData(0) = 1
    .AddItem "Entfernen"
    .ItemData(1) = 2
End With

With CmVrs
    .AddItem "Postversand"
    .ItemData(0) = 0
    .AddItem "Emailversand"
    .ItemData(1) = 1
    .AddItem "Downloadlink"
    .ItemData(2) = 2
    .ListIndex = GlRVs - 1 'Standard-Rechnungsversandweg
End With

With TrLi1
    Select Case GlSty
    Case 8: .Appearance = xtpAppearanceOffice2013
    Case 7: .Appearance = xtpAppearanceOffice2013
    Case Else: .Appearance = xtpAppearanceResource
    End Select
    .Checkboxes = True
    .Font.SIZE = GlTFt.SIZE
    .Font.Name = GlTFt.Name
    .ForeColor = -2147483641
    .FullRowSelect = False
    .HideSelection = False
    .HotTracking = False
    .Icons = ImMan.Icons
    .IconSize = 16
    .LabelEdit = xtpTreeViewLabelManual
    .Scroll = True
    .ShowLines = xtpTreeViewShowLines
    .ShowPlusMinus = True
    .SingleSel = False
End With

Set Knote = TrLi1.Nodes.Add(, , "P801", "Adressen", IC16_Folder_View)
With Knote
    .Bold = True
    .Checked = False
End With

CmVer.ListIndex = 0
CmZil.ListIndex = 0
CmMan.ListIndex = GlSMa - 1
CmSer.ListIndex = 0
CmOut.ListIndex = 0
CmRep.ListIndex = 0
CmPas.ListIndex = 1
CmMit.ListIndex = 0
CmVrs.ListIndex = 0
CmGrp.ListIndex = 0

With TxKop
    .Pattern = "\d*"
    .SetMask "0", "_"
    .Text = 2
End With

With TxNum
    .Pattern = "\d*"
    .SetMask "00000", "_____"
    .Text = "00001"
End With

Select Case GlBut
Case RibTab_Mitarbeit:
    ChSer.Caption = "Kalenderspalte"
    ChRep.Caption = "Onlinetermine"
Case RibTab_Mandanten:
    ChSer.Caption = "Kalenderspalte"
    ChRep.Caption = "Onlinetermine"
End Select

FM.BackColor = GlBak
Rahm0.BackColor = GlBak
Rahm1.BackColor = GlBak
ChVer.BackColor = GlBak
ChZil.BackColor = GlBak
ChKop.BackColor = GlBak
ChGrp.BackColor = GlBak
ChMan.BackColor = GlBak
ChSer.BackColor = GlBak
ChOut.BackColor = GlBak
ChRep.BackColor = GlBak
ChPas.BackColor = GlBak
ChMit.BackColor = GlBak
ChNum.BackColor = GlBak
ChVrs.BackColor = GlBak

Set ImMan = Nothing

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "TInit " & Err.Number
Resume Next

End Sub
Private Sub FAnp()
On Error GoTo SuErr

Dim KatNr As Long
Dim ZalZi As Long
Dim Kopie As Long
Dim TheNr As Long
Dim AdNum As Long
Dim SerAd As Integer
Dim OutAd As Integer
Dim SynAd As Integer
Dim Passi As Integer
Dim Mitar As Integer
Dim Versa As Integer
Dim Grupp As Boolean
Dim GrEnf As Boolean

Set FM = frmAdrAnpa
Set CmVer = FM.cmbVersi
Set CmZil = FM.cmbZaZie
Set CmMan = FM.cmbManda
Set CmSer = FM.cmbSerie
Set CmRep = FM.cmbRepli
Set CmPas = FM.cmbPassi
Set CmMit = FM.cmbMitar
Set CmGrp = FM.cmbGrupp
Set CmVrs = FM.cmbVersa
Set TxKop = FM.txtKopie
Set TxNum = FM.txtNumme
Set ChRep = FM.chkRepli
Set ChPas = FM.chkPassi
Set ChVer = FM.chkKatal
Set ChZil = FM.chkZahlu
Set ChKop = FM.chkKopie
Set ChGrp = FM.chkGrupp
Set ChMan = FM.chkManda
Set ChSer = FM.chkSerie
Set ChMit = FM.chkMitar
Set ChNum = FM.chkNumme
Set ChVrs = FM.chkVersa

If ChVer.Value = xtpChecked Then
    KatNr = CmVer.ItemData(CmVer.ListIndex)
Else
    KatNr = 0
End If

If ChZil.Value = xtpChecked Then
    ZalZi = CmZil.ItemData(CmZil.ListIndex)
Else
    ZalZi = 0
End If

If ChKop.Value = xtpChecked Then
    Kopie = CInt(TxKop.Text)
Else
    Kopie = 0
End If

If ChMan.Value = xtpChecked Then
    TheNr = CmMan.ItemData(CmMan.ListIndex)
Else
    TheNr = 0
End If

If ChGrp.Value = xtpChecked Then
    Grupp = True
End If

If CmGrp.ListIndex = 1 Then
    GrEnf = True
End If

If ChSer.Value = xtpChecked Then
    SerAd = CmSer.ItemData(CmSer.ListIndex)
End If

If ChOut.Value = xtpChecked Then
    OutAd = CmOut.ItemData(CmOut.ListIndex)
End If

If ChRep.Value = xtpChecked Then
    SynAd = CmRep.ItemData(CmRep.ListIndex)
End If

If ChPas.Value = xtpChecked Then
    Passi = CmPas.ItemData(CmPas.ListIndex)
End If

If ChMit.Value = xtpChecked Then
    Mitar = CmMit.ItemData(CmMit.ListIndex)
End If

If ChNum.Value = xtpChecked Then
    AdNum = Val(TxNum.Text)
Else
    AdNum = 0
End If

If ChVrs.Value = xtpChecked Then
    Versa = CmVrs.ItemData(CmVrs.ListIndex)
Else
    Versa = -1
End If

S_AdAn KatNr, ZalZi, Kopie, Grupp, TheNr, SerAd, OutAd, SynAd, Passi, Mitar, GrEnf, AdNum, Versa

Exit Sub

SuErr:
If GlDbg = True Then MsgBox Err.Description, 48, "FAnp " & Err.Number
Resume Next

End Sub
Private Sub btnHilfe_Click()
On Error Resume Next

Dim TeTit As String
Dim TeMai As String
Dim TeInh As String
Dim TeFus As String

TeTit = IniGetOpt("Hilfe", 50911)
TeMai = IniGetOpt("Hilfe", 50912)
TeInh = IniGetOpt("Hilfe", 50913)
TeFus = IniGetOpt("Hilfe", 50914)

SMeFr TeTit, TeMai, TeInh, TeFus, False, 1, True, Me.hwnd
    
End Sub
Private Sub btnSchließ_Click()
    Unload Me
End Sub
Private Sub btnWeiter_Click()
    FAnp
    Unload Me
End Sub

Private Sub chkGrupp_Click()
On Error Resume Next

Set ChGrp = Me.chkGrupp
Set CmGrp = Me.cmbGrupp
Set TrLi1 = Me.trvList1

If ChGrp.Value = xtpChecked Then
    CmGrp.Enabled = True
    TrLi1.Enabled = True
Else
    CmGrp.Enabled = False
    TrLi1.Enabled = False
End If

End Sub

Private Sub chkKatal_Click()
On Error Resume Next

Set ChVer = Me.chkKatal
Set CmVer = Me.cmbVersi

If ChVer.Value = xtpChecked Then
    CmVer.Enabled = True
Else
    CmVer.Enabled = False
End If

End Sub

Private Sub chkKopie_Click()
On Error Resume Next

Set FM = frmAdrAnpa
Set ChKop = FM.chkKopie
Set TxKop = FM.txtKopie
Set UpCo1 = FM.updCont1

If ChKop.Value = xtpChecked Then
    TxKop.Enabled = True
    UpCo1.Enabled = True
Else
    TxKop.Enabled = False
    UpCo1.Enabled = False
End If

End Sub

Private Sub chkManda_Click()
On Error Resume Next

Set ChMan = Me.chkManda
Set CmMan = Me.cmbManda

If ChMan.Value = xtpChecked Then
    CmMan.Enabled = True
Else
    CmMan.Enabled = False
End If

End Sub

Private Sub chkMitar_Click()
On Error Resume Next

Set ChMit = Me.chkMitar
Set CmMit = Me.cmbMitar

If ChMit.Value = xtpChecked Then
    CmMit.Enabled = True
Else
    CmMit.Enabled = False
End If

End Sub
Private Sub chkNumme_Click()
On Error Resume Next

Set FM = frmAdrAnpa
Set ChNum = FM.chkNumme
Set TxNum = FM.txtNumme

If ChNum.Value = xtpChecked Then
    TxNum.Enabled = True
Else
    TxNum.Enabled = False
End If

End Sub

Private Sub chkOutlo_Click()
On Error Resume Next

Set ChOut = Me.chkOutlo
Set CmOut = Me.cmbOutlo

If ChOut.Value = xtpChecked Then
    CmOut.Enabled = True
Else
    CmOut.Enabled = False
End If

End Sub

Private Sub chkPassi_Click()
On Error Resume Next

Set ChPas = Me.chkPassi
Set CmPas = Me.cmbPassi

If ChPas.Value = xtpChecked Then
    CmPas.Enabled = True
Else
    CmPas.Enabled = False
End If

End Sub
Private Sub chkRepli_Click()
On Error Resume Next

Set ChRep = Me.chkRepli
Set CmRep = Me.cmbRepli

If ChRep.Value = xtpChecked Then
    CmRep.Enabled = True
Else
    CmRep.Enabled = False
End If

End Sub
Private Sub chkSerie_Click()
On Error Resume Next

Set ChSer = Me.chkSerie
Set CmSer = Me.cmbSerie

If ChSer.Value = xtpChecked Then
    CmSer.Enabled = True
Else
    CmSer.Enabled = False
End If

End Sub

Private Sub chkVersa_Click()
On Error Resume Next

Set ChVrs = Me.chkVersa
Set CmVrs = Me.cmbVersa

If ChVrs.Value = xtpChecked Then
    CmVrs.Enabled = True
Else
    CmVrs.Enabled = False
End If

End Sub
Private Sub chkZahlu_Click()
On Error Resume Next

Set ChZil = Me.chkZahlu
Set CmZil = Me.cmbZaZie

If ChZil.Value = xtpChecked Then
    CmZil.Enabled = True
Else
    CmZil.Enabled = False
End If

End Sub
Private Sub Form_Load()
    FInit
    AFont FM
    AdGru 4
    SFrame 1, Me.hwnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAdrAnpa = Nothing
End Sub

Private Sub trvList1_NodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Dim TreKy As String

Set TrLi1 = Me.trvList1

For Each Knote In TrLi1.Nodes
    Knote.Image = IC16_Folder_Close
Next Knote

Node.Image = IC16_Folder_Open
TrLi1.Nodes(1).Image = IC16_Folder_View

If Node.Key = "P801" Then
    For Each Knote In TrLi1.Nodes
        Knote.Checked = Node.Checked
    Next Knote
End If

End Sub
Private Sub trvList1_NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
On Error Resume Next

Dim TreKy As String

Set TrLi1 = Me.trvList1
    
For Each Knote In TrLi1.Nodes
    Knote.Image = IC16_Folder_Close
Next Knote

Node.Image = IC16_Folder_Open
TrLi1.Nodes(1).Image = IC16_Folder_View

End Sub

